"""Read-only access to classic Outlook Desktop over COM.

Everything here is read-only: it never sends, moves, deletes, marks read, or
edits source mail. It extracts each message into a detached ``MessageData``
(plain Python, no live COM references) so the rest of the tool — the .eml
builder, identity, export — needs no COM and is unit-testable without Outlook.

COM runs synchronously on the calling thread (this is a CLI, not an async
server), so a single ``pythoncom.CoInitialize()`` is enough.
"""
from __future__ import annotations

import os
import tempfile
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from email.utils import formataddr
from typing import Iterator

# DASL property tags (unicode variants) used to reach fields the Outlook object
# model does not expose directly.
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_SENT_REPRESENTING_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

OL_FOLDER_INBOX = 6


class OutlookError(Exception):
    pass


@dataclass
class AttachmentData:
    filename: str
    data: bytes
    mime_type: str | None = None
    content_id: str | None = None
    inline: bool = False


@dataclass
class MessageData:
    entry_id: str
    message_class: str
    subject: str
    sender_name: str
    sender_smtp: str
    to: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    bcc: list[str] = field(default_factory=list)
    sent_on: datetime | None = None
    received_time: datetime | None = None
    html_body: str = ""
    text_body: str = ""
    internet_message_id: str = ""
    transport_headers: str = ""
    attachments: list[AttachmentData] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def is_mail(self) -> bool:
        return self.message_class.startswith("IPM.Note")

    @property
    def is_smime(self) -> bool:
        return "SMIME" in self.message_class.upper()

    @property
    def is_encrypted(self) -> bool:
        # IPM.Note.SMIME is encrypted; IPM.Note.SMIME.MultipartSigned is signed.
        mc = self.message_class.upper()
        return mc == "IPM.NOTE.SMIME" or mc.endswith(".SMIME")


def _get_prop(accessor, tag: str):
    try:
        val = accessor.GetProperty(tag)
        return val if val not in (None, "") else None
    except Exception:
        return None


@contextmanager
def outlook_session(store_name: str = ""):
    """Connect to the running classic Outlook and yield (namespace, store).

    Requires classic OUTLOOK.EXE to be running in the same user session.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        # Prefer attaching to the already-running classic Outlook (the charter's
        # premise) rather than launching one. GetActiveObject attaches to the
        # live instance; Dispatch is a fallback that attaches to the running
        # singleton if present. We never use DispatchEx (which can spawn a new
        # instance in an unexpected profile).
        app = None
        try:
            app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            try:
                app = win32com.client.Dispatch("Outlook.Application")
            except Exception as exc:  # pragma: no cover - needs live Outlook
                raise OutlookError(
                    "Could not connect to a running classic Outlook. Start "
                    "classic Outlook Desktop (OUTLOOK.EXE, not the new Outlook) "
                    "in this user session and try again."
                ) from exc
        try:
            namespace = app.GetNamespace("MAPI")
        except Exception as exc:  # pragma: no cover - needs live Outlook
            raise OutlookError(f"Connected to Outlook but MAPI was unavailable: {exc}") from exc

        store = None
        if store_name:
            for i in range(namespace.Stores.Count):
                candidate = namespace.Stores.Item(i + 1)
                if candidate.DisplayName == store_name:
                    store = candidate
                    break
            if store is None:
                raise OutlookError(f"Store not found: {store_name!r}")
        else:
            store = namespace.DefaultStore
        yield namespace, store
    finally:
        pythoncom.CoUninitialize()


def _walk(folder, prefix: list[str]) -> Iterator[tuple[list[str], object]]:
    segments = prefix + [folder.Name]
    yield segments, folder
    subfolders = folder.Folders
    for i in range(subfolders.Count):
        yield from _walk(subfolders.Item(i + 1), segments)


def list_folders(store) -> Iterator[tuple[list[str], object]]:
    """Yield (segments, folder_com) for every folder under the store root.

    ``segments`` is the list of raw Outlook folder names from the root down; a
    name that itself contains '/' stays a single segment, so callers never
    mistake it for a path separator.
    """
    root = store.GetRootFolder()
    subfolders = root.Folders
    for i in range(subfolders.Count):
        yield from _walk(subfolders.Item(i + 1), [])


def find_folder(store, wanted_path: str):
    """Resolve an export-config folder path (escaped form) to a COM folder."""
    from .paths import parse_config_path

    target = parse_config_path(wanted_path)
    for segments, folder in list_folders(store):
        if segments == target:
            return folder
    return None


def _resolve_sender_smtp(mail) -> str:
    smtp = _get_prop(mail.PropertyAccessor, PR_SENDER_SMTP_ADDRESS)
    if smtp:
        return smtp
    smtp = _get_prop(mail.PropertyAccessor, PR_SENT_REPRESENTING_SMTP)
    if smtp:
        return smtp
    try:
        sender = mail.Sender
        if sender is not None:
            try:
                eu = sender.GetExchangeUser()
                if eu and eu.PrimarySmtpAddress:
                    return eu.PrimarySmtpAddress
            except Exception:
                pass
            smtp = _get_prop(sender.PropertyAccessor, PR_SMTP_ADDRESS)
            if smtp:
                return smtp
    except Exception:
        pass
    addr = getattr(mail, "SenderEmailAddress", "") or ""
    # EX-type (/O=...) addresses are not usable SMTP; prefer empty over garbage.
    return "" if addr.startswith("/") else addr


def _resolve_recipient_smtp(recipient) -> str:
    smtp = _get_prop(recipient.PropertyAccessor, PR_SMTP_ADDRESS)
    if smtp:
        return smtp
    try:
        ae = recipient.AddressEntry
        if ae is not None:
            try:
                eu = ae.GetExchangeUser()
                if eu and eu.PrimarySmtpAddress:
                    return eu.PrimarySmtpAddress
            except Exception:
                pass
            smtp = _get_prop(ae.PropertyAccessor, PR_SMTP_ADDRESS)
            if smtp:
                return smtp
    except Exception:
        pass
    addr = getattr(recipient, "Address", "") or ""
    if addr and not addr.startswith("/"):
        return addr
    return getattr(recipient, "Name", "") or ""


def _to_dt(value) -> datetime | None:
    """Convert a pywintypes datetime to a timezone-aware datetime.

    pywin32 returns a datetime carrying tzinfo (usually UTC). We preserve the
    instant and offset rather than collapsing to naive local time, so the .eml
    Date header keeps the original moment.
    """
    if value is None:
        return None
    try:
        tz = getattr(value, "tzinfo", None) or timezone.utc
        return datetime.fromtimestamp(value.timestamp(), tz=tz)
    except Exception:
        try:
            return datetime(value.year, value.month, value.day,
                            value.hour, value.minute, value.second,
                            tzinfo=timezone.utc)
        except Exception:
            return None


def _extract_attachments(mail) -> list[AttachmentData]:
    out: list[AttachmentData] = []
    attachments = mail.Attachments
    count = attachments.Count
    if count == 0:
        return out
    tmpdir = tempfile.mkdtemp(prefix="oma_att_")
    try:
        for i in range(1, count + 1):
            att = attachments.Item(i)
            filename = getattr(att, "FileName", None) or f"attachment_{i}.bin"
            content_id = _get_prop(att.PropertyAccessor, PR_ATTACH_CONTENT_ID)
            hidden = _get_prop(att.PropertyAccessor, PR_ATTACHMENT_HIDDEN)
            inline = bool(content_id) or bool(hidden)
            safe = "".join(c for c in filename if c not in '\\/:*?"<>|') or f"att_{i}.bin"
            tmp_path = os.path.join(tmpdir, f"{i}_{safe}")
            try:
                att.SaveAsFile(tmp_path)
                with open(tmp_path, "rb") as fh:
                    data = fh.read()
            except Exception as exc:
                raise OutlookError(
                    f"attachment {filename!r} could not be read: {exc}"
                ) from exc
            out.append(AttachmentData(
                filename=filename, data=data,
                content_id=content_id, inline=inline,
            ))
    finally:
        for fn in os.listdir(tmpdir):
            try:
                os.remove(os.path.join(tmpdir, fn))
            except OSError:
                pass
        try:
            os.rmdir(tmpdir)
        except OSError:
            pass
    return out


def extract_message(mail) -> MessageData:
    """Detach one Outlook item into plain data. May raise OutlookError."""
    pa = mail.PropertyAccessor
    message_class = getattr(mail, "MessageClass", "") or ""

    to, cc, bcc = [], [], []
    warnings: list[str] = []
    try:
        recipients = mail.Recipients
        total = recipients.Count
        for i in range(1, total + 1):
            try:
                r = recipients.Item(i)
                smtp = _resolve_recipient_smtp(r)
                display = getattr(r, "Name", "") or ""
                # formataddr quotes display names containing commas/specials so a
                # name like "Doe, Jane" does not split into two addresses.
                formatted = formataddr((display, smtp)) if smtp else (display or "")
                rtype = getattr(r, "Type", 1)
                if rtype == 2:
                    cc.append(formatted)
                elif rtype == 3:
                    bcc.append(formatted)
                else:
                    to.append(formatted)
            except Exception as exc:
                warnings.append(f"recipient-{i}-unreadable: {exc}")
    except Exception as exc:
        warnings.append(f"recipients-enumeration-failed: {exc}")

    if mail.MessageClass and "SMIME" in str(message_class).upper():
        warnings.append(f"smime-item:{message_class}")

    return MessageData(
        entry_id=getattr(mail, "EntryID", ""),
        message_class=message_class,
        subject=getattr(mail, "Subject", "") or "",
        sender_name=getattr(mail, "SenderName", "") or "",
        sender_smtp=_resolve_sender_smtp(mail),
        to=to, cc=cc, bcc=bcc,
        warnings=warnings,
        sent_on=_to_dt(getattr(mail, "SentOn", None)),
        received_time=_to_dt(getattr(mail, "ReceivedTime", None)),
        html_body=getattr(mail, "HTMLBody", "") or "",
        text_body=getattr(mail, "Body", "") or "",
        internet_message_id=_get_prop(pa, PR_INTERNET_MESSAGE_ID) or "",
        transport_headers=_get_prop(pa, PR_TRANSPORT_MESSAGE_HEADERS) or "",
        attachments=_extract_attachments(mail),
    )


def iter_folder_items(folder, date_from: date | None = None,
                      date_to: date | None = None) -> Iterator[object]:
    """Yield mail COM items in a folder, optionally filtered by received date.

    Yields the live COM items newest-first; the caller extracts and releases
    them one at a time to keep memory bounded on large folders.
    """
    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    count = items.Count
    for i in range(1, count + 1):
        item = items.Item(i)
        if date_from or date_to:
            rt = _to_dt(getattr(item, "ReceivedTime", None))
            if rt is not None:
                d = rt.date()
                if date_from and d < date_from:
                    continue
                if date_to and d > date_to:
                    continue
        yield item
