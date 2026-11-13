"""Rebuild a standard RFC 822 / MIME ``.eml`` from a detached ``MessageData``.

Pure: no COM, no I/O — takes data, returns bytes, so it is unit-testable with
synthetic messages. Ordinary mail (text, HTML, attachments, inline CID images)
is reproduced faithfully. A rebuilt message is not a byte-for-byte original and
makes no claim to carry a valid original signature; see README fidelity notes.
"""
from __future__ import annotations

import mimetypes
from email.message import EmailMessage
from email.policy import SMTP
from email.utils import format_datetime, formataddr

from .identity import message_identity
from .outlook_com import MessageData


def _split_type(filename: str, default: str = "application/octet-stream") -> tuple[str, str]:
    ctype, _ = mimetypes.guess_type(filename)
    if not ctype:
        ctype = default
    maintype, _, subtype = ctype.partition("/")
    return maintype or "application", subtype or "octet-stream"


def _clean_cid(content_id: str) -> str:
    return (content_id or "").strip().lstrip("<").rstrip(">")


def build_eml(md: MessageData) -> tuple[bytes, dict]:
    """Return (eml_bytes, info).

    ``info`` carries fields the manifest records: message_id, whether it was
    generated, and a short list of fidelity notes for this message.
    """
    msg = EmailMessage()
    notes: list[str] = []

    message_id, generated = message_identity(md)
    msg["Message-ID"] = message_id

    from_addr = md.sender_smtp or ""
    if from_addr:
        # formataddr quotes a display name containing commas or specials.
        msg["From"] = formataddr((md.sender_name or "", from_addr))
    elif md.sender_name:
        msg["From"] = formataddr((md.sender_name, ""))
        notes.append("sender-smtp-unresolved")

    if md.to:
        msg["To"] = ", ".join(md.to)
    if md.cc:
        msg["Cc"] = ", ".join(md.cc)
    if md.bcc:
        msg["Bcc"] = ", ".join(md.bcc)

    dt = md.sent_on or md.received_time
    if dt is not None:
        msg["Date"] = format_datetime(dt)
    msg["Subject"] = md.subject or ""

    if not md.is_mail:
        notes.append(f"non-mail-item:{md.message_class}")
    if md.is_smime:
        # Signed/encrypted mail is rebuilt from readable parts; the original
        # signature is not reproduced and encrypted content is not decrypted.
        notes.append(
            "encrypted-smime:content-not-decrypted" if md.is_encrypted
            else "signed-smime:signature-not-preserved"
        )
    notes.extend(md.warnings)

    # Body: text + optional HTML alternative.
    text = md.text_body or ""
    html = md.html_body or ""
    msg.set_content(text if text else " ")
    if html:
        msg.add_alternative(html, subtype="html")

    inline = [a for a in md.attachments if a.inline and a.content_id]
    regular = [a for a in md.attachments if not (a.inline and a.content_id)]

    # Inline CID images attach to the HTML alternative as multipart/related so
    # they render in place rather than as loose attachments.
    if inline:
        if html and len(msg.get_payload()) >= 2:
            html_part = msg.get_payload()[1]
            for att in inline:
                maintype, subtype = _split_type(att.filename, "application/octet-stream")
                try:
                    html_part.add_related(
                        att.data, maintype=maintype, subtype=subtype,
                        cid=f"<{_clean_cid(att.content_id)}>",
                        filename=att.filename, disposition="inline",
                    )
                except Exception:
                    regular.append(att)
                    notes.append(f"inline-fallback:{att.filename}")
        else:
            # No HTML to anchor them to — carry as regular attachments.
            regular.extend(inline)
            notes.append("inline-without-html")

    for att in regular:
        maintype, subtype = _split_type(att.filename)
        msg.add_attachment(
            att.data, maintype=maintype, subtype=subtype, filename=att.filename,
        )

    return msg.as_bytes(policy=SMTP), {
        "message_id": message_id,
        "generated_id": generated,
        "notes": notes,
    }
