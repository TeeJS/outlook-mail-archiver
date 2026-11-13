"""Message identity and dedupe policy.

Prefer the original ``Message-ID`` header. When it is missing or malformed,
synthesise a deterministic substitute from the message's own content so that
re-exporting the same message yields the same id (idempotent) — and so that two
genuinely different messages are very unlikely to collide. The substitute does
NOT depend on the folder, so moving a message between folders does not change
its id.
"""
from __future__ import annotations

import hashlib
import re

from .outlook_com import MessageData

_MSGID_RE = re.compile(r"^<[^<>@\s]+@[^<>@\s]+>$")


def _normalize_existing(raw: str) -> str | None:
    raw = (raw or "").strip()
    if not raw:
        return None
    if not raw.startswith("<"):
        raw = "<" + raw
    if not raw.endswith(">"):
        raw = raw + ">"
    return raw if _MSGID_RE.match(raw) else None


def body_sha256(md: MessageData) -> str:
    """Stable hash of the message body (html preferred, else plain)."""
    body = md.html_body or md.text_body or ""
    return hashlib.sha256(body.encode("utf-8", "replace")).hexdigest()


def _attachment_fingerprint(md: MessageData) -> str:
    # Hash the attachment BYTES, not just name+length: two attachments of equal
    # length but different content must not produce the same message identity.
    parts = [
        f"{a.filename}:{hashlib.sha256(a.data).hexdigest()}"
        for a in md.attachments
    ]
    return "|".join(sorted(parts))


def message_identity(md: MessageData) -> tuple[str, bool]:
    """Return (message_id, generated).

    ``generated`` is True when the id was synthesised because the source lacked
    a usable ``Message-ID``.
    """
    existing = _normalize_existing(md.internet_message_id)
    if existing:
        return existing, False

    dt = md.sent_on or md.received_time
    stamp = dt.isoformat() if dt else ""
    recipients = ",".join(sorted(r.lower() for r in (md.to + md.cc)))
    material = "\x1f".join([
        stamp,
        (md.sender_smtp or "").lower(),
        recipients,
        " ".join((md.subject or "").split()).lower(),
        body_sha256(md),
        _attachment_fingerprint(md),
    ])
    digest = hashlib.sha256(material.encode("utf-8", "replace")).hexdigest()[:32]
    return f"<oma-{digest}@outlook-mail-archiver.local>", True
