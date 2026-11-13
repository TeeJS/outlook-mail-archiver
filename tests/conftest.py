"""Synthetic fixtures — no Outlook/COM required.

The COM imports in outlook_com live inside outlook_session(), so MessageData /
AttachmentData import cleanly without pywin32, and the pure modules (identity,
eml_builder, paths, reconcile, import_driver, and the export orchestration with
COM monkeypatched) are all testable here.
"""
from __future__ import annotations

import itertools
from datetime import datetime

import pytest

from outlook_mail_archiver.outlook_com import AttachmentData, MessageData

_entry_seq = itertools.count(1)


def make_message(**over) -> MessageData:
    # Unique EntryID per message unless the caller sets one, so tests don't
    # accidentally trip the physical-item dedup.
    base = dict(
        entry_id=f"E{next(_entry_seq)}",
        message_class="IPM.Note",
        subject="Hello world",
        sender_name="Alice Example",
        sender_smtp="alice@example.com",
        to=["Bob <bob@example.com>"],
        cc=[],
        bcc=[],
        sent_on=datetime(2026, 5, 1, 9, 30, 0),
        received_time=datetime(2026, 5, 1, 9, 31, 0),
        html_body="<p>Hello <b>world</b></p>",
        text_body="Hello world",
        internet_message_id="<abc123@mail.example.com>",
        transport_headers="",
        attachments=[],
    )
    base.update(over)
    return MessageData(**base)


@pytest.fixture
def message():
    return make_message


def png_1x1() -> bytes:
    # Minimal 1x1 PNG.
    return bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
        "0000000a49444154789c6360000002000154a24f0b0000000049454e44ae426082"
    )
