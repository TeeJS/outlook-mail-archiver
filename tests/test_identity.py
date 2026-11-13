from conftest import make_message

from outlook_mail_archiver.outlook_com import AttachmentData
from outlook_mail_archiver.identity import message_identity


def test_existing_message_id_preserved():
    mid, generated = message_identity(make_message(internet_message_id="<x@y.com>"))
    assert mid == "<x@y.com>"
    assert generated is False


def test_malformed_id_is_regenerated():
    mid, generated = message_identity(make_message(internet_message_id="not-an-id"))
    assert generated is True
    assert mid.endswith("@outlook-mail-archiver.local>")


def test_generated_id_is_deterministic():
    a = message_identity(make_message(internet_message_id=""))
    b = message_identity(make_message(internet_message_id=""))
    assert a == b
    assert a[1] is True


def test_different_body_different_id():
    a, _ = message_identity(make_message(internet_message_id="", html_body="<p>one</p>"))
    b, _ = message_identity(make_message(internet_message_id="", html_body="<p>two</p>"))
    assert a != b


def test_equal_length_different_attachment_bytes_differ():
    # Regression: fingerprint must hash bytes, not just length.
    m1 = make_message(internet_message_id="",
                      attachments=[AttachmentData("a.bin", b"aa")])
    m2 = make_message(internet_message_id="",
                      attachments=[AttachmentData("a.bin", b"bb")])
    assert message_identity(m1)[0] != message_identity(m2)[0]
