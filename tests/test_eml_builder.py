import email
from email.policy import default as default_policy

from conftest import make_message, png_1x1

from outlook_mail_archiver.outlook_com import AttachmentData
from outlook_mail_archiver.eml_builder import build_eml


def _parse(md):
    data, info = build_eml(md)
    msg = email.message_from_bytes(data, policy=default_policy)
    return msg, info, data


def test_basic_text_html_alternative():
    msg, info, _ = _parse(make_message())
    assert msg["Subject"] == "Hello world"
    assert msg.get_content_type() == "multipart/alternative"
    types = {p.get_content_type() for p in msg.walk()}
    assert "text/plain" in types and "text/html" in types


def test_comma_display_name_is_single_address():
    # Regression: "Doe, Jane" must not split into two From addresses.
    msg, _, _ = _parse(make_message(sender_name="Doe, Jane",
                                    sender_smtp="jane@example.com"))
    addrs = msg["From"].addresses
    assert len(addrs) == 1
    assert addrs[0].addr_spec == "jane@example.com"


def test_date_header_is_timezone_aware():
    from datetime import datetime, timezone
    dt = datetime(2026, 5, 1, 9, 30, tzinfo=timezone.utc)
    msg, _, _ = _parse(make_message(sent_on=dt, received_time=None))
    assert msg["Date"] is not None
    assert "0000" in msg["Date"] or "+00:00" in str(msg["Date"])


def test_message_id_preserved():
    msg, info, _ = _parse(make_message(internet_message_id="<keep@me.com>"))
    assert msg["Message-ID"] == "<keep@me.com>"
    assert info["generated_id"] is False


def test_regular_attachment():
    md = make_message(attachments=[AttachmentData("doc.txt", b"hello", inline=False)])
    msg, _, _ = _parse(md)
    names = [p.get_filename() for p in msg.walk() if p.get_filename()]
    assert "doc.txt" in names


def test_inline_cid_becomes_related_inline():
    md = make_message(attachments=[
        AttachmentData("img.png", png_1x1(), content_id="image001", inline=True)
    ])
    msg, _, _ = _parse(md)
    related = [p for p in msg.walk() if p.get_content_type() == "multipart/related"]
    assert related, "inline image should produce a multipart/related"
    img = [p for p in msg.walk() if p.get_content_type() == "image/png"]
    assert img
    assert img[0]["Content-ID"] == "<image001>"
    assert "inline" in (img[0]["Content-Disposition"] or "")


def test_smime_signed_noted():
    msg, info, _ = _parse(make_message(message_class="IPM.Note.SMIME.MultipartSigned"))
    assert any("signed-smime" in n for n in info["notes"])
