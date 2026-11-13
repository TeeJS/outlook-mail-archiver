from datetime import datetime, timezone

from outlook_mail_archiver.state import advance_high_water_mark, load_state


def test_watermark_starts_empty(tmp_path):
    assert load_state(tmp_path, 3).high_water_mark is None


def test_watermark_advances_forward_only(tmp_path):
    d1 = datetime(2026, 1, 1, tzinfo=timezone.utc)
    advance_high_water_mark(tmp_path, 3, d1)
    assert load_state(tmp_path, 3).high_water_mark == d1

    # an older date must not move it backward
    advance_high_water_mark(tmp_path, 3, datetime(2025, 1, 1, tzinfo=timezone.utc))
    assert load_state(tmp_path, 3).high_water_mark == d1

    # a newer date advances it
    d2 = datetime(2026, 6, 1, tzinfo=timezone.utc)
    advance_high_water_mark(tmp_path, 3, d2)
    assert load_state(tmp_path, 3).high_water_mark == d2


def test_watermark_is_per_account(tmp_path):
    advance_high_water_mark(tmp_path, 3, datetime(2026, 1, 1, tzinfo=timezone.utc))
    assert load_state(tmp_path, 9).high_water_mark is None
