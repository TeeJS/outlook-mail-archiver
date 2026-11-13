from outlook_mail_archiver.paths import (
    display_path, parse_config_path, sanitize_segment, zip_path,
)


def test_literal_slash_in_name_is_one_segment():
    assert parse_config_path(r"Qlik\/RMM\/SAP Alerts") == ["Qlik/RMM/SAP Alerts"]


def test_separator_vs_escaped_slash():
    assert parse_config_path(r"Inbox/SAP\/apps") == ["Inbox", "SAP/apps"]


def test_display_roundtrip():
    segs = ["Inbox", "SAP/apps"]
    assert parse_config_path(display_path(segs)) == segs


def test_zip_path_sanitizes_slash_in_name():
    # A folder literally named 'A/B' becomes ONE component (with a stable tag),
    # never two, and never bare 'A_B' (which could collide with a real 'A_B').
    mapped = zip_path(["A/B"])
    assert "/" not in mapped
    assert mapped.startswith("A_B~")
    assert zip_path(["A", "B"]) == "A/B"


def test_zip_path_distinct_and_deterministic():
    # 'A/B' and a literal 'A_B' map to DIFFERENT, STABLE paths regardless of
    # order — the order-dependent-collision bug must not recur.
    a = zip_path(["A/B"])
    b = zip_path(["A_B"])
    assert a != b
    assert a == zip_path(["A/B"])      # deterministic
    assert b == "A_B"                  # clean name is unchanged


def test_tilde_marker_is_reserved():
    # A literal name containing '~' gets its own suffix, so it can never equal
    # another name's generated '~<hash>' suffix.
    mapped = zip_path(["A_B~a1b2c3"])
    assert mapped != "A_B~a1b2c3"
    assert mapped.startswith("A_B~a1b2c3~")
    assert zip_path(["A/B"]) != mapped


def test_sanitize_trailing_dot_and_reserved():
    assert sanitize_segment("report.") == "report"
    assert sanitize_segment("COM1").upper() == "COM1_"
    assert sanitize_segment("a.b.c") == "a.b.c"  # interior dots preserved


def test_sanitize_all_separators_stays_nonempty():
    assert sanitize_segment("///") == "___"
    assert sanitize_segment("") == "_"
