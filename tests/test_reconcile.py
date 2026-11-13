from outlook_mail_archiver.import_driver import parse_import_output
from outlook_mail_archiver.reconcile import ImportCounts, ReconcileResult, reconcile


def _manifest(exported, export_failed=0):
    messages = [{"status": "exported"} for _ in range(exported)]
    return {"messages": messages, "totals": {"failed": export_failed}}


def test_counts_consistent():
    r = reconcile(_manifest(10), ImportCounts(imported=8, duplicates=2, failed=0))
    assert r.balanced and r.counts_consistent


def test_unresolved_blocks_consistency():
    r = reconcile(_manifest(10), ImportCounts(imported=5, duplicates=2, failed=0))
    assert r.unresolved == 3
    assert not r.balanced and not r.counts_consistent


def test_import_failures_block_consistency():
    r = reconcile(_manifest(10), ImportCounts(imported=8, duplicates=0, failed=2))
    assert r.balanced and not r.counts_consistent


def test_export_failures_block_consistency():
    r = reconcile(_manifest(10, export_failed=1), ImportCounts(imported=10, failed=0))
    assert not r.counts_consistent


def test_empty_export_and_empty_log_is_not_consistent():
    # Regression: 0 == 0 + 0 + 0 must NOT report success of any kind.
    r = reconcile(_manifest(0), ImportCounts())
    assert not r.counts_consistent


def test_parse_mail_archiver_output():
    text = (
        "=== Import Results ===\n"
        "Total Emails: 15234\n"
        "Imported Successfully: 15120\n"
        "Failed: 42\n"
        "Skipped (duplicates): 44\n"
    )
    counts = parse_import_output(text)
    assert counts.imported == 15120
    assert counts.failed == 42
    assert counts.duplicates == 44
