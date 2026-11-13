from outlook_mail_archiver.verify import ArchiveClient, ArchiveError, verify


def _manifest(records, totals=None, errors=None):
    return {
        "messages": [dict(status="exported", **r) for r in records],
        "totals": totals or {},
        "errors": errors or [],
    }


def _client(present_ids, account_id=1):
    # present_ids: a set of normalised (no-bracket) Message-IDs the archive holds,
    # or an Exception to raise while indexing.
    def provider():
        if isinstance(present_ids, Exception):
            raise present_ids
        return set(present_ids)
    return ArchiveClient("http://x", "key", account_id, ids_provider=provider)


def test_all_present_is_verified():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"},
                     {"message_id": "<b@x>", "entry_id": "2"}])
    # archive stores them without angle brackets
    r = verify(man, _client({"a@x", "b@x"}))
    assert r.checked == 2 and r.present == 2 and r.verified_success


def test_missing_blocks_verification():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}])
    r = verify(man, _client(set()))
    assert r.missing == ["<a@x>"] and not r.verified_success


def test_conflict_cannot_be_fully_verified():
    # Two distinct source items share one Message-ID; the archive holds one, so
    # the pair can never be fully verified -> blocks success.
    man = _manifest([{"message_id": "<dup@x>", "entry_id": "A"},
                     {"message_id": "<dup@x>", "entry_id": "B"}])
    r = verify(man, _client({"dup@x"}))
    assert r.present == 1 and r.unresolved_conflicts == 1
    assert not r.verified_success


def test_index_error_is_unresolved():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}])
    r = verify(man, _client(ArchiveError("boom")))
    assert r.errors and r.checked == 1 and not r.verified_success


def test_source_failure_blocks_verified_success():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}], totals={"failed": 1})
    r = verify(man, _client({"a@x"}))
    assert r.present == 1 and r.source_failed == 1
    assert not r.verified_success


def test_source_folder_error_blocks_verified_success():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}],
                    errors=[{"folder": "X", "reason": "folder-not-found"}])
    r = verify(man, _client({"a@x"}))
    assert r.source_errors == 1 and not r.verified_success


def test_degraded_mail_blocks_verified_success():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}], totals={"degraded": 1})
    r = verify(man, _client({"a@x"}))
    assert r.source_degraded == 1 and not r.verified_success
