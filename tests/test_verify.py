from outlook_mail_archiver.verify import ArchiveClient, ArchiveError, verify


def _manifest(records, totals=None, errors=None):
    return {
        "messages": [dict(status="exported", **r) for r in records],
        "totals": totals or {},
        "errors": errors or [],
    }


def _client(responses, account_id=1):
    # responses: message_id -> list[candidate dict] or an Exception to raise
    def finder(mid):
        r = responses.get(mid, [])
        if isinstance(r, Exception):
            raise r
        return r
    return ArchiveClient("http://x", "key", account_id, finder=finder)


def test_all_present_is_verified():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"},
                     {"message_id": "<b@x>", "entry_id": "2"}])
    client = _client({
        "<a@x>": [{"messageId": "<a@x>", "accountId": 1}],
        "<b@x>": [{"messageId": "<b@x>", "accountId": 1}],
    })
    r = verify(man, client)
    assert r.checked == 2 and r.present == 2 and r.verified_success


def test_missing_blocks_verification():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}])
    r = verify(man, _client({"<a@x>": []}))
    assert r.missing == ["<a@x>"] and not r.verified_success


def test_wrong_account_is_not_a_match():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}])
    r = verify(man, _client({"<a@x>": [{"messageId": "<a@x>", "accountId": 999}]}))
    assert not r.verified_success and r.missing == ["<a@x>"]


def test_conflict_cannot_be_fully_verified():
    # Two distinct source items share one Message-ID; the archive holds one, so
    # the pair can never be fully verified -> blocks success.
    man = _manifest([{"message_id": "<dup@x>", "entry_id": "A"},
                     {"message_id": "<dup@x>", "entry_id": "B"}])
    r = verify(man, _client({"<dup@x>": [{"messageId": "<dup@x>", "accountId": 1}]}))
    assert r.unresolved_conflicts == 1
    assert not r.verified_success


def test_http_error_is_unresolved():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}])
    r = verify(man, _client({"<a@x>": ArchiveError("boom")}))
    assert r.errors and not r.verified_success


def test_candidate_without_accountid_is_not_a_match():
    # Regression: an unattributed candidate (no accountId) must NOT be accepted.
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}])
    r = verify(man, _client({"<a@x>": [{"messageId": "<a@x>"}]}))
    assert r.missing == ["<a@x>"] and not r.verified_success


def test_source_failure_blocks_verified_success():
    # Regression: a source export failure must not vanish behind a verified subset.
    man = _manifest(
        [{"message_id": "<a@x>", "entry_id": "1"}],
        totals={"failed": 1},
    )
    r = verify(man, _client({"<a@x>": [{"messageId": "<a@x>", "accountId": 1}]}))
    assert r.present == 1 and r.source_failed == 1
    assert not r.verified_success


def test_source_folder_error_blocks_verified_success():
    man = _manifest(
        [{"message_id": "<a@x>", "entry_id": "1"}],
        errors=[{"folder": "X", "reason": "folder-not-found"}],
    )
    r = verify(man, _client({"<a@x>": [{"messageId": "<a@x>", "accountId": 1}]}))
    assert r.source_errors == 1 and not r.verified_success


def test_malformed_accountid_does_not_crash():
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1"}])
    r = verify(man, _client({"<a@x>": [{"messageId": "<a@x>", "accountId": "oops"}]}))
    assert not r.verified_success and r.missing == ["<a@x>"]
