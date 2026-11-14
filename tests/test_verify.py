import io
import urllib.error

import pytest

from outlook_mail_archiver.verify import (
    ArchiveClient, ArchiveError, ArchiveIndex, VerifyResult,
    confirmed_max_date, verify,
)


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


def _client_index(present_ids, subject_counts, account_id=1):
    # Inject a full ArchiveIndex (Message-IDs + subject multiset) for the
    # subject-fallback path.
    def provider():
        return ArchiveIndex(message_ids=set(present_ids),
                            subject_counts=dict(subject_counts))
    return ArchiveClient("http://x", "key", account_id, index_provider=provider)


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


# --- subject fallback (Message-ID re-derived by the archive) -----------------

def test_subject_fallback_confirms_rederived_message_id():
    # Archive holds the message but under a different Message-ID; exact match
    # fails, subject match (one copy available) confirms it.
    man = _manifest([{"message_id": "<x@geopod>", "entry_id": "1",
                      "subject": "UPS Power Restored 12:52"}])
    r = verify(man, _client_index(set(), {"ups power restored 12:52": 1}))
    assert r.present == 0 and r.present_fallback == 1 and r.confirmed == 1
    assert not r.missing and not r.ambiguous and r.verified_success


def test_subject_fallback_normalises_whitespace_and_case():
    man = _manifest([{"message_id": "<x@geopod>", "entry_id": "1",
                      "subject": "  Hello   WORLD "}])
    r = verify(man, _client_index(set(), {"hello world": 1}))
    assert r.present_fallback == 1 and r.verified_success


def test_subject_absent_is_missing_not_fallback():
    man = _manifest([{"message_id": "<x@geopod>", "entry_id": "1",
                      "subject": "Only In Export"}])
    r = verify(man, _client_index(set(), {"something else": 2}))
    assert r.present_fallback == 0 and r.missing == ["<x@geopod>"]
    assert not r.verified_success


def test_subject_slot_exhausted_is_ambiguous():
    # Two exported messages share a subject, but the archive holds only one copy
    # of it and neither matches by Message-ID: one is covered, one is ambiguous.
    man = _manifest([{"message_id": "<a@geopod>", "entry_id": "1", "subject": "Dup"},
                     {"message_id": "<b@geopod>", "entry_id": "2", "subject": "Dup"}])
    r = verify(man, _client_index(set(), {"dup": 1}))
    assert r.present_fallback == 1 and len(r.ambiguous) == 1
    assert not r.verified_success


def test_exact_match_reserves_subject_slot_before_fallback():
    # One message matches by Message-ID (consuming the single archive copy of
    # 'Shared'); a second with the same subject must NOT be confirmed by fallback.
    man = _manifest([{"message_id": "<real@host>", "entry_id": "1", "subject": "Shared"},
                     {"message_id": "<x@geopod>", "entry_id": "2", "subject": "Shared"}])
    r = verify(man, _client_index({"real@host"}, {"shared": 1}))
    assert r.present == 1 and r.present_fallback == 0
    assert len(r.ambiguous) == 1 and not r.verified_success


def test_fallback_can_be_disabled():
    man = _manifest([{"message_id": "<x@geopod>", "entry_id": "1", "subject": "S"}])
    r = verify(man, _client_index(set(), {"s": 1}), fallback_subject=False)
    assert r.present_fallback == 0 and r.ambiguous == ["<x@geopod>"]


def test_ids_only_client_leaves_fallback_inert():
    # Legacy ids_provider path carries no subject multiset, so behaviour matches
    # the pre-fallback exact-only verify.
    man = _manifest([{"message_id": "<a@x>", "entry_id": "1", "subject": "S"}])
    r = verify(man, _client(set()))
    assert r.present_fallback == 0 and r.missing == ["<a@x>"]


# --- confirmed_max_date (incremental watermark advance) ----------------------

def _m(mid, date, status="exported"):
    return {"message_id": mid, "date": date, "status": status}


def test_confirmed_max_date_uses_newest_present():
    man = {"messages": [_m("<a>", "2026-01-01T00:00:00"),
                        _m("<b>", "2026-03-01T00:00:00")]}
    r = VerifyResult()  # nothing missing
    assert confirmed_max_date(man, r).isoformat() == "2026-03-01T00:00:00"


def test_confirmed_max_date_excludes_strict_missing():
    # The newest message is strictly missing -> watermark must not advance to it.
    man = {"messages": [_m("<a>", "2026-01-01T00:00:00"),
                        _m("<late>", "2026-06-01T00:00:00")]}
    r = VerifyResult(missing=["<late>"])
    assert confirmed_max_date(man, r).isoformat() == "2026-01-01T00:00:00"


def test_confirmed_max_date_counts_ambiguous_and_conflicts_as_present():
    # ambiguous is NOT excluded (it is in the archive) -> it can set the mark.
    man = {"messages": [_m("<a>", "2026-01-01T00:00:00"),
                        _m("<amb>", "2026-09-01T00:00:00")]}
    r = VerifyResult(ambiguous=["<amb>"])
    assert confirmed_max_date(man, r).isoformat() == "2026-09-01T00:00:00"


def test_confirmed_max_date_none_when_all_missing_or_dateless():
    man = {"messages": [_m("<a>", None), _m("<b>", "bad-date")]}
    assert confirmed_max_date(man, VerifyResult()) is None


# --- HTTP layer: rate-limit (429) resilience ---------------------------------

def _http_error(code, body=b"", retry_after=None):
    headers = {"Retry-After": retry_after} if retry_after is not None else {}
    return urllib.error.HTTPError("http://x/p", code, "err", headers, io.BytesIO(body))


class _FakeResp:
    def __init__(self, payload: bytes):
        self._payload = payload

    def read(self):
        return self._payload

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def _patch_urlopen(monkeypatch, outcomes):
    """outcomes: list of either an Exception to raise or bytes to return."""
    seq = list(outcomes)

    def fake_urlopen(req, timeout=None):
        item = seq.pop(0)
        if isinstance(item, Exception):
            raise item
        return _FakeResp(item)

    monkeypatch.setattr("urllib.request.urlopen", fake_urlopen)


def test_get_retries_on_429_then_succeeds(monkeypatch):
    slept = []
    client = ArchiveClient("http://x", "key", 1, throttle=0,
                           sleep=lambda s: slept.append(s))
    _patch_urlopen(monkeypatch, [
        _http_error(429), _http_error(429), b'{"ok": true}',
    ])
    assert client._get("/p") == {"ok": True}
    assert len(slept) == 2  # two backoff waits before the success


def test_get_honours_retry_after(monkeypatch):
    slept = []
    client = ArchiveClient("http://x", "key", 1, throttle=0,
                           sleep=lambda s: slept.append(s))
    _patch_urlopen(monkeypatch, [_http_error(429, retry_after="5"), b'{"ok": 1}'])
    client._get("/p")
    assert slept == [5.0]  # exact Retry-After delta honoured, not backoff


def test_get_gives_up_after_max_retries(monkeypatch):
    client = ArchiveClient("http://x", "key", 1, throttle=0,
                           max_retries=2, sleep=lambda s: None)
    _patch_urlopen(monkeypatch, [_http_error(429)] * 5)
    with pytest.raises(ArchiveError):
        client._get("/p")


def test_get_does_not_retry_on_404(monkeypatch):
    calls = {"n": 0}

    def fake_urlopen(req, timeout=None):
        calls["n"] += 1
        raise _http_error(404)

    monkeypatch.setattr("urllib.request.urlopen", fake_urlopen)
    client = ArchiveClient("http://x", "key", 1, throttle=0, sleep=lambda s: None)
    with pytest.raises(ArchiveError):
        client._get("/p")
    assert calls["n"] == 1  # non-transient: no retry
