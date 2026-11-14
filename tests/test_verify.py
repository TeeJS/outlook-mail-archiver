import io
import urllib.error

import pytest

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
