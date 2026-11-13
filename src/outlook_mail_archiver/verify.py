"""Per-message destination verification.

Aggregate import counts can balance while hiding a wrong or mismatched message,
so the real acceptance evidence is a per-Message-ID lookup against the
Mail-Archiver REST API: for each exported message, confirm the archive holds a
message whose ``messageId`` matches exactly AND that belongs to the target
account.

Safeguards (per review):
  * An exact ``messageId`` AND ``accountId`` match is required — not a text hit.
  * A Message-ID conflict (two distinct source items share one Message-ID) can
    never be fully verified by Message-ID alone: the archive dedups to one, so
    at most one of the pair is confirmable. Such records are reported UNRESOLVED
    and block success.
  * Any HTTP/search/parse error is UNRESOLVED (never silently "present").

The HTTP layer is injected (``finder``), so the matching logic is unit-tested;
the live endpoint shape must still be validated against a real instance before
it is relied on (charter pre-transfer verification).
"""
from __future__ import annotations

import json
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from dataclasses import dataclass, field


@dataclass
class VerifyResult:
    checked: int = 0
    present: int = 0
    missing: list[str] = field(default_factory=list)
    unresolved_conflicts: int = 0
    errors: list[dict] = field(default_factory=list)

    @property
    def verified_success(self) -> bool:
        return (
            self.checked > 0
            and self.present == self.checked
            and not self.missing
            and not self.errors
            and self.unresolved_conflicts == 0
        )

    def as_dict(self) -> dict:
        return {
            "checked": self.checked,
            "present": self.present,
            "missing": self.missing,
            "unresolved_conflicts": self.unresolved_conflicts,
            "errors": self.errors,
            "verified_success": self.verified_success,
        }


class ArchiveError(Exception):
    pass


class ArchiveClient:
    """Thin REST client. ``find`` returns candidate records for a Message-ID.

    Each candidate is a dict with at least ``messageId`` and ``accountId``.
    Pass a custom ``finder`` (message_id -> list[dict]) to test without HTTP.
    """

    def __init__(self, base_url: str, api_key: str, account_id: int, *, finder=None,
                 timeout: float = 30.0):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.account_id = account_id
        self.timeout = timeout
        self._finder = finder

    def find(self, message_id: str) -> list[dict]:
        if self._finder is not None:
            return self._finder(message_id)
        return self._http_find(message_id)

    def _http_find(self, message_id: str) -> list[dict]:  # pragma: no cover - live only
        bare = message_id.strip().lstrip("<").rstrip(">")
        query = urllib.parse.urlencode({
            "q": bare, "accountId": self.account_id, "pageSize": 50,
        })
        url = f"{self.base_url}/api/v1/emails?{query}"
        req = urllib.request.Request(url, headers={
            "Authorization": f"Bearer {self.api_key}",
            "Accept": "application/json",
        })
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except (urllib.error.URLError, json.JSONDecodeError, ValueError) as exc:
            raise ArchiveError(f"archive query failed: {exc}") from exc
        items = payload.get("items", payload if isinstance(payload, list) else [])
        return items if isinstance(items, list) else []


def _matches(candidate: dict, message_id: str, account_id: int) -> bool:
    cid = str(candidate.get("messageId", "")).strip()
    want = message_id.strip()
    if cid != want and cid != want.strip("<>"):
        return False
    acct = candidate.get("accountId")
    return acct is None or int(acct) == int(account_id)


def verify(manifest: dict, client: ArchiveClient) -> VerifyResult:
    result = VerifyResult()
    groups: dict[str, list[dict]] = defaultdict(list)
    for m in manifest.get("messages", []):
        if m.get("status") == "exported":
            groups[m["message_id"]].append(m)

    for message_id, records in groups.items():
        result.checked += len(records)
        try:
            candidates = client.find(message_id)
        except ArchiveError as exc:
            result.errors.append({"message_id": message_id, "error": str(exc)})
            continue

        found = any(_matches(c, message_id, client.account_id) for c in candidates)
        if len(records) == 1:
            if found:
                result.present += 1
            else:
                result.missing.append(message_id)
        else:
            # Conflict: the archive holds at most one for this Message-ID, so the
            # others cannot be confirmed by Message-ID alone.
            if found:
                result.present += 1
                result.unresolved_conflicts += len(records) - 1
            else:
                result.missing.append(message_id)
                result.unresolved_conflicts += len(records)
    return result
