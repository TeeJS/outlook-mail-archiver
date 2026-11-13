"""Per-message destination verification.

Aggregate import counts can balance while hiding a wrong or mismatched message,
so the real acceptance evidence is confirming, per message, that the archive
holds it. The Mail-Archiver REST API's search result does NOT carry the
``Message-ID`` (only the message *detail* does), so we build the set of
Message-IDs the archive holds for the target account — list the account's
messages, read each one's detail for its ``messageId`` — and check every
exported Message-ID against that set.

Safeguards (per review):
  * Only Message-IDs whose detail reports the target ``accountId`` are counted,
    so a cross-account or unattributed hit is never accepted.
  * A Message-ID conflict (two distinct source items share one Message-ID) can
    never be fully verified: the archive dedups to one, so such records are
    reported UNRESOLVED and block success.
  * Any HTTP/parse error while indexing is a global error → not verified.

The HTTP layer is injected (``ids_provider``) so the matching logic is
unit-tested; the live index path is exercised against a real instance.
"""
from __future__ import annotations

import json
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from dataclasses import dataclass, field


def _norm(message_id: str) -> str:
    """Normalise a Message-ID for comparison: trim and drop angle brackets.

    The archive stores the id without ``<>``; the manifest keeps them.
    """
    return (message_id or "").strip().lstrip("<").rstrip(">").strip()


@dataclass
class VerifyResult:
    checked: int = 0
    present: int = 0
    missing: list[str] = field(default_factory=list)
    unresolved_conflicts: int = 0
    errors: list[dict] = field(default_factory=list)
    source_failed: int = 0
    source_errors: int = 0
    source_degraded: int = 0
    skipped_unsupported: int = 0

    @property
    def verified_success(self) -> bool:
        return (
            self.checked > 0
            and self.present == self.checked
            and not self.missing
            and not self.errors
            and self.unresolved_conflicts == 0
            and self.source_failed == 0
            and self.source_errors == 0
            and self.source_degraded == 0
        )

    def as_dict(self) -> dict:
        return {
            "checked": self.checked,
            "present": self.present,
            "missing": self.missing,
            "unresolved_conflicts": self.unresolved_conflicts,
            "errors": self.errors,
            "source_failed": self.source_failed,
            "source_errors": self.source_errors,
            "source_degraded": self.source_degraded,
            "skipped_unsupported": self.skipped_unsupported,
            "verified_success": self.verified_success,
        }


class ArchiveError(Exception):
    pass


class ArchiveClient:
    """Reads the set of Message-IDs the archive holds for one account.

    Pass ``ids_provider`` (a callable returning a set of normalised Message-IDs)
    to test without HTTP.
    """

    def __init__(self, base_url: str, api_key: str, account_id: int, *,
                 ids_provider=None, timeout: float = 30.0, page_size: int = 100):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.account_id = account_id
        self.timeout = timeout
        self.page_size = page_size
        self._ids_provider = ids_provider

    def account_message_ids(self, from_date=None) -> set[str]:
        """Return the normalised Message-IDs the archive holds for the account.

        ``from_date`` (a date) limits the query to messages on/after that date,
        so an incremental run only indexes the recent window.
        """
        if self._ids_provider is not None:
            return self._ids_provider()
        return self._http_index(from_date)

    def _get(self, path: str):  # pragma: no cover - live only
        req = urllib.request.Request(self.base_url + path, headers={
            "Authorization": f"Bearer {self.api_key}",
            "Accept": "application/json",
        })
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except (urllib.error.URLError, json.JSONDecodeError, ValueError) as exc:
            raise ArchiveError(f"archive request failed ({path}): {exc}") from exc

    def _http_index(self, from_date=None) -> set[str]:  # pragma: no cover - live only
        ids: set[str] = set()
        page = 1
        while True:
            params = {
                "accountId": self.account_id, "pageSize": self.page_size, "page": page,
            }
            if from_date is not None:
                params["from"] = from_date.isoformat()
            q = urllib.parse.urlencode(params)
            data = self._get(f"/api/v1/emails?{q}")
            if isinstance(data, list):
                items, total_pages = data, 1
            elif isinstance(data, dict):
                items = data.get("items", []) or []
                total_pages = int(data.get("totalPages", 1) or 1)
            else:
                raise ArchiveError(f"unexpected list response: {type(data).__name__}")
            for it in items:
                eid = it.get("id") if isinstance(it, dict) else None
                if eid is None:
                    continue
                detail = self._get(f"/api/v1/emails/{eid}")
                if not isinstance(detail, dict):
                    continue
                try:
                    same_account = int(detail.get("accountId")) == int(self.account_id)
                except (TypeError, ValueError):
                    same_account = False
                mid = _norm(str(detail.get("messageId", "")))
                if same_account and mid:
                    ids.add(mid)
            if page >= total_pages:
                break
            page += 1
        return ids


def _as_int(value, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def verify(manifest: dict, client: ArchiveClient) -> VerifyResult:
    result = VerifyResult()
    totals = manifest.get("totals", {}) or {}
    result.source_failed = _as_int(totals.get("failed", 0))
    result.source_degraded = _as_int(totals.get("degraded", 0))
    result.skipped_unsupported = _as_int(totals.get("skipped_unsupported", 0))
    result.source_errors = len(manifest.get("errors", []) or [])

    groups: dict[str, list[dict]] = defaultdict(list)
    for m in manifest.get("messages", []):
        if m.get("status") == "exported":
            groups[m["message_id"]].append(m)

    try:
        present_ids = client.account_message_ids()
    except ArchiveError as exc:
        # Could not index the archive — nothing can be confirmed.
        result.checked = sum(len(v) for v in groups.values())
        result.errors.append({"error": str(exc)})
        return result

    for message_id, records in groups.items():
        result.checked += len(records)
        found = _norm(message_id) in present_ids
        if len(records) == 1:
            if found:
                result.present += 1
            else:
                result.missing.append(message_id)
        else:
            # Conflict: the archive holds at most one for this Message-ID.
            if found:
                result.present += 1
                result.unresolved_conflicts += len(records) - 1
            else:
                result.missing.append(message_id)
                result.unresolved_conflicts += len(records)
    return result
