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

Subject fallback: some sources emit a Message-ID the archive re-derives on
import (SendGrid/Marketo/Freshservice ids, ids containing ``;``), so the
exact-id check cannot find a message that is in fact present. A reserved-slot
subject multiset backs a secondary match — every archive message can confirm at
most one exported message, so a subject hit is only claimed while an unclaimed
copy of that subject remains. Such matches are reported SEPARATELY
(``present_fallback``); a subject that is in the archive but has no unclaimed
copy left is ``ambiguous`` (cannot confirm THIS message) and still blocks
success, as does a subject the archive does not hold at all (``missing``).

The HTTP layer is injected (``index_provider`` / ``ids_provider``) so the
matching logic is unit-tested; the live index path is exercised against a real
instance.
"""
from __future__ import annotations

import json
import random
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from dataclasses import dataclass, field


def _norm(message_id: str) -> str:
    """Normalise a Message-ID for comparison: trim and drop angle brackets.

    The archive stores the id without ``<>``; the manifest keeps them.
    """
    return (message_id or "").strip().lstrip("<").rstrip(">").strip()


def _subj(subject: str) -> str:
    """Normalise a subject for the fallback match: collapse whitespace, lower.

    Some sources store a real ``Message-ID`` the archive re-derives on import
    (e.g. SendGrid/Marketo/Freshservice ids, ids containing ``;``), so the
    exact-id check cannot find them even though the message is present. Subject
    is the only field this archive exposes that is both populated and stable
    across import (``from`` is frequently blank; the date fields are stamped at
    import), so it backs the secondary match.
    """
    return " ".join((subject or "").split()).lower()


@dataclass
class ArchiveIndex:
    """What the archive holds for one account, for matching against the export.

    ``message_ids`` is the primary (exact) key. ``subject_counts`` is a multiset
    of normalised subjects, used only as a reserved-slot fallback so a message
    present under a re-derived Message-ID is not falsely reported missing.
    """
    message_ids: set = field(default_factory=set)
    subject_counts: dict = field(default_factory=dict)


@dataclass
class VerifyResult:
    checked: int = 0
    present: int = 0
    present_fallback: int = 0
    missing: list[str] = field(default_factory=list)
    ambiguous: list[str] = field(default_factory=list)
    unresolved_conflicts: int = 0
    errors: list[dict] = field(default_factory=list)
    source_failed: int = 0
    source_errors: int = 0
    source_degraded: int = 0
    skipped_unsupported: int = 0

    @property
    def confirmed(self) -> int:
        """Messages confirmed present: exact Message-ID plus subject fallback."""
        return self.present + self.present_fallback

    @property
    def verified_success(self) -> bool:
        return (
            self.checked > 0
            and self.confirmed == self.checked
            and not self.missing
            and not self.ambiguous
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
            "present_fallback": self.present_fallback,
            "confirmed": self.confirmed,
            "missing": self.missing,
            "ambiguous": self.ambiguous,
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
                 ids_provider=None, index_provider=None, timeout: float = 30.0,
                 page_size: int = 100, throttle: float = 0.03, max_retries: int = 6,
                 backoff_cap: float = 30.0, sleep=time.sleep):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.account_id = account_id
        self.timeout = timeout
        self.page_size = page_size
        self._ids_provider = ids_provider
        self._index_provider = index_provider
        # Rate-limit resilience: a light pause between requests keeps a large
        # index (tens of thousands of detail GETs) under the archive's limiter,
        # and 429/503 are retried with exponential backoff honouring Retry-After.
        self.throttle = throttle
        self.max_retries = max_retries
        self.backoff_cap = backoff_cap
        self._sleep = sleep

    def account_index(self, from_date=None) -> ArchiveIndex:
        """Return the archive's ``ArchiveIndex`` (Message-IDs + subject multiset).

        ``from_date`` (a date) limits the query to messages on/after that date,
        so an incremental run only indexes the recent window.
        """
        if self._index_provider is not None:
            return self._index_provider()
        if self._ids_provider is not None:
            # Legacy id-only injection: no subject multiset, so fallback is inert.
            return ArchiveIndex(message_ids=set(self._ids_provider()),
                                subject_counts={})
        return self._http_index(from_date)

    def account_message_ids(self, from_date=None) -> set[str]:
        """Back-compat: just the Message-ID set from :meth:`account_index`."""
        return self.account_index(from_date).message_ids

    @staticmethod
    def _retry_after_seconds(exc: "urllib.error.HTTPError") -> float | None:
        """Parse a Retry-After header (delta-seconds only) if present."""
        try:
            value = exc.headers.get("Retry-After") if exc.headers else None
        except Exception:
            value = None
        if not value:
            return None
        try:
            return max(0.0, float(value.strip()))
        except (TypeError, ValueError):
            return None  # HTTP-date form is ignored; backoff is used instead

    def _get(self, path: str):  # pragma: no cover - live only
        req = urllib.request.Request(self.base_url + path, headers={
            "Authorization": f"Bearer {self.api_key}",
            "Accept": "application/json",
        })
        if self.throttle:
            self._sleep(self.throttle)
        attempt = 0
        while True:
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except urllib.error.HTTPError as exc:
                # 429 Too Many Requests / 503 Service Unavailable are transient:
                # wait (Retry-After if given, else exponential backoff) and retry.
                if exc.code in (429, 503) and attempt < self.max_retries:
                    hinted = self._retry_after_seconds(exc)
                    backoff = min(self.backoff_cap, (2 ** attempt) + random.random())
                    self._sleep(hinted if hinted is not None else backoff)
                    attempt += 1
                    continue
                raise ArchiveError(f"archive request failed ({path}): {exc}") from exc
            except (urllib.error.URLError, json.JSONDecodeError, ValueError) as exc:
                raise ArchiveError(f"archive request failed ({path}): {exc}") from exc

    def _http_index(self, from_date=None) -> ArchiveIndex:  # pragma: no cover - live only
        ids: set[str] = set()
        subjects: Counter = Counter()
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
                if not same_account:
                    continue
                mid = _norm(str(detail.get("messageId", "")))
                if mid:
                    ids.add(mid)
                subj = _subj(str(detail.get("subject", "")))
                if subj:
                    subjects[subj] += 1
            if page >= total_pages:
                break
            page += 1
        return ArchiveIndex(message_ids=ids, subject_counts=dict(subjects))


def _as_int(value, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def verify(manifest: dict, client: ArchiveClient, *,
           fallback_subject: bool = True) -> VerifyResult:
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
        index = client.account_index()
    except ArchiveError as exc:
        # Could not index the archive — nothing can be confirmed.
        result.checked = sum(len(v) for v in groups.values())
        result.errors.append({"error": str(exc)})
        return result

    present_ids = index.message_ids
    # Reserved-slot subject multiset: every archive message can back at most one
    # exported message, so an exact match consumes its subject slot before the
    # fallback pass may claim any remaining slot.
    subj_avail: Counter = Counter(index.subject_counts)

    not_found: list[tuple[str, list[dict]]] = []
    for message_id, records in groups.items():
        result.checked += len(records)
        if _norm(message_id) in present_ids:
            # Exact Message-ID hit (a conflict group dedups to one in the archive).
            result.present += 1
            if len(records) > 1:
                result.unresolved_conflicts += len(records) - 1
            subj = _subj(records[0].get("subject"))
            if subj and subj_avail.get(subj, 0) > 0:
                subj_avail[subj] -= 1
        else:
            not_found.append((message_id, records))

    for message_id, records in not_found:
        subj = _subj(records[0].get("subject"))
        if fallback_subject and subj and subj_avail.get(subj, 0) > 0:
            # A message with this subject is present under a re-derived
            # Message-ID; consume the slot so it is never double-counted.
            subj_avail[subj] -= 1
            result.present_fallback += 1
            if len(records) > 1:
                result.unresolved_conflicts += len(records) - 1
        elif subj and index.subject_counts.get(subj, 0) > 0:
            # Subject is in the archive but every copy is already accounted for;
            # cannot confirm THIS message — flag, do not claim present.
            result.ambiguous.append(message_id)
            if len(records) > 1:
                result.unresolved_conflicts += len(records)
        else:
            # Archive holds nothing with this subject: genuinely unconfirmed.
            result.missing.append(message_id)
            if len(records) > 1:
                result.unresolved_conflicts += len(records)
    return result
