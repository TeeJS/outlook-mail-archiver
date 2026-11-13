"""Stage 1 — export.

Reads the selected Outlook folders (read-only), rebuilds each message as .eml,
and writes a durable, versioned export package:

    <staging>/export-<UTC timestamp>/
        messages.zip     ZIP whose folder tree mirrors the source (sanitised)
        manifest.json    one record per message + per-folder counts + totals

Guarantees reflected in the charter:
  * Every selected message is accounted for: exported, skipped-duplicate (the
    same message reached via two selected folders in one run), skipped-
    unsupported (non-mail item), or failed-with-reason. Nothing is dropped
    silently.
  * Interrupted writes do not look complete: the ZIP and manifest are written to
    ``.part`` files and renamed at the end; a package without manifest.json is
    incomplete.
  * Idempotency is at the DESTINATION: Mail-Archiver skips a message whose
    Message-ID it already holds, so re-exporting and re-importing never creates a
    second archive copy. The export deliberately does NOT skip based on prior
    local packages — trusting an unverified prior manifest could emit an empty
    ZIP with no recovery if that package was deleted or never imported.
"""
from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from . import __version__
from .config import Config
from .eml_builder import build_eml
from .identity import body_sha256
from .outlook_com import (
    extract_message, find_folder, iter_folder_items, list_folders, outlook_session,
)
from .paths import display_path, parse_config_path, zip_path
from .verify import _norm

MANIFEST_VERSION = 1
_SUBJECT_SAFE = re.compile(r"[^A-Za-z0-9._-]+")

# Notes that mean a REAL mail item lost content (incomplete, not merely a
# documented format limitation). These degrade a message: it is exported but
# must never count as a clean, complete archival success. Declared limitations
# (signed/encrypted S/MIME, non-mail items) are reported, not degraded.
_DEGRADED_PREFIXES = (
    "recipient-",                    # a specific recipient was unreadable
    "recipients-enumeration-failed",  # the recipient list could not be read
    "sender-smtp-unresolved",        # no usable sender address
    "inline-fallback",               # an inline image could not be anchored
    "inline-without-html",           # inline images with no HTML to attach to
)


def _is_degraded(notes: list[str]) -> bool:
    return any(str(n).startswith(_DEGRADED_PREFIXES) for n in notes)


class _RollingZip:
    """Writes .eml entries across one or more ZIP parts, each kept under a byte
    cap so every part can be uploaded through a size-limited web form.

    Part 1 is ``messages.zip`` (so small exports look unchanged); later parts are
    ``messages-02.zip``, ``messages-03.zip``, … A single message is never split:
    if it alone exceeds the cap it still goes in its own part.
    """

    def __init__(self, pkg: Path, cap_bytes: int):
        self.pkg = pkg
        self.cap = max(1, cap_bytes)
        self.part = 0
        self.zf = None
        self.bytes = 0
        self.parts: list[str] = []
        self._open_next()

    def _open_next(self) -> None:
        if self.zf is not None:
            self.zf.close()
        self.part += 1
        name = "messages.zip" if self.part == 1 else f"messages-{self.part:02d}.zip"
        self.zf = ZipFile(self.pkg / name, "w", ZIP_DEFLATED)
        self.bytes = 0
        self.parts.append(name)

    def write(self, arcname: str, data: bytes) -> str:
        if self.bytes > 0 and self.bytes + len(data) > self.cap:
            self._open_next()
        self.zf.writestr(arcname, data)
        self.bytes += len(data)
        return self.parts[-1]

    def close(self) -> None:
        if self.zf is not None:
            self.zf.close()
            self.zf = None


@dataclass
class ExportSummary:
    package_dir: Path | None = None
    processed: int = 0
    exported: int = 0
    skipped_duplicate: int = 0
    skipped_unsupported: int = 0
    skipped_already_archived: int = 0
    conflicts: int = 0
    degraded: int = 0
    failed: int = 0
    per_folder: dict = field(default_factory=dict)
    errors: list[dict] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.failed == 0


def _safe_name(subject: str, msg_hash: str, seq: int) -> str:
    stub = _SUBJECT_SAFE.sub("_", (subject or "no-subject")).strip("_")[:48] or "msg"
    return f"{seq:05d}_{stub}_{msg_hash[:12]}.eml"


def _resolve_targets(store, specs: list[str]) -> list[tuple[list[str], object, str]]:
    """Resolve the folders to export to (segments, folder_com, display) tuples.

    A spec of "*" means the whole store: the tree is walked ONCE (list_folders),
    instead of resolving each folder name separately (which would re-walk the
    tree per name). Otherwise each explicit spec is resolved by name.
    """
    if any((s or "").strip() == "*" for s in specs):
        return [(segs, folder, display_path(segs)) for segs, folder in list_folders(store)]
    targets = []
    for spec in specs:
        segs = parse_config_path(spec)
        targets.append((segs, find_folder(store, spec), display_path(segs)))
    return targets


def run_export(config: Config, *, incremental: bool | None = None,
               full_sweep: bool = False, archive_client=None,
               max_zip_bytes: int | None = None) -> ExportSummary:
    if not config.export.folders:
        raise ValueError(
            "No folders selected for export. Set export.folders in config.toml "
            '(a list of folder paths, or ["*"] for the whole mailbox; use '
            "`outlook-mail-archiver folders` to list them)."
        )

    staging = config.export.staging_dir
    staging.mkdir(parents=True, exist_ok=True)

    # Incremental mode: skip messages already in the archive, narrowed to a date
    # window from the verified high-water mark. First run (no watermark) or a
    # --full sweep scans everything; the archive's dedup makes that harmless.
    incremental = config.export.incremental if incremental is None else incremental
    archived_ids: set[str] = set()
    window_from = None
    eff_date_from = config.export.date_from
    if incremental:
        from datetime import timedelta
        from .state import load_state
        from .verify import ArchiveClient

        ma = config.mail_archiver
        if not ma.base_url or not ma.api_key or not ma.import_account_id:
            raise ValueError(
                "Incremental export needs mail_archiver.base_url, api_key and "
                "import_account_id in config.toml."
            )
        hwm = load_state(staging, ma.import_account_id).high_water_mark
        if hwm is not None and not full_sweep:
            window_from = (hwm - timedelta(days=config.export.overlap_days)).date()
            eff_date_from = window_from if eff_date_from is None else max(eff_date_from, window_from)
        if archive_client is None:
            archive_client = ArchiveClient(ma.base_url, ma.api_key, ma.import_account_id)
        archived_ids = {_norm(x) for x in archive_client.account_message_ids(from_date=window_from)}

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%SZ")
    pkg = staging / f"export-{stamp}"
    pkg.mkdir(parents=True, exist_ok=False)
    manifest_part = pkg / "manifest.json.part"
    cap_bytes = max_zip_bytes if max_zip_bytes is not None else config.export.max_zip_gb * 1024**3

    summary = ExportSummary(package_dir=pkg)
    messages: list[dict] = []
    conflicts_list: list[dict] = []
    seen_entry_ids: set[str] = set()        # true physical-item dedup
    id_index: dict[str, dict] = {}          # message_id -> first {entry_id, body}
    excluded_note = (
        "Only folders listed in export.folders were read; all others were "
        "excluded by selection."
    )

    rz = _RollingZip(pkg, cap_bytes)
    with outlook_session(config.outlook_store) as (_ns, store):
        store_name = store.DisplayName
        targets = _resolve_targets(store, config.export.folders)
        try:
            for segments, folder, disp in targets:
                if folder is None:
                    summary.failed += 1
                    summary.errors.append({"folder": disp, "reason": "folder-not-found"})
                    summary.per_folder[disp] = {"found": False}
                    continue

                folder_zip = zip_path(segments)
                counts = {"found": True, "exported": 0, "skipped": 0, "failed": 0}
                seq = 0

                for item in iter_folder_items(
                    folder, eff_date_from, config.export.date_to
                ):
                    summary.processed += 1
                    seq += 1
                    entry_id = getattr(item, "EntryID", "")
                    try:
                        md = extract_message(item)
                    except Exception as exc:
                        summary.failed += 1
                        counts["failed"] += 1
                        summary.errors.append({
                            "folder": disp, "entry_id": entry_id,
                            "reason": f"extract-failed: {exc}",
                        })
                        continue

                    # The extracted record carries the authoritative EntryID.
                    entry_id = md.entry_id or entry_id

                    if not md.is_mail:
                        summary.skipped_unsupported += 1
                        counts["skipped"] += 1
                        messages.append({
                            "folder": disp, "zip_folder": folder_zip,
                            "message_class": md.message_class,
                            "status": "skipped-unsupported",
                            "subject": md.subject,
                        })
                        continue

                    try:
                        eml_bytes, info = build_eml(md)
                    except Exception as exc:
                        summary.failed += 1
                        counts["failed"] += 1
                        summary.errors.append({
                            "folder": disp, "entry_id": entry_id,
                            "subject": md.subject, "reason": f"build-failed: {exc}",
                        })
                        continue

                    mid = info["message_id"]

                    # Incremental: the archive already holds this message -> skip.
                    if incremental and _norm(mid) in archived_ids:
                        summary.skipped_already_archived += 1
                        counts["skipped"] += 1
                        messages.append({
                            "folder": disp, "message_id": mid,
                            "status": "skipped-already-archived",
                            "subject": md.subject,
                        })
                        continue

                    bhash = body_sha256(md)

                    # True duplicate = the same physical item (same EntryID)
                    # reached twice in the selection. Only this is skipped.
                    if entry_id and entry_id in seen_entry_ids:
                        summary.skipped_duplicate += 1
                        counts["skipped"] += 1
                        messages.append({
                            "folder": disp, "zip_folder": folder_zip,
                            "message_id": mid, "entry_id": entry_id,
                            "status": "duplicate-item", "subject": md.subject,
                        })
                        continue
                    if entry_id:
                        seen_entry_ids.add(entry_id)

                    # Distinct items sharing a Message-ID are a CONFLICT, not a
                    # duplicate: export both (lose nothing) and flag it. The
                    # archive dedups on Message-ID at import, so one of the pair
                    # will be skipped there — verification must block on this.
                    notes = list(info["notes"])
                    prev = id_index.get(mid)
                    if prev is not None and prev["entry_id"] != entry_id:
                        kind = "different-body" if prev["body_sha256"] != bhash else "same-body"
                        notes.append(f"message-id-conflict:{kind}")
                        summary.conflicts += 1
                        conflicts_list.append({
                            "message_id": mid, "kind": kind,
                            "entry_ids": [prev["entry_id"], entry_id],
                        })
                    else:
                        id_index[mid] = {"entry_id": entry_id, "body_sha256": bhash}

                    msg_hash = hashlib.sha256(
                        (entry_id or mid).encode("utf-8")).hexdigest()
                    name = _safe_name(md.subject, msg_hash, seq)
                    arc = f"{folder_zip}/{name}" if folder_zip else name
                    zip_file = rz.write(arc, eml_bytes)

                    summary.exported += 1
                    counts["exported"] += 1
                    degraded = _is_degraded(notes)
                    if degraded:
                        summary.degraded += 1
                    dt = md.sent_on or md.received_time
                    messages.append({
                        "folder": disp,
                        "zip_path": arc,
                        "zip_file": zip_file,
                        "message_id": mid,
                        "entry_id": entry_id,
                        "generated_id": info["generated_id"],
                        "subject": md.subject,
                        "sender": md.sender_smtp,
                        "date": dt.isoformat() if dt else None,
                        "sha256": hashlib.sha256(eml_bytes).hexdigest(),
                        "size": len(eml_bytes),
                        "body_sha256": bhash,
                        "attachments": len(md.attachments),
                        "notes": notes,
                        "degraded": degraded,
                        "status": "exported",
                    })

                summary.per_folder[disp] = counts
        finally:
            rz.close()

    manifest = {
        "manifest_version": MANIFEST_VERSION,
        "tool_version": __version__,
        "export_id": f"export-{stamp}",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "source_store": store_name,
        "selection": {
            "folders": [display_path(parse_config_path(s)) for s in config.export.folders],
            "date_from": config.export.date_from.isoformat() if config.export.date_from else None,
            "date_to": config.export.date_to.isoformat() if config.export.date_to else None,
            "excluded_note": excluded_note,
        },
        "totals": {
            "processed": summary.processed,
            "exported": summary.exported,
            "skipped_duplicate": summary.skipped_duplicate,
            "skipped_unsupported": summary.skipped_unsupported,
            "skipped_already_archived": summary.skipped_already_archived,
            "conflicts": summary.conflicts,
            "degraded": summary.degraded,
            "failed": summary.failed,
        },
        "incremental": {
            "enabled": bool(incremental),
            "full_sweep": bool(full_sweep),
            "window_from": window_from.isoformat() if window_from else None,
        },
        "zip_parts": rz.parts,
        "per_folder": summary.per_folder,
        "conflicts": conflicts_list,
        "errors": summary.errors,
        "messages": messages,
    }

    manifest_part.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    # The ZIP parts are written directly; the manifest is written LAST (atomic
    # rename), so a package that has manifest.json is known-complete.
    manifest_part.replace(pkg / "manifest.json")
    return summary
