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
from .outlook_com import extract_message, find_folder, iter_folder_items, outlook_session
from .paths import display_path, parse_config_path, zip_path

MANIFEST_VERSION = 1
_SUBJECT_SAFE = re.compile(r"[^A-Za-z0-9._-]+")


@dataclass
class ExportSummary:
    package_dir: Path | None = None
    processed: int = 0
    exported: int = 0
    skipped_duplicate: int = 0
    skipped_unsupported: int = 0
    conflicts: int = 0
    failed: int = 0
    per_folder: dict = field(default_factory=dict)
    errors: list[dict] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.failed == 0


def _safe_name(subject: str, msg_hash: str, seq: int) -> str:
    stub = _SUBJECT_SAFE.sub("_", (subject or "no-subject")).strip("_")[:48] or "msg"
    return f"{seq:05d}_{stub}_{msg_hash[:12]}.eml"


def run_export(config: Config) -> ExportSummary:
    if not config.export.folders:
        raise ValueError(
            "No folders selected for export. Set export.folders in config.toml "
            "(use `outlook-mail-archiver folders` to list them)."
        )

    staging = config.export.staging_dir
    staging.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%SZ")
    pkg = staging / f"export-{stamp}"
    pkg.mkdir(parents=True, exist_ok=False)
    zip_part = pkg / "messages.zip.part"
    manifest_part = pkg / "manifest.json.part"

    summary = ExportSummary(package_dir=pkg)
    messages: list[dict] = []
    conflicts_list: list[dict] = []
    seen_entry_ids: set[str] = set()        # true physical-item dedup
    id_index: dict[str, dict] = {}          # message_id -> first {entry_id, body}
    excluded_note = (
        "Only folders listed in export.folders were read; all others were "
        "excluded by selection."
    )

    with outlook_session(config.outlook_store) as (_ns, store):
        store_name = store.DisplayName
        with ZipFile(zip_part, "w", ZIP_DEFLATED) as zf:
            for spec in config.export.folders:
                segments = parse_config_path(spec)
                disp = display_path(segments)
                folder = find_folder(store, spec)
                if folder is None:
                    summary.failed += 1
                    summary.errors.append({"folder": disp, "reason": "folder-not-found"})
                    summary.per_folder[disp] = {"found": False}
                    continue

                folder_zip = zip_path(segments)
                counts = {"found": True, "exported": 0, "skipped": 0, "failed": 0}
                seq = 0

                for item in iter_folder_items(
                    folder, config.export.date_from, config.export.date_to
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
                    zf.writestr(arc, eml_bytes)

                    summary.exported += 1
                    counts["exported"] += 1
                    dt = md.sent_on or md.received_time
                    messages.append({
                        "folder": disp,
                        "zip_path": arc,
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
                        "status": "exported",
                    })

                summary.per_folder[disp] = counts

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
            "conflicts": summary.conflicts,
            "failed": summary.failed,
        },
        "per_folder": summary.per_folder,
        "conflicts": conflicts_list,
        "errors": summary.errors,
        "messages": messages,
    }

    manifest_part.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    # Commit atomically: rename the ZIP, then the manifest last so a package that
    # has manifest.json is known-complete.
    zip_part.replace(pkg / "messages.zip")
    manifest_part.replace(pkg / "manifest.json")
    return summary
