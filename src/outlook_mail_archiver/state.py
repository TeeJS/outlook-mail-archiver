"""Durable incremental-sync state.

A small per-account JSON file in the staging directory records the high-water
mark: the newest message date that has been **verified** present in the archive.
The export uses it to limit each incremental run to recent mail; it is advanced
only by a successful ``verify``, so the window never moves past mail that has not
actually been confirmed archived.

Losing this file is harmless: the next run simply scans wider (a full sweep),
and the archive's own Message-ID dedup prevents duplicates.
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path


@dataclass
class SyncState:
    account_id: int
    high_water_mark: datetime | None = None


def state_path(staging_dir: Path, account_id: int) -> Path:
    return Path(staging_dir) / f"sync-state-{account_id}.json"


def load_state(staging_dir: Path, account_id: int) -> SyncState:
    path = state_path(staging_dir, account_id)
    if not path.exists():
        return SyncState(account_id=account_id)
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return SyncState(account_id=account_id)
    hwm = raw.get("high_water_mark")
    dt = None
    if hwm:
        try:
            dt = datetime.fromisoformat(hwm)
        except ValueError:
            dt = None
    return SyncState(account_id=raw.get("account_id", account_id), high_water_mark=dt)


def save_state(state: SyncState, staging_dir: Path) -> None:
    path = state_path(staging_dir, state.account_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "account_id": state.account_id,
        "high_water_mark": state.high_water_mark.isoformat() if state.high_water_mark else None,
        "updated_utc": datetime.now(timezone.utc).isoformat(),
    }
    tmp = path.with_suffix(".json.part")
    tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    tmp.replace(path)


def advance_high_water_mark(staging_dir: Path, account_id: int,
                            newest: datetime | None) -> SyncState:
    """Move the watermark forward to ``newest`` (never backward). No-op if None."""
    state = load_state(staging_dir, account_id)
    if newest is not None and (state.high_water_mark is None
                               or newest > state.high_water_mark):
        state.high_water_mark = newest
        save_state(state, staging_dir)
    return state
