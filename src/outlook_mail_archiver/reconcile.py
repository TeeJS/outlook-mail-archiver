"""Reconcile an import result against an export manifest.

The rule from the charter: for the exported inventory,

    exported == imported + existing-duplicates + explicit-failures

with nothing unresolved. "Archived" means the destination accepted the message,
not merely that a ZIP was written. Aggregate counts can hide a mismatch, so a
balanced count is necessary but the stronger per-Message-ID check (via the
Mail-Archiver REST API) is the preferred confirmation where it is available.
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ImportCounts:
    imported: int = 0
    duplicates: int = 0
    failed: int = 0


@dataclass
class ReconcileResult:
    exported: int
    imported: int
    duplicates: int
    failed: int
    export_failed: int = 0

    @property
    def accounted(self) -> int:
        return self.imported + self.duplicates + self.failed

    @property
    def unresolved(self) -> int:
        return self.exported - self.accounted

    @property
    def balanced(self) -> bool:
        return self.unresolved == 0

    @property
    def counts_consistent(self) -> bool:
        # The IMPORTER'S COUNTS add up: something was exported, the numbers
        # balance, and neither side reported failures. This is necessary but NOT
        # sufficient — aggregate counts can hide a wrong-id/mismatched message.
        # Only per-message `verify` against the archive proves acceptance.
        return (
            self.exported > 0
            and self.balanced
            and self.failed == 0
            and self.export_failed == 0
        )

    def as_dict(self) -> dict:
        return {
            "exported": self.exported,
            "imported": self.imported,
            "duplicates": self.duplicates,
            "failed": self.failed,
            "export_failed": self.export_failed,
            "unresolved": self.unresolved,
            "balanced": self.balanced,
            "counts_consistent": self.counts_consistent,
        }


def exported_count(manifest: dict) -> int:
    return sum(1 for m in manifest.get("messages", []) if m.get("status") == "exported")


def load_manifest(package_dir: str | Path) -> dict:
    path = Path(package_dir) / "manifest.json"
    return json.loads(path.read_text(encoding="utf-8"))


def reconcile(manifest: dict, counts: ImportCounts) -> ReconcileResult:
    return ReconcileResult(
        exported=exported_count(manifest),
        imported=counts.imported,
        duplicates=counts.duplicates,
        failed=counts.failed,
        export_failed=int(manifest.get("totals", {}).get("failed", 0)),
    )
