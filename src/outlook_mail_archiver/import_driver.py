"""Stage 2 — import driver.

Mail-Archiver ingests a ZIP of .eml files via its CLI (``--import-eml``) or the
web upload form; there is no per-message import API. This driver does NOT spawn a
shell by default (a build/run tool quietly spawning shells is exactly what
endpoint protection flags). By default it prints the exact, copy-ready command
for you to run, then parses the output you paste back to reconcile.

Only when ``mail_archiver.import_command`` is explicitly set in config does it
run that command for you.
"""
from __future__ import annotations

import re
import subprocess
from pathlib import Path

from .config import Config
from .reconcile import ImportCounts

# Mail-Archiver CLI summary lines, e.g.:
#   Total Emails: 15234
#   Imported Successfully: 15120
#   Failed: 42
#   Skipped (duplicates): 44
_PATTERNS = {
    "imported": re.compile(r"Imported(?:\s+Successfully)?\s*:?\s*(\d+)", re.I),
    "failed": re.compile(r"\bFailed\s*:?\s*(\d+)", re.I),
    "duplicates": re.compile(r"Skipped\s*\(?duplicates?\)?\s*:?\s*(\d+)", re.I),
}


def documented_command(config: Config, zip_path: Path) -> str:
    acct = config.mail_archiver.import_account_id or "<account-id>"
    return (
        "docker compose exec mailarchive-app dotnet MailArchiver.dll "
        f"--import-eml --file /data/import/{zip_path.name} --account-id {acct}"
    )


def parse_import_output(text: str) -> ImportCounts:
    def grab(key: str) -> int:
        m = _PATTERNS[key].search(text or "")
        return int(m.group(1)) if m else 0

    return ImportCounts(
        imported=grab("imported"),
        duplicates=grab("duplicates"),
        failed=grab("failed"),
    )


def run_import(config: Config, zip_path: Path) -> tuple[str, ImportCounts | None]:
    """Return (output_text, counts_or_None).

    With ``import_command`` configured, runs it and parses the output. Otherwise
    returns the documented command and no counts (the user runs it and feeds the
    output to ``reconcile``).
    """
    cmd = config.mail_archiver.import_command.strip()
    if not cmd:
        return documented_command(config, zip_path), None

    # Explicit opt-in only. The configured command is the user's own; we pass the
    # package path and never inject credentials.
    full = cmd.replace("{zip}", str(zip_path)).replace("{name}", zip_path.name)
    proc = subprocess.run(full, shell=True, capture_output=True, text=True)
    output = (proc.stdout or "") + (proc.stderr or "")
    return output, parse_import_output(output)
