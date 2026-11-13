"""Local configuration loader.

Config lives in a git-ignored ``config.toml`` (see ``config.example.toml``). It
holds the destination instance details and the export selection. Nothing loaded
here is ever written into export packages, manifests, or logs.
"""
from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path


class ConfigError(Exception):
    """Raised when the configuration is missing or invalid."""


@dataclass(frozen=True)
class ExportConfig:
    staging_dir: Path
    folders: list[str] = field(default_factory=list)
    date_from: date | None = None
    date_to: date | None = None
    incremental: bool = False
    overlap_days: int = 14


@dataclass(frozen=True)
class MailArchiverConfig:
    base_url: str = ""
    import_account_id: int = 0
    import_command: str = ""
    api_key: str = ""


@dataclass(frozen=True)
class Config:
    outlook_store: str
    export: ExportConfig
    mail_archiver: MailArchiverConfig
    source_path: Path


def _parse_date(value: str, field_name: str) -> date | None:
    value = (value or "").strip()
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ConfigError(f"{field_name} must be ISO yyyy-mm-dd, got {value!r}") from exc


def load_config(path: str | Path = "config.toml") -> Config:
    """Load and validate ``config.toml``.

    Only the fields the requested operation needs are validated hard; e.g. the
    export selection is checked when exporting, not merely on load.
    """
    path = Path(path)
    if not path.exists():
        raise ConfigError(
            f"Config file not found: {path}. Copy config.example.toml to "
            f"config.toml and fill it in."
        )
    with path.open("rb") as fh:
        raw = tomllib.load(fh)

    outlook = raw.get("outlook", {})
    export = raw.get("export", {})
    archiver = raw.get("mail_archiver", {})

    staging = (export.get("staging_dir") or "").strip()
    if not staging:
        raise ConfigError("export.staging_dir is required.")

    folders = [f for f in (export.get("folders") or []) if str(f).strip()]

    return Config(
        outlook_store=(outlook.get("store") or "").strip(),
        export=ExportConfig(
            staging_dir=Path(staging),
            folders=folders,
            date_from=_parse_date(export.get("date_from", ""), "export.date_from"),
            date_to=_parse_date(export.get("date_to", ""), "export.date_to"),
            incremental=bool(export.get("incremental", False)),
            overlap_days=int(export.get("overlap_days", 14) or 14),
        ),
        mail_archiver=MailArchiverConfig(
            base_url=(archiver.get("base_url") or "").strip(),
            import_account_id=int(archiver.get("import_account_id") or 0),
            import_command=(archiver.get("import_command") or "").strip(),
            api_key=(archiver.get("api_key") or "").strip(),
        ),
        source_path=path,
    )
