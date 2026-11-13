"""Folder-path handling.

Outlook folder names can contain characters that collide with path separators —
notably a literal '/' (e.g. 'Qlik/RMM/SAP Alerts', 'SAP/apps'), '.', and others.
A folder tree is therefore carried as a list of raw name *segments*, never a
pre-joined string, so one folder with a slash in its name never splits into two.

Three renderings of the same segment list:
  * display_path  — human/config form, a literal '/' in a name escaped as '\\/'.
  * zip_path      — safe for a ZIP entry and the destination folder tree: each
                    segment sanitised, joined with '/'.
"""
from __future__ import annotations

import hashlib
import re

# Windows reserved device names (any case), which cannot be a path component.
_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_UNSAFE = re.compile(r'[/\\:*?"<>|\x00-\x1f]')


def parse_config_path(spec: str) -> list[str]:
    """Split a config folder path into segments.

    '/' separates segments; a literal slash inside one name is escaped '\\/'.
    """
    segments: list[str] = []
    cur: list[str] = []
    i = 0
    while i < len(spec):
        ch = spec[i]
        if ch == "\\" and i + 1 < len(spec) and spec[i + 1] == "/":
            cur.append("/")
            i += 2
            continue
        if ch == "/":
            segments.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    segments.append("".join(cur))
    return [s for s in segments if s != ""]


def display_path(segments: list[str]) -> str:
    """Render segments for humans/config, escaping a literal '/' in a name."""
    return "/".join(s.replace("/", "\\/") for s in segments)


def sanitize_segment(name: str) -> str:
    """Make one folder name safe as a single path component.

    Replaces separators and other unsafe characters, strips trailing dots and
    spaces (invalid on Windows), and disambiguates reserved device names. A
    period inside the name is preserved; only trailing dots are stripped.
    """
    cleaned = _UNSAFE.sub("_", name).rstrip(" .")
    if not cleaned:
        cleaned = "_"
    if cleaned.upper() in _RESERVED:
        cleaned = cleaned + "_"
    return cleaned


def map_segment(name: str) -> str:
    """Map one raw folder name to a stable, safe path component.

    Deterministic and order-independent: a name that is already safe maps to
    itself; a name that had to be sanitised gets a short suffix derived from the
    RAW name, so two distinct raw names that sanitise to the same base (e.g.
    'A/B' and 'A_B') never collide and always map the same way across runs.
    """
    safe = sanitize_segment(name)
    if safe == name:
        return safe
    tag = hashlib.sha1(name.encode("utf-8")).hexdigest()[:6]
    return f"{safe}~{tag}"


def zip_path(segments: list[str]) -> str:
    """Render segments as a POSIX ZIP/destination path.

    Each segment is mapped independently (so nested/ancestor segments are stable
    too) and deterministically, with no dependence on the order folders were
    selected in.
    """
    return "/".join(map_segment(s) for s in segments)
