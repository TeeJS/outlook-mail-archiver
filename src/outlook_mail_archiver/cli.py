"""Command-line interface.

    outlook-mail-archiver folders              list the Outlook folder tree
    outlook-mail-archiver export               run Stage 1 export
    outlook-mail-archiver import <package>     Stage 2: show/run the import
    outlook-mail-archiver reconcile <package>  reconcile import output vs manifest
    outlook-mail-archiver verify <package>     per-message check vs the archive
    outlook-mail-archiver pilot                export, then show the import command

All source access is read-only. Exit status is non-zero when anything is left
failed or unresolved.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import ConfigError, load_config
from .import_driver import parse_import_output, run_import
from .reconcile import ImportCounts, load_manifest, reconcile


def _load(args) -> "object":
    try:
        return load_config(args.config)
    except ConfigError as exc:
        print(f"config error: {exc}", file=sys.stderr)
        raise SystemExit(2)


def cmd_folders(args) -> int:
    from .outlook_com import OutlookError, list_folders, outlook_session
    from .paths import display_path

    config = _load(args)
    try:
        with outlook_session(config.outlook_store) as (_ns, store):
            print(f"# Store: {store.DisplayName}")
            for segments, folder in list_folders(store):
                try:
                    count = folder.Items.Count
                except Exception:
                    count = "?"
                print(f"{display_path(segments)}\t[{count}]")
    except OutlookError as exc:
        print(f"outlook error: {exc}", file=sys.stderr)
        return 2
    return 0


def cmd_export(args) -> int:
    from .export import run_export
    from .outlook_com import OutlookError

    config = _load(args)
    try:
        summary = run_export(config)
    except OutlookError as exc:
        print(f"outlook error: {exc}", file=sys.stderr)
        return 2
    except ValueError as exc:
        print(f"{exc}", file=sys.stderr)
        return 2

    print(f"Export package: {summary.package_dir}")
    print(f"  processed:             {summary.processed}")
    print(f"  exported:              {summary.exported}")
    print(f"  skipped (duplicate):   {summary.skipped_duplicate}")
    print(f"  skipped (unsupported): {summary.skipped_unsupported}")
    print(f"  message-id conflicts:  {summary.conflicts}")
    print(f"  failed:                {summary.failed}")
    if summary.conflicts:
        print("  note: distinct messages share a Message-ID; both were exported,"
              " but the archive will keep one. See manifest 'conflicts'.")
    if summary.errors:
        print("  errors:")
        for e in summary.errors[:20]:
            print(f"    - {e}")
        if len(summary.errors) > 20:
            print(f"    ... and {len(summary.errors) - 20} more (see manifest.json)")
    if not summary.ok:
        print("Export completed WITH failures (see manifest.json).", file=sys.stderr)
        return 1
    print("Export completed cleanly.")
    return 0


def cmd_import(args) -> int:
    config = _load(args)
    pkg = Path(args.package)
    zip_path = pkg / "messages.zip"
    if not zip_path.exists():
        print(f"no messages.zip in {pkg}", file=sys.stderr)
        return 2

    output, counts = run_import(config, zip_path)
    if counts is None:
        print("Mail-Archiver import is file-based. Place the ZIP in the importer's")
        print("allowed directory and run:\n")
        print(f"    {output}\n")
        print("Then reconcile the result, e.g.:")
        print(f"    outlook-mail-archiver reconcile {pkg} --import-output import.log")
        return 0

    print(output)
    result = reconcile(load_manifest(pkg), counts)
    _print_reconcile(result)
    return 0 if result.counts_consistent else 1


def cmd_reconcile(args) -> int:
    _load(args)  # validates config exists
    pkg = Path(args.package)
    if args.import_output:
        text = Path(args.import_output).read_text(encoding="utf-8", errors="replace")
    else:
        print("Paste the Mail-Archiver import summary, then Ctrl-Z/Ctrl-D:",
              file=sys.stderr)
        text = sys.stdin.read()
    counts = parse_import_output(text)
    result = reconcile(load_manifest(pkg), counts)
    _print_reconcile(result)
    return 0 if result.counts_consistent else 1


def cmd_verify(args) -> int:
    from .verify import ArchiveClient, verify

    config = _load(args)
    ma = config.mail_archiver
    if not ma.base_url or not ma.api_key or not ma.import_account_id:
        print("verify needs mail_archiver.base_url, api_key and import_account_id "
              "in config.toml.", file=sys.stderr)
        return 2
    manifest = load_manifest(Path(args.package))
    client = ArchiveClient(ma.base_url, ma.api_key, ma.import_account_id)
    result = verify(manifest, client)
    print("Per-message verification against the archive:")
    print(f"  checked:              {result.checked}")
    print(f"  present:              {result.present}")
    print(f"  missing:              {len(result.missing)}")
    print(f"  unresolved conflicts: {result.unresolved_conflicts}")
    print(f"  query errors:         {len(result.errors)}")
    print(f"  source failed:        {result.source_failed}")
    print(f"  source folder errors: {result.source_errors}")
    print(f"  non-mail skipped:     {result.skipped_unsupported}")
    for e in result.errors[:10]:
        print(f"    - {e}")
    if result.source_failed or result.source_errors:
        print("  note: the export itself was incomplete (failures/errors); "
              "verification cannot report full success.")
    if result.skipped_unsupported:
        print("  note: non-mail items were skipped by design (see manifest); "
              "they are not part of the mail transfer.")
    if result.verified_success:
        print("  => VERIFIED: every exported message is present in the archive.")
        return 0
    print("  => NOT verified. Do not treat the transfer as complete.", file=sys.stderr)
    return 1


def cmd_pilot(args) -> int:
    rc = cmd_export(args)
    if rc not in (0, 1):
        return rc
    # Show the import command for the newest package.
    config = _load(args)
    packages = sorted(config.export.staging_dir.glob("export-*"))
    if packages:
        print("\nNext — import this package:")
        print(f"    outlook-mail-archiver import {packages[-1]}")
    return rc


def _print_reconcile(result) -> None:
    print("Reconciliation (exported = imported + duplicates + failed):")
    print(f"  exported:       {result.exported}")
    print(f"  imported:       {result.imported}")
    print(f"  duplicates:     {result.duplicates}")
    print(f"  failed (import):{result.failed}")
    print(f"  failed (export):{result.export_failed}")
    print(f"  unresolved:     {result.unresolved}")
    if result.counts_consistent:
        print("  => counts consistent (necessary, NOT sufficient). Run "
              "`verify` for per-message archive confirmation before trusting it.")
    elif result.exported == 0:
        print("  => nothing was exported; not a success. Check the export.")
    elif not result.balanced:
        print("  => UNRESOLVED messages: counts do not balance. Do not treat as done.")
    else:
        print("  => balanced, but there were failures (import and/or export); "
              "see the reports.")


def build_parser() -> argparse.ArgumentParser:
    # A shared parent so --config is accepted both before and after the command
    # (outlook-mail-archiver --config x export  AND  ... export --config x).
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--config", default="config.toml", help="path to config.toml")

    p = argparse.ArgumentParser(prog="outlook-mail-archiver", description=__doc__,
                                parents=[common])
    sub = p.add_subparsers(dest="command", required=True)

    sub.add_parser("folders", parents=[common],
                   help="list the Outlook folder tree").set_defaults(func=cmd_folders)
    sub.add_parser("export", parents=[common],
                   help="run the export stage").set_defaults(func=cmd_export)

    pi = sub.add_parser("import", parents=[common],
                        help="show or run the import for a package")
    pi.add_argument("package", help="export package directory")
    pi.set_defaults(func=cmd_import)

    pr = sub.add_parser("reconcile", parents=[common],
                        help="reconcile import output vs manifest")
    pr.add_argument("package", help="export package directory")
    pr.add_argument("--import-output", help="file with the importer's summary output")
    pr.set_defaults(func=cmd_reconcile)

    pv = sub.add_parser("verify", parents=[common],
                        help="per-message check that exports reached the archive")
    pv.add_argument("package", help="export package directory")
    pv.set_defaults(func=cmd_verify)

    pp = sub.add_parser("pilot", parents=[common],
                        help="export then show the import command")
    pp.set_defaults(func=cmd_pilot)
    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
