# Scheduled Archive Sync — Plan (awaiting sign-off)

Status: **DRAFT — not started.** Nothing touches the Unraid box until approved.
Trigger: **Option 2 — SMB dropbox (chosen).** PC never SSHes to Unraid.

## Charter

1. **One thing it must do:** nightly, capture only NEW Outlook mail into the
   archive with zero manual steps.
2. **Wrong if shipped without:** silent gaps. Each run must `verify` and only
   advance the watermark on confirmed import; the overlap window re-checks the
   boundary so nothing slips.
3. **Off-limits workaround:** the manual web upload, and any "fire and forget"
   that imports without verifying.
4. **Deploy target:** PC (Windows Task Scheduler) runs export/verify + drops the
   zip; the mail-archiver container on Unraid runs the import, triggered by an
   Unraid-side cron. Backup: staging dir retained until verified; git for code.
5. **Done =** a scheduled run with new mail lands it in the archive, `verify`
   confirms it, the watermark advances, and no human touched anything.

## How it works (decoupled, no SSH)

**PC nightly task** (one PowerShell wrapper, in order):

1. `verify` the PREVIOUS package (Unraid imported it during the day) →
   advances the watermark on success.
2. `export --incremental` → package with messages.zip (new mail only).
3. Copy messages.zip to the SMB dropbox as `<name>.tmp`, then rename to
   `<name>.zip` (atomic — the watcher never sees a half-written file).

**Unraid User Scripts cron** (bash, runs frequently, e.g. every 15 min):

- For each `*.zip` in the dropbox:
  `docker exec <container> dotnet MailArchiver.dll --import-eml --file /import/<zip> --account-id 3`
- On exit 0, move the zip to `done/`; on failure, move to `failed/` and log.

**Why verify-previous-then-export:** it needs no cross-machine timing. Last
night's drop is imported by the next run, so each run confirms the prior one
before exporting the next. Re-import is safe — the CLI dedups by Message-ID — so
a retried or double-processed zip never duplicates mail. The overlap window
covers any lag.

## Pieces

**Unraid (needs your OK — directive #11):**
- One share for the dropbox (host dir, e.g. `/mnt/user/appdata/mail-archiver/import`),
  with `done/` and `failed/` subdirs. Reachable from the PC as a UNC path.
- Bind-mount that host dir → `/import` on the container, and add `/import` to the
  allow-list via env `LocalImport__AllowedPaths__0=/import` (empty by default).
  I'll generate the updated Unraid template.
- Confirm the running image has `--import-eml` (safe probe: run it with no args,
  it prints usage).

**What I'll build (no box access needed):**
- The PowerShell wrapper (verify → export → atomic drop, with logging).
- The Unraid User Scripts bash watcher (import each zip, move to done/failed, log).
- The updated Unraid container template (bind-mount + env).

## Open — need from you
- **SMB dropbox share path:** the UNC path the PC writes to (e.g.
  `\\192.168.1.25\mail-import`). Propose/confirm, and I'll match the host mount.
- **Schedule (defaults, override freely):** PC export+drop **daily 2:00 AM**;
  Unraid watcher **every 15 min**.

## Success criteria
A scheduled run with new mail: the watcher imports it (exit 0), the next PC run's
`verify` shows it present, the watermark advances, zero manual steps. Verified by
watching one real cycle end to end.
