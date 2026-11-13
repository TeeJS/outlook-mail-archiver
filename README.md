# outlook-mail-archiver

Archive mail from local **classic Outlook Desktop** into a self-hosted
[Mail-Archiver](https://github.com/s1t5/mail-archiver) instance — **without
Microsoft Graph**.

It reads your mail through Outlook's local COM automation (the already-running,
already-authenticated classic Outlook session — the same path as
[outlook-desktop-mcp](https://github.com/TeeJS/outlook-desktop-mcp)), rebuilds
each message as a standard `.eml`, and feeds a ZIP to Mail-Archiver's EML
importer. No Graph API, no Entra app registration, no OAuth.

Source access is strictly **read-only**: it never sends, moves, deletes, marks
read, or edits your mail.

## How it works — two stages

```
classic Outlook (COM, read-only) ─▶ rebuild .eml ─▶ ZIP by folder + manifest
                                                           │
                                                           ▼
                             Mail-Archiver EML import ─▶ reconcile ─▶ verify
```

**Stage 1 — export** produces a durable, versioned package:

```
<staging>/export-<UTC timestamp>/
    messages.zip     folder tree mirrors the source (names sanitised safely)
    manifest.json    one record per message + per-folder counts + totals
```

**Stage 2 — import** is separate: Mail-Archiver ingests the ZIP (its CLI
`--import-eml`, or the web upload). Then you **reconcile** counts and, for real
acceptance evidence, **verify** per-Message-ID against the archive.

## Requirements

- Windows with **classic** Outlook Desktop running (new/modern Outlook is **not**
  supported — it has no COM)
- Python 3.12+
- A running Mail-Archiver instance with an account of type **IMPORT**

## Install

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e .
```

## Configure

```powershell
Copy-Item config.example.toml config.toml
```

Edit `config.toml` (it is git-ignored and never leaves your machine):

- `export.staging_dir` — where packages are written (private mail; keep it local)
- `export.folders` — explicit folder paths to export (nothing is omitted silently)
- `export.date_from` / `date_to` — optional inclusive date bounds
- `mail_archiver.base_url`, `import_account_id`, `api_key` — for import/verify

List folder paths to copy into `export.folders`:

```powershell
.\.venv\Scripts\python.exe -m outlook_mail_archiver.cli folders
```

A literal `/` inside one folder name (e.g. `Qlik/RMM/SAP Alerts`) is written in
config with a backslash: `Qlik\/RMM\/SAP Alerts`.

## Use

```powershell
# Stage 1 — export the selected folders
python -m outlook_mail_archiver.cli export

# Stage 2 — show (or run) the import for a package
python -m outlook_mail_archiver.cli import <staging>\export-<stamp>

# reconcile the importer's reported counts against the manifest
python -m outlook_mail_archiver.cli reconcile <package> --import-output import.log

# per-message acceptance check against the archive (the real evidence)
python -m outlook_mail_archiver.cli verify <package>
```

`pilot` runs an export and prints the import command for the newest package —
use it for the small representative first run before a full transfer.

### Mail-Archiver import setup (one-time, on the server)

The CLI import reads from a whitelisted directory. In the Mail-Archiver
`docker-compose.yml`:

```yaml
environment:
  - LocalImport__AllowedPaths__0=/data/import
volumes:
  - /path/on/host/import:/data/import
```

Place `messages.zip` in that directory and run:

```bash
docker compose exec mailarchive-app dotnet MailArchiver.dll \
  --import-eml --file /data/import/messages.zip --account-id <id>
```

The web **Import EML** form is a fallback that needs no server change.

## Verification — counts are not enough

Idempotency is at the **destination**: Mail-Archiver skips a message whose
`Message-ID` it already holds, so re-importing never creates a second copy.

- `reconcile` checks the importer's counts balance: `exported = imported +
  duplicates + failed`. This is **necessary but not sufficient** — aggregate
  counts can hide a mismatched message.
- `verify` is the real evidence: it looks up each exported `Message-ID` in the
  archive via the read-only REST API and confirms an exact `messageId` +
  `accountId` match. Only this reports `VERIFIED`.

> The `verify` endpoint shape is implemented to the documented REST API but has
> not yet been exercised against a live instance; validate it on the real
> account as part of the first pilot (it needs `base_url`, `api_key`,
> `import_account_id`).

## Identity & dedupe

- The original `Message-ID` is used when present and well-formed.
- When absent, a **deterministic** substitute is synthesised from the message's
  own content (date, sender, recipients, normalised subject, body hash,
  attachment byte-hashes) — folder-independent, so moving a message does not
  change its id, and stable across re-exports.
- Two **distinct** items that share one `Message-ID` are a **conflict**: both are
  exported (nothing is lost) and flagged in the manifest. The archive dedups on
  `Message-ID` at import, so one of the pair is kept there — `verify` reports
  such a pair as unresolved rather than claiming success.

## Fidelity & limitations

A rebuilt `.eml` is a faithful copy of ordinary mail — plain/HTML body, subject,
dates (with offset), resolved SMTP addresses for sender and recipients,
attachments, and inline (CID) images as `multipart/related`. It is **not** a
byte-for-byte original:

- **Signed / encrypted (S/MIME):** rebuilt from readable parts; the original
  signature is not reproduced and encrypted content is not decrypted. These are
  detected and noted, never claimed as faithful.
- **Calendar items, contacts, and other non-mail** are recorded as
  `skipped-unsupported` (not silently dropped).
- **Edited-in-place mail:** if a message keeps its `Message-ID` but its body
  changed, the archive will not update the already-held copy. The export records
  `body_sha256` so the change is detectable; it is not an in-place update.

## Security

- Read-only against Outlook; it attaches to the running session and never
  launches Outlook or stores Outlook credentials.
- Mail-Archiver credentials / API key live only in the git-ignored `config.toml`
  and are never written into packages, manifests, or logs.
- Export packages contain private mail — keep `staging_dir` off any shared or
  public location.

## Status

Stage 1 (export) is implemented and proven against a real mailbox (read-only).
Stage 2 import driver, reconciliation, and per-message verification are
implemented and unit-tested; the live import and `verify` paths are exercised
with your instance's credentials as the first-pilot verification step.

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

## License

MIT (see LICENSE).
