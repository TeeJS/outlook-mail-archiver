# outlook-mail-archiver

Archive mail from local **classic Outlook Desktop** into a self-hosted
[Mail-Archiver](https://github.com/s1t5/mail-archiver) instance — **without Microsoft Graph**.

## Why

Reads mail through Outlook's local COM automation (the same connection used by
[outlook-desktop-mcp](https://github.com/TeeJS/outlook-desktop-mcp)), rebuilds each
message as a standard `.eml`, and feeds it to Mail-Archiver's EML import. No Graph
API, no Entra app registration, no OAuth — it rides the already-authenticated
Outlook session.

## Status

🚧 Early groundwork. Feasibility proven (COM read → `.eml` rebuild → Mail-Archiver
import); design and implementation in progress.

## How it works (planned)

```
classic Outlook (COM) ──▶ rebuild .eml ──▶ ZIP by folder ──▶ Mail-Archiver EML import
```

- **Read:** Outlook Desktop COM (classic `OUTLOOK.EXE` must be running; 64-bit).
- **Convert:** per-message rebuild (headers + HTML/plain body + attachments).
  Outlook's native MAPI→MIME converter is unavailable to out-of-process callers
  under Click-to-Run, so messages are rebuilt from their parts.
- **Import:** Mail-Archiver account of type `IMPORT`; ZIP folder paths become
  archive folders; duplicates are skipped by `Message-ID` (safe to re-run).

## Requirements

- Windows with classic Outlook Desktop (new/modern Outlook is **not** supported)
- Python 3.12+
- A running Mail-Archiver instance

## License

TBD
