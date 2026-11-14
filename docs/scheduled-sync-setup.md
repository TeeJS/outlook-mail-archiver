# Scheduled Archive Sync — Setup

One-time setup for the hands-off nightly sync. Three parts: the Unraid
container, the Unraid import watcher, and the PC scheduled task. Scripts live in
`scripts/`. Design and rationale: [scheduled-sync-plan.md](scheduled-sync-plan.md).

Dropbox share: `\\192.168.1.25\archives\outlook-email-archiver`
→ Unraid host dir `/mnt/user/archives/outlook-email-archiver`
→ mounted into the container at `/app/imports` (already an allowed import path).

---

## 1. Unraid — mount the dropbox into the container

In the Unraid web UI → **Docker** → click the **mail-archiver-sync** container → **Edit**:

1. Click **Add another Path, Port, Variable, Label or Device** → Config Type **Path**.
   - **Name:** `Import Dropbox`
   - **Container Path:** `/app/imports`
   - **Host Path:** `/mnt/user/archives/outlook-email-archiver`
   - **Access Mode:** Read/Write
2. **Apply** (the container recreates).

No environment variable is needed — `/app/imports` is in the shipped
`LocalImport.AllowedPaths` by default.

**Confirm the container name** (the watcher uses `mail-archiver-sync`):
```bash
docker ps --format '{{.Names}}'
```
If it differs, edit `CONTAINER=` at the top of the watcher script in step 2.

**Confirm the CLI import exists** (safe — imports nothing, just prints usage):
```bash
docker exec -w /app mail-archiver-sync dotnet /app/MailArchiver.dll --import-eml
```
Expect `ERROR: --file <path> is required for EML import`. That's success.

---

## 2. Unraid — install the import watcher

Needs the **User Scripts** plugin (Apps → search "User Scripts").

1. Settings → **User Scripts** → **Add New Script**, name it `mail-import-watcher`.
2. Click the gear → **Edit Script**, and paste the entire contents of
   [`scripts/unraid-import-watcher.sh`](../scripts/unraid-import-watcher.sh).
3. **Save**.
4. Set its schedule → **Custom**, cron: `*/15 * * * *` (every 15 minutes).

It imports each `*.zip` in the dropbox, moves successes to `done/` and failures
to `failed/`, and logs to `import.log` in the dropbox. A lock file stops two runs
overlapping.

---

## 3. PC — schedule the sync

Run this **once** in non-admin PowerShell to register both tasks. It prompts for
your Windows password (stored encrypted by Task Scheduler) so the task can reach
the network share while you're logged off. No elevation needed.

```powershell
$cred = Get-Credential -UserName "$env:USERDOMAIN\$env:USERNAME" -Message "Windows password (so the nightly task can reach the share)"
$ps   = 'D:\Github\outlook-mail-archiver\scripts\sync.ps1'

# Nightly incremental — 2:00 AM daily
$a1 = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ps`""
$t1 = New-ScheduledTaskTrigger -Daily -At 2:00AM
Register-ScheduledTask -TaskName 'MailArchiver Sync (daily)' -Action $a1 -Trigger $t1 `
  -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Limited -Force

# Monthly full sweep (integrity backstop) — 3:00 AM on the 1st
$a2 = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ps`" -Full"
$t2 = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Sunday -At 3:00AM
Register-ScheduledTask -TaskName 'MailArchiver Sync (full sweep)' -Action $a2 -Trigger $t2 `
  -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Limited -Force
```

The daily task: verify the previous drop → `export --incremental` → drop the new
ZIP. The full sweep re-exports everything (dedup makes the re-import mostly
skips) so back-dated or moved mail outside the 14-day overlap can't be missed.

Logs: `D:\mail-archive-staging\sync-logs\sync-YYYYMMDD.log`.

**Outlook must be running** (classic desktop) when the task fires — the export
reads it over COM. If the PC may be asleep at 2 AM, tick *"Wake the computer to
run this task"* in Task Scheduler → the task → Conditions.

---

## 4. First-run test

```powershell
D:\Github\outlook-mail-archiver\scripts\sync.ps1
```
Then within ~15 min check the archive got the new mail:
- Unraid: `cat /mnt/user/archives/outlook-email-archiver/import.log` and that the
  zip moved to `done/`.
- Next PC run (or run `sync.ps1` again): the verify step should report the dropped
  mail present and advance the watermark.

If a zip lands in `failed/`, the reason is in `import.log`.
