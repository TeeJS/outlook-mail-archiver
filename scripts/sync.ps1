<#
    sync.ps1 — hands-off Outlook -> Mail-Archiver sync (PC side).

    Each run, in order:
      1. verify the PREVIOUS package (the Unraid watcher imported it during the
         day) -> advances the incremental high-water mark over confirmed mail.
      2. export --incremental  (or --full with -Full) -> a new package.
      3. atomically drop the new messages*.zip part(s) to the SMB dropbox for
         the Unraid watcher to import.

    Safe to re-run and safe to miss a run: the importer dedups by Message-ID and
    the export overlap window re-checks the boundary, so nothing is lost or
    duplicated. Nothing here touches the archive directly — it only writes ZIPs
    to the dropbox share.

    Register with Task Scheduler (see docs/scheduled-sync-setup.md). Non-admin;
    runs as the current user with network access to the share.
#>
[CmdletBinding()]
param(
    [string]$RepoDir    = 'D:\Github\outlook-mail-archiver',
    [string]$Python     = 'D:\Github\outlook-mail-archiver\.venv\Scripts\python.exe',
    [string]$StagingDir = 'D:\mail-archive-staging',
    [string]$Share      = '\\192.168.1.25\archives\outlook-email-archiver',
    [string]$LogDir     = 'D:\mail-archive-staging\sync-logs',
    [switch]$Full        # monthly backstop: export EVERYTHING (ignores the watermark)
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$logFile = Join-Path $LogDir ("sync-{0}.log" -f (Get-Date -Format 'yyyyMMdd'))

function Log([string]$msg) {
    $line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

function Invoke-Cli([string[]]$cliArgs) {
    Push-Location $RepoDir
    # A non-zero CLI exit must NOT throw here. verify intentionally exits non-zero
    # when mail is ambiguous/degraded but still present, and PowerShell 7.4+ turns a
    # non-zero native exit into a TERMINATING error while $ErrorActionPreference is
    # 'Stop' ($PSNativeCommandUseErrorActionPreference). Relax it for the call so
    # verify can't abort the run; callers gate on the returned exit code instead.
    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $Python -m outlook_mail_archiver.cli @cliArgs 2>&1 | ForEach-Object { Log "  | $_" }
        return $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prevEAP
        Pop-Location
    }
}

try {
    Log ("=== sync run start ({0}) ===" -f ($(if ($Full) { 'FULL sweep' } else { 'incremental' })))

    # 1) Verify the previous package (imported during the day) -> advance watermark.
    #    verify's exit code is non-zero whenever anything is degraded/ambiguous/
    #    conflicted even though the mail is present, so it is logged, never fatal.
    $prev = Get-ChildItem $StagingDir -Directory -Filter 'export-*' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($prev) {
        Log "verify previous package: $($prev.Name)"
        $rc = Invoke-Cli @('verify', $prev.FullName)
        Log "verify exit=$rc"
    } else {
        Log "no previous package to verify (first run)"
    }

    # 2) Export.
    $exportArgs = if ($Full) { @('export','--full') } else { @('export','--incremental') }
    Log ("export {0}" -f ($exportArgs[1]))
    $rc = Invoke-Cli $exportArgs
    Log "export exit=$rc"
    if ($rc -ne 0) { throw "export failed (exit $rc)" }

    # 3) Identify the package this export produced.
    $new = Get-ChildItem $StagingDir -Directory -Filter 'export-*' |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $new -or ($prev -and $new.FullName -eq $prev.FullName)) {
        Log "no new package produced; nothing to drop"
        Log "=== sync run done (no new mail) ==="; return
    }
    $parts = Get-ChildItem $new.FullName -Filter 'messages*.zip' -ErrorAction SilentlyContinue
    if (-not $parts) {
        Log "package $($new.Name) has no messages*.zip; nothing to drop"
        Log "=== sync run done (no new mail) ==="; return
    }

    # 4) Drop each ZIP part atomically: copy to .tmp, then server-side rename to
    #    .zip so the Unraid watcher (globbing *.zip) never sees a partial file.
    if (-not (Test-Path $Share)) { throw "dropbox share not reachable: $Share" }
    foreach ($p in $parts) {
        $destName = "{0}__{1}" -f $new.Name, $p.Name   # unique per package + part
        $dest = Join-Path $Share $destName
        Log "drop $($p.Name) -> $destName ($([math]::Round($p.Length/1MB,1)) MB)"
        Copy-Item $p.FullName "$dest.tmp" -Force
        Move-Item "$dest.tmp" $dest -Force
    }
    Log ("dropped {0} part(s) from {1}" -f $parts.Count, $new.Name)
    Log "=== sync run done ==="
}
catch {
    Log "ERROR: $($_.Exception.Message)"
    exit 1
}
