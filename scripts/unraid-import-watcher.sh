#!/bin/bash
# unraid-import-watcher.sh — Unraid side of the hands-off Outlook sync.
#
# Imports each ZIP the PC drops in the dropbox into the Mail-Archiver IMPORT
# account via the container's own CLI, then files the ZIP away. Install it in
# the "User Scripts" plugin and schedule it every 15 minutes (custom cron
# */15 * * * *). See docs/scheduled-sync-setup.md.
#
# Re-importing is safe: Mail-Archiver dedups by Message-ID. The PC writes each
# file as .tmp then renames to .zip, so the *.zip glob never catches a partial.

set -u

# ---- settings -------------------------------------------------------------
CONTAINER="mail-archiver"                              # confirm: docker ps --format '{{.Names}}'
ACCOUNT_ID=3                                            # NWPIPE.COM import account
DROP="/mnt/user/archives/outlook-email-archiver"       # host side of the dropbox share
DLL="/app/MailArchiver.dll"                             # entrypoint inside the container
CONTAINER_IMPORT_DIR="/app/imports"                    # dropbox mounted here (already an allowed path)
# ---------------------------------------------------------------------------

DONE="$DROP/done"
FAILED="$DROP/failed"
LOG="$DROP/import.log"
mkdir -p "$DONE" "$FAILED"

# Only one watcher at a time (an import can outrun the 15-minute cron).
exec 9>"$DROP/.watcher.lock"
flock -n 9 || exit 0

log() { echo "$(date '+%F %T')  $*" >> "$LOG"; }

shopt -s nullglob
zips=("$DROP"/*.zip)
[ ${#zips[@]} -eq 0 ] && exit 0

for zip in "${zips[@]}"; do
    name="$(basename "$zip")"
    log "importing $name"
    if docker exec -w /app "$CONTAINER" dotnet "$DLL" \
            --import-eml --file "$CONTAINER_IMPORT_DIR/$name" --account-id "$ACCOUNT_ID" >> "$LOG" 2>&1; then
        mv -f "$zip" "$DONE/"
        log "OK   $name -> done/"
    else
        rc=$?
        mv -f "$zip" "$FAILED/"
        log "FAIL $name (exit $rc) -> failed/"
    fi
done
