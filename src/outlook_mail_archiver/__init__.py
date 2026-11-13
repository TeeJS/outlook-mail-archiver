"""outlook-mail-archiver: archive classic Outlook Desktop mail into Mail-Archiver
via local COM, no Microsoft Graph.

Two stages, deliberately separate:
  * export  — read Outlook (read-only COM), rebuild each message as .eml, write a
              ZIP mirroring the folder tree plus a durable manifest.
  * import  — hand the ZIP to Mail-Archiver's EML importer and reconcile the
              result against the manifest (destination acceptance, not ZIP-written,
              is success).
"""

__version__ = "0.1.0"
