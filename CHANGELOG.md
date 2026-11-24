# Changes

## Installable Outlook PST converter — 2026-11-24

- Added the `email-funes` console command for import, audit, recovery, stats and export, with separately installable native, RTF and Windows COM dependencies.
- Required classification or an explicit all-record choice for the first unclassified export. Preserved existing PST provenance, selection policies and turn mappings.
- Made import diagnostics content-free by default and protected local database, report, recovery and export paths. Added offline wheel tests, Windows/Linux checks, secret scans and optional-dependency audits.
- Documented dedicated local Funes indexing, untrusted correspondence and the normalized message contract for future adapters.

## Metadata recovery and enriched export — 2026-11-20

- Added the PST metadata audit and second-pass SQLite recovery commands to the repository.
- Added native MAPI sender fields, structured To/Cc/Bcc recipients, item classification, unambiguous local Exchange address resolution, optional RTF body recovery, and field provenance.
- Preserved the original SQLite database, existing message/thread identities, and attachment-filename-only storage.
- Hardened per-message rollback accounting, Unicode text handling, source snapshot checks, and aggregate recovery logs.
- Added read-only export previews, an email-only default for classified archives, optional meetings/all-item/Bcc selection, and normalized recipient rendering.
- Persisted export policy alongside stable Funes mappings to prevent incompatible selection or rendering changes.
- Documented the audit → recovery → preview → JSONL workflow, portability limits, and Funes rebuild decisions.
- Added synthetic regression tests for metadata recovery and enriched exports; kept mailbox databases and logs outside source control.

## Repository edition — 2026-11-20

- Consolidated the earlier EML prototype and the version 3 PST/SQLite/Funes scripts into one repository.
- Separated shared SQLite storage and filters from the optional Outlook reader.
- Added the archive → SQLite → JSONL → local Funes → optional Hugging Face workflow, Mermaid diagrams, synthetic examples, and contributor guidance.
- Corrected installation guidance to use the official `libpff-python` package.
- Added regression coverage for conservative deduplication, resumable imports, and safe export folders.
- Kept original PST data, SQLite databases, generated exports, and extraction logs out of source control.

The direct reader replaced an earlier COM-only approach because some Outlook profiles could not initialize a usable session. The Outlook route remains available as an alternative. The original EML converter remains an experimental snapshot shortcut; future source readers should use the SQLite stage.
