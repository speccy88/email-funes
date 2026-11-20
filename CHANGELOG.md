# Changes

## Repository edition — 2026-11-20

- Consolidated the earlier EML prototype and the version 3 PST/SQLite/Funes scripts into one repository.
- Separated shared SQLite storage and filters from the optional Outlook reader.
- Added the archive → SQLite → JSONL → local Funes → optional Hugging Face workflow, Mermaid diagrams, synthetic examples, and contributor guidance.
- Corrected installation guidance to use the official `libpff-python` package.
- Added regression coverage for conservative deduplication, resumable imports, and safe export folders.
- Kept original PST data, SQLite databases, generated exports, and extraction logs out of source control.

The direct reader replaced an earlier COM-only approach because some Outlook profiles could not initialize a usable session. The Outlook route remains available as an alternative. The original EML converter remains an experimental snapshot shortcut; future source readers should use the SQLite stage.
