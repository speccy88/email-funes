#!/usr/bin/env python3
"""Export stored emails to Funes .funes.jsonl files.

Maintains a stable per-conversation sequence map in SQLite, so re-exporting
and importing new mail won't renumber old turns. Run after PST import.

Usage:
  py sqlite_to_funes.py --db "C:\\path\\mail.sqlite" --out "C:\\path\\funes_exports"
  funes index --check /path/to/funes_exports
  funes index /path/to/funes_exports
"""
from __future__ import annotations

import argparse
from dataclasses import asdict, dataclass
from datetime import datetime
import json
import sqlite3
import shutil
import tempfile
from pathlib import Path


@dataclass(frozen=True)
class ExportPolicy:
    """Choices that must remain stable once turn IDs have been assigned."""

    selection: str = "all"
    classified: bool = False
    recipient_source: str = "legacy"
    include_bcc: bool = False
    format_version: int = 1

    def serialize(self) -> str:
        return json.dumps(asdict(self), sort_keys=True, separators=(",", ":"))


def table_exists(db: sqlite3.Connection, name: str) -> bool:
    return db.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def resolve_policy(db: sqlite3.Connection, *, include_meetings: bool = False,
                   include_all: bool = False, include_bcc: bool = False) -> ExportPolicy:
    """Select email records when the database contains recovery classification."""
    columns = {row[1] for row in db.execute("PRAGMA table_info(messages)")}
    if not columns:
        raise ValueError("The database has no messages table")
    if include_meetings and include_all:
        raise ValueError("Choose --include-meetings or --include-all, not both")
    classified = "message_category" in columns
    if include_meetings and not classified:
        raise ValueError("--include-meetings requires message classification; run metadata recovery first")
    selection = "all" if include_all or not classified else "email-and-meetings" if include_meetings else "email"
    recipient_source = "normalized" if table_exists(db, "recovery_recipients") else "legacy"
    return ExportPolicy(selection, classified, recipient_source, include_bcc)


def selection_sql(policy: ExportPolicy) -> tuple[str, tuple[str, ...]]:
    """Return a parameterized predicate; category names are the recovery v1 values."""
    if policy.selection == "all":
        return "1", ()
    categories = ("email / ipm.note", "email")
    if policy.selection == "email-and-meetings":
        categories += ("meeting request", "meeting")
    placeholders = ",".join("?" for _ in categories)
    return f"LOWER(TRIM(COALESCE(m.message_category,''))) IN ({placeholders})", categories


def validate_policy(db: sqlite3.Connection, policy: ExportPolicy) -> None:
    """Check compatibility without writing tables or changing existing identities."""
    if table_exists(db, "funes_export_policy"):
        row = db.execute("SELECT policy_json FROM funes_export_policy WHERE singleton=1").fetchone()
        if row and row[0] != policy.serialize():
            raise ValueError(
                "Export policy differs from the stored Funes mapping. Keep the original options, "
                "or deliberately rebuild the mapping in a separate database copy and use a fresh Funes memory."
            )
    else:
        row = None
    if not table_exists(db, "funes_map"):
        return
    mapped = db.execute("SELECT COUNT(*) FROM funes_map").fetchone()[0]
    if mapped and row is None and policy != ExportPolicy():
        raise ValueError(
            "Existing Funes mapping has no export policy and uses legacy rendering. "
            "Enriched selection or recipient changes require a deliberately rebuilt mapping "
            "in a separate database copy and a fresh Funes memory."
        )
    predicate, params = selection_sql(policy)
    excluded = db.execute(
        f"SELECT COUNT(*) FROM funes_map fm LEFT JOIN messages m ON m.id=fm.message_id "
        f"WHERE m.id IS NULL OR NOT ({predicate})", params,
    ).fetchone()[0]
    if excluded:
        raise ValueError(
            f"Existing Funes mapping contains {excluded:,} missing or excluded messages; "
            "refusing to omit or renumber mapped turns. Use a deliberately rebuilt mapping and fresh Funes memory."
        )


def remember_policy(db: sqlite3.Connection, policy: ExportPolicy) -> None:
    validate_policy(db, policy)
    db.execute("""
        CREATE TABLE IF NOT EXISTS funes_export_policy (
            singleton INTEGER PRIMARY KEY CHECK(singleton=1),
            policy_json TEXT NOT NULL
        )
    """)
    db.execute("INSERT OR IGNORE INTO funes_export_policy(singleton,policy_json) VALUES(1,?)",
               (policy.serialize(),))


def summarize(db: sqlite3.Connection, policy: ExportPolicy) -> dict:
    """Count selections without creating mapping tables or reading message content."""
    predicate, params = selection_sql(policy)
    total = db.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
    selected, missing_dates = db.execute(
        f"SELECT COUNT(*), COALESCE(SUM(NULLIF(TRIM(sent_utc),'') IS NULL "
        f"AND NULLIF(TRIM(received_utc),'') IS NULL),0) FROM messages m WHERE {predicate}", params,
    ).fetchone()
    categories = dict(db.execute(
        "SELECT COALESCE(NULLIF(TRIM(message_category),''),'Unclassified'), COUNT(*) "
        "FROM messages GROUP BY 1 ORDER BY 1"
    )) if policy.classified else {"Unclassified (legacy schema)": total}
    return {
        "total": total, "categories": categories, "selected": selected,
        "excluded": total - selected, "missing_dates": missing_dates,
        "exportable": selected - missing_dates,
        "mapped": db.execute("SELECT COUNT(*) FROM funes_map").fetchone()[0]
        if table_exists(db, "funes_map") else 0,
    }


def print_summary(summary: dict, policy: ExportPolicy) -> None:
    print(f"Selection: {policy.selection}; recipient source: {policy.recipient_source}; Bcc: {'included' if policy.include_bcc else 'omitted'}")
    if not policy.classified:
        print("Legacy database: message classification is unavailable; all records remain eligible. Run metadata recovery to distinguish emails from other Outlook items.")
    print(f"Database records: {summary['total']:,}; selected: {summary['selected']:,}; excluded: {summary['excluded']:,}")
    print(f"Selected records with missing dates: {summary['missing_dates']:,}; exportable by selection/date availability: {summary['exportable']:,}; existing mappings: {summary['mapped']:,}")
    for category, count in summary["categories"].items():
        print(f"  {category}: {count:,}")


def ensure_map(db: sqlite3.Connection):
    db.execute("""
        CREATE TABLE IF NOT EXISTS funes_map (
            message_id INTEGER NOT NULL PRIMARY KEY REFERENCES messages(id),
            session_id TEXT NOT NULL,
            turn_uuid TEXT NOT NULL,
            seq INTEGER NOT NULL,
            UNIQUE(session_id,seq),
            UNIQUE(session_id,turn_uuid)
        )
    """)
    db.execute("CREATE INDEX IF NOT EXISTS funes_map_session ON funes_map(session_id,seq)")
    db.commit()


def map_new_messages(db: sqlite3.Connection, policy: ExportPolicy | None = None) -> tuple[int, int]:
    """Assign a permanent seq exactly once. Late arrivals append to their session."""
    policy = policy or resolve_policy(db)
    remember_policy(db, policy)
    predicate, params = selection_sql(policy)
    next_seq = {sid: int(n) for sid, n in db.execute(
        "SELECT session_id, MAX(seq)+1 FROM funes_map GROUP BY session_id")}
    added, no_date = 0, 0
    for mid, key, msg_key, sent, received in db.execute(f"""
        SELECT m.id, m.conversation_key, m.message_key, m.sent_utc, m.received_utc
        FROM messages m LEFT JOIN funes_map fm ON m.id=fm.message_id
        WHERE fm.message_id IS NULL AND ({predicate})
        ORDER BY m.conversation_key, COALESCE(NULLIF(m.sent_utc,''),m.received_utc),m.id
    """, params):
        if not ((sent or "").strip() or (received or "").strip()):
            no_date += 1
            continue  # Funes requires a valid UTC timestamp; do not invent one.
        sid = key
        seq = next_seq.get(sid, 0)
        turn = msg_key  # already starts mid- or content-, contains no colons
        db.execute("INSERT INTO funes_map(message_id,session_id,turn_uuid,seq) VALUES(?,?,?,?)",
                   (mid, sid, turn, seq))
        next_seq[sid] = seq + 1
        added += 1
        if added % 5000 == 0:
            db.commit()
    db.commit()
    return added, no_date


def recipient_lines(db: sqlite3.Connection, message_id: int, to: str, cc: str,
                    policy: ExportPolicy) -> tuple[str, str, str]:
    """Prefer typed recipient rows and preserve names when no SMTP address exists."""
    typed: dict[str, list[str]] = {}
    if policy.recipient_source == "normalized":
        for kind, name, raw, smtp in db.execute(
            "SELECT recipient_type,display_name,address_raw,smtp_address "
            "FROM recovery_recipients WHERE message_id=? ORDER BY ordinal", (message_id,),
        ):
            kind = {"1": "to", "2": "cc", "3": "bcc"}.get(str(kind).lower(), str(kind).lower())
            if kind not in ("to", "cc", "bcc") or (kind == "bcc" and not policy.include_bcc):
                continue
            values = typed.setdefault(kind, [])
            address = (smtp or "").strip() or (raw or "").strip()
            name = (name or "").strip()
            value = f"{name} <{address}>" if name and address and name != address else address or name
            if value:
                values.append(value)
    # An existing type with empty values is still evidence from the recipient table.
    # Falling back only when that type is absent avoids reinstating stale display lists.
    to = "; ".join(typed["to"]) if "to" in typed else to or ""
    cc = "; ".join(typed["cc"]) if "cc" in typed else cc or ""
    bcc = "; ".join(typed.get("bcc", []))
    if policy.include_bcc and "bcc" not in typed and table_exists(db, "recovery_metadata"):
        row = db.execute("SELECT display_bcc FROM recovery_metadata WHERE message_id=?", (message_id,)).fetchone()
        bcc = (row[0] or "") if row else ""
    return to, cc, bcc


def turns(db: sqlite3.Connection, policy: ExportPolicy | None = None):
    """Read mapped messages ordered by session and permanent sequence."""
    policy = policy or resolve_policy(db)
    validate_policy(db, policy)
    category_sql = "m.message_category" if policy.classified else "''"
    sql = f"""
        SELECT m.id, fm.session_id, fm.turn_uuid, fm.seq,
               COALESCE(NULLIF(m.sent_utc,''), m.received_utc) AS ts,
               m.sender_name, m.sender_email, m.to_list, m.cc_list,
               m.subject, m.body, m.attachment_names_json, m.first_folder, {category_sql}
        FROM funes_map fm JOIN messages m ON m.id=fm.message_id
        ORDER BY fm.session_id, fm.seq
    """
    previous_session, expected_sequence = None, 0
    for mid, sid, turn, seq, ts, sender, email, to, cc, subject, body, att_json, folder, category in db.execute(sql):
        if sid != previous_session:
            previous_session, expected_sequence = sid, 0
        if seq != expected_sequence:
            raise ValueError(f"Non-dense Funes sequence in session {sid!r}; retain mapped rows or rebuild the export mapping deliberately")
        expected_sequence += 1
        if not ts or not ts.endswith("Z"):
            raise ValueError(f"Invalid UTC timestamp in message {turn!r}")
        datetime.fromisoformat(ts[:-1] + "+00:00")
        if ":" in sid or ":" in turn or not sid or not turn or seq < 0:
            raise ValueError(f"Invalid Funes identity in message {turn!r}")
        attachments = json.loads(att_json)
        if not isinstance(attachments, list) or any(not isinstance(name, str) for name in attachments):
            raise ValueError(f"Invalid attachment filename array in message {turn!r}")
        to, cc, bcc = recipient_lines(db, mid, to, cc, policy)
        header = [
            f"From: {sender} <{email}>" if sender else f"From: {email}",
            f"To: {to}",
            f"Cc: {cc}" if cc else None,
            f"Bcc: {bcc}" if bcc else None,
            f"Date (UTC): {ts}",
            f"Subject: {subject}",
            f"PST folder: {folder}",
            f"Outlook item category: {category}" if category else None,
            f"Attachment filenames: {', '.join(attachments)}" if attachments else None,
        ]
        lines = [line for line in header if line]
        text = "\n".join(lines) + "\n\n" + body
        yield {
            "format": 1,
            "session_id": sid,
            "turn_uuid": turn,
            "seq": seq,
            "ts": ts,
            "role": "member",
            "harness": "outlook-pst",
            "blocks": [{"block_type": "text", "text": text}],
        }


def check_destination(out_dir: Path) -> None:
    """Never mix snapshots or overwrite an older export."""
    if out_dir.exists() and (not out_dir.is_dir() or any(out_dir.iterdir())):
        raise ValueError(f"Output must be a new or empty directory: {out_dir}. Use a new snapshot folder.")


def export(db: sqlite3.Connection, out_dir: Path, per_file: int = 2000,
           policy: ExportPolicy | None = None) -> int:
    """Stage the whole snapshot, then publish its directory after success."""
    if per_file < 1:
        raise ValueError("per_file must be at least 1")
    check_destination(out_dir)
    out_dir.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=".email-funes-export-", dir=out_dir.parent))
    file_count, turn_count = 0, 0
    fh = None
    try:
        for turn in turns(db, policy):
            if turn_count % per_file == 0:
                if fh is not None:
                    fh.close()
                file_count += 1
                # Staging extensions prevent indexing unfinished export files.
                temp_path = staging / f"outlook-{file_count:06d}.tmp"
                fh = temp_path.open("w", encoding="utf-8", newline="\n")
            fh.write(json.dumps(turn, ensure_ascii=False, separators=(",", ":")) + "\n")
            turn_count += 1
        if fh is not None:
            fh.close()
        for temp_path in staging.glob("*.tmp"):
            temp_path.replace(temp_path.with_suffix(".funes.jsonl"))
        # Check again immediately before publishing; never remove existing data.
        check_destination(out_dir)
        if out_dir.exists():
            out_dir.rmdir()
        staging.replace(out_dir)
    finally:
        if fh is not None and not fh.closed:
            fh.close()
        if staging.exists():
            shutil.rmtree(staging)
    print(f"Exported {turn_count:,} turns to {file_count:,} .funes.jsonl files in {out_dir}")
    print("Stable turn IDs and sequence numbers are retained in SQLite funes_map.")
    return turn_count


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--db", required=True, type=Path, help="SQLite database generated by an email-funes importer")
    p.add_argument("--out", type=Path, help="New or empty directory for Funes import files (not needed for --dry-run)")
    p.add_argument("--dry-run", action="store_true", help="Read-only selection summary; do not create mappings or JSONL")
    selection = p.add_mutually_exclusive_group()
    selection.add_argument("--include-meetings", action="store_true", help="Include meeting messages alongside emails in enriched databases")
    selection.add_argument("--include-all", action="store_true", help="Include every Outlook item category, including unclassified records")
    p.add_argument("--include-bcc", action="store_true", help="Include recovered Bcc recipients in exported text (omitted by default)")
    p.add_argument("--turns-per-file", type=int, default=2000, help="Maximum turns per output file (default: 2000)")
    args = p.parse_args(argv)
    if not args.db.is_file() or args.turns_per_file < 1:
        p.error("--db must be an existing SQLite file, and --turns-per-file must be >= 1")
    if not args.dry_run and args.out is None:
        p.error("--out is required unless --dry-run is selected")
    if args.out is not None and args.out.resolve() == args.db.resolve():
        p.error("--out must be a directory, not the database file")
    try:
        if not args.dry_run:
            check_destination(args.out)
    except ValueError as exc:
        p.error(str(exc))
    # Opening mode=ro is also enforced by SQLite, not just by avoiding write calls.
    uri = args.db.resolve().as_uri() + ("?mode=ro" if args.dry_run else "?mode=rw")
    db = sqlite3.connect(uri, uri=True, timeout=120)
    db.execute("PRAGMA foreign_keys=ON")
    try:
        policy = resolve_policy(db, include_meetings=args.include_meetings,
                                include_all=args.include_all, include_bcc=args.include_bcc)
        print_summary(summarize(db, policy), policy)
        validate_policy(db, policy)
        if args.dry_run:
            print("Dry run complete. No database changes or JSONL files were made; message content and timestamps have not been fully validated.")
            return 0
        ensure_map(db)
        new, missing = map_new_messages(db, policy)
        print(f"Assigned new stable Funes turn IDs: {new:,}; skipped messages with missing dates: {missing:,}")
        export(db, args.out, args.turns_per_file, policy)
        return 0
    except (ValueError, sqlite3.Error) as exc:
        print(f"ERROR: {exc}")
        return 2
    finally:
        db.close()


if __name__ == "__main__":
    raise SystemExit(main())
