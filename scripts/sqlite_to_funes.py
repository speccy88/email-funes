#!/usr/bin/env python3
"""Export unique emails from pst_to_sqlite.py to Funes .funes.jsonl files.

Maintains a stable per-conversation sequence map in SQLite, so re-exporting
and importing new mail won't renumber old turns. Run after PST import.

Usage:
  py sqlite_to_funes.py --db "C:\\path\\mail.sqlite" --out "C:\\path\\funes_exports"
  funes index --check /path/to/funes_exports
  funes index /path/to/funes_exports
"""
from __future__ import annotations

import argparse
from datetime import datetime
import json
import sqlite3
import shutil
import tempfile
from pathlib import Path


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


def map_new_messages(db: sqlite3.Connection) -> tuple[int, int]:
    """Assign a permanent seq exactly once. Late arrivals append to their session."""
    next_seq = {sid: int(n) for sid, n in db.execute(
        "SELECT session_id, MAX(seq)+1 FROM funes_map GROUP BY session_id")}
    added, no_date = 0, 0
    for mid, key, msg_key, sent, received in db.execute("""
        SELECT m.id, m.conversation_key, m.message_key, m.sent_utc, m.received_utc
        FROM messages m LEFT JOIN funes_map fm ON m.id=fm.message_id
        WHERE fm.message_id IS NULL
        ORDER BY m.conversation_key, COALESCE(m.sent_utc,m.received_utc),m.id
    """):
        if not (sent or received):
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


def turns(db: sqlite3.Connection):
    """Read mapped messages ordered by session and permanent sequence."""
    sql = """
        SELECT fm.session_id, fm.turn_uuid, fm.seq,
               COALESCE(m.sent_utc, m.received_utc) AS ts,
               m.sender_name, m.sender_email, m.to_list, m.cc_list,
               m.subject, m.body, m.attachment_names_json, m.first_folder
        FROM funes_map fm JOIN messages m ON m.id=fm.message_id
        ORDER BY fm.session_id, fm.seq
    """
    previous_session, expected_sequence = None, 0
    for sid, turn, seq, ts, sender, email, to, cc, subject, body, att_json, folder in db.execute(sql):
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
        header = [
            f"From: {sender} <{email}>" if sender else f"From: {email}",
            f"To: {to}",
            f"Cc: {cc}" if cc else None,
            f"Date (UTC): {ts}",
            f"Subject: {subject}",
            f"PST folder: {folder}",
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


def export(db: sqlite3.Connection, out_dir: Path, per_file: int = 2000) -> int:
    """Stage the whole snapshot, then publish its directory after success."""
    if per_file < 1:
        raise ValueError("per_file must be at least 1")
    check_destination(out_dir)
    out_dir.parent.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=".email-funes-export-", dir=out_dir.parent))
    file_count, turn_count = 0, 0
    fh = None
    try:
        for turn in turns(db):
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
    p.add_argument("--out", required=True, type=Path, help="Directory for Funes import files")
    p.add_argument("--turns-per-file", type=int, default=2000, help="Maximum turns per output file (default: 2000)")
    args = p.parse_args(argv)
    if not args.db.is_file() or args.turns_per_file < 1:
        p.error("--db must be an existing SQLite file, and --turns-per-file must be >= 1")
    if args.out.resolve() == args.db.resolve():
        p.error("--out must be a directory, not the database file")
    try:
        check_destination(args.out)
    except ValueError as exc:
        p.error(str(exc))
    db = sqlite3.connect(str(args.db), timeout=120)
    db.execute("PRAGMA foreign_keys=ON")
    try:
        ensure_map(db)
        new, missing = map_new_messages(db)
        print(f"Assigned new stable Funes turn IDs: {new:,}; skipped messages with missing dates: {missing:,}")
        export(db, args.out, args.turns_per_file)
        return 0
    finally:
        db.close()


if __name__ == "__main__":
    raise SystemExit(main())
