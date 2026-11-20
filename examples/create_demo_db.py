#!/usr/bin/env python3
"""Create a small synthetic email database for testing the SQLite/Funes workflow.

Run from the repository root:
    python examples/create_demo_db.py --db data/demo.sqlite

No PST reader, Outlook profile, network connection, or personal email is needed.
"""
from __future__ import annotations

import argparse
from dataclasses import replace
import json
from pathlib import Path
import sys

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from email_funes_common import Message, message_fingerprint, open_database, sha, show_stats, source_seen, store_message


def demo_message(
    number: int,
    subject: str,
    body: str,
    timestamp: str,
    sender: str,
    recipient: str,
    conversation: str,
    attachments: tuple[str, ...] = (),
) -> Message:
    """Return one fully synthetic record compatible with the common schema."""
    internet_id = f"demo-{number}@example.com"
    sender_name = "Engineering Demo" if sender.startswith("engineering") else "Planning Demo"
    fingerprint = message_fingerprint(
        internet_message_id=internet_id, sender_name=sender_name,
        sender_email=sender, to_list=recipient, cc_list="", timestamp=timestamp,
        subject=subject, full_body=body, attachment_names=list(attachments),
        attachment_count=len(attachments),
    )
    return Message(
        source_key=sha(f"synthetic-demo-v1/inbox/{number}"),
        message_key="mid-" + sha(internet_id + "\0" + fingerprint),
        fingerprint=fingerprint,
        internet_message_id=internet_id,
        conversation_key=conversation,
        sent_utc=timestamp,
        received_utc=timestamp,
        sender_name=sender_name,
        sender_email=sender,
        to_list=recipient,
        cc_list="",
        subject=subject,
        body=body,
        attachment_names_json=json.dumps(attachments),
        attachment_count=len(attachments),
        folder="Synthetic archive/Inbox",
    )


def demo_records(enriched: bool = False) -> list[Message]:
    first = demo_message(
        1,
        "Demo pump inspection: observations",
        "Synthetic example. Inspection found a worn coupling. "
        "Replace the coupling and inspect the alignment before the next trial.",
        "2026-11-02T14:00:00Z",
        "engineering@example.com",
        "planning@example.com",
        "demo-pump-inspection",
        ("inspection-notes.pdf",),
    )
    second = demo_message(
        2,
        "Re: Demo pump inspection: observations",
        "Synthetic example. The coupling was replaced and the alignment checked. "
        "The trial is scheduled for November 4. Décision : conserver le compte rendu.",
        "2026-11-03T15:30:00Z",
        "planning@example.com",
        "engineering@example.com",
        "demo-pump-inspection",
    )
    third = demo_message(
        3,
        "Demo ventilation project: document review",
        "Synthetic example. The draft schedule needs one final review. "
        "No installation date has been approved.",
        "2026-11-05T09:00:00Z",
        "engineering@example.com",
        "planning@example.com",
        "demo-ventilation-project",
        ("draft-schedule.xlsx",),
    )
    duplicate = replace(
        first,
        source_key=sha("synthetic-demo-v1/project-folder/1"),
        folder="Synthetic archive/Project copies",
    )
    records = [first, second, third, duplicate]
    if enriched:
        records.extend([
            demo_message(
                4, "Demo pump inspection: meeting request",
                "Synthetic meeting request. Review the inspection findings on November 3.",
                "2026-11-02T16:00:00Z", "planning@example.com", "engineering@example.com",
                "demo-pump-inspection",
            ),
            demo_message(
                5, "Demo pump inspection: calendar appointment",
                "Synthetic calendar item. Equipment review in the workshop.",
                "2026-11-03T16:00:00Z", "planning@example.com", "engineering@example.com",
                "demo-calendar-appointment",
            ),
        ])
    return records


def add_demo_metadata(db) -> None:
    """Populate the recovery schema with synthetic metadata; no PST is opened."""
    from pst_sqlite_recover import ensure_schema

    ensure_schema(db)
    for mid, internet_id, sender, to in db.execute(
        "SELECT id,internet_message_id,sender_email,to_list FROM messages"
    ):
        if internet_id == "demo-4@example.com":
            klass, category = "IPM.Schedule.Meeting.Request", "Meeting request"
        elif internet_id == "demo-5@example.com":
            klass, category = "IPM.Appointment", "Appointment"
        else:
            klass, category = "IPM.Note", "Email / IPM.Note"
        recipients = [("to", "Recipient Demo", to)]
        if internet_id == "demo-1@example.com":
            recipients.extend([
                ("cc", "Maintenance Demo", "maintenance@example.com"),
                ("bcc", "Archive Demo", "archive@example.com"),
            ])
        db.execute(
            "UPDATE messages SET message_class=?,message_category=?,sender_raw_address=?,"
            "sender_email_source='synthetic',recipient_to_source='synthetic',"
            "body_source='synthetic' WHERE id=?", (klass, category, sender, mid),
        )
        db.execute(
            "INSERT INTO recovery_metadata(message_id,mapi_sender_raw,mapi_sender_smtp,"
            "display_to,recovered_internet_id,recipient_count) VALUES(?,?,?,?,?,?)",
            (mid, sender, sender, to, internet_id, len(recipients)),
        )
        db.executemany(
            "INSERT INTO recovery_recipients(message_id,ordinal,recipient_type,display_name,"
            "address_type,address_raw,smtp_address) VALUES(?,?,?,?,'SMTP',?,?)",
            [(mid, ordinal, kind, name, address, address)
             for ordinal, (kind, name, address) in enumerate(recipients)],
        )
    db.commit()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--db", type=Path, default=Path("data/demo.sqlite"))
    parser.add_argument("--enriched", action="store_true",
                        help="Create a new demo with recovered metadata and non-email Outlook items")
    args = parser.parse_args(argv)
    database = args.db.expanduser().resolve()
    if args.enriched and database.exists():
        parser.error("--enriched needs a new demo database path")
    db = open_database(database)
    added = duplicates = already_seen = 0
    try:
        for record in demo_records(args.enriched):
            if source_seen(db, record.source_key):
                already_seen += 1
                continue
            result = store_message(db, record)
            if result == "kept":
                added += 1
            else:
                duplicates += 1
        db.commit()
        if args.enriched:
            add_demo_metadata(db)
        print(f"Synthetic demo: new={added}; duplicate occurrences={duplicates}; already seen={already_seen}")
        print(f"Database: {database}")
        show_stats(db)
    finally:
        db.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
