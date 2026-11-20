"""Synthetic EML regressions: attachment exclusion and truthful message dates."""
from email.message import EmailMessage
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import eml_to_funes as eml


class EMLTests(unittest.TestCase):
    def test_attachment_subtrees_are_not_indexed(self):
        outer = EmailMessage()
        outer.set_content('Body to preserve.')
        outer.add_attachment('Text attachment payload.', subtype='plain', filename='notes.txt')
        nested = EmailMessage()
        nested['Subject'] = 'Attached message'
        nested.set_content('Nested email attachment payload.')
        outer.add_attachment(nested, filename='forwarded.eml')
        self.assertEqual(eml.message_text(outer), 'Body to preserve.')
        names = [p.get_filename() for p in eml.visible_parts(outer) if p.get_filename()]
        self.assertEqual(names, ['notes.txt', 'forwarded.eml'])

    def test_date_timezone_is_preserved(self):
        self.assertEqual(eml.utc_timestamp('Tue, 03 Nov 2026 10:00:00 -0500'), '2026-11-03T15:00:00Z')

    def test_invalid_date_is_not_replaced_with_file_time(self):
        for value in ['', 'not a date', 'Tue, 03 Nov 2026 10:00:00']:
            with self.subTest(value=value), self.assertRaises(ValueError):
                eml.utc_timestamp(value)

    def test_reused_message_id_retains_body_variants(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'source'
            source.mkdir()
            for i in range(2):
                message = EmailMessage()
                message['Message-ID'] = '<same@example.com>'
                message['Date'] = 'Tue, 03 Nov 2026 10:00:00 -0500'
                message['From'] = 'sender@example.com'
                message['To'] = 'receiver@example.com'
                message['Subject'] = 'Synthetic variants'
                message.set_content(f'Variant {i}')
                (source / f'{i}.eml').write_bytes(message.as_bytes())
            target = Path(directory) / 'output'
            self.assertEqual(eml.convert(source, target, None, None, 0), 0)
            lines = [line for path in target.glob('*.funes.jsonl') for line in path.read_text(encoding='utf-8').splitlines()]
            self.assertEqual(len(lines), 2)


if __name__ == '__main__':
    unittest.main()
