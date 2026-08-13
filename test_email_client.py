from __future__ import annotations

import tempfile
import unittest
from datetime import date
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import MagicMock

from sop_reporter.config import (
    AttachmentConfig,
    EmailConfig,
    EmailSearchConfig,
)
from sop_reporter.email_client import GmailIMAPClient


def _message_bytes(message_id: str, filename: str, payload: bytes) -> bytes:
    message = EmailMessage()
    message["From"] = "reports@example.com"
    message["To"] = "sop@example.com"
    message["Subject"] = "Daily SOP"
    message["Message-ID"] = message_id
    message.set_content("Attached report")
    message.add_attachment(
        payload,
        maintype="application",
        subtype="octet-stream",
        filename=filename,
    )
    return message.as_bytes()


class GmailIMAPClientTests(unittest.TestCase):
    def setUp(self) -> None:
        self.config = EmailConfig(
            account="sop@example.com",
            imap_host="imap.gmail.com",
            imap_port=993,
            mailbox="INBOX",
            search=EmailSearchConfig(
                sender="reports@example.com",
                subject_contains="Daily SOP",
                since_days=14,
                unread_only=False,
            ),
            attachments=AttachmentConfig(("*.xlsx",)),
        )

    def test_search_fetch_mime_and_attachment_download(self) -> None:
        imap = MagicMock()
        imap.login.return_value = ("OK", [b"logged in"])
        imap.select.return_value = ("OK", [b"2"])
        imap.search.return_value = ("OK", [b"1 2"])
        imap.fetch.side_effect = [
            ("OK", [(b"1 (BODY[])", _message_bytes("<one@example>", "daily.xlsx", b"xlsx-data")), b")"]),
            ("OK", [(b"2 (BODY[])", _message_bytes("<two@example>", "notes.txt", b"text")), b")"]),
        ]
        factory = MagicMock(return_value=imap)
        client = GmailIMAPClient(
            self.config,
            account="sop@example.com",
            app_password="abcdefghijklmnop",
            imap_factory=factory,
        )

        with tempfile.TemporaryDirectory() as directory:
            records = client.fetch_messages(
                Path(directory),
                today=date(2026, 8, 13),
            )
            self.assertEqual(len(records), 2)
            self.assertEqual(len(records[0].attachments), 1)
            self.assertEqual(records[0].attachments[0].path.read_bytes(), b"xlsx-data")
            self.assertEqual(len(records[1].attachments), 0)

        factory.assert_called_once_with("imap.gmail.com", 993, timeout=30)
        search_args = imap.search.call_args.args
        self.assertIn("SINCE", search_args)
        self.assertIn("30-Jul-2026", search_args)
        self.assertIn("FROM", search_args)
        self.assertIn("SUBJECT", search_args)
        imap.select.assert_called_once_with("INBOX", readonly=True)
        imap.close.assert_called_once()
        imap.logout.assert_called_once()

    def test_handled_message_is_not_downloaded(self) -> None:
        imap = MagicMock()
        imap.login.return_value = ("OK", [])
        imap.select.return_value = ("OK", [])
        imap.search.return_value = ("OK", [b"1"])
        imap.fetch.return_value = (
            "OK",
            [(b"1 (BODY[])", _message_bytes("<one@example>", "daily.xlsx", b"data"))],
        )
        client = GmailIMAPClient(
            self.config,
            "sop@example.com",
            "abcdefghijklmnop",
            imap_factory=MagicMock(return_value=imap),
        )
        with tempfile.TemporaryDirectory() as directory:
            records = client.fetch_messages(
                Path(directory),
                handled_message_ids={"<one@example>"},
            )
            self.assertEqual(records, [])
            self.assertEqual(list(Path(directory).iterdir()), [])

    def test_blank_filters_fall_back_to_latest_inbox_messages(self) -> None:
        config = EmailConfig(
            account="sheetssop@gmail.com",
            imap_host="imap.gmail.com",
            imap_port=993,
            mailbox="INBOX",
            search=EmailSearchConfig(
                sender="",
                subject_contains="",
                since_days=14,
                unread_only=False,
            ),
            attachments=AttachmentConfig(("*.xlsx",)),
        )
        imap = MagicMock()
        imap.login.return_value = ("OK", [])
        imap.select.return_value = ("OK", [b"3"])
        imap.search.side_effect = [
            ("OK", [b""]),
            ("OK", [b""]),
            ("OK", [b"1 2 3"]),
        ]
        imap.fetch.side_effect = [
            ("OK", [(b"1", _message_bytes("<one>", "one.xlsx", b"1"))]),
            ("OK", [(b"2", _message_bytes("<two>", "two.xlsx", b"2"))]),
            ("OK", [(b"3", _message_bytes("<three>", "three.xlsx", b"3"))]),
        ]
        client = GmailIMAPClient(
            config,
            "sheetssop@gmail.com",
            "abcdefghijklmnop",
            imap_factory=MagicMock(return_value=imap),
        )

        with tempfile.TemporaryDirectory() as directory:
            records = client.fetch_messages(Path(directory), today=date(2026, 8, 13))

        self.assertEqual(len(records), 3)
        self.assertEqual(imap.search.call_count, 3)
        self.assertEqual(imap.search.call_args_list[-1].args, (None, "ALL"))


if __name__ == "__main__":
    unittest.main()
