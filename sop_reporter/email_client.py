from __future__ import annotations

import fnmatch
import hashlib
import imaplib
import logging
import os
import re
from dataclasses import dataclass
from datetime import date, timedelta
from email import policy
from email.message import Message
from email.parser import BytesParser
from pathlib import Path
from typing import Callable, Iterable

from sop_reporter.config import EmailConfig
from sop_reporter.exceptions import EmailClientError


LOGGER = logging.getLogger(__name__)
INVALID_FILENAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass(frozen=True)
class DownloadedAttachment:
    original_filename: str
    path: Path
    content_type: str
    size: int


@dataclass(frozen=True)
class EmailRecord:
    sequence_id: str
    message_id: str
    subject: str
    sender: str
    sent_date: str
    attachments: tuple[DownloadedAttachment, ...]


def _safe_filename(filename: str) -> str:
    leaf = filename.replace("\\", "/").split("/")[-1]
    cleaned = INVALID_FILENAME_CHARACTERS.sub("_", leaf).strip(" .")
    return cleaned[:180] or "attachment.xlsx"


def _quote_search_value(value: str) -> str:
    return value.replace("\\", "").replace('"', "")


def verify_credentials(config: EmailConfig, account: str, app_password: str) -> str | None:
    """Try a Gmail login. Returns None when accepted, else why it was refused.

    Used before storing credentials so a wrong app password is never saved.
    A network problem is reported separately from a rejected password, because
    only the latter means the password itself is wrong.
    """
    connection = None
    try:
        connection = imaplib.IMAP4_SSL(config.imap_host, config.imap_port, timeout=30)
    except OSError as exc:
        return (
            f"Could not reach {config.imap_host}: {exc}. "
            "Check the internet connection and try again."
        )
    try:
        status, _ = connection.login(account, app_password)
        if status != "OK":
            return "Gmail rejected this address and app password."
        return None
    except imaplib.IMAP4.error:
        return (
            "Gmail rejected this address and app password. Use a 16-character "
            "app password generated at myaccount.google.com/apppasswords, not "
            "the normal Gmail password."
        )
    except OSError as exc:
        return f"The connection to Gmail failed: {exc}"
    finally:
        try:
            connection.logout()
        except Exception:
            pass


class GmailIMAPClient:
    def __init__(
        self,
        config: EmailConfig,
        account: str,
        app_password: str,
        imap_factory: Callable[..., imaplib.IMAP4_SSL] = imaplib.IMAP4_SSL,
    ) -> None:
        self.config = config
        self.account = account
        self._app_password = app_password
        self._imap_factory = imap_factory

    def build_search_criteria(self, today: date | None = None) -> tuple[str, ...]:
        reference = today or date.today()
        since = reference - timedelta(days=self.config.search.since_days)
        criteria: list[str] = ["SINCE", since.strftime("%d-%b-%Y")]
        if self.config.search.sender:
            criteria.extend(
                ["FROM", f'"{_quote_search_value(self.config.search.sender)}"']
            )
        if self.config.search.subject_contains:
            criteria.extend(
                [
                    "SUBJECT",
                    f'"{_quote_search_value(self.config.search.subject_contains)}"',
                ]
            )
        if self.config.search.unread_only:
            criteria.append("UNSEEN")
        return tuple(criteria)

    def build_gmail_raw_query(self) -> str:
        parts = [f"newer_than:{self.config.search.since_days}d"]
        if self.config.search.sender:
            parts.append(f'from:("{_quote_search_value(self.config.search.sender)}")')
        if self.config.search.subject_contains:
            parts.append(
                f'subject:("{_quote_search_value(self.config.search.subject_contains)}")'
            )
        if self.config.search.unread_only:
            parts.append("is:unread")
        return " ".join(parts)

    def fetch_messages(
        self,
        download_dir: Path,
        handled_message_ids: Iterable[str] = (),
        today: date | None = None,
    ) -> list[EmailRecord]:
        handled = set(handled_message_ids)
        download_dir.mkdir(parents=True, exist_ok=True)
        connection: imaplib.IMAP4_SSL | None = None
        selected = False
        records: list[EmailRecord] = []
        try:
            connection = self._imap_factory(
                self.config.imap_host,
                self.config.imap_port,
                timeout=30,
            )
            try:
                status, _ = connection.login(self.account, self._app_password)
            except imaplib.IMAP4.error as exc:
                # Gmail reports a bad app password by raising here rather than
                # returning a non-OK status, so it has to be caught at the call
                # or the generic handler below strips its meaning.
                raise EmailClientError(
                    f"Gmail rejected the stored app password for {self.account}: {exc}",
                    authentication_failed=True,
                ) from exc
            if status != "OK":
                raise EmailClientError(
                    f"Gmail rejected the stored app password for {self.account}",
                    authentication_failed=True,
                )
            status, select_data = connection.select(self.config.mailbox, readonly=True)
            if status != "OK":
                raise EmailClientError(
                    f"Gmail mailbox could not be opened: {self.config.mailbox}"
                )
            selected = True

            mailbox_count = "unknown"
            if select_data and select_data[0] is not None:
                raw_count = select_data[0]
                mailbox_count = (
                    raw_count.decode("ascii", errors="replace")
                    if isinstance(raw_count, bytes)
                    else str(raw_count)
                )
            criteria = self.build_search_criteria(today=today)
            LOGGER.info(
                "Gmail account %s mailbox %s reports %s total message(s); search=%s",
                self.account,
                self.config.mailbox,
                mailbox_count,
                " ".join(criteria),
            )

            status, search_data = connection.search(
                None, *criteria
            )
            if status != "OK":
                raise EmailClientError("Gmail IMAP search failed")
            sequence_ids = search_data[0].split() if search_data and search_data[0] else []

            # Gmail occasionally returns an empty result for a valid RFC IMAP
            # SINCE search even while the web Inbox visibly contains mail. Retry
            # using Gmail's native search extension before declaring the Inbox empty.
            if not sequence_ids:
                raw_query = self.build_gmail_raw_query()
                LOGGER.warning(
                    "Standard Gmail IMAP search returned zero; retrying with X-GM-RAW %s",
                    raw_query,
                )
                status, search_data = connection.search(
                    None,
                    "X-GM-RAW",
                    f'"{raw_query}"',
                )
                if status != "OK":
                    raise EmailClientError("Gmail native search failed")
                sequence_ids = (
                    search_data[0].split() if search_data and search_data[0] else []
                )

            # During setup, blank sender/subject filters mean "show me the Inbox."
            # If both date-search forms fail, inspect the latest messages directly.
            if (
                not sequence_ids
                and not self.config.search.sender
                and not self.config.search.subject_contains
                and not self.config.search.unread_only
            ):
                LOGGER.warning(
                    "Both date searches returned zero; falling back to latest Inbox messages"
                )
                status, search_data = connection.search(None, "ALL")
                if status != "OK":
                    raise EmailClientError("Gmail ALL fallback search failed")
                all_ids = search_data[0].split() if search_data and search_data[0] else []
                sequence_ids = all_ids[-250:]
            LOGGER.info("Gmail search returned %d message(s)", len(sequence_ids))

            for raw_sequence_id in sequence_ids:
                sequence_id = raw_sequence_id.decode("ascii", errors="replace")
                status, fetched = connection.fetch(raw_sequence_id, "(BODY.PEEK[])")
                if status != "OK":
                    raise EmailClientError(
                        f"Gmail failed to fetch message sequence {sequence_id}"
                    )
                raw_message = self._extract_raw_message(fetched)
                parsed = BytesParser(policy=policy.default).parsebytes(raw_message)
                message_id = self._message_id(parsed, raw_message)
                if message_id in handled:
                    LOGGER.debug("Skipping handled Gmail message %s", message_id)
                    continue
                attachments = self._download_attachments(
                    parsed,
                    message_id=message_id,
                    download_dir=download_dir,
                )
                records.append(
                    EmailRecord(
                        sequence_id=sequence_id,
                        message_id=message_id,
                        subject=str(parsed.get("Subject", "")),
                        sender=str(parsed.get("From", "")),
                        sent_date=str(parsed.get("Date", "")),
                        attachments=tuple(attachments),
                    )
                )
            return records
        except EmailClientError:
            raise
        except (imaplib.IMAP4.error, OSError) as exc:
            raise EmailClientError(f"Gmail IMAP operation failed: {exc}") from exc
        except Exception as exc:
            raise EmailClientError(f"Unexpected Gmail processing failure: {exc}") from exc
        finally:
            if connection is not None:
                if selected:
                    try:
                        connection.close()
                    except Exception:
                        LOGGER.debug("IMAP mailbox close failed", exc_info=True)
                try:
                    connection.logout()
                except Exception:
                    LOGGER.debug("IMAP logout failed", exc_info=True)

    @staticmethod
    def _extract_raw_message(fetched: object) -> bytes:
        if not isinstance(fetched, list):
            raise EmailClientError("Gmail returned an unexpected FETCH response")
        for item in fetched:
            if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
                return item[1]
        raise EmailClientError("Gmail FETCH response did not contain a message body")

    @staticmethod
    def _message_id(message: Message, raw_message: bytes) -> str:
        header = str(message.get("Message-ID", "")).strip()
        if header:
            return header
        return "sha256:" + hashlib.sha256(raw_message).hexdigest()

    def _download_attachments(
        self,
        message: Message,
        message_id: str,
        download_dir: Path,
    ) -> list[DownloadedAttachment]:
        token = hashlib.sha256(message_id.encode("utf-8")).hexdigest()[:12]
        downloaded: list[DownloadedAttachment] = []
        attachment_index = 0
        for part in message.walk():
            if part.is_multipart():
                continue
            filename = part.get_filename()
            if not filename:
                continue
            if not any(
                fnmatch.fnmatch(filename.casefold(), pattern.casefold())
                for pattern in self.config.attachments.filename_patterns
            ):
                continue
            payload = part.get_payload(decode=True)
            if payload is None:
                LOGGER.warning("Skipping attachment with no payload: %s", filename)
                continue
            attachment_index += 1
            safe_name = _safe_filename(filename)
            destination = download_dir / f"{token}_{attachment_index:02d}_{safe_name}"
            temporary = destination.with_suffix(destination.suffix + ".part")
            try:
                with temporary.open("wb") as handle:
                    handle.write(payload)
                    handle.flush()
                    os.fsync(handle.fileno())
                os.replace(temporary, destination)
            except OSError as exc:
                try:
                    temporary.unlink(missing_ok=True)
                except OSError:
                    pass
                raise EmailClientError(
                    f"Could not save Gmail attachment {filename}: {exc}"
                ) from exc
            downloaded.append(
                DownloadedAttachment(
                    original_filename=filename,
                    path=destination,
                    content_type=part.get_content_type(),
                    size=len(payload),
                )
            )
            LOGGER.info("Downloaded attachment %s", filename)
        return downloaded
