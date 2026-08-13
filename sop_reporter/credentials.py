from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable

from sop_reporter.exceptions import CredentialsError
from sop_reporter.gui.first_run import EnteredCredentials, prompt_for_credentials


SERVICE_NAME = "SOPReporter Gmail"


@dataclass(frozen=True)
class GmailCredentials:
    account: str
    app_password: str = field(repr=False)


class CredentialManager:
    def __init__(
        self,
        prompt: Callable[[str], EnteredCredentials] = prompt_for_credentials,
        keyring_module: Any | None = None,
    ) -> None:
        self._prompt = prompt
        self._keyring = keyring_module

    def _get_keyring(self) -> Any:
        if self._keyring is not None:
            return self._keyring
        try:
            import keyring
        except ImportError as exc:
            raise CredentialsError(
                "The keyring package is not installed. Reinstall SOP Reporter or its requirements."
            ) from exc
        self._keyring = keyring
        return keyring

    def ensure_credentials(self, configured_account: str) -> GmailCredentials:
        keyring_module = self._get_keyring()
        account = configured_account.strip()
        if account:
            try:
                stored = keyring_module.get_password(SERVICE_NAME, account)
            except Exception as exc:
                raise CredentialsError(
                    "Windows Credential Manager could not be read"
                ) from exc
            if stored:
                return GmailCredentials(
                    account=account,
                    app_password="".join(stored.split()),
                )

        entered = self._prompt(account)
        normalized_password = "".join(entered.app_password.split())
        try:
            keyring_module.set_password(
                SERVICE_NAME,
                entered.account,
                normalized_password,
            )
        except Exception as exc:
            raise CredentialsError(
                "Windows Credential Manager could not store the Gmail app password"
            ) from exc
        return GmailCredentials(
            account=entered.account,
            app_password=normalized_password,
        )
