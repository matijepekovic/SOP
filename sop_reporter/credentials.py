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

    def clear(self, account: str) -> None:
        """Forget the stored app password so the next run asks again."""
        account = account.strip()
        if not account:
            return
        keyring_module = self._get_keyring()
        try:
            keyring_module.delete_password(SERVICE_NAME, account)
        except Exception:
            # Nothing stored for this account, which is the desired end state.
            pass

    def ensure_credentials(
        self,
        configured_account: str,
        verifier: Callable[[str, str], str | None] | None = None,
        *,
        force_prompt: bool = False,
    ) -> GmailCredentials:
        """Return working Gmail credentials, asking for them if needed.

        ``verifier`` is given the account and password and returns None when
        Gmail accepts them, or a sentence explaining the refusal. Nothing is
        written to Credential Manager until it succeeds, so a mistyped app
        password can never be stored and become impossible to change.
        """
        keyring_module = self._get_keyring()
        account = configured_account.strip()
        if account and not force_prompt:
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

        problem = ""
        while True:
            entered = self._prompt(account, problem) if problem else self._prompt(account)
            normalized_password = "".join(entered.app_password.split())
            account = entered.account

            if verifier is not None:
                problem = verifier(entered.account, normalized_password) or ""
                if problem:
                    # Ask again rather than storing something that does not work.
                    continue

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
