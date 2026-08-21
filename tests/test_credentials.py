from __future__ import annotations

import unittest

from sop_reporter.credentials import SERVICE_NAME, CredentialManager
from sop_reporter.exceptions import CredentialsCancelledError
from sop_reporter.gui.first_run import EnteredCredentials


class FakeKeyring:
    def __init__(self, stored: dict[tuple[str, str], str] | None = None) -> None:
        self.stored = dict(stored or {})
        self.writes = 0

    def get_password(self, service, account):
        return self.stored.get((service, account))

    def set_password(self, service, account, password):
        self.stored[(service, account)] = password
        self.writes += 1

    def delete_password(self, service, account):
        del self.stored[(service, account)]


class ScriptedPrompt:
    """Returns queued answers, recording the reasons it was re-shown."""

    def __init__(self, *answers: tuple[str, str]) -> None:
        self.answers = list(answers)
        self.problems: list[str] = []
        self.calls = 0

    def __call__(self, account: str = "", problem: str = "") -> EnteredCredentials:
        self.calls += 1
        self.problems.append(problem)
        if not self.answers:
            raise CredentialsCancelledError("cancelled")
        email, password = self.answers.pop(0)
        return EnteredCredentials(account=email, app_password=password)


GOOD = "abcd efgh ijkl mnop"
BAD = "wrongwrongwrong"


def only_good(_account: str, password: str) -> str | None:
    if "".join(password.split()) == "".join(GOOD.split()):
        return None
    return "Gmail rejected this address and app password."


class CredentialTests(unittest.TestCase):
    def test_stored_password_is_reused_without_prompting(self) -> None:
        keyring = FakeKeyring({(SERVICE_NAME, "a@b.com"): GOOD})
        prompt = ScriptedPrompt()
        manager = CredentialManager(prompt=prompt, keyring_module=keyring)

        credentials = manager.ensure_credentials("a@b.com")

        self.assertEqual(credentials.account, "a@b.com")
        self.assertEqual(prompt.calls, 0)

    def test_a_rejected_password_is_never_stored_and_prompts_again(self) -> None:
        keyring = FakeKeyring()
        prompt = ScriptedPrompt(("a@b.com", BAD), ("a@b.com", GOOD))
        manager = CredentialManager(prompt=prompt, keyring_module=keyring)

        credentials = manager.ensure_credentials("", verifier=only_good)

        # Asked twice, and the second prompt explained why it reopened.
        self.assertEqual(prompt.calls, 2)
        self.assertEqual(prompt.problems[0], "")
        self.assertIn("rejected", prompt.problems[1])
        # Only the working password was ever written.
        self.assertEqual(keyring.writes, 1)
        self.assertEqual(keyring.stored[(SERVICE_NAME, "a@b.com")], "abcdefghijklmnop")
        self.assertEqual(credentials.app_password, "abcdefghijklmnop")

    def test_cancelling_the_retry_stores_nothing(self) -> None:
        keyring = FakeKeyring()
        prompt = ScriptedPrompt(("a@b.com", BAD))
        manager = CredentialManager(prompt=prompt, keyring_module=keyring)

        with self.assertRaises(CredentialsCancelledError):
            manager.ensure_credentials("", verifier=only_good)

        self.assertEqual(keyring.writes, 0)
        self.assertEqual(keyring.stored, {})

    def test_force_prompt_replaces_a_stored_password(self) -> None:
        keyring = FakeKeyring({(SERVICE_NAME, "a@b.com"): "stale"})
        prompt = ScriptedPrompt(("a@b.com", GOOD))
        manager = CredentialManager(prompt=prompt, keyring_module=keyring)

        manager.ensure_credentials("a@b.com", verifier=only_good, force_prompt=True)

        self.assertEqual(prompt.calls, 1)
        self.assertEqual(keyring.stored[(SERVICE_NAME, "a@b.com")], "abcdefghijklmnop")

    def test_a_failed_change_leaves_the_working_password_in_place(self) -> None:
        keyring = FakeKeyring({(SERVICE_NAME, "a@b.com"): "abcdefghijklmnop"})
        prompt = ScriptedPrompt(("a@b.com", BAD))
        manager = CredentialManager(prompt=prompt, keyring_module=keyring)

        with self.assertRaises(CredentialsCancelledError):
            manager.ensure_credentials("a@b.com", verifier=only_good, force_prompt=True)

        # The credential that still works was not overwritten or removed.
        self.assertEqual(keyring.stored[(SERVICE_NAME, "a@b.com")], "abcdefghijklmnop")

    def test_clear_forgets_the_stored_password(self) -> None:
        keyring = FakeKeyring({(SERVICE_NAME, "a@b.com"): GOOD})
        prompt = ScriptedPrompt(("a@b.com", GOOD))
        manager = CredentialManager(prompt=prompt, keyring_module=keyring)

        manager.clear("a@b.com")
        self.assertEqual(keyring.stored, {})

        # Clearing an account with nothing stored is not an error.
        manager.clear("a@b.com")

        manager.ensure_credentials("a@b.com", verifier=only_good)
        self.assertEqual(prompt.calls, 1)


if __name__ == "__main__":
    unittest.main()
