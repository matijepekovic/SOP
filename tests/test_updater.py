from __future__ import annotations

import io
import json
import tempfile
import unittest
import urllib.error
from pathlib import Path
from unittest import mock

from sop_reporter.exceptions import UpdateError
from sop_reporter.updater import Updater, is_newer, parse_version


class _FakeResponse(io.BytesIO):
    def __init__(self, payload: bytes, headers: dict | None = None) -> None:
        super().__init__(payload)
        self.headers = headers or {}

    def __enter__(self) -> "_FakeResponse":
        return self

    def __exit__(self, *_args) -> None:
        self.close()


def _release_payload(tag: str = "v0.4.0", size: int = 11, *, asset: bool = True) -> dict:
    payload: dict = {
        "tag_name": tag,
        "body": "Release notes",
        "html_url": f"https://github.com/owner/repo/releases/tag/{tag}",
        "assets": [],
    }
    if asset:
        payload["assets"] = [
            {
                "name": "SOPReporter.exe",
                "browser_download_url": "https://example.invalid/SOPReporter.exe",
                "size": size,
            }
        ]
    return payload


class VersionTests(unittest.TestCase):
    def test_parses_tags_with_and_without_prefix(self) -> None:
        self.assertEqual(parse_version("v0.3.1"), (0, 3, 1))
        self.assertEqual(parse_version("0.3.1"), (0, 3, 1))
        self.assertEqual(parse_version("1.2.3-beta"), (1, 2, 3))

    def test_rejects_unparseable_version(self) -> None:
        with self.assertRaises(UpdateError):
            parse_version("not-a-version")

    def test_ordering_is_numeric_not_lexicographic(self) -> None:
        self.assertTrue(is_newer("0.10.0", "0.9.0"))
        self.assertFalse(is_newer("0.3.1", "0.3.1"))
        self.assertFalse(is_newer("0.2.0", "0.3.1"))


class CheckTests(unittest.TestCase):
    def _updater(self, payload: dict, current: str = "0.3.1") -> Updater:
        opener = mock.Mock(
            return_value=_FakeResponse(json.dumps(payload).encode("utf-8"))
        )
        return Updater(
            repository="owner/repo",
            current_version=current,
            executable=Path("C:/app/SOPReporter.exe"),
            opener=opener,
        )

    def test_newer_release_is_reported(self) -> None:
        release = self._updater(_release_payload("v0.4.0")).check()
        assert release is not None
        self.assertEqual(release.version, "0.4.0")
        self.assertEqual(release.asset_name, "SOPReporter.exe")

    def test_same_version_reports_no_update(self) -> None:
        self.assertIsNone(self._updater(_release_payload("v0.3.1")).check())

    def test_older_release_reports_no_update(self) -> None:
        self.assertIsNone(self._updater(_release_payload("v0.2.0")).check())

    def test_release_without_the_asset_is_an_error(self) -> None:
        updater = self._updater(_release_payload("v0.4.0", asset=False))
        with self.assertRaises(UpdateError) as caught:
            updater.check()
        self.assertIn("does not attach", str(caught.exception))

    def test_missing_release_explains_how_to_fix_it(self) -> None:
        opener = mock.Mock(
            side_effect=urllib.error.HTTPError(
                "https://api.github.com", 404, "Not Found", {}, None
            )
        )
        updater = Updater(
            repository="owner/repo",
            current_version="0.3.1",
            executable=Path("C:/app/SOPReporter.exe"),
            opener=opener,
        )
        with self.assertRaises(UpdateError) as caught:
            updater.check()
        self.assertIn("No published release", str(caught.exception))

    def test_unreachable_github_is_reported_clearly(self) -> None:
        opener = mock.Mock(side_effect=urllib.error.URLError("offline"))
        updater = Updater(
            repository="owner/repo",
            current_version="0.3.1",
            executable=Path("C:/app/SOPReporter.exe"),
            opener=opener,
        )
        with self.assertRaises(UpdateError) as caught:
            updater.check()
        self.assertIn("Could not reach GitHub", str(caught.exception))


class DownloadAndApplyTests(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.folder = Path(self._temp.name)
        self.exe = self.folder / "SOPReporter.exe"
        self.exe.write_bytes(b"old-build")
        self.addCleanup(self._temp.cleanup)

    def _updater(self, opener, current: str = "0.3.1") -> Updater:
        return Updater(
            repository="owner/repo",
            current_version=current,
            executable=self.exe,
            opener=opener,
        )

    def test_download_writes_the_asset_next_to_the_executable(self) -> None:
        body = b"new-build!!"
        opener = mock.Mock(
            return_value=_FakeResponse(body, {"Content-Length": str(len(body))})
        )
        updater = self._updater(opener)
        release = updater._release_from_entry(_release_payload("v0.4.0", len(body)))
        staged = updater.download(release)
        self.assertTrue(staged.is_file())
        self.assertEqual(staged.read_bytes(), body)

    def test_truncated_download_is_rejected_and_cleaned_up(self) -> None:
        opener = mock.Mock(return_value=_FakeResponse(b"short"))
        updater = self._updater(opener)
        release = updater._release_from_entry(_release_payload("v0.4.0", size=999))
        with self.assertRaises(UpdateError) as caught:
            updater.download(release)
        self.assertIn("incomplete", str(caught.exception))
        self.assertFalse((self.folder / "SOPReporter.update.exe").exists())

    def test_apply_swaps_the_build_and_keeps_the_previous_one(self) -> None:
        staged = self.folder / "SOPReporter.update.exe"
        staged.write_bytes(b"new-build")
        updater = self._updater(mock.Mock())

        previous = updater.apply(staged)

        self.assertEqual(self.exe.read_bytes(), b"new-build")
        self.assertTrue(previous.is_file())
        self.assertEqual(previous.read_bytes(), b"old-build")
        self.assertFalse(staged.exists())

    def test_apply_without_a_staged_file_fails(self) -> None:
        updater = self._updater(mock.Mock())
        with self.assertRaises(UpdateError):
            updater.apply(self.folder / "missing.exe")

    def test_cleanup_removes_superseded_builds(self) -> None:
        (self.folder / "SOPReporter.previous-0.2.0.exe").write_bytes(b"x")
        (self.folder / "SOPReporter.previous-0.3.0.exe").write_bytes(b"x")
        updater = self._updater(mock.Mock())

        removed = updater.cleanup_previous_builds()

        self.assertEqual(removed, 2)
        self.assertTrue(self.exe.is_file())
        self.assertEqual(list(self.folder.glob("*.previous-*.exe")), [])

    def test_running_from_source_refuses_to_self_update(self) -> None:
        updater = self._updater(mock.Mock())
        with mock.patch.object(Updater, "is_frozen", property(lambda _self: False)):
            with self.assertRaises(UpdateError) as caught:
                updater.ensure_updatable()
        self.assertIn("running from source", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
