"""Desktop self-update against GitHub Releases.

The published executable checks the repository's latest release, compares it
to the running version, and can download and swap itself in place.

Windows will not let a running executable be overwritten or deleted, but it
does allow it to be *renamed*. The swap therefore renames the running
executable aside, moves the freshly downloaded build into its place, starts
the new build, and exits. The renamed file is removed on the next start.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from sop_reporter.exceptions import UpdateError


LOGGER = logging.getLogger(__name__)

ASSET_NAME = "SOPReporter.exe"
PREVIOUS_BUILD_MARKER = ".previous-"
RELAUNCH_FLAG = "--wait-for-previous"
USER_AGENT = "SOPReporter-Updater"
_REQUEST_TIMEOUT = 30
_DOWNLOAD_TIMEOUT = 300
# Windows process creation flags; defined literally so this module stays
# importable on non-Windows machines for testing.
_DETACHED_PROCESS = 0x00000008
_CREATE_NO_WINDOW = 0x08000000


def parse_version(text: str) -> tuple[int, ...]:
    """Turn ``v0.3.1`` or ``0.3.1-beta`` into a comparable tuple."""
    cleaned = str(text).strip().lstrip("vV")
    parts: list[int] = []
    for chunk in cleaned.split("."):
        digits = ""
        for character in chunk:
            if not character.isdigit():
                break
            digits += character
        if not digits:
            break
        parts.append(int(digits))
    if not parts:
        raise UpdateError(f"Could not read a version number from {text!r}")
    return tuple(parts)


def is_newer(candidate: str, current: str) -> bool:
    return parse_version(candidate) > parse_version(current)


@dataclass(frozen=True)
class ReleaseInfo:
    version: str
    tag: str
    asset_name: str
    asset_url: str
    asset_size: int
    notes: str
    html_url: str


class Updater:
    def __init__(
        self,
        *,
        repository: str,
        current_version: str,
        executable: Path | None = None,
        asset_name: str = ASSET_NAME,
        include_prereleases: bool = False,
        opener: Callable[..., Any] = urllib.request.urlopen,
    ) -> None:
        self.repository = repository.strip().strip("/")
        self.current_version = current_version
        self.asset_name = asset_name
        self.include_prereleases = include_prereleases
        self._opener = opener
        self._executable = Path(executable) if executable else None

    # ------------------------------------------------------------------
    # Environment
    # ------------------------------------------------------------------
    @property
    def is_frozen(self) -> bool:
        return bool(getattr(sys, "frozen", False))

    @property
    def executable(self) -> Path:
        if self._executable is not None:
            return self._executable
        return Path(sys.executable).resolve()

    def ensure_updatable(self) -> None:
        if not self.is_frozen:
            raise UpdateError(
                "SOP Reporter is running from source, not from SOPReporter.exe. "
                "Self-update only applies to the packaged application."
            )
        folder = self.executable.parent
        probe = folder / ".update_write_test"
        try:
            probe.write_text("ok", encoding="utf-8")
            probe.unlink()
        except OSError as exc:
            raise UpdateError(
                f"The folder holding SOPReporter.exe is not writable: {folder}. "
                "Move the application somewhere writable, such as your Desktop, "
                "or reinstall it there."
            ) from exc

    # ------------------------------------------------------------------
    # Check
    # ------------------------------------------------------------------
    def check(self) -> ReleaseInfo | None:
        """Return the latest release when it is newer than the running build."""
        release = self._latest_release()
        if release is None:
            return None
        if not is_newer(release.version, self.current_version):
            LOGGER.info(
                "SOP Reporter %s is current (latest published: %s)",
                self.current_version,
                release.version,
            )
            return None
        LOGGER.info(
            "Update available: %s -> %s", self.current_version, release.version
        )
        return release

    def _latest_release(self) -> ReleaseInfo | None:
        if self.include_prereleases:
            payload = self._request_json(
                f"https://api.github.com/repos/{self.repository}/releases?per_page=10"
            )
            releases = [item for item in payload if not item.get("draft")]
            if not releases:
                return None
            entry = releases[0]
        else:
            entry = self._request_json(
                f"https://api.github.com/repos/{self.repository}/releases/latest"
            )
        return self._release_from_entry(entry)

    def _release_from_entry(self, entry: Any) -> ReleaseInfo | None:
        if not isinstance(entry, dict):
            raise UpdateError("GitHub returned an unexpected release payload")
        tag = str(entry.get("tag_name", "")).strip()
        if not tag:
            return None
        asset = None
        for candidate in entry.get("assets") or []:
            if str(candidate.get("name", "")) == self.asset_name:
                asset = candidate
                break
        if asset is None:
            raise UpdateError(
                f"Release {tag} does not attach {self.asset_name}. "
                "The build may still be running."
            )
        return ReleaseInfo(
            version=tag.lstrip("vV"),
            tag=tag,
            asset_name=self.asset_name,
            asset_url=str(asset.get("browser_download_url", "")),
            asset_size=int(asset.get("size", 0) or 0),
            notes=str(entry.get("body") or "").strip(),
            html_url=str(entry.get("html_url") or ""),
        )

    def _request_json(self, url: str) -> Any:
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "application/vnd.github+json",
            },
        )
        try:
            with self._opener(request, timeout=_REQUEST_TIMEOUT) as response:
                return json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                raise UpdateError(
                    "No published release was found for "
                    f"{self.repository}. Publish a release with "
                    f"{self.asset_name} attached, then check again."
                ) from exc
            if exc.code in {403, 429}:
                raise UpdateError(
                    "GitHub is rate limiting update checks. Try again later."
                ) from exc
            raise UpdateError(f"GitHub returned HTTP {exc.code} for {url}") from exc
        except (urllib.error.URLError, OSError) as exc:
            raise UpdateError(f"Could not reach GitHub: {exc}") from exc
        except json.JSONDecodeError as exc:
            raise UpdateError("GitHub returned a malformed release response") from exc

    # ------------------------------------------------------------------
    # Download
    # ------------------------------------------------------------------
    def download(
        self,
        release: ReleaseInfo,
        *,
        progress: Callable[[int, int], None] | None = None,
    ) -> Path:
        if not release.asset_url:
            raise UpdateError(f"Release {release.tag} has no downloadable asset URL")
        destination = self.executable.with_name(f"{self.executable.stem}.update.exe")
        request = urllib.request.Request(
            release.asset_url, headers={"User-Agent": USER_AGENT}
        )
        received = 0
        try:
            with self._opener(request, timeout=_DOWNLOAD_TIMEOUT) as response:
                total = int(response.headers.get("Content-Length") or release.asset_size)
                with destination.open("wb") as handle:
                    while True:
                        chunk = response.read(262_144)
                        if not chunk:
                            break
                        handle.write(chunk)
                        received += len(chunk)
                        if progress is not None:
                            progress(received, total)
        except (urllib.error.URLError, OSError) as exc:
            destination.unlink(missing_ok=True)
            raise UpdateError(f"Downloading the update failed: {exc}") from exc

        if release.asset_size and received != release.asset_size:
            destination.unlink(missing_ok=True)
            raise UpdateError(
                "The downloaded update was incomplete "
                f"({received} of {release.asset_size} bytes). Try again."
            )
        if received == 0:
            destination.unlink(missing_ok=True)
            raise UpdateError("The downloaded update was empty")
        LOGGER.info("Downloaded update %s to %s", release.version, destination)
        return destination

    # ------------------------------------------------------------------
    # Apply
    # ------------------------------------------------------------------
    def apply(self, staged: Path) -> Path:
        """Swap ``staged`` into place, returning the renamed previous build."""
        staged = Path(staged)
        if not staged.is_file():
            raise UpdateError(f"The downloaded update is missing: {staged}")
        target = self.executable
        previous = target.with_name(
            f"{target.stem}{PREVIOUS_BUILD_MARKER}{self.current_version}{target.suffix}"
        )
        previous.unlink(missing_ok=True)
        try:
            os.replace(target, previous)
        except OSError as exc:
            staged.unlink(missing_ok=True)
            raise UpdateError(
                f"Could not move the running application aside: {exc}"
            ) from exc
        try:
            os.replace(staged, target)
        except OSError as exc:
            # Put the working build back so the app still starts next time.
            try:
                os.replace(previous, target)
            except OSError:
                LOGGER.exception("Rollback failed; %s holds the previous build", previous)
            staged.unlink(missing_ok=True)
            raise UpdateError(f"Could not install the update: {exc}") from exc
        LOGGER.info("Installed update over %s (previous build kept at %s)", target, previous)
        return previous

    def relaunch(self) -> None:
        """Start the newly installed build and let it wait for this one to exit."""
        target = self.executable
        creation_flags = 0
        if os.name == "nt":
            creation_flags = _DETACHED_PROCESS | _CREATE_NO_WINDOW
        try:
            subprocess.Popen(  # noqa: S603 - launching our own executable
                [str(target), RELAUNCH_FLAG],
                cwd=str(target.parent),
                close_fds=True,
                creationflags=creation_flags,
            )
        except OSError as exc:
            raise UpdateError(
                f"The update installed, but SOP Reporter could not restart: {exc}. "
                "Start it again from the Start menu or its shortcut."
            ) from exc
        LOGGER.info("Relaunched %s after update", target)

    def cleanup_previous_builds(self) -> int:
        """Delete builds left behind by earlier updates. Best effort."""
        target = self.executable
        removed = 0
        pattern = f"{target.stem}{PREVIOUS_BUILD_MARKER}*{target.suffix}"
        try:
            candidates = list(target.parent.glob(pattern))
        except OSError:
            return 0
        for stale in candidates:
            try:
                stale.unlink()
                removed += 1
            except OSError:
                # Still locked by the outgoing process; the next start clears it.
                LOGGER.debug("Could not remove %s yet", stale)
        for leftover in target.parent.glob(f"{target.stem}.update.exe"):
            try:
                leftover.unlink()
            except OSError:
                LOGGER.debug("Could not remove stale download %s", leftover)
        if removed:
            LOGGER.info("Removed %d superseded build(s)", removed)
        return removed

    # ------------------------------------------------------------------
    # Convenience
    # ------------------------------------------------------------------
    def install(
        self,
        release: ReleaseInfo,
        *,
        progress: Callable[[int, int], None] | None = None,
    ) -> Path:
        """Download and swap in ``release`` without restarting."""
        self.ensure_updatable()
        staged = self.download(release, progress=progress)
        return self.apply(staged)
