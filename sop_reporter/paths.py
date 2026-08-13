from __future__ import annotations

import os
import shutil
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from sop_reporter.exceptions import ConfigurationError


APP_NAME = "SOPReporter"


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"))


def _is_writable_directory(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix=".sop-write-test-", dir=path, delete=True):
            pass
        return True
    except OSError:
        return False


def _appdata_root() -> Path:
    configured = os.environ.get("APPDATA")
    if configured:
        return Path(configured) / APP_NAME
    return Path.home() / "AppData" / "Roaming" / APP_NAME


@dataclass(frozen=True)
class AppPaths:
    resource_root: Path
    runtime_root: Path

    @classmethod
    def discover(cls) -> "AppPaths":
        override = os.environ.get("SOP_REPORTER_DATA_DIR")
        if _is_frozen():
            resource_root = Path(getattr(sys, "_MEIPASS")).resolve()
            preferred_runtime = Path(sys.executable).resolve().parent
        else:
            resource_root = Path(__file__).resolve().parent.parent
            preferred_runtime = resource_root

        if override:
            runtime_root = Path(override).expanduser().resolve()
            if not _is_writable_directory(runtime_root):
                raise ConfigurationError(
                    f"SOP_REPORTER_DATA_DIR is not writable: {runtime_root}"
                )
        elif _is_writable_directory(preferred_runtime):
            runtime_root = preferred_runtime
        else:
            runtime_root = _appdata_root()
            if not _is_writable_directory(runtime_root):
                raise ConfigurationError(
                    f"Neither the executable folder nor AppData is writable: {runtime_root}"
                )

        return cls(resource_root=resource_root, runtime_root=runtime_root)

    @property
    def config_dir(self) -> Path:
        return self.runtime_root / "config"

    @property
    def app_config_path(self) -> Path:
        return self.config_dir / "app_config.yaml"

    @property
    def extraction_config_path(self) -> Path:
        return self.config_dir / "extraction_rules.yaml"

    @property
    def state_path(self) -> Path:
        return self.runtime_root / "state.json"

    @property
    def logs_dir(self) -> Path:
        return self.runtime_root / "logs"

    @property
    def default_icon_path(self) -> Path:
        return self.resource_root / "assets" / "tray_icon.ico"

    def resource(self, relative_path: str | Path) -> Path:
        return self.resource_root / relative_path

    def resolve_runtime_directory(self, configured: str) -> Path:
        candidate = Path(configured).expanduser()
        if candidate.is_absolute():
            return candidate
        return self.runtime_root / candidate

    def ensure_layout(self) -> None:
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.logs_dir.mkdir(parents=True, exist_ok=True)

        copies = (
            (
                self.resource("config/app_config.default.yaml"),
                self.app_config_path,
            ),
            (
                self.resource("config/extraction_rules.default.yaml"),
                self.extraction_config_path,
            ),
        )
        for source, destination in copies:
            if destination.exists():
                continue
            if not source.is_file():
                raise ConfigurationError(f"Bundled default config is missing: {source}")
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)

