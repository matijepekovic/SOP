from __future__ import annotations

import logging
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Callable

import yaml

from sop_reporter.exceptions import UpdateError
from sop_reporter.pipeline import RunResult, RunStatus


LOGGER = logging.getLogger(__name__)


class ControlWindow:
    """Small operator-facing dashboard for the persistent tray application."""

    def __init__(
        self,
        *,
        job_runner,
        run_now: Callable[[], None],
        exit_app: Callable[[], None],
        app_config_path: Path,
        extraction_config_path: Path,
        reports_dir: Path,
        logs_dir: Path,
        updater=None,
        current_version: str = "",
        check_updates_on_startup: bool = False,
    ) -> None:
        self.job_runner = job_runner
        self._run_now_callback = run_now
        self._exit_callback = exit_app
        self.app_config_path = Path(app_config_path)
        self.extraction_config_path = Path(extraction_config_path)
        self.reports_dir = Path(reports_dir)
        self.logs_dir = Path(logs_dir)
        self.updater = updater
        self.current_version = current_version
        self._pending_release = None
        self._update_busy = False
        self._events: queue.Queue[tuple[str, object | None]] = queue.Queue()

        self.root = tk.Tk()
        self.root.title("SOP Reporter")
        self.root.geometry("760x650")
        self.root.minsize(700, 580)
        self.root.protocol("WM_DELETE_WINDOW", self.hide)

        self.sender = tk.StringVar()
        self.subject = tk.StringVar()
        self.since_days = tk.StringVar()
        self.unread_only = tk.BooleanVar()
        self.schedule_enabled = tk.BooleanVar()
        self.schedule_time = tk.StringVar()
        self.printer_enabled = tk.BooleanVar()
        self.printer_output = tk.StringVar()
        self.printer_name = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.update_status = tk.StringVar(
            value=f"Version {current_version}" if current_version else "Version unknown"
        )

        self._build()
        self._load()
        self.root.after(150, self._drain_events)
        if check_updates_on_startup and self.updater is not None:
            self.root.after(2000, lambda: self._check_updates(silent=True))

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="SOP Reporter", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="Olympia market • one report and print job per Sub Status",
        ).pack(anchor="w", pady=(0, 14))

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(0, 12))
        self.run_button = ttk.Button(actions, text="Run Now", command=self._run_now)
        self.run_button.pack(side="left")
        ttk.Button(actions, text="Open Reports", command=lambda: self._open(self.reports_dir)).pack(side="left", padx=6)
        ttk.Button(actions, text="Open Logs", command=lambda: self._open(self.logs_dir)).pack(side="left")
        ttk.Button(actions, text="Open Rule File", command=lambda: self._open(self.extraction_config_path)).pack(side="left", padx=6)

        status_frame = ttk.LabelFrame(outer, text="Current status", padding=12)
        status_frame.pack(fill="x", pady=(0, 12))
        ttk.Label(status_frame, textvariable=self.status, wraplength=680).pack(anchor="w")

        email = ttk.LabelFrame(outer, text="Gmail search", padding=12)
        email.pack(fill="x", pady=(0, 12))
        self._field(email, 0, "Sender contains", self.sender)
        self._field(email, 1, "Subject contains", self.subject)
        self._field(email, 2, "Search last (days)", self.since_days, width=10)
        ttk.Checkbutton(email, text="Unread messages only", variable=self.unread_only).grid(row=3, column=1, sticky="w", pady=4)
        ttk.Label(
            email,
            text="For testing, leave sender and subject blank. Only .xlsx/.xlsm attachments are downloaded.",
            foreground="#555555",
            wraplength=620,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        schedule = ttk.LabelFrame(outer, text="Automatic schedule", padding=12)
        schedule.pack(fill="x", pady=(0, 12))
        ttk.Checkbutton(schedule, text="Enabled Monday–Friday", variable=self.schedule_enabled).grid(row=0, column=0, sticky="w")
        ttk.Label(schedule, text="Run time (24-hour)").grid(row=0, column=1, padx=(30, 8))
        ttk.Entry(schedule, textvariable=self.schedule_time, width=9).grid(row=0, column=2)

        printer = ttk.LabelFrame(outer, text="Printing", padding=12)
        printer.pack(fill="x", pady=(0, 12))
        ttk.Checkbutton(printer, text="Send output automatically", variable=self.printer_enabled).grid(row=0, column=0, sticky="w", pady=4)
        output_row = ttk.Frame(printer)
        output_row.grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Radiobutton(
            output_row, text="Save as PDF", variable=self.printer_output, value="pdf"
        ).pack(side="left")
        ttk.Radiobutton(
            output_row, text="Send to printer", variable=self.printer_output, value="printer"
        ).pack(side="left", padx=(12, 0))
        ttk.Label(printer, text="Excel printer name").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(printer, textvariable=self.printer_name).grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=4)
        printer.columnconfigure(1, weight=1)
        ttk.Label(printer, text="PDFs are saved next to the reports, laid out exactly as they would print. Blank printer name uses the Windows default. Output is forced to 17×11 landscape either way.", foreground="#555555", wraplength=620).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        updates = ttk.LabelFrame(outer, text="Software updates", padding=12)
        updates.pack(fill="x", pady=(0, 12))
        update_buttons = ttk.Frame(updates)
        update_buttons.pack(fill="x")
        self.check_update_button = ttk.Button(
            update_buttons,
            text="Check for Updates",
            command=lambda: self._check_updates(silent=False),
        )
        self.check_update_button.pack(side="left")
        self.install_update_button = ttk.Button(
            update_buttons,
            text="Install and Restart",
            command=self._install_update,
            state="disabled",
        )
        self.install_update_button.pack(side="left", padx=6)
        ttk.Label(
            updates,
            textvariable=self.update_status,
            wraplength=680,
            foreground="#555555",
        ).pack(anchor="w", pady=(8, 0))
        if self.updater is None:
            self.check_update_button.configure(state="disabled")
            self.update_status.set("Updates are disabled in the configuration file.")

        footer = ttk.Frame(outer)
        footer.pack(fill="x", pady=(4, 0))
        ttk.Button(footer, text="Save Settings", command=self._save).pack(side="left")
        ttk.Button(footer, text="Exit SOP Reporter", command=self._exit_callback).pack(side="right")
        ttk.Label(
            footer,
            text="Saved settings take effect after exiting and reopening the app.",
            foreground="#8A4B08",
        ).pack(side="left", padx=12)

    @staticmethod
    def _field(parent, row: int, label: str, variable: tk.Variable, width: int | None = None) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        entry = ttk.Entry(parent, textvariable=variable, width=width)
        entry.grid(row=row, column=1, sticky="ew", padx=(12, 0), pady=4)
        parent.columnconfigure(1, weight=1)

    def _load(self) -> None:
        with self.app_config_path.open("r", encoding="utf-8") as handle:
            data = yaml.safe_load(handle) or {}
        email = data.get("email", {})
        search = email.get("search", {})
        schedule = data.get("schedule", {})
        printer = data.get("printer", {})
        self.sender.set(search.get("from", ""))
        self.subject.set(search.get("subject_contains", ""))
        self.since_days.set(str(search.get("since_days", 14)))
        self.unread_only.set(bool(search.get("unread_only", False)))
        self.schedule_enabled.set(bool(schedule.get("enabled", True)))
        self.schedule_time.set(str(schedule.get("time", "07:00")))
        self.printer_enabled.set(bool(printer.get("enabled", True)))
        self.printer_output.set(str(printer.get("output", "printer")))
        self.printer_name.set(str(printer.get("name", "")))

    def _save(self) -> None:
        try:
            days = int(self.since_days.get())
            if days < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid setting", "Search days must be a whole number of 1 or more.", parent=self.root)
            return
        time_value = self.schedule_time.get().strip()
        try:
            hour, minute = (int(value) for value in time_value.split(":"))
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid setting", "Schedule time must use 24-hour HH:MM format.", parent=self.root)
            return

        with self.app_config_path.open("r", encoding="utf-8") as handle:
            data = yaml.safe_load(handle) or {}
        search = data.setdefault("email", {}).setdefault("search", {})
        search["from"] = self.sender.get().strip()
        search["subject_contains"] = self.subject.get().strip()
        search["since_days"] = days
        search["unread_only"] = self.unread_only.get()
        schedule = data.setdefault("schedule", {})
        schedule["enabled"] = self.schedule_enabled.get()
        schedule["time"] = time_value
        printer = data.setdefault("printer", {})
        printer["enabled"] = self.printer_enabled.get()
        printer["output"] = self.printer_output.get() or "printer"
        printer["name"] = self.printer_name.get().strip()

        temporary = self.app_config_path.with_suffix(".yaml.tmp")
        with temporary.open("w", encoding="utf-8", newline="\n") as handle:
            yaml.safe_dump(data, handle, sort_keys=False, allow_unicode=True)
        os.replace(temporary, self.app_config_path)
        self.status.set("Settings saved. Exit and reopen SOP Reporter, then click Run Now.")
        messagebox.showinfo(
            "Settings saved",
            "Exit and reopen SOP Reporter so the new Gmail filters, schedule, and printer settings take effect.",
            parent=self.root,
        )

    def _run_now(self) -> None:
        if self.job_runner.is_running:
            self.status.set("A run is already in progress.")
            return
        self.set_running()
        self._run_now_callback()

    def set_running(self) -> None:
        self._events.put(("running", None))

    def set_result(self, result: RunResult) -> None:
        self._events.put(("result", result))

    # ------------------------------------------------------------------
    # Software updates
    # ------------------------------------------------------------------
    def check_for_updates(self) -> None:
        """Public entry point, safe to call from the tray thread."""
        self.root.after(0, lambda: self._check_updates(silent=False))

    def _check_updates(self, *, silent: bool) -> None:
        """Ask GitHub for a newer release. Network work runs off the UI thread."""
        if self.updater is None or self._update_busy:
            return
        self._update_busy = True
        self.check_update_button.configure(state="disabled")
        self.install_update_button.configure(state="disabled")
        self.update_status.set("Checking for updates…")
        threading.Thread(
            target=self._check_updates_worker,
            args=(silent,),
            name="SOP-Update-Check",
            daemon=True,
        ).start()

    def _check_updates_worker(self, silent: bool) -> None:
        try:
            release = self.updater.check()
        except UpdateError as exc:
            self._events.put(("update_error", (str(exc), silent)))
        except Exception as exc:  # pragma: no cover - defensive
            LOGGER.exception("Unexpected failure while checking for updates")
            self._events.put(("update_error", (str(exc), silent)))
        else:
            self._events.put(("update_checked", (release, silent)))

    def _install_update(self) -> None:
        release = self._pending_release
        if self.updater is None or release is None or self._update_busy:
            return
        confirmed = messagebox.askyesno(
            "Install update",
            f"Install version {release.version} and restart SOP Reporter now?\n\n"
            "Any run in progress will be interrupted.",
            parent=self.root,
        )
        if not confirmed:
            return
        if self.job_runner.is_running:
            messagebox.showinfo(
                "SOP Reporter is busy",
                "A fetch/report/print run is in progress. Wait for it to finish, "
                "then install the update.",
                parent=self.root,
            )
            return
        self._update_busy = True
        self.check_update_button.configure(state="disabled")
        self.install_update_button.configure(state="disabled")
        self.update_status.set(f"Downloading version {release.version}…")
        threading.Thread(
            target=self._install_update_worker,
            args=(release,),
            name="SOP-Update-Install",
            daemon=True,
        ).start()

    def _install_update_worker(self, release) -> None:
        def progress(received: int, total: int) -> None:
            if total:
                percent = int(received * 100 / total)
                self._events.put(("update_progress", (release.version, percent)))

        try:
            self.updater.install(release, progress=progress)
            self.updater.relaunch()
        except UpdateError as exc:
            self._events.put(("update_error", (str(exc), False)))
        except Exception as exc:  # pragma: no cover - defensive
            LOGGER.exception("Unexpected failure while installing the update")
            self._events.put(("update_error", (str(exc), False)))
        else:
            self._events.put(("update_installed", release))

    def _on_update_checked(self, release, silent: bool) -> None:
        self._update_busy = False
        self.check_update_button.configure(state="normal")
        if release is None:
            self._pending_release = None
            self.install_update_button.configure(state="disabled")
            self.update_status.set(
                f"Version {self.current_version} is the latest available."
            )
            if not silent:
                messagebox.showinfo(
                    "No update available",
                    f"SOP Reporter {self.current_version} is up to date.",
                    parent=self.root,
                )
            return
        self._pending_release = release
        self.install_update_button.configure(state="normal")
        self.update_status.set(
            f"Version {release.version} is available "
            f"(you have {self.current_version}). "
            "Choose Install and Restart to apply it."
        )

    def _on_update_error(self, message: str, silent: bool) -> None:
        self._update_busy = False
        self.check_update_button.configure(state="normal")
        self.install_update_button.configure(
            state="normal" if self._pending_release else "disabled"
        )
        self.update_status.set(f"Update check failed: {message}")
        if not silent:
            messagebox.showerror("Update failed", message, parent=self.root)

    def _drain_events(self) -> None:
        while True:
            try:
                event, payload = self._events.get_nowait()
            except queue.Empty:
                break
            if event == "running":
                self.status.set("Checking Gmail and processing matching Excel attachments…")
                self.run_button.configure(state="disabled")
            elif event == "result" and isinstance(payload, RunResult):
                self.status.set(payload.message)
                self.run_button.configure(state="normal")
                if payload.status == RunStatus.FAILED:
                    messagebox.showerror("SOP Reporter failed", payload.message, parent=self.root)
            elif event == "update_checked" and isinstance(payload, tuple):
                self._on_update_checked(*payload)
            elif event == "update_error" and isinstance(payload, tuple):
                self._on_update_error(*payload)
            elif event == "update_progress" and isinstance(payload, tuple):
                version, percent = payload
                self.update_status.set(f"Downloading version {version}… {percent}%")
            elif event == "update_installed":
                version = getattr(payload, "version", "")
                self.update_status.set(
                    f"Version {version} installed. Restarting SOP Reporter…"
                )
                self._exit_callback()
        if self.root.winfo_exists():
            self.root.after(150, self._drain_events)

    def run(self) -> None:
        self.root.mainloop()

    def show(self) -> None:
        self.root.after(0, self._show_now)

    def _show_now(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def hide(self) -> None:
        self.root.withdraw()

    def close(self) -> None:
        try:
            self.root.after(0, self.root.destroy)
        except tk.TclError:
            pass

    @staticmethod
    def _open(path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.suffix:
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            path.mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))  # type: ignore[attr-defined]
