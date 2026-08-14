from __future__ import annotations

import ctypes
import gc
import logging
import os
import time
from pathlib import Path
from typing import Any

from sop_reporter.config import PrinterConfig
from sop_reporter.exceptions import PrintError


LOGGER = logging.getLogger(__name__)

# Raw Excel constants are intentional. Frozen one-file applications cannot rely
# on win32com's generated constants/type-library cache being writable.
XL_PAPER_TABLOID = 3
XL_TYPE_PDF = 0
XL_PORTRAIT = 1
XL_LANDSCAPE = 2


def _excel_process_ids(psutil_module: Any) -> set[int]:
    process_ids: set[int] = set()
    for process in psutil_module.process_iter(["pid", "name"]):
        try:
            if str(process.info.get("name", "")).casefold() == "excel.exe":
                process_ids.add(int(process.info["pid"]))
        except (psutil_module.NoSuchProcess, psutil_module.AccessDenied):
            continue
    return process_ids


def _pid_from_window_handle(hwnd: int) -> int | None:
    if not hwnd:
        return None
    process_id = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(
        ctypes.c_void_p(hwnd), ctypes.byref(process_id)
    )
    return int(process_id.value) or None


class ExcelPrinter:
    def __init__(self, config: PrinterConfig) -> None:
        self.config = config

    def print_workbook(self, workbook_path: Path) -> None:
        if not self.config.enabled:
            LOGGER.info("Printing is disabled; generated report was not sent to a printer")
            return
        if os.name != "nt":
            raise PrintError("Excel printing is available only on Windows")
        workbook_path = Path(workbook_path).resolve()
        if not workbook_path.is_file():
            raise PrintError(f"Report workbook does not exist: {workbook_path}")

        try:
            import psutil
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise PrintError(
                "Windows Excel automation dependencies are not installed"
            ) from exc

        before_pids = _excel_process_ids(psutil)
        owned_pids: set[int] = set()
        excel = None
        workbook = None
        print_invoked = False
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False

            after_dispatch_pids = _excel_process_ids(psutil)
            diff_pids = after_dispatch_pids - before_pids
            try:
                hwnd_pid = _pid_from_window_handle(int(excel.Hwnd))
            except Exception:
                hwnd_pid = None
            if hwnd_pid is not None and hwnd_pid not in before_pids:
                owned_pids.add(hwnd_pid)
            elif hwnd_pid is not None and hwnd_pid in before_pids:
                LOGGER.warning(
                    "Excel COM reported a pre-existing process; SOP Reporter will never force-terminate it"
                )
            elif len(diff_pids) == 1:
                owned_pids.update(diff_pids)
            else:
                LOGGER.warning(
                    "Could not uniquely identify the isolated Excel process; orphan cleanup is limited"
                )

            workbook = excel.Workbooks.Open(
                str(workbook_path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                Notify=False,
            )
            orientation = (
                XL_LANDSCAPE
                if self.config.orientation == "landscape"
                else XL_PORTRAIT
            )
            for index in range(1, int(workbook.Worksheets.Count) + 1):
                worksheet = workbook.Worksheets.Item(index)
                page_setup = worksheet.PageSetup
                page_setup.PaperSize = XL_PAPER_TABLOID
                page_setup.Orientation = orientation
                page_setup.Zoom = False
                page_setup.FitToPagesWide = self.config.fit_to_pages_wide
                page_setup.FitToPagesTall = (
                    False
                    if self.config.fit_to_pages_tall == 0
                    else self.config.fit_to_pages_tall
                )

            if self.config.output == "pdf":
                # Same page setup as printing, so the PDF is a faithful preview
                # of the sheet of paper rather than a differently laid out file.
                pdf_path = workbook_path.with_suffix(".pdf")
                print_invoked = True
                workbook.ExportAsFixedFormat(
                    Type=XL_TYPE_PDF, Filename=str(pdf_path), OpenAfterPublish=False
                )
                LOGGER.info(
                    "Saved %s as Tabloid %s PDF", pdf_path.name, self.config.orientation
                )
            else:
                print_arguments: dict[str, Any] = {"Copies": self.config.copies}
                if self.config.name:
                    print_arguments["ActivePrinter"] = self.config.name
                print_invoked = True
                workbook.PrintOut(**print_arguments)
                LOGGER.info(
                    "Printed %s on Tabloid %s%s",
                    workbook_path.name,
                    self.config.orientation,
                    f" using {self.config.name}" if self.config.name else "",
                )
        except Exception as exc:
            raise PrintError(
                f"Excel could not print {workbook_path.name}: {exc}",
                print_invoked=print_invoked,
            ) from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    LOGGER.exception("Excel workbook did not close cleanly")
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    LOGGER.exception("The isolated Excel instance did not quit cleanly")
            workbook = None
            excel = None
            gc.collect()
            if com_initialized:
                pythoncom.CoUninitialize()
            self._cleanup_owned_excel_processes(psutil, owned_pids)

    @staticmethod
    def _cleanup_owned_excel_processes(psutil_module: Any, owned_pids: set[int]) -> None:
        if not owned_pids:
            return
        deadline = time.monotonic() + 5.0
        while time.monotonic() < deadline:
            remaining = [pid for pid in owned_pids if psutil_module.pid_exists(pid)]
            if not remaining:
                return
            time.sleep(0.2)
        for process_id in owned_pids:
            try:
                process = psutil_module.Process(process_id)
                if process.name().casefold() != "excel.exe":
                    continue
                LOGGER.warning("Terminating orphaned SOP Reporter Excel process %d", process_id)
                process.terminate()
                try:
                    process.wait(timeout=3)
                except psutil_module.TimeoutExpired:
                    process.kill()
            except (psutil_module.NoSuchProcess, psutil_module.AccessDenied):
                continue
