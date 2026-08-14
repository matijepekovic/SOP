from __future__ import annotations

import ctypes
import os
import time


ERROR_ALREADY_EXISTS = 183


class AlreadyRunningError(RuntimeError):
    pass


class SingleInstance:
    def __init__(self, handle: int | None = None) -> None:
        self._handle = handle

    @classmethod
    def acquire(
        cls,
        name: str = "Local\\SOPReporter.Tray.Instance",
        *,
        wait_seconds: float = 0.0,
    ) -> "SingleInstance":
        """Claim the single-instance mutex.

        ``wait_seconds`` lets a freshly installed build wait for the older
        process that launched it to finish exiting. Without it the relaunch
        after a self-update would race the outgoing process and lose.
        """
        deadline = time.monotonic() + max(0.0, wait_seconds)
        while True:
            try:
                return cls._try_acquire(name)
            except AlreadyRunningError:
                if time.monotonic() >= deadline:
                    raise
                time.sleep(0.25)

    @classmethod
    def _try_acquire(cls, name: str) -> "SingleInstance":
        if os.name != "nt":
            return cls()
        kernel32 = ctypes.windll.kernel32
        kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p]
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
        kernel32.CloseHandle.restype = ctypes.c_bool
        kernel32.SetLastError(0)
        handle = kernel32.CreateMutexW(None, False, name)
        if not handle:
            raise OSError("Windows could not create the SOP Reporter instance mutex")
        if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
            kernel32.CloseHandle(ctypes.c_void_p(handle))
            raise AlreadyRunningError("SOP Reporter is already running")
        return cls(int(handle))

    def close(self) -> None:
        if self._handle is not None and os.name == "nt":
            ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(self._handle))
            self._handle = None

    def __enter__(self) -> "SingleInstance":
        return self

    def __exit__(self, *_args) -> None:
        self.close()
