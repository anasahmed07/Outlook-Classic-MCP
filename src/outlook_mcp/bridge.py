"""Persistent STA thread for Outlook COM operations.

All COM calls run on a dedicated Single-Threaded Apartment (STA) thread
holding one ``Outlook.Application`` Dispatch handle. The MCP event loop
schedules work via :meth:`OutlookBridge.call` and awaits the result, so
the event loop never blocks on COM and COM never sees a non-STA thread.

Why one persistent thread instead of a fresh client per call:
* Reuses the same Dispatch — no per-call ``CoCreateInstance`` cost.
* Survives Outlook auto-launching (Dispatch will spawn ``OUTLOOK.EXE``
  if it isn't running, and the same connection serves every later call).
* Matches Outlook's strict STA threading model.
"""

from __future__ import annotations

import asyncio
import logging
import queue
import subprocess
import threading
import time
from typing import Any, Callable

logger = logging.getLogger("outlook_mcp.bridge")

# Cold-launching Outlook from a fully closed state can take 15–25s before
# the MAPI session answers `CurrentUser.Name`. The COM thread polls up to
# `_ATTACH_TIMEOUT_SEC`; the main thread waits slightly longer in start().
_ATTACH_TIMEOUT_SEC = 60
_READY_TIMEOUT_SEC = _ATTACH_TIMEOUT_SEC + 5
_CALL_TIMEOUT_SEC = 60


def _find_outlook_exe() -> str | None:
    """Look up OUTLOOK.EXE via the App Paths registry key."""
    try:
        import winreg
    except ImportError:
        return None
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(
                hive,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
            ) as key:
                path, _ = winreg.QueryValueEx(key, None)
                if path:
                    return path
        except OSError:
            continue
    return None


def _launch_outlook() -> bool:
    """Spawn OUTLOOK.EXE detached from this process. Returns True if started."""
    path = _find_outlook_exe()
    if not path:
        logger.warning("Could not locate OUTLOOK.EXE in registry App Paths")
        return False
    try:
        flags = 0
        # DETACHED_PROCESS so Outlook outlives this MCP process.
        flags |= getattr(subprocess, "DETACHED_PROCESS", 0)
        flags |= getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        subprocess.Popen(  # noqa: S603 - path comes from HKLM registry
            [path],
            creationflags=flags,
            close_fds=True,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        logger.info("Launched Outlook from %s", path)
        return True
    except OSError as exc:
        logger.warning("Failed to launch Outlook (%s): %s", path, exc)
        return False


class OutlookBridge:
    """Owns the COM thread and dispatches work to it."""

    def __init__(self) -> None:
        self._thread: threading.Thread | None = None
        self._queue: queue.Queue = queue.Queue()
        self._ready = threading.Event()
        self._shutdown = threading.Event()
        self._init_error: BaseException | None = None
        self._outlook: Any = None
        self._namespace: Any = None
        # Captured on the COM thread, safe to read from any thread.
        self._mailbox_name: str = "?"

    def start(self) -> None:
        """Spawn the COM thread and wait for it to attach to Outlook.

        Raises whatever the COM thread raised if Dispatch failed (with a
        friendly message — Outlook normally auto-launches; failure here
        usually means the user denied a UAC prompt or Outlook is mid-
        crash recovery).
        """
        self._thread = threading.Thread(
            target=self._run, daemon=True, name="outlook-com"
        )
        self._thread.start()
        if not self._ready.wait(timeout=_READY_TIMEOUT_SEC):
            raise RuntimeError(
                f"Outlook COM thread did not become ready within "
                f"{_READY_TIMEOUT_SEC}s. If Outlook didn't auto-launch, "
                "open it manually and retry."
            )
        if self._init_error is not None:
            raise self._init_error
        # Read the cached primitive — touching self._namespace from this
        # (main) thread would raise RPC_E_WRONG_THREAD because the COM
        # interface is marshalled to the bridge thread.
        logger.info("Bridge ready (mailbox: %s)", self._mailbox_name)

    def _run(self) -> None:
        import pythoncom
        from win32com.client import dynamic

        # Outlook is STA-only. DISABLE_OLE1DDE matches what Office
        # itself initializes COM with on its threads.
        pythoncom.CoInitializeEx(
            pythoncom.COINIT_APARTMENTTHREADED | pythoncom.COINIT_DISABLE_OLE1DDE
        )

        # Use dynamic (late-bound) Dispatch on purpose: pywin32's gencache
        # caches a typed proxy whose internal references carry thread
        # affinity. When install.bat pre-warms the typelib in one process
        # and the bridge later runs in a new process on a different STA
        # thread, calls into that cached wrapper raise RPC_E_WRONG_THREAD
        # (0x8001010E). The dynamic wrapper is pure IDispatch::Invoke —
        # slower per call, but marshals correctly across apartments.
        #
        # When Outlook is not running, Dispatch() spawns OUTLOOK.EXE but
        # the MAPI session isn't immediately usable — `CurrentUser.Name`
        # raises E_ABORT (-2147467260) until the profile finishes loading.
        # We poll, and if the first attempt fails we also kick off
        # OUTLOOK.EXE ourselves in case Dispatch's auto-launch was blocked.
        deadline = time.monotonic() + _ATTACH_TIMEOUT_SEC
        launched = False
        stage = "init"
        last_exc: BaseException | None = None
        while time.monotonic() < deadline and not self._shutdown.is_set():
            try:
                stage = "Dispatch(Outlook.Application)"
                self._outlook = dynamic.Dispatch("Outlook.Application")
                stage = "GetNamespace('MAPI')"
                self._namespace = self._outlook.GetNamespace("MAPI")
                # Capture the mailbox name as a plain string on this
                # thread so start() can log it from the main thread
                # without touching a COM proxy.
                stage = "CurrentUser.Name"
                self._mailbox_name = str(self._namespace.CurrentUser.Name)
                last_exc = None
                break
            except BaseException as exc:  # noqa: BLE001
                last_exc = exc
                self._outlook = None
                self._namespace = None
                if not launched:
                    logger.info(
                        "Outlook not ready (stage=%s): %s — launching it",
                        stage,
                        exc,
                    )
                    _launch_outlook()
                    launched = True
                time.sleep(1.5)

        if last_exc is not None:
            self._init_error = RuntimeError(
                f"Outlook bridge failed during '{stage}' after "
                f"{_ATTACH_TIMEOUT_SEC}s: {last_exc}"
            )
            self._init_error.__cause__ = last_exc
            self._ready.set()
            pythoncom.CoUninitialize()
            return

        self._ready.set()
        while not self._shutdown.is_set():
            try:
                func, args, kwargs, done, holder = self._queue.get(timeout=0.5)
            except queue.Empty:
                continue
            try:
                holder["value"] = func(self._outlook, self._namespace, *args, **kwargs)
            except BaseException as exc:  # noqa: BLE001 - propagated to caller
                holder["error"] = exc
            finally:
                done.set()

        pythoncom.CoUninitialize()

    async def call(self, func: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
        """Run ``func(outlook, namespace, *args, **kwargs)`` on the COM thread."""
        if self._thread is None or not self._thread.is_alive():
            raise RuntimeError("Bridge is not running. Did you forget to call start()?")
        done = threading.Event()
        holder: dict[str, Any] = {}
        self._queue.put((func, args, kwargs, done, holder))

        loop = asyncio.get_running_loop()
        signaled = await loop.run_in_executor(
            None, lambda: done.wait(timeout=_CALL_TIMEOUT_SEC)
        )
        if not signaled:
            raise TimeoutError(
                f"Outlook operation timed out after {_CALL_TIMEOUT_SEC}s. "
                "Outlook may be waiting on a dialog (security prompt, "
                "credential prompt, etc.) — check the Outlook window."
            )
        if "error" in holder:
            raise holder["error"]
        return holder.get("value")

    def stop(self) -> None:
        self._shutdown.set()
        if self._thread is not None:
            self._thread.join(timeout=5)
