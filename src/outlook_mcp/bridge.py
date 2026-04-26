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
import threading
from typing import Any, Callable

logger = logging.getLogger("outlook_mcp.bridge")

_READY_TIMEOUT_SEC = 15
_CALL_TIMEOUT_SEC = 60


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
        stage = "init"
        try:
            # Use dynamic (late-bound) Dispatch on purpose: pywin32's
            # gencache caches a typed proxy whose internal references
            # carry thread affinity. When install.bat pre-warms the
            # typelib in one process and the bridge later runs in a new
            # process on a different STA thread, calls into that cached
            # wrapper raise RPC_E_WRONG_THREAD (0x8001010E). The dynamic
            # wrapper is pure IDispatch::Invoke — slower per call, but
            # marshals correctly across apartments.
            stage = "Dispatch(Outlook.Application)"
            self._outlook = dynamic.Dispatch("Outlook.Application")
            stage = "GetNamespace('MAPI')"
            self._namespace = self._outlook.GetNamespace("MAPI")
            # Capture the mailbox name as a plain string on this thread
            # so start() can log it from the main thread without
            # touching a COM proxy.
            stage = "CurrentUser.Name"
            self._mailbox_name = str(self._namespace.CurrentUser.Name)
        except BaseException as exc:  # noqa: BLE001 - re-raised by start()
            self._init_error = RuntimeError(
                f"Outlook bridge failed during '{stage}': {exc}"
            )
            self._init_error.__cause__ = exc
            self._ready.set()
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
