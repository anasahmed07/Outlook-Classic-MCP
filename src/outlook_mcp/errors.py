"""Domain errors and COM-error formatting."""

from __future__ import annotations

from typing import Any


class OutlookError(Exception):
    """Domain-level error from the Outlook wrapper.

    The message is safe to surface to an LLM/agent — it should suggest a
    corrective next step where possible.
    """


# HRESULTs that mean the COM channel to OUTLOOK.EXE is gone (the process
# exited or crashed), as opposed to an error *inside* a live Outlook.
_DISCONNECT_HRESULTS = frozenset(
    {
        0x800706BA,  # RPC_S_SERVER_UNAVAILABLE — "The RPC server is unavailable"
        0x800706BE,  # RPC_S_CALL_FAILED — server died mid-call
        0x800706BF,  # RPC_S_CALL_FAILED_DNE — call never reached the server
        0x80010108,  # RPC_E_DISCONNECTED — "The object invoked has disconnected"
        0x80010007,  # RPC_E_SERVER_DIED
        0x80010012,  # RPC_E_SERVER_DIED_DNE
        0x800401FD,  # CO_E_OBJNOTCONNECTED — "Object is not connected to server"
    }
)


def _com_hresults(exc: BaseException) -> set[int]:
    """Extract every HRESULT carried by a pywin32 ``com_error``.

    ``com_error.args`` is ``(hresult, message, excepinfo, argerror)``.
    Dispatch-level failures often surface as DISP_E_EXCEPTION with the
    real code tucked into ``excepinfo[5]`` (the scode), so collect both.
    """
    found: set[int] = set()
    args: tuple[Any, ...] = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        found.add(args[0] & 0xFFFFFFFF)
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6:
        scode = args[2][5]
        if isinstance(scode, int):
            found.add(scode & 0xFFFFFFFF)
    return found


def is_disconnect_error(exc: BaseException | None) -> bool:
    """True if ``exc`` (or anything in its cause chain) means Outlook died.

    Client code sometimes wraps the raw ``com_error`` in an
    :class:`OutlookError`, so walk ``__cause__``/``__context__`` too.
    """
    seen: set[int] = set()
    while exc is not None and id(exc) not in seen:
        seen.add(id(exc))
        if _com_hresults(exc) & _DISCONNECT_HRESULTS:
            return True
        exc = exc.__cause__ or exc.__context__
    return False


def format_com_error(exc: BaseException) -> str:
    """Render a pywin32 ``com_error`` (or any exception) as a friendly string.

    pywin32 raises ``pythoncom.com_error`` whose ``.args`` is a 4-tuple:
    ``(hresult, message, exc_info, arg_index)``. We surface the HRESULT
    in hex plus whatever short message we can find.
    """
    args: tuple[Any, ...] = getattr(exc, "args", ())
    if len(args) >= 2 and isinstance(args[0], int):
        hresult = args[0] & 0xFFFFFFFF
        msg = args[1] or ""
        detail = ""
        exc_info = args[2] if len(args) >= 3 else None
        if isinstance(exc_info, tuple) and len(exc_info) >= 3 and exc_info[2]:
            detail = f" — {exc_info[2]}"
        return f"COM Error 0x{hresult:08X}: {msg}{detail}"
    return f"{type(exc).__name__}: {exc}"
