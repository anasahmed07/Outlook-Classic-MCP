"""Domain errors and COM-error formatting."""

from __future__ import annotations

from typing import Any


class OutlookError(Exception):
    """Domain-level error from the Outlook wrapper.

    The message is safe to surface to an LLM/agent — it should suggest a
    corrective next step where possible.
    """


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
