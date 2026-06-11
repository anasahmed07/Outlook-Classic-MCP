"""Tests for errors — disconnect detection and COM-error formatting."""

from outlook_mcp.errors import OutlookError, format_com_error, is_disconnect_error


class FakeComError(Exception):
    """Stands in for pythoncom.com_error: detection only reads ``.args``."""


def _rpc_unavailable() -> FakeComError:
    # pywin32 reports HRESULTs as signed ints; -2147023174 == 0x800706BA.
    return FakeComError(-2147023174, "The RPC server is unavailable.", None, None)


def test_disconnect_detected_from_signed_hresult():
    assert is_disconnect_error(_rpc_unavailable())


def test_disconnect_detected_from_unsigned_hresult():
    assert is_disconnect_error(FakeComError(0x80010108, "disconnected", None, None))


def test_disconnect_detected_in_excepinfo_scode():
    # DISP_E_EXCEPTION at the top, real failure in excepinfo[5].
    excepinfo = (0, "Microsoft Outlook", "boom", None, 0, -2147023174)
    exc = FakeComError(-2147352567, "Exception occurred.", excepinfo, None)
    assert is_disconnect_error(exc)


def test_disconnect_detected_through_cause_chain():
    wrapped = OutlookError("Item not found.")
    wrapped.__cause__ = _rpc_unavailable()
    assert is_disconnect_error(wrapped)


def test_not_found_is_not_disconnect():
    # MAPI_E_NOT_FOUND — Outlook alive, item missing.
    assert not is_disconnect_error(
        FakeComError(-2147221233, "The operation failed.", None, None)
    )


def test_plain_exceptions_are_not_disconnect():
    assert not is_disconnect_error(ValueError("nope"))
    assert not is_disconnect_error(None)


def test_format_com_error_renders_hex():
    msg = format_com_error(_rpc_unavailable())
    assert "0x800706BA" in msg
