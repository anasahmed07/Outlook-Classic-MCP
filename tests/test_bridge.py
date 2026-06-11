"""Tests for the bridge reconnect cycle, with COM faked out.

These run the real OutlookBridge thread machinery (CoInitializeEx and
all) but replace ``dynamic.Dispatch`` so no Outlook is needed.
"""

from __future__ import annotations

import asyncio
import sys
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "win32", reason="bridge requires pywin32"
)


def _fake_outlook():
    namespace = SimpleNamespace(CurrentUser=SimpleNamespace(Name="Test User"))
    return SimpleNamespace(GetNamespace=lambda kind: namespace)


def _rpc_unavailable() -> Exception:
    import pythoncom

    return pythoncom.com_error(
        -2147023174, "The RPC server is unavailable.", None, None
    )


@pytest.fixture
def bridge(monkeypatch):
    from win32com.client import dynamic

    from outlook_mcp.bridge import OutlookBridge

    monkeypatch.setattr(dynamic, "Dispatch", lambda name: _fake_outlook())
    b = OutlookBridge()
    b.start()
    yield b
    b.stop()


def test_call_runs_on_com_thread(bridge):
    result = asyncio.run(bridge.call(lambda outlook, ns: ns.CurrentUser.Name))
    assert result == "Test User"


def test_reconnects_and_retries_after_disconnect(bridge):
    attempts = []

    def flaky(outlook, namespace):
        attempts.append(namespace)
        if len(attempts) == 1:
            raise _rpc_unavailable()
        return "recovered"

    assert asyncio.run(bridge.call(flaky)) == "recovered"
    # Retried once, against a freshly attached namespace.
    assert len(attempts) == 2
    assert attempts[0] is not attempts[1]


def test_non_disconnect_errors_are_not_retried(bridge):
    attempts = []

    def broken(outlook, namespace):
        attempts.append(1)
        raise ValueError("bad input")

    with pytest.raises(ValueError):
        asyncio.run(bridge.call(broken))
    assert len(attempts) == 1


def test_error_surfaces_when_reattach_fails(bridge, monkeypatch):
    from win32com.client import dynamic

    import outlook_mcp.bridge as bridge_mod

    # After the disconnect, make re-attach fail fast: Dispatch raises and
    # the attach deadline is immediate.
    def dead_dispatch(name):
        raise _rpc_unavailable()

    monkeypatch.setattr(dynamic, "Dispatch", dead_dispatch)
    monkeypatch.setattr(bridge_mod, "_ATTACH_TIMEOUT_SEC", 0)
    monkeypatch.setattr(bridge_mod, "_launch_outlook", lambda: False)

    def dies(outlook, namespace):
        raise _rpc_unavailable()

    with pytest.raises(RuntimeError, match="could not be relaunched"):
        asyncio.run(bridge.call(dies))
