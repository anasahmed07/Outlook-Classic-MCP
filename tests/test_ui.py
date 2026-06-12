"""MCP Apps UI layer: resource registration, tool metadata, result shape."""

from __future__ import annotations

import asyncio
import re

import pytest

from outlook_mcp import ui as ui_mod
from outlook_mcp.server import build_server

UI_TOOLS = {
    "outlook_list_mails": "ui://outlook/mail-list.html",
    "outlook_search_mails": "ui://outlook/mail-list.html",
    "outlook_get_mail": "ui://outlook/mail-view.html",
    "outlook_list_events": "ui://outlook/calendar.html",
    "outlook_list_contacts": "ui://outlook/contacts.html",
    "outlook_search_contacts": "ui://outlook/contacts.html",
    "outlook_list_tasks": "ui://outlook/tasks.html",
}


@pytest.fixture(scope="module")
def server():
    mcp, _bridge = build_server()
    return mcp


def test_ui_tools_declare_resource_uri(server):
    tools = asyncio.run(server.list_tools())
    by_name = {t.name: t for t in tools}
    for name, uri in UI_TOOLS.items():
        meta = by_name[name].meta
        assert meta and meta["ui"]["resourceUri"] == uri, name


def test_non_ui_tools_have_no_ui_meta(server):
    tools = asyncio.run(server.list_tools())
    for t in tools:
        if t.name not in UI_TOOLS:
            assert not (t.meta or {}).get("ui"), t.name


def test_ui_resources_registered_and_render(server):
    resources = asyncio.run(server.list_resources())
    uris = {str(r.uri) for r in resources}
    for uri in set(UI_TOOLS.values()):
        assert uri in uris
    for r in resources:
        assert r.mimeType == ui_mod.MIME_TYPE
        contents = asyncio.run(server.read_resource(r.uri))
        html = list(contents)[0].content
        assert html.lstrip().startswith("<!DOCTYPE html>")
        # the shared bridge must be injected
        assert "ui/initialize" in html
        # sandbox CSP blocks external loads: views must be self-contained
        assert not re.search(r'<script[^>]+src\s*=', html)
        assert not re.search(r'<link[^>]+href\s*=', html)
        assert ui_mod._UI_DIR.joinpath("_common.html").exists()
        assert "<!--%COMMON%-->" not in html
        # ui/initialize params must follow the MCP Apps spec: `appInfo` is
        # required — hosts that validate (Claude Desktop) silently drop the
        # handshake otherwise and the view renders empty.
        assert "appInfo:" in html
        assert "clientInfo" not in html


def test_ui_result_carries_both_representations():
    res = ui_mod.ui_result("**markdown**", {"items": [], "count": 0})
    assert res.content[0].text == "**markdown**"
    assert res.structuredContent == {"items": [], "count": 0}
    assert not res.isError


def test_ui_meta_rejects_unknown_view():
    with pytest.raises(ValueError):
        ui_mod.ui_meta("nope")
