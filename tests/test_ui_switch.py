"""OUTLOOK_MCP_UI switch: no automatic cards, explicit show tools instead."""

from __future__ import annotations

import asyncio

from outlook_mcp.server import build_server


def _build(monkeypatch, ui=None):
    monkeypatch.delenv("OUTLOOK_MCP_UI", raising=False)
    if ui is not None:
        monkeypatch.setenv("OUTLOOK_MCP_UI", ui)
    mcp, _ = build_server()
    return mcp, {t.name: t for t in asyncio.run(mcp.list_tools())}


def test_default_unchanged(monkeypatch):
    _, tools = _build(monkeypatch)
    assert "outlook_show_mails" not in tools
    assert tools["outlook_list_mails"].meta["ui"]["resourceUri"].startswith("ui://")


def test_off_removes_card_meta_and_adds_show_tools(monkeypatch):
    _, tools = _build(monkeypatch, "off")
    for name in ("outlook_list_mails", "outlook_search_mails", "outlook_get_mail",
                 "outlook_list_events", "outlook_list_contacts", "outlook_list_tasks"):
        assert not (tools[name].meta or {}).get("ui"), name
    assert tools["outlook_show_mails"].meta["ui"]["resourceUri"].endswith("mail-list.html")
    assert tools["outlook_show_mail"].meta["ui"]["resourceUri"].endswith("mail-view.html")


def test_off_still_registers_resources(monkeypatch):
    mcp, _ = _build(monkeypatch, "off")
    uris = {str(r.uri) for r in asyncio.run(mcp.list_resources())}
    assert "ui://outlook/mail-list.html" in uris


def test_show_tools_describe_when_to_use_them(monkeypatch):
    _, tools = _build(monkeypatch, "off")
    for name in ("outlook_show_mails", "outlook_show_mail"):
        text = tools[name].description.lower()
        assert "card" in text and "widget" in text, name


def test_instructions_point_to_show_tools_only_when_off(monkeypatch):
    mcp_off, _ = _build(monkeypatch, "off")
    mcp_on, _ = _build(monkeypatch)
    assert "outlook_show_mails" in mcp_off._mcp_server.instructions
    assert "outlook_show_mails" not in mcp_on._mcp_server.instructions
