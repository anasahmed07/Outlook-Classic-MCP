"""Tests for utils.formatting — markdown renderers + helpers."""

import json

from outlook_mcp.utils.formatting import format_response, truncate


def test_truncate_short_passthrough():
    assert truncate("hi", 10) == "hi"


def test_truncate_long_appends_ellipsis():
    out = truncate("a" * 100, 10)
    assert out.endswith("…")
    assert len(out) <= 11


def test_truncate_handles_none():
    assert truncate(None) == ""


def test_format_json_passthrough():
    out = format_response({"a": 1}, "json")
    assert json.loads(out) == {"a": 1}


def test_format_markdown_mail_collection():
    payload = {
        "count": 1,
        "folder": "Inbox",
        "items": [
            {
                "subject": "Hello",
                "from": "Alice",
                "from_address": "alice@example.com",
                "received": "2026-04-25T10:00:00",
                "unread": True,
                "has_attachments": False,
                "preview": "body preview",
                "entry_id": "abc",
            }
        ],
    }
    out = format_response(payload, "markdown")
    assert "**Hello**" in out
    assert "Alice" in out
    assert "alice@example.com" in out
    assert "abc" in out


def test_format_markdown_categories_collection():
    payload = {
        "count": 2,
        "items": [
            {"name": "Work", "color": 1},
            {"name": "Home", "color": 7},
        ],
    }
    out = format_response(payload, "markdown")
    assert "Work" in out
    assert "Home" in out
    assert "color 1" in out


def test_format_markdown_rules_collection():
    payload = {
        "count": 1,
        "items": [{"index": 1, "name": "Move newsletters", "enabled": True}],
    }
    out = format_response(payload, "markdown")
    assert "Move newsletters" in out
    assert "ON" in out


def test_format_markdown_mail_detail():
    payload = {
        "subject": "Project update",
        "from": "Bob",
        "from_address": "bob@example.com",
        "to": "anas@example.com",
        "received": "2026-04-25T11:00:00",
        "body": "Here's the update.",
        "html_body": "<p>Here's the update.</p>",
        "attachments": [],
    }
    out = format_response(payload, "markdown")
    assert out.startswith("# Project update")
    assert "Bob" in out
    assert "Here's the update." in out
