"""Tests for utils.formatting — markdown renderers + helpers."""

import datetime as dt
import json

from outlook_mcp.utils.formatting import format_response, to_iso, truncate


def test_to_iso_relabels_fake_utc_as_local():
    # pywin32 hands us Outlook's local wall-clock time tagged as UTC.
    # The wall-clock value must be preserved and get the real local offset.
    fake_utc = dt.datetime(2026, 6, 10, 16, 33, 22, tzinfo=dt.timezone.utc)
    expected = dt.datetime(2026, 6, 10, 16, 33, 22).astimezone()
    assert to_iso(fake_utc) == expected.isoformat()
    assert "16:33:22" in to_iso(fake_utc)


def test_to_iso_naive_gets_local_offset():
    naive = dt.datetime(2026, 1, 5, 9, 0, 0)
    assert to_iso(naive) == naive.astimezone().isoformat()


def test_to_iso_handles_none_and_str():
    assert to_iso(None) is None
    assert to_iso("2026-01-01") == "2026-01-01"


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
