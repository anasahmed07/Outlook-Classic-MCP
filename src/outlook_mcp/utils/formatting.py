"""Markdown rendering for tool responses.

Tools accept ``response_format='markdown' | 'json'``. ``json`` returns
``json.dumps`` of the raw payload; ``markdown`` runs the renderers
below, which detect the payload shape (mail list, event list, contact
list, task list, folder list, single-item detail) and produce a chat-
friendly view.
"""

from __future__ import annotations

import datetime as dt
import json
from typing import Any

from outlook_mcp.schemas import ResponseFormat


def truncate(text: str | None, limit: int = 500) -> str:
    if not text:
        return ""
    text = text.replace("\r\n", "\n").strip()
    if len(text) <= limit:
        return text
    return text[:limit].rstrip() + "…"


def to_iso(value: Any) -> str | None:
    """Convert a pywintypes/datetime/None to ISO-8601 string."""
    if value is None:
        return None
    try:
        return value.isoformat()
    except AttributeError:
        return str(value)


def from_iso(value: str | None) -> dt.datetime | None:
    """Parse an ISO-8601 string into a naive datetime (Outlook uses local)."""
    if value is None:
        return None
    from outlook_mcp.errors import OutlookError

    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as exc:
        raise OutlookError(
            f"Invalid datetime '{value}'. Use ISO-8601 format like "
            "'2026-04-25T14:30:00'."
        ) from exc
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone().replace(tzinfo=None)
    return parsed


def format_response(data: Any, fmt: ResponseFormat | str) -> str:
    if isinstance(fmt, str):
        fmt = ResponseFormat(fmt)
    if fmt == ResponseFormat.JSON:
        return json.dumps(data, indent=2, default=str, ensure_ascii=False)
    return _to_markdown(data)


def _to_markdown(data: Any) -> str:
    if isinstance(data, dict):
        items = data.get("items")
        if isinstance(items, list) and items:
            return _render_collection(data)
        return _render_detail(data)
    return "```json\n" + json.dumps(data, indent=2, default=str) + "\n```"


def _render_collection(data: dict[str, Any]) -> str:
    items = data["items"]
    sample = items[0]
    lines: list[str] = []

    header_bits = [f"**{data.get('count', len(items))} result(s)**"]
    if "folder" in data:
        header_bits.append(f"folder: `{data['folder']}`")
    if "query" in data:
        header_bits.append(f"query: `{data['query']}`")
    lines.append(" — ".join(header_bits))
    lines.append("")

    if "from" in sample and "received" in sample:
        for m in items:
            unread_marker = "(unread) " if m.get("unread") else ""
            attach_marker = " [attach]" if m.get("has_attachments") else ""
            lines.append(
                f"- {unread_marker}**{m.get('subject', '(no subject)')}**{attach_marker}"
            )
            lines.append(f"  - from: {m.get('from')} <{m.get('from_address')}>")
            lines.append(f"  - received: {m.get('received')}")
            if m.get("preview"):
                lines.append(f"  - preview: {m['preview']}")
            lines.append(f"  - id: `{m.get('entry_id')}`")
    elif "start" in sample and "end" in sample and "subject" in sample:
        for ev in items:
            lines.append(f"- **{ev.get('subject', '(no subject)')}**")
            lines.append(f"  - {ev.get('start')} -> {ev.get('end')}")
            if ev.get("location"):
                lines.append(f"  - location: {ev['location']}")
            if ev.get("organizer"):
                lines.append(f"  - organizer: {ev['organizer']}")
            lines.append(f"  - id: `{ev.get('entry_id')}`")
    elif "full_name" in sample:
        for c in items:
            line = f"- **{c.get('full_name')}**"
            if c.get("email"):
                line += f" — {c['email']}"
            lines.append(line)
            if c.get("company") or c.get("job_title"):
                lines.append(
                    f"  - {c.get('job_title') or ''} @ {c.get('company') or ''}".strip()
                )
            lines.append(f"  - id: `{c.get('entry_id')}`")
    elif "due_date" in sample:
        for t in items:
            mark = "[x] " if t.get("complete") else "[ ] "
            lines.append(f"- {mark}**{t.get('subject')}**")
            if t.get("due_date"):
                lines.append(f"  - due: {t['due_date']}")
            lines.append(f"  - id: `{t.get('entry_id')}`")
    elif "path" in sample and "item_count" in sample:
        for f in items:
            lines.append(
                f"- `{f['path']}` — {f['item_count']} items, "
                f"{f.get('unread_count', 0)} unread"
            )
    elif "name" in sample and "color" in sample:
        # Categories
        for cat in items:
            lines.append(f"- **{cat['name']}** (color {cat['color']})")
    elif "enabled" in sample and "name" in sample:
        # Rules
        for r in items:
            mark = "ON " if r.get("enabled") else "OFF"
            lines.append(f"- [{mark}] {r['name']}")
    else:
        lines.append("```json")
        lines.append(json.dumps(items, indent=2, default=str))
        lines.append("```")

    if data.get("has_more"):
        lines.append("")
        lines.append(
            f"_More results available — pass `offset={data.get('next_offset')}` to continue._"
        )
    return "\n".join(lines)


def _render_detail(data: dict[str, Any]) -> str:
    lines: list[str] = []
    if "subject" in data and ("body" in data or "html_body" in data):
        lines.append(f"# {data.get('subject', '(no subject)')}")
        if data.get("from"):
            lines.append(f"**From:** {data['from']} <{data.get('from_address', '')}>")
        if data.get("to"):
            lines.append(f"**To:** {data['to']}")
        if data.get("cc"):
            lines.append(f"**CC:** {data['cc']}")
        if data.get("received"):
            lines.append(f"**Received:** {data['received']}")
        if data.get("start"):
            lines.append(f"**When:** {data['start']} -> {data.get('end')}")
        if data.get("location"):
            lines.append(f"**Location:** {data['location']}")
        if data.get("attachments"):
            lines.append("")
            lines.append("**Attachments:**")
            for att in data["attachments"]:
                lines.append(
                    f"- [{att['index']}] {att['filename']} "
                    f"({att.get('size_bytes', '?')} bytes)"
                )
        lines.append("")
        lines.append("---")
        lines.append("")
        body = data.get("body", "")
        lines.append(body if body else "_(no body)_")
        return "\n".join(lines)
    return "```json\n" + json.dumps(data, indent=2, default=str) + "\n```"
