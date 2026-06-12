"""MCP Apps UI layer (SEP-1865).

Serves self-contained HTML views as ``ui://`` resources and provides the
helpers tools use to attach a view to themselves:

- ``ui_meta(view)``      -> ``_meta`` dict for the ``@mcp.tool`` decorator
- ``ui_result(text, d)`` -> ``CallToolResult`` carrying markdown for the
                            model *and* ``structuredContent`` for the app
- ``register_ui(mcp)``   -> registers every view as a resource

Hosts that support MCP Apps (Claude, Claude Desktop, VS Code Copilot, …)
render the view in a sandboxed iframe in place of the text result; hosts
that don't simply show the markdown — nothing changes for them.

Each view is a single HTML file with inline CSS/JS only (the sandbox CSP
blocks external loads by default). The shared bridge + stylesheet live in
``_common.html`` and are injected at read time.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from mcp.types import CallToolResult, TextContent

_UI_DIR = Path(__file__).parent

MIME_TYPE = "text/html;profile=mcp-app"

# view name -> (title, description, extra `ui` resource metadata)
_VIEWS: dict[str, tuple[str, str, dict[str, Any]]] = {
    "mail-list": (
        "Outlook mail list",
        "Interactive inbox view: open, mark, flag, and delete mails.",
        {},
    ),
    "mail-view": (
        "Outlook mail reader",
        "Full reading view for a single mail with attachment list.",
        {},
    ),
    "calendar": (
        "Outlook calendar agenda",
        "Agenda view of upcoming events grouped by day.",
        {},
    ),
    "contacts": (
        "Outlook contacts",
        "Searchable contact cards (saved contacts + org directory).",
        {"permissions": {"clipboardWrite": {}}},
    ),
    "tasks": (
        "Outlook tasks",
        "Task checklist: complete tasks and add new ones.",
        {},
    ),
}

_render_cache: dict[str, str] = {}


def view_uri(view: str) -> str:
    return f"ui://outlook/{view}.html"


def ui_meta(view: str) -> dict[str, Any]:
    """``meta=`` value for a tool that renders ``view`` when called."""
    if view not in _VIEWS:
        raise ValueError(f"unknown UI view: {view}")
    return {"ui": {"resourceUri": view_uri(view)}}


def ui_result(text: str, data: dict[str, Any]) -> CallToolResult:
    """Tool result with markdown for the model and data for the app."""
    return CallToolResult(
        content=[TextContent(type="text", text=text)],
        structuredContent=data,
    )


def _render(view: str) -> str:
    if view not in _render_cache:
        common = (_UI_DIR / "_common.html").read_text(encoding="utf-8")
        html = (_UI_DIR / f"{view}.html").read_text(encoding="utf-8")
        _render_cache[view] = html.replace("<!--%COMMON%-->", common)
    return _render_cache[view]


def _make_reader(view: str):
    # FastMCP requires a zero-arg function for static (non-template) resources.
    def _reader() -> str:
        return _render(view)

    return _reader


def register_ui(mcp) -> None:
    """Register every view as a ``ui://`` resource on the FastMCP server."""
    for view, (title, description, extra_ui_meta) in _VIEWS.items():
        _reader = _make_reader(view)
        mcp.resource(
            view_uri(view),
            name=f"outlook-ui-{view}",
            title=title,
            description=description,
            mime_type=MIME_TYPE,
            meta={"ui": {"prefersBorder": True, **extra_ui_meta}},
        )(_reader)
