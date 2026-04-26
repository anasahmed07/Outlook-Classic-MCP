"""MCP tool wrappers for Outlook categories."""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from outlook_mcp.client import categories as cat_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_categories",
        annotations={
            "title": "List Outlook color categories",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_categories(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """List the color categories configured in this Outlook profile."""
        data = await bridge.call(cat_client.list_categories)
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_set_category",
        annotations={
            "title": "Set categories on a mail / event / task",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_set_category(
        entry_id: Annotated[str, Field(description="EntryID of the item.")],
        categories: Annotated[
            str,
            Field(
                description=(
                    "Comma-separated category names (e.g. 'Important' or "
                    "'Work, Follow-up'). Empty string clears all categories."
                ),
            ),
        ],
    ) -> str:
        """Replace the categories on an item (mail / event / task)."""
        data = await bridge.call(
            cat_client.set_category, entry_id=entry_id, categories=categories
        )
        return format_response(data, "json")
