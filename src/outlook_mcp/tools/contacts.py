"""MCP tool wrappers for contacts."""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from outlook_mcp.client import contacts as contacts_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_contacts",
        annotations={
            "title": "List Outlook contacts",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_contacts(
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        offset: Annotated[int, Field(ge=0)] = 0,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """List contacts from the default Contacts folder, sorted by name."""
        data = await bridge.call(
            contacts_client.list_contacts, limit=limit, offset=offset
        )
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_search_contacts",
        annotations={
            "title": "Search Outlook contacts",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_search_contacts(
        query: Annotated[str, Field(min_length=1, description="Search term.")],
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """Search contacts by name, email, company, or job title (substring)."""
        data = await bridge.call(
            contacts_client.search_contacts, query=query, limit=limit
        )
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_get_contact",
        annotations={
            "title": "Get Outlook contact",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_get_contact(
        entry_id: Annotated[str, Field(description="EntryID of the contact.")],
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """Fetch full details for one contact."""
        data = await bridge.call(contacts_client.get_contact, entry_id=entry_id)
        return format_response(data, response_format)
