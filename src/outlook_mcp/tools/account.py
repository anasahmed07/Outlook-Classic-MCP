"""MCP tool wrappers for account / sanity checks."""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from outlook_mcp.client import account as account_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_whoami",
        annotations={
            "title": "Show current Outlook user and accounts",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_whoami(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """Return the current Outlook profile user and the list of mail accounts."""
        data = await bridge.call(account_client.whoami)
        return format_response(data, response_format)
