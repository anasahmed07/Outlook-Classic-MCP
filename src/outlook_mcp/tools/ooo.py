"""MCP tool wrappers for Out-of-Office status."""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from outlook_mcp.client import ooo as ooo_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_get_out_of_office",
        annotations={
            "title": "Check Out-of-Office (auto-reply) status",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_get_out_of_office(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """Check whether Out-of-Office auto-reply is currently enabled."""
        data = await bridge.call(ooo_client.get_out_of_office)
        return format_response(data, response_format)
