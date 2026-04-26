"""MCP tool wrappers for mail rules."""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from outlook_mcp.client import rules as rules_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_rules",
        annotations={
            "title": "List Outlook mail rules",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_rules(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """List all mail rules with their enabled status."""
        data = await bridge.call(rules_client.list_rules)
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_toggle_rule",
        annotations={
            "title": "Enable or disable a mail rule",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_toggle_rule(
        rule_name: Annotated[
            str,
            Field(min_length=1, description="Exact rule name (use outlook_list_rules)."),
        ],
        enabled: Annotated[bool, Field(description="True to enable, False to disable.")],
    ) -> str:
        """Enable or disable a mail rule by name. Modifies live mail rules."""
        data = await bridge.call(
            rules_client.toggle_rule, rule_name=rule_name, enabled=enabled
        )
        return format_response(data, "json")
