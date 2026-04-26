"""MCP tool wrappers for tasks."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from outlook_mcp.client import tasks as tasks_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_tasks",
        annotations={
            "title": "List Outlook tasks",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_tasks(
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        include_completed: Annotated[bool, Field()] = False,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """List tasks (to-dos). By default only incomplete tasks are returned."""
        data = await bridge.call(
            tasks_client.list_tasks, limit=limit, include_completed=include_completed
        )
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_create_task",
        annotations={
            "title": "Create Outlook task",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_create_task(
        subject: Annotated[str, Field(description="Task subject.")],
        due_date: Annotated[Optional[str], Field(description="ISO-8601 date/datetime.")] = None,
        body: Annotated[Optional[str], Field()] = None,
        importance: Annotated[str, Field(description="low/normal/high.")] = "normal",
        reminder: Annotated[Optional[str], Field(description="ISO-8601 reminder time.")] = None,
    ) -> str:
        """Create a new task in the default Tasks folder."""
        data = await bridge.call(
            tasks_client.create_task,
            subject=subject,
            due_date=due_date,
            body=body,
            importance=importance,
            reminder=reminder,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_complete_task",
        annotations={
            "title": "Mark Outlook task complete",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_complete_task(
        entry_id: Annotated[str, Field(description="EntryID of the task.")],
    ) -> str:
        """Mark a task as 100% complete."""
        data = await bridge.call(tasks_client.complete_task, entry_id=entry_id)
        return format_response(data, "json")
