"""MCP tool wrappers for calendar / meetings."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from outlook_mcp.client import calendar as cal_client
from outlook_mcp.schemas import Recurrence
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_events",
        annotations={
            "title": "List Outlook calendar events",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_events(
        start: Annotated[Optional[str], Field(description="ISO-8601 start of range. Defaults to now.")] = None,
        end: Annotated[Optional[str], Field(description="ISO-8601 end of range. Defaults to start+14d.")] = None,
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        include_recurrences: Annotated[bool, Field()] = True,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """List calendar events in a date range, including recurring instances."""
        data = await bridge.call(
            cal_client.list_events,
            start=start,
            end=end,
            limit=limit,
            include_recurrences=include_recurrences,
        )
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_get_event",
        annotations={
            "title": "Get Outlook calendar event",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_get_event(
        entry_id: Annotated[str, Field(description="EntryID of the event.")],
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """Fetch an event with attendees, body, and recurrence info."""
        data = await bridge.call(cal_client.get_event, entry_id=entry_id)
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_create_event",
        annotations={
            "title": "Create Outlook calendar event",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_create_event(
        subject: Annotated[str, Field(description="Event subject.")],
        start: Annotated[str, Field(description="ISO-8601 start time (e.g. '2026-05-01T14:00:00').")],
        end: Annotated[str, Field(description="ISO-8601 end time.")],
        location: Annotated[Optional[str], Field()] = None,
        body: Annotated[Optional[str], Field()] = None,
        attendees: Annotated[
            Optional[list[str]],
            Field(description="Email addresses. Adding any attendee turns this into a meeting."),
        ] = None,
        is_online_meeting: Annotated[bool, Field()] = False,
        reminder_minutes: Annotated[Optional[int], Field(ge=0, le=10080)] = 15,
        recurrence: Annotated[
            Optional[Recurrence],
            Field(description="Recurrence pattern. Omit for a single occurrence."),
        ] = None,
    ) -> str:
        """Create a calendar event or meeting invite."""
        data = await bridge.call(
            cal_client.create_event,
            subject=subject,
            start=start,
            end=end,
            location=location,
            body=body,
            attendees=attendees,
            is_online_meeting=is_online_meeting,
            reminder_minutes=reminder_minutes,
            recurrence=recurrence,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_update_event",
        annotations={
            "title": "Update Outlook calendar event",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_update_event(
        entry_id: Annotated[str, Field(description="EntryID of the event.")],
        subject: Annotated[Optional[str], Field()] = None,
        start: Annotated[Optional[str], Field()] = None,
        end: Annotated[Optional[str], Field()] = None,
        location: Annotated[Optional[str], Field()] = None,
        body: Annotated[Optional[str], Field()] = None,
    ) -> str:
        """Update fields on an existing event. Only non-null fields are written."""
        data = await bridge.call(
            cal_client.update_event,
            entry_id=entry_id,
            subject=subject,
            start=start,
            end=end,
            location=location,
            body=body,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_event",
        annotations={
            "title": "Delete Outlook calendar event",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_delete_event(
        entry_id: Annotated[str, Field(description="EntryID of the event.")],
    ) -> str:
        """Delete a calendar event (sends cancellation if it has attendees)."""
        data = await bridge.call(cal_client.delete_event, entry_id=entry_id)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_respond_event",
        annotations={
            "title": "Respond to a meeting invite",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_respond_event(
        entry_id: Annotated[str, Field(description="EntryID of the meeting.")],
        response: Annotated[str, Field(description="One of: 'accept', 'tentative', 'decline'.")],
        send_response: Annotated[bool, Field()] = True,
    ) -> str:
        """Accept / tentatively accept / decline a meeting invite."""
        data = await bridge.call(
            cal_client.respond_event,
            entry_id=entry_id,
            response=response,
            send_response=send_response,
        )
        return format_response(data, "json")
