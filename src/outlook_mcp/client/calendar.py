"""Calendar (appointment / meeting) COM operations."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.constants import (
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_MEETING,
    OL_MEETING_ACCEPTED,
    OL_MEETING_DECLINED,
    OL_MEETING_TENTATIVE,
    OL_RECURS_DAILY,
    OL_RECURS_MONTHLY,
    OL_RECURS_WEEKLY,
    OL_RECURS_YEARLY,
    OL_TO,
)
from outlook_mcp.errors import OutlookError
from outlook_mcp.schemas import Recurrence
from outlook_mcp.utils.formatting import from_iso, to_iso, truncate


_RECURRENCE_TYPE_MAP = {
    "daily": OL_RECURS_DAILY,
    "weekly": OL_RECURS_WEEKLY,
    "monthly": OL_RECURS_MONTHLY,
    "yearly": OL_RECURS_YEARLY,
}


def _event_summary(ev: Any) -> dict[str, Any]:
    recur_state = _safe_get(ev, "RecurrenceState", 0)
    return {
        "entry_id": _safe_get(ev, "EntryID"),
        "subject": _safe_get(ev, "Subject", ""),
        "start": to_iso(_safe_get(ev, "Start")),
        "end": to_iso(_safe_get(ev, "End")),
        "location": _safe_get(ev, "Location", ""),
        "organizer": _safe_get(ev, "Organizer", ""),
        "is_recurring": bool(recur_state),
        "all_day": bool(_safe_get(ev, "AllDayEvent", False)),
        "preview": truncate(_safe_get(ev, "Body", ""), 200),
    }


def _event_full(ev: Any) -> dict[str, Any]:
    attendees = []
    if _safe_get(ev, "Recipients"):
        for r in ev.Recipients:
            attendees.append(
                {
                    "name": _safe_get(r, "Name"),
                    "address": _safe_get(r, "Address"),
                    "type": _safe_get(r, "Type"),
                    "response": _safe_get(r, "MeetingResponseStatus"),
                }
            )
    return {
        **_event_summary(ev),
        "body": _safe_get(ev, "Body", ""),
        "attendees": attendees,
        "reminder_minutes": _safe_get(ev, "ReminderMinutesBeforeStart"),
        "categories": _safe_get(ev, "Categories", ""),
    }


def _apply_recurrence(appt: Any, spec: Recurrence) -> None:
    pattern = appt.GetRecurrencePattern()
    pattern.RecurrenceType = _RECURRENCE_TYPE_MAP[spec.type]
    pattern.Interval = spec.interval
    if spec.occurrences is not None:
        pattern.Occurrences = spec.occurrences
    if spec.end_date is not None:
        pattern.PatternEndDate = from_iso(spec.end_date)


def list_events(
    outlook: Any,
    namespace: Any,
    *,
    start: str | None = None,
    end: str | None = None,
    limit: int = 50,
    include_recurrences: bool = True,
) -> dict[str, Any]:
    cal = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = cal.Items
    if include_recurrences:
        items.IncludeRecurrences = True
    items.Sort("[Start]")

    start_dt = from_iso(start) or dt.datetime.now()
    end_dt = from_iso(end) or (start_dt + dt.timedelta(days=14))

    restrict = (
        f"[Start] >= '{start_dt.strftime('%m/%d/%Y %H:%M %p')}' AND "
        f"[Start] <= '{end_dt.strftime('%m/%d/%Y %H:%M %p')}'"
    )
    filtered = items.Restrict(restrict)

    results: list[dict[str, Any]] = []
    for ev in filtered:
        results.append(_event_summary(ev))
        if len(results) >= limit:
            break

    return {
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(results),
        "items": results,
    }


def get_event(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    return _event_full(get_item_by_id(namespace, entry_id))


def create_event(
    outlook: Any,
    namespace: Any,
    *,
    subject: str,
    start: str,
    end: str,
    location: str | None = None,
    body: str | None = None,
    attendees: list[str] | None = None,
    is_online_meeting: bool = False,
    reminder_minutes: int | None = 15,
    recurrence: Recurrence | None = None,
) -> dict[str, Any]:
    appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
    appt.Subject = subject
    appt.Start = from_iso(start)
    appt.End = from_iso(end)
    if location:
        appt.Location = location
    if body:
        appt.Body = body
    if reminder_minutes is not None:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
    if attendees:
        appt.MeetingStatus = OL_MEETING
        for addr in attendees:
            rec = appt.Recipients.Add(addr)
            rec.Type = OL_TO
        appt.Recipients.ResolveAll()
    if recurrence is not None:
        _apply_recurrence(appt, recurrence)
    appt.Save()
    if attendees:
        appt.Send()
    return {
        "status": "created",
        "entry_id": appt.EntryID,
        "subject": appt.Subject,
        "start": to_iso(appt.Start),
        "end": to_iso(appt.End),
    }


def update_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
) -> dict[str, Any]:
    ev = get_item_by_id(namespace, entry_id)
    if subject is not None:
        ev.Subject = subject
    if start is not None:
        ev.Start = from_iso(start)
    if end is not None:
        ev.End = from_iso(end)
    if location is not None:
        ev.Location = location
    if body is not None:
        ev.Body = body
    ev.Save()
    return {"status": "updated", "entry_id": entry_id}


def delete_event(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    ev = get_item_by_id(namespace, entry_id)
    subject = _safe_get(ev, "Subject", "")
    ev.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id}


def respond_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    response: str,
    send_response: bool = True,
) -> dict[str, Any]:
    ev = get_item_by_id(namespace, entry_id)
    code = {
        "accept": OL_MEETING_ACCEPTED,
        "tentative": OL_MEETING_TENTATIVE,
        "decline": OL_MEETING_DECLINED,
    }.get(response.lower())
    if code is None:
        raise OutlookError("response must be one of: 'accept', 'tentative', 'decline'.")
    resp = ev.Respond(code, True)
    if send_response and resp is not None:
        resp.Send()
    return {"status": "responded", "response": response}
