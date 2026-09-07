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
from outlook_mcp.utils.formatting import from_iso, to_dasl_date, to_iso, truncate

# How far before `start` to widen the Restrict() lower bound so events that
# begin earlier but still overlap the requested window (multi-day / all-day
# appointments) aren't excluded. Outlook's DASL Restrict() is only reliable
# ANDing two bounds on the *same* property (verified live), so we can't add
# a "[End] > start" clause directly — instead we over-fetch on [Start] and
# apply the precise overlap check in Python below.
_OVERLAP_LOOKBACK = dt.timedelta(days=3)


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


def _as_naive(value: Any) -> dt.datetime | None:
    if not hasattr(value, "year"):
        return None
    return dt.datetime(
        value.year, value.month, value.day, value.hour, value.minute, value.second
    )


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
    # Outlook requires the collection sorted ascending by [Start] BEFORE
    # IncludeRecurrences is set, or recurring series expand incorrectly.
    items.Sort("[Start]")
    if include_recurrences:
        items.IncludeRecurrences = True

    start_dt = from_iso(start) or dt.datetime.now()
    end_dt = from_iso(end) or (start_dt + dt.timedelta(days=14))

    # DASL property query (locale-independent ISO dates) instead of a plain
    # Jet literal — see to_dasl_date() for why the plain "[Start] >= 'mm/dd
    # /yyyy ...'" form silently returns the wrong events on non-US locales.
    restrict = (
        f"@SQL=\"urn:schemas:calendar:dtstart\" >= "
        f"'{to_dasl_date(start_dt - _OVERLAP_LOOKBACK)}' AND "
        f"\"urn:schemas:calendar:dtstart\" <= '{to_dasl_date(end_dt)}'"
    )
    filtered = items.Restrict(restrict)

    results: list[dict[str, Any]] = []
    truncated = False
    for ev in filtered:
        ev_start = _as_naive(_safe_get(ev, "Start"))
        ev_end = _as_naive(_safe_get(ev, "End"))
        if ev_start is None:
            continue
        # Restrict() above only bounds [Start]; apply the real overlap test
        # (and, for un-expanded recurring masters, the real range check —
        # Restrict() does not reliably filter those by date, verified live:
        # masters years outside the range still pass through when
        # IncludeRecurrences is left off).
        if ev_start > end_dt:
            continue
        if ev_end is not None and ev_end <= start_dt:
            continue
        if ev_end is None and ev_start < start_dt:
            continue
        if len(results) >= limit:
            truncated = True
            break
        results.append(_event_summary(ev))

    return {
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(results),
        "items": results,
        "truncated": truncated,
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
