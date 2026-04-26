"""Task COM operations."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.constants import (
    IMPORTANCE_MAP,
    OL_CLASS_TASK,
    OL_FOLDER_TASKS,
    OL_IMPORTANCE_NORMAL,
    OL_TASK_ITEM,
)
from outlook_mcp.utils.formatting import from_iso, to_iso


def _task_summary(t: Any) -> dict[str, Any]:
    return {
        "entry_id": _safe_get(t, "EntryID"),
        "subject": _safe_get(t, "Subject", ""),
        "due_date": to_iso(_safe_get(t, "DueDate")),
        "start_date": to_iso(_safe_get(t, "StartDate")),
        "complete": bool(_safe_get(t, "Complete", False)),
        "percent_complete": _safe_get(t, "PercentComplete", 0),
        "importance": _safe_get(t, "Importance"),
        "status": _safe_get(t, "Status"),
    }


def list_tasks(
    outlook: Any,
    namespace: Any,
    *,
    limit: int = 50,
    include_completed: bool = False,
) -> dict[str, Any]:
    folder = namespace.GetDefaultFolder(OL_FOLDER_TASKS)
    items = folder.Items
    items.Sort("[DueDate]")
    results: list[dict[str, Any]] = []
    for t in items:
        if _safe_get(t, "Class") != OL_CLASS_TASK:
            continue
        if not include_completed and _safe_get(t, "Complete", False):
            continue
        results.append(_task_summary(t))
        if len(results) >= limit:
            break
    return {"count": len(results), "items": results}


def create_task(
    outlook: Any,
    namespace: Any,
    *,
    subject: str,
    due_date: str | None = None,
    body: str | None = None,
    importance: str = "normal",
    reminder: str | None = None,
) -> dict[str, Any]:
    task = outlook.CreateItem(OL_TASK_ITEM)
    task.Subject = subject
    if due_date:
        task.DueDate = from_iso(due_date)
    if body:
        task.Body = body
    task.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)
    if reminder:
        task.ReminderSet = True
        task.ReminderTime = from_iso(reminder)
    task.Save()
    return {"status": "created", "entry_id": task.EntryID, "subject": task.Subject}


def complete_task(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    task = get_item_by_id(namespace, entry_id)
    task.MarkComplete()
    task.Save()
    return {"status": "completed", "entry_id": entry_id}
