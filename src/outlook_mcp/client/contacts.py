"""Contact COM operations."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.constants import OL_CLASS_CONTACT, OL_FOLDER_CONTACTS


def _contact_summary(c: Any) -> dict[str, Any]:
    return {
        "entry_id": _safe_get(c, "EntryID"),
        "full_name": _safe_get(c, "FullName", ""),
        "email": _safe_get(c, "Email1Address"),
        "company": _safe_get(c, "CompanyName"),
        "job_title": _safe_get(c, "JobTitle"),
        "mobile": _safe_get(c, "MobileTelephoneNumber"),
    }


def _contact_full(c: Any) -> dict[str, Any]:
    return {
        **_contact_summary(c),
        "business_phone": _safe_get(c, "BusinessTelephoneNumber"),
        "home_phone": _safe_get(c, "HomeTelephoneNumber"),
        "address": _safe_get(c, "BusinessAddress"),
        "notes": _safe_get(c, "Body", ""),
    }


def list_contacts(outlook: Any, namespace: Any, *, limit: int = 50, offset: int = 0) -> dict[str, Any]:
    folder = namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)
    items = folder.Items
    items.Sort("[FullName]")
    results: list[dict[str, Any]] = []
    skipped = 0
    for c in items:
        if _safe_get(c, "Class") != OL_CLASS_CONTACT:
            continue
        if skipped < offset:
            skipped += 1
            continue
        results.append(_contact_summary(c))
        if len(results) >= limit:
            break
    return {
        "count": len(results),
        "offset": offset,
        "items": results,
        "has_more": len(results) == limit,
    }


def search_contacts(outlook: Any, namespace: Any, *, query: str, limit: int = 25) -> dict[str, Any]:
    folder = namespace.GetDefaultFolder(OL_FOLDER_CONTACTS)
    q = query.lower()
    results: list[dict[str, Any]] = []
    for c in folder.Items:
        if _safe_get(c, "Class") != OL_CLASS_CONTACT:
            continue
        haystack = " ".join(
            str(_safe_get(c, attr, "") or "")
            for attr in ("FullName", "Email1Address", "CompanyName", "JobTitle")
        ).lower()
        if q in haystack:
            results.append(_contact_summary(c))
            if len(results) >= limit:
                break
    return {"query": query, "count": len(results), "items": results}


def get_contact(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    return _contact_full(get_item_by_id(namespace, entry_id))
