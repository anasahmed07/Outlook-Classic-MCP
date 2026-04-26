"""Category COM operations.

Outlook categories are profile-wide (not per-store), so we read them
straight off ``namespace.Categories``.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.folders import get_item_by_id


def list_categories(outlook: Any, namespace: Any) -> dict[str, Any]:
    items = []
    cats = namespace.Categories
    for i in range(cats.Count):
        cat = cats.Item(i + 1)
        items.append({"name": cat.Name, "color": cat.Color})
    return {"count": len(items), "items": items}


def set_category(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    categories: str,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    item.Categories = categories
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "subject": getattr(item, "Subject", ""),
        "categories": item.Categories or "",
    }
