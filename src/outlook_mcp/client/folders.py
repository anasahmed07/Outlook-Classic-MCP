"""Folder resolution + list/create.

All public functions take ``(outlook, namespace, **kwargs)`` so they can
run on the bridge thread directly. Internal helpers take ``namespace``
since they don't need the application object.
"""

from __future__ import annotations

from typing import Any

import pythoncom

from outlook_mcp.constants import DEFAULT_FOLDER_MAP, OL_FOLDER_INBOX
from outlook_mcp.errors import OutlookError


def _safe_get(item: Any, attr: str, default: Any = None) -> Any:
    try:
        return getattr(item, attr)
    except Exception:
        return default


def resolve_folder(namespace: Any, folder: str | None) -> Any:
    """Resolve a folder spec to a MAPIFolder COM object.

    Accepted forms:
      * None / "" -> default Inbox
      * one of the well-known names (``inbox``, ``sent``, ...)
      * a slash-delimited path (``Inbox/Projects/Quinn``), optionally
        prefixed with a store display name.
    """
    if not folder:
        return namespace.GetDefaultFolder(OL_FOLDER_INBOX)

    key = folder.strip().lower()
    if key in DEFAULT_FOLDER_MAP:
        return namespace.GetDefaultFolder(DEFAULT_FOLDER_MAP[key])

    segments = [s for s in folder.replace("\\", "/").split("/") if s]
    if not segments:
        raise OutlookError(f"Invalid folder spec: '{folder}'")

    first = segments[0].lower()
    root = None
    for store in namespace.Stores:
        if store.DisplayName.lower() == first:
            root = store.GetRootFolder()
            segments = segments[1:]
            break
    if root is None:
        root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent

    current = root
    for seg in segments:
        seg_lower = seg.lower()
        match = None
        for sub in current.Folders:
            if sub.Name.lower() == seg_lower:
                match = sub
                break
        if match is None:
            raise OutlookError(
                f"Folder '{seg}' not found under '{current.Name}'. "
                "Use outlook_list_folders to see available paths."
            )
        current = match
    return current


def get_item_by_id(namespace: Any, entry_id: str, store_id: str | None = None) -> Any:
    try:
        if store_id:
            return namespace.GetItemFromID(entry_id, store_id)
        return namespace.GetItemFromID(entry_id)
    except pythoncom.com_error as exc:
        raise OutlookError(
            f"Item with id '{entry_id}' not found. The id may be stale or "
            "the item may have been deleted."
        ) from exc


def list_folders(outlook: Any, namespace: Any, *, root: str | None = None, max_depth: int = 4) -> list[dict[str, Any]]:
    if root:
        start = resolve_folder(namespace, root)
    else:
        start = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent

    out: list[dict[str, Any]] = []

    def walk(folder: Any, path: str, depth: int) -> None:
        items_obj = _safe_get(folder, "Items")
        out.append(
            {
                "name": folder.Name,
                "path": path,
                "item_count": items_obj.Count if items_obj else 0,
                "unread_count": _safe_get(folder, "UnReadItemCount", 0),
                "default_item_type": _safe_get(folder, "DefaultItemType", -1),
            }
        )
        if depth >= max_depth:
            return
        for sub in folder.Folders:
            walk(sub, f"{path}/{sub.Name}", depth + 1)

    walk(start, start.Name, 0)
    return out


def create_folder(outlook: Any, namespace: Any, *, parent: str | None, name: str) -> dict[str, Any]:
    parent_folder = resolve_folder(namespace, parent)
    new_folder = parent_folder.Folders.Add(name)
    return {
        "name": new_folder.Name,
        "path": f"{parent_folder.Name}/{new_folder.Name}",
        "entry_id": new_folder.EntryID,
    }
