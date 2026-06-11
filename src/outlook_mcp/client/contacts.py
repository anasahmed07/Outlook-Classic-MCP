"""Contact COM operations.

Corporate accounts commonly have an *empty* default Contacts folder —
real entries live in sibling contact folders (Recipient Cache, etc.)
and in the directory (Exchange Global Address List). So list/search
cover every contact folder in every store, and search can also scan
the GAL.
"""

from __future__ import annotations

from typing import Any, Iterator

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.constants import OL_CLASS_CONTACT, OL_CONTACT_ITEM


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


def _contact_folders(namespace: Any) -> Iterator[tuple[Any, str]]:
    """Yield ``(folder, path)`` for every contact folder in every store.

    ``DefaultItemType`` uses the same OlItemType enum as CreateItem, so
    contact folders are the ones reporting ``OL_CONTACT_ITEM`` (2).
    """

    def walk(folder: Any, path: str, depth: int) -> Iterator[tuple[Any, str]]:
        if _safe_get(folder, "DefaultItemType") == OL_CONTACT_ITEM:
            yield folder, path
        if depth >= 3:
            return
        try:
            subs = list(folder.Folders)
        except Exception:
            return
        for sub in subs:
            yield from walk(sub, f"{path}/{sub.Name}", depth + 1)

    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
        except Exception:
            continue
        yield from walk(root, str(store.DisplayName), 0)


def _matches_all(haystack: str, words: list[str]) -> bool:
    return all(w in haystack for w in words)


def list_contacts(outlook: Any, namespace: Any, *, limit: int = 50, offset: int = 0) -> dict[str, Any]:
    results: list[dict[str, Any]] = []
    skipped = 0
    for folder, path in _contact_folders(namespace):
        items = folder.Items
        try:
            items.Sort("[FullName]")
        except Exception:
            pass
        for c in items:
            if _safe_get(c, "Class") != OL_CLASS_CONTACT:
                continue
            if skipped < offset:
                skipped += 1
                continue
            row = _contact_summary(c)
            row["folder"] = path
            results.append(row)
            if len(results) >= limit:
                return {
                    "count": len(results),
                    "offset": offset,
                    "items": results,
                    "has_more": True,
                }
    return {
        "count": len(results),
        "offset": offset,
        "items": results,
        "has_more": False,
    }


def _search_directory(namespace: Any, words: list[str], limit: int) -> list[dict[str, Any]]:
    """Scan the Global Address List for entries matching all words.

    Matching is on the display name; details (SMTP, title, company) are
    fetched only for matches to keep the scan cheap. Large GALs take a
    few seconds — acceptable for an explicit search call.
    """
    out: list[dict[str, Any]] = []
    try:
        entries = namespace.GetGlobalAddressList().AddressEntries
    except Exception:
        return out  # no Exchange directory on this profile
    for entry in entries:
        name = str(_safe_get(entry, "Name", "") or "")
        if not _matches_all(name.lower(), words):
            continue
        row: dict[str, Any] = {
            "full_name": name,
            "email": None,
            "company": None,
            "job_title": None,
            "mobile": None,
            "source": "directory",
        }
        try:
            exu = entry.GetExchangeUser()
        except Exception:
            exu = None
        if exu is not None:
            row["email"] = _safe_get(exu, "PrimarySmtpAddress")
            row["company"] = _safe_get(exu, "CompanyName")
            row["job_title"] = _safe_get(exu, "JobTitle")
            row["mobile"] = _safe_get(exu, "MobileTelephoneNumber")
        out.append(row)
        if len(out) >= limit:
            break
    return out


def search_contacts(
    outlook: Any,
    namespace: Any,
    *,
    query: str,
    limit: int = 25,
    include_directory: bool = True,
) -> dict[str, Any]:
    words = [w for w in query.lower().split() if w] or [query.lower()]
    results: list[dict[str, Any]] = []
    seen_emails: set[str] = set()

    for folder, path in _contact_folders(namespace):
        if len(results) >= limit:
            break
        for c in folder.Items:
            if _safe_get(c, "Class") != OL_CLASS_CONTACT:
                continue
            haystack = " ".join(
                str(_safe_get(c, attr, "") or "")
                for attr in ("FullName", "Email1Address", "CompanyName", "JobTitle")
            ).lower()
            if _matches_all(haystack, words):
                row = _contact_summary(c)
                row["folder"] = path
                row["source"] = "contacts"
                results.append(row)
                if row.get("email"):
                    seen_emails.add(str(row["email"]).lower())
                if len(results) >= limit:
                    break

    searched_directory = False
    if include_directory and len(results) < limit:
        searched_directory = True
        for row in _search_directory(namespace, words, limit - len(results)):
            if row.get("email") and str(row["email"]).lower() in seen_emails:
                continue
            results.append(row)

    return {
        "query": query,
        "count": len(results),
        "items": results,
        "searched_directory": searched_directory,
    }


def get_contact(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    return _contact_full(get_item_by_id(namespace, entry_id))


def resolve_name(outlook: Any, namespace: Any, *, name: str) -> dict[str, Any]:
    """Resolve a name or address against the address book / directory.

    Same mechanism Outlook uses when you type a name in To: and press
    Ctrl+K. Unique matches resolve; ambiguous ones don't (no UI is
    shown), in which case the caller should try a fuller name.
    """
    recipient = namespace.CreateRecipient(name)
    recipient.Resolve()
    if not recipient.Resolved:
        return {
            "resolved": False,
            "query": name,
            "note": (
                "No unique match. Try the full display name or the exact "
                "address, or use outlook_search_contacts to browse matches."
            ),
        }
    entry = recipient.AddressEntry
    smtp = _safe_get(entry, "Address")
    try:
        exu = entry.GetExchangeUser()
    except Exception:
        exu = None
    if exu is not None and _safe_get(exu, "PrimarySmtpAddress"):
        smtp = exu.PrimarySmtpAddress
    return {
        "resolved": True,
        "query": name,
        "display_name": _safe_get(entry, "Name"),
        "smtp_address": smtp,
    }
