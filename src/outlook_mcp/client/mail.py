"""Mail COM operations."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id, resolve_folder
from outlook_mcp.constants import (
    IMPORTANCE_MAP,
    OL_CLASS_MAIL,
    OL_CLASS_MEETING_REQUEST,
    OL_FORMAT_HTML,
    OL_FORMAT_PLAIN,
    OL_IMPORTANCE_NORMAL,
    OL_MAIL_ITEM,
)
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.formatting import from_iso, to_iso, truncate
from outlook_mcp.utils.paths import validate_attachment_path, validate_output_dir
from outlook_mcp.utils.safety import safe_dasl


def _mail_summary(item: Any) -> dict[str, Any]:
    attachments = _safe_get(item, "Attachments")
    return {
        "entry_id": _safe_get(item, "EntryID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "has_attachments": attachments.Count > 0 if attachments else False,
        "importance": _safe_get(item, "Importance"),
        "preview": truncate(_safe_get(item, "Body", ""), 200),
    }


def _mail_full(item: Any, include_body: bool = True) -> dict[str, Any]:
    attachments = []
    if _safe_get(item, "Attachments"):
        for i, att in enumerate(item.Attachments, start=1):
            attachments.append(
                {
                    "index": i,
                    "filename": att.FileName,
                    "size_bytes": _safe_get(att, "Size"),
                }
            )
    result = {
        "entry_id": _safe_get(item, "EntryID"),
        "conversation_id": _safe_get(item, "ConversationID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "cc": _safe_get(item, "CC", ""),
        "bcc": _safe_get(item, "BCC", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "sent": to_iso(_safe_get(item, "SentOn")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "importance": _safe_get(item, "Importance"),
        "categories": _safe_get(item, "Categories", ""),
        "attachments": attachments,
    }
    if include_body:
        result["body"] = _safe_get(item, "Body", "")
        result["html_body"] = _safe_get(item, "HTMLBody", "")
    return result


def list_mails(
    outlook: Any,
    namespace: Any,
    *,
    folder: str | None = "inbox",
    limit: int = 25,
    offset: int = 0,
    unread_only: bool = False,
    since: str | None = None,
    until: str | None = None,
    from_address: str | None = None,
) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    clauses: list[str] = []
    if unread_only:
        clauses.append("[UnRead] = True")
    since_dt = from_iso(since)
    until_dt = from_iso(until)
    if since_dt:
        clauses.append(f"[ReceivedTime] >= '{since_dt.strftime('%m/%d/%Y %H:%M %p')}'")
    if until_dt:
        clauses.append(f"[ReceivedTime] <= '{until_dt.strftime('%m/%d/%Y %H:%M %p')}'")

    if clauses:
        items = items.Restrict(" AND ".join(clauses))

    from_lower = from_address.lower() if from_address else None
    results: list[dict[str, Any]] = []
    skipped = 0
    for item in items:
        cls = _safe_get(item, "Class")
        if cls not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
            continue
        if from_lower:
            sender = (_safe_get(item, "SenderEmailAddress") or "").lower()
            if from_lower not in sender:
                continue
        if skipped < offset:
            skipped += 1
            continue
        results.append(_mail_summary(item))
        if len(results) >= limit:
            break

    return {
        "folder": f.Name,
        "count": len(results),
        "offset": offset,
        "limit": limit,
        "items": results,
        "has_more": len(results) == limit,
        "next_offset": offset + len(results) if len(results) == limit else None,
    }


def search_mails(
    outlook: Any,
    namespace: Any,
    *,
    query: str,
    folder: str | None = "inbox",
    limit: int = 25,
    scope: str = "subject_body",
) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    if scope == "dasl":
        # Caller is explicitly passing a raw DASL filter; don't mangle it.
        filtered = items.Restrict(query)
    elif scope == "subject":
        esc = safe_dasl(query)
        filtered = items.Restrict(
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{esc}%'"
        )
    elif scope == "from":
        esc = safe_dasl(query)
        filtered = items.Restrict(
            f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{esc}%' OR "
            f"\"urn:schemas:httpmail:fromname\" LIKE '%{esc}%'"
        )
    else:  # subject_body
        esc = safe_dasl(query)
        filtered = items.Restrict(
            f"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{esc}%' OR "
            f"\"urn:schemas:httpmail:textdescription\" LIKE '%{esc}%')"
        )

    results: list[dict[str, Any]] = []
    for item in filtered:
        cls = _safe_get(item, "Class")
        if cls not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
            continue
        results.append(_mail_summary(item))
        if len(results) >= limit:
            break

    return {
        "query": query,
        "scope": scope,
        "folder": f.Name,
        "count": len(results),
        "items": results,
    }


def get_mail(outlook: Any, namespace: Any, *, entry_id: str, include_body: bool = True) -> dict[str, Any]:
    return _mail_full(get_item_by_id(namespace, entry_id), include_body=include_body)


def send_mail(
    outlook: Any,
    namespace: Any,
    *,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
    importance: str = "normal",
    save_only: bool = False,
) -> dict[str, Any]:
    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    if bcc:
        mail.BCC = "; ".join(bcc)
    mail.Subject = subject
    if html:
        mail.BodyFormat = OL_FORMAT_HTML
        mail.HTMLBody = body
    else:
        mail.BodyFormat = OL_FORMAT_PLAIN
        mail.Body = body
    mail.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)

    for raw_path in attachments or []:
        mail.Attachments.Add(validate_attachment_path(raw_path))

    if save_only:
        mail.Save()
        return {
            "status": "saved_to_drafts",
            "entry_id": mail.EntryID,
            "subject": mail.Subject,
        }

    mail.Send()
    return {
        "status": "sent",
        "to": to,
        "cc": cc or [],
        "bcc": bcc or [],
        "subject": subject,
    }


def reply_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    attachments: list[str] | None = None,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    if html:
        reply.BodyFormat = OL_FORMAT_HTML
        reply.HTMLBody = body + (reply.HTMLBody or "")
    else:
        reply.Body = body + "\n\n" + (reply.Body or "")
    for raw_path in attachments or []:
        reply.Attachments.Add(validate_attachment_path(raw_path))
    reply.Send()
    return {
        "status": "sent",
        "reply_all": reply_all,
        "in_reply_to": entry_id,
        "subject": reply.Subject,
    }


def forward_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    to: list[str],
    body: str = "",
    cc: list[str] | None = None,
    html: bool = False,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id)
    fwd = original.Forward()
    fwd.To = "; ".join(to)
    if cc:
        fwd.CC = "; ".join(cc)
    if body:
        if html:
            fwd.BodyFormat = OL_FORMAT_HTML
            fwd.HTMLBody = body + (fwd.HTMLBody or "")
        else:
            fwd.Body = body + "\n\n" + (fwd.Body or "")
    fwd.Send()
    return {"status": "sent", "forwarded": entry_id, "to": to, "subject": fwd.Subject}


def move_mail(outlook: Any, namespace: Any, *, entry_id: str, target_folder: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    target = resolve_folder(namespace, target_folder)
    moved = item.Move(target)
    return {"status": "moved", "new_entry_id": moved.EntryID, "folder": target.Name}


def delete_mail(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    subject = _safe_get(item, "Subject", "")
    item.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id}


def mark_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    read: bool | None = None,
    flagged: bool | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    if read is not None:
        item.UnRead = not read
    if flagged is not None:
        item.FlagStatus = 2 if flagged else 0
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "unread": bool(item.UnRead),
        "flagged": item.FlagStatus == 2,
    }


def save_attachments(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    output_dir: str,
    attachment_index: int | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    out_dir = validate_output_dir(output_dir)
    saved: list[str] = []
    attachments = list(item.Attachments)
    if attachment_index is not None:
        if attachment_index < 1 or attachment_index > len(attachments):
            raise OutlookError(
                f"attachment_index {attachment_index} out of range "
                f"(message has {len(attachments)} attachments, 1-indexed)."
            )
        attachments = [attachments[attachment_index - 1]]
    import os

    for att in attachments:
        target = os.path.join(out_dir, att.FileName)
        att.SaveAsFile(target)
        saved.append(target)
    return {
        "status": "saved",
        "count": len(saved),
        "files": saved,
        "output_dir": out_dir,
    }
