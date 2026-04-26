"""MCP tool wrappers for mail."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from outlook_mcp.client import mail as mail_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_mails",
        annotations={
            "title": "List Outlook mails",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_mails(
        folder: Annotated[
            str,
            Field(
                description=(
                    "Folder name. Either a well-known name (inbox, sent, drafts, "
                    "deleted, junk) or a path like 'Inbox/Projects/Quinn'."
                ),
            ),
        ] = "inbox",
        limit: Annotated[int, Field(ge=1, le=100, description="Max items.")] = 25,
        offset: Annotated[int, Field(ge=0, description="Pagination offset.")] = 0,
        unread_only: Annotated[bool, Field(description="Return only unread.")] = False,
        since: Annotated[Optional[str], Field(description="ISO-8601 lower bound on ReceivedTime.")] = None,
        until: Annotated[Optional[str], Field(description="ISO-8601 upper bound on ReceivedTime.")] = None,
        from_address: Annotated[Optional[str], Field(description="Substring match on sender email.")] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """List mail items from a folder, newest first."""
        data = await bridge.call(
            mail_client.list_mails,
            folder=folder,
            limit=limit,
            offset=offset,
            unread_only=unread_only,
            since=since,
            until=until,
            from_address=from_address,
        )
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_search_mails",
        annotations={
            "title": "Search Outlook mails",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_search_mails(
        query: Annotated[str, Field(min_length=1, description="Search keywords or DASL filter.")],
        folder: Annotated[str, Field(description="Folder to search in.")] = "inbox",
        scope: Annotated[
            str,
            Field(
                description=(
                    "Where to look: 'subject_body' (default), 'subject', 'from', "
                    "or 'dasl' to pass `query` as a raw DASL @SQL filter."
                ),
            ),
        ] = "subject_body",
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """Search a mail folder by subject, body, or sender."""
        data = await bridge.call(
            mail_client.search_mails,
            query=query,
            folder=folder,
            limit=limit,
            scope=scope,
        )
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_get_mail",
        annotations={
            "title": "Get full Outlook mail",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_get_mail(
        entry_id: Annotated[str, Field(min_length=1, description="EntryID of the mail item.")],
        include_body: Annotated[bool, Field(description="Include full body + HTML body.")] = True,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """Fetch full body, headers, and attachment list for one mail item."""
        data = await bridge.call(
            mail_client.get_mail, entry_id=entry_id, include_body=include_body
        )
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_send_mail",
        annotations={
            "title": "Send Outlook mail",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_send_mail(
        to: Annotated[list[str], Field(min_length=1, description="Recipient addresses.")],
        subject: Annotated[str, Field(description="Subject line.")],
        body: Annotated[str, Field(description="Message body. Plain text unless html=True.")],
        cc: Annotated[Optional[list[str]], Field(description="CC recipients.")] = None,
        bcc: Annotated[Optional[list[str]], Field(description="BCC recipients.")] = None,
        html: Annotated[bool, Field(description="Treat body as HTML.")] = False,
        attachments: Annotated[Optional[list[str]], Field(description="Absolute paths to local files.")] = None,
        importance: Annotated[str, Field(description="One of: 'low', 'normal', 'high'.")] = "normal",
        save_only: Annotated[bool, Field(description="If true, save to Drafts instead of sending.")] = False,
    ) -> str:
        """Compose and send a new mail. Set save_only=True to save to Drafts."""
        data = await bridge.call(
            mail_client.send_mail,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            html=html,
            attachments=attachments,
            importance=importance,
            save_only=save_only,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_reply_mail",
        annotations={
            "title": "Reply to Outlook mail",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_reply_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to reply to.")],
        body: Annotated[str, Field(description="Reply body. Quoted original is appended.")],
        reply_all: Annotated[bool, Field(description="Reply to all recipients.")] = False,
        html: Annotated[bool, Field(description="Treat body as HTML.")] = False,
        attachments: Annotated[Optional[list[str]], Field(description="Files to attach.")] = None,
    ) -> str:
        """Reply (or reply-all) to an existing mail."""
        data = await bridge.call(
            mail_client.reply_mail,
            entry_id=entry_id,
            body=body,
            reply_all=reply_all,
            html=html,
            attachments=attachments,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_forward_mail",
        annotations={
            "title": "Forward Outlook mail",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_forward_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to forward.")],
        to: Annotated[list[str], Field(min_length=1, description="Forward recipients.")],
        body: Annotated[str, Field(description="Optional note above the forwarded mail.")] = "",
        cc: Annotated[Optional[list[str]], Field(description="CC recipients.")] = None,
        html: Annotated[bool, Field(description="Treat body as HTML.")] = False,
    ) -> str:
        """Forward an existing mail with an optional added note."""
        data = await bridge.call(
            mail_client.forward_mail,
            entry_id=entry_id,
            to=to,
            body=body,
            cc=cc,
            html=html,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_move_mail",
        annotations={
            "title": "Move Outlook mail to folder",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_move_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to move.")],
        target_folder: Annotated[str, Field(description="Destination folder.")],
    ) -> str:
        """Move a mail to another folder. Returns the new EntryID."""
        data = await bridge.call(
            mail_client.move_mail, entry_id=entry_id, target_folder=target_folder
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_mail",
        annotations={
            "title": "Delete Outlook mail (move to Deleted Items)",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_delete_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to delete.")],
    ) -> str:
        """Delete a mail (Outlook moves it to Deleted Items)."""
        data = await bridge.call(mail_client.delete_mail, entry_id=entry_id)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_mark_mail",
        annotations={
            "title": "Mark Outlook mail read/unread or flag it",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_mark_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail.")],
        read: Annotated[Optional[bool], Field(description="True=mark read, False=unread, None=no change.")] = None,
        flagged: Annotated[Optional[bool], Field(description="True=flag for follow-up, False=clear flag.")] = None,
    ) -> str:
        """Toggle read state and/or follow-up flag on a mail."""
        data = await bridge.call(
            mail_client.mark_mail, entry_id=entry_id, read=read, flagged=flagged
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_save_attachments",
        annotations={
            "title": "Save Outlook mail attachments to disk",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_save_attachments(
        entry_id: Annotated[str, Field(description="EntryID of the mail.")],
        output_dir: Annotated[str, Field(description="Absolute directory under your user profile.")],
        attachment_index: Annotated[
            Optional[int], Field(ge=1, description="1-indexed attachment. Omit to save all.")
        ] = None,
    ) -> str:
        """Save one or all attachments from a mail to a local directory."""
        data = await bridge.call(
            mail_client.save_attachments,
            entry_id=entry_id,
            output_dir=output_dir,
            attachment_index=attachment_index,
        )
        return format_response(data, "json")
