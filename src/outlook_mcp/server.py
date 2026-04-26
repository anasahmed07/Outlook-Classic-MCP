"""FastMCP server construction and tool registration."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from outlook_mcp.bridge import OutlookBridge
from outlook_mcp.tools import register_all

INSTRUCTIONS = """\
This MCP server gives you full access to Microsoft Outlook desktop on
Windows via COM automation. It can read and send mail, manage calendar
events, contacts, tasks, color categories, mail rules, and check
Out-of-Office status.

Authentication piggybacks on whatever account Outlook is signed into —
no Microsoft Graph API, no Entra app registration, no OAuth tokens.

PREREQUISITE: classic Outlook (OUTLOOK.EXE) must be installed. The
server auto-launches Outlook on first call; the new "modern" Outlook
(olk.exe) is NOT supported.

Tool categories (all prefixed `outlook_`):
  - Mail: list, search, get, send, reply, forward, move, delete, mark, save_attachments
  - Folders: list_folders, create_folder
  - Calendar: list_events, get_event, create_event, update_event, delete_event, respond_event
  - Contacts: list_contacts, search_contacts, get_contact
  - Tasks: list_tasks, create_task, complete_task
  - Categories: list_categories, set_category
  - Rules: list_rules, toggle_rule  (modifies live mail rules — confirm first)
  - Out-of-Office: get_out_of_office
  - Account: whoami  (sanity check on the bound mailbox)

Most read tools accept response_format='markdown' (default) or 'json'.
Item references use Outlook EntryID strings; list tools return them on
every item — pass them back to detail/edit/delete tools.
"""


def build_server() -> tuple[FastMCP, OutlookBridge]:
    """Construct the FastMCP instance, bridge, and wire all tools.

    The bridge is *not* started here — entrypoint.main() does that so
    the readiness wait happens after logging is configured.
    """
    mcp = FastMCP("outlook_mcp", instructions=INSTRUCTIONS)
    bridge = OutlookBridge()
    register_all(mcp, bridge)
    return mcp, bridge
