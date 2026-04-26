"""Outlook COM enum values used across the client modules.

We hard-code the integer values rather than depending on
``win32com.client.constants``: that requires the typelib cache to have
been generated for the current pywin32 install, which isn't always true
on a fresh machine.
"""

from __future__ import annotations

# OlDefaultFolders
OL_FOLDER_DELETED = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_SENT = 5
OL_FOLDER_INBOX = 6
OL_FOLDER_CALENDAR = 9
OL_FOLDER_CONTACTS = 10
OL_FOLDER_NOTES = 12
OL_FOLDER_TASKS = 13
OL_FOLDER_DRAFTS = 16
OL_FOLDER_JUNK = 23

DEFAULT_FOLDER_MAP: dict[str, int] = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT,
    "drafts": OL_FOLDER_DRAFTS,
    "deleted": OL_FOLDER_DELETED,
    "trash": OL_FOLDER_DELETED,
    "outbox": OL_FOLDER_OUTBOX,
    "junk": OL_FOLDER_JUNK,
    "spam": OL_FOLDER_JUNK,
    "calendar": OL_FOLDER_CALENDAR,
    "contacts": OL_FOLDER_CONTACTS,
    "tasks": OL_FOLDER_TASKS,
    "notes": OL_FOLDER_NOTES,
}

# OlObjectClass — the .Class property on items
OL_CLASS_MAIL = 43
OL_CLASS_APPOINTMENT = 26
OL_CLASS_CONTACT = 40
OL_CLASS_TASK = 48
OL_CLASS_MEETING_REQUEST = 53

# OlItemType — first arg to Application.CreateItem
OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1
OL_CONTACT_ITEM = 2
OL_TASK_ITEM = 3

# OlImportance
OL_IMPORTANCE_LOW = 0
OL_IMPORTANCE_NORMAL = 1
OL_IMPORTANCE_HIGH = 2

IMPORTANCE_MAP: dict[str, int] = {
    "low": OL_IMPORTANCE_LOW,
    "normal": OL_IMPORTANCE_NORMAL,
    "high": OL_IMPORTANCE_HIGH,
}

# OlMailRecipientType
OL_TO = 1
OL_CC = 2
OL_BCC = 3

# OlBodyFormat
OL_FORMAT_PLAIN = 1
OL_FORMAT_HTML = 2

# OlMeetingResponse
OL_MEETING_TENTATIVE = 2
OL_MEETING_ACCEPTED = 3
OL_MEETING_DECLINED = 4

# OlRecurrenceType
OL_RECURS_DAILY = 0
OL_RECURS_WEEKLY = 1
OL_RECURS_MONTHLY = 2
OL_RECURS_YEARLY = 5

# OlMeetingStatus
OL_MEETING = 1

# Out-of-office MAPI property tag (PR_OOF_STATE)
OOF_PROPTAG = "http://schemas.microsoft.com/mapi/proptag/0x661D000B"
