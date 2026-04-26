# Tool reference

Every `outlook_*` tool, with parameters, defaults, return shape, and notes on chaining. Skim the table of contents, then jump to the tools you need.

## Contents

- [Mail](#mail) — list_mails, search_mails, get_mail, send_mail, reply_mail, forward_mail, move_mail, delete_mail, mark_mail, save_attachments
- [Folders](#folders) — list_folders, create_folder
- [Calendar](#calendar) — list_events, get_event, create_event, update_event, delete_event, respond_event
- [Contacts](#contacts) — list_contacts, search_contacts, get_contact
- [Tasks](#tasks) — list_tasks, create_task, complete_task
- [Categories](#categories) — list_categories, set_category
- [Rules](#rules) — list_rules, toggle_rule
- [Out-of-Office](#out-of-office) — get_out_of_office
- [Account](#account) — whoami
- [Common return-field glossary](#common-return-field-glossary)

---

## Mail

### `outlook_list_mails`

List mail items from a folder, newest first. Read-only.

| Param            | Type      | Default     | Notes |
| ---------------- | --------- | ----------- | ----- |
| `folder`         | string    | `"inbox"`   | Well-known name or path. See SKILL.md → Folder references. |
| `limit`          | int 1–100 | `25`        | Max items to return. |
| `offset`         | int ≥0    | `0`         | Skip this many before returning. Use with the returned `next_offset` to paginate. |
| `unread_only`    | bool      | `false`     | If true, only unread mails. |
| `since`          | ISO-8601  | `null`      | Lower bound on `ReceivedTime`. |
| `until`          | ISO-8601  | `null`      | Upper bound on `ReceivedTime`. |
| `from_address`   | string    | `null`      | **Substring** match on sender. See gotcha re: `EX:/O=...` addresses. |
| `response_format`| `markdown`/`json` | `markdown` | Use `json` to extract `entry_id`s. |

**Returns** (`json` shape): `{ folder, count, offset, limit, items: [...], has_more, next_offset }`. Each item has: `entry_id, subject, from, from_address, to, received, unread, has_attachments, importance, preview` (200-char body excerpt).

### `outlook_search_mails`

Search a single folder by subject/body, subject-only, sender, or raw DASL. Read-only.

| Param            | Type     | Default          | Notes |
| ---------------- | -------- | ---------------- | ----- |
| `query`          | string   | required         | Search text, or a DASL @SQL filter when `scope='dasl'`. |
| `folder`         | string   | `"inbox"`        | Where to search. |
| `scope`          | enum     | `"subject_body"` | `subject_body` (default), `subject`, `from`, or `dasl`. |
| `limit`          | int 1–100| `25`             | |
| `response_format`| str      | `markdown`       | |

**Returns**: `{ query, scope, folder, count, items: [...] }`. Items have the same summary shape as `list_mails`.

`scope='dasl'` is for power use — pass a complete `@SQL=...` filter and the server applies it raw. Only reach for this when subject_body/subject/from can't express what the user wants.

### `outlook_get_mail`

Fetch the full body, all headers, and the attachment manifest for one mail. Read-only.

| Param          | Type   | Default | Notes |
| -------------- | ------ | ------- | ----- |
| `entry_id`     | string | required | From a list/search result. |
| `include_body` | bool   | `true`   | If false, omits `body` and `html_body`. Useful when you only need metadata. |
| `response_format` | str | `markdown` | |

**Returns**: `{ entry_id, conversation_id, subject, from, from_address, to, cc, bcc, received, sent, unread, importance, categories, attachments: [{index, filename, size_bytes}], body, html_body }`.

`attachments[].index` is **1-indexed**; pass it to `save_attachments` to save a single file.

### `outlook_send_mail`

Compose and send a new mail, or save it to Drafts. Has external side effect.

| Param          | Type      | Default   | Notes |
| -------------- | --------- | --------- | ----- |
| `to`           | list[str] | required  | One or more SMTP addresses. |
| `subject`      | string    | required  | |
| `body`         | string    | required  | Plain text unless `html=true`. |
| `cc`           | list[str] | `null`    | |
| `bcc`          | list[str] | `null`    | |
| `html`         | bool      | `false`   | When true, `body` is HTML. |
| `attachments`  | list[str] | `null`    | Absolute paths under user profile. |
| `importance`   | enum      | `"normal"`| `low` / `normal` / `high`. |
| `save_only`    | bool      | `false`   | **Save to Drafts instead of sending.** |

**Returns** (sent): `{ status: "sent", to, cc, bcc, subject }`. (Drafts): `{ status: "saved_to_drafts", entry_id, subject }`.

Always confirm the recipient list and subject with the user before calling this tool unless they have explicitly authorized you to send.

### `outlook_reply_mail`

Reply (or reply-all) to an existing mail. The original message is appended below your body, the same way Outlook's Reply button does it. Has external side effect.

| Param         | Type      | Default | Notes |
| ------------- | --------- | ------- | ----- |
| `entry_id`    | string    | required | The mail being replied to. |
| `body`        | string    | required | Your reply text. The quoted original is appended automatically. |
| `reply_all`   | bool      | `false`  | If true, includes the original CC list. |
| `html`        | bool      | `false`  | |
| `attachments` | list[str] | `null`   | |

**Returns**: `{ status: "sent", reply_all, in_reply_to, subject }`.

This sends immediately; there's no `save_only` flag on `reply_mail`. To stage a reply for review, copy the original's recipients yourself and call `send_mail` with `save_only=true` instead.

### `outlook_forward_mail`

Forward an existing mail to new recipients with an optional note above. Has external side effect.

| Param      | Type      | Default | Notes |
| ---------- | --------- | ------- | ----- |
| `entry_id` | string    | required | |
| `to`       | list[str] | required | |
| `body`     | string    | `""`     | Optional note prepended to the forwarded content. |
| `cc`       | list[str] | `null`   | |
| `html`     | bool      | `false`  | |

**Returns**: `{ status: "sent", forwarded, to, subject }`.

### `outlook_move_mail`

Move a mail to another folder.

| Param           | Type   | Default | Notes |
| --------------- | ------ | ------- | ----- |
| `entry_id`      | string | required | |
| `target_folder` | string | required | Well-known name or path. |

**Returns**: `{ status: "moved", new_entry_id, folder }`.

The `entry_id` changes when an item moves stores. **Use the returned `new_entry_id`** if you need to act on the moved item again.

### `outlook_delete_mail`

Soft-delete (moves to Deleted Items). Reversible by the user from Outlook.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

**Returns**: `{ status: "deleted", subject, entry_id }`.

### `outlook_mark_mail`

Toggle read state and/or follow-up flag.

| Param      | Type | Default | Notes |
| ---------- | ---- | ------- | ----- |
| `entry_id` | string | required | |
| `read`     | bool/null | `null` | `true` = mark read, `false` = mark unread, `null` = no change. |
| `flagged`  | bool/null | `null` | `true` = flag for follow-up, `false` = clear flag, `null` = no change. |

**Returns**: `{ status: "updated", entry_id, unread, flagged }`.

### `outlook_save_attachments`

Save one or all attachments from a mail to a local directory.

| Param              | Type    | Default | Notes |
| ------------------ | ------- | ------- | ----- |
| `entry_id`         | string  | required | |
| `output_dir`       | string  | required | Absolute path under user profile. Created if missing. |
| `attachment_index` | int ≥1  | `null`  | 1-indexed. Omit to save all. |

**Returns**: `{ status: "saved", count, files: [absolute paths], output_dir }`.

---

## Folders

### `outlook_list_folders`

Walk the folder tree under a root.

| Param            | Type    | Default | Notes |
| ---------------- | ------- | ------- | ----- |
| `root`           | string  | `null`  | Folder to start from. Default = the default mailbox root. |
| `max_depth`      | int 1–10| `4`     | How deep to walk. |
| `response_format`| str     | `markdown` | |

**Returns**: `{ count, items: [{name, path, item_count, unread_count, default_item_type}, ...] }`. The `path` strings are exactly what you pass back as a `folder` parameter elsewhere.

### `outlook_create_folder`

Create a sub-folder under a parent.

| Param    | Type   | Default   | Notes |
| -------- | ------ | --------- | ----- |
| `name`   | string | required  | New folder name. |
| `parent` | string | `"inbox"` | Where to put it. |

**Returns**: `{ name, path, entry_id }`.

---

## Calendar

### `outlook_list_events`

List calendar events in a date range, including expanded recurring instances.

| Param                 | Type    | Default        | Notes |
| --------------------- | ------- | -------------- | ----- |
| `start`               | ISO-8601| now            | |
| `end`                 | ISO-8601| `start + 14d`  | |
| `limit`               | int 1–200| `50`          | |
| `include_recurrences` | bool    | `true`         | If false, only the master entries — usually you want true. |
| `response_format`     | str     | `markdown`     | |

**Returns**: `{ start, end, count, items: [...] }`. Items: `entry_id, subject, start, end, location, organizer, is_recurring, all_day, preview` (200-char body excerpt).

### `outlook_get_event`

Full event detail, including attendees and their RSVP status.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |
| `response_format` | str | `markdown` | |

**Returns**: summary fields + `body, attendees: [{name, address, type, response}], reminder_minutes, categories`.

### `outlook_create_event`

Create a calendar event or meeting invite. **Adding any attendee turns this into a meeting that is sent immediately on success — there is no draft state for meeting invites.**

| Param               | Type           | Default | Notes |
| ------------------- | -------------- | ------- | ----- |
| `subject`           | string         | required | |
| `start`             | ISO-8601       | required | |
| `end`               | ISO-8601       | required | |
| `location`          | string         | `null`  | |
| `body`              | string         | `null`  | |
| `attendees`         | list[str]      | `null`  | Email addresses. Adding any value here makes this a meeting and sends invites. |
| `is_online_meeting` | bool           | `false` | Reserved — current behavior is to mark the meeting; the actual Teams/Zoom link is added by Outlook clients. |
| `reminder_minutes`  | int 0–10080    | `15`    | Minutes before start. |
| `recurrence`        | Recurrence obj | `null`  | See SKILL.md → Recurrence. |

**Returns**: `{ status: "created", entry_id, subject, start, end }`.

Confirm attendee list, times, and recurrence with the user before calling.

### `outlook_update_event`

Update fields on an event. Only non-null fields are written. Does **not** modify recurrence — for cadence changes, delete and recreate.

| Param      | Type     | Default | Notes |
| ---------- | -------- | ------- | ----- |
| `entry_id` | string   | required | |
| `subject`  | string   | `null`  | |
| `start`    | ISO-8601 | `null`  | |
| `end`      | ISO-8601 | `null`  | |
| `location` | string   | `null`  | |
| `body`     | string   | `null`  | |

**Returns**: `{ status: "updated", entry_id }`. If the event has attendees, Outlook may send an updated-meeting notification when this saves.

### `outlook_delete_event`

Delete a calendar event. **If the event has attendees, this sends a cancellation notice.**

**Returns**: `{ status: "deleted", subject, entry_id }`.

### `outlook_respond_event`

Respond to a meeting invite.

| Param           | Type   | Default | Notes |
| --------------- | ------ | ------- | ----- |
| `entry_id`      | string | required | |
| `response`      | enum   | required | `accept` / `tentative` / `decline`. |
| `send_response` | bool   | `true`   | Set false to record locally without emailing the organizer. |

**Returns**: `{ status: "responded", response }`.

---

## Contacts

### `outlook_list_contacts`

List contacts from the default Contacts folder, sorted by full name.

| Param            | Type     | Default | Notes |
| ---------------- | -------- | ------- | ----- |
| `limit`          | int 1–200| `50`    | |
| `offset`         | int ≥0   | `0`     | |
| `response_format`| str      | `markdown` | |

**Returns**: `{ count, offset, items: [...], has_more }`. Items: `entry_id, full_name, email, company, job_title, mobile`.

### `outlook_search_contacts`

Substring search across name, email, company, and job title.

| Param            | Type    | Default | Notes |
| ---------------- | ------- | ------- | ----- |
| `query`          | string  | required | |
| `limit`          | int 1–100| `25`   | |
| `response_format`| str     | `markdown` | |

### `outlook_get_contact`

Full contact record. Returns the summary fields plus `business_phone, home_phone, address, notes`.

---

## Tasks

### `outlook_list_tasks`

List tasks from the default Tasks folder, sorted by due date.

| Param                | Type     | Default | Notes |
| -------------------- | -------- | ------- | ----- |
| `limit`              | int 1–200| `50`    | |
| `include_completed`  | bool     | `false` | Default hides done tasks. |
| `response_format`    | str      | `markdown` | |

**Items**: `entry_id, subject, due_date, start_date, complete, percent_complete, importance, status`.

### `outlook_create_task`

| Param        | Type     | Default   | Notes |
| ------------ | -------- | --------- | ----- |
| `subject`    | string   | required  | |
| `due_date`   | ISO-8601 | `null`    | |
| `body`       | string   | `null`    | |
| `importance` | enum     | `"normal"`| low/normal/high. |
| `reminder`   | ISO-8601 | `null`    | Sets a reminder time. |

**Returns**: `{ status: "created", entry_id, subject }`.

### `outlook_complete_task`

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

Marks the task 100% complete. Returns `{ status: "completed", entry_id }`.

---

## Categories

### `outlook_list_categories`

Returns the color categories defined in the user's Outlook profile: `{ count, items: [{name, color}, ...] }`. Categories are profile-wide, not per-folder.

### `outlook_set_category`

Replace the categories on any item (mail, event, task).

| Param        | Type   | Default | Notes |
| ------------ | ------ | ------- | ----- |
| `entry_id`   | string | required | |
| `categories` | string | required | **Comma-separated names.** Empty string clears all. e.g. `"Important"` or `"Work, Follow-up"`. |

This **replaces** existing categories rather than adding to them. To add `Foo` to an item that already has `Bar`, send `"Bar, Foo"`. Get the current value first via `get_mail` / `get_event` if needed.

---

## Rules

### `outlook_list_rules`

Returns the user's mail rules with their on/off state: `{ count, items: [{index, name, enabled}] }`.

### `outlook_toggle_rule`

| Param       | Type   | Default | Notes |
| ----------- | ------ | ------- | ----- |
| `rule_name` | string | required | **Exact** name from `list_rules`. |
| `enabled`   | bool   | required | `true` to enable, `false` to disable. |

This change is live the moment it's saved — no staging buffer. Confirm the rule name with the user before calling.

---

## Out-of-Office

### `outlook_get_out_of_office`

Reports whether OOO auto-reply is currently on. Returns `{ out_of_office: bool, status: "on"|"off" }`, or `{ out_of_office: null, status: "unknown", note: ... }` if the property isn't readable on this profile.

There is **no tool to enable, disable, or schedule OOO**. Tell users to manage it via Outlook → File → Automatic Replies.

---

## Account

### `outlook_whoami`

Returns the bound user and the list of accounts: `{ current_user, accounts: [{display_name, smtp_address, user_name, account_type}, ...] }`. Useful as a sanity check when the user has multiple mailboxes or you want to confirm which mailbox you're acting on.

---

## Common return-field glossary

- `entry_id` — opaque, stable handle for an item. Pass back verbatim. Becomes invalid on delete; changes on cross-store move.
- `conversation_id` — groups mails in a thread. Same value across replies/forwards in one conversation.
- `from` — display name of the sender.
- `from_address` — sender address. **For Exchange senders, this is an `EX:/O=...` distinguished name, not SMTP.** Match by substring.
- `received` / `sent` / `start` / `end` / `due_date` — ISO-8601 strings.
- `unread` — bool. Note `mark_mail` returns `unread` (not `read`).
- `importance` — integer (0=low, 1=normal, 2=high).
- `categories` — comma-separated string of category names; empty string = none.
- `preview` — 200-char body excerpt. Not a substitute for `get_mail` / `get_event` when you need the full body.
- `has_more` / `next_offset` — pagination signals on list endpoints.
