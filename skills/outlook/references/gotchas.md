# Gotchas

Every quirk and failure mode worth knowing, organized by *what the user sees*. Read this before doing anything that has external side effects, and consult the matching entry the moment something looks weird.

## "You said you sent it but I don't see it in Sent Items."

Almost always **Outlook's Programmatic Access security setting** silently blocking the send. On corporate-managed Outlook installs, sending mail from automation can be gated by Trust Center. Symptoms: `send_mail` / `reply_mail` / `forward_mail` returns success, the mail never appears in Sent Items, no error is surfaced to you.

**Tell the user**: open Outlook → File → Options → Trust Center → Trust Center Settings → Programmatic Access. The setting that lets the integration send is "Never warn me about suspicious activity". On locked-down corporate machines IT may have to whitelist the Python interpreter as a trusted publisher, in which case there's nothing the user can do at the application layer.

`delete_mail` is also gated by this setting.

## Sender filter "from_address=alice@example.com" returns nothing

Exchange senders' `from_address` is **not an SMTP address**. It's an `EX:/O=ExchangeLabs/OU=...` distinguished name. The substring filter on `list_mails` matches against this raw string, which won't contain the SMTP address.

**Fix**: pass a name fragment that *will* appear in the DN — usually the user's first name, last name, or company alias. `from_address="sarah"` is far more reliable than `from_address="sarah@example.com"`. Or use `outlook_search_mails(scope="from", query="sarah")` which matches both the from-name and from-email fields via DASL.

## Toggling a rule changed the user's actual mail flow before they confirmed

`outlook_toggle_rule` modifies a live mail rule the moment it returns successfully. There's no staging buffer, no preview, no undo. If you call it on the wrong rule name, the user's mail is now being filed (or not filed) differently for real.

**Always**: `outlook_list_rules` first → confirm the **exact** rule name with the user → only then `outlook_toggle_rule`.

## "Set my OOO to on for the next two weeks" — and it just doesn't

The integration has `outlook_get_out_of_office` (read state) but no tool to set it. The auto-reply state lives in an Exchange property that requires EWS or Graph to write, neither of which the integration uses.

**Tell the user** to set OOO via Outlook → File → Automatic Replies. After they do, you can confirm it took effect with `outlook_get_out_of_office`.

## "Why did everyone get an email when I just rescheduled the meeting?"

Updating an event with attendees usually triggers Outlook to send an updated-meeting notification. Likewise, deleting an event with attendees sends a cancellation. There is no flag to suppress this.

**Workaround**: warn the user up front. If they want a quiet local-only change, they have to do it in Outlook directly with "Don't send updates" — the integration can't suppress it.

## Created a meeting and the invite went out before the user reviewed

`outlook_create_event` with `attendees=[...]` sends the invite **the moment the call succeeds**. There is no draft state for meeting invites the way there is for plain mail (`save_only=true`).

**Mitigation**: confirm subject, time, and attendee list with the user before calling. If the user wants to compose carefully, create the event without `attendees` first, then have them add attendees manually in Outlook.

## "I replied — can you save that to drafts so I can edit it first?"

`outlook_reply_mail` doesn't have a `save_only` flag. It sends.

**Workaround**: build the reply yourself and save it via `outlook_send_mail(save_only=true)`:

```
to = [original.from_address]                        # or all of original.to/cc for reply-all
subject = "Re: " + original.subject                 # don't double-prefix
body = your_reply + "\n\n" + original.body          # quote manually
outlook_send_mail(to=to, subject=subject, body=body, save_only=true)
```

Tell the user the draft is in the Drafts folder for them to review.

## Item not found / EntryID errors after a move

Moving an item with `outlook_move_mail` returns a `new_entry_id`. The original `entry_id` no longer resolves.

**Always** use `new_entry_id` for any subsequent action on the moved item. If you've lost track, re-list the destination folder.

## "Schedule it daily for the next 6 months" — and then they want to change cadence

`outlook_update_event` cannot modify the recurrence pattern. It only rewrites scalar fields (subject, start, end, location, body).

**Fix**: to change cadence (or end date or interval), delete the event and recreate it with the new `recurrence` object. If it has attendees, this means a cancellation + a new invite — warn the user.

## Set categories overwrote the existing tags

`outlook_set_category` **replaces** the categories field, it doesn't append. Categories are stored as a comma-separated string and the tool overwrites that whole string.

**Pattern**: read first, merge, then set.

```
existing = outlook_get_mail(entry_id, include_body=false).categories
new_value = ", ".join(filter(None, [existing, "Important"]))
outlook_set_category(entry_id=entry_id, categories=new_value)
```

To **clear** all categories, pass an empty string.

## "It says my category 'Urgent' doesn't exist."

`outlook_set_category` doesn't validate the names against the profile's defined categories. If you set a category that isn't defined, Outlook will accept the string but the item won't get the color and the user won't see a recognizable tag in the UI.

**Best practice**: `outlook_list_categories` first; only assign names that come back from there. If the user wants a brand-new category, they need to create it in Outlook (Home → Categorize → All Categories → New) — there is no `create_category` tool.

## Attachment / output_dir failures

The integration enforces a **user-profile sandbox** on all filesystem paths. Passing `D:\stuff\file.pdf` or `\\fileserver\share\foo` will fail with a sandbox error.

**Tell the user** their options:

1. Move the file under their user profile (`C:\Users\<them>\...`) — easiest.
2. Set the env var `OUTLOOK_MCP_ALLOW_ANY_PATH=1` in the integration's environment to disable the sandbox — they have to restart the integration after changing it. Use only when they really need it (e.g. corporate file shares).

## Path errors specifically about "must be absolute"

Relative paths like `Documents/file.pdf` always fail. Always pass `C:\\Users\\<user>\\Documents\\file.pdf` (or with forward slashes — both work).

## First call of the session takes 5–10 seconds

Outlook's COM surface needs to warm up the first time. Subsequent calls are fast. **Don't retry** as if the slow first call had failed; you'll just queue up duplicate sends.

## "Outlook is open but the integration says it isn't ready"

The user is probably running the **new** Outlook (`olk.exe`), which doesn't expose COM. Only **classic** Outlook (`OUTLOOK.EXE`, the one bundled with Microsoft 365) is supported.

**Tell the user**: in the new Outlook, top-right, toggle "New Outlook" off. Then reopen and the integration will reconnect.

## DASL filter syntax errors when using `scope='dasl'`

`scope='dasl'` passes the entire `query` string into Outlook's `Items.Restrict` as a raw `@SQL=` filter — it must be valid DASL.

**Examples that work**:

```
@SQL="urn:schemas:httpmail:subject" LIKE '%budget%'
@SQL="urn:schemas:httpmail:hasattachment" = 1
@SQL="urn:schemas:httpmail:read" = 0 AND "urn:schemas:httpmail:importance" = 2
```

If you don't actually need DASL, use `scope='subject_body'` / `'subject'` / `'from'` instead — they construct the filter for you and handle escaping.

## Unread count looks wrong

`unread` on a mail item reflects whether the item is unread. `unread_count` on a folder (from `list_folders`) is Outlook's cached count and can lag the actual state by seconds, especially right after a sync.

If the numbers disagree, trust the per-item `unread` flag.

## Multi-mailbox: list_mails on "inbox" hit the wrong mailbox

`folder="inbox"` always resolves to the **default** mailbox's Inbox. If the user has multiple stores (a primary + a shared mailbox + an archive), the shortcut won't reach the others.

**Fix**: use a store-qualified path. `outlook_list_folders(root=null)` shows the top-level store names; pass `"<Store Display Name>/Inbox"`. Confirm the active mailbox with `outlook_whoami` if you're not sure which is which.

## Send error mentions "ResolveAll" or "recipient could not be resolved"

The address you passed to `attendees=` (or in `to`/`cc`/`bcc`) couldn't be resolved against the user's address book or directory. Possible causes:

- A typo in the address.
- An internal alias that the user's directory recognizes (`alice`) but Outlook needs the full SMTP form for (`alice@example.com`).
- An external address being rejected because the user's tenant blocks external invites.

**Recover**: ask the user for the full SMTP address; retry.
