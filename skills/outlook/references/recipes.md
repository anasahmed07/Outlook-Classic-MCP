# Recipes

Worked examples for common multi-step Outlook workflows. Each recipe shows the *order of tool calls*, what to extract from each return, and where the user-confirmation moment goes.

## Contents

- [1. Triage unread mail](#1-triage-unread-mail)
- [2. Draft a grounded reply](#2-draft-a-grounded-reply)
- [3. Reply-all with an attachment](#3-reply-all-with-an-attachment)
- [4. Save attachments from a message](#4-save-attachments-from-a-message)
- [5. Find messages from someone](#5-find-messages-from-someone)
- [6. Move a thread to a project folder](#6-move-a-thread-to-a-project-folder)
- [7. File a project folder you don't have yet](#7-file-a-project-folder-you-dont-have-yet)
- [8. Weekly inbox digest](#8-weekly-inbox-digest)
- [9. What's on my calendar this week?](#9-whats-on-my-calendar-this-week)
- [10. Schedule a one-off meeting](#10-schedule-a-one-off-meeting)
- [11. Schedule a recurring meeting](#11-schedule-a-recurring-meeting)
- [12. Reschedule an existing event](#12-reschedule-an-existing-event)
- [13. Accept or decline a meeting invite](#13-accept-or-decline-a-meeting-invite)
- [14. Look up a contact](#14-look-up-a-contact)
- [15. Create a follow-up task from an email](#15-create-a-follow-up-task-from-an-email)
- [16. Tag mail with a color category](#16-tag-mail-with-a-color-category)
- [17. Toggle a mail rule](#17-toggle-a-mail-rule)
- [18. Confirm OOO status](#18-confirm-ooo-status)
- [19. Multi-mailbox: act on a shared mailbox](#19-multi-mailbox-act-on-a-shared-mailbox)

---

## 1. Triage unread mail

> "What's unread in my inbox today?"

```
1. outlook_list_mails(folder="inbox", unread_only=true, limit=25, response_format="json")
2. Summarize the items for the user (sender, subject, time, preview).
3. If the user asks to act on one, capture its entry_id and route to recipe 2/4/6.
```

Use `since` if the user wants only today: `since="<today>T00:00:00"`.

## 2. Draft a grounded reply

> "Reply to that email from Bob about the Q3 numbers — tell him we'll have it by Friday."

```
1. outlook_search_mails(query="Q3", folder="inbox", scope="subject_body",
                       response_format="json")
2. Pick the right item by subject/sender. Note its entry_id.
3. outlook_get_mail(entry_id=<id>, include_body=true, response_format="json")
   → read the full body so the reply is on-topic.
4. Draft the reply text yourself, grounded in step 3's body.
5. Show the draft to the user and ask "send or save to drafts?".
6. Either:
     a) outlook_reply_mail(entry_id=<id>, body=<draft>)         # sends
     b) Recreate as a draft via outlook_send_mail(
          to=[from_address from step 3], subject="Re: <subject>",
          body=<draft>, save_only=true)                         # stages
```

Why step 3: list/search results carry only a 200-character preview. Replying off the preview produces vague, generic responses.

Why "send or save": `reply_mail` sends immediately. There is no `save_only` flag on `reply_mail`. If the user wants to review before sending, you must build it as a draft via `send_mail(save_only=true)`.

## 3. Reply-all with an attachment

```
1. outlook_search_mails(...) → entry_id
2. outlook_get_mail(entry_id, include_body=true) → confirm context
3. Confirm with the user.
4. outlook_reply_mail(
      entry_id=<id>,
      body=<reply text>,
      reply_all=true,
      attachments=["C:\\Users\\<user>\\Documents\\report.pdf"])
```

Attachment paths must be absolute and under the user's profile.

## 4. Save attachments from a message

> "Save the PDF from Sarah's email to my Downloads."

```
1. outlook_search_mails(query="<keyword>", scope="from", response_format="json")
   or list/search until you have the entry_id.
2. outlook_get_mail(entry_id, include_body=false, response_format="json")
   → check the `attachments` array; note indexes.
3. outlook_save_attachments(
      entry_id=<id>,
      output_dir="C:\\Users\\<user>\\Downloads",
      attachment_index=2)         # omit to save all
```

Returns absolute paths to the saved files — surface them to the user.

## 5. Find messages from someone

> "Did Mira email me about the offsite?"

Two options. The first is faster but only filters; the second searches text content.

```
A) outlook_list_mails(folder="inbox", from_address="mira", limit=10,
                      response_format="json")
   # substring match on sender — works even when the address is EX:/O=...

B) outlook_search_mails(query="mira offsite", folder="inbox",
                        scope="subject_body", response_format="json")
   # actually scans subject + body
```

If the user says the magic word "from" they often want option A. Combine: search for "offsite", then filter the results by checking `from` in code.

## 6. Move a thread to a project folder

```
1. outlook_search_mails(...) → entry_id
2. outlook_move_mail(entry_id=<id>, target_folder="Inbox/Projects/Acme")
3. Capture new_entry_id from the response if you'll act on the moved item again.
```

The `entry_id` you started with is no longer valid after the move; use `new_entry_id`.

## 7. File a project folder you don't have yet

```
1. outlook_list_folders(root="inbox", max_depth=4)  → confirm path doesn't exist
2. outlook_create_folder(name="Acme", parent="Inbox/Projects")
3. Now move mails into it (recipe 6).
```

`create_folder` will fail if `parent` doesn't exist; create parents one level at a time if needed.

## 8. Weekly inbox digest

> "What did I get this past week from anyone outside the team?"

```
1. outlook_list_mails(
      folder="inbox",
      since="<today - 7d>T00:00:00",
      until="<today>T23:59:59",
      limit=100,
      response_format="json")
2. In your post-processing, drop items whose from_address contains
   the user's company domain.
3. Group by from / by day / by importance, summarize.
```

Don't try to express "outside the team" as a DASL filter — it's much easier to filter in your own analysis.

## 9. What's on my calendar this week?

```
1. outlook_list_events(
      start="<this Monday>T00:00:00",
      end="<this Sunday>T23:59:59",
      include_recurrences=true,
      response_format="markdown")
```

Markdown is fine here — the result is the answer. Use `json` if the user is going to ask follow-ups like "decline the 4pm".

## 10. Schedule a one-off meeting

> "30 minutes with alice@example.com on Tuesday at 2pm to discuss the proposal."

```
1. Confirm: "Tuesday May 5 at 2:00–2:30 PM, attendees: alice@example.com,
   subject: 'Discuss proposal'. Send the invite?"
2. outlook_create_event(
      subject="Discuss proposal",
      start="2026-05-05T14:00:00",
      end="2026-05-05T14:30:00",
      attendees=["alice@example.com"],
      reminder_minutes=15)
```

Adding `attendees` makes this a meeting and **sends the invite immediately on success**. There's no draft state — confirm before calling.

## 11. Schedule a recurring meeting

> "Weekly 1:1 with my manager Tuesdays at 10am for the rest of the year."

```
outlook_create_event(
   subject="1:1 with <manager>",
   start="2026-05-05T10:00:00",
   end="2026-05-05T10:30:00",
   attendees=["manager@example.com"],
   recurrence={
      "type": "weekly",
      "interval": 1,
      "end_date": "2026-12-31"
   })
```

If the user prefers a count, use `occurrences` instead of `end_date`. Omit both for indefinite.

## 12. Reschedule an existing event

```
1. outlook_list_events(...) or outlook_get_event(entry_id) → confirm correct event
2. outlook_update_event(
      entry_id=<id>,
      start="<new start>",
      end="<new end>")
```

`update_event` rewrites only the fields you pass. It cannot change recurrence cadence — for that, delete and recreate.

If the event has attendees, Outlook will likely send an updated-meeting notification when the change saves. Mention that to the user.

## 13. Accept or decline a meeting invite

```
1. outlook_search_mails(query="<organizer or subject>", scope="subject_body")
   or outlook_list_events / outlook_get_event to locate the meeting → entry_id
2. outlook_respond_event(entry_id=<id>, response="accept")
   # or "tentative" / "decline"
```

Pass `send_response=false` if the user wants to update locally without notifying the organizer.

## 14. Look up a contact

```
A) outlook_search_contacts(query="<name fragment>", response_format="json")
   → grab entry_id of the right hit
B) outlook_get_contact(entry_id=<id>) for full details
```

`search_contacts` matches across name, email, company, and job title — if the user only knows a company, that works.

## 15. Create a follow-up task from an email

> "Remind me to follow up on Bob's contract email next Monday."

```
1. outlook_search_mails(...) → confirm the email exists, capture subject
2. outlook_create_task(
      subject="Follow up: " + <email subject>,
      due_date="<next Monday>T09:00:00",
      reminder="<next Monday>T09:00:00",
      body="Re: Bob's email about <subject>; entry_id=<id>")
```

Stashing the original `entry_id` in the task body gives you (and the user) a way to jump back to the source email later.

## 16. Tag mail with a color category

```
1. outlook_list_categories()  → confirm the category name exists in the user's profile
2. outlook_get_mail(entry_id, include_body=false, response_format="json")
   → read existing categories so you don't clobber them
3. outlook_set_category(
      entry_id=<id>,
      categories="<existing>, Important")    # comma-separated; empty clears all
```

Categories must already exist in the user's Outlook profile. If the user wants a new category, they create it in Outlook itself first — there is no `create_category` tool.

## 17. Toggle a mail rule

> "Turn off the rule that auto-files Jira mail."

```
1. outlook_list_rules()  → find the exact name (e.g. "Auto-file Jira to Engineering/Jira")
2. Confirm with the user: "Disable rule 'Auto-file Jira to Engineering/Jira'?"
3. outlook_toggle_rule(rule_name="Auto-file Jira to Engineering/Jira",
                       enabled=false)
```

The change is live immediately. Always confirm the exact name; partial / fuzzy matches are not supported.

## 18. Confirm OOO status

```
outlook_get_out_of_office()
```

If `status: "on"`, surface that to the user. If they want to **change** OOO state, you cannot — direct them to Outlook → File → Automatic Replies.

## 19. Multi-mailbox: act on a shared mailbox

When the user has more than one Outlook account or store (e.g. a personal mailbox plus a shared "Support" mailbox), the well-known names (`inbox`, `sent`, etc.) only resolve against the **default** mailbox. To target another store, qualify the path with the store display name:

```
1. outlook_whoami()                           # confirm what's bound
2. outlook_list_folders(root=null, max_depth=2)
   # the top-level entries are store names; pick the store you want
3. outlook_list_mails(folder="Support Mailbox/Inbox", limit=25)
   # store name + path
```

The store display name is whatever Outlook shows in the folder pane (e.g. `"Mailbox - support@example.com"`, `"Online Archive - you@example.com"`, `"Personal Folders"`). Match it case-insensitively.
