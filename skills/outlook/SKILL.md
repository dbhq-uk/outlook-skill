---
name: outlook
description: Microsoft 365 Outlook mail and calendar through Microsoft Graph - read and search the inbox, draft, reply and send, attachments, categories and folders, calendar events, invitations and free/busy. Use for "check my email", "draft an email", "reply to", "my Outlook calendar", "am I free", "book a meeting", "Microsoft 365 mail". Not for Gmail or Google Calendar, and not for PST files (use outlook-to-md).
---

# outlook

Microsoft 365 Outlook mail and calendar through the Microsoft Graph API. Two scripts do the
work: `${CLAUDE_SKILL_DIR}/scripts/outlook-mail.sh` and
`${CLAUDE_SKILL_DIR}/scripts/outlook-calendar.sh`. The index below has every command. For the
exact arguments and behaviour, read `references/commands.md`.

## Rules

**1. Draft, show, then send.** Every mail command makes a draft; only `send` sends. Show the
user the draft (From, To, Cc, Bcc, subject, body, attachments) and wait for an explicit "send
it" before running `send`. The same goes for invitations: `create` an event with no attendees,
show the user the event and the attendee list, and run `invite` only after they approve.

**2. Replies go to everyone.** `reply`, `mdreply` and `followup` are reply-all: every original To
and Cc recipient is on the draft. Read the original's full To and Cc before replying, and check
the recipients the draft prints. People Cc assistants and colleagues; dropping them is a real
harm. For a reply to the sender only, trim the draft with `update to` and `update cc ""`.

**3. Read the whole message.** `preview` is a 200-character snippet, for finding a message only.
Before you summarise, answer or act on an email, `read` it end to end, including the attachment
list and any request in the body. In a long chain, find where the current message ends and the
quoted chain begins.

**4. Check the time.** Run `date` (and `date -u`) before anything that involves today, deadlines
or "yesterday"; a session can span days. Graph timestamps are UTC. Calendar times are wall-clock
in `OUTLOOK_TZ`, which defaults to the system zone, and a server is often UTC when the user is
not. If the calendar warns that it is using UTC and that is not the user's zone, set
`OUTLOOK_TZ` (for example `Europe/London`) before quoting a time.

**5. Refusals and permission prompts are the send gap working.** Never work around one, for
example by calling Graph with curl yourself or by rewording a command so it is not recognised.
- With `OUTLOOK_READ_ONLY=1`, every command that writes or sends refuses. Tell the user; do not
  unset it.
- `send`, `invite`, `respond`, `cancel`, `create ... --send-invites` and
  `update ... --notify-attendees` may raise a permission prompt. That prompt is the user's
  approval, so show them what will be sent first.
- `create` refuses attendees without `--send-invites`. `update` refuses to change a meeting you
  organise without `--notify-attendees`, and `delete` refuses one (use `cancel`).

**6. Confirm the From address.** To send as an alias, get the exact address from `aliases`, set
it with `update <draft-id> from <address>` (or `OUTLOOK_FROM_ADDRESS` for every draft), and
show the user the From line before sending.

**7. Setup is the user's.** If an account is not configured, or a refresh is refused because
the sign-in has lapsed, ask the user to run `${CLAUDE_SKILL_DIR}/scripts/outlook-setup.sh`
(add `--account <name>` for another mailbox). Do not run it yourself: it opens a browser and
waits for a person to paste back a URL.

## Command index

Mail is `outlook-mail.sh <verb>`, calendar is `outlook-calendar.sh <verb>`. Add
`--account <name>` before the verb for another mailbox. Listings print short IDs, which every
command accepts. A message keeps its ID when it moves to another folder.

**Read mail**
- `inbox`, `unread`, `focused`, `sent`, `drafts`, `flagged` `[count]`: list, newest first
- `folder <name> [count]`, `from <address> [count]`, `category <name> [count]`: list by folder, sender or category
- `search <text or KQL> [count|all]`: search; KQL such as `subject:x AND from:y` works
- `read <id>`: the whole message. `preview <id>`: a snippet, for navigation only
- `thread <id>`: the conversation, oldest first. `stats`: inbox totals

**Write mail** (all make drafts)
- `draft <to> <subject> <body>`, `mddraft <to> <subject> <markdown> [--cc x] [--bcc y]`: a new draft; prefer `mddraft`
- `reply <id> <body>`, `mdreply <id> <markdown>`: reply-all draft; prefer `mdreply`
- `forward <id> <to> [markdown comment]`, `followup <sent-id> [markdown]`: forward, or chase your own sent mail
- `update <draft-id> <field> <value>`: `subject`, `body`, `mdbody` (keeps the quoted chain), `to`, `cc`, `bcc`, `importance`, `from`
- `aliases`: the addresses this mailbox can send as
- `attach <draft-id> <file> [--inline <cid>]`: add a file, up to 150 MB
- `signature <draft-id> <html-file>`: add an HTML signature; its local images go inline
- `send <draft-id>`: sends; prints From, To, Cc, Bcc, subject and attachments first

**Attachments**
- `attachments <id>`: list. `download <id> [attachment-id]`: save to `./inbox/`

**Organise mail**
- `markread`, `markunread`, `flag`, `unflag`, `junk`, `notjunk`, `archive` `<id>`
- `delete <id>`: to Deleted Items, where it can be restored
- `move <id> <folder>`, `batch-move <folder> <ids...>` (or IDs on stdin): move; `Parent/Child` targets one folder
- `categories`, `categorize <id> "<a, b>" | --add <c> | --remove <c>`: apply categories
- `mkcategory <name> [colour]`, `rccategory <name> <colour>`, `rmcategory <name>`: the master list
- `folders`, `subfolders [folder]`, `mkdir <name> [parent]`, `rename <old> <new>`, `rmdir <name> [--force]`: folders
- `export <folder> <dir> [--since YYYY-MM-DD] [--count N]`: raw `.eml` for the outlook-to-md skill to archive

**Calendar**
- `events [count]`, `today`, `week`, `day <YYYY-MM-DD>`, `search <text> [days]`: list, with short IDs
- `read <event-id>`, `calendars`, `free <start> <end>`: details, calendars, busy periods
- `create <subject> <start> <end> [location]`, `quick <subject> <start>`: sends nothing
- `invite <event-id> <emails> [optional]`: sends invitations
- `respond <event-id> accept|decline|tentative [message]`: tells the organiser
- `update <event-id> <field> <value> [--notify-attendees]`: subject, location, start or end
- `cancel <event-id> [message]`: a meeting you organise; tells every attendee
- `delete <event-id>`: removes an event that notifies nobody

**Accounts** (`outlook-token.sh`)
- `list`, `status`, `test`, `refresh`, `get`: accounts, connection, a valid token

Times are `YYYY-MM-DDTHH:MM`. Markdown commands need `pandoc`. Throttling (HTTP 429) is
retried by the scripts, so do not loop a command. `batch-move` prints `FAILED <id>` for each
message it did not move; run it again with those IDs.
