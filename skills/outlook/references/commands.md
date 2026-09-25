# outlook command reference

The full syntax of every command. `SKILL.md` holds the rules and a one-line index; read this
file when you need the exact arguments or the behaviour behind a command.

`$M` below is `scripts/outlook-mail.sh`, `$C` is `scripts/outlook-calendar.sh`, `$T` is
`scripts/outlook-token.sh` and `$S` is `scripts/outlook-setup.sh`, all in this skill's
directory: use the full paths SKILL.md gives. Every command takes `--account <name>`
(or `-a <name>`) before the verb; otherwise `OUTLOOK_ACCOUNT`, then `default`.

**IDs.** Listings print a short ID (the last 20 characters of the Graph ID). Every command that
takes an ID accepts the short or the full form. A message keeps its ID when it moves to another
folder, Deleted Items included, so an ID from a listing still works after `move`, `batch-move`
or `delete`.

**Folder names.** `move`, `batch-move`, `folder`, `mkdir`, `rename` and `rmdir` resolve a name
the same way: a bare name matches case-insensitively anywhere in the tree (the shallowest wins
on a tie), and `Parent/Child` targets one folder. Deleted Items is only searched through a path.

## Reading mail

```bash
$M inbox [count]            # newest first, default 10
$M unread [count]
$M focused [count]          # the Focused inbox
$M sent [count]             # your sent items
$M folder "Projects" [count]
$M from "john@example.com" [count]
$M flagged [count]          # flagged for follow-up
$M category "Follow up" [count]   # messages with this category, any folder
$M thread <message-id>      # the whole conversation, oldest first
$M read <message-id>        # the full message: headers, recipients, body, attachment list
$M preview <message-id>     # subject, sender, date and a snippet only
$M drafts [count]
$M stats                    # inbox totals
```

`search` takes free text or KQL, and a count (default 10, up to 1000, or `all`). Results are
ranked by Graph, then sorted newest first:

```bash
$M search "project update"
$M search "invoice" 50
$M search 'subject:invoice AND from:jane@example.com'
$M search 'from:acme.com AND body:renewal' all
```

`preview` is a snippet of about 200 characters. It is for finding the right message, never for
reading one.

## Drafting and sending

Every command below makes a **draft**. Only `send` sends.

```bash
$M draft "to@example.com" "Subject" "Plain text body"
$M mddraft "to@example.com" "Subject" "**Markdown** body"     # converted to HTML
$M mddraft "a@example.com; b@example.com" "Subject" "Body" --cc "c@example.com" --bcc "audit@example.com"
$M reply <message-id> "Plain text reply"                       # reply-all
$M mdreply <message-id> "**Markdown** reply"                   # reply-all
$M forward <message-id> "a@x.com, b@y.com" ["optional markdown comment"]
$M followup <sent-message-id> ["optional markdown body"]       # a chaser on your own sent mail
$M send <draft-id>
```

- Recipient lists are comma- or semicolon-separated. Each draft command prints To, Cc and Bcc
  as Graph stored them.
- `reply`, `mdreply` and `followup` use Graph's `createReplyAll`, so every original To and Cc
  recipient is on the draft. For a reply to the sender only, trim the draft with `update to`
  and `update cc ""`.
- Prefer `mdreply` and `mddraft` for anything professional: plain text looks poor in Outlook.
- `mddraft`, `mdreply`, `forward` with a comment, `followup` and `update mdbody` need `pandoc`.
  The HTML they build sets the font, size, line height and paragraph spacing inline; that is
  all in `md_to_html` in `outlook-mail.sh`.
- `send` reads the draft back and prints From, To, Cc, Bcc, Subject and the attachments before
  it posts. It posts nothing if it cannot read them.

### Changing a draft

```bash
$M update <draft-id> subject "New subject"
$M update <draft-id> body "Plain text body"      # replaces the whole body, quoted chain and signature included
$M update <draft-id> mdbody "**Markdown** body"  # keeps the quoted chain and any signature block
$M update <draft-id> to "new@example.com"        # replaces To
$M update <draft-id> cc "one@example.com, two@example.com"   # adds to Cc, deduplicated
$M update <draft-id> bcc "bcc@example.com"       # adds to Bcc
$M update <draft-id> cc ""                       # clears Cc (bcc "" clears Bcc)
$M update <draft-id> importance high             # high, normal or low
$M update <draft-id> from "alias@example.com"
```

`mdreply` and `followup` put an invisible marker where the quoted chain starts, and
`update mdbody` splits on it, so editing a reply keeps the chain. Use `mdbody`, not `body`, on a
reply.

### Sending as an alias

```bash
$M aliases                                        # the addresses this mailbox can send as
$M update <draft-id> from "alias@example.com"     # works on any draft, replies included
OUTLOOK_FROM_ADDRESS="alias@example.com" $M mdreply <message-id> "Thanks"
```

- `OUTLOOK_FROM_ADDRESS` sets the From on every draft that `draft`, `mddraft`, `reply`,
  `mdreply`, `followup` and `forward` make. Each prints the From it set. `update from`
  overrides it on one draft, and `send` warns if the draft would go from another address.
- Get the exact alias from `aliases`; never guess one. An address not in the list warns rather
  than blocks, because SendAs rights on a shared mailbox never appear there. If it is not
  permitted, `send` fails with `ErrorSendAsDenied` and nothing goes.
- The tenant must have `SendFromAliasEnabled`. Without it Exchange rewrites the From back to the
  primary address, so check that a test send arrived as the alias.
- `OUTLOOK_FROM_NAME` is usually ignored: Exchange uses the mailbox's own display name.
- An alias on a domain with no DKIM or DMARC may be spam-filtered by strict receivers.

## Attachments, inline images and signatures

```bash
$M attachments <message-id>                    # list
$M download <message-id> [attachment-id]       # to ./inbox/ (under CLAUDE_PROJECT_DIR if set)
$M attach <draft-id> /path/to/file             # up to 150 MB; run again for more files
$M attach <draft-id> logo.png --inline logo    # shown by <img src="cid:logo"> in the body
$M signature <draft-id> /path/to/signature.html
```

- Files under 3 MB go up in one request; larger ones in 4 MB chunks.
- `signature` works on an HTML draft (`mddraft` and `mdreply` always make one; a plain `draft`
  is refused). Every `<img>` whose src is a local file is uploaded inline and its src rewritten
  to `cid:`, so it shows without the reader allowing remote images. Relative paths are read
  from the HTML file's directory. The block goes above the quoted chain. `update mdbody` keeps
  it, running `signature` again replaces it, and `update body` removes it.

## Organising mail

```bash
$M markread <message-id>
$M markunread <message-id>
$M flag <message-id>
$M unflag <message-id>
$M junk <message-id>                 # to Junk Email
$M notjunk <message-id>              # back to the Inbox
$M archive <message-id>
$M delete <message-id>               # to Deleted Items, where it can be restored
$M move <message-id> "Clients/Acme"
$M batch-move "Projects" <id1> <id2> <id3>        # `bulk-move` is the same command
some_command_that_prints_ids | $M batch-move "Projects"
```

`batch-move` resolves the folder once and moves 20 messages per Graph `$batch` request. It
prints `FAILED <id>` for every message it did not move, including every message in a batch
that failed as a whole, and exits 1. Messages Graph throttled are retried on their own. To
reorganise a whole inbox, group IDs by destination and run `batch-move` once per folder.

### Categories

```bash
$M categories                                    # the master list
$M categorize <message-id> "Red category, Invoices"   # replaces the message's categories
$M categorize <message-id> ""                    # clears them
$M categorize <message-id> --add "Follow up"     # adds one, keeps the rest
$M categorize <message-id> --remove "Follow up"  # removes one, keeps the rest
$M mkcategory "Follow up" red                    # a colour name or presetN; safe to re-run
$M rccategory "Follow up" "dark blue"            # recolour
$M rmcategory "Follow up"                        # from the master list only
```

Prefer `--add` and `--remove` unless you mean to replace the whole list. There is no rename:
Graph fixes a category's name once it exists. `rmcategory` leaves the label on messages that
already carry it.

### Folders

```bash
$M folders                        # top level
$M subfolders ["Important"]       # default: Inbox
$M mkdir "Projects"               # top level
$M mkdir "Acme" "Clients"         # under a folder
$M rename "Old Name" "New Name"   # refuses system folders
$M rmdir "Empty Folder"           # to Deleted Items
$M rmdir "Old Folder" --force     # also when it holds messages; never a system folder
```

## Exporting mail to an archive

```bash
$M export "Inbox/Clients" ./staging/ [--since 2026-07-01] [--count 50]
```

Then append `./staging/` to the archive with the outlook-to-md skill, using its `--append`
mode; its own instructions give the command. If that skill is not installed, say so and stop
at the export.

`export` writes each message as raw `.eml`, newest first (default cap 1000). The staging
layout becomes the archive's folder grouping. `--append` dedupes on `Message-ID`, so an
overlapping `--since` is harmless; a message with no `Message-ID` (some drafts) is archived
again on every overlapping run.

## Calendar

Times are `YYYY-MM-DDTHH:MM` wall-clock in `OUTLOOK_TZ`, which defaults to the system zone.
Every listing prints each event's short ID. A recurring meeting appears once per occurrence.

```bash
$C events [count]                 # upcoming, default 10
$C today
$C week
$C day 2026-07-20
$C search "board meeting" [days]  # subject and location, default the next 90 days
$C read <event-id>
$C calendars
$C free "2026-02-05T09:00" "2026-02-05T17:00"   # busy periods; free and cancelled events ignored
```

### Creating events and inviting people

```bash
$C create "Subject" "2026-02-05T14:00" "2026-02-05T15:00" ["Location"]   # no attendees: nothing is sent
$C quick "Team standup" "2026-02-05T09:00"                                  # one hour, nothing is sent
$C invite <event-id> "a@x.com, b@y.com" [optional]                          # SENDS invitations
$C create "Subject" "<start>" "<end>" "" "a@x.com" --send-invites           # SENDS on creation
```

- The two-step flow is the default: `create` with no attendees, show the user the event and the
  attendee list, and run `invite` only after they approve. Re-inviting an address does nothing,
  so `invite` can be run again to add people.
- `create` refuses an attendee list without `--send-invites`, and creates nothing. Use the flag
  only when the user has approved that exact list in this conversation.

### Responding, changing, cancelling

```bash
$C respond <event-id> accept|decline|tentative ["message"]   # SENDS to the organiser
$C update <event-id> <field> <value>                          # subject, location, start or end
$C update <event-id> start "2026-02-05T15:00" --notify-attendees   # SENDS updates to attendees
$C cancel <event-id> ["message"]                              # a meeting you organise; SENDS cancellations
$C delete <event-id>                                          # removes it; notifies nobody
```

- `update` on a meeting you organise with attendees refuses and names who would hear. Show the
  user the change and that list, then run it again with `--notify-attendees` once they approve.
- `delete` refuses a meeting you organise, because Graph would send a cancellation. Use
  `cancel` for that.

## Tokens and accounts

```bash
$T list      # configured accounts
$T status    # connected, and as whom
$T test      # inbox totals, refreshing first if needed
$T refresh   # force a refresh
$T get       # a valid access token, refreshed if stale
$S [--account work]                 # the user runs this, not the agent
```

Credentials live in `~/.dbhq/outlook/<account>/`. A token refreshes itself when a command needs
it. A refresh that fails leaves `credentials.json` unchanged; one that Microsoft refuses (the
sign-in has lapsed) needs the user to run setup again. Setup opens a browser and waits for the
user to paste back a URL, so ask them to run it.

## Errors and limits

- **Throttled (HTTP 429).** Retried after Graph's `Retry-After`, up to three times
  (`OUTLOOK_MAX_RETRIES`). A 503 or 504 is retried for reads only. Do not loop a command
  yourself.
- **Timeouts.** Every request gives up after 120 seconds (600 for an upload chunk, a download
  or an export).
- **Permission denied.** Ask the user to run setup again to re-consent.
- **Read-only mode.** With `OUTLOOK_READ_ONLY=1` every command that writes or sends refuses
  before any request. Tell the user; do not unset it.
