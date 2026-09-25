# Reference

Every command, argument and environment variable in both skills. For how to use them
together, see the [guides](README.md#doing).

Script paths are written here as `mail.sh`, `calendar.sh`, `token.sh` and `setup.sh`. In full
they are `~/.claude/skills/outlook/scripts/outlook-<name>.sh` for a local install,
and `${CLAUDE_SKILL_DIR}/scripts/outlook-<name>.sh` inside the skill itself.

## Conventions

**Account selection.** Every `outlook` script takes `--account <name>` or `-a <name>`
*before* the command, falling back to `$OUTLOOK_ACCOUNT` and then `default`.

**Message IDs.** Listings print a 20-character short ID (the tail of the full Graph ID). Every
command taking an ID accepts either form. Short IDs resolve from `id_cache.json`, which every
listing populates, so list-then-act costs no extra API call. A moved message has a new ID.

**Folder names.** `move`, `batch-move`, `folder`, `mkdir`, `rename` and `rmdir` resolve names
identically: a bare name matches case-insensitively anywhere in the tree, shallowest wins on a
tie; `Parent/Child` targets one specific folder.

**Times.** Calendar commands take and print `YYYY-MM-DDTHH:MM` wall-clock in the resolved
timezone (see [`OUTLOOK_TZ`](#environment-variables)).

**Exit codes.** `0` on success, `1` on usage error, a missing account, an unresolvable ID or a
Graph error. `batch-move` also exits `1` when any single message failed or any short ID could
not be resolved, having moved the rest.

## mail.sh

### Reading

| Command | Notes |
|---|---|
| `inbox [count]` | Newest first, default 10 |
| `unread [count]` | Unread in the inbox |
| `focused [count]` | Focused inbox only, sorted client-side |
| `sent [count]` | Sent items |
| `drafts [count]` | Drafts |
| `flagged [count]` | Flagged for follow-up, across folders |
| `category <name> [count]` | Messages carrying a category, across folders, newest first. Says when there are more than `count` |
| `folder <name> [count]` | Any folder, resolved by name |
| `from <email> [count]` | One sender, newest first |
| `search <query> [count]` | Default 10, max 1000, or `all` |
| `thread <id> [count]` | Whole conversation, oldest first, across inbox and sent |
| `read <id>` | Full body, full `To:`/`Cc:`, attachment list |
| `preview <id>` | About 200 characters. Navigation only |
| `export <folder> <dir> [--since YYYY-MM-DD] [--count N]` | Write messages as `.eml` for archiving. `--count` defaults to 1000, newest first |

`search` takes free text or Graph KQL: `subject:`, `from:`, `to:`, `body:`, with `AND`, `OR`
and `NOT`. A bare email address is turned into a `from:` match. Results come back ranked by
Graph, then sorted newest-first.

```bash
mail.sh search 'subject:invoice AND from:jane@example.com'
mail.sh search 'from:acme.com AND body:renewal' all
```

### Writing and sending

| Command | Notes |
|---|---|
| `draft <to> <subject> <body> [--cc <emails>] [--bcc <emails>]` | Plain text |
| `mddraft <to> <subject> <markdown> [--cc <emails>] [--bcc <emails>]` | Markdown to HTML. Needs `pandoc` |
| `reply <id> <body>` | **Reply-all.** Plain text |
| `mdreply <id> <markdown>` | **Reply-all.** Markdown. Prefer this |
| `forward <id> <to-emails> [markdown-comment]` | Quoted message and its attachments |
| `followup <sent-id> [markdown-body]` | Chaser on your own sent message. **Reply-all** |
| `update <draft-id> <field> <value>` | See below |
| `send <draft-id>` | The only verb that sends. Prints From, To, Cc, Bcc, Subject and attachments first, read back from Graph |
| `aliases` | Addresses this mailbox may send as |

Recipient lists are comma- or semicolon-separated. Nothing is sent until `send`.

`update` fields:

| Field | Behaviour |
|---|---|
| `subject` | Replaces |
| `body` | Plain text. Replaces the whole body, quoted chain included |
| `mdbody` | Markdown. Preserves the quoted chain on `mdreply` / `followup` drafts |
| `to` | Replaces the To line |
| `cc` | Appends, deduped case-insensitively. `""` clears |
| `bcc` | Appends, deduped case-insensitively. `""` clears |
| `from` | Send-as address. Works on any draft, including replies |
| `importance` | `high`, `normal`, `low` |

The chain preservation in `mdbody` works off an invisible `<span
data-mdreply-chain-start="1">` marker injected when the reply draft is created. `body` has no
knowledge of it.

### Attachments

| Command | Notes |
|---|---|
| `attachments <id>` | List |
| `download <id> [attachment-id]` | All, or one. Saves to `inbox/` under `$CLAUDE_PROJECT_DIR`, else the current directory |
| `attach <draft-id> <file> [--inline <cid>]` | One file per call, repeat for more. `--inline` uploads it as an inline image the body shows with `<img src="cid:<cid>">` |
| `signature <draft-id> <html-file>` | Adds an HTML signature to an HTML draft, above the quoted chain. Each `<img>` whose quoted `src` is a local file is uploaded inline and pointed at by `cid:`. Running it again replaces the block; `update mdbody` keeps it |

Under 3 MB, `attach` sends a single base64 upload. At or above 3 MB it opens a Graph upload
session and streams 4 MB chunks with a progress indicator, which is what carries files up to
Graph's 150 MB ceiling.

### Triage and organising

| Command | Notes |
|---|---|
| `markread <id>` / `markunread <id>` | |
| `flag <id>` / `unflag <id>` | Follow-up flag. List with `flagged` |
| `categorize <id> <cats>` | Comma-separated. **Replaces** the whole list. `""` clears |
| `categorize <id> --add <cat>` | Adds one, leaves the rest |
| `categorize <id> --remove <cat>` | Removes one, leaves the rest |
| `categories` | The mailbox's master category list |
| `mkcategory <name> [colour]` | Colour is a name (`red`, `dark blue`, …) or `presetN`. Re-running on an existing name reports rather than errors |
| `rccategory <name> <colour>` | Recolour |
| `rmcategory <name>` | Removes from the master list only. Messages keep the label |
| `junk <id>` / `notjunk <id>` | To Junk Email, or back to the Inbox |
| `archive <id>` | To the Archive folder |
| `delete <id>` | Moves the message to Deleted Items, where it can be restored. It gets a new ID there |
| `move <id> <folder>` | |
| `batch-move <folder> <id…>` | IDs as arguments or on stdin. Batches of 20 via Graph `$batch` |

There is no `rename` for categories: Graph makes `displayName` immutable once a category
exists.

### Folders

| Command | Notes |
|---|---|
| `folders` | Top level |
| `subfolders [parent]` | Default `inbox` |
| `mkdir <name> [parent]` | Top-level, or a subfolder of `parent` |
| `rename <folder> <new-name>` | Refuses well-known system folders |
| `rmdir <folder> [--force]` | Moves the folder, and everything in it, to Deleted Items. Refuses system folders always, and a folder holding messages without `--force` |
| `stats` | Inbox totals and unread count |

## calendar.sh

| Command | Notes |
|---|---|
| `events [count]` | The next `count` events (default 10) in the coming year. Each occurrence of a recurring meeting is its own row |
| `today` | |
| `week` | Today and the next seven days |
| `day <YYYY-MM-DD>` | |
| `search <text> [days]` | Subject and location, default next 90 days |
| `read <id>` | Details, including attendees and responses |
| `calendars` | |
| `create <subject> <start> <end> [location] [attendees --send-invites]` | Without attendees, **nothing is sent**. Attendees need `--send-invites`, because the invitations go out immediately; without it `create` refuses and creates nothing |
| `invite <id> <emails> [required\|optional]` | Sends invitations. Re-inviting an address is a no-op |
| `quick <subject> <start>` | One hour, no location |
| `update <id> <field> <value> [--notify-attendees]` | `subject`, `location`, `start`, `end`. A meeting you organise sends every attendee an update, so it needs `--notify-attendees` |
| `respond <id> <accept\|decline\|tentative> [comment]` | Notifies the organiser |
| `cancel <id> [comment]` | Withdraws a meeting you organise and tells attendees |
| `delete <id>` | For an event that notifies nobody. A meeting you organise is refused, because deleting it sends a cancellation: use `cancel` |
| `free <start> <end>` | Free, or what is in the way |

Every listing prints each event's short ID (the last 20 characters) and caches the full IDs,
so `read`, `update`, `respond`, `cancel`, `delete` and `invite` take the ID straight off the
screen. Listings page through Graph's results rather than stopping at its default of 10. A
short ID that is not in the last listing is looked up from 30 days back to a year ahead, which
finds a single occurrence of a recurring meeting. Two events whose IDs end the same way are
printed with their full IDs, and the short form is refused rather than guessed.

`free` ignores events shown as free and events that have been cancelled.

Graph rejects a `start` later than the current `end` and an `end` earlier than the current
`start`, so moving an event to another day means updating the safe bound first.

## token.sh

| Command | Notes |
|---|---|
| `refresh` | Force a token refresh |
| `get` | Print a valid access token, for direct Graph calls. Refreshes first if it is expired or within 60 seconds of expiry |
| `test` | Full round trip: token (refreshed if needed), Graph call, inbox counts |
| `status` | Connected or expired, and whose mailbox |
| `list` | Configured accounts |

## setup.sh

```bash
setup.sh                      # configure the default account
setup.sh --account work       # add another, reusing the app registration if found
```

Registers (or reuses) an Azure app, runs the OAuth sign-in, and writes
`~/.dbhq/outlook/<account>/`. The manual equivalent is
[`references/setup.md`](../skills/outlook/references/setup.md).

Delegated permissions requested, and the whole of what the pack can do:
`Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `User.Read`, `offline_access`.

## outlook_to_md.py

```
outlook_to_md.py [-h] [--include-deleted] [--timezone TZ] [--verbose] [--append]
                 [--owner-email EMAIL] pst_file output_dir
```

| Argument | Notes |
|---|---|
| `pst_file` | A PST file, or a directory of `.eml` files |
| `output_dir` | Created if absent |
| `--append` | Skip emails already archived, matched on `Message-ID`. Without it the run overwrites |
| `--include-deleted` | Include deleted items from the PST (passes `-D` to `readpst`) |
| `--timezone TZ` | Render every date in this IANA zone, e.g. `Europe/London`. An unknown name is refused. Without it, each date keeps the offset the message was sent with |
| `--owner-email EMAIL` | Fixes `MAILER-DAEMON` senders in sent items |
| `--verbose`, `-v` | Per-email logging |

Run it with the skill's own interpreter: `~/.claude/skills/outlook-to-md/.venv/bin/python`.

### Backends

A directory input is handled directly and needs nothing else. A PST file is read with
**readpst** (`pst-utils`): `readpst -e -8 -o <tmp> <pst>`, with `-D` added by
`--include-deleted`. Without `readpst` a PST run stops and says how to install it.

### Output

```
output/
├── emails/<Folder>/<timestamp>_from-x_to-y_Subject/
│   ├── email.md              # YAML frontmatter, body as markdown, original headers
│   ├── email.eml             # RFC 822 original
│   ├── attachment_001_*      # extracted attachments
│   └── checksums.sha256      # SHA256 of every file in this folder
├── index.csv                 # date, sender, recipient, subject, pst_folder, attachment count
├── index.md                  # timeline by year and month
├── extraction_log.txt        # errors and statistics
└── manifest.sha256           # hashes every checksums.sha256 and the index
```

`manifest.sha256` also records every source the archive has been built or appended from, each
with its own SHA256 and timestamp. Verify with `sha256sum -c manifest.sha256`.

`pst_folder` in the index means "the folder this message came from", whether that was a PST
folder or a Graph one.

### Behaviour on bad input

| Situation | What happens |
|---|---|
| Corrupt email | Logged to `extraction_log.txt`, processing continues |
| Encoding problems | UTF-8, then latin-1, then raw bytes |
| Two emails at the same timestamp | Suffixed `-001`, `-002` |
| Path too long | Subject truncated, uniqueness preserved |
| No `Message-ID` header | Nothing to dedupe on, so re-archived on every overlapping `--append` |

### Throughput

Roughly 5,000 emails an hour without attachments, 2,000 with. A 300 MB PST of one to three
thousand messages takes five to fifteen minutes.

## Environment variables

| Variable | Effect |
|---|---|
| `OUTLOOK_ACCOUNT` | Account to use, unless `--account` is given. Default `default` |
| `OUTLOOK_TZ` | Timezone for every calendar time. Falls back to `/etc/timezone`, `timedatectl`, the `/etc/localtime` symlink, then `Europe/London` |
| `OUTLOOK_FROM_ADDRESS` | Default From on every draft: `draft`, `mddraft`, `reply`, `mdreply`, `followup`, `forward`. `update from` overrides it on one draft |
| `OUTLOOK_FROM_NAME` | Usually ignored - Exchange overrides the display name for addresses the mailbox owns |
| `OUTLOOK_READ_ONLY` | Set to `1` and every command that writes or sends refuses before any request is made. Only the listings and reads run. Empty, `0`, `false`, `no` and `off` leave it off |
| `CLAUDE_PROJECT_DIR` | Where `download` writes its `inbox/` directory. Falls back to the current directory |

## Files on disk

| Path | Contents |
|---|---|
| `~/.dbhq/outlook/<account>/config.json` | Client ID, secret, tenant, redirect URI, scopes (`600`) |
| `~/.dbhq/outlook/<account>/credentials.json` | Access and refresh tokens (`600`) |
| `~/.dbhq/outlook/<account>/id_cache.json` | Short ID to full Graph ID |
| `~/.dbhq/outlook/<account>/event_id_cache.json` | Full IDs of the events the last calendar listing printed |
| `~/.dbhq/outlook/<account>/.token.lock` | Empty. Held during a token refresh so two commands do not refresh at once |

A pre-multi-account install with flat `~/.dbhq/outlook/*.json` files is migrated into
`default/` on the first run of any script.

## Requirements

| Skill | Required | Optional |
|---|---|---|
| `outlook` | `azure-cli`, `jq`, `curl` | `pandoc`, for every markdown command |
| `outlook-to-md` | `python3` 3.9+ | `readpst` (`pst-utils`), needed for `.pst` files only |

`install.sh` checks these per skill, so a missing `azure-cli` skips `outlook` and leaves
`outlook-to-md` installed rather than failing the lot.
