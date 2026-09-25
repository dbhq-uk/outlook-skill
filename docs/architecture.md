# How the pack is built

Why there are two skills, what the safety rails actually are, and the decisions that are
easier to live with once you know why they were made.

## Two skills, one problem

Mail arrives in two states: the correspondence you are handling now, and the correspondence
someone handed you in a box. `outlook` deals with the first, over the network, with
credentials. `outlook-to-md` deals with the second, entirely offline, with no credentials at
all - it reads files.

They are separate skills because their dependencies are disjoint and their risk profiles are
nothing alike. A machine with no `azure-cli` can still turn a PST into markdown; `install.sh`
checks requirements per skill and skips rather than fails, so you can take either half on its
own.

They join at one point: `.eml` files. `outlook export` writes them, `outlook-to-md`
reads a directory of them, and the result is one archive that spans the PST you were given and
the mail that has arrived since.

## No server, no daemon, no service

Four bash scripts and the small library they share, `curl`, `jq`, and the Microsoft Graph v1.0 REST API. Nothing runs between
commands, nothing listens on a port, and there is no component of this pack you have to trust
that is not a file in the repository.

Authentication is a **delegated** OAuth flow, not application permissions. The distinction
matters: a delegated token can only ever do what the signed-in person can do in their own
mailbox. There is no tenant-wide grant, and nothing here can reach another person's mail.
Delegated does not always mean no admin: whether a user may consent to `Mail.Send` and
`Mail.ReadWrite` is the tenant's user-consent policy, and many tenants require an admin to
approve an app that asks for them. The app is a **public client**: setup signs in with the
authorisation-code flow and PKCE, so there is no client secret to store, and none to expire. It
used to register a Web app with a two-year secret shared by every account, which would have
stopped them all on the same day. Tokens live under `~/.dbhq/outlook/<account>/` at mode `600` and refresh
themselves when a command needs it. A refresh that fails leaves the stored tokens exactly as
they were, so a network blip never costs you the sign-in.

Five scopes are requested and no more: `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`,
`User.Read`, `offline_access`.

### The scope that was deliberately left out

`MailboxSettings.Read` would let the calendar script read the mailbox's own timezone, which
would remove the single most annoying failure in the pack. It is not requested, because
reading a person's mailbox settings is a real permission and a calendar convenience is not a
good enough reason to hold it.

So the script resolves the timezone from the machine instead - `OUTLOOK_TZ`, `/etc/timezone`,
`timedatectl`, the `/etc/localtime` symlink, `Europe/London` - and when it lands on UTC
without being told to, it warns on stderr before every command. Being loudly approximate beats
being quietly wrong: same instant, wrong wall-clock, missed meeting.

## The send gap

Every outward action in the pack is split in two, and the second half is a separate verb you
have to run on purpose.

| Creating | Sending |
|---|---|
| `draft`, `mddraft`, `reply`, `mdreply`, `forward`, `followup` | `send` |
| `create` (an event with no attendees) | `invite` |

This is not a UI nicety. An agent driving a mailbox is one confident inference away from
mailing a client, and the gap is where a person gets to look.

A split in two is only a gap if something stops the second half running straight after the
first. Instructions in `SKILL.md` ask the agent to wait for a yes, but an instruction is only as
strong as the agent's reading of it. So the gap is held in four layers, and each one says what
it does not cover.

**The scripts refuse the one-command sends.** `create` with an attendee list sends the
invitations the moment the event exists, so it refuses unless `--send-invites` is on the
command. Changing a meeting you organise sends every attendee an update, so `update` refuses
one unless `--notify-attendees` is given. Deleting a meeting you organise sends the attendees a
cancellation, which is Graph's documented behaviour and not the silent delete it looks like,
so `delete` refuses one and points at `cancel`. None of this stops `send`, `invite`, `respond`
or `cancel`, which are sends by name.

**`OUTLOOK_READ_ONLY=1` refuses every command that writes or sends**, in both scripts, before
a token is read or a request is made. It is an allow-list: each script names the verbs that
only read, and every other verb is refused, so a verb added later is refused until somebody
decides it only reads. It suits a triage session. It is the only layer that works the same
under every agent.

**The plugin asks before a command that sends.** Installed as a Claude Code plugin, the pack
registers a `PreToolUse` hook, `hooks/send-gate.sh`, that answers "ask" for `send`, `invite`,
`respond`, `cancel`, `--send-invites` and `--notify-attendees`, and stays silent for everything
else. It reads the command as written, so a script reached through a variable or a renamed copy
is not caught. Claude Code documents that a hook's "ask" prompts in Manual and auto mode; it
does not document that one prompts in bypassPermissions mode.

**Ask rules prompt in every mode.** Claude Code lists an explicit ask rule among the things no
permission mode approves on its own, bypassPermissions included. `hooks/ask-rules.json` holds
rules for the same commands, and `install.sh` offers to add them to `~/.claude/settings.json`.
They go in your own settings, so they are never added without a yes: `install.sh --ask-rules`,
or `y` at the prompt. `hooks/install-ask-rules.sh` adds them on its own, and only the ones that
are missing. The symlink install does not load the plugin's hook, so for that install these
rules are the Claude Code layer.

Codex loads neither the hook nor the rules. Under Codex the gap is the scripts, read-only mode
and the instructions.

## Replies keep everyone on the thread

`reply`, `mdreply` and `followup` call Graph's `createReplyAll`, so the draft carries every
original `To:` and `Cc:`. Trimming to sender-only is possible and takes an extra command.

The default is that way round because of an actual incident: a reply silently dropped two Cc'd
assistants from an estate-agent thread, and they had to be looped back in afterwards. The
failure mode is invisible from your side - the recipients simply stop seeing the conversation,
and nobody tells you. Losing a recipient by accident is a real harm; sending to one person too
many is an embarrassment. The default protects against the first.

The same reasoning shapes `preview` versus `read`. `preview` returns about 200 characters and
is documented for navigation only, because a short preview does not mean a short message, and
the requests and deadlines live below the fold.

## Short IDs

A Microsoft Graph message ID runs to several hundred characters. Pasting one into a terminal
is unpleasant; putting a screenful of them into an agent's context is worse.

Listings therefore print the last 20 characters, and every command that takes an ID accepts
either form. Every listing writes the short-to-full mapping into `id_cache.json`, so the
common list-then-act flow resolves from cache with no extra API call. On a miss the resolver
cascades through the folders a message might be in.

The cost is a rule you have to know: **moving a message gives it a new ID**. That comes from
a choice made here. The skill uses Graph's default IDs, which change when a message changes
folder. Graph can return immutable IDs instead, through the `Prefer: IdType="ImmutableId"`
header, and those survive a move; the skill does not ask for them today. The short-ID cache
makes the rule easy to forget. Re-list from the destination before acting on a message you
have just moved.

## Email HTML that survives Outlook

Every markdown-to-HTML conversion in the pack goes through one helper, `md_to_html` in
`outlook-mail.sh`, and applies its styling **inline** on each element:

| Property | Value |
|---|---|
| Font | `'Aptos', 'Aptos Display', 'Segoe UI', Roboto, sans-serif` |
| Size | `14px` |
| Line height | `1.5`, or `1.6` on replies and forwards |
| Colour | `#333` |
| Paragraph margin | `0 0 14px 0`, on every `<p>` |

Inline rather than a `<style>` block because Outlook strips the block. The per-paragraph
margin is there for the same reason: without it, paragraphs collapse into each other until
Outlook happens to re-render the draft after an edit, which looks like the tool mangled your
email. Aptos is the Microsoft 365 default since 2024, and the stack degrades to Segoe UI and
then to a system sans elsewhere.

One helper, one font stack variable. Change it there and every command changes with it.

## Two bash constraints worth knowing

**GNU and BSD both.** The scripts run on Linux and WSL (GNU `date`, `dd`, `sed`) and on macOS
(BSD tools, and `/bin/bash` 3.2). Dates try GNU `date -d` first and fall back to BSD
`date -j -f` or `date -v+Nd`; upload chunks are read with `tail -c` and `head -c`, which both
have. A `macos` CI job runs the offline suites on a Mac, and `helpers_test.sh` runs the date
helpers against a stand-in for BSD `date` on Linux too.

**Attachments never touch the command line.** Base64 goes to a temp file and reaches `jq` and
`curl` through `--rawfile`, because Linux's `MAX_ARG_STRLEN` is about 128 KB and passing the
payload as an argument fails with "Argument list too long" for anything over roughly 96 KB.
Above 3 MB the script opens a Graph upload session and streams 4 MB chunks instead, which is
what makes attachments up to Graph's 150 MB ceiling possible at all.

**Bulk moves go through `$batch`.** `batch-move` resolves the destination folder once and
sends moves twenty at a time through Graph's batch endpoint, reporting per-message failures
and exiting non-zero if any of them failed. Looping `move` would be one round trip per message
plus one folder lookup per message. A batch answer with no `responses` array (an error, a
gateway page, nothing at all) fails every message in that batch by ID: it once printed
"0 moved, 0 failed" and exited 0. Messages Graph throttled inside a batch are sent again, and
only those, after the longest `Retry-After` among them.

**Every request has a timeout and survives throttling.** All Graph calls go through
`outlook_curl` in `lib/graph.sh`, which adds `--connect-timeout` and `--max-time` and reads the
status from the response headers. Outlook allows four concurrent requests per mailbox and
answers the rest with 429 and `Retry-After`. A 429 is retried on any method, because Graph did
nothing with the request. A 503 or 504 is retried only on `GET`, because a POST that timed out
at a gateway may still have sent the mail. An HTTP error with no JSON body becomes a JSON error,
so no caller mistakes an empty 503 for success. `graph_test.sh` also checks that every `curl`
written in the scripts carries both timeouts.

## The archive is built to be stood behind

`outlook-to-md` produces a chain of custody, not just a pile of markdown:

1. Every email folder gets a `checksums.sha256` over its own files
2. `manifest.sha256` hashes every one of those, plus the index
3. The manifest records each source the archive was built or appended from, with that source's
   own SHA256 and a timestamp

So `sha256sum -c manifest.sha256` is a single command that says whether the archive is still
exactly what came out of the sources, and the manifest says what those sources were. The
target here is the archive someone asks you to produce a year later.

The original `.eml` is kept beside every `email.md` for the same reason. The markdown is a
convenience; the RFC 822 original is the evidence.

Deduplication on append is by `Message-ID`, and there is no content-hash fallback. Received
mail always carries the header, but some drafts and malformed messages do not, and those are
re-archived on every overlapping run. It is a narrow gap and it is documented rather than
papered over.

## One PST reader

A `.pst` is read by `readpst` from `pst-utils`, and by nothing else. It writes each message as
an `.eml` file in a folder tree that mirrors the PST (`readpst -j 0 -e -8`, plus `-D` when
`--include-deleted` is given), and that tree goes through the same `.eml` path as a live-mail
export. So there is one parser for messages, and it is the one the tests cover most.

`-j 0` turns off readpst's parallel jobs. With `-e`, readpst may split one folder's messages
across jobs, and then it sometimes never writes the last few of them. It exits 0 and reports
no error, so the archive is short and says it is complete. On the CI sample PST, readpst 0.6.76
wrote 70 or 71 messages by default and 71 every time with `-j 0`.

There used to be a second, Python PST backend, preferred whenever it was installed. It failed on
every message, its last release was in 2022, and it was the only reason for a Python 3.11
ceiling and a pinned `setuptools`. It was removed on 25 Sep 2026. The Python side now installs on
any interpreter from 3.9 up.

A directory of `.eml` files needs no `readpst` and is checked first. CI runs the suite on 3.9,
3.11 and 3.13 with `readpst` stubbed, and a separate job installs `pst-utils` and converts a
real, public PST end to end on 3.13, checking the emails, folders and attachments that come out.

## Testing

The bash suites are offline by construction: they extract the real functions out of the
scripts and exercise them against a mocked `api_call`, so no account, credentials or network
are involved. Suites are discovered rather than listed, so a new `*_test.sh` runs without
anyone remembering to register it.

CI installs `pandoc` before running them and **fails the build if any suite skips a test**. The
`md_to_html` assertions skip themselves when `pandoc` is missing, and that helper builds the
body of every email the skill sends - a skipped test inside a green build is exactly the thing
worth failing on.
