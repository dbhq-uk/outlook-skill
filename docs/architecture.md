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

Four bash scripts, `curl`, `jq`, and the Microsoft Graph v1.0 REST API. Nothing runs between
commands, nothing listens on a port, and there is no component of this pack you have to trust
that is not a file in the repository.

Authentication is a **delegated** OAuth flow, not application permissions. The distinction
matters: a delegated token can only ever do what the signed-in person can do in their own
mailbox. There is no tenant-wide grant, no admin consent, and nothing here can reach another
person's mail. Tokens live under `~/.dbhq/outlook/<account>/` at mode `600` and refresh
themselves when a command needs it.

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

## Nothing leaves without a second command

Every outward action in the pack is split in two, and the second half is always a separate
verb you have to run on purpose.

| Creating | Sending |
|---|---|
| `draft`, `mddraft`, `reply`, `mdreply`, `forward`, `followup` | `send` |
| `create` (an event with no attendees) | `invite` |

`create` does have a form that takes attendees and invites them immediately. It exists
because sometimes the list really has already been agreed, and it is documented as the
exception rather than the shape of the workflow.

This is not a UI nicety. An agent driving a mailbox is one confident inference away from
mailing a client, and the gap is where a person gets to look.

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

The cost is a rule you have to know: **moving a message gives it a new ID**. That is Graph's
behaviour, not a choice made here, but the short-ID cache makes it easy to forget. Re-list
from the destination before acting on a message you have just moved.

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

**Attachments never touch the command line.** Base64 goes to a temp file and reaches `jq` and
`curl` through `--rawfile`, because Linux's `MAX_ARG_STRLEN` is about 128 KB and passing the
payload as an argument fails with "Argument list too long" for anything over roughly 96 KB.
Above 3 MB the script opens a Graph upload session and streams 4 MB chunks instead, which is
what makes attachments up to Graph's 150 MB ceiling possible at all.

**Bulk moves go through `$batch`.** `batch-move` resolves the destination folder once and
sends moves twenty at a time through Graph's batch endpoint, reporting per-message failures
and exiting non-zero if any of them failed. Looping `move` would be one round trip per message
plus one folder lookup per message.

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

## The Python version problem

`libratom` is the preferred PST backend. It pins `numpy==1.23.5`, whose newest wheel is cp311
- so it cannot install on Python 3.12 or later, which is what most current systems ship.

`setup.sh` handles this rather than failing: it prefers a 3.9 to 3.11 interpreter if one is on
your `PATH`, and otherwise builds the virtualenv without `libratom` and tells you plainly that
PST extraction now depends on `readpst`. A directory of `.eml` files needs neither backend and
is checked before either.

CI asserts both paths, because a fallback nobody exercises is a fallback that has already
broken. The test matrix runs 3.9, 3.11 and 3.13 without `libratom`, and a separate job installs
the full documented requirements on 3.11 and proves `libratom` imports and is detected. If
that job ever goes red, `libratom` has moved and `setup.sh`'s interpreter cap needs revisiting.

## Testing

The bash suites are offline by construction: they extract the real functions out of the
scripts and exercise them against a mocked `api_call`, so no account, credentials or network
are involved. Suites are discovered rather than listed, so a new `*_test.sh` runs without
anyone remembering to register it.

CI installs `pandoc` before running them and **fails the build if any suite skips a test**. The
`md_to_html` assertions skip themselves when `pandoc` is missing, and that helper builds the
body of every email the skill sends - a skipped test inside a green build is exactly the thing
worth failing on.
