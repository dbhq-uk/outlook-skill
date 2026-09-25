# Handling a thread

How to take a message from arriving to answered without losing a recipient, mangling the
formatting, or sending before you meant to.

Paths below are shortened to `mail.sh` for the mail script. In a real terminal that is
`~/.claude/skills/outlook/scripts/outlook-mail.sh`, and inside the skill it is
`${CLAUDE_SKILL_DIR}/scripts/outlook-mail.sh`.

## Read the whole message first

```bash
mail.sh read <message-id>
```

`read` prints the full body, both recipient lists and the attachment list. `preview` prints
about 200 characters and is for navigation only. A short preview does not mean a short
message, and the part below the fold is where deadlines and requests live.

If the message is one turn of a long exchange, get the rest of it:

```bash
mail.sh thread <message-id>
```

That returns the whole conversation oldest-first, spanning inbox and sent items, so you can
see what was already promised before you answer.

## Reply, and keep everyone on it

```bash
mail.sh mdreply <message-id> "Thanks - **Friday** works. Draft attached."
```

This creates a draft. It does not send.

`reply`, `mdreply` and `followup` all call Graph's `createReplyAll`, so the draft carries
every original `To:` and `Cc:` recipient. That is deliberate. Silently dropping a Cc is a
failure you do not find out about: the assistant, the colleague or the audit address simply
stops seeing the thread, and nobody tells you. Estate agents, solicitors and accountants Cc
people as a matter of routine.

The three reply commands print every recipient they set. Read that output. If the list looks
shorter than you expected, the original had recipients you missed - go back and check before
sending.

**Prefer `mdreply` to `reply`.** Plain-text replies look broken in Outlook next to the HTML
they are quoting.

### Trimming to sender-only

When you actually mean sender-only, create the reply first and then cut it down:

```bash
mail.sh update <draft-id> to "sender@example.com"
mail.sh update <draft-id> cc ""
```

`to` replaces the line. `cc` and `bcc` **append** to what is already there, deduped
case-insensitively - passing an empty string is how you clear them.

### Editing a reply body

```bash
mail.sh update <draft-id> mdbody "Rewritten body in **markdown**"
```

Use `mdbody`, not `body`, on any draft made by `mdreply` or `followup`. The reply chain is
marked with an invisible span when the draft is created, and `mdbody` splits on that marker
to keep the quoted history. Plain `body` replaces the lot, quoted chain included.

## Send as an alias

Never guess an alias. List what the mailbox can actually send as:

```bash
mail.sh aliases
```

Then set it on the draft - this works on any draft, including replies and forwards, which is
why it is the only alias mechanism you need:

```bash
mail.sh update <draft-id> from "dan@example.com"
mail.sh send <draft-id>
```

Every draft command and `update from` print the From address. `send` reads the draft back
from Graph and prints From, To, Cc, Bcc, Subject and the attachment names before it posts, and
if it cannot read the draft it sends nothing. Which identity a message goes out as matters as
much as who receives it, so confirm it before sending.

To default every draft to an alias, export `OUTLOOK_FROM_ADDRESS`. It applies to `draft`,
`mddraft`, `reply`, `mdreply`, `followup` and `forward`, and each prints the From it set. If the
variable is set and a draft would still go from another address, `send` warns.

Three things to know before relying on this:

- The tenant must have `SendFromAliasEnabled` set. Without it Exchange quietly rewrites the
  From back to the primary address, so prove it with a real test send rather than assuming.
- An address the mailbox does not list warns rather than blocks, because SendAs rights on a
  *shared* mailbox are real and never appear in this mailbox's alias list. If the right is
  genuinely absent, `send` fails with `ErrorSendAsDenied` and nothing goes out.
- Check the alias domain has DKIM and DMARC before sending externally. The send succeeding
  and the message being delivered are different questions.

## Attachments

Coming in:

```bash
mail.sh attachments <message-id>          # list them
mail.sh download <message-id>             # all of them
mail.sh download <message-id> <att-id>    # one
```

Files land in `inbox/` under `CLAUDE_PROJECT_DIR`, or the current directory if that is unset.

Going out, one call per file:

```bash
mail.sh attach <draft-id> /path/to/report.pdf
mail.sh attach <draft-id> /path/to/data.xlsx
```

Under 3 MB it is a single base64 upload. Above that the script opens a Graph upload session
and streams 4 MB chunks with a progress bar, which is what makes 150 MB attachments possible
at all - Graph's simple upload cannot carry them.

### Signatures and inline images

Many mail clients block remote images until the reader allows them, so a logo linked from a
website often arrives as a broken box. An image attached inline and referenced by `cid:` shows
straight away.

```bash
mail.sh signature <draft-id> ~/signature/dan.html
mail.sh attach <draft-id> chart.png --inline chart   # one image, for <img src="cid:chart">
```

`signature` takes an HTML file. Each `<img>` whose quoted `src` is a local file, relative to
the HTML file or absolute, is uploaded as an inline attachment and its `src` rewritten to
`cid:`. Remote URLs and existing `cid:` or `data:` images are left alone. The block goes just
above the quoted chain on a reply, or at the end of a new message. It needs an HTML draft, so
make new mail with `mddraft`.

The block is marked, so `update mdbody` keeps it when it rewrites the message, and running
`signature` again replaces it and skips images already attached. `update body`, the plain-text
form, replaces the whole body and the signature with it.

## Forwards and chasers

```bash
mail.sh forward <message-id> "a@x.com, b@y.com" "FYI - **deadline is Friday**."
```

The forward draft carries the quoted message and its attachments. The comment is markdown and
optional.

```bash
mail.sh followup <sent-message-id>
mail.sh followup <sent-message-id> "Any thoughts on this when you get a moment?"
```

`followup` replies to your own sent message, so it lands in the existing thread rather than
starting a new one. Find the original with `mail.sh sent 20` first.

## Filing what is done

```bash
mail.sh archive <message-id>
mail.sh move <message-id> "Clients/Acme"
mail.sh batch-move "Projects" <id1> <id2> <id3>
```

A bare folder name matches case-insensitively anywhere in the tree, shallowest wins on a tie;
a `Parent/Child` path targets one specific folder. Use the path form whenever the same name
exists twice.

`batch-move` resolves the destination once and moves in batches of 20 through Graph's
`$batch` endpoint, so it is far quicker than looping `move`. IDs come from arguments or from
stdin:

```bash
mail.sh folder "Newsletters" 200 | awk '/^\[/ {print $2}' | mail.sh batch-move "Archive/Newsletters"
```

Listing rows are `[n] <short-id> | date | from | subject`, so the short ID is the second
whitespace-separated field of any line starting with `[`.

Moving a message does not change its ID, so the IDs from one listing still work after a move.
You can move a message twice without listing its new folder in between.

## A session that only reads

```bash
export OUTLOOK_READ_ONLY=1
```

Every command that writes or sends then refuses before it reaches Graph, and the listings and
reads still work. That includes flagging and categorising below, so leave it unset for a
session that files mail.

## Triage without moving anything

```bash
mail.sh flag <message-id>
mail.sh categorize <message-id> --add "Follow up"
mail.sh categorize <message-id> --remove "Follow up"
```

Prefer `--add` and `--remove` to the comma-separated form. The bare form *replaces* every
category on the message, which quietly discards labels someone else's rule or another agent
put there. Use the replacing form only when that is what you mean.

To list the messages that carry a category, in any folder:

```bash
mail.sh category "Follow up"              # newest 10, with short IDs
mail.sh category "Follow up" 50
```

Master categories are separate from what is on a message:

```bash
mail.sh categories                        # what the mailbox has
mail.sh mkcategory "Follow up" red
mail.sh rccategory "Follow up" "dark blue"
mail.sh rmcategory "Follow up"
```

There is no rename - Graph makes `displayName` immutable once a category exists. And
`rmcategory` only removes it from the master list: messages already carrying the label keep
it until you strip them with `categorize --remove`.
