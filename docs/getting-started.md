# Getting started

Install the pack, authenticate a mailbox, read your inbox, and send one message you wrote
yourself. About ten minutes, most of it waiting on Microsoft's consent screen.

This is the happy path and nothing else. Options, alternatives and every other verb are in
the [reference](reference.md).

## Before you start

You need a Microsoft 365 mailbox you can sign into, and `azure-cli`, `jq` and `curl` on the
machine. `pandoc` is optional but you want it - without it the markdown commands cannot
convert anything, and plain-text email looks it.

```bash
sudo apt install jq curl pandoc          # Debian / Ubuntu
brew install jq curl pandoc              # macOS
```

Azure CLI has its own installer, documented by
[Microsoft](https://learn.microsoft.com/cli/azure/install-azure-cli).

## 1. Install the pack

```
/plugin marketplace add dbhq-uk/marketplace
/plugin install outlook@dbhq
```

Both skills arrive together. If you would rather run from a clone, `./install.sh` symlinks
the same two directories into `~/.claude/skills/` - see [dev-setup](dev-setup.md).

The examples below spell out the script path as `~/.claude/skills/outlook/scripts/`,
which is where a local install puts it. A plugin install lives elsewhere and you will not
normally type the path at all: you ask in plain language and the skill runs the command.

## 2. Authenticate

```bash
~/.claude/skills/outlook/scripts/outlook-setup.sh
```

It registers an Azure app for you (or reuses one it finds), opens a sign-in, and asks you to
consent to five delegated permissions: `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`,
`User.Read` and `offline_access`. Those are the whole of what the pack can do. Nothing there
lets it read anyone else's mailbox, and nothing lets it act while you are not signed in
beyond the refresh token's life.

Credentials land in `~/.dbhq/outlook/default/`, mode `600`, and never go anywhere else.

Check it worked:

```bash
~/.claude/skills/outlook/scripts/outlook-token.sh test
```

```
Connection successful!
Inbox: 1,284 total, 12 unread
```

If that fails, [accounts and tokens](guides/accounts.md) covers every way it can, and
[`references/setup.md`](../skills/outlook/references/setup.md) walks the Azure
registration by hand.

## 3. Read the inbox

```bash
~/.claude/skills/outlook/scripts/outlook-mail.sh inbox 5
```

```
Fetching inbox (5 messages)...

[1] AAkALgAAAAAAHYQDEapm | 2026-07-30 | jane@example.com | Contract for review
[2] AAkALgAAAAAAHYQDEbn4 | 2026-07-30 | no-reply@example.org | Your invoice
...
```

The number in brackets is just the row. The 20-character token after it is a **short ID** -
the tail of the real Graph ID, which runs to several hundred characters. Every command that
takes an ID accepts the short form, and listings cache the mapping, so the obvious
list-then-act flow costs no extra API calls.

Now open one properly:

```bash
~/.claude/skills/outlook/scripts/outlook-mail.sh read AAkALgAAAAAAHYQDEapm
```

That prints the full body, the complete `To:` and `Cc:` lists, and any attachments. There is
also a `preview` verb; it returns roughly the first 200 characters and exists for finding the
right message in a list, never for deciding what a message says.

## 4. Write one to yourself

```bash
~/.claude/skills/outlook/scripts/outlook-mail.sh mddraft \
  "you@example.com" "First draft from the terminal" \
  "This is **markdown**, converted to HTML with the Aptos stack Outlook expects."
```

```
Draft created.
Draft ID: AAkALgAAAAAAHYQDEcm2
To: you@example.com
Subject: First draft from the terminal
```

Nothing has been sent. Open Outlook and the draft is sitting there, formatted. That gap
between creating and sending is the pack's core safety property, not an inconvenience: no
verb in either skill puts a message in front of another human without a second, separate
command.

Send it:

```bash
~/.claude/skills/outlook/scripts/outlook-mail.sh send AAkALgAAAAAAHYQDEcm2
```

## 5. Look at the calendar

```bash
~/.claude/skills/outlook/scripts/outlook-calendar.sh today
```

If the first line of output is a note about times being in UTC, read it. The script reports
wall-clock time in the system timezone, which on a server or container is UTC while you are
not - so a 14:00 London meeting comes back as 13:00. Export `OUTLOOK_TZ=Europe/London` (or
your own zone) and the note goes away.

```bash
export OUTLOOK_TZ=Europe/London
```

Put that in your shell profile once and forget it.

## Where to go next

You now have a working install. In practice you will drive all of this by asking - *"check my
email"*, *"reply to Jane and say the contract looks fine"*, *"am I free Thursday
afternoon"* - and the skill picks the commands.

- [Handling a thread](guides/email.md) for replies, recipients, attachments and aliases
- [Running the calendar](guides/calendar.md) for the two-step invite flow
- [Reference](reference.md) for every verb and flag
