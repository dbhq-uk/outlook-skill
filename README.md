<div align="center">

<img src="assets/logo.svg" alt="outlook skill for Claude Code, by DBHQ" width="560">

# outlook

**Your Microsoft 365 mail, calendar and archives in the terminal - driven by Claude Code or Codex**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Claude Code](https://img.shields.io/badge/Claude_Code-Plugin-blueviolet)](https://code.claude.com/docs/en/plugins)
[![Platform](https://img.shields.io/badge/Platform-Linux%20%7C%20macOS%20%7C%20WSL-lightgrey)]()

A free, open-source tool by [DBHQ](https://dbhq.uk) - documented at [skills.dbhq.uk](https://skills.dbhq.uk/outlook/)

</div>

---

Read your inbox, draft and send properly formatted replies and forwards, triage with flags and
categories, manage attachments up to 150 MB, and run your calendar - including responding to
invites and inviting attendees - all from Claude Code or Codex, in plain language.
Multi-account, OAuth-based, and built with the safety rails that matter for real
correspondence.

Two skills ship in this pack:

| Skill | What it does | Needs |
|---|---|---|
| **`outlook`** | Live Microsoft 365 mail and calendar via the Graph API | OAuth, network |
| **`outlook-to-md`** | Turns PST exports and live mail into integrity-verified markdown, offline | Nothing but a file |

They cover the two halves of the same problem: the mail you are handling now, and the mail you
were handed in a box - and they join up, so one archive spans both. `outlook-to-md` itself
needs no credentials and makes no network calls; it reads files on disk, whether they came out
of a PST or out of `outlook`.

## What makes it different

**Reply-all by default.** Replies preserve every original `To:` and `Cc:` recipient, so you
never silently drop someone from a thread. Trim to sender-only when you actually mean to.
Losing a recipient by accident is invisible from your side; sending to one person too many is
not.

**Reads the whole message, never the preview.** The skill is instructed to open the full body
end-to-end before summarising or replying, so deadlines, attachments and requests buried below
the fold are not missed.

**Nothing leaves without a second command.** Drafting and sending are different verbs, and so
are creating an event and inviting anyone to it. An agent driving a mailbox is one confident
inference away from mailing a client; the gap is where a person gets to look.

**Time-aware.** It anchors "today", "tomorrow" and "by EOD" against the real clock and tracks
BST against UTC, so scheduled sends and deadline arithmetic are correct. Calendar times are
wall-clock in a timezone it tells you about rather than assumes.

**Professional formatting.** Markdown drafts convert to clean HTML with the Microsoft 365
Aptos font stack and inline styles - including per-paragraph margins - that survive Outlook's
rendering.

**It asks for five permissions and no more.** `Mail.ReadWrite`, `Mail.Send`,
`Calendars.ReadWrite`, `User.Read`, `offline_access`, all delegated. Nothing tenant-wide,
nothing that can reach another mailbox, and no admin consent to obtain.

## What it covers

```bash
outlook-mail.sh      inbox · unread · focused · sent · drafts · flagged · folder · from
                           search · thread · read · preview · export
                           draft · mddraft · reply · mdreply · forward · followup
                           update · send · aliases
                           attachments · download · attach
                           markread · flag · categorize · categories · junk · archive · delete
                           move · batch-move · mkdir · rename · rmdir · folders · stats

outlook-calendar.sh  events · today · week · day · search · read · calendars
                           create · invite · quick · update · respond · cancel · delete · free

outlook-token.sh     refresh · get · test · status · list

outlook_to_md.py           <pst-or-eml-dir> <output-dir> [--append --timezone --owner-email …]
```

You normally type none of this - you ask, and the skill picks the command. Every argument,
flag and default is in [docs/reference.md](docs/reference.md).

## Install

### As a Claude Code plugin (recommended)

```
/plugin marketplace add dbhq-uk/marketplace
/plugin install outlook@dbhq
```

### Any agent (Cursor, Copilot, Windsurf, Gemini, Cline and more)

```bash
npx skills add dbhq-uk/outlook-skill
```

The [skills.sh](https://skills.sh) CLI installs into whichever agent directories
it finds, so this works outside Claude Code and Codex too.

### Local install (Claude Code or Codex)

```bash
git clone https://github.com/dbhq-uk/outlook-skill.git
cd outlook-skill
./install.sh          # Claude Code: symlinks into ~/.claude/skills (edits are live)
./install-codex.sh    # Codex: installs into ~/.codex/skills
```

[`install.sh`](install.sh) and [`install-codex.sh`](install-codex.sh) are the
same install two ways: Claude Code substitutes `${CLAUDE_SKILL_DIR}`, so the
whole skill directory is symlinked untouched, while Codex does not, so its
`SKILL.md` is rewritten at install time. Re-run the Codex one after editing
`SKILL.md`.

## Requirements

Checked per skill - a missing dependency skips that skill rather than failing the install, so
you can take either half on its own.

| Skill | Required | Optional |
|---|---|---|
| `outlook` | `azure-cli` · `jq` · `curl` | `pandoc` (markdown-formatted emails) |
| `outlook-to-md` | `python3` (3.9+) | `readpst` (`pst-utils`; fallback PST backend) |

`outlook-to-md` provisions its own virtualenv on install. Its preferred PST backend pins an
old numpy and cannot install on Python 3.12 or later, so `setup.sh` picks a 3.9-3.11
interpreter when one exists and otherwise builds without it and says so - see [the Python
version problem](docs/architecture.md#the-python-version-problem).

## Documentation

**[The documentation index](docs/README.md)** reaches everything. Start with [getting
started](docs/getting-started.md) to set it up, [handling a thread](docs/guides/email.md) and
[running the calendar](docs/guides/calendar.md) to use it, [the
reference](docs/reference.md) to look a command up, and [how the pack is
built](docs/architecture.md) to understand why it behaves the way it does. There are also
guides for [markdown archives](docs/guides/archiving.md) and [accounts and
tokens](docs/guides/accounts.md).

Hacking on it, or running from source with live edits: [docs/dev-setup.md](docs/dev-setup.md),
[CONTRIBUTING.md](CONTRIBUTING.md), and [AGENTS.md](AGENTS.md) if you are an AI agent doing so.

## Credentials and privacy

No secrets live in this repository. Your tokens are stored locally under `~/.dbhq/outlook/`
and used only to talk to Microsoft Graph directly from your machine.

## Also from DBHQ

Fifteen free agent skills, all of them installable from the same marketplace and
all documented at **[skills.dbhq.uk](https://skills.dbhq.uk)**.

| Skill | What it does |
|---|---|
| [trello](https://skills.dbhq.uk/trello/) | Your boards, run from your agent |
| [legwork](https://skills.dbhq.uk/legwork/) | Research that settles a decision, and says when it cannot |
| [dovetail](https://skills.dbhq.uk/dovetail/) | Checks whether your repository still agrees with itself |
| [verve](https://skills.dbhq.uk/verve/) | Strips AI tells from prose and puts a voice back |
| [vela](https://skills.dbhq.uk/vela/) | Compiler-exact code search, in any language you index |
| [garmin](https://skills.dbhq.uk/garmin/) | Your Garmin data, answered in the terminal |
| [imager](https://skills.dbhq.uk/imager/) | Images from GPT Image 2, costed before it spends |
| [gitview](https://skills.dbhq.uk/gitview/) | Which branches are finished, and safe to delete |
| [atlassian](https://skills.dbhq.uk/atlassian/) | Jira issues and Confluence pages |
| [pennyblack](https://skills.dbhq.uk/pennyblack/) | A physical letter, posted from the terminal |
| [buildwork](https://skills.dbhq.uk/buildwork/) | Your open issues, run as parallel agents |
| [deskwork](https://skills.dbhq.uk/deskwork/) | What an agent noticed, tracked as real work |
| [groupwork](https://skills.dbhq.uk/groupwork/) | A second agent on the work, adversary or partner |

Plus [heliograph](https://skills.dbhq.uk/heliograph/), for a machine you cannot log into.

## Licence

[MIT](LICENSE) © 2026 DBHQ Consulting Ltd
