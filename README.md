<div align="center">

<img src="assets/logo.svg" alt="Outlook skill for Claude Code, by DBHQ" width="420">

# Outlook for Claude Code

**Microsoft 365 email and calendar in your terminal - driven by Claude Code or Codex, powered by the Microsoft Graph API**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Claude Code](https://img.shields.io/badge/Claude_Code-Plugin-blueviolet)](https://code.claude.com/docs/en/plugins)
[![Platform](https://img.shields.io/badge/Platform-Linux%20%7C%20macOS%20%7C%20WSL-lightgrey)]()

A free, open-source tool by [DBHQ](https://dbhq.uk)

</div>

---

Read your inbox, draft and send properly formatted replies and forwards, triage with flags and categories, manage attachments up to 150 MB, and run your calendar - including responding to invites and inviting attendees - all from Claude Code or Codex, in plain language. Multi-account, OAuth-based, and built with the safety rails that matter for real correspondence.

## Why it is different

- **Reply-all by default** - replies preserve every original `To:` and `Cc:` recipient, so you never silently drop someone from a thread. Trim to sender-only when you actually mean to.
- **Reads the whole message, never the preview** - the skill is instructed to open the full body end-to-end before summarising or replying, so deadlines, attachments and requests buried below the fold are not missed.
- **Time-aware** - anchors "today", "tomorrow" and "by EOD" against the real clock and tracks BST vs UTC, so scheduled sends and deadline maths are correct.
- **Professional formatting** - markdown drafts convert to clean HTML with the Microsoft 365 Aptos font stack and inline styles (including per-paragraph margins) that survive Outlook's rendering.
- **Draft-first, always** - nothing is sent, forwarded, or invited without showing you the draft and waiting for your explicit go-ahead.

## Features

**Email - reading and triage**
- Inbox, unread, focused, sent, drafts, flagged, and any folder by name (recursive, `Parent/Child` paths supported)
- Full-message reading with attachment listing; whole-conversation thread view (oldest first)
- Search: free text or KQL field operators (`subject:`, `from:`, `to:`, `body:`, AND/OR/NOT), paged up to 1000 results, newest first
- Flag / unflag for follow-up, categories (list master categories, apply, clear), junk / not-junk, mark read/unread

**Email - writing and sending**
- Plain and markdown drafts, reply-all-safe replies, forwards with an optional markdown comment, follow-up chasers
- Draft editing: subject, body (plain or markdown, reply chain preserved), To/Cc/Bcc (deduped), importance (high/normal/low)
- Attachments up to 150 MB (automatic chunked upload with progress), download to `./inbox/`

**Organisation**
- Archive, move, and batch-move (batches of 20 via Graph `$batch`, short IDs resolved automatically)
- Folder create / rename / delete (system folders protected, non-empty needs `--force`), subfolder listing, inbox stats

**Calendar**
- Upcoming, today, this week, any date, and event search by subject/location
- Event details, create, quick 1-hour create, update, delete
- Two-step meeting flow: create the event first (nothing sent), then `invite` attendees (required or optional) once approved - invitations only go out at that point
- Respond to invitations (accept / decline / tentative, with comment) and cancel meetings you organise
- Free/busy availability; all times are wall-clock in your timezone, with day/week/free windows offset-qualified so they cannot drift across a midnight boundary

> **Set `OUTLOOK_TZ` if your machine is not in your own timezone.** Calendar times default to the *system* timezone; on a server or container that is usually UTC, which would report a 14:00 London meeting as 13:00. Export `OUTLOOK_TZ=Europe/London` (or your zone) — the calendar script warns you when it is falling back to UTC.

**Accounts**
- Multiple accounts under `~/.outlook/<account>/`, selected by `--account` flag or `OUTLOOK_ACCOUNT` env var; one Azure app registration can be reused across mailboxes

## Install

### As a Claude Code plugin (recommended)

```
/plugin marketplace add dbhq-uk/marketplace
/plugin install outlook-graph@dbhq
```

Then run the one-time setup the skill points you to, and talk to it in plain language: *"check my email"*, *"draft a reply to the last message from Sam"*, *"am I free Thursday afternoon"*.

### Local install (Claude Code or Codex)

```bash
git clone https://github.com/dbhq-uk/outlook-graph-skill.git
cd outlook-graph-skill
./install.sh          # Claude Code: symlinks into ~/.claude/skills (edits are live)
./install-codex.sh    # Codex: installs into ~/.codex/skills
```

## Setup

First run launches `outlook-setup.sh`, which registers an Azure app and authenticates you via OAuth. Credentials are stored per account under `~/.outlook/<account>/` and never leave your machine. Tokens refresh automatically. See [`skills/outlook-graph/references/setup.md`](skills/outlook-graph/references/setup.md) for manual steps.

## Command reference

You normally just talk to the skill in plain language, but every command is also usable directly.

`outlook-mail.sh`:

| Area | Commands |
|---|---|
| Read | `inbox` · `unread` · `focused` · `sent` · `drafts` · `flagged` · `folder <name>` · `from <email>` · `search <query>` · `thread <id>` · `read <id>` · `preview <id>` |
| Write | `draft` · `mddraft` · `reply` · `mdreply` · `forward <id> <to> [comment]` · `followup <sent-id>` · `update <draft> subject\|body\|mdbody\|to\|cc\|bcc\|importance` · `send <draft>` |
| Attachments | `attachments <id>` · `download <id> [att-id]` · `attach <draft> <file>` (up to 150 MB) |
| Triage | `markread` · `markunread` · `flag` · `unflag` · `categorize <id> <cats>` · `categories` · `junk` · `notjunk` · `archive` · `delete` |
| Organise | `move <id> <folder>` · `batch-move <folder> <ids…>` · `mkdir` · `rename` · `rmdir [--force]` · `folders` · `subfolders` · `stats` |

`outlook-calendar.sh`:

| Area | Commands |
|---|---|
| View | `events` · `today` · `week` · `day <date>` · `search <text> [days]` · `read <id>` · `calendars` |
| Create | `create <subject> <start> <end> [location] [attendees]` · `invite <id> <emails> [required\|optional]` · `quick <subject> <start>` |
| Manage | `update <id> <field> <value>` · `respond <id> accept\|decline\|tentative [comment]` · `cancel <id> [comment]` · `delete <id>` |
| Availability | `free <start> <end>` |

`outlook-token.sh`: `refresh` · `get` · `test` · `status` · `list` (accounts).

All scripts accept `--account <name>` / `-a <name>` (or `OUTLOOK_ACCOUNT`) before the command.

## Development

Want to hack on this skill or run it from source with live edits? See [`docs/dev-setup.md`](docs/dev-setup.md).

## Requirements

`azure-cli` · `jq` · `curl` · `pandoc` (optional, for markdown-formatted emails)

## Credentials and privacy

No secrets live in this repository. Your tokens are stored locally under `~/.outlook/` and used only to talk to Microsoft Graph directly from your machine.

## License

[MIT](LICENSE) © 2026 DBHQ Consulting Ltd
