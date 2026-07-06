<div align="center">

# 📬 Outlook for Claude Code

**Microsoft 365 email and calendar in your terminal - driven by Claude Code or Codex, powered by the Microsoft Graph API**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Claude Code](https://img.shields.io/badge/Claude_Code-Plugin-blueviolet)](https://code.claude.com/docs/en/plugins)
[![Platform](https://img.shields.io/badge/Platform-Linux%20%7C%20macOS%20%7C%20WSL-lightgrey)]()

A free, open-source tool by [DBHQ](https://dbhq.uk)

</div>

---

Read your inbox, draft and send properly formatted replies, manage attachments up to 150 MB, and work with your calendar - all from Claude Code or Codex, in plain language. Multi-account, OAuth-based, and built with the safety rails that matter for real correspondence.

## Why it is different

- **Reply-all by default** - replies preserve every original `To:` and `Cc:` recipient, so you never silently drop someone from a thread. Trim to sender-only when you actually mean to.
- **Reads the whole message, never the preview** - the skill is instructed to open the full body end-to-end before summarising or replying, so deadlines, attachments and requests buried below the fold are not missed.
- **Time-aware** - anchors "today", "tomorrow" and "by EOD" against the real clock and tracks BST vs UTC, so scheduled sends and deadline maths are correct.
- **Professional formatting** - markdown drafts convert to clean HTML with inline styles that survive Outlook's rendering.

## Features

- Inbox, unread, focused, sent, search, and any folder by name
- Full-message reading with attachment listing and download
- Plain and markdown drafts, reply-all-safe replies, follow-up chasers
- Attachments up to 150 MB (automatic chunked upload)
- Folder management, archive, move, mark read/unread
- Calendar: upcoming, today, this week, event details, quick-create, free/busy
- Multiple accounts, selected by flag or environment variable

## Install

### As a Claude Code plugin (recommended)

```
/plugin marketplace add dbhq-uk/marketplace
/plugin install outlook@dbhq
```

Then run the one-time setup the skill points you to, and talk to it in plain language: *"check my email"*, *"draft a reply to the last message from Sam"*, *"am I free Thursday afternoon"*.

### Local install (Claude Code or Codex)

```bash
git clone https://github.com/dbhq-uk/outlook-skill.git
cd outlook-skill
./install.sh          # Claude Code: symlinks into ~/.claude/skills (edits are live)
./install-codex.sh    # Codex: installs into ~/.codex/skills
```

## Setup

First run launches `outlook-setup.sh`, which registers an Azure app and authenticates you via OAuth. Credentials are stored per account under `~/.outlook/<account>/` and never leave your machine. Tokens refresh automatically. See [`skills/outlook/references/setup.md`](skills/outlook/references/setup.md) for manual steps.

## Requirements

`azure-cli` · `jq` · `curl` · `pandoc` (optional, for markdown-formatted emails)

## Credentials and privacy

No secrets live in this repository. Your tokens are stored locally under `~/.outlook/` and used only to talk to Microsoft Graph directly from your machine.

## License

[MIT](LICENSE) © 2026 DBHQ Consulting Ltd
