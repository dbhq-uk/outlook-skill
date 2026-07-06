# Developer setup - Outlook skill

Set the skill up from source with a **live symlink install**, so your edits are active immediately in Claude Code (and Codex). End users don't need this - they install via the [DBHQ marketplace](../README.md#install).

## Prerequisites

- `git` (and the GitHub CLI `gh` if you'll push changes)
- `jq`, `curl`, `azure-cli` (the `az` command), and optionally `pandoc` (for markdown-formatted emails)

## 1. Clone

```bash
git clone https://github.com/dbhq-uk/outlook-skill.git ~/dbhq-outlook
cd ~/dbhq-outlook
```

## 2. Install (symlink)

```bash
./install.sh          # Claude Code: symlinks into ~/.claude/skills (edits are live)
./install-codex.sh    # Codex: installs into ~/.codex/skills
```

The committed skill is plugin-native (script paths use `${CLAUDE_PLUGIN_ROOT}`). The installers symlink the `scripts/` directory into your skills folder and generate `SKILL.md` with those paths rewritten for a non-plugin install. Scripts are live-linked; **after editing a `SKILL.md`, re-run `./install.sh`** to regenerate the installed copy.

## 3. Credentials

Complete the setup the installer offers, or run it directly:

```bash
~/.claude/skills/outlook/scripts/outlook-setup.sh
```

It registers (or reuses) an Azure app and signs you in via OAuth. Credentials are stored under `~/.outlook/<account>/`, never in the repo, and tokens refresh automatically. If you've already set this up on another machine, you can copy `~/.outlook/` across instead of re-authenticating.

## 4. Verify

```bash
~/.claude/skills/outlook/scripts/outlook-token.sh test
```

Then, in Claude Code, try *"check my email"*.

## Working across machines

Editing a **script** under `~/dbhq-outlook` is live immediately. After editing a **`SKILL.md`**, re-run `./install.sh`. If you develop on more than one machine, `git pull` before you start and `git push` when done to keep them in sync.
