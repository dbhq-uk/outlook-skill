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

The committed skill references its scripts via `${CLAUDE_SKILL_DIR}` (the skill's own directory), which Claude Code substitutes for personal, project and plugin installs alike. So `install.sh` symlinks the **whole skill directory** into `~/.claude/skills/` - `SKILL.md`, `scripts/` and `references/` are all live, and every edit (including `SKILL.md`) takes effect with no re-run. Re-run `install.sh` only when you add a new skill directory. Codex does not substitute `${CLAUDE_SKILL_DIR}`, so `install-codex.sh` rewrites it to the install path - **re-run `./install-codex.sh` after editing a `SKILL.md`** for Codex.

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

Editing **anything** under `~/dbhq-outlook` (scripts or `SKILL.md`) is live immediately in Claude Code - the whole skill directory is symlinked. For Codex, re-run `./install-codex.sh` after a `SKILL.md` edit. If you develop on more than one machine, `git pull` before you start and `git push` when done to keep them in sync.
