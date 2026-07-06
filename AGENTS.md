# AGENTS.md

Guidance for AI agents (and people) working in this repository.

## What this is

The **Outlook** skill for AI coding agents - Microsoft 365 email and calendar via the Microsoft Graph API. It follows the [Agent Skills](https://agentskills.io) layout (`skills/<name>/SKILL.md`) and ships as a [Claude Code plugin](https://code.claude.com/docs/en/plugins).

## Layout

```
.claude-plugin/plugin.json     # plugin manifest
skills/outlook/SKILL.md        # the skill (agent-facing instructions)
skills/outlook/scripts/        # bash scripts (self-contained: jq + curl + az)
skills/outlook/references/     # manual setup guide
install.sh / install-codex.sh  # local symlink installers (Claude / Codex)
```

## Conventions

- Scripts are self-contained: they read credentials from `~/.outlook/<account>/` and have no bundled-path dependencies, so they run from any location.
- SKILL.md references scripts via `${CLAUDE_PLUGIN_ROOT}`, which Claude Code substitutes for plugin installs. The installers rewrite that variable to the install path for local symlink installs (Claude and Codex).
- Shell scripts use `set -e`; errors go to stderr, structured output to stdout.
- No secrets in the repo - credentials live under `~/.outlook/`.
- House style: British English, plain hyphens.

## Validating a change

```bash
bash -n skills/outlook/scripts/*.sh     # scripts parse
claude plugin validate .                # manifest + structure
```
