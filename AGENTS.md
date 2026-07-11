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
- SKILL.md references scripts via `${CLAUDE_SKILL_DIR}` (the skill's own directory), which Claude Code substitutes for personal, project, and plugin installs alike. `install.sh` therefore symlinks the whole skill directory into `~/.claude/skills/` (no rewrite). `install-codex.sh` still rewrites the variable to the install path, since Codex does not substitute it.
- Shell scripts use `set -e`; errors go to stderr, structured output to stdout.
- No secrets in the repo - credentials live under `~/.outlook/`.
- House style: British English, plain hyphens.

## Validating a change

```bash
bash -n skills/outlook/scripts/*.sh          # scripts parse
shellcheck skills/outlook/scripts/*.sh       # lint (warnings should be clean)
bash skills/outlook/tests/helpers_test.sh    # offline unit tests (no account needed)
claude plugin validate .                     # manifest + structure
```

`tests/helpers_test.sh` extracts the pure helpers (search encoding/paging/sort,
folder resolution, token-expiry logic) from the scripts and runs them against a
mocked Graph API, so it catches regressions without a live mailbox. Anything that
needs real Graph calls (token refresh, actual search results) still wants a
manual smoke test against a configured account.
