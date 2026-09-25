# AGENTS.md

Guidance for AI agents (and people) working in this repository.

## What this is

The **Outlook** skill for AI coding agents - Microsoft 365 email and calendar via the Microsoft Graph API. It follows the [Agent Skills](https://agentskills.io) layout (`skills/<name>/SKILL.md`) and ships as a [Claude Code plugin](https://code.claude.com/docs/en/plugins).

## Layout

```
.claude-plugin/plugin.json        # plugin manifest
.github/workflows/validate.yml    # CI: parse, unit tests, frontmatter, py matrix, macOS
hooks/                            # Claude Code send gap: PreToolUse hook, ask rules, their installer
skills/outlook/SKILL.md     # the live mail/calendar skill (agent-facing)
skills/outlook/scripts/     # bash scripts (jq + curl + az)
skills/outlook/scripts/lib/ # graph.sh: token code shared by the scripts, sourced not run
skills/outlook/references/  # commands.md (full command reference) and setup.md (manual setup); ship with the skill
skills/outlook-to-md/SKILL.md     # the offline archive skill (agent-facing)
skills/outlook-to-md/scripts/     # outlook_to_md.py, run from its own .venv
docs/                             # human-facing documentation, see docs/README.md
install.sh / install-codex.sh     # local symlink installers (Claude / Codex)
```

Two audiences, kept apart deliberately. `skills/*/SKILL.md` and `skills/*/references/` are
loaded by the agent at runtime and ship with the skill. `docs/` is for people and is not
installed - so a fact needed at runtime belongs in the skill, not only in `docs/`.

## Conventions

- Scripts read credentials from `~/.dbhq/outlook/<account>/`. Their one bundled dependency is `scripts/lib/graph.sh`, which holds the token code and is sourced from beside each script's real location (symlinks followed), so they run from any location. Token or refresh logic goes in that file, never back into a script: it was once copied into three scripts, and a bug in it existed three times.
- SKILL.md references scripts via `${CLAUDE_SKILL_DIR}` (the skill's own directory), which Claude Code substitutes for personal, project, and plugin installs alike. `install.sh` therefore symlinks the whole skill directory into `~/.claude/skills/` (no rewrite). `install-codex.sh` still rewrites the variable to the install path, since Codex does not substitute it.
- Shell scripts use `set -e`; errors go to stderr, structured output to stdout.
- No secrets in the repo - credentials live under `~/.dbhq/outlook/`.
- **A verb that sends is a decision, not an edit.** Every new verb must be classed in
  `send_gap_test.sh` as read or write, and only a verb that reads goes in the script's
  `READ_ONLY_VERBS`. If it sends, it also needs a rule in `hooks/ask-rules.json` and a case in
  `hooks/send-gate.sh`, and `send_gate_test.sh` must list it. See `docs/architecture.md`,
  "The send gap".
- Never add ask rules to a user's settings without their yes. `install.sh` offers them on a
  terminal or with `--ask-rules`, and not otherwise.
- Scripts run on GNU tools (Linux, WSL) and on BSD tools with bash 3.2 (macOS). A `date`,
  `dd`, `sed` or `stat` flag that only one of them has needs a fallback for the other, and
  no bash 4 features (associative arrays, `mapfile`, `${x,,}`). The `macos` CI job runs the
  offline suites with the Mac's own tools and `/bin/bash`.
- `SKILL.md` loads in full on every trigger, so it keeps to the rules and a one-line command
  index, under 1,500 words. The full syntax goes in `references/commands.md`. A new verb needs
  a line in one of them; `docs_test.sh` fails otherwise.
- House style: British English, plain hyphens.

## Validating a change

```bash
bash -n skills/outlook/scripts/*.sh    # scripts parse
shellcheck skills/outlook/scripts/*.sh # lint (warnings should be clean)
bash skills/outlook/tests/helpers_test.sh  # offline unit tests (no account needed)
bash skills/outlook/tests/token_test.sh    # token refresh against a fake curl
bash skills/outlook/tests/calendar_test.sh # calendar verbs against a fake curl
bash skills/outlook/tests/mail_test.sh     # draft From and send summary against a fake curl
bash skills/outlook/tests/graph_test.sh    # timeouts, 429/503 retries and batch-move failures
bash skills/outlook/tests/setup_test.sh    # setup: public client, PKCE, no secret, the old-secret path
bash skills/outlook/tests/docs_test.sh     # SKILL.md word limit and dashes; every verb documented
bash skills/outlook/tests/send_gap_test.sh # read-only mode and the calendar send flags
bash skills/outlook/tests/send_gate_test.sh # the hook, the ask rules and their installer
python3 -m pytest skills/outlook-to-md/tests/ -q # archive suite (no PST needed)
claude plugin validate .                     # manifest + structure
```

`tests/helpers_test.sh` extracts the pure helpers (search encoding/paging/sort,
folder resolution, token-expiry logic) from the scripts and runs them against a
mocked Graph API, so it catches regressions without a live mailbox.
`tests/token_test.sh` sources `scripts/lib/graph.sh` and runs the real refresh
against a fake `curl` and a throwaway `HOME`, so it never touches real
credentials. Anything that needs real Graph calls (actual search results) still
wants a manual smoke test against a configured account.
