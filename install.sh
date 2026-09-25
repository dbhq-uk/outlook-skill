#!/bin/bash
# Install the Outlook skill pack into ~/.claude/skills/ as a live symlink install.
#
# SKILL.md references scripts via ${CLAUDE_SKILL_DIR}, which Claude Code
# substitutes to the skill's own directory for personal, project, and plugin
# installs alike. So this script symlinks the whole skill directory into
# ~/.claude/skills/ - every edit (scripts AND SKILL.md) is immediately live,
# with no per-file rewrite. Re-run only when you add a new skill directory.

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_ROOT="$HOME/.claude/skills"

echo "=== Outlook skill pack installer (Claude Code) ==="
echo

# --- Dependencies, per skill ---
# Checked per skill rather than globally: the two skills share no dependencies,
# so a missing azure-cli must not block someone who only wants outlook-to-md.
# A skill whose required tools are absent is skipped with a reason, not fatal.
missing_for() {
  local skill="$1" missing=""
  case "$skill" in
    outlook)
      command -v az   >/dev/null 2>&1 || missing="$missing azure-cli"
      command -v jq   >/dev/null 2>&1 || missing="$missing jq"
      command -v curl >/dev/null 2>&1 || missing="$missing curl"
      ;;
    outlook-to-md)
      command -v python3 >/dev/null 2>&1 || missing="$missing python3"
      ;;
  esac
  echo "$missing"
}

command -v pandoc  >/dev/null 2>&1 || echo "Optional: pandoc not found (needed for markdown-formatted emails)."
command -v readpst >/dev/null 2>&1 || echo "Optional: readpst not found (pst-utils; needed to read .pst files, not for live-mail archives)."
echo

# --- Retire skills that have been renamed -----------------------------------
# pst-to-markdown became outlook-to-md once it also ingested live mail, and
# outlook-graph became outlook on 17 Sep 2026. An old install is a symlink to a
# directory that no longer exists, so without this it lingers beside the new
# skill as a dangling duplicate the agent may still match - and a dangling
# outlook-graph is worse than a dangling pst-to-markdown, because an agent
# reading its SKILL.md would call four scripts that are no longer there.
for stale in pst-to-markdown outlook-graph; do
  if [ -e "$SKILLS_ROOT/$stale" ] || [ -L "$SKILLS_ROOT/$stale" ]; then
    echo "Removing renamed skill '$stale'"
    rm -rf "${SKILLS_ROOT:?}/$stale"
  fi
done

# --- Install each skill in this repo as a full-directory symlink ---
mkdir -p "$SKILLS_ROOT"
INSTALLED=0
for src in "$SCRIPT_DIR"/skills/*/; do
  src="${src%/}"
  name="$(basename "$src")"
  target="$SKILLS_ROOT/$name"

  MISSING="$(missing_for "$name")"
  if [ -n "$MISSING" ]; then
    echo "Skipping '$name' - missing required:$MISSING"
    echo "  macOS:  brew install$MISSING"
    echo "  Ubuntu: sudo apt install$MISSING"
    continue
  fi

  echo "Installing '$name' -> $target"
  rm -rf "$target"            # replace any prior copy or partial-symlink install
  ln -sfn "$src" "$target"    # whole-directory symlink; ${CLAUDE_SKILL_DIR} resolves it
  chmod +x "$src"/scripts/*.sh 2>/dev/null || true
  INSTALLED=$((INSTALLED + 1))

  # Skills carrying a setup.sh provision their own environment (outlook-to-md
  # builds its Python venv). Non-fatal: the skill is installed either way.
  if [ -x "$src/setup.sh" ]; then
    echo "  Running $name setup..."
    "$src/setup.sh" || echo "  Setup failed for '$name'; re-run $src/setup.sh when ready."
  fi
done

if [ "$INSTALLED" -eq 0 ]; then
  echo
  echo "Nothing installed - every skill was missing a required dependency."
  exit 1
fi

echo
echo "Installed as directory symlinks - all edits (scripts and SKILL.md) are live. Re-run only when adding a new skill."
echo

# --- Setup / credentials (outlook only; outlook-to-md needs none) ---
SETUP="$SKILLS_ROOT/outlook/scripts/outlook-setup.sh"
if [ ! -e "$SKILLS_ROOT/outlook" ]; then
  echo "outlook was not installed - skipping credential setup."
# Every home the credentials have had, so an existing install is never sent
# back through setup. outlook-token.sh moves them on its first run; this only
# needs to RECOGNISE them, which is why it looks at all four paths.
elif [ -f "$HOME/.dbhq/outlook/default/credentials.json" ] \
  || [ -f "$HOME/.dbhq/outlook-graph/default/credentials.json" ] \
  || [ -f "$HOME/.outlook-graph/default/credentials.json" ] \
  || [ -f "$HOME/.outlook-graph/credentials.json" ]; then
  echo "Existing Outlook credentials found. Re-run setup any time with:"
  echo "  $SETUP"
else
  echo "No credentials found. Launching setup..."
  echo
  "$SETUP" || echo "Setup skipped or failed; run '$SETUP' when ready."
fi

echo
echo "Done. Try: 'check my email', 'what's on my calendar today', or 'extract archive.pst'"
