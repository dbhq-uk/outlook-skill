#!/bin/bash
# Install the Outlook skill into ~/.claude/skills/ as a live symlink install.
#
# SKILL.md references scripts via ${CLAUDE_SKILL_DIR}, which Claude Code
# substitutes to the skill's own directory for personal, project, and plugin
# installs alike. So this script symlinks the whole skill directory into
# ~/.claude/skills/ - every edit (scripts AND SKILL.md) is immediately live,
# with no per-file rewrite. Re-run only when you add a new skill directory.

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_ROOT="$HOME/.claude/skills"

echo "=== Outlook skill installer (Claude Code) ==="
echo

# --- Dependencies ---
MISSING=""
command -v az >/dev/null 2>&1   || MISSING="$MISSING azure-cli"
command -v jq >/dev/null 2>&1   || MISSING="$MISSING jq"
command -v curl >/dev/null 2>&1 || MISSING="$MISSING curl"
if [ -n "$MISSING" ]; then
  echo "Missing required dependencies:$MISSING"
  echo "  macOS:  brew install$MISSING"
  echo "  Ubuntu: sudo apt install$MISSING"
  exit 1
fi
command -v pandoc >/dev/null 2>&1 || echo "Optional: pandoc not found (needed for markdown-formatted emails)."
echo "Dependencies OK."
echo

# --- Install each skill in this repo as a full-directory symlink ---
mkdir -p "$SKILLS_ROOT"
for src in "$SCRIPT_DIR"/skills/*/; do
  src="${src%/}"
  name="$(basename "$src")"
  target="$SKILLS_ROOT/$name"
  echo "Installing '$name' -> $target"
  rm -rf "$target"            # replace any prior copy or partial-symlink install
  ln -sfn "$src" "$target"    # whole-directory symlink; ${CLAUDE_SKILL_DIR} resolves it
  chmod +x "$src"/scripts/*.sh 2>/dev/null || true
done

echo
echo "Installed as directory symlinks - all edits (scripts and SKILL.md) are live. Re-run only when adding a new skill."
echo

# --- Setup / credentials ---
SETUP="$SKILLS_ROOT/outlook/scripts/outlook-setup.sh"
if [ -f "$HOME/.outlook/default/credentials.json" ] || [ -f "$HOME/.outlook/credentials.json" ]; then
  echo "Existing Outlook credentials found. Re-run setup any time with:"
  echo "  $SETUP"
else
  echo "No credentials found. Launching setup..."
  echo
  "$SETUP" || echo "Setup skipped or failed; run '$SETUP' when ready."
fi

echo
echo "Done. Try: 'check my email' or 'what's on my calendar today'"
