#!/bin/bash
# Install the Outlook skill into ~/.claude/skills/ as a live symlink install.
#
# The committed skill is plugin-native: SKILL.md references scripts via
# ${CLAUDE_PLUGIN_ROOT}, which Claude Code substitutes for marketplace/plugin
# installs. For a local symlink install (edit-and-see-live), this script rewrites
# that variable to the installed path and symlinks the scripts so your edits are
# immediately live. Re-run this script after editing a SKILL.md.

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

# --- Install each skill in this repo ---
for src in "$SCRIPT_DIR"/skills/*/; do
  src="${src%/}"
  name="$(basename "$src")"
  target="$SKILLS_ROOT/$name"
  echo "Installing '$name' -> $target"
  mkdir -p "$target"
  # Live symlinks for the parts you edit often
  [ -d "$src/scripts" ]    && ln -sfn "$src/scripts"    "$target/scripts"
  [ -d "$src/references" ] && ln -sfn "$src/references" "$target/references"
  chmod +x "$src"/scripts/*.sh 2>/dev/null || true
  # Generate SKILL.md with the plugin-root variable rewritten to the install path.
  # Generic (any skill), so cross-skill references within a pack also resolve.
  sed 's#\${CLAUDE_PLUGIN_ROOT}/skills/#$HOME/.claude/skills/#g' \
    "$src/SKILL.md" > "$target/SKILL.md"
done

echo
echo "Installed. Scripts are symlinked (edits are live); re-run this script after editing a SKILL.md."
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
