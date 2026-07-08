#!/bin/bash
# Install the Outlook skill into ~/.codex/skills/ for Codex.
#
# Codex does not substitute ${CLAUDE_SKILL_DIR}, so this script rewrites that
# variable to each skill's installed Codex path and symlinks the scripts (edits
# stay live). Re-run after editing a SKILL.md.

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_ROOT="$HOME/.codex/skills"

echo "=== Outlook skill installer (Codex) ==="
echo

MISSING=""
command -v az >/dev/null 2>&1   || MISSING="$MISSING azure-cli"
command -v jq >/dev/null 2>&1   || MISSING="$MISSING jq"
command -v curl >/dev/null 2>&1 || MISSING="$MISSING curl"
if [ -n "$MISSING" ]; then
  echo "Missing required dependencies:$MISSING"
  exit 1
fi
command -v pandoc >/dev/null 2>&1 || echo "Optional: pandoc not found (needed for markdown-formatted emails)."
echo "Dependencies OK."
echo

for src in "$SCRIPT_DIR"/skills/*/; do
  src="${src%/}"
  name="$(basename "$src")"
  target="$SKILLS_ROOT/$name"
  echo "Installing '$name' -> $target"
  mkdir -p "$target"
  [ -d "$src/scripts" ]    && ln -sfn "$src/scripts"    "$target/scripts"
  [ -d "$src/references" ] && ln -sfn "$src/references" "$target/references"
  chmod +x "$src"/scripts/*.sh 2>/dev/null || true
  # Rewrite ${CLAUDE_SKILL_DIR} to this skill's Codex path. Cross-skill refs of
  # the form ${CLAUDE_SKILL_DIR}/../<other> then resolve to a sibling skill.
  sed "s#\${CLAUDE_SKILL_DIR}#$target#g" \
    "$src/SKILL.md" > "$target/SKILL.md"
done

echo
echo "Installed for Codex. Run scripts/outlook-setup.sh if you have not configured credentials."
