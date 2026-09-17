#!/bin/bash
# Install the Outlook skill pack into ~/.codex/skills/ for Codex.
#
# Codex does not substitute ${CLAUDE_SKILL_DIR}, so this script rewrites that
# variable to each skill's installed Codex path and symlinks the scripts (edits
# stay live). Re-run after editing a SKILL.md.

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILLS_ROOT="$HOME/.codex/skills"

echo "=== Outlook skill pack installer (Codex) ==="
echo

# Per-skill dependency checks - see install.sh for why this is not global.
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
command -v readpst >/dev/null 2>&1 || echo "Optional: readpst not found (pst-utils; fallback PST backend if libratom fails)."
echo

# pst-to-markdown became outlook-to-md once it also ingested live mail, and
# outlook-graph became outlook on 17 Sep 2026. An old Codex install holds
# symlinks into a directory that no longer exists, so clear it out rather than
# leave a broken duplicate skill beside the new one.
for stale in pst-to-markdown outlook-graph; do
  if [ -e "$SKILLS_ROOT/$stale" ] || [ -L "$SKILLS_ROOT/$stale" ]; then
    echo "Removing renamed skill '$stale'"
    rm -rf "${SKILLS_ROOT:?}/$stale"
  fi
done

INSTALLED=0
for src in "$SCRIPT_DIR"/skills/*/; do
  src="${src%/}"
  name="$(basename "$src")"
  target="$SKILLS_ROOT/$name"

  MISSING="$(missing_for "$name")"
  if [ -n "$MISSING" ]; then
    echo "Skipping '$name' - missing required:$MISSING"
    continue
  fi

  echo "Installing '$name' -> $target"
  mkdir -p "$target"
  [ -d "$src/scripts" ]    && ln -sfn "$src/scripts"    "$target/scripts"
  [ -d "$src/references" ] && ln -sfn "$src/references" "$target/references"
  # The rewritten SKILL.md tells the user to run ${CLAUDE_SKILL_DIR}/setup.sh,
  # so setup.sh (and the requirements.txt it reads) must exist at $target too -
  # otherwise the one command the skill documents resolves to nothing.
  [ -f "$src/setup.sh" ]         && ln -sfn "$src/setup.sh"         "$target/setup.sh"
  [ -f "$src/requirements.txt" ] && ln -sfn "$src/requirements.txt" "$target/requirements.txt"
  chmod +x "$src"/scripts/*.sh 2>/dev/null || true
  INSTALLED=$((INSTALLED + 1))

  if [ -x "$src/setup.sh" ]; then
    echo "  Running $name setup..."
    "$src/setup.sh" || echo "  Setup failed for '$name'; re-run $src/setup.sh when ready."
  fi
  # The venv lives beside the source, but SKILL.md paths are rewritten to
  # $target - so it has to be reachable from there too, or every
  # ${CLAUDE_SKILL_DIR}/.venv/bin/python line resolves to nothing.
  [ -d "$src/.venv" ] && ln -sfn "$src/.venv" "$target/.venv"

  # Rewrite ${CLAUDE_SKILL_DIR} to this skill's Codex path. Cross-skill refs of
  # the form ${CLAUDE_SKILL_DIR}/../<other> then resolve to a sibling skill.
  sed "s#\${CLAUDE_SKILL_DIR}#$target#g" \
    "$src/SKILL.md" > "$target/SKILL.md"
done

if [ "$INSTALLED" -eq 0 ]; then
  echo
  echo "Nothing installed - every skill was missing a required dependency."
  exit 1
fi

echo
echo "Installed for Codex. Run scripts/outlook-setup.sh if you have not configured Outlook credentials."
