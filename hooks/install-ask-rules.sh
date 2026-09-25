#!/bin/bash
# Add the outlook ask rules to Claude Code's user settings, with your consent.
#
# An ask rule makes Claude Code stop and ask you before a matching command runs,
# and Claude Code documents it as one of the few things no permission mode
# auto-approves - bypassPermissions included. The rules in ask-rules.json match
# the commands that send something: outlook-mail.sh send, outlook-calendar.sh
# invite, respond and cancel, and the --send-invites and --notify-attendees
# flags. Reading, drafting and filing are left alone.
#
#   hooks/install-ask-rules.sh                  show the rules, then ask
#   hooks/install-ask-rules.sh --yes            add them without asking
#   hooks/install-ask-rules.sh --settings FILE  use another settings file
#
# Only the rules that are missing are added. Everything else in the file is
# kept, and the old file is saved beside it first. Without --yes and without a
# terminal to ask on, nothing is changed.

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
RULES_FILE="$HERE/ask-rules.json"
SETTINGS="${CLAUDE_CONFIG_DIR:-$HOME/.claude}/settings.json"
YES=0

while [ $# -gt 0 ]; do
    case "$1" in
        -y|--yes) YES=1; shift ;;
        --settings)
            [ $# -ge 2 ] || { echo "Error: --settings needs a file" >&2; exit 1; }
            SETTINGS="$2"; shift 2 ;;
        -h|--help) sed -n '2,18p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
        *) echo "Error: unknown option: $1" >&2; exit 1 ;;
    esac
done

command -v jq >/dev/null 2>&1 || { echo "Error: jq is required." >&2; exit 1; }
RULES=$(jq -ce '.permissions.ask | arrays' "$RULES_FILE") \
    || { echo "Error: $RULES_FILE has no permissions.ask list." >&2; exit 1; }

# Write through a symlinked settings file (a dotfiles repo, say) rather than
# replacing the link with a plain file.
TARGET="$SETTINGS"
if [ -L "$SETTINGS" ]; then
    TARGET=$(readlink -f "$SETTINGS" 2>/dev/null) || TARGET="$SETTINGS"
fi

if [ -e "$TARGET" ]; then
    if ! jq -e 'type == "object"
                and ((.permissions // {}) | type == "object")
                and ((.permissions.ask // []) | type == "array")' "$TARGET" >/dev/null 2>&1; then
        echo "Error: $SETTINGS is not a settings object with a permissions.ask list, so it was left alone." >&2
        echo "Fix it, or add the rules by hand from $RULES_FILE." >&2
        exit 1
    fi
    CURRENT=$(cat "$TARGET")
else
    CURRENT='{}'
fi

MISSING=$(printf '%s' "$CURRENT" | jq -c --argjson r "$RULES" '
    (.permissions.ask // []) as $have
    | [$r[] | select(. as $x | $have | any(.[]; . == $x) | not)]')
COUNT=$(printf '%s' "$MISSING" | jq 'length')

if [ "$COUNT" -eq 0 ]; then
    echo "All $(printf '%s' "$RULES" | jq 'length') outlook ask rules are already in $SETTINGS. Nothing to do."
    exit 0
fi

echo "These $COUNT ask rules would be added to permissions.ask in $SETTINGS:"
printf '%s' "$MISSING" | jq -r '.[] | "  " + .'
echo "Claude Code will then ask you before any of these commands runs, in every permission mode."

if [ "$YES" -ne 1 ]; then
    if [ ! -t 0 ]; then
        echo "No terminal to ask on, and no --yes, so nothing was changed."
        exit 0
    fi
    printf 'Add them? [y/N] '
    read -r answer || answer=""
    case "$answer" in
        y|Y|yes|Yes|YES) ;;
        *) echo "Nothing was changed."; exit 0 ;;
    esac
fi

dir=$(dirname "$TARGET")
mkdir -p "$dir"
if [ -e "$TARGET" ]; then
    cp -p "$TARGET" "$TARGET.before-outlook-ask-rules"
fi
tmp=$(mktemp "$dir/.settings.XXXXXX")
trap 'rm -f "$tmp"' EXIT
printf '%s' "$CURRENT" | jq --argjson m "$MISSING" '
    .permissions = ((.permissions // {}) | .ask = ((.ask // []) + $m))' > "$tmp"
mv -f "$tmp" "$TARGET"
trap - EXIT

echo "Added $COUNT ask rules to $SETTINGS."
[ -e "$TARGET.before-outlook-ask-rules" ] && echo "The previous file is at $TARGET.before-outlook-ask-rules."
echo "Remove them any time with /permissions in Claude Code, or by editing the file."
