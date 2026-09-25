#!/bin/bash
# Live check: does Exchange keep the reply chain marker in a saved draft?
#
# mdreply, followup and forward put an empty span between the new message and
# the quoted history:
#
#     <span data-mdreply-chain-start="1"></span>
#
# `update mdbody` keeps everything from that span onwards. If Exchange strips
# the span when it saves a draft, the next mdbody edit replaces the whole body
# and the quoted history is lost without a word. The offline suites cannot
# tell, because they answer Graph from fixtures. This asks the real server.
#
# It needs a configured account and pandoc. It writes to that mailbox, and
# nothing else:
#   1. it creates a draft with no recipients, so the draft cannot be sent;
#   2. it writes a body with the marker, as mdreply does, and reads it back;
#   3. it runs the real `outlook-mail.sh update <draft> mdbody`, and reads the
#      body back again;
#   4. it moves the draft to Deleted Items, whatever happened.
# Nothing is sent. CI never runs it: the name does not end in _test.sh.
#
#   bash skills/outlook/tests/chain_marker_live.sh [--account <name>]
#
# Exit 0: Exchange kept the marker, and mdbody kept the quoted history.
# Exit 1: it did not. The output says which step lost it.
# Exit 2: the check could not run (no account, no pandoc, read-only mode, or a
#         Graph error).
set -u

SCRIPTS="$(cd "$(dirname "${BASH_SOURCE[0]}")/../scripts" && pwd)"
# shellcheck source=../scripts/lib/graph.sh
. "$SCRIPTS/lib/graph.sh"

if [ "${1:-}" = "--account" ] || [ "${1:-}" = "-a" ]; then
    [ -n "${2:-}" ] || { echo "Error: $1 requires an account name" >&2; exit 2; }
    export OUTLOOK_ACCOUNT="$2"
fi
# The From on a test draft does not matter, and must not fail the check.
export OUTLOOK_FROM_ADDRESS=""

MARKER='<span data-mdreply-chain-start="1"></span>'
HISTORY='chain-marker-check: quoted history'

cannot_run() { echo "Cannot run the check: $1" >&2; exit 2; }

if outlook_read_only; then
    cannot_run "OUTLOOK_READ_ONLY is set, and this check writes a draft."
fi
command -v pandoc >/dev/null 2>&1 || cannot_run "update mdbody needs pandoc."
# outlook-token.sh prints only the token on success, and its error otherwise.
if ! TOKEN=$(bash "$SCRIPTS/outlook-token.sh" get); then
    cannot_run "no token for account '${OUTLOOK_ACCOUNT:-default}': $TOKEN"
fi

graph() {  # graph <METHOD> <path> [json-body]
    local args=(-s --connect-timeout 10 --max-time 60 -X "$1" -H "Authorization: Bearer $TOKEN")
    [ $# -ge 3 ] && args+=(-H "Content-Type: application/json" -d "$3")
    curl "${args[@]}" "https://graph.microsoft.com/v1.0$2"
}
graph_error() { printf '%s' "$1" | jq -r '.error.message // .error.code // empty' 2>/dev/null; }
saved_body() { graph GET "/me/messages/$DRAFT?\$select=body" | jq -r '.body.content // empty'; }

# What happened to the marker in a saved body: kept, changed or stripped.
# "changed" means the attribute survived but the span is no longer the exact
# string mdbody searches for, which loses the history just the same.
marker_state() {
    if [[ "$1" == *"$MARKER"* ]]; then echo kept
    elif [[ "$1" == *data-mdreply-chain-start* ]]; then echo changed
    else echo stripped
    fi
}

# 1. A draft with no recipients.
created=$(graph POST /me/messages \
    "$(jq -n '{subject: "outlook-skill chain marker check (safe to delete)", body: {contentType: "HTML", content: "<p>placeholder</p>"}}')")
DRAFT=$(printf '%s' "$created" | jq -r '.id // empty' 2>/dev/null)
[ -n "$DRAFT" ] || cannot_run "Graph did not create the draft: $(graph_error "$created")"

cleanup() {
    if bash "$SCRIPTS/outlook-mail.sh" delete "$DRAFT" >/dev/null 2>&1; then
        echo "Moved the check draft to Deleted Items."
    else
        echo "Could not move the check draft to Deleted Items. Delete the draft" >&2
        echo "'outlook-skill chain marker check (safe to delete)' by hand." >&2
    fi
}
trap cleanup EXIT
echo "Created a draft with no recipients. It cannot be sent."

# 2. The body mdreply writes: new text, the marker, then the quoted history.
first="<p>First version of the reply.</p>
<br/>
${MARKER}
<div><p>${HISTORY}</p></div>"
patched=$(graph PATCH "/me/messages/$DRAFT" "$(jq -n --arg b "$first" '{body: {contentType: "HTML", content: $b}}')")
err=$(graph_error "$patched")
[ -z "$err" ] || cannot_run "Graph did not save the body: $err"

saved=$(saved_body)
state=$(marker_state "$saved")
echo "1. The marker after Exchange saved the draft: $state"

# 3. The real mdbody edit.
if ! out=$(bash "$SCRIPTS/outlook-mail.sh" update "$DRAFT" mdbody "Second version of the reply." 2>&1); then
    cannot_run "update mdbody failed: $out"
fi
after=$(saved_body)
if [[ "$after" == *"$HISTORY"* ]]; then history=kept; else history=lost; fi
echo "2. The quoted history after update mdbody: $history"
after_state=$(marker_state "$after")
echo "3. The marker after update mdbody: $after_state"

if [ "$state" = kept ] && [ "$history" = kept ] && [ "$after_state" = kept ]; then
    echo "Result: Exchange keeps the chain marker, and update mdbody keeps the quoted history."
    exit 0
fi
echo "Result: the quoted history is NOT safe. update mdbody would lose it on a reply."
case "$state" in
    changed) printf 'Exchange rewrote the span. The saved body around it: %s\n' \
                 "$(printf '%s' "$saved" | tr '\n' ' ' | grep -o '.\{0,60\}data-mdreply-chain-start.\{0,60\}' | head -1)" ;;
    stripped) echo "Exchange removed the span from the saved body." ;;
esac
exit 1
