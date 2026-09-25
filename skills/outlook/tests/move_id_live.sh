#!/bin/bash
# Live check: does a message keep its ID when it moves to another folder?
#
# Every request the scripts make asks Graph for immutable IDs (lib/graph.sh).
# Those should keep their value when an item moves, so an ID from a listing
# still works after `move`, `batch-move` or `delete`. The offline suites prove
# the header is sent. Only a real mailbox can say what Graph does with it.
#
# It needs a configured account. It writes to that mailbox, and sends nothing:
#   1. it creates a draft with no recipients, so the draft cannot be sent;
#   2. it moves the draft to Deleted Items with the real `outlook-mail.sh
#      delete`;
#   3. it reads the draft by the ID it had BEFORE the move, with the real
#      `outlook-mail.sh read`, and checks that it is now in Deleted Items.
# The draft stays in Deleted Items, where `delete` put it. CI never runs this:
# the name does not end in _test.sh.
#
#   bash skills/outlook/tests/move_id_live.sh [--account <name>]
#
# Exit 0: the ID survived the move.
# Exit 1: it did not, so an ID goes stale when its message moves.
# Exit 2: the check could not run (no account, read-only mode, or a Graph
#         error).
set -u

SCRIPTS="$(cd "$(dirname "${BASH_SOURCE[0]}")/../scripts" && pwd)"
# shellcheck source=../scripts/lib/graph.sh
. "$SCRIPTS/lib/graph.sh"

if [ "${1:-}" = "--account" ] || [ "${1:-}" = "-a" ]; then
    [ -n "${2:-}" ] || { echo "Error: $1 requires an account name" >&2; exit 2; }
    export OUTLOOK_ACCOUNT="$2"
fi

SUBJECT="outlook-skill move ID check (safe to delete)"

cannot_run() { echo "Cannot run the check: $1" >&2; exit 2; }

if outlook_read_only; then
    cannot_run "OUTLOOK_READ_ONLY is set, and this check writes a draft."
fi
# outlook-token.sh prints only the token on success, and its error otherwise.
if ! TOKEN=$(bash "$SCRIPTS/outlook-token.sh" get); then
    cannot_run "no token for account '${OUTLOOK_ACCOUNT:-default}': $TOKEN"
fi

graph() {  # graph <METHOD> <path> [json-body], asking for immutable IDs as the scripts do
    local args=(-s --connect-timeout 10 --max-time 60 -X "$1" -H "Authorization: Bearer $TOKEN" -H "$OUTLOOK_PREFER_IDS")
    [ $# -ge 3 ] && args+=(-H "Content-Type: application/json" -d "$3")
    curl "${args[@]}" "https://graph.microsoft.com/v1.0$2"
}
graph_error() { printf '%s' "$1" | jq -r '.error.message // .error.code // empty' 2>/dev/null; }

# 1. A draft with no recipients.
created=$(graph POST /me/messages \
    "$(jq -n --arg s "$SUBJECT" '{subject: $s, body: {contentType: "Text", content: "Checks that a message keeps its ID when it moves."}}')")
DRAFT=$(printf '%s' "$created" | jq -r '.id // empty' 2>/dev/null)
[ -n "$DRAFT" ] || cannot_run "Graph did not create the draft: $(graph_error "$created")"
echo "Created a draft with no recipients. It cannot be sent."

# 2. The real delete: a move to Deleted Items.
if ! out=$(bash "$SCRIPTS/outlook-mail.sh" delete "$DRAFT" 2>&1); then
    echo "The draft '$SUBJECT' is still in Drafts. Delete it by hand." >&2
    cannot_run "outlook-mail.sh delete failed: $out"
fi
echo "Moved the draft to Deleted Items with outlook-mail.sh delete."

# 3. The real read, by the ID from before the move, and where the draft is now.
deleted_folder=$(graph GET "/me/mailFolders/deleteditems?\$select=id" | jq -r '.id // empty' 2>/dev/null)
now_in=$(graph GET "/me/messages/$DRAFT?\$select=parentFolderId" | jq -r '.parentFolderId // empty' 2>/dev/null)
if out=$(bash "$SCRIPTS/outlook-mail.sh" read "$DRAFT" 2>&1) && [[ "$out" == *"$SUBJECT"* ]]; then
    read_ok=yes
else
    read_ok=no
fi
echo "1. outlook-mail.sh read, by the ID from before the move: $([ "$read_ok" = yes ] && echo found || echo 'not found')"

if [ "$read_ok" = yes ] && [ -n "$deleted_folder" ] && [ "$now_in" = "$deleted_folder" ]; then
    echo "2. The draft is in Deleted Items: yes"
    echo "Result: the ID survived the move. IDs from a listing stay valid after move, batch-move and delete."
    exit 0
fi
if [ "$read_ok" = yes ]; then
    echo "2. The draft is in Deleted Items: no"
    echo "Result: the old ID still reads, but the draft is not in Deleted Items, so the move did not happen as expected."
else
    echo "Result: the ID changed when the draft moved. Graph did not return an immutable ID, so an ID from a listing goes stale after a move."
fi
echo "The draft '$SUBJECT' is in Deleted Items or Drafts. Delete it by hand."
exit 1
