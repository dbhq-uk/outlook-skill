#!/bin/bash
# Offline tests for tests/chain_marker_live.sh, the live check of whether
# Exchange keeps the reply chain marker in a saved draft.
#
# The live check is run by hand against a real mailbox, so this proves its
# logic against a fake Exchange instead: one that keeps the marker, one that
# strips it, and one that rewrites it. A fake `curl` first on PATH holds the
# draft body between calls, and a throwaway HOME holds a fresh token. Nothing
# here touches ~/.dbhq or the network.
#
#   bash skills/outlook/tests/chain_marker_test.sh
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CHECK="$TESTS_DIR/chain_marker_live.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }

if ! command -v pandoc >/dev/null 2>&1; then
    echo "skip - chain marker check tests (pandoc not installed)"
    exit 0
fi

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# Real Graph IDs are over 100 characters, so they pass resolve_message_id as is.
DRAFT_ID="CHECKDRAFT$(printf 'c%.0s' $(seq 1 110))"

# --- Fake Exchange -------------------------------------------------------------
# Logs "METHOD URL" and each request body. A PATCH saves the body the way
# FAKE_EXCHANGE says: keep it, strip the marker, or rewrite the marker.
mkdir -p "$TMP/bin"
cat > "$TMP/bin/curl" <<'FAKE'
#!/bin/bash
url="" method=GET data="" prev=""
for a in "$@"; do
  case "$prev" in -X) method="$a" ;; -d|--data-binary) data="$a" ;; esac
  case "$a" in https://*) url="$a" ;; esac
  prev="$a"
done
printf '%s %s\n' "$method" "$url" >> "$FAKE_CURL_LOG"
[ -n "$data" ] && printf 'BODY %s\n' "$(printf '%s' "$data" | jq -c .)" >> "$FAKE_CURL_LOG"
path="${url#https://graph.microsoft.com/v1.0}"
marker='<span data-mdreply-chain-start="1"></span>'
case "$method $path" in
  "POST /me/messages")
      [ -n "${FAKE_CREATE_ERROR:-}" ] && { printf '%s' '{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}'; exit 0; }
      printf '%s' "$data" | jq -r '.body.content' > "$FAKE_BODY"
      printf '{"id":"%s"}' "$FAKE_DRAFT_ID" ;;
  "PATCH /me/messages/"*)
      body=$(printf '%s' "$data" | jq -r '.body.content')
      case "$FAKE_EXCHANGE" in
        strip)   body="${body//"$marker"/}" ;;
        rewrite) body="${body//"$marker"/<span data-mdreply-chain-start=\"1\" style=\"\"></span>}" ;;
      esac
      printf '%s' "$body" > "$FAKE_BODY"
      printf '{"id":"%s"}' "$FAKE_DRAFT_ID" ;;
  "GET /me/messages/"*)
      jq -n --arg id "$FAKE_DRAFT_ID" --rawfile b "$FAKE_BODY" '{id: $id, body: {contentType: "html", content: $b}}' ;;
  "POST /me/messages/"*/move) printf '{"id":"%s"}' "$FAKE_DRAFT_ID" ;;
  *) printf '{}' ;;
esac
exit 0
FAKE
chmod +x "$TMP/bin/curl"
export PATH="$TMP/bin:$PATH"
export FAKE_CURL_LOG="$TMP/curl.log" FAKE_BODY="$TMP/body.html" FAKE_DRAFT_ID="$DRAFT_ID"

# --- Fixture HOME with a token that does not need refreshing --------------------
FIX_HOME="$TMP/home"
ACC="$FIX_HOME/.dbhq/outlook/default"
mkdir -p "$ACC"
chmod 700 "$FIX_HOME/.dbhq" "$FIX_HOME/.dbhq/outlook" "$ACC"
printf '%s' '{"client_id":"c"}' > "$ACC/config.json"
printf '{"access_token":"tok","refresh_token":"r","expires_at":%s}' "$(( $(date +%s) + 3600 ))" > "$ACC/credentials.json"
chmod 600 "$ACC/credentials.json"

check() {  # check <keep|strip|rewrite>: runs the live check, prints its output
    : > "$FAKE_CURL_LOG"
    : > "$FAKE_BODY"
    HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default FAKE_EXCHANGE="$1" bash "$CHECK" 2>&1
}
moved_to() {
    grep -A1 "^POST https://graph.microsoft.com/v1.0/me/messages/$DRAFT_ID/move\$" "$FAKE_CURL_LOG" \
        | sed -n 's/^BODY //p' | jq -r '.destinationId'
}
sends() { grep -c '/send$' "$FAKE_CURL_LOG" || true; }
created() { grep -A1 '^POST https://graph.microsoft.com/v1.0/me/messages$' "$FAKE_CURL_LOG" | sed -n 's/^BODY //p'; }

# An Exchange that keeps the marker: the check passes.
out=$(check keep); rc=$?
eq "keep: the check exits 0" "0" "$rc"
has "keep: reports the marker kept after a save" "1. The marker after Exchange saved the draft: kept" "$out"
has "keep: reports the history kept after mdbody" "2. The quoted history after update mdbody: kept" "$out"
has "keep: reports the marker kept after mdbody" "3. The marker after update mdbody: kept" "$out"
eq "keep: the draft has no recipients" "0" "$(created | jq '[.toRecipients, .ccRecipients, .bccRecipients] | map(. // [] | length) | add')"
eq "keep: nothing is sent" "0" "$(sends)"
eq "keep: the draft is moved to Deleted Items" "deleteditems" "$(moved_to)"
eq "keep: the edit went through the real update mdbody" "1" \
   "$(grep '^BODY' "$FAKE_CURL_LOG" | grep -c 'Second version of the reply')"

# An Exchange that strips the empty span: the check fails and says why.
out=$(check strip); rc=$?
eq "strip: the check exits 1" "1" "$rc"
has "strip: reports the marker stripped" "1. The marker after Exchange saved the draft: stripped" "$out"
has "strip: reports the history lost" "2. The quoted history after update mdbody: lost" "$out"
has "strip: says the span was removed" "Exchange removed the span from the saved body." "$out"
eq "strip: the draft is still moved to Deleted Items" "deleteditems" "$(moved_to)"
eq "strip: nothing is sent" "0" "$(sends)"

# An Exchange that keeps the attribute but rewrites the span: mdbody's exact
# match misses it, so this fails too, and shows the rewritten span.
out=$(check rewrite); rc=$?
eq "rewrite: the check exits 1" "1" "$rc"
has "rewrite: reports the marker changed" "1. The marker after Exchange saved the draft: changed" "$out"
has "rewrite: shows the span as saved" 'data-mdreply-chain-start="1" style=""' "$out"
eq "rewrite: the draft is still moved to Deleted Items" "deleteditems" "$(moved_to)"

# It cannot run: read-only mode, or Graph refuses the draft. Exit 2, no writes.
: > "$FAKE_CURL_LOG"
out=$(HOME="$FIX_HOME" OUTLOOK_READ_ONLY=1 FAKE_EXCHANGE=keep bash "$CHECK" 2>&1); rc=$?
eq "read-only: the check exits 2" "2" "$rc"
has "read-only: says why" "OUTLOOK_READ_ONLY is set" "$out"
eq "read-only: no request is made" "0" "$(grep -c . "$FAKE_CURL_LOG" || true)"

out=$(FAKE_CREATE_ERROR=1 check keep); rc=$?
eq "create refused: the check exits 2" "2" "$rc"
has "create refused: passes on Graph's message" "Access is denied." "$out"
eq "create refused: nothing to move" "" "$(moved_to)"

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
