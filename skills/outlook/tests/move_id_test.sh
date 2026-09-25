#!/bin/bash
# Offline tests for tests/move_id_live.sh, the live check that a message keeps
# its ID when it moves.
#
# The live check is run by hand against a real mailbox, so this proves its
# logic against a fake Exchange that behaves like Graph: a move gives the
# message a new ID and the old one stops working, unless the request asked for
# immutable IDs. A second fake ignores the header and always changes the ID.
# A fake `curl` first on PATH holds the message between calls, and a throwaway
# HOME holds a fresh token. Nothing here touches ~/.dbhq or the network.
#
#   bash skills/outlook/tests/move_id_test.sh
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CHECK="$TESTS_DIR/move_id_live.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# --- Fake Exchange -------------------------------------------------------------
# State: the message's current ID, its folder and its subject, one per file.
# FAKE_IDS=immutable honours Prefer: IdType="ImmutableId"; FAKE_IDS=mutable
# ignores it. Real Graph IDs are over 100 characters, so each fake ID is too.
mkdir -p "$TMP/bin"
cat > "$TMP/bin/curl" <<'FAKE'
#!/bin/bash
url="" method=GET data="" prev="" immutable=0
for a in "$@"; do
  case "$prev" in
    -X) method="$a" ;;
    -d|--data-binary) data="$a" ;;
    -H) [ "$a" = 'Prefer: IdType="ImmutableId"' ] && immutable=1 ;;
  esac
  case "$a" in https://*) url="$a" ;; esac
  prev="$a"
done
[ "$FAKE_IDS" = mutable ] && immutable=0
printf '%s %s immutable=%s\n' "$method" "$url" "$immutable" >> "$FAKE_CURL_LOG"
[ -n "$data" ] && printf 'BODY %s\n' "$(printf '%s' "$data" | jq -c .)" >> "$FAKE_CURL_LOG"
path="${url#https://graph.microsoft.com/v1.0}"
S="$FAKE_STATE"
long() { printf '%s%s' "$1" "$(printf 'x%.0s' $(seq 1 110))"; }
not_found() { printf '%s' '{"error":{"code":"ErrorItemNotFound","message":"The specified object was not found in the store."}}'; }
current=$(cat "$S/id" 2>/dev/null)
case "$method $path" in
  "POST /me/messages")
      long "MSG1-" > "$S/id"; echo DRAFTS-ID > "$S/folder"
      printf '%s' "$data" | jq -r '.subject' > "$S/subject"
      jq -n --arg id "$(cat "$S/id")" '{id: $id}' ;;
  "POST /me/messages/"*/move)
      id="${path#/me/messages/}"; id="${id%/move}"
      [ "$id" = "$current" ] || { not_found; exit 0; }
      dest=$(printf '%s' "$data" | jq -r '.destinationId')
      [ "$dest" = deleteditems ] && dest=DELETEDITEMS-ID
      echo "$dest" > "$S/folder"
      # A move gives the message a new ID unless the request asked for an
      # immutable one.
      [ "$immutable" = 1 ] || long "MSG2-" > "$S/id"
      jq -n --arg id "$(cat "$S/id")" '{id: $id}' ;;
  "GET /me/mailFolders/deleteditems"*) printf '%s' '{"id":"DELETEDITEMS-ID"}' ;;
  "GET /me/messages/"*)
      id="${path#/me/messages/}"; id="${id%%\?*}"
      [ "$id" = "$current" ] || { not_found; exit 0; }
      jq -n --arg id "$current" --arg f "$(cat "$S/folder")" --arg s "$(cat "$S/subject")" \
          '{id: $id, parentFolderId: $f, subject: $s, from: {emailAddress: {address: "me@example.com"}},
            toRecipients: [], receivedDateTime: "2026-09-25T10:00:00Z", body: {contentType: "text", content: "x"}, attachments: []}' ;;
  *) printf '{}' ;;
esac
exit 0
FAKE
chmod +x "$TMP/bin/curl"
export PATH="$TMP/bin:$PATH"
export FAKE_CURL_LOG="$TMP/curl.log"

# --- Fixture HOME with a token that does not need refreshing --------------------
FIX_HOME="$TMP/home"
ACC="$FIX_HOME/.dbhq/outlook/default"
mkdir -p "$ACC"
chmod 700 "$FIX_HOME/.dbhq" "$FIX_HOME/.dbhq/outlook" "$ACC"
printf '%s' '{"client_id":"c"}' > "$ACC/config.json"
printf '{"access_token":"tok","refresh_token":"r","expires_at":%s}' "$(( $(date +%s) + 3600 ))" > "$ACC/credentials.json"
chmod 600 "$ACC/credentials.json"

n=0
check() {  # check <immutable|mutable>: runs the live check, prints its output
    n=$((n + 1))
    export FAKE_STATE="$TMP/state.$n"
    mkdir -p "$FAKE_STATE"
    : > "$FAKE_CURL_LOG"
    HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default FAKE_IDS="$1" bash "$CHECK" 2>&1
}
created() { grep -A1 '^POST https://graph.microsoft.com/v1.0/me/messages ' "$FAKE_CURL_LOG" | sed -n 's/^BODY //p'; }
sends() { grep -c '/send ' "$FAKE_CURL_LOG" || true; }

# Graph honours the header: the ID survives the move and the check passes.
out=$(check immutable); rc=$?
eq "immutable: the check exits 0" "0" "$rc"
has "immutable: the old ID still reads" "1. outlook-mail.sh read, by the ID from before the move: found" "$out"
has "immutable: the draft is in Deleted Items" "2. The draft is in Deleted Items: yes" "$out"
eq "immutable: the draft was moved to Deleted Items by the real delete" "deleteditems" \
   "$(grep -A1 '/move immutable=' "$FAKE_CURL_LOG" | sed -n 's/^BODY //p' | jq -r '.destinationId')"
eq "immutable: the draft has no recipients" "0" \
   "$(created | jq '[.toRecipients, .ccRecipients, .bccRecipients] | map(. // [] | length) | add')"
eq "immutable: nothing is sent" "0" "$(sends)"
eq "immutable: every request asked for immutable IDs" "0" "$(grep -c 'immutable=0$' "$FAKE_CURL_LOG" || true)"

# Graph ignores the header: the ID changes, and the check fails and says so.
out=$(check mutable); rc=$?
eq "mutable: the check exits 1" "1" "$rc"
has "mutable: the old ID no longer reads" "1. outlook-mail.sh read, by the ID from before the move: not found" "$out"
has "mutable: says the ID changed" "the ID changed when the draft moved" "$out"
eq "mutable: nothing is sent" "0" "$(sends)"

# It cannot run in read-only mode: exit 2 and no request at all.
: > "$FAKE_CURL_LOG"
out=$(HOME="$FIX_HOME" OUTLOOK_READ_ONLY=1 FAKE_IDS=immutable bash "$CHECK" 2>&1); rc=$?
eq "read-only: the check exits 2" "2" "$rc"
has "read-only: says why" "OUTLOOK_READ_ONLY is set" "$out"
eq "read-only: no request is made" "0" "$(grep -c . "$FAKE_CURL_LOG" || true)"

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
