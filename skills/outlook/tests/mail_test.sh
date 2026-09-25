#!/bin/bash
# Offline tests for the draft and send verbs of outlook-mail.sh, run as a user
# would run them.
#
# A fake `curl` first on PATH answers Graph from fixtures and logs every request
# with its body, and a throwaway HOME holds a fresh token. Nothing is sent to
# anyone and nothing here touches ~/.dbhq.
#
#   bash skills/outlook/tests/mail_test.sh
#
# What this pins: OUTLOOK_FROM_ADDRESS reaches the drafts that reply, mdreply,
# followup and forward create, and `send` shows From, To, Cc, Bcc, Subject and
# the attachments before it posts, and posts nothing if it cannot read them.
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
MAIL="$(cd "$TESTS_DIR/../scripts" && pwd)/outlook-mail.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# IDs over 100 characters pass through resolve_message_id untouched.
MSG_ID="MSG$(printf 'm%.0s' $(seq 1 110))"
DRAFT_ID="DRAFT$(printf 'd%.0s' $(seq 1 110))"

# --- Fake curl -----------------------------------------------------------------
# Logs "METHOD URL" then the request body (if any) on the next line, prefixed
# "BODY ". A PATCH answers with the draft plus whatever was patched, as Graph
# does, or with $FAKE_PATCH_ERROR when that is set.
mkdir -p "$TMP/bin"
cat > "$TMP/bin/curl" <<'FAKE'
#!/bin/bash
url="" method=GET data="" prev=""
for a in "$@"; do
  case "$prev" in -X) method="$a" ;; -d) data="$a" ;; esac
  case "$a" in https://*) url="$a" ;; esac
  prev="$a"
done
printf '%s %s\n' "$method" "$url" >> "$FAKE_CURL_LOG"
[ -n "$data" ] && printf 'BODY %s\n' "$(printf '%s' "$data" | jq -c .)" >> "$FAKE_CURL_LOG"
path="${url#https://graph.microsoft.com/v1.0}"
case "$method $path" in
  "GET /me?"*)              printf '%s' '{"mail":"dan@example.com","proxyAddresses":["SMTP:dan@example.com","smtp:alias@example.com"]}' ;;
  "POST "*/createReplyAll)  cat "$FAKE_DRAFT" ;;
  "POST "*/createForward)   cat "$FAKE_DRAFT" ;;
  "PATCH /me/messages/"*)
      if [ -n "${FAKE_PATCH_ERROR:-}" ]; then cat "$FAKE_PATCH_ERROR"; else jq -c --argjson d "$data" '. + $d' "$FAKE_DRAFT"; fi ;;
  "GET /me/messages/"*/attachments*) cat "$FAKE_ATTACHMENTS" ;;
  "GET /me/messages/"*)     cat "$FAKE_DRAFT_READ" ;;
  "POST /me/messages/"*/send) : ;;
  *) printf '{}' ;;
esac
exit 0
FAKE
chmod +x "$TMP/bin/curl"
export PATH="$TMP/bin:$PATH"
export FAKE_CURL_LOG="$TMP/curl.log" FAKE_DRAFT="$TMP/draft.json" \
       FAKE_DRAFT_READ="$TMP/draft_read.json" FAKE_ATTACHMENTS="$TMP/attachments.json"

printf '{"id":"%s","subject":"RE: Hello","isDraft":true,"body":{"contentType":"HTML","content":"<p>quoted chain</p>"},"toRecipients":[{"emailAddress":{"address":"alice@example.com"}}],"ccRecipients":[{"emailAddress":{"address":"carol@example.com"}}],"bccRecipients":[],"from":null}' \
    "$DRAFT_ID" > "$FAKE_DRAFT"
cp "$FAKE_DRAFT" "$FAKE_DRAFT_READ"
printf '%s' '{"value":[{"name":"report.pdf","isInline":false},{"name":"logo.png","isInline":true}]}' > "$FAKE_ATTACHMENTS"

# --- Fixture HOME with a token that does not need refreshing --------------------
FIX_HOME="$TMP/home"
ACC="$FIX_HOME/.dbhq/outlook/default"
mkdir -p "$ACC"
chmod 700 "$FIX_HOME/.dbhq" "$FIX_HOME/.dbhq/outlook" "$ACC"
printf '%s' '{"client_id":"c","client_secret":"s"}' > "$ACC/config.json"
printf '{"access_token":"tok","refresh_token":"r","expires_at":%s}' "$(( $(date +%s) + 3600 ))" > "$ACC/credentials.json"
chmod 600 "$ACC/credentials.json"

mail() {  # mail <from-address or ""> <verb> [args...]
    local from="$1"; shift
    : > "$FAKE_CURL_LOG"
    HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default OUTLOOK_FROM_ADDRESS="$from" bash "$MAIL" "$@" 2>&1
}
# The `from` address in the body of the PATCH to the draft, or "none".
patched_from() {
    grep -A1 "^PATCH https://graph.microsoft.com/v1.0/me/messages/$DRAFT_ID\$" "$FAKE_CURL_LOG" \
        | sed -n 's/^BODY //p' | jq -r '.from.emailAddress.address // "none"' | tail -1
}
patches() { grep -c "^PATCH https://graph.microsoft.com/v1.0/me/messages/$DRAFT_ID\$" "$FAKE_CURL_LOG" || true; }

########################################
# OUTLOOK_FROM_ADDRESS reaches every reply-family draft.
########################################
if command -v pandoc >/dev/null 2>&1; then
    out=$(mail alias@example.com mdreply "$MSG_ID" "**Thanks**")
    eq "mdreply PATCHes from = OUTLOOK_FROM_ADDRESS" "alias@example.com" "$(patched_from)"
    has "mdreply prints the From it set" "From:    alias@example.com" "$out"
    eq "mdreply still keeps the quoted chain" "1" \
       "$(grep '^BODY' "$FAKE_CURL_LOG" | grep -c 'data-mdreply-chain-start')"

    out=$(mail alias@example.com followup "$MSG_ID" "Any news?")
    eq "followup PATCHes from = OUTLOOK_FROM_ADDRESS" "alias@example.com" "$(patched_from)"

    out=$(mail alias@example.com forward "$MSG_ID" "bob@example.com" "FYI")
    eq "forward with a comment PATCHes from = OUTLOOK_FROM_ADDRESS" "alias@example.com" "$(patched_from)"

    out=$(mail "" mdreply "$MSG_ID" "**Thanks**")
    eq "mdreply without OUTLOOK_FROM_ADDRESS sets no from" "none" "$(patched_from)"
    has "mdreply without OUTLOOK_FROM_ADDRESS says the mailbox default is used" "From:    (mailbox default)" "$out"
else
    echo "skip - mdreply/followup/forward-with-comment tests (pandoc not installed)"
fi

out=$(mail alias@example.com reply "$MSG_ID" "Thanks")
eq "reply PATCHes from = OUTLOOK_FROM_ADDRESS" "alias@example.com" "$(patched_from)"
has "reply prints the From it set" "From:    alias@example.com" "$out"

out=$(mail alias@example.com forward "$MSG_ID" "bob@example.com")
eq "forward without a comment PATCHes from = OUTLOOK_FROM_ADDRESS" "alias@example.com" "$(patched_from)"

out=$(mail "" reply "$MSG_ID" "Thanks")
eq "reply without OUTLOOK_FROM_ADDRESS makes no PATCH" "0" "$(patches)"

# If the From cannot be set, say so and exit non-zero rather than leave a draft
# that looks ready to go from the wrong address.
printf '%s' '{"error":{"code":"ErrorSendAsDenied","message":"Not allowed to send as that address."}}' > "$TMP/patch_error.json"
out=$(FAKE_PATCH_ERROR="$TMP/patch_error.json" mail alias@example.com reply "$MSG_ID" "Thanks"); rc=$?
eq "reply exits non-zero when the From cannot be set" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "reply says the From was not set" "its From could not be set to alias@example.com" "$out"

########################################
# send shows what is about to go before it posts.
########################################
printf '{"id":"%s","subject":"RE: Hello","isDraft":true,"from":{"emailAddress":{"address":"alias@example.com"}},"toRecipients":[{"emailAddress":{"address":"alice@example.com"}}],"ccRecipients":[{"emailAddress":{"address":"carol@example.com"}}],"bccRecipients":[{"emailAddress":{"address":"audit@example.com"}}]}' \
    "$DRAFT_ID" > "$FAKE_DRAFT_READ"
out=$(mail alias@example.com send "$DRAFT_ID")
has "send prints From"        "From:        alias@example.com" "$out"
has "send prints To"          "To:          alice@example.com" "$out"
has "send prints Cc"          "Cc:          carol@example.com" "$out"
has "send prints Bcc"         "Bcc:         audit@example.com" "$out"
has "send prints Subject"     "Subject:     RE: Hello" "$out"
has "send prints attachments" "Attachments: report.pdf, logo.png (inline)" "$out"
eq "send prints the summary before it sends" "1" \
   "$(printf '%s\n' "$out" | awk '/From:        alias/{f=NR} /^Sending/{s=NR} END{print (f && s && f < s) ? 1 : 0}')"
eq "send reads the draft before posting it" "1" \
   "$(awk -v d="$DRAFT_ID" '$0 ~ "^GET .*/me/messages/"d"\\?" {g=NR} $0 ~ "^POST .*/me/messages/"d"/send$" {p=NR} END{print (g && p && g < p) ? 1 : 0}' "$FAKE_CURL_LOG")"
eq "send posts once" "1" "$(grep -c "^POST .*/me/messages/$DRAFT_ID/send\$" "$FAKE_CURL_LOG")"

# A draft with no from goes out from the primary address, and send says which.
printf '{"id":"%s","subject":"Hi","isDraft":true,"from":null,"toRecipients":[{"emailAddress":{"address":"alice@example.com"}}],"ccRecipients":[],"bccRecipients":[]}' \
    "$DRAFT_ID" > "$FAKE_DRAFT_READ"
out=$(mail "" send "$DRAFT_ID")
has "send names the primary address when from is unset" "From:        dan@example.com (mailbox default)" "$out"
has "send prints (none) for an empty Cc" "Cc:          (none)" "$out"

# The From that will go out differs from OUTLOOK_FROM_ADDRESS: warn.
out=$(mail alias@example.com send "$DRAFT_ID")
has "send warns when the draft's From is not OUTLOOK_FROM_ADDRESS" "but this draft sends from dan@example.com" "$out"

# If the draft cannot be read, nothing is sent.
printf '%s' '{"error":{"code":"ErrorItemNotFound","message":"The specified object was not found."}}' > "$FAKE_DRAFT_READ"
out=$(mail "" send "$DRAFT_ID"); rc=$?
eq "send exits non-zero when the draft cannot be read" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "send posts nothing when the draft cannot be read" "0" "$(grep -c '/send$' "$FAKE_CURL_LOG" || true)"

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
