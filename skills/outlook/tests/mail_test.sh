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
# `delete` and `rmdir` move to Deleted Items and never send a DELETE.
# `attach --inline` and `signature` upload inline images, and `update mdbody`
# keeps a signature block. `draft` and `mddraft` take recipient lists and
# --cc/--bcc. reply, mdreply and followup create a reply-all draft and print
# every To, Cc and Bcc. The reply verbs put the chain marker between the new
# text and the quoted history, and `update mdbody` keeps everything from the
# first marker onwards.
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
  case "$prev" in -X) method="$a" ;; -d|--data-binary) data="$a" ;; esac
  case "$a" in https://*) url="$a" ;; esac
  prev="$a"
done
printf '%s %s\n' "$method" "$url" >> "$FAKE_CURL_LOG"
# A body sent as @file (the attachment upload) is read from the file.
case "$data" in @*) data=$(cat "${data#@}") ;; esac
[ -n "$data" ] && printf 'BODY %s\n' "$(printf '%s' "$data" | jq -c .)" >> "$FAKE_CURL_LOG"
path="${url#https://graph.microsoft.com/v1.0}"
case "$method $path" in
  "GET /me?"*)              printf '%s' '{"mail":"dan@example.com","proxyAddresses":["SMTP:dan@example.com","smtp:alias@example.com"]}' ;;
  "POST /me/messages")      printf '%s' "$data" | jq -c '. + {id: "NEWDRAFTxxxxxxxxxxxxxxxxxxxxxx"}' ;;
  "POST "*/createReplyAll)  cat "$FAKE_DRAFT" ;;
  "POST "*/createForward)   cat "$FAKE_DRAFT" ;;
  "PATCH /me/messages/"*)
      if [ -n "${FAKE_PATCH_ERROR:-}" ]; then cat "$FAKE_PATCH_ERROR"; else jq -c --argjson d "$data" '. + $d' "$FAKE_DRAFT"; fi ;;
  "GET /me/mailFolders?"*)  printf '%s' '{"value":[{"id":"FOLDER1","displayName":"Old Projects"}]}' ;;
  "GET /me/mailFolders/FOLDER1?"*) printf '{"id":"FOLDER1","totalItemCount":%s}' "${FAKE_FOLDER_COUNT:-0}" ;;
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

########################################
# delete and rmdir: a move to Deleted Items, never a DELETE
########################################
# A Graph DELETE on a message skips Deleted Items and lands in Recoverable
# Items, where the user will not look for it.
out=$(mail "" delete "$MSG_ID"); rc=$?
eq "delete exits 0" "0" "$rc"
eq "delete posts one move" "1" "$(grep -c "^POST https://graph.microsoft.com/v1.0/me/messages/$MSG_ID/move\$" "$FAKE_CURL_LOG")"
eq "delete moves to deleteditems" "deleteditems" \
   "$(grep -A1 "/me/messages/$MSG_ID/move\$" "$FAKE_CURL_LOG" | sed -n 's/^BODY //p' | jq -r '.destinationId')"
eq "delete sends no DELETE" "0" "$(grep -c '^DELETE ' "$FAKE_CURL_LOG" || true)"
has "delete says where the message went" "Moved to Deleted Items" "$out"

out=$(FAKE_FOLDER_COUNT=3 mail "" rmdir "Old Projects"); rc=$?
eq "rmdir of a folder with messages and no --force exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "rmdir without --force changes nothing" "0" "$(grep -c '^\(POST\|DELETE\) ' "$FAKE_CURL_LOG" || true)"
has "rmdir without --force says what --force does" "move it, and everything in it, to Deleted Items" "$out"

out=$(FAKE_FOLDER_COUNT=3 mail "" rmdir "Old Projects" --force); rc=$?
eq "rmdir --force exits 0" "0" "$rc"
eq "rmdir --force moves the folder to deleteditems" "deleteditems" \
   "$(grep -A1 '^POST https://graph.microsoft.com/v1.0/me/mailFolders/FOLDER1/move$' "$FAKE_CURL_LOG" | sed -n 's/^BODY //p' | jq -r '.destinationId')"
eq "rmdir --force sends no DELETE" "0" "$(grep -c '^DELETE ' "$FAKE_CURL_LOG" || true)"
has "rmdir says where the folder went" "Moved folder 'Old Projects' to Deleted Items" "$out"

out=$(mail "" rmdir "Old Projects"); rc=$?
eq "rmdir of an empty folder needs no --force" "1" "$(grep -c '^POST .*/mailFolders/FOLDER1/move$' "$FAKE_CURL_LOG")"
eq "rmdir of an empty folder sends no DELETE" "0" "$(grep -c '^DELETE ' "$FAKE_CURL_LOG" || true)"

########################################
# draft and mddraft: recipient lists, --cc and --bcc
########################################
new_draft() {  # the JSON body of the POST that created the draft
    grep -A1 '^POST https://graph.microsoft.com/v1.0/me/messages$' "$FAKE_CURL_LOG" | sed -n 's/^BODY //p'
}
addrs() { new_draft | jq -r "[.$1[]?.emailAddress.address] | join(\",\")"; }

out=$(mail "" draft "a@example.com; b@example.com" "Hi" "Body"); rc=$?
eq "draft with two addresses exits 0" "0" "$rc"
eq "draft splits a semicolon list into two To recipients" "a@example.com,b@example.com" "$(addrs toRecipients)"
eq "draft without --cc sends no Cc" "false" "$(new_draft | jq 'has("ccRecipients")')"
has "draft prints every To recipient" "To: a@example.com, b@example.com" "$out"
has "draft prints an empty Cc as (none)" "Cc: (none)" "$out"

out=$(mail "" draft "a@example.com, b@example.com" "Hi" "Body" --cc "c@example.com;d@example.com" --bcc "audit@example.com")
eq "draft splits a comma list too" "a@example.com,b@example.com" "$(addrs toRecipients)"
eq "draft --cc sets Cc" "c@example.com,d@example.com" "$(addrs ccRecipients)"
eq "draft --bcc sets Bcc" "audit@example.com" "$(addrs bccRecipients)"
eq "the options never land in the body or subject" "Hi|Body" "$(new_draft | jq -r '"\(.subject)|\(.body.content)"')"
has "draft prints Cc" "Cc: c@example.com, d@example.com" "$out"
has "draft prints Bcc" "Bcc: audit@example.com" "$out"

out=$(mail "" mddraft --cc "c@example.com" "a@example.com;b@example.com" "Hi" "**Body**")
eq "mddraft splits a list into two To recipients" "a@example.com,b@example.com" "$(addrs toRecipients)"
eq "mddraft --cc before the arguments works" "c@example.com" "$(addrs ccRecipients)"
has "mddraft still converts the markdown" "<strong>Body</strong>" "$(new_draft | jq -r '.body.content')"

out=$(mail "" draft " ; , " "Hi" "Body"); rc=$?
eq "a To list with no address is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "a refused draft creates nothing" "" "$(new_draft)"
out=$(mail "" draft "a@example.com" "Hi" "Body" --cc); rc=$?
eq "--cc with no list is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "--cc with no list creates nothing" "" "$(new_draft)"

########################################
# attach --inline and signature
########################################
IMG="$TMP/img"
mkdir -p "$IMG/icons"
printf 'PNGDATA-logo' > "$IMG/logo.png"
printf 'PNGDATA-phone' > "$IMG/icons/phone.png"
b64() { base64 < "$1" | tr -d '\n'; }
posted_attachments() {  # the JSON body of every attachment POST, one per line
    grep -A1 "^POST https://graph.microsoft.com/v1.0/me/messages/$DRAFT_ID/attachments\$" "$FAKE_CURL_LOG" | sed -n 's/^BODY //p'
}
patched_body() {  # the body of the last PATCH, on one line
    grep -A1 "^PATCH https://graph.microsoft.com/v1.0/me/messages/$DRAFT_ID\$" "$FAKE_CURL_LOG" | sed -n 's/^BODY //p' \
        | tail -1 | jq -r '.body.content' | tr '\n' ' '
}
printf '%s' '{"value":[]}' > "$FAKE_ATTACHMENTS"

out=$(mail "" attach "$DRAFT_ID" "$IMG/logo.png" --inline logo); rc=$?
eq "attach --inline exits 0" "0" "$rc"
eq "attach --inline posts isInline, the content ID and the bytes" "true|logo|logo.png|$(b64 "$IMG/logo.png")" \
   "$(posted_attachments | jq -r '"\(.isInline)|\(.contentId)|\(.name)|\(.contentBytes)"')"
has "attach --inline says it is inline" "(inline, cid:logo)" "$out"

out=$(mail "" attach "$DRAFT_ID" --inline logo "$IMG/logo.png")
eq "--inline may come before the file" "logo" "$(posted_attachments | jq -r '.contentId')"

out=$(mail "" attach "$DRAFT_ID" "$IMG/logo.png")
eq "attach without --inline is an ordinary attachment" "false|none" \
   "$(posted_attachments | jq -r '"\(has("isInline"))|\(.contentId // "none")"')"

# Over 3 MB the upload goes through a session, which has to say inline too.
head -c 3200000 /dev/zero > "$IMG/big.bin"
out=$(mail "" attach "$DRAFT_ID" "$IMG/big.bin" --inline big)
eq "a large inline file opens an inline upload session" "true|big|big.bin" \
   "$(grep -A1 '/attachments/createUploadSession$' "$FAKE_CURL_LOG" | sed -n 's/^BODY //p' | jq -r '.AttachmentItem | "\(.isInline)|\(.contentId)|\(.name)"')"

out=$(mail "" attach "$DRAFT_ID" "$IMG/logo.png" --inline "a b"); rc=$?
eq "a content ID with a space is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "a refused content ID uploads nothing" "0" "$(posted_attachments | grep -c . || true)"

# A reply draft: new message, the chain marker, then the quoted chain, which
# carries a signature block of its own from an earlier message.
CHAIN='<span data-mdreply-chain-start="1"></span>'
S_START='<span data-outlook-signature-start="1"></span>'
S_END='<span data-outlook-signature-end="1"></span>'
reply_body="<div>Thanks, see you then.</div><br/>${CHAIN}<p>quoted</p>${S_START}<p>OLD QUOTED SIG</p>${S_END}"
jq -n --arg id "$DRAFT_ID" --arg b "$reply_body" '{id: $id, body: {contentType: "html", content: $b}}' > "$FAKE_DRAFT_READ"
cat > "$IMG/sig.html" <<'HTML'
<table><tr><td><img src="logo.png" alt="DBHQ"></td>
<td>Dan<br><img src='icons/phone.png'> 01234<br><img src="https://example.com/remote.png"><img src="cid:kept"></td></tr></table>
HTML

out=$(mail "" signature "$DRAFT_ID" "$IMG/sig.html"); rc=$?
eq "signature exits 0" "0" "$rc"
eq "signature uploads each local image once, inline" "2" "$(posted_attachments | jq -s '[.[] | select(.isInline == true)] | length')"
eq "signature uploads the right bytes under each content ID" \
   "sig-1-phone.png=$(b64 "$IMG/icons/phone.png") sig-2-logo.png=$(b64 "$IMG/logo.png")" \
   "$(posted_attachments | jq -r '"\(.contentId)=\(.contentBytes)"' | sort | tr '\n' ' ' | sed 's/ $//')"
body=$(patched_body)
has "signature points the logo at its cid" 'src="cid:sig-2-logo.png"' "$body"
has "signature points the icon at its cid" "src='cid:sig-1-phone.png'" "$body"
has "signature leaves a remote image alone" 'src="https://example.com/remote.png"' "$body"
has "signature leaves an existing cid alone" 'src="cid:kept"' "$body"
eq "the block sits between the message and the chain" "1" \
   "$(printf '%s' "$body" | awk -v s="$S_START" -v c="$CHAIN" '{a=index($0,"see you then"); b=index($0,s); d=index($0,c)} END{print (a && b && d && a < b && b < d) ? 1 : 0}')"
has "the quoted chain is untouched" "${CHAIN}<p>quoted</p>${S_START}<p>OLD QUOTED SIG</p>${S_END}" "$body"

# Run again on the draft as it now is: the block is replaced, not doubled, and
# images already inline are not uploaded again.
jq -n --arg id "$DRAFT_ID" --arg b "$body" '{id: $id, body: {contentType: "html", content: $b}}' > "$FAKE_DRAFT_READ"
printf '%s' '{"value":[{"contentId":"sig-1-phone.png","isInline":true},{"contentId":"sig-2-logo.png","isInline":true}]}' > "$FAKE_ATTACHMENTS"
out=$(mail "" signature "$DRAFT_ID" "$IMG/sig.html")
eq "a second run uploads nothing" "0" "$(posted_attachments | grep -c . || true)"
eq "a second run leaves one block before the chain" "1" \
   "$(patched_body | awk -v c="$CHAIN" '{print substr($0, 1, index($0, c))}' | grep -o 'data-outlook-signature-start' | wc -l | tr -d ' ')"
printf '%s' '{"value":[]}' > "$FAKE_ATTACHMENTS"

# update mdbody keeps the block and the chain. It used to keep only the chain,
# so the signature was silently dropped on the first edit.
sig_draft="<div>Thanks, see you then.</div><br/>${S_START}<p>Dan</p><img src=\"cid:sig-2-logo.png\">${S_END}<br/>${CHAIN}<p>quoted</p>${S_START}<p>OLD QUOTED SIG</p>${S_END}"
jq -n --arg id "$DRAFT_ID" --arg b "$sig_draft" '{id: $id, body: {contentType: "html", content: $b}}' > "$FAKE_DRAFT_READ"
out=$(mail "" update "$DRAFT_ID" mdbody "Rewritten **reply**")
new=$(patched_body)
has "update mdbody writes the new text" "<strong>reply</strong>" "$new"
eq "update mdbody replaces the old text" "gone" "$([ "${new#*see you then}" = "$new" ] && echo gone || echo kept)"
has "update mdbody keeps the signature block" "${S_START}<p>Dan</p><img src=\"cid:sig-2-logo.png\">${S_END}" "$new"
has "update mdbody keeps the chain" "${CHAIN}<p>quoted</p>${S_START}<p>OLD QUOTED SIG</p>${S_END}" "$new"
eq "update mdbody puts the block before the chain" "1" \
   "$(printf '%s' "$new" | awk -v s="$S_START" -v c="$CHAIN" '{a=index($0,"<strong>reply"); b=index($0,s); d=index($0,c)} END{print (a && b && d && a < b && b < d) ? 1 : 0}')"

# A plain-text draft, a missing image: nothing is uploaded or changed.
jq -n --arg id "$DRAFT_ID" '{id: $id, body: {contentType: "text", content: "plain"}}' > "$FAKE_DRAFT_READ"
out=$(mail "" signature "$DRAFT_ID" "$IMG/sig.html"); rc=$?
eq "signature on a plain-text draft exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "signature on a plain-text draft changes nothing" "0" "$(grep -c '^\(POST\|PATCH\) ' "$FAKE_CURL_LOG" || true)"

jq -n --arg id "$DRAFT_ID" --arg b "$reply_body" '{id: $id, body: {contentType: "html", content: $b}}' > "$FAKE_DRAFT_READ"
printf '<img src="logo.png"><img src="missing.png">' > "$IMG/broken.html"
out=$(mail "" signature "$DRAFT_ID" "$IMG/broken.html"); rc=$?
eq "a missing image makes signature exit non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "a missing image is named" "missing.png" "$out"
eq "a missing image uploads and changes nothing" "0" "$(grep -c '^\(POST\|PATCH\) ' "$FAKE_CURL_LOG" || true)"

########################################
# Reply-all: every reply verb keeps all the recipients and prints them all
########################################
# The printed To and Cc lines are how the user checks who a reply goes to
# before it is sent. A sender-only reply would drop the Cc line silently.
printf '{"id":"%s","subject":"RE: Plans","isDraft":true,"body":{"contentType":"HTML","content":"<p>QUOTED HISTORY</p>"},"toRecipients":[{"emailAddress":{"address":"alice@example.com"}},{"emailAddress":{"address":"bob@example.com"}}],"ccRecipients":[{"emailAddress":{"address":"carol@example.com"}},{"emailAddress":{"address":"dave@example.com"}}],"bccRecipients":[{"emailAddress":{"address":"audit@example.com"}}],"from":null}' \
    "$DRAFT_ID" > "$FAKE_DRAFT"
reply_all_posts() { grep -c "^POST https://graph.microsoft.com/v1.0/me/messages/$MSG_ID/createReplyAll\$" "$FAKE_CURL_LOG" || true; }
sender_only_posts() { grep -c '/createReply$' "$FAKE_CURL_LOG" || true; }

reply_verbs="reply"
command -v pandoc >/dev/null 2>&1 && reply_verbs="reply mdreply followup"
for verb in $reply_verbs; do
    out=$(mail "" "$verb" "$MSG_ID" "Thanks")
    eq "$verb creates a reply-all draft" "1" "$(reply_all_posts)"
    eq "$verb never creates a sender-only reply" "0" "$(sender_only_posts)"
    has "$verb prints every To recipient" "To:      alice@example.com, bob@example.com" "$out"
    has "$verb prints every Cc recipient" "Cc:      carol@example.com, dave@example.com" "$out"
    has "$verb prints the Bcc" "Bcc:     audit@example.com" "$out"
done
command -v pandoc >/dev/null 2>&1 || echo "skip - mdreply/followup reply-all tests (pandoc not installed)"

# No Cc on the original: no Cc line, rather than an empty one.
jq '.ccRecipients = [] | .bccRecipients = []' "$FAKE_DRAFT" > "$TMP/no_cc.json" && mv "$TMP/no_cc.json" "$FAKE_DRAFT"
out=$(mail "" reply "$MSG_ID" "Thanks")
eq "reply prints no Cc line when there is no Cc" "0" "$(printf '%s\n' "$out" | grep -c '^Cc:' || true)"
has "reply still prints both To recipients" "To:      alice@example.com, bob@example.com" "$out"

########################################
# The chain marker that update mdbody relies on
########################################
# mdreply, followup and forward put an empty span between the new message and
# the quoted history. update mdbody keeps everything from that span onwards.
# Whether Exchange keeps the span in a saved draft is not something a fixture
# can answer: tests/chain_marker_live.sh asks the real server.
if command -v pandoc >/dev/null 2>&1; then
    # The body of the last PATCH to the draft, on one line.
    last_patch_body() {
        grep -A1 "^PATCH https://graph.microsoft.com/v1.0/me/messages/$DRAFT_ID\$" "$FAKE_CURL_LOG" \
            | sed -n 's/^BODY //p' | tail -1 | jq -j '.body.content' | tr '\n' ' '
    }
    # 1 when the three strings appear in this order in $1.
    in_order() {
        printf '%s' "$1" | awk -v a="$2" -v b="$3" -v c="$4" \
            '{x=index($0,a); y=index($0,b); z=index($0,c)} END{print (x && y && z && x < y && y < z) ? 1 : 0}'
    }

    for verb in mdreply followup; do
        out=$(mail "" "$verb" "$MSG_ID" "New **text**")
        eq "$verb puts the marker between the new text and the quoted history" "1" \
           "$(in_order "$(last_patch_body)" "<strong>text</strong>" "$CHAIN" "QUOTED HISTORY")"
    done
    out=$(mail "" forward "$MSG_ID" "bob@example.com" "New **text**")
    eq "forward with a comment puts the marker between the comment and the message" "1" \
       "$(in_order "$(last_patch_body)" "<strong>text</strong>" "$CHAIN" "QUOTED HISTORY")"

    # update mdbody on a reply draft: the new text replaces the old, and the
    # chain is kept byte for byte from the marker onwards.
    chain_draft="<p>OLD TEXT</p><br/>${CHAIN}<div><p>QUOTED HISTORY</p></div>"
    jq -n --arg id "$DRAFT_ID" --arg b "$chain_draft" '{id: $id, body: {contentType: "html", content: $b}}' > "$FAKE_DRAFT_READ"
    out=$(mail "" update "$DRAFT_ID" mdbody "Rewritten **reply**"); rc=$?
    new=$(last_patch_body)
    eq "update mdbody on a reply exits 0" "0" "$rc"
    eq "update mdbody reads the draft body before it writes" "1" \
       "$(awk -v d="$DRAFT_ID" '$0 ~ "^GET .*/me/messages/"d"\\?\\$select=body$" {g=NR} $0 ~ "^PATCH .*/me/messages/"d"$" {p=NR} END{print (g && p && g < p) ? 1 : 0}' "$FAKE_CURL_LOG")"
    eq "update mdbody drops the old text" "gone" "$([ "${new#*OLD TEXT}" = "$new" ] && echo gone || echo kept)"
    eq "update mdbody keeps the chain from the marker onwards" "${CHAIN}<div><p>QUOTED HISTORY</p></div>" "${CHAIN}${new#*"$CHAIN"}"
    eq "update mdbody puts the new text before the chain" "1" "$(in_order "$new" "<strong>reply</strong>" "$CHAIN" "QUOTED HISTORY")"

    # A reply to a reply: the quoted history holds an earlier marker of its
    # own. The split is on the FIRST marker, so none of the history is lost.
    nested="<p>OLD TEXT</p><br/>${CHAIN}<p>THEIR REPLY</p><br/>${CHAIN}<p>MY FIRST MESSAGE</p>"
    jq -n --arg id "$DRAFT_ID" --arg b "$nested" '{id: $id, body: {contentType: "html", content: $b}}' > "$FAKE_DRAFT_READ"
    out=$(mail "" update "$DRAFT_ID" mdbody "Rewritten")
    new=$(last_patch_body)
    has "update mdbody splits on the first marker and keeps the whole history" \
        "${CHAIN}<p>THEIR REPLY</p><br/>${CHAIN}<p>MY FIRST MESSAGE</p>" "$new"
    eq "update mdbody on a reply to a reply drops only the new text" "gone" \
       "$([ "${new#*OLD TEXT}" = "$new" ] && echo gone || echo kept)"

    # No marker: the whole body is replaced. This is what would happen to
    # every reply if Exchange stripped the marker on save, and why the live
    # check exists.
    jq -n --arg id "$DRAFT_ID" '{id: $id, body: {contentType: "html", content: "<p>OLD TEXT</p><p>UNMARKED HISTORY</p>"}}' > "$FAKE_DRAFT_READ"
    out=$(mail "" update "$DRAFT_ID" mdbody "Rewritten")
    new=$(last_patch_body)
    eq "update mdbody with no marker replaces the whole body" "gone" \
       "$([ "${new#*UNMARKED HISTORY}" = "$new" ] && echo gone || echo kept)"
    eq "update mdbody with no marker adds no marker" "0" "$(printf '%s' "$new" | grep -c 'data-mdreply-chain-start' || true)"
else
    echo "skip - chain marker tests (pandoc not installed)"
fi

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
