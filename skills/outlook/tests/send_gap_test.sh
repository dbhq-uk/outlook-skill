#!/bin/bash
# Offline tests for the send gap as the scripts enforce it, run as a user would
# run them.
#
# A fake `curl` first on PATH logs every request and answers Graph from
# fixtures, and a throwaway HOME holds the token. Nothing is sent to anyone and
# nothing here touches ~/.dbhq.
#
#   bash skills/outlook/tests/send_gap_test.sh
#
# What this pins:
#   - OUTLOOK_READ_ONLY=1 refuses every write verb in both scripts before any
#     request is made, token refresh included, and lets every read verb run.
#   - Every verb in each script's dispatch is classed as read or write here, so
#     a new verb cannot slip past read-only mode unclassified.
#   - calendar `create` with attendees needs --send-invites.
#   - calendar `update` of a meeting you organise needs --notify-attendees, and
#     `delete` of one is refused in favour of `cancel`, because Graph notifies
#     the attendees in both cases.
#
# The literal '$select' strings are Graph query parameters (SC2016).
# shellcheck disable=SC2016
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPTS="$(cd "$TESTS_DIR/../scripts" && pwd)"
MAIL="$SCRIPTS/outlook-mail.sh"
CAL="$SCRIPTS/outlook-calendar.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }
lacks() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" "does not contain: $2" "$3"; else eq "$1" ok ok; fi; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# --- Fake curl -----------------------------------------------------------------
# Logs "METHOD URL" and, on the next line, "BODY <json>" for a request with a
# body. The event read that update and delete make is answered from
# $FAKE_EVENT; a created or patched event is echoed back with an id.
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
[ -n "$data" ] && printf 'BODY %s\n' "$(printf '%s' "$data" | jq -c . 2>/dev/null || printf '%s' "$data")" >> "$FAKE_CURL_LOG"
path="${url#https://graph.microsoft.com/v1.0}"
case "$method $path" in
  "POST https://login.microsoftonline.com"*|"POST "*oauth2*)
      printf '{"access_token":"new","refresh_token":"r2","expires_in":3600}' ;;
  "GET /me/events/"*'$select=isOrganizer'*) cat "$FAKE_EVENT" ;;
  "POST /me/calendar/events")
      printf '%s' "$data" | jq -c '. + {id: "NEWEVENTIDxxxxxxxxxxxxxxxxxxxxxxxxxx"}' ;;
  "PATCH "*) printf '%s' "$data" | jq -c '. + {id: "x", subject: "s", start: {dateTime: "a"}, end: {dateTime: "b"}}' ;;
  "DELETE "*|"POST "*) : ;;
  *) printf '{"value":[]}' ;;
esac
exit 0
FAKE
chmod +x "$TMP/bin/curl"
export PATH="$TMP/bin:$PATH"
export FAKE_CURL_LOG="$TMP/curl.log" FAKE_EVENT="$TMP/event.json"

# --- Fixture HOME -----------------------------------------------------------------
FIX_HOME="$TMP/home"
ACC="$FIX_HOME/.dbhq/outlook/default"
mkdir -p "$ACC"
chmod 700 "$FIX_HOME/.dbhq" "$FIX_HOME/.dbhq/outlook" "$ACC"
printf '%s' '{"client_id":"c","client_secret":"s"}' > "$ACC/config.json"
fresh_token()   { printf '{"access_token":"tok","refresh_token":"r","expires_at":%s}' "$(( $(date +%s) + 3600 ))" > "$ACC/credentials.json"; }
expired_token() { printf '{"access_token":"tok","refresh_token":"r","expires_at":%s}' "$(( $(date +%s) - 3600 ))" > "$ACC/credentials.json"; }
fresh_token

run() {  # run <script> [args...] - OUTLOOK_READ_ONLY comes from the caller's environment
    local script="$1"; shift
    : > "$FAKE_CURL_LOG"
    HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default OUTLOOK_TZ=Europe/London OUTLOOK_FROM_ADDRESS='' \
        bash "$script" "$@" 2>&1
}
requests() { grep -vc '^BODY ' "$FAKE_CURL_LOG" || true; }

# Every verb the script's top-level dispatch accepts, one per line.
dispatch_verbs() {
    awk '/^case "\$1" in/{c=1; next} c && /^esac/{exit} c && /^    [a-z][a-z|-]*\)$/{print}' "$1" \
        | sed 's/^ *//; s/)$//' | tr '|' '\n' | sort
}

# The classification this test holds. A verb that is in the script and in
# neither list fails the first check below, so it has to be decided on.
MAIL_READ="inbox unread focused sent from search read preview aliases drafts flagged thread category categories folders subfolders folder stats attachments download export"
MAIL_WRITE="draft mddraft reply update mdreply followup forward send markread markunread flag unflag categorize mkcategory rccategory rmcategory junk notjunk delete archive move batch-move bulk-move mkdir rename rmdir attach"
CAL_READ="events today week read calendars day search free"
CAL_WRITE="create invite quick update delete cancel respond"

# Message IDs over 100 characters pass through resolve_message_id untouched.
LONG_ID="MSG$(printf 'm%.0s' $(seq 1 110))"

words_sorted() { printf '%s\n' "$@" | tr ' ' '\n' | grep . | sort; }

eq "every mail verb is classed as read or write" \
   "$(words_sorted "$MAIL_READ $MAIL_WRITE" | tr '\n' ' ')" "$(dispatch_verbs "$MAIL" | tr '\n' ' ')"
eq "every calendar verb is classed as read or write" \
   "$(words_sorted "$CAL_READ $CAL_WRITE" | tr '\n' ' ')" "$(dispatch_verbs "$CAL" | tr '\n' ' ')"
eq "no mail verb is classed as both" "" "$(words_sorted "$MAIL_READ" "$MAIL_WRITE" | uniq -d | tr '\n' ' ')"
eq "no calendar verb is classed as both" "" "$(words_sorted "$CAL_READ" "$CAL_WRITE" | uniq -d | tr '\n' ' ')"

# The scripts' own allow-lists match the read verbs here.
script_read_verbs() { sed -n 's/^READ_ONLY_VERBS=(\(.*\))$/\1/p' "$1" | tr ' ' '\n' | grep . | sort | tr '\n' ' '; }
eq "outlook-mail.sh allows exactly the mail read verbs" "$(words_sorted "$MAIL_READ" | tr '\n' ' ')" "$(script_read_verbs "$MAIL")"
eq "outlook-calendar.sh allows exactly the calendar read verbs" "$(words_sorted "$CAL_READ" | tr '\n' ' ')" "$(script_read_verbs "$CAL")"

# A verb classed as read makes no write call in its own branch of the dispatch.
writes_in_branch() {  # writes_in_branch <script> <verb>
    awk -v v="$2" '/^case "\$1" in/{c=1; next} c && /^esac/{exit}
        c && /^    [a-z][a-z|-]*\)$/{ n=$1; sub(/\)$/, "", n); k=split(n, a, "|"); in_v=0; for (i=1;i<=k;i++) if (a[i]==v) in_v=1; next }
        c && in_v && /api_call (PATCH|POST|DELETE|PUT)/{print NR}' "$1"
}
for verb in $MAIL_READ; do eq "mail read verb $verb makes no write call" "" "$(writes_in_branch "$MAIL" "$verb")"; done
for verb in $CAL_READ; do eq "calendar read verb $verb makes no write call" "" "$(writes_in_branch "$CAL" "$verb")"; done
eq "the write-call scan finds mail send's POST" "1" "$(writes_in_branch "$MAIL" send | grep -c .)"

########################################
# OUTLOOK_READ_ONLY=1: every write verb refuses and makes no request at all.
# The token is expired, so even a token refresh would show up in the log.
########################################
expired_token
export OUTLOOK_READ_ONLY=1
for verb in $MAIL_WRITE; do
    out=$(run "$MAIL" "$verb" AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA "b@example.com" "c"); rc=$?
    eq "read-only: mail $verb exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
    has "read-only: mail $verb says why" "Refused: OUTLOOK_READ_ONLY is set" "$out"
    eq "read-only: mail $verb makes no request" "0" "$(requests)"
done
for verb in $CAL_WRITE; do
    out=$(run "$CAL" "$verb" AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA "2026-10-01T10:00" "2026-10-01T11:00" "" "b@example.com" --send-invites); rc=$?
    eq "read-only: calendar $verb exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
    has "read-only: calendar $verb says why" "Refused: OUTLOOK_READ_ONLY is set" "$out"
    eq "read-only: calendar $verb makes no request" "0" "$(requests)"
done

# --account before the verb is still gated on the verb.
out=$(run "$MAIL" --account default send AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA); rc=$?
eq "read-only: --account default send is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "read-only: --account default send makes no request" "0" "$(requests)"

# An unknown verb is refused too: the gate is an allow-list.
out=$(run "$MAIL" frobnicate); rc=$?
has "read-only: an unknown verb is refused" "Refused: OUTLOOK_READ_ONLY is set" "$out"

# Read verbs still run. With no arguments most print their usage; none refuses.
fresh_token
for verb in $MAIL_READ; do
    out=$(run "$MAIL" "$verb")
    lacks "read-only: mail $verb still runs" "Refused: OUTLOOK_READ_ONLY" "$out"
done
for verb in $CAL_READ; do
    out=$(run "$CAL" "$verb")
    lacks "read-only: calendar $verb still runs" "Refused: OUTLOOK_READ_ONLY" "$out"
done
out=$(run "$MAIL")
lacks "read-only: the usage text still prints" "Refused" "$out"
has "the usage text names read-only mode" "OUTLOOK_READ_ONLY=1" "$out"

# Any value but empty, 0, false, no or off turns it on.
expired_token
for v in true yes TRUE on; do
    out=$(OUTLOOK_READ_ONLY=$v run "$MAIL" markread "$LONG_ID")
    eq "OUTLOOK_READ_ONLY=$v refuses" "0" "$(requests)"
done
fresh_token
for v in 0 false no off ""; do
    out=$(OUTLOOK_READ_ONLY=$v run "$MAIL" markread "$LONG_ID")
    eq "OUTLOOK_READ_ONLY='$v' does not refuse" "1" "$(grep -c '^PATCH ' "$FAKE_CURL_LOG")"
done
unset OUTLOOK_READ_ONLY

########################################
# create: attendees need --send-invites.
########################################
fresh_token
out=$(run "$CAL" create "Kickoff" 2026-10-01T10:00 2026-10-01T11:00 "" "a@example.com, b@example.com"); rc=$?
eq "create with attendees and no --send-invites exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "create with attendees and no --send-invites says how" "--send-invites" "$out"
eq "create with attendees and no --send-invites makes no request" "0" "$(requests)"

out=$(run "$CAL" create "Kickoff" 2026-10-01T10:00 2026-10-01T11:00 "" "a@example.com, b@example.com" --send-invites); rc=$?
eq "create with --send-invites succeeds" "0" "$rc"
eq "create with --send-invites posts both attendees" "a@example.com,b@example.com" \
   "$(sed -n 's/^BODY //p' "$FAKE_CURL_LOG" | jq -r '[.attendees[].emailAddress.address] | join(",")')"
eq "the flag never reaches Graph as an argument" "0" "$(grep -c -- '--send-invites' "$FAKE_CURL_LOG")"

out=$(run "$CAL" create --send-invites "Kickoff" 2026-10-01T10:00 2026-10-01T11:00 "Room 1" "a@example.com")
eq "--send-invites before the arguments works too" "Kickoff|Room 1|a@example.com" \
   "$(sed -n 's/^BODY //p' "$FAKE_CURL_LOG" | jq -r '[.subject, .location.displayName, .attendees[0].emailAddress.address] | join("|")')"

out=$(run "$CAL" create "Focus time" 2026-10-01T10:00 2026-10-01T11:00 "Desk"); rc=$?
eq "create without attendees needs no flag" "0" "$rc"
eq "create without attendees posts no attendees" "false" \
   "$(sed -n 's/^BODY //p' "$FAKE_CURL_LOG" | jq 'has("attendees")')"

out=$(run "$CAL" create "Focus time" 2026-10-01T10:00 2026-10-01T11:00 "Desk" "  ; ,"); rc=$?
eq "an attendee list with no address is not an invite" "0" "$rc"

########################################
# update and delete on a meeting you organise.
########################################
EVT="EVENTIDxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
organised_meeting() { printf '{"isOrganizer":true,"attendees":[{"emailAddress":{"address":"a@example.com"}},{"emailAddress":{"address":"b@example.com"}}]}' > "$FAKE_EVENT"; }
own_event()         { printf '{"isOrganizer":true,"attendees":[]}' > "$FAKE_EVENT"; }
their_meeting()     { printf '{"isOrganizer":false,"attendees":[{"emailAddress":{"address":"boss@example.com"}}]}' > "$FAKE_EVENT"; }

organised_meeting
out=$(run "$CAL" update "$EVT" start 2026-10-01T09:00); rc=$?
eq "update of an organised meeting without the flag exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "update of an organised meeting names who would hear" "a@example.com, b@example.com" "$out"
eq "update of an organised meeting without the flag patches nothing" "0" "$(grep -c '^PATCH ' "$FAKE_CURL_LOG")"

out=$(run "$CAL" update "$EVT" start 2026-10-01T09:00 --notify-attendees); rc=$?
eq "update with --notify-attendees succeeds" "0" "$rc"
eq "update with --notify-attendees patches once" "1" "$(grep -c "^PATCH https://graph.microsoft.com/v1.0/me/calendar/events/$EVT\$" "$FAKE_CURL_LOG")"
eq "update with --notify-attendees patches the value, not the flag" "2026-10-01T09:00" \
   "$(sed -n 's/^BODY //p' "$FAKE_CURL_LOG" | jq -r '.start.dateTime')"

own_event
out=$(run "$CAL" update "$EVT" subject "New title"); rc=$?
eq "update of your own event needs no flag" "1" "$(grep -c '^PATCH ' "$FAKE_CURL_LOG")"
their_meeting
out=$(run "$CAL" update "$EVT" subject "New title"); rc=$?
eq "update of someone else's meeting needs no flag" "1" "$(grep -c '^PATCH ' "$FAKE_CURL_LOG")"

organised_meeting
out=$(run "$CAL" delete "$EVT"); rc=$?
eq "delete of an organised meeting exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "delete of an organised meeting points to cancel" "cancel <event-id>" "$out"
eq "delete of an organised meeting deletes nothing" "0" "$(grep -c '^DELETE ' "$FAKE_CURL_LOG")"

own_event
out=$(run "$CAL" delete "$EVT"); rc=$?
eq "delete of your own event deletes it" "1" "$(grep -c "^DELETE https://graph.microsoft.com/v1.0/me/calendar/events/$EVT\$" "$FAKE_CURL_LOG")"

printf '%s' '{"error":{"code":"ErrorItemNotFound","message":"The specified object was not found."}}' > "$FAKE_EVENT"
out=$(run "$CAL" delete "$EVT"); rc=$?
eq "delete exits non-zero when the event cannot be read" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "delete deletes nothing when the event cannot be read" "0" "$(grep -c '^DELETE ' "$FAKE_CURL_LOG")"
out=$(run "$CAL" update "$EVT" subject "x"); rc=$?
eq "update patches nothing when the event cannot be read" "0" "$(grep -c '^PATCH ' "$FAKE_CURL_LOG")"

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
