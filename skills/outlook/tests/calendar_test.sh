#!/bin/bash
# Offline tests for outlook-calendar.sh, run as a user would run it.
#
# A fake `curl` first on PATH answers Graph from fixture files and logs every
# request, and a throwaway HOME holds a fresh token, so no account or network
# is needed and nothing is sent to anyone. Nothing here touches ~/.dbhq.
#
#   bash skills/outlook/tests/calendar_test.sh
#
# What this pins: listings page through @odata.nextLink instead of stopping at
# Graph's default 10, every listed event carries an ID that read, respond and
# cancel accept, `events` shows each occurrence of a recurring meeting, and
# `free` ignores events that do not block time.
#
# The literal '$skip' and '$top' strings are Graph query parameters (SC2016).
# shellcheck disable=SC2016
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CAL="$(cd "$TESTS_DIR/../scripts" && pwd)/outlook-calendar.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }
lacks() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" "does not contain: $2" "$3"; else eq "$1" ok ok; fi; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# --- Fake curl -----------------------------------------------------------------
# calendarView: page 1 from $FAKE_PAGE1, and a URL carrying $skip gets
# $FAKE_PAGE2. Anything posted is logged and answered with an empty 202.
mkdir -p "$TMP/bin"
cat > "$TMP/bin/curl" <<'FAKE'
#!/bin/bash
url="" method=GET prev=""
for a in "$@"; do
  case "$prev" in -X) method="$a" ;; esac
  case "$a" in https://*) url="$a" ;; esac
  prev="$a"
done
printf '%s %s\n' "$method" "$url" >> "$FAKE_CURL_LOG"
case "$url" in
  */me/calendar/calendarView*'$skip'*) cat "$FAKE_PAGE2" ;;
  */me/calendar/calendarView*)         cat "$FAKE_PAGE1" ;;
  */me/events\?*)                      printf '{"value":[]}' ;;
  */me/calendar/events/*)              printf '{"id":"x","subject":"Read me","start":{"dateTime":"2026-09-01T09:00:00","timeZone":"Europe/London"},"end":{"dateTime":"2026-09-01T10:00:00","timeZone":"Europe/London"}}' ;;
  *) [ "$method" = GET ] && printf '{"value":[]}' ;;
esac
exit 0
FAKE
chmod +x "$TMP/bin/curl"
export PATH="$TMP/bin:$PATH"
export FAKE_CURL_LOG="$TMP/curl.log" FAKE_PAGE1="$TMP/page1.json" FAKE_PAGE2="$TMP/page2.json"

# --- Fixture HOME with a token that does not need refreshing --------------------
FIX_HOME="$TMP/home"
ACC="$FIX_HOME/.dbhq/outlook/default"
mkdir -p "$ACC"
chmod 700 "$FIX_HOME/.dbhq" "$FIX_HOME/.dbhq/outlook" "$ACC"
printf '%s' '{"client_id":"c","client_secret":"s"}' > "$ACC/config.json"
printf '{"access_token":"tok","refresh_token":"r","expires_at":%s}' "$(( $(date +%s) + 3600 ))" > "$ACC/credentials.json"
chmod 600 "$ACC/credentials.json"

cal() {
    : > "$FAKE_CURL_LOG"
    HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default OUTLOOK_TZ=Europe/London bash "$CAL" "$@" 2>&1
}

# A realistic-length event ID with a distinct tail.
eid() { printf 'AAMkADkzNjdmZjRlLTc4ZWYtNDdiNC1hZDdiLWE3MjZmNTg0YjAwNgBGAAAAAAD%s' "$1"; }
event() {  # id day subject [showAs] [isCancelled] [type]
    printf '{"id":"%s","subject":"%s","start":{"dateTime":"%sT09:00:00.0000000","timeZone":"Europe/London"},"end":{"dateTime":"%sT10:00:00.0000000","timeZone":"Europe/London"},"location":{"displayName":"Room"},"showAs":"%s","isCancelled":%s,"type":"%s"}' \
        "$1" "$3" "$2" "$2" "${4:-busy}" "${5:-false}" "${6:-singleInstance}"
}

########################################
# A busy week: 12 events over two pages. All 12 must show, each with an ID.
########################################
{
    printf '{"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/calendar/calendarView?startDateTime=a&endDateTime=b&$skip=10","value":['
    for i in 1 2 3 4 5 6 7 8 9 10; do
        [ "$i" -gt 1 ] && printf ','
        event "$(eid "WEEKEVENT$(printf '%03d' "$i")AAAA=")" "2026-09-2$((i % 7))" "Meeting $i"
    done
    printf ']}'
} > "$FAKE_PAGE1"
{
    printf '{"value":['
    event "$(eid "WEEKEVENT011AAAA=")" "2026-09-28" "Meeting 11"; printf ','
    event "$(eid "WEEKEVENT012AAAA=")" "2026-09-29" "Meeting 12"
    printf ']}'
} > "$FAKE_PAGE2"

out=$(cal week)
eq "week shows all 12 events across two pages" "12" "$(printf '%s\n' "$out" | grep -c '^\[[0-9]*\] ')"
has "week shows the 12th event" "Meeting 12" "$out"
eq "week follows the next page" "2" "$(grep -c 'calendarView' "$FAKE_CURL_LOG")"
eq "week asks for 100 per page" "1" "$(head -1 "$FAKE_CURL_LOG" | grep -c '\$top=100')"
has "week prints each event's short ID" "[12] $(eid "WEEKEVENT012AAAA=" | tail -c 20) | 2026-09-29 09:00-10:00 | Meeting 12 | Room" "$out"

# The ID on screen is one the other commands take.
short=$(eid "WEEKEVENT012AAAA=" | tail -c 20)
out=$(cal read "$short")
eq "read accepts a short ID from week" "1" "$(grep -c "GET https://graph.microsoft.com/v1.0/me/calendar/events/$(eid "WEEKEVENT012AAAA=")$" "$FAKE_CURL_LOG")"
has "read shows the event" "Subject: Read me" "$out"

for verb in today "day 2026-09-21" "search Meeting"; do
    # shellcheck disable=SC2086 # the verb and its argument split on purpose
    out=$(cal $verb)
    has "$verb prints short IDs" "$(eid "WEEKEVENT001AAAA=" | tail -c 20)" "$out"
done

########################################
# A recurring series: `events` lists each occurrence on its own row, and
# respond takes an occurrence's short ID.
########################################
OCC1=$(eid "SERIESxOCC20260901AA=")
OCC2=$(eid "SERIESxOCC20260908AA=")
OCC3=$(eid "SERIESxOCC20260915AA=")
{
    printf '{"value":['
    event "$OCC1" "2026-09-01" "Weekly sync" busy false occurrence; printf ','
    event "$OCC2" "2026-09-08" "Weekly sync" busy false occurrence; printf ','
    event "$OCC3" "2026-09-15" "Weekly sync" busy false occurrence
    printf ']}'
} > "$FAKE_PAGE1"

out=$(cal events)
eq "events lists each occurrence" "3" "$(printf '%s\n' "$out" | grep -c 'Weekly sync')"
eq "events reads calendarView, not the series list" "1" "$(grep -c '/me/calendar/calendarView' "$FAKE_CURL_LOG")"
eq "events does not read /me/calendar/events" "0" "$(grep -c '/me/calendar/events?' "$FAKE_CURL_LOG")"
has "events prints the second occurrence's short ID" "$(printf '%s' "$OCC2" | tail -c 20)" "$out"

cal respond "$(printf '%s' "$OCC2" | tail -c 20)" accept >/dev/null
eq "respond resolves an occurrence's short ID (from the listing)" "1" \
   "$(grep -c "POST https://graph.microsoft.com/v1.0/me/events/$OCC2/accept$" "$FAKE_CURL_LOG")"

rm -f "$ACC/event_id_cache.json"
cal respond "$(printf '%s' "$OCC3" | tail -c 20)" decline >/dev/null
eq "respond resolves an occurrence's short ID with no listing cached" "1" \
   "$(grep -c "POST https://graph.microsoft.com/v1.0/me/events/$OCC3/decline$" "$FAKE_CURL_LOG")"

cal cancel "$(printf '%s' "$OCC1" | tail -c 20)" >/dev/null
eq "cancel accepts a short ID from events" "1" \
   "$(grep -c "POST https://graph.microsoft.com/v1.0/me/events/$OCC1/cancel$" "$FAKE_CURL_LOG")"

########################################
# Two IDs with the same last 20 characters are printed in full, and the short
# form is refused rather than guessed.
########################################
{
    printf '{"value":['
    event "AAAAoneSERIES-SAME-TAIL-FOR-BOTH=" "2026-09-02" "Clash A"; printf ','
    event "BBBBtwoSERIES-SAME-TAIL-FOR-BOTH=" "2026-09-03" "Clash B"
    printf ']}'
} > "$FAKE_PAGE1"
out=$(cal today)
has "clashing IDs are printed in full (A)" "AAAAoneSERIES-SAME-TAIL-FOR-BOTH=" "$out"
has "clashing IDs are printed in full (B)" "BBBBtwoSERIES-SAME-TAIL-FOR-BOTH=" "$out"
out=$(cal respond "$(printf '%s' "AAAAoneSERIES-SAME-TAIL-FOR-BOTH=" | tail -c 20)" accept); rc=$?
eq "a clashing short ID is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "a clashing short ID sends nothing" "0" "$(grep -c '^POST' "$FAKE_CURL_LOG")"

########################################
# free: an event shown as free, or cancelled, does not block time.
########################################
printf '{"value":[%s]}' "$(event "$(eid FREEMARKER0001AAA=)" 2026-09-01 "Out of office note" free)" > "$FAKE_PAGE1"
has "free: an event marked free leaves you free" "You are FREE" "$(cal free 2026-09-01T08:00 2026-09-01T18:00)"
printf '{"value":[%s]}' "$(event "$(eid CANCELLED0001AAA=)" 2026-09-01 "Cancelled call" busy true)" > "$FAKE_PAGE1"
has "free: a cancelled event leaves you free" "You are FREE" "$(cal free 2026-09-01T08:00 2026-09-01T18:00)"
printf '{"value":[%s,%s]}' "$(event "$(eid BUSY0001AAAAAAA=)" 2026-09-01 "Real meeting")" \
    "$(event "$(eid FREEMARKER0002AAA=)" 2026-09-01 "Reminder" free)" > "$FAKE_PAGE1"
out=$(cal free 2026-09-01T08:00 2026-09-01T18:00)
has "free: a busy event still counts" "You have 1 event(s)" "$out"
has "free: the busy event is listed with its ID" "$(eid BUSY0001AAAAAAA= | tail -c 20)" "$out"
lacks "free: the free event is not listed" "Reminder" "$out"

########################################
# A Graph error is reported and the command exits non-zero.
########################################
printf '%s' '{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}' > "$FAKE_PAGE1"
out=$(cal week); rc=$?
eq "week on a Graph error exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "week on a Graph error says why" "Access is denied." "$out"

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
