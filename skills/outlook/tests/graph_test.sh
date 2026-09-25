#!/bin/bash
# Offline tests for how the scripts handle Graph failures: timeouts on every
# curl call, retries for throttling (429) and for 503/504 on safe methods, and
# batch-move reporting a batch that failed as a whole.
#
# A fake `curl` first on PATH answers from numbered response files, one per
# call of each method, and writes the status and headers to the file curl's -D
# names, as real curl does. A throwaway HOME holds a fresh token. Nothing is
# sent to anyone and nothing here touches ~/.dbhq.
#
#   bash skills/outlook/tests/graph_test.sh
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPTS="$(cd "$TESTS_DIR/../scripts" && pwd)"
MAIL="$SCRIPTS/outlook-mail.sh"
CAL="$SCRIPTS/outlook-calendar.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }
hasnt() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" "no: $2" "$3"; else eq "$1" ok ok; fi; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# --- Fake curl -----------------------------------------------------------------
# The Nth call with method M answers from $FAKE_DIR/M.N, else $FAKE_DIR/M.default.
# A response file is a status line, any header lines, a blank line, then the
# body. Each call is logged as "METHOD URL max-time=<value or none>", then the
# request body on a "BODY " line if there is one. With -f and a status of 400
# or more it writes no body and exits 22, as curl does.
mkdir -p "$TMP/bin"
cat > "$TMP/bin/curl" <<'FAKE'
#!/bin/bash
url="" method=GET data="" prev="" hdr="" out="" fail=0 maxtime=none
for a in "$@"; do
  case "$prev" in
    -X) method="$a" ;;
    -d|--data-binary) data="$a" ;;
    -D) hdr="$a" ;;
    -o) out="$a" ;;
    --max-time) maxtime="$a" ;;
  esac
  case "$a" in https://*) url="$a" ;; -f|-sf) fail=1 ;; esac
  prev="$a"
done
n=$(( $(cat "$FAKE_DIR/$method.count" 2>/dev/null || echo 0) + 1 ))
echo "$n" > "$FAKE_DIR/$method.count"
# A body on stdin (an upload chunk) is binary: it is kept byte for byte in
# $FAKE_DIR/<method>.<n>.body, and its Content-Range is logged.
if [ "$data" = "@-" ]; then
  cat > "$FAKE_DIR/$method.$n.body"
  data=""
  prev=""; for a in "$@"; do [ "$prev" = "-H" ] && case "$a" in Content-Range:*) printf 'RANGE %s\n' "${a#Content-Range: }" >> "$FAKE_CURL_LOG" ;; esac; prev="$a"; done
fi
case "$data" in @*) data=$(cat "${data#@}") ;; esac
printf '%s %s max-time=%s\n' "$method" "$url" "$maxtime" >> "$FAKE_CURL_LOG"
[ -n "$data" ] && printf 'BODY %s\n' "$(printf '%s' "$data" | jq -c . 2>/dev/null || printf '%s' "$data")" >> "$FAKE_CURL_LOG"

resp="$FAKE_DIR/$method.$n"
[ -f "$resp" ] || resp="$FAKE_DIR/$method.default"
[ -f "$resp" ] || { printf '200\n\n{}' > "$TMP_EMPTY"; resp="$TMP_EMPTY"; }

status=$(head -1 "$resp")
headers=$(awk 'NR > 1 && /^$/ {exit} NR > 1 {print}' "$resp")
body=$(awk 'f {print} NR > 1 && /^$/ {f=1}' "$resp")
if [ -n "$hdr" ]; then
  { printf 'HTTP/1.1 %s X\r\n' "$status"
    [ -n "$headers" ] && printf '%s\n' "$headers" | sed 's/$/\r/'
    printf '\r\n'; } > "$hdr"
fi
if [ "$fail" = 1 ] && [ "$status" -ge 400 ]; then exit 22; fi
if [ -n "$out" ]; then printf '%s' "$body" > "$out"; else printf '%s' "$body"; fi
exit 0
FAKE
chmod +x "$TMP/bin/curl"
export PATH="$TMP/bin:$PATH"
export FAKE_CURL_LOG="$TMP/curl.log" TMP_EMPTY="$TMP/empty.resp"

# --- Fixture HOME with a token that does not need refreshing --------------------
FIX_HOME="$TMP/home"
ACC="$FIX_HOME/.dbhq/outlook/default"
mkdir -p "$ACC"
chmod 700 "$FIX_HOME/.dbhq" "$FIX_HOME/.dbhq/outlook" "$ACC"
printf '%s' '{"client_id":"c","client_secret":"s"}' > "$ACC/config.json"
printf '{"access_token":"tok","refresh_token":"r","expires_at":%s}' "$(( $(date +%s) + 3600 ))" > "$ACC/credentials.json"
chmod 600 "$ACC/credentials.json"

# Start a scenario: a fresh response directory and an empty log.
scenario() {
    FAKE_DIR="$TMP/fake.$1"
    rm -rf "$FAKE_DIR"; mkdir -p "$FAKE_DIR"
    export FAKE_DIR
    : > "$FAKE_CURL_LOG"
}
# respond <method> <n|default> <status> <body> [header...]
respond() {
    local method="$1" n="$2" status="$3" body="$4"; shift 4
    { printf '%s\n' "$status"; for h in "$@"; do printf '%s\n' "$h"; done; printf '\n%s' "$body"; } \
        > "$FAKE_DIR/$method.$n"
}
calls() { grep -c "^$1 " "$FAKE_CURL_LOG" || true; }
run() {  # run <script> [args...]; stdout and stderr together, exit code in $rc
    local script="$1"; shift
    out=$(HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default bash "$script" "$@" 2>&1); rc=$?
}

# Message IDs over 25 characters are used as given, with no lookup.
ID0="AAMkAGmsg0-$(printf 'a%.0s' $(seq 1 30))"
ID1="AAMkAGmsg1-$(printf 'b%.0s' $(seq 1 30))"
ID2="AAMkAGmsg2-$(printf 'c%.0s' $(seq 1 30))"

########################################
# The retry rule and the wait, as pure functions.
########################################
# shellcheck source=../scripts/lib/graph.sh
. "$SCRIPTS/lib/graph.sh"
retries() { outlook_should_retry "$1" "$2" && echo retry || echo no; }
eq "429 on POST is retried"       retry "$(retries POST 429)"
eq "429 on GET is retried"        retry "$(retries GET 429)"
eq "503 on GET is retried"        retry "$(retries GET 503)"
eq "504 on GET is retried"        retry "$(retries get 504)"
eq "503 on POST is not retried"   no    "$(retries POST 503)"
eq "504 on PATCH is not retried"  no    "$(retries PATCH 504)"
eq "500 on GET is not retried"    no    "$(retries GET 500)"
eq "no status is not retried"     no    "$(retries GET '')"
eq "Retry-After in seconds is honoured"       3  "$(outlook_retry_wait 3 0)"
eq "Retry-After over the cap is capped"       60 "$(outlook_retry_wait 300 0)"
eq "no Retry-After backs off 1s first"        1  "$(outlook_retry_wait '' 0)"
eq "no Retry-After backs off 4s on the third" 4  "$(outlook_retry_wait '' 2)"
eq "an HTTP-date Retry-After falls back"      2  "$(outlook_retry_wait 'Wed, 21 Oct 2026 07:28:00 GMT' 1)"

########################################
# api_call retries a 429, honouring Retry-After, and then succeeds.
########################################
scenario cal429
respond GET 1 429 '{"error":{"code":"TooManyRequests","message":"Too many requests"}}' 'Retry-After: 1'
respond GET 2 200 '{"value":[{"name":"Calendar","isDefaultCalendar":true,"color":"auto"}]}'
start=$(date +%s)
run "$CAL" calendars
elapsed=$(( $(date +%s) - start ))
eq "calendar: a 429 is retried once" "2" "$(calls GET)"
eq "calendar: the retry succeeds" "0" "$rc"
has "calendar: the result is the retried response" "Calendar" "$out"
has "calendar: the retry is announced on stderr" "answered HTTP 429. Retrying in 1s" "$out"
eq "calendar: it waited the Retry-After" "1" "$([ "$elapsed" -ge 1 ] && echo 1 || echo 0)"

scenario mail429
respond GET 1 429 '{"error":{"code":"ApplicationThrottled","message":"Throttled"}}' 'Retry-After: 1'
respond GET 2 200 '{"id":"ARCH"}'
respond POST 1 200 "{\"responses\":[{\"id\":\"0\",\"status\":201,\"body\":{}}]}"
run "$MAIL" batch-move archive "$ID0"
eq "mail: a throttled GET is retried" "2" "$(calls GET)"
eq "mail: the command then succeeds" "0" "$rc"
has "mail: and moves the message" "Done: 1 moved, 0 failed." "$out"

########################################
# 503 and 504 are retried on GET only, and retries are capped.
########################################
scenario get503
respond GET 1 503 '' 'Retry-After: 1'
respond GET 2 504 '' 'Retry-After: 1'
respond GET 3 200 '{"value":[{"name":"Calendar","isDefaultCalendar":true,"color":"auto"}]}'
run "$CAL" calendars
eq "a GET that gets 503 then 504 is retried twice" "3" "$(calls GET)"
has "and then succeeds" "Calendar" "$out"

scenario cap
respond GET default 429 '{"error":{"code":"TooManyRequests","message":"Still busy"}}' 'Retry-After: 0'
out=$(HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default OUTLOOK_MAX_RETRIES=2 bash "$CAL" calendars 2>&1)
eq "retries stop at OUTLOOK_MAX_RETRIES (1 try + 2 retries)" "3" "$(calls GET)"
has "the last retry is announced as the last" "retry 2 of 2" "$out"
hasnt "and there is no third" "retry 3 of" "$out"

scenario post503
respond GET default 200 '{"id":"ARCH"}'
respond POST default 503 '' 'Retry-After: 1'
run "$MAIL" batch-move archive "$ID0" "$ID1"
eq "a POST that gets a 503 is not retried" "1" "$(calls POST)"
eq "and the command exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "an empty 503 is reported, not read as success" "Microsoft Graph answered HTTP 503 with no error details." "$out"

########################################
# batch-move: a batch that fails as a whole fails every message in it.
########################################
scenario whole
respond GET default 200 '{"id":"ARCH"}'
respond POST default 400 '{"error":{"code":"BadRequest","message":"Invalid batch payload"}}'
run "$MAIL" batch-move archive "$ID0" "$ID1" "$ID2"
eq "whole-batch error: exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "whole-batch error: the summary counts every message" "Done: 0 moved, 3 failed." "$out"
has "whole-batch error: says the batch failed" "the whole batch of 3 failed: Invalid batch payload" "$out"
has "whole-batch error: names the first ID"  "FAILED $ID0" "$out"
has "whole-batch error: names the second ID" "FAILED $ID1" "$out"
has "whole-batch error: names the third ID"  "FAILED $ID2" "$out"

scenario notjson
respond GET default 200 '{"id":"ARCH"}'
respond POST default 200 '<html>gateway</html>'
run "$MAIL" batch-move archive "$ID0"
eq "a batch answer that is not JSON: exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "a batch answer that is not JSON: the message counts as failed" "Done: 0 moved, 1 failed." "$out"

########################################
# batch-move: items Graph throttled are sent again, and only those.
########################################
scenario items
respond GET default 200 '{"id":"ARCH"}'
respond POST 1 200 '{"responses":[
  {"id":"0","status":201,"body":{}},
  {"id":"1","status":429,"headers":{"Retry-After":"1"},"body":{"error":{"code":"ApplicationThrottled","message":"Mailbox concurrency"}}},
  {"id":"2","status":404,"body":{"error":{"code":"ErrorItemNotFound","message":"Not found"}}}]}'
respond POST 2 200 '{"responses":[{"id":"1","status":201,"body":{}}]}'
run "$MAIL" batch-move archive "$ID0" "$ID1" "$ID2"
eq "throttled item: the batch is sent a second time" "2" "$(calls POST)"
second=$(grep '^BODY' "$FAKE_CURL_LOG" | sed -n '2p' | sed 's/^BODY //')
eq "throttled item: the second batch holds only the throttled message" "1" \
   "$(printf '%s' "$second" | jq -r '.requests | map(.id) | join(",")')"
eq "throttled item: and asks to move the right message" "/me/messages/$ID1/move" \
   "$(printf '%s' "$second" | jq -r '.requests[0].url')"
has "throttled item: the retry is announced" "throttled: 1 message(s), retrying in 1s" "$out"
has "throttled item: the summary counts it as moved" "Done: 2 moved, 1 failed." "$out"
has "throttled item: the real failure names its ID" "FAILED $ID2 [404]: Not found" "$out"
hasnt "throttled item: the retried ID is not reported failed" "FAILED $ID1" "$out"
eq "throttled item: exits non-zero for the real failure" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"

scenario missing
respond GET default 200 '{"id":"ARCH"}'
respond POST default 200 '{"responses":[{"id":"0","status":201,"body":{}}]}'
run "$MAIL" batch-move archive "$ID0" "$ID1"
has "a message with no response in the batch counts as failed" "Done: 1 moved, 1 failed." "$out"
has "and is named" "FAILED $ID1" "$out"

########################################
# A large attachment goes up in 4 MB chunks, read with tail and head rather
# than GNU-only dd flags, so it works on macOS too. The chunks, put back
# together, are the file.
########################################
scenario chunks
# A draft ID over 100 characters is used as given, with no lookup.
DRAFT="AAMkAGdraft-$(printf 'd%.0s' $(seq 1 100))"
BIG="$TMP/big.bin"
head -c $((9 * 1048576 + 123)) /dev/urandom > "$BIG"
respond POST default 200 '{"uploadUrl":"https://upload.example.com/session/1"}'
respond PUT default 200 ''
run "$MAIL" attach "$DRAFT" "$BIG"
eq "chunked upload: exits 0" "0" "$rc"
eq "chunked upload: three PUTs for 9 MB" "3" "$(calls PUT)"
eq "chunked upload: the ranges cover the file" \
   "bytes 0-4194303/9437307 bytes 4194304-8388607/9437307 bytes 8388608-9437306/9437307" \
   "$(sed -n 's/^RANGE //p' "$FAKE_CURL_LOG" | tr '\n' ' ' | sed 's/ $//')"
cat "$FAKE_DIR"/PUT.1.body "$FAKE_DIR"/PUT.2.body "$FAKE_DIR"/PUT.3.body > "$TMP/rebuilt.bin"
eq "chunked upload: the chunks put back together are the file" "same" \
   "$(cmp -s "$BIG" "$TMP/rebuilt.bin" && echo same || echo different)"
eq "chunked upload: the upload has the long transfer timeout" "3" "$(grep -c '^PUT .* max-time=600$' "$FAKE_CURL_LOG" || true)"
eq "the script uses no GNU-only dd flags" "0" "$(grep -v '^[[:space:]]*#' "$MAIL" | grep -c 'iflag=' || true)"

########################################
# Every curl call has a timeout.
########################################
# Every call the scenarios above made went out with --max-time.
scenario timeouts
respond GET default 200 '{"id":"ARCH"}'
respond POST default 200 '{"responses":[{"id":"0","status":201,"body":{}}]}'
run "$MAIL" batch-move archive "$ID0"
run "$CAL" calendars
eq "every request the scripts made had --max-time" "0" "$(grep -c 'max-time=none' "$FAKE_CURL_LOG" || true)"
eq "and there were requests to check" "1" "$([ "$(grep -c ' max-time=' "$FAKE_CURL_LOG")" -ge 3 ] && echo 1 || echo 0)"

# And every curl command written in the scripts carries --connect-timeout and
# --max-time. Continuation lines are joined first, so a flag on the next line
# counts. A curl that is only mentioned (command -v curl, an error message, a
# comment) is not a call.
curl_calls_without_timeout() {
    awk '
        /^[[:space:]]*#/ { next }
        { line = line $0 }
        /\\$/ { sub(/\\$/, "", line); next }
        {
            if (line ~ /(^|[[:space:](|;&!])curl[[:space:]]+-/ &&
                (line !~ /--max-time/ || line !~ /--connect-timeout/))
                print FILENAME ": " line
            line = ""
        }' "$@"
}
missing=$(curl_calls_without_timeout "$SCRIPTS"/*.sh "$SCRIPTS"/lib/*.sh)
eq "every curl call in the scripts has --connect-timeout and --max-time" "" "$missing"
found=$(awk '/^[[:space:]]*#/ {next} /(^|[[:space:](|;&!])curl[[:space:]]+-/' "$SCRIPTS"/*.sh "$SCRIPTS"/lib/*.sh | wc -l)
eq "the check found curl calls to look at" "1" "$([ "$found" -ge 5 ] && echo 1 || echo 0)"
# The checker itself catches a call with no timeout.
printf 'x=$(curl -s -X GET "https://example.com" \\\n    -H "a: b")\n' > "$TMP/bad.sh"
eq "the check catches a curl call with no timeout" "1" \
   "$([ -n "$(curl_calls_without_timeout "$TMP/bad.sh")" ] && echo 1 || echo 0)"

echo
echo "PASS=$PASS FAIL=$FAIL"
[ "$FAIL" -eq 0 ]
