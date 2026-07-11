#!/bin/bash
# Offline unit tests for the pure/logic helpers in outlook-mail.sh.
#
# These extract the real functions from the script and exercise them with a
# mocked api_call + date, so no Microsoft account or network is required.
# They cover: URL-encoding, KQL detection, search paging/sort/cap, folder
# resolution (BFS + Parent/Child paths), and the token-expiry decision.
#
#   bash skills/outlook/tests/helpers_test.sh
#
# Requires: jq, grep, awk (same tools the skill itself uses).
#
# The literal '$search=...' strings below are expected URLs, not expansions
# (SC2016); the mocked api_call definitions are invoked indirectly by the
# extracted functions (SC2317); ACCESS_TOKEN is read inside the eval'd api_call
# (SC2034). All intentional in this harness.
# shellcheck disable=SC2016,SC2317,SC2034
set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
MAIL="$SCRIPT_DIR/scripts/outlook-mail.sh"
GRAPH_URL="https://graph.microsoft.com/v1.0"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }

# Pull a function definition (from `name() {` to the first line that is just `}`)
# out of the live script so the tests track the real implementation.
extract_fn() { awk "/^$1\\(\\) \\{/{f=1} f{print} f&&/^\\}/{exit}" "$MAIL"; }

eval "$(extract_fn urlencode)"
eval "$(extract_fn run_message_search)"
eval "$(extract_fn _find_folder_by_name)"
eval "$(extract_fn _resolve_folder_path)"
eval "$(extract_fn resolve_folder_id)"

########################################
# urlencode
########################################
eq "urlencode space"     "a%20b"             "$(urlencode 'a b')"
eq "urlencode ampersand" "Q3%20%26%20Q4"     "$(urlencode 'Q3 & Q4')"
eq "urlencode colon"     "subject%3Ainvoice" "$(urlencode 'subject:invoice')"
eq "urlencode at/plus"   "a%2Bb%40c"         "$(urlencode 'a+b@c')"

########################################
# run_message_search: paging + newest-first sort + cap + encoding + errors
########################################
# Page 1 returns two messages + a nextLink; page 2 returns two more, no link.
api_call() {
    local endpoint="$2"; printf '%s' "$endpoint" > /tmp/outlook_test_last_url
    if [[ "$endpoint" == *skiptoken* ]]; then
        echo '{"value":[{"id":"m3","receivedDateTime":"2026-04-01T00:00:00Z"},{"id":"m4","receivedDateTime":"2026-01-15T00:00:00Z"}]}'
    else
        echo '{"@odata.nextLink":"'"$GRAPH_URL"'/me/messages?$skiptoken=ABC","value":[{"id":"m1","receivedDateTime":"2026-03-01T00:00:00Z"},{"id":"m2","receivedDateTime":"2026-02-01T00:00:00Z"}]}'
    fi
}
eq "search paged, newest-first" "m3,m1,m2,m4" "$(run_message_search 'project' 10 | jq -r '[.value[].id]|join(",")')"
eq "search cap stops paging"    "2"           "$(run_message_search 'project' 2 | jq '.value|length')"
run_message_search 'Q3 & Q4' 1 >/dev/null
eq "search free-text encoded+quoted" \
   '/me/messages?$search=%22Q3%20%26%20Q4%22&$top=1&$select=id,subject,from,receivedDateTime,isRead,bodyPreview' \
   "$(cat /tmp/outlook_test_last_url)"
run_message_search 'subject:invoice' 1 >/dev/null
eq "search KQL wrapped in quotes (colon encoded)" \
   '/me/messages?$search=%22subject%3Ainvoice%22&$top=1&$select=id,subject,from,receivedDateTime,isRead,bodyPreview' \
   "$(cat /tmp/outlook_test_last_url)"
run_message_search '"already quoted"' 1 >/dev/null
eq "search does not double-quote" \
   '/me/messages?$search=%22already%20quoted%22&$top=1&$select=id,subject,from,receivedDateTime,isRead,bodyPreview' \
   "$(cat /tmp/outlook_test_last_url)"
api_call() { echo '{"error":{"code":"BadRequest","message":"nope"}}'; }
eq "search propagates error" "nope" "$(run_message_search 'x' 5 | jq -r '.error.message')"

########################################
# resolve_folder_id: BFS + path (mock tree)
#   top-level: Inbox(IB) Archive(AR) Clients(CL)
#   IB->Projects(PR)->Acme(A1) ; CL->Acme(A2)
########################################
api_call() {
    local endpoint="$2"
    case "$endpoint" in
      "/me/mailFolders?\$top=200") echo '{"value":[{"displayName":"Inbox","id":"IB"},{"displayName":"Archive","id":"AR"},{"displayName":"Clients","id":"CL"}]}';;
      "/me/mailFolders/IB/childFolders?\$top=200") echo '{"value":[{"displayName":"Projects","id":"PR"}]}';;
      "/me/mailFolders/PR/childFolders?\$top=200") echo '{"value":[{"displayName":"Acme","id":"A1"}]}';;
      "/me/mailFolders/CL/childFolders?\$top=200") echo '{"value":[{"displayName":"Acme","id":"A2"}]}';;
      "/me/mailFolders/inbox") echo '{"id":"IB"}';;
      *) echo '{"value":[]}';;
    esac
}
eq "folder top-level"            "CL" "$(resolve_folder_id 'Clients')"
eq "folder case-insensitive"     "PR" "$(resolve_folder_id 'projects')"
eq "folder well-known alias"     "IB" "$(resolve_folder_id 'Inbox')"
eq "folder ambiguous->shallowest" "A2" "$(resolve_folder_id 'Acme')"
eq "folder path Clients/Acme"    "A2" "$(resolve_folder_id 'Clients/Acme')"
eq "folder path Inbox/Projects/Acme" "A1" "$(resolve_folder_id 'Inbox/Projects/Acme')"
eq "folder not found -> rc1"     "1"  "$(resolve_folder_id 'Nope' >/dev/null; echo $?)"

########################################
# token-expiry decision (mirrors ensure_valid_token's local check)
########################################
NOW=1000000
decide() { if [ -n "$1" ] && [ "$NOW" -lt "$(( $2 - 60 ))" ]; then echo cached; else echo refresh; fi; }
eq "token fresh -> cached"        "cached"  "$(decide tok 1000100)"
eq "token within margin -> refresh" "refresh" "$(decide tok 1000030)"
eq "token expired -> refresh"     "refresh" "$(decide tok 999000)"
eq "token missing -> refresh"     "refresh" "$(decide '' 1000100)"
eq "token no expires_at -> refresh" "refresh" "$(decide tok 0)"

########################################
# api_call: empty 204/202 success must stay empty (NOT become an error);
# only a transport failure (curl non-zero) becomes a NetworkError; an
# InvalidAuthenticationToken body triggers one refresh + retry.
########################################
eval "$(extract_fn api_call)"
ACCESS_TOKEN="tok"
refresh_access_token() { echo "newtok"; }

_graph_request() { return 0; }                       # empty body, success (204)
unset OUTLOOK_TOKEN_RETRIED
empty_out=$(api_call GET /x)
eq "api_call empty-204 stays empty" "" "$empty_out"
eq "api_call empty-204 is not an error" "false" \
   "$(printf '%s' "$empty_out" | jq -e 'has("error")' >/dev/null 2>&1 && echo true || echo false)"

_graph_request() { return 7; }                       # transport failure, empty
unset OUTLOOK_TOKEN_RETRIED
eq "api_call transport-fail -> NetworkError" "NetworkError" "$(api_call GET /x | jq -r '.error.code')"

_graph_request() { printf '%s' '{"value":[1,2,3]}'; return 0; }   # normal body
unset OUTLOOK_TOKEN_RETRIED
eq "api_call passthrough" "3" "$(api_call GET /x | jq '.value|length')"

echo 0 > /tmp/outlook_test_calls                     # auth error then success
_graph_request() {
    local n; n=$(cat /tmp/outlook_test_calls); n=$((n+1)); echo "$n" > /tmp/outlook_test_calls
    if [ "$n" = 1 ]; then printf '%s' '{"error":{"code":"InvalidAuthenticationToken","message":"expired"}}'
    else printf '%s' '{"value":["ok"]}'; fi
    return 0
}
unset OUTLOOK_TOKEN_RETRIED
eq "api_call auth-retry recovers" "ok" "$(api_call GET /x | jq -r '.value[0]')"

rm -f /tmp/outlook_test_last_url /tmp/outlook_test_calls
echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
