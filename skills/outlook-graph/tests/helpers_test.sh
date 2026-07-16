#!/bin/bash
# Offline unit tests for the pure/logic helpers in outlook-graph-mail.sh.
#
# These extract the real functions from the script and exercise them with a
# mocked api_call + date, so no Microsoft account or network is required.
# They cover: URL-encoding, KQL detection, search paging/sort/cap, folder
# resolution (BFS + Parent/Child paths), and the token-expiry decision.
#
#   bash skills/outlook-graph/tests/helpers_test.sh
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
MAIL="$SCRIPT_DIR/scripts/outlook-graph-mail.sh"
CAL="$SCRIPT_DIR/scripts/outlook-graph-calendar.sh"
GRAPH_URL="https://graph.microsoft.com/v1.0"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }

# Pull a function definition (from `name() {` to the first line that is just `}`)
# out of the live script so the tests track the real implementation.
extract_fn() { awk "/^$1\\(\\) \\{/{f=1} f{print} f&&/^\\}/{exit}" "$MAIL"; }

extract_cal_fn() { awk "/^$1\\(\\) \\{/{f=1} f{print} f&&/^\\}/{exit}" "$CAL"; }

eval "$(extract_fn urlencode)"
eval "$(extract_fn run_message_search)"
eval "$(extract_fn _find_folder_by_name)"
eval "$(extract_fn _resolve_folder_path)"
eval "$(extract_fn resolve_folder_id)"
eval "$(extract_fn recipients_to_json)"
eval "$(extract_fn sendable_addresses)"
eval "$(extract_fn address_in_list)"
eval "$(extract_fn warn_if_not_sendable)"
eval "$(extract_fn from_to_json)"
eval "$(extract_fn md_to_html)"
eval "$(extract_cal_fn attendees_to_json)"
eval "$(extract_cal_fn resolve_event_id)"

########################################
# recipients_to_json / attendees_to_json
########################################
eq "recipients split+trim" "a@x.com,b@y.com,c@z.com" \
   "$(recipients_to_json ' a@x.com, b@y.com ; c@z.com ' | jq -r '[.[].emailAddress.address]|join(",")')"
eq "recipients drops empties" "1" \
   "$(recipients_to_json 'a@x.com,, ;' | jq 'length')"
eq "attendees default required" "a@x.com|required" \
   "$(attendees_to_json 'a@x.com' | jq -r '.[0] | "\(.emailAddress.address)|\(.type)"')"
eq "attendees typed optional" "optional" \
   "$(attendees_to_json 'a@x.com' optional | jq -r '.[0].type')"
eq "attendees empty -> []" "0" "$(attendees_to_json '' | jq 'length')"

########################################
# html_to_text: the `read` renderer. Block tags MUST become line breaks before
# tags are stripped, or paragraphs run together ("Hello,This is...") and list
# items merge - which would garble the very thing the skill insists on reading
# end-to-end.
########################################
# Eval the real assignment out of the script (rather than scraping the raw text,
# which still carries its shell quote-escaping) so the tests exercise the exact
# jq program the script runs.
eval "$(sed -n "/^HTML_TO_TEXT='/,/^'\$/p" "$MAIL")"
htt() { jq -rn --arg h "$1" "$HTML_TO_TEXT"' $h | html_to_text'; }

eq "html_to_text separates paragraphs" $'Hello,\nWorld' \
   "$(htt '<p>Hello,</p><p>World</p>')"
eq "html_to_text bullets each on own line" $'- one\n- two' \
   "$(htt '<ul><li>one</li><li>two</li></ul>')"
eq "html_to_text honours <br>" $'Regards,\nDan' \
   "$(htt 'Regards,<br>Dan')"
eq "html_to_text decodes entities" 'A & B < C > D' \
   "$(htt '<p>A &amp; B &lt; C &gt; D</p>')"
eq "html_to_text drops style/script content" 'Body' \
   "$(htt '<style>p{color:red}</style><p>Body</p>')"
eq "html_to_text collapses blank runs" $'A\n\nB' \
   "$(htt '<p>A</p><div></div><div></div><p>B</p>')"
eq "html_to_text empty body -> empty" "" "$(htt '')"

########################################
# md_to_html (skipped when pandoc is unavailable)
########################################
if command -v pandoc >/dev/null 2>&1; then
    FONT_STACK="'Aptos', 'Aptos Display', 'Segoe UI', Roboto, sans-serif"
    html=$(md_to_html $'First para\n\nSecond **bold** para')
    case "$html" in *"'Aptos'"*) eq "md_to_html uses Aptos stack" ok ok;; *) eq "md_to_html uses Aptos stack" "contains 'Aptos'" "$html";; esac
    case "$html" in *'<p style="margin: 0 0 14px 0;">'*) eq "md_to_html inlines <p> margins" ok ok;; *) eq "md_to_html inlines <p> margins" "contains <p style=" "$html";; esac
    case "$html" in *'line-height: 1.5'*) eq "md_to_html default line-height" ok ok;; *) eq "md_to_html default line-height" "1.5" "$html";; esac
    case "$(md_to_html 'x' 1.6)" in *'line-height: 1.6'*) eq "md_to_html reply line-height" ok ok;; *) eq "md_to_html reply line-height" "1.6" "?";; esac
else
    echo "skip - md_to_html tests (pandoc not installed)"
fi

########################################
# resolve_event_id: passthrough, upcoming hit, past fallback, miss
########################################
api_call() {
    local endpoint="$2"
    case "$endpoint" in
      "/me/calendar/events?"*) echo '{"value":[{"id":"AAAAlongupcomingevent111"},{"id":"AAAAlongupcomingevent222"}]}';;
      "/me/events?"*)          echo '{"value":[{"id":"BBBBlongpasteventXYZ99999"}]}';;
      *) echo '{"value":[]}';;
    esac
}
today_start() { echo "2026-01-01T00:00:00Z"; }
LONG_ID="AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
eq "event long id passthrough" "$LONG_ID" "$(resolve_event_id "$LONG_ID")"
eq "event short id upcoming"   "AAAAlongupcomingevent222" "$(resolve_event_id 'event222')"
eq "event short id past fallback" "BBBBlongpasteventXYZ99999" "$(resolve_event_id 'XYZ99999')"
eq "event id miss -> rc1" "1" "$(resolve_event_id 'nomatch' >/dev/null; echo $?)"

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
#   top-level: Deleted Items(DI) Inbox(IB) Archive(AR) Clients(CL)
#   IB->Projects(PR)->Acme(A1) ; CL->Acme(A2)
#   DI->Projects(GHOST) - a deleted folder keeps its name in the bin, and Graph
#   lists Deleted Items BEFORE Inbox, so a bare name must NOT resolve to it.
########################################
api_call() {
    local endpoint="$2"
    case "$endpoint" in
      "/me/mailFolders?\$top=200") echo '{"value":[{"displayName":"Deleted Items","id":"DI"},{"displayName":"Inbox","id":"IB"},{"displayName":"Archive","id":"AR"},{"displayName":"Clients","id":"CL"}]}';;
      "/me/mailFolders/DI/childFolders?\$top=200") echo '{"value":[{"displayName":"Projects","id":"GHOST"},{"displayName":"Acme","id":"GHOST2"}]}';;
      "/me/mailFolders/IB/childFolders?\$top=200") echo '{"value":[{"displayName":"Projects","id":"PR"}]}';;
      "/me/mailFolders/PR/childFolders?\$top=200") echo '{"value":[{"displayName":"Acme","id":"A1"}]}';;
      "/me/mailFolders/CL/childFolders?\$top=200") echo '{"value":[{"displayName":"Acme","id":"A2"}]}';;
      "/me/mailFolders/inbox") echo '{"id":"IB"}';;
      "/me/mailFolders/deleteditems?\$select=id") echo '{"id":"DI"}';;
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
# The regression that matters: a same-named ghost in the bin must never win.
eq "bare name skips Deleted Items ghost" "PR" "$(resolve_folder_id 'Projects')"
eq "bare name skips bin even when only match" "1" \
   "$(api_call() { case "$2" in "/me/mailFolders?\$top=200") echo '{"value":[{"displayName":"Deleted Items","id":"DI"},{"displayName":"Inbox","id":"IB"}]}';; "/me/mailFolders/DI/childFolders?\$top=200") echo '{"value":[{"displayName":"Old Project","id":"GHOST"}]}';; "/me/mailFolders/deleteditems?\$select=id") echo '{"id":"DI"}';; *) echo '{"value":[]}';; esac; }; resolve_folder_id 'Old Project' >/dev/null; echo $?)"
eq "explicit bin path still resolves" "GHOST" \
   "$(api_call() { case "$2" in "/me/mailFolders?\$top=200") echo '{"value":[{"displayName":"Deleted Items","id":"DI"},{"displayName":"Inbox","id":"IB"}]}';; "/me/mailFolders/DI/childFolders?\$top=200") echo '{"value":[{"displayName":"Old Project","id":"GHOST"}]}';; "/me/mailFolders/deleteditems?\$select=id") echo '{"id":"DI"}';; *) echo '{"value":[]}';; esac; }; resolve_folder_id 'Deleted Items/Old Project')"

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
eval "$(extract_fn _api_call)"
eval "$(extract_fn api_call)"
eval "$(extract_fn api_call_file)"
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

########################################
# api_call_file: the attachment path streams its body from a file rather than an
# argv string, but must inherit the SAME retry/error behaviour as api_call. It
# is the call most exposed to a mid-run token expiry (a slow upload, issued
# after the draft is created), and an empty body from a dropped token would be
# read as success by callers that only check for ".error".
########################################
body_file=$(mktemp); printf '%s' '{"name":"big.pdf"}' > "$body_file"

_graph_request_file() { return 0; }                  # empty body, success (204)
unset OUTLOOK_TOKEN_RETRIED
eq "api_call_file empty-204 stays empty" "" "$(api_call_file POST /x "$body_file")"

_graph_request_file() { return 7; }                  # transport failure, empty
unset OUTLOOK_TOKEN_RETRIED
eq "api_call_file transport-fail -> NetworkError" "NetworkError" \
   "$(api_call_file POST /x "$body_file" | jq -r '.error.code')"

# Auth error, then success: the retry must re-read the body file and use the
# refreshed token.
echo 0 > /tmp/outlook_test_calls
_graph_request_file() {
    local n; n=$(cat /tmp/outlook_test_calls); n=$((n+1)); echo "$n" > /tmp/outlook_test_calls
    if [ "$n" = 1 ]; then printf '%s' '{"error":{"code":"InvalidAuthenticationToken","message":"expired"}}'
    else jq -c --arg tok "$ACCESS_TOKEN" '{name: .name, token: $tok}' < "$3"; fi
    return 0
}
unset OUTLOOK_TOKEN_RETRIED
retry_out=$(api_call_file POST /x "$body_file")
eq "api_call_file auth-retry re-reads body file" "big.pdf" "$(printf '%s' "$retry_out" | jq -r '.name')"
eq "api_call_file auth-retry uses refreshed token" "newtok" "$(printf '%s' "$retry_out" | jq -r '.token')"

########################################
# Send-as-alias: sendable_addresses / is_sendable_address / from_to_json.
# Graph marks the primary address with an uppercase "SMTP:" prefix and aliases
# with lowercase "smtp:". Getting that casing rule wrong would either hide every
# alias or mistake an alias for the primary, so it is pinned here.
########################################
api_call() {
    echo '{"mail":"dan@example.com","userPrincipalName":"dan@example.onmicrosoft.com","proxyAddresses":["smtp:alias1@example.com","SMTP:dan@example.com","X500:/o=ExchangeLabs/cn=Recipients/cn=abc","sip:dan@example.com","smtp:alias2@other.co.uk"]}'
}
eq "sendable_addresses primary first, then aliases" \
   "dan@example.com,alias1@example.com,alias2@other.co.uk" \
   "$(sendable_addresses | paste -sd, -)"
eq "sendable_addresses drops X500/sip entries" "3" "$(sendable_addresses | wc -l)"
eq "address_in_list matches alias" "0" \
   "$(sendable_addresses | address_in_list 'alias2@other.co.uk'; echo $?)"
eq "address_in_list is case-insensitive" "0" \
   "$(sendable_addresses | address_in_list 'Alias1@EXAMPLE.com'; echo $?)"
eq "address_in_list rejects unknown" "1" \
   "$(sendable_addresses | address_in_list 'nope@example.com'; echo $?)"
# grep -F, not a regex: a '.' in a domain must not match any character, or
# 'aliasX@other.co.uk' would masquerade as a known address.
eq "address_in_list treats input as literal, not regex" "1" \
   "$(sendable_addresses | address_in_list 'aliasX@other.co.uk'; echo $?)"
eq "address_in_list rejects a substring of a known address" "1" \
   "$(sendable_addresses | address_in_list 'alias1@example.co'; echo $?)"

# A mailbox with no proxyAddresses at all (some tenants) must still report the
# one address it can send as, rather than an empty list.
api_call() { echo '{"mail":"solo@example.com","userPrincipalName":"solo@example.onmicrosoft.com"}'; }
eq "sendable_addresses falls back to .mail" "solo@example.com" "$(sendable_addresses | paste -sd, -)"
api_call() { echo '{"userPrincipalName":"upn-only@example.com"}'; }
eq "sendable_addresses falls back to UPN" "upn-only@example.com" "$(sendable_addresses | paste -sd, -)"

# Aliases but NO uppercase-tagged primary. The primary must still appear, and
# must lead - otherwise `aliases` crowns the first ALIAS "(primary)" and
# `update from <primary>` warns that the real primary is not your address.
api_call() { echo '{"mail":"real@example.com","proxyAddresses":["smtp:alias@example.com"]}'; }
eq "sendable_addresses keeps primary when no SMTP: tag" \
   "real@example.com,alias@example.com" "$(sendable_addresses | paste -sd, -)"

# The primary is routinely repeated as a lowercase smtp: entry; listing it twice
# would render it twice in `aliases`.
api_call() { echo '{"mail":"dan@example.com","proxyAddresses":["SMTP:dan@example.com","smtp:DAN@example.com","smtp:alias@example.com"]}'; }
eq "sendable_addresses dedups primary case-insensitively" \
   "dan@example.com,alias@example.com" "$(sendable_addresses | paste -sd, -)"

# A failed lookup must yield an empty list, so warn_if_not_sendable can say
# "could not check" rather than blaming a valid alias for a network blip.
api_call() { echo '{"error":{"code":"NetworkError","message":"boom"}}'; }
eq "sendable_addresses empty on Graph error" "" "$(sendable_addresses)"
eq "warn_if_not_sendable says 'not checked' on lookup failure" "1" \
   "$(warn_if_not_sendable 'real@example.com' 2>&1 >/dev/null | grep -c 'has not been checked')"
eq "warn_if_not_sendable does NOT blame the address on lookup failure" "0" \
   "$(warn_if_not_sendable 'real@example.com' 2>&1 >/dev/null | grep -c 'is not one of')"

# Sending name:"" makes Outlook render the bare address instead of the mailbox
# display name, so a blank name must be omitted from the payload entirely.
eq "from_to_json omits blank name" '{"emailAddress":{"address":"a@x.com"}}' \
   "$(from_to_json 'a@x.com' | jq -c .)"
eq "from_to_json omits name when unset" '{"emailAddress":{"address":"a@x.com"}}' \
   "$(from_to_json 'a@x.com' '' | jq -c .)"
eq "from_to_json includes name when given" '{"emailAddress":{"address":"a@x.com","name":"Dan G"}}' \
   "$(from_to_json 'a@x.com' 'Dan G' | jq -c .)"

rm -f "$body_file" /tmp/outlook_test_last_url /tmp/outlook_test_calls
echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
