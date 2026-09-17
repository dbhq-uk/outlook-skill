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
CAL="$SCRIPT_DIR/scripts/outlook-calendar.sh"
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
eval "$(extract_fn category_colour_to_preset)"
eval "$(extract_fn category_json)"
eval "$(extract_fn resolve_category_id)"
eval "$(extract_fn categories_add)"
eval "$(extract_fn categories_remove)"

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

########################################
# export helpers: staging filename + --since validation.
# The filename only has to be unique and sortable - outlook_to_md.py derives the
# archive folder name from the message's own headers, not from this name.
########################################
eval "$(extract_fn export_eml_filename)"
eval "$(extract_fn export_since_filter)"

eq "export filename from Graph timestamp" "20260729_101200_AAAAAAAAAAAAAAAAAAAA.eml" \
   "$(export_eml_filename '2026-07-29T10:12:00Z' 'PREFIXAAAAAAAAAAAAAAAAAAAA')"
eq "export filename drops fractional seconds" "20260729_101200_abcdefghijklmnopqrst.eml" \
   "$(export_eml_filename '2026-07-29T10:12:00.5230000Z' 'abcdefghijklmnopqrst')"
# The id is server-supplied; a '/' in it would escape the output directory.
case "$(export_eml_filename '2026-07-29T10:12:00Z' 'aaaa/bbbb/cccc/dddd/eeee')" in
  */*) eq "export filename strips path separators" "no slash" "contains a slash";;
  *)   eq "export filename strips path separators" ok ok;;
esac

eq "since filter builds a Graph filter" "receivedDateTime ge 2026-07-01T00:00:00Z" \
   "$(export_since_filter '2026-07-01')"
# A malformed date must fail here. Sent to Graph it comes back as a generic
# BadRequest, or worse silently matches nothing and looks like "no new mail".
eq "since rejects a human-written date" "1" "$(export_since_filter '1 July 2026' 2>/dev/null; echo $?)"
eq "since rejects an empty value"       "1" "$(export_since_filter '' 2>/dev/null; echo $?)"
eq "since rejects a partial date"       "1" "$(export_since_filter '2026-07' 2>/dev/null; echo $?)"

########################################
# export_list_messages: paging + cap + --since encoding + error propagation.
# $top caps at 1000 per page and a mail folder can hold far more, so paging is
# required rather than optional.
########################################
eval "$(extract_fn export_list_messages)"

api_call() {
    local endpoint="$2"; printf '%s' "$endpoint" > /tmp/outlook_test_last_url
    if [[ "$endpoint" == *skiptoken* ]]; then
        echo '{"value":[{"id":"m3","receivedDateTime":"2026-04-01T00:00:00Z"}]}'
    else
        echo '{"@odata.nextLink":"'"$GRAPH_URL"'/me/messages?$skiptoken=ABC","value":[{"id":"m1","receivedDateTime":"2026-03-01T00:00:00Z"},{"id":"m2","receivedDateTime":"2026-02-01T00:00:00Z"}]}'
    fi
}
eq "export pages through nextLink" "3" "$(export_list_messages FID '' 100 | jq '.value|length')"
eq "export cap stops paging"       "2" "$(export_list_messages FID '' 2 | jq '.value|length')"
eq "export sorts newest-first" "m3,m1,m2" \
   "$(export_list_messages FID '' 100 | jq -r '[.value[].id]|join(",")')"

export_list_messages FID '2026-07-01' 1 >/dev/null
eq "export encodes --since into \$filter" "1" \
   "$(grep -c 'receivedDateTime%20ge%202026-07-01T00%3A00%3A00Z' /tmp/outlook_test_last_url)"
eq "export scopes the query to the folder" "1" \
   "$(grep -c '/me/mailFolders/FID/messages' /tmp/outlook_test_last_url)"

# A bad --since must stop before any request is issued: capture the URL from
# the last real call, run the rejected one, then check nothing overwrote it.
last_url_before=$(cat /tmp/outlook_test_last_url)
eq "export rejects a bad --since without calling Graph" "1" \
   "$(export_list_messages FID 'last tuesday' 5 >/dev/null 2>&1; echo $?)"
eq "export rejects a bad --since without calling Graph (no request issued)" \
   "$last_url_before" "$(cat /tmp/outlook_test_last_url)"

api_call() { echo '{"error":{"code":"BadRequest","message":"nope"}}'; }
eq "export propagates a Graph error as rc1" "1" \
   "$(export_list_messages FID '' 10 >/dev/null 2>&1; echo $?)"
eq "export reports the Graph error message" "1" \
   "$(export_list_messages FID '' 10 2>&1 >/dev/null | grep -c 'nope')"

########################################
# category_colour_to_preset: Graph stores colours as opaque presetN values.
# Both a friendly name and a raw preset are accepted, because presetN means
# nothing to a reader and a name cannot reach a preset Microsoft adds later.
########################################
eq "colour name maps to preset" "preset0" "$(category_colour_to_preset red)"
eq "colour name is case-insensitive" "preset0" "$(category_colour_to_preset RED)"
eq "colour name ignores spaces" "preset22" "$(category_colour_to_preset 'dark blue')"
eq "colour grey spelling" "preset12" "$(category_colour_to_preset grey)"
eq "colour gray spelling" "preset12" "$(category_colour_to_preset gray)"
eq "raw preset passes through" "preset7" "$(category_colour_to_preset preset7)"
eq "out-of-range preset rejected" "1" "$(category_colour_to_preset preset25 >/dev/null; echo $?)"
eq "unknown colour rejected" "1" "$(category_colour_to_preset mauve >/dev/null; echo $?)"

########################################
# category_json / resolve_category_id: Graph addresses a master category by
# GUID, so a display name must be resolved first. The match is exact and
# case-insensitive, deliberately NOT a substring match: renaming or deleting
# the wrong category on a fuzzy match cannot be undone. The fixture's second
# entry contains the first one's name, so a substring implementation fails here.
########################################
mock_cats() {
    api_call() { echo '{"value":[{"id":"C1","displayName":"Follow up","color":"preset0"},{"id":"C2","displayName":"Follow up later","color":"preset4"}]}'; }
}
eq "resolve category by exact name" "C1" "$(mock_cats; resolve_category_id 'Follow up')"
eq "resolve category is case-insensitive" "C1" "$(mock_cats; resolve_category_id 'FOLLOW UP')"
eq "resolve category does not substring match" "" "$(mock_cats; resolve_category_id 'Follow')"
eq "resolve category absent is empty" "" "$(mock_cats; resolve_category_id 'Nope')"
eq "category_json carries the colour" "preset4" "$(mock_cats; category_json 'Follow up later' | jq -r '.color')"

# category_json / resolve_category_id on an unreadable master list must fail
# CLOSED: a non-zero exit and the error body on stdout, never the same empty
# "not found" shape a genuinely absent category produces. (This replaces an
# earlier "exits cleanly" / exit-0 assertion: that pinned the crash-guard fix's
# shape at the time, not the requirement - an API error reported as success is
# exactly the bug this pins against now.)
mock_cat_error() { api_call() { echo '{"error":{"code":"NetworkError","message":"boom"}}'; }; }
eq "category_json on API error exits non-zero" "1" \
   "$(mock_cat_error; category_json 'Follow up' >/dev/null 2>&1; echo $?)"
eq "category_json on API error prints the error body, not empty" "NetworkError" \
   "$(mock_cat_error; category_json 'Follow up' 2>/dev/null | jq -r '.error.code')"
eq "resolve_category_id on API error exits non-zero" "1" \
   "$(mock_cat_error; resolve_category_id 'Follow up' >/dev/null 2>&1; echo $?)"
eq "resolve_category_id on API error carries the error body, not an empty id" "NetworkError" \
   "$(mock_cat_error; resolve_category_id 'Follow up' 2>/dev/null | jq -r '.error.code')"

########################################
# categories_add / categories_remove: `categorize` replaces the message's whole
# category list, so adding one label meant every caller had to read-modify-write
# by hand. These do it once, correctly: order is preserved, nothing the caller
# did not name is touched, and both are no-ops when there is nothing to do.
########################################
eq "add to empty list" '["Follow up"]' \
   "$(echo '[]' | categories_add 'Follow up')"
eq "add preserves existing and order" '["Invoices","Project X","Follow up"]' \
   "$(echo '["Invoices","Project X"]' | categories_add 'Follow up')"
eq "add of a present category does not duplicate" '["Invoices","Follow up"]' \
   "$(echo '["Invoices","Follow up"]' | categories_add 'Follow up')"
eq "add is case-insensitive about duplicates" '["Follow up"]' \
   "$(echo '["Follow up"]' | categories_add 'FOLLOW UP')"
eq "remove takes only the named one" '["Invoices","Project X"]' \
   "$(echo '["Invoices","Follow up","Project X"]' | categories_remove 'Follow up')"
eq "remove is case-insensitive" '["Invoices"]' \
   "$(echo '["Invoices","Follow up"]' | categories_remove 'FOLLOW UP')"
eq "remove of an absent category is a no-op" '["Invoices"]' \
   "$(echo '["Invoices"]' | categories_remove 'Nope')"

########################################
# CLI-level integration: run the real script as a subprocess, so the
# `categorize` dispatcher's flag handling and every call site that reads the
# master category list are proven end-to-end - not just at the unit level,
# where the dispatcher's `case "$3" in ... esac` is not an extractable
# function. A throwaway HOME provides a valid, non-expired token (so no
# network call is needed to refresh it), and `curl` is shadowed with a bash
# function - exported so the child `bash "$MAIL"` process inherits it in
# place of the real binary - that logs every request and answers from small
# fixture files a test can swap between "ok" and "API error".
########################################
CLI_TMP=$(mktemp -d)
CLI_HOME="$CLI_TMP/home"
CLI_LOG="$CLI_TMP/curl.log"
CLI_MASTERCATS="$CLI_TMP/mastercats.json"
CLI_MASTERCATS_POST="$CLI_TMP/mastercats_post.json"
CLI_MASTERCATS_PATCH="$CLI_TMP/mastercats_patch.json"
CLI_CURRENTCATS="$CLI_TMP/currentcats.json"
mkdir -p "$CLI_HOME/.dbhq/outlook/default"
printf '%s' '{"client_id":"test-client","client_secret":"test-secret"}' \
    > "$CLI_HOME/.dbhq/outlook/default/config.json"
printf '%s' '{"access_token":"test-token","refresh_token":"test-refresh","expires_at":9999999999}' \
    > "$CLI_HOME/.dbhq/outlook/default/credentials.json"
chmod 600 "$CLI_HOME/.dbhq/outlook/default/credentials.json"
export CLI_LOG CLI_MASTERCATS CLI_MASTERCATS_POST CLI_MASTERCATS_PATCH CLI_CURRENTCATS

# >100 chars so resolve_message_id's "looks like a full ID" short-circuit
# returns it unchanged - no cache file or listing API call needed to resolve it.
CLI_MSG_ID=$(printf 'M%.0s' $(seq 1 110))

curl() {
    local url="" method="GET" data="" prev="" arg
    for arg in "$@"; do
        case "$prev" in
            -X) method="$arg" ;;
            -d) data="$arg" ;;
        esac
        case "$arg" in
            https://graph.microsoft.com/*) url="$arg" ;;
        esac
        prev="$arg"
    done
    printf '%s %s\n' "$method" "$url" >> "$CLI_LOG"
    case "$url" in
        *'/me/outlook/masterCategories')
            case "$method" in
                GET) cat "$CLI_MASTERCATS" ;;
                POST) cat "$CLI_MASTERCATS_POST" ;;
                *) printf '{}' ;;
            esac
            ;;
        *'/me/outlook/masterCategories/'*)
            case "$method" in
                PATCH) cat "$CLI_MASTERCATS_PATCH" ;;
                DELETE) printf '' ;;
                *) printf '{}' ;;
            esac
            ;;
        *'$select=categories'*)
            cat "$CLI_CURRENTCATS"
            ;;
        *'/me/messages/'*)
            if [ "$method" = "PATCH" ] && [ -n "$data" ]; then
                printf '%s' "$data" | jq -c '{categories: (.categories // [])}'
            else
                printf '{}'
            fi
            ;;
        *) printf '{}' ;;
    esac
}
export -f curl

run_mail_cli() { HOME="$CLI_HOME" OUTLOOK_ACCOUNT=default bash "$MAIL" "$@"; }

# Captures stdout/stderr/exit code of a run_mail_cli call into CLI_OUT/CLI_ERR/CLI_RC.
run_and_capture() {
    local outfile errfile
    outfile=$(mktemp); errfile=$(mktemp)
    run_mail_cli "$@" > "$outfile" 2> "$errfile"
    CLI_RC=$?
    CLI_OUT=$(cat "$outfile"); CLI_ERR=$(cat "$errfile")
    rm -f "$outfile" "$errfile"
}

contains() { case "$1" in *"$2"*) return 0 ;; *) return 1 ;; esac; }

########################################
# A near-miss flag (wrong case, typo, unsupported verb) must be
# rejected - never fall through to the replace form, which would PATCH the
# message's categories to a single-element list containing the flag text
# itself, wiping every real category while still reporting success.
########################################
: > "$CLI_LOG"
run_and_capture categorize "$CLI_MSG_ID" --Add "Follow up"
eq "categorize --Add (wrong case) is rejected" "1" "$CLI_RC"
if contains "$CLI_ERR" "--add" && contains "$CLI_ERR" "--remove"; then
    eq "categorize --Add error names the two valid flags" ok ok
else
    eq "categorize --Add error names the two valid flags" "mentions --add and --remove" "$CLI_ERR"
fi
eq "categorize --Add makes no API call before rejecting" "" "$(cat "$CLI_LOG")"

: > "$CLI_LOG"
run_and_capture categorize "$CLI_MSG_ID" --rm "Follow up"
eq "categorize --rm (unsupported verb) is rejected" "1" "$CLI_RC"
eq "categorize --rm makes no API call before rejecting" "" "$(cat "$CLI_LOG")"

: > "$CLI_LOG"
run_and_capture categorize "$CLI_MSG_ID" -add "Follow up"
eq "categorize -add (single dash) is rejected" "1" "$CLI_RC"
eq "categorize -add makes no API call before rejecting" "" "$(cat "$CLI_LOG")"

: > "$CLI_LOG"
run_and_capture categorize "$CLI_MSG_ID" --append "Follow up"
eq "categorize --append (typo verb) is rejected" "1" "$CLI_RC"
eq "categorize --append makes no API call before rejecting" "" "$(cat "$CLI_LOG")"

: > "$CLI_LOG"
run_and_capture categorize "$CLI_MSG_ID" --add "Follow up" "Extra"
eq "categorize --add with a stray extra argument is rejected" "1" "$CLI_RC"
eq "categorize --add stray-argument case makes no API call before rejecting" "" "$(cat "$CLI_LOG")"

# The legitimate replace and clear forms must still work unchanged.
printf '%s' '{"categories":["Old"]}' > "$CLI_CURRENTCATS"
: > "$CLI_LOG"
run_and_capture categorize "$CLI_MSG_ID" "A,B"
eq "categorize plain replace still works" "0" "$CLI_RC"
eq "categorize plain replace sets the given categories" "Categories set: A, B" "$CLI_OUT"

: > "$CLI_LOG"
run_and_capture categorize "$CLI_MSG_ID" ""
eq "categorize empty string still clears" "0" "$CLI_RC"
eq "categorize empty string clears output" "Categories cleared" "$CLI_OUT"

########################################
# An API error reading the master category list must never look like "no such
# category" (mkcategory/rccategory/rmcategory), and must never be reported as a
# false "not in the master list" (categorize --add). A transient read failure
# that reads as "absent" would have mkcategory create a duplicate. Exercised at
# all four call sites under a mocked failure of the masterCategories GET.
########################################
printf '%s' '{"error":{"code":"NetworkError","message":"boom"}}' > "$CLI_MASTERCATS"

: > "$CLI_LOG"
run_and_capture mkcategory "Follow up" red
eq "mkcategory on unreadable master list fails closed (non-zero)" "1" "$CLI_RC"
if contains "$CLI_ERR" "could not read the master category list"; then
    eq "mkcategory reports the read failure clearly" ok ok
else
    eq "mkcategory reports the read failure clearly" "mentions could not read the master category list" "$CLI_ERR"
fi
eq "mkcategory does not claim success on read failure" "0" \
    "$(contains "$CLI_OUT$CLI_ERR" "Category created" && echo 1 || echo 0)"
eq "mkcategory issues no create POST when the list is unreadable" "0" \
    "$(grep -c 'POST .*masterCategories$' "$CLI_LOG")"

: > "$CLI_LOG"
run_and_capture rccategory "Follow up" red
eq "rccategory on unreadable master list fails closed (non-zero)" "1" "$CLI_RC"
eq "rccategory does NOT misreport an unreadable list as 'no category named'" "0" \
    "$(contains "$CLI_ERR" "no category named" && echo 1 || echo 0)"
if contains "$CLI_ERR" "could not read the master category list"; then
    eq "rccategory reports the read failure clearly" ok ok
else
    eq "rccategory reports the read failure clearly" "mentions could not read the master category list" "$CLI_ERR"
fi
eq "rccategory issues no PATCH when the list is unreadable" "0" \
    "$(grep -c 'PATCH .*masterCategories/' "$CLI_LOG")"

: > "$CLI_LOG"
run_and_capture rmcategory "Follow up"
eq "rmcategory on unreadable master list fails closed (non-zero)" "1" "$CLI_RC"
eq "rmcategory does NOT misreport an unreadable list as 'no category named'" "0" \
    "$(contains "$CLI_ERR" "no category named" && echo 1 || echo 0)"
if contains "$CLI_ERR" "could not read the master category list"; then
    eq "rmcategory reports the read failure clearly" ok ok
else
    eq "rmcategory reports the read failure clearly" "mentions could not read the master category list" "$CLI_ERR"
fi
eq "rmcategory issues no DELETE when the list is unreadable" "0" \
    "$(grep -c 'DELETE .*masterCategories/' "$CLI_LOG")"

# categorize --add: the master-list check is advisory only, so a failed check
# must NOT block the message-level write, and must NOT be misreported as "not
# in the master list" (we simply don't know).
printf '%s' '{"categories":["Old"]}' > "$CLI_CURRENTCATS"
: > "$CLI_LOG"
run_and_capture categorize "$CLI_MSG_ID" --add "Follow up"
eq "categorize --add still succeeds when the master-list check fails" "0" "$CLI_RC"
eq "categorize --add still applies the category despite the check failing" \
    "Categories set: Old, Follow up" "$CLI_OUT"
eq "categorize --add does NOT falsely claim 'not in the master list'" "0" \
    "$(contains "$CLI_ERR" "is not in the master list" && echo 1 || echo 0)"
if contains "$CLI_ERR" "could not check"; then
    eq "categorize --add surfaces the check failure as a warning" ok ok
else
    eq "categorize --add surfaces the check failure as a warning" "mentions could not check" "$CLI_ERR"
fi
eq "categorize --add still PATCHes the message despite the check failing" "1" \
    "$(grep -c "PATCH .*/me/messages/$CLI_MSG_ID\$" "$CLI_LOG")"

########################################
# mkcategory/rccategory must report what Graph actually returned,
# not the value the caller asked for - Graph can answer 200 with the colour
# unchanged, and a report that echoes the request is a false success.
########################################
printf '%s' '{"value":[]}' > "$CLI_MASTERCATS"
printf '%s' '{"id":"NEWCAT","displayName":"Follow Up","color":"preset3"}' > "$CLI_MASTERCATS_POST"
run_and_capture mkcategory "Follow up" red
eq "mkcategory reports the server's displayName/color, not the request" \
    "Category created: Follow Up (preset3)" "$CLI_OUT"

printf '%s' '{"value":[{"id":"CAT1","displayName":"Follow up","color":"preset0"}]}' > "$CLI_MASTERCATS"
printf '%s' '{"id":"CAT1","displayName":"Follow up","color":"preset0"}' > "$CLI_MASTERCATS_PATCH"
run_and_capture rccategory "Follow up" darkblue
eq "rccategory reports the server's actual colour, not the requested one" \
    "Category recoloured: Follow up (preset0)" "$CLI_OUT"

unset -f curl
rm -rf "$CLI_TMP"

rm -f "$body_file" /tmp/outlook_test_last_url /tmp/outlook_test_calls
echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
