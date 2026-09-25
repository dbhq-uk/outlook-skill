#!/bin/bash
# Offline tests for the shared token code in scripts/lib/graph.sh.
#
# The real functions are sourced from the library, and `curl` is a fake
# executable placed first on PATH, so no account or network is needed. Every
# test runs against a throwaway HOME: nothing here touches ~/.dbhq.
#
#   bash skills/outlook/tests/token_test.sh
#
# What this pins: a failed refresh never changes credentials.json (the bug that
# emptied it and lost the refresh token), a good refresh keeps the old refresh
# token when Microsoft sends no new one, the 60-second expiry margin, and the
# lock that stops two commands refreshing the same account at once.
#
# The fake curl reads FAKE_CURL_MODE and FAKE_CURL_LOG from the environment
# (SC2030/SC2031 are about the per-test subshell exports, which is intended).
# shellcheck disable=SC2030,SC2031
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPTS="$(cd "$TESTS_DIR/../scripts" && pwd)"
LIB="$SCRIPTS/lib/graph.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
contains() { case "$1" in *"$2"*) return 0 ;; *) return 1 ;; esac; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# --- Fake curl ---------------------------------------------------------------
# Token endpoint: answers per FAKE_CURL_MODE, the way real curl run with
# --fail-with-body would (an HTTP error prints the body and exits 22).
# Graph: answers {"value":[]}.
mkdir -p "$TMP/bin"
cat > "$TMP/bin/curl" <<'FAKE'
#!/bin/bash
url=""
for a in "$@"; do case "$a" in https://*) url="$a" ;; esac; done
{ printf -- '--- call\n'; printf '%s\n' "$@"; } >> "${FAKE_CURL_LOG:-/dev/null}"
case "$url" in
  https://login.microsoftonline.com/*)
    case "${FAKE_CURL_MODE:-ok}" in
      exit6)        echo "curl: (6) Could not resolve host: login.microsoftonline.com" >&2; exit 6 ;;
      timeout)      echo "curl: (28) Operation timed out" >&2; exit 28 ;;
      empty200)     exit 0 ;;
      html502)      printf '<html><body>502 Bad Gateway</body></html>'; exit 22 ;;
      html200)      printf '<html><body>Sign in to the Wi-Fi</body></html>'; exit 0 ;;
      invalid_grant) printf '{"error":"invalid_grant","error_description":"AADSTS700082: The refresh token has expired."}'; exit 22 ;;
      noaccess)     printf '{"token_type":"Bearer","expires_in":3600}'; exit 0 ;;
      ok)           printf '{"token_type":"Bearer","access_token":"new-access","refresh_token":"new-refresh","expires_in":3600}'; exit 0 ;;
      ok_norefresh) printf '{"token_type":"Bearer","access_token":"new-access","expires_in":"3599"}'; exit 0 ;;
    esac
    ;;
  https://graph.microsoft.com/*/mailFolders/inbox) printf '{"totalItemCount":3,"unreadItemCount":1}'; exit 0 ;;
  https://graph.microsoft.com/*) printf '{"value":[]}'; exit 0 ;;
esac
exit 0
FAKE
chmod +x "$TMP/bin/curl"
export PATH="$TMP/bin:$PATH"

# --- Fixture HOME --------------------------------------------------------------
FIX_HOME="$TMP/home"
CONFIG_DIR="$FIX_HOME/.dbhq/outlook/default"
CONFIG_FILE="$CONFIG_DIR/config.json"
CREDS_FILE="$CONFIG_DIR/credentials.json"
# shellcheck disable=SC2034 # read by lib/graph.sh
ACCOUNT=default
mkdir -p "$CONFIG_DIR"
chmod 700 "$FIX_HOME/.dbhq" "$FIX_HOME/.dbhq/outlook" "$CONFIG_DIR"
printf '%s' '{"client_id":"test-client","client_secret":"test-secret"}' > "$CONFIG_FILE"

# Write credentials with a given expires_at (omit the argument for none).
write_creds() {
    if [ $# -gt 0 ]; then
        printf '{"access_token":"old-access","refresh_token":"old-refresh","expires_at":%s}' "$1" > "$CREDS_FILE"
    else
        printf '%s' '{"access_token":"old-access","refresh_token":"old-refresh"}' > "$CREDS_FILE"
    fi
    chmod 600 "$CREDS_FILE"
    cp "$CREDS_FILE" "$TMP/creds.before"
}
unchanged() { cmp -s "$CREDS_FILE" "$TMP/creds.before" && echo unchanged || echo changed; }
calls() { grep -c -- '^--- call$' "$FAKE_CURL_LOG" 2>/dev/null || true; }

export FAKE_CURL_LOG="$TMP/curl.log"

# Source the real library into this shell. `date` is a function so the expiry
# tests can fix the clock; with FAKE_NOW unset it is the real date.
# shellcheck source=../scripts/lib/graph.sh
. "$LIB"
date() { if [ -n "${FAKE_NOW:-}" ] && [ "${1:-}" = "+%s" ]; then echo "$FAKE_NOW"; else command date "$@"; fi; }

########################################
# A failed refresh leaves credentials.json byte-for-byte unchanged and returns
# non-zero. Before the fix, each of these emptied the file.
########################################
for mode in exit6 timeout empty200 html502 html200 invalid_grant noaccess; do
    write_creds 0
    : > "$FAKE_CURL_LOG"
    out=$(FAKE_CURL_MODE=$mode refresh_access_token 2>"$TMP/err"); rc=$?
    err=$(cat "$TMP/err")
    eq "refresh [$mode] returns non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
    eq "refresh [$mode] leaves credentials.json unchanged" "unchanged" "$(unchanged)"
    eq "refresh [$mode] prints no token" "" "$out"
    if contains "$err" "credentials.json is unchanged"; then
        eq "refresh [$mode] says the file is unchanged" ok ok
    else
        eq "refresh [$mode] says the file is unchanged" "mentions 'credentials.json is unchanged'" "$err"
    fi
    eq "refresh [$mode] leaves no temp file behind" "0" \
       "$(find "$CONFIG_DIR" -name '.credentials.*' | wc -l | tr -d ' ')"
done

write_creds 0
FAKE_CURL_MODE=invalid_grant refresh_access_token >/dev/null 2>"$TMP/err"
if contains "$(cat "$TMP/err")" "AADSTS700082" && contains "$(cat "$TMP/err")" "outlook-setup.sh --account default"; then
    eq "invalid_grant shows Microsoft's reason and how to sign in again" ok ok
else
    eq "invalid_grant shows Microsoft's reason and how to sign in again" "AADSTS700082 + outlook-setup.sh" "$(cat "$TMP/err")"
fi

write_creds 0
FAKE_CURL_MODE=exit6 refresh_access_token >/dev/null 2>"$TMP/err"
if contains "$(cat "$TMP/err")" "curl exit 6"; then
    eq "network failure names the curl exit code" ok ok
else
    eq "network failure names the curl exit code" "mentions curl exit 6" "$(cat "$TMP/err")"
fi

# An empty or unreadable credentials file (what the old bug left behind) gets a
# clear "sign in again", and no request is made with an empty refresh token.
: > "$CREDS_FILE"
: > "$FAKE_CURL_LOG"
FAKE_CURL_MODE=ok refresh_access_token >/dev/null 2>"$TMP/err"; rc=$?
eq "empty credentials.json -> non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "empty credentials.json -> no request" "0" "$(calls)"
if contains "$(cat "$TMP/err")" "no refresh token"; then
    eq "empty credentials.json says there is no refresh token" ok ok
else
    eq "empty credentials.json says there is no refresh token" "mentions 'no refresh token'" "$(cat "$TMP/err")"
fi

########################################
# A good refresh writes the new tokens, stamps expires_at, stays at 600.
########################################
write_creds 0
: > "$FAKE_CURL_LOG"
before=$(command date +%s)
out=$(FAKE_CURL_MODE=ok refresh_access_token 2>"$TMP/err"); rc=$?
after=$(command date +%s)
eq "refresh ok returns 0" "0" "$rc"
eq "refresh ok prints the new access token" "new-access" "$out"
eq "refresh ok stores the new access token" "new-access" "$(jq -r .access_token "$CREDS_FILE")"
eq "refresh ok stores the new refresh token" "new-refresh" "$(jq -r .refresh_token "$CREDS_FILE")"
at=$(jq -r .expires_at "$CREDS_FILE")
eq "refresh ok stamps expires_at = now + expires_in" "1" \
   "$([ "$at" -ge $((before + 3600)) ] && [ "$at" -le $((after + 3600)) ] && echo 1 || echo 0)"
eq "refresh ok keeps the file at 600" "600" "$(stat -c %a "$CREDS_FILE" 2>/dev/null || stat -f %Lp "$CREDS_FILE")"
eq "refresh ok leaves no temp file behind" "0" \
   "$(find "$CONFIG_DIR" -name '.credentials.*' | wc -l | tr -d ' ')"
eq "refresh sends the stored refresh token" "1" "$(grep -c '^refresh_token=old-refresh$' "$FAKE_CURL_LOG")"
eq "refresh sends grant_type=refresh_token" "1" "$(grep -c '^grant_type=refresh_token$' "$FAKE_CURL_LOG")"
eq "refresh runs curl with --fail-with-body" "1" "$(grep -c '^--fail-with-body$' "$FAKE_CURL_LOG")"
eq "refresh runs curl with a --max-time" "1" "$(grep -c '^--max-time$' "$FAKE_CURL_LOG")"
eq "refresh runs curl with a --connect-timeout" "1" "$(grep -c '^--connect-timeout$' "$FAKE_CURL_LOG")"

# The fixture config is an install from before PKCE: it has a client secret,
# and must keep sending it, or that install stops working on upgrade.
eq "an old config with a secret still sends it" "1" "$(grep -c '^client_secret=test-secret$' "$FAKE_CURL_LOG")"

# A public client (every setup since PKCE) has no secret, and the refresh must
# send no client_secret parameter at all, not even an empty one: Microsoft
# refuses a secret from a public client.
cp "$CONFIG_FILE" "$TMP/config.before"
printf '%s' '{"client_id":"public-client","tenant":"common"}' > "$CONFIG_FILE"
write_creds 0
: > "$FAKE_CURL_LOG"
out=$(FAKE_CURL_MODE=ok refresh_access_token 2>"$TMP/err"); rc=$?
eq "public client: refresh ok" "0" "$rc"
eq "public client: refresh stores the new token" "new-access" "$(jq -r .access_token "$CREDS_FILE")"
eq "public client: the request carries no client_secret" "0" "$(grep -c '^client_secret' "$FAKE_CURL_LOG" || true)"
eq "public client: the request carries the client_id" "1" "$(grep -c '^client_id=public-client$' "$FAKE_CURL_LOG")"
cp "$TMP/config.before" "$CONFIG_FILE"

# Microsoft does not always send a new refresh token. The old one must survive,
# or the next refresh has nothing to send.
write_creds 0
FAKE_CURL_MODE=ok_norefresh refresh_access_token >/dev/null 2>&1
eq "refresh without a new refresh_token keeps the old one" "old-refresh" "$(jq -r .refresh_token "$CREDS_FILE")"
eq "refresh without a new refresh_token still stores the access token" "new-access" "$(jq -r .access_token "$CREDS_FILE")"
eq "a string expires_in is still honoured" "1" \
   "$(at=$(jq -r .expires_at "$CREDS_FILE"); now=$(command date +%s); [ "$at" -gt $((now + 3500)) ] && echo 1 || echo 0)"

########################################
# ensure_valid_token: the real function, with a fixed clock. The token is used
# until it is within 60 seconds of expiry. These fail if the margin or the
# comparison changes.
########################################
export FAKE_NOW=1000000
write_creds $((FAKE_NOW + 61)); : > "$FAKE_CURL_LOG"
eq "61s left -> stored token"          "old-access" "$(FAKE_CURL_MODE=ok ensure_valid_token)"
eq "61s left -> no request"            "0" "$(calls)"
write_creds $((FAKE_NOW + 60)); : > "$FAKE_CURL_LOG"
eq "60s left -> refreshed token"       "new-access" "$(FAKE_CURL_MODE=ok ensure_valid_token)"
eq "60s left -> one request"           "1" "$(calls)"
write_creds $((FAKE_NOW - 3600)); : > "$FAKE_CURL_LOG"
eq "expired -> refreshed token"        "new-access" "$(FAKE_CURL_MODE=ok ensure_valid_token)"
write_creds; : > "$FAKE_CURL_LOG"
eq "no expires_at -> refreshed token"  "new-access" "$(FAKE_CURL_MODE=ok ensure_valid_token)"
printf '{"refresh_token":"old-refresh","expires_at":%s}' $((FAKE_NOW + 3600)) > "$CREDS_FILE"; : > "$FAKE_CURL_LOG"
eq "no access_token -> refreshed token" "new-access" "$(FAKE_CURL_MODE=ok ensure_valid_token)"
write_creds $((FAKE_NOW - 1))
FAKE_CURL_MODE=exit6 ensure_valid_token >/dev/null 2>&1; rc=$?
eq "expired + failed refresh -> non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "expired + failed refresh -> file unchanged" "unchanged" "$(unchanged)"
unset FAKE_NOW

########################################
# The lock: a command that finds the token expired while another is refreshing
# waits, then uses the token the other one wrote instead of refreshing again.
########################################
if command -v flock >/dev/null 2>&1; then
    write_creds 0
    : > "$FAKE_CURL_LOG"
    far=$(( $(command date +%s) + 3600 ))
    (
        flock 9
        sleep 1
        printf '{"access_token":"other-access","refresh_token":"other-refresh","expires_at":%s}' "$far" > "$CREDS_FILE"
    ) 9>> "$CONFIG_DIR/.token.lock" &
    holder=$!
    sleep 0.3
    got=$(FAKE_CURL_MODE=ok ensure_valid_token 2>/dev/null)
    wait "$holder"
    eq "waits for the lock, then uses the token the holder wrote" "other-access" "$got"
    eq "no second refresh after waiting for the lock" "0" "$(calls)"
else
    echo "skip - lock test (flock not installed)"
fi

########################################
# The entry scripts, run as a user would, against the fixture HOME.
########################################
run_cli() { HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default OUTLOOK_TZ=Europe/London "$@"; }

write_creds 0
out=$(FAKE_CURL_MODE=empty200 run_cli bash "$SCRIPTS/outlook-token.sh" refresh 2>&1); rc=$?
eq "token.sh refresh on an empty 200 exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "token.sh refresh on an empty 200 does not claim success" "0" \
   "$(contains "$out" "refreshed successfully" && echo 1 || echo 0)"
eq "token.sh refresh on an empty 200 leaves the file unchanged" "unchanged" "$(unchanged)"

write_creds 0
out=$(FAKE_CURL_MODE=ok run_cli bash "$SCRIPTS/outlook-token.sh" refresh 2>&1); rc=$?
eq "token.sh refresh ok exits 0" "0" "$rc"
eq "token.sh refresh ok stores the new token" "new-access" "$(jq -r .access_token "$CREDS_FILE")"

write_creds 0
FAKE_CURL_MODE=exit6 run_cli bash "$SCRIPTS/outlook-mail.sh" inbox 1 >/dev/null 2>&1; rc=$?
eq "mail.sh with a failed refresh exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "mail.sh with a failed refresh leaves the file unchanged" "unchanged" "$(unchanged)"

write_creds 0
FAKE_CURL_MODE=html502 run_cli bash "$SCRIPTS/outlook-calendar.sh" today >/dev/null 2>&1; rc=$?
eq "calendar.sh with a failed refresh exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "calendar.sh with a failed refresh leaves the file unchanged" "unchanged" "$(unchanged)"

# `get` refreshes a stale token before printing it. It used to print whatever
# was stored, so a hand-written Graph call made with it failed with 401.
write_creds 0; : > "$FAKE_CURL_LOG"
out=$(FAKE_CURL_MODE=ok run_cli bash "$SCRIPTS/outlook-token.sh" get 2>/dev/null); rc=$?
eq "token.sh get with an expired token exits 0" "0" "$rc"
eq "token.sh get with an expired token prints a refreshed token" "new-access" "$out"
eq "token.sh get with an expired token makes one refresh request" "1" "$(calls)"
eq "token.sh get with an expired token stores the refreshed token" "new-access" "$(jq -r .access_token "$CREDS_FILE")"

write_creds $(( $(command date +%s) + 30 )); : > "$FAKE_CURL_LOG"
eq "token.sh get within the 60s margin refreshes" "new-access" \
   "$(FAKE_CURL_MODE=ok run_cli bash "$SCRIPTS/outlook-token.sh" get 2>/dev/null)"

write_creds $(( $(command date +%s) + 3600 )); : > "$FAKE_CURL_LOG"
eq "token.sh get with a fresh token prints it" "old-access" \
   "$(FAKE_CURL_MODE=ok run_cli bash "$SCRIPTS/outlook-token.sh" get 2>/dev/null)"
eq "token.sh get with a fresh token makes no request" "0" "$(calls)"

write_creds 0
out=$(FAKE_CURL_MODE=exit6 run_cli bash "$SCRIPTS/outlook-token.sh" get 2>/dev/null); rc=$?
eq "token.sh get with a failed refresh exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "token.sh get with a failed refresh prints no token" "" "$out"

write_creds 0; : > "$FAKE_CURL_LOG"
out=$(FAKE_CURL_MODE=ok run_cli bash "$SCRIPTS/outlook-token.sh" test 2>&1); rc=$?
eq "token.sh test refreshes an expired token instead of failing" "0" "$rc"
eq "token.sh test calls Graph with the refreshed token" "1" \
   "$(grep -c '^Authorization: Bearer new-access$' "$FAKE_CURL_LOG")"

# The library is found through a symlink to the script itself, not only through
# a symlinked skill directory.
far=$(( $(command date +%s) + 3600 ))
write_creds "$far"
ln -s "$SCRIPTS/outlook-token.sh" "$TMP/bin/linked-token"
eq "a symlinked script still finds lib/graph.sh" "old-access" \
   "$(run_cli bash "$TMP/bin/linked-token" get 2>&1)"

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
