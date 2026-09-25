#!/bin/bash
# Offline tests for outlook-setup.sh: a public client, PKCE, and no client
# secret, plus the path that keeps an install from before PKCE working.
#
# Setup runs for real against fakes first on PATH: `az` logs what it was asked
# and answers from fixtures, `curl` answers the token endpoint and Graph, and
# `xdg-open` plays the browser, taking the sign-in URL and handing back a
# redirect URL for the pasted-URL prompt. Every run uses a throwaway HOME.
# Nothing here calls Azure or Microsoft, and nothing touches ~/.dbhq.
#
#   bash skills/outlook/tests/setup_test.sh
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPTS="$(cd "$TESTS_DIR/../scripts" && pwd)"
SETUP="$SCRIPTS/outlook-setup.sh"
NATIVE="https://login.microsoftonline.com/common/oauth2/nativeclient"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }
hasnt() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" "no: $2" "$3"; else eq "$1" ok ok; fi; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT
mkdir -p "$TMP/bin"

# --- Fake az -------------------------------------------------------------------
# One line per call in $FAKE_AZ_LOG. `az rest` writes its --body to
# $FAKE_AZ_BODY and exits $FAKE_AZ_REST_RC. `ad app list` answers
# $FAKE_AZ_EXISTING (empty: no app of that name).
cat > "$TMP/bin/az" <<'FAKE'
#!/bin/bash
printf 'az %s\n' "$*" >> "$FAKE_AZ_LOG"
case "$1 $2 $3" in
  "account show "*) exit 0 ;;
  "ad app list")    printf '%s\n' "${FAKE_AZ_EXISTING:-}" ;;
  "ad app create")  printf 'new-app-id\n' ;;
  "ad app show")
      case "$*" in
        *web.redirectUris*)          printf '["%s","https://other.example/callback"]\n' "https://login.microsoftonline.com/common/oauth2/nativeclient" ;;
        *publicClient.redirectUris*) printf 'null\n' ;;
      esac ;;
  "ad app credential") printf 'A-NEW-SECRET\n' ;;
  "rest "*|"rest --method PATCH")
      prev=""; for a in "$@"; do [ "$prev" = "--body" ] && printf '%s' "$a" > "$FAKE_AZ_BODY"; prev="$a"; done
      exit "${FAKE_AZ_REST_RC:-0}" ;;
esac
exit 0
FAKE

# --- Fake curl -----------------------------------------------------------------
# Each call is logged as "--- call" and then one argument per line. The token
# endpoint answers with tokens; Graph's inbox answers with counts.
cat > "$TMP/bin/curl" <<'FAKE'
#!/bin/bash
url=""
for a in "$@"; do case "$a" in https://*) url="$a" ;; esac; done
{ printf -- '--- call\n'; printf '%s\n' "$@"; } >> "$FAKE_CURL_LOG"
case "$url" in
  https://login.microsoftonline.com/*/token)
      printf '{"token_type":"Bearer","access_token":"access-1","refresh_token":"refresh-1","expires_in":3600}' ;;
  https://graph.microsoft.com/*/mailFolders/inbox)
      printf '{"totalItemCount":7,"unreadItemCount":2}' ;;
  *) printf '{}' ;;
esac
exit 0
FAKE

# --- Fake browser --------------------------------------------------------------
# Keeps the sign-in URL, and writes the URL the user would paste back: the
# native-client page with a code and the same state (or $FAKE_STATE instead).
cat > "$TMP/bin/xdg-open" <<'FAKE'
#!/bin/bash
printf '%s' "$1" > "$FAKE_AUTH_URL"
state=$(printf '%s' "$1" | tr '?&' '\n\n' | sed -n 's/^state=//p')
printf '%s?code=THE.CODE-123&state=%s&session_state=abc\n' \
    "https://login.microsoftonline.com/common/oauth2/nativeclient" "${FAKE_STATE:-$state}" > "$FAKE_REDIRECT"
FAKE
chmod +x "$TMP/bin/az" "$TMP/bin/curl" "$TMP/bin/xdg-open"

export FAKE_AZ_LOG="$TMP/az.log" FAKE_AZ_BODY="$TMP/az-body.json" FAKE_CURL_LOG="$TMP/curl.log" \
       FAKE_AUTH_URL="$TMP/auth-url" FAKE_REDIRECT="$TMP/redirect"

# run_setup <home> <account> [answer...]: runs setup with the answers typed in
# order, then pastes the redirect URL once the fake browser has written it.
# Sets $out and $rc.
run_setup() {
    local home="$1" account="$2"; shift 2
    rm -f "$FAKE_AUTH_URL" "$FAKE_REDIRECT" "$FAKE_AZ_BODY"
    : > "$FAKE_AZ_LOG"; : > "$FAKE_CURL_LOG"
    out=$( { for a in "$@"; do printf '%s\n' "$a"; done
             for _ in $(seq 1 50); do [ -s "$FAKE_REDIRECT" ] && break; sleep 0.1; done
             cat "$FAKE_REDIRECT" 2>/dev/null; } \
           | HOME="$home" PATH="$TMP/bin:$PATH" bash "$SETUP" --account "$account" 2>&1 ); rc=$?
}
# The value of one form field in the token request, or "(absent)".
token_field() {
    awk -v k="$1=" 'index($0, k) == 1 { print substr($0, length(k) + 1); f = 1 } END { if (!f) print "(absent)" }' "$FAKE_CURL_LOG"
}
url_param() { tr '?&' '\n\n' < "$FAKE_AUTH_URL" | sed -n "s/^$1=//p"; }
legacy_config() {  # legacy_config <home> <account> <client_id>: an install from before PKCE
    mkdir -p "$1/.dbhq/outlook/$2"
    chmod 700 "$1/.dbhq" "$1/.dbhq/outlook" "$1/.dbhq/outlook/$2"
    printf '{"client_id":"%s","client_secret":"old-secret","tenant":"common"}' "$3" > "$1/.dbhq/outlook/$2/config.json"
    printf '{"access_token":"a","refresh_token":"r","expires_at":%s}' "$(( $(date +%s) + 3600 ))" > "$1/.dbhq/outlook/$2/credentials.json"
    chmod 600 "$1/.dbhq/outlook/$2/config.json" "$1/.dbhq/outlook/$2/credentials.json"
}

########################################
# A fresh setup: a new public-client app, PKCE, and no secret anywhere.
########################################
H="$TMP/home-fresh"; mkdir -p "$H"
run_setup "$H" default
CFG="$H/.dbhq/outlook/default/config.json"
eq "fresh: setup exits 0" "0" "$rc"
has "fresh: the connection test passes" "Connection successful!" "$out"
create=$(grep '^az ad app create' "$FAKE_AZ_LOG")
has "fresh: the app is registered with a public-client redirect URI" "--public-client-redirect-uris $NATIVE" "$create"
has "fresh: public client flows are allowed" "--is-fallback-public-client true" "$create"
hasnt "fresh: the app has no web redirect URI" "--web-redirect-uris" "$create"
eq "fresh: no client secret is created" "0" "$(grep -c 'credential reset' "$FAKE_AZ_LOG" || true)"
eq "fresh: the sign-in uses PKCE with S256" "S256" "$(url_param code_challenge_method)"
verifier=$(token_field code_verifier)
eq "fresh: the verifier is 64 characters" "64" "${#verifier}"
eq "fresh: the challenge is the SHA-256 of the verifier sent" \
   "$(printf '%s' "$verifier" | openssl dgst -sha256 -binary | openssl base64 -A | tr '+/' '-_' | tr -d '=')" \
   "$(url_param code_challenge)"
eq "fresh: the pasted code is redeemed" "THE.CODE-123" "$(token_field code)"
eq "fresh: the token request sends no client_secret" "(absent)" "$(token_field client_secret)"
eq "fresh: the token request is an authorisation-code grant" "authorization_code" "$(token_field grant_type)"
eq "fresh: config.json has no client_secret" "false" "$(jq 'has("client_secret")' "$CFG")"
eq "fresh: config.json has the new app's ID" "new-app-id" "$(jq -r .client_id "$CFG")"
eq "fresh: config.json is 600" "600" "$(stat -c %a "$CFG" 2>/dev/null || stat -f %Lp "$CFG")"
eq "fresh: credentials.json holds the tokens" "refresh-1" "$(jq -r .refresh_token "$H/.dbhq/outlook/default/credentials.json")"
eq "fresh: expires_at is stamped" "number" "$(jq -r '.expires_at | type' "$H/.dbhq/outlook/default/credentials.json")"

# And a refresh then works, with no secret to send.
jq '.expires_at = 0' "$H/.dbhq/outlook/default/credentials.json" > "$TMP/c" && cp "$TMP/c" "$H/.dbhq/outlook/default/credentials.json"
: > "$FAKE_CURL_LOG"
out=$(HOME="$H" PATH="$TMP/bin:$PATH" bash "$SCRIPTS/outlook-token.sh" refresh 2>&1); rc=$?
eq "fresh: a refresh afterwards exits 0" "0" "$rc"
eq "fresh: the refresh sends no client_secret" "(absent)" "$(token_field client_secret)"
eq "fresh: the refresh sends the stored refresh token" "refresh-1" "$(token_field refresh_token)"

########################################
# A second account reuses the public app: no Azure calls and no secret.
########################################
run_setup "$H" work Y
eq "second account: setup exits 0" "0" "$rc"
eq "second account: nothing is asked of Azure" "" "$(cat "$FAKE_AZ_LOG")"
eq "second account: it reuses the app" "new-app-id" "$(jq -r .client_id "$H/.dbhq/outlook/work/config.json")"
eq "second account: no client_secret" "false" "$(jq 'has("client_secret")' "$H/.dbhq/outlook/work/config.json")"
eq "second account: the token request sends no client_secret" "(absent)" "$(token_field client_secret)"

########################################
# The state must match: a URL from another sign-in is refused before any
# token request, and nothing is written.
########################################
H="$TMP/home-state"; mkdir -p "$H"
FAKE_STATE=not-this-one run_setup "$H" default
eq "wrong state: setup exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "wrong state: says why" "different sign-in" "$out"
eq "wrong state: no token request" "0" "$(grep -c -- '--- call' "$FAKE_CURL_LOG" || true)"
eq "wrong state: no config written" "no" "$([ -e "$H/.dbhq/outlook/default/config.json" ] && echo yes || echo no)"

########################################
# An install from before PKCE, kept on its secret by choice: it still works.
########################################
H="$TMP/home-keep"; mkdir -p "$H"
legacy_config "$H" default legacy-app-id
cp "$H/.dbhq/outlook/default/config.json" "$TMP/default.before"
run_setup "$H" work Y n
eq "keep secret: setup exits 0" "0" "$rc"
eq "keep secret: nothing is asked of Azure" "" "$(cat "$FAKE_AZ_LOG")"
eq "keep secret: the token request sends the old secret" "old-secret" "$(token_field client_secret)"
eq "keep secret: and uses PKCE too" "64" "$(token_field code_verifier | tr -d '\n' | wc -c | tr -d ' ')"
eq "keep secret: the new config keeps the secret" "old-secret" "$(jq -r .client_secret "$H/.dbhq/outlook/work/config.json")"
has "keep secret: warns it will expire" "stops working when it expires" "$out"
eq "keep secret: the other account's config is untouched" "same" \
   "$(cmp -s "$TMP/default.before" "$H/.dbhq/outlook/default/config.json" && echo same || echo changed)"

########################################
# The migration: reuse an old app and convert it to a public client.
########################################
H="$TMP/home-convert"; mkdir -p "$H"
legacy_config "$H" default legacy-app-id
legacy_config "$H" work legacy-app-id
run_setup "$H" work y Y Y
eq "convert: setup exits 0" "0" "$rc"
eq "convert: the app is patched through Graph" "1" \
   "$(grep -c "^az rest --method PATCH --uri https://graph.microsoft.com/v1.0/applications(appId='legacy-app-id')" "$FAKE_AZ_LOG" || true)"
eq "convert: public client flows are allowed" "true" "$(jq -r .isFallbackPublicClient "$FAKE_AZ_BODY")"
eq "convert: the native-client URI moves to the public client platform" "[\"$NATIVE\"]" \
   "$(jq -c .publicClient.redirectUris "$FAKE_AZ_BODY")"
eq "convert: it leaves the web platform, and other web URIs stay" '["https://other.example/callback"]' \
   "$(jq -c .web.redirectUris "$FAKE_AZ_BODY")"
eq "convert: no new secret is created" "0" "$(grep -c 'credential reset' "$FAKE_AZ_LOG" || true)"
eq "convert: the token request sends no client_secret" "(absent)" "$(token_field client_secret)"
eq "convert: the account's config drops its secret" "false" "$(jq 'has("client_secret")' "$H/.dbhq/outlook/work/config.json")"
eq "convert: and keeps the same app" "legacy-app-id" "$(jq -r .client_id "$H/.dbhq/outlook/work/config.json")"
has "convert: names the other account still on the old secret" "Account 'default' still uses this app's old secret" "$out"
eq "convert: the other account still has its secret, so it keeps working" "old-secret" \
   "$(jq -r .client_secret "$H/.dbhq/outlook/default/config.json")"

########################################
# Re-running setup on an account whose app is found by name converts it too.
########################################
H="$TMP/home-byname"; mkdir -p "$H"
legacy_config "$H" default legacy-app-id
FAKE_AZ_EXISTING=legacy-app-id run_setup "$H" default y Y
eq "by name: setup exits 0" "0" "$rc"
eq "by name: the found app is converted" "1" "$(grep -c '^az rest --method PATCH' "$FAKE_AZ_LOG" || true)"
eq "by name: no new app is created" "0" "$(grep -c '^az ad app create' "$FAKE_AZ_LOG" || true)"
eq "by name: config.json drops the secret" "false" "$(jq 'has("client_secret")' "$H/.dbhq/outlook/default/config.json")"

########################################
# If the app cannot be converted, setup stops, says what to do in the portal,
# and leaves the account as it was.
########################################
H="$TMP/home-denied"; mkdir -p "$H"
legacy_config "$H" default legacy-app-id
cp "$H/.dbhq/outlook/default/config.json" "$TMP/denied.before"
FAKE_AZ_EXISTING=legacy-app-id FAKE_AZ_REST_RC=1 run_setup "$H" default y Y
eq "denied: setup exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "denied: gives the portal steps" "Allow public client flows" "$out"
eq "denied: no sign-in is attempted" "0" "$(grep -c -- '--- call' "$FAKE_CURL_LOG" || true)"
eq "denied: the account's config is unchanged" "same" \
   "$(cmp -s "$TMP/denied.before" "$H/.dbhq/outlook/default/config.json" && echo same || echo changed)"

echo
echo "PASS=$PASS FAIL=$FAIL"
[ "$FAIL" -eq 0 ]
