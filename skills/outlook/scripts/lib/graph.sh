# shellcheck shell=bash
# Shared token handling, the Graph request wrapper (timeouts and throttling)
# and the read-only gate for the outlook scripts. Sourced, never run.
#
# outlook-mail.sh, outlook-calendar.sh and outlook-token.sh each carried their
# own copy of this code, and a bug in it (a failed refresh wrote an empty
# response over credentials.json and lost the refresh token) existed three
# times as a result. There is one copy now, here.
#
# The caller sets these before calling anything below:
#   ACCOUNT      account name, used in messages
#   CONFIG_DIR   ~/.dbhq/outlook/<account>
#   CONFIG_FILE  $CONFIG_DIR/config.json    (client_id; client_secret only on
#                an app registered before setup moved to PKCE)
#   CREDS_FILE   $CONFIG_DIR/credentials.json
#
# Every function here is safe to call under `set -e` and inside `$(...)`: a
# failure is a non-zero return with a message on stderr, never an early exit,
# and stdout carries only the access token.
#
# shellcheck disable=SC2154 # ACCOUNT, CONFIG_DIR, CONFIG_FILE, CREDS_FILE come from the caller

OUTLOOK_TOKEN_URL="https://login.microsoftonline.com/common/oauth2/v2.0/token"
OUTLOOK_AUTHORIZE_URL="https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
OUTLOOK_REDIRECT_URI="https://login.microsoftonline.com/common/oauth2/nativeclient"
OUTLOOK_SCOPE="offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read"

# A token is used until it is within this many seconds of expiry.
OUTLOOK_TOKEN_MARGIN=60

# How long to wait for another command that is refreshing the same account.
OUTLOOK_TOKEN_LOCK_WAIT="${OUTLOOK_TOKEN_LOCK_WAIT:-30}"

# Print the stored access token if it is still good for more than the margin.
# Returns 1, printing nothing, when it is missing, expired or about to expire.
# A missing or zero expires_at always counts as expired.
outlook_cached_token() {
    local access_token expires_at now
    access_token=$(jq -r '.access_token // empty' "$CREDS_FILE" 2>/dev/null) || return 1
    expires_at=$(jq -r '(.expires_at | numbers | floor) // 0' "$CREDS_FILE" 2>/dev/null) || return 1
    [ -n "$expires_at" ] || expires_at=0
    now=$(date +%s)

    if [ -n "$access_token" ] && [ "$now" -lt "$((expires_at - OUTLOOK_TOKEN_MARGIN))" ]; then
        printf '%s\n' "$access_token"
        return 0
    fi
    return 1
}

# Run a command while holding this account's token lock, so two commands that
# find the token expired at the same moment do not both refresh it. flock is
# not installed everywhere (macOS lacks it); without it the refresh still runs,
# and the write below is still atomic, so the worst case is one extra refresh.
_outlook_with_token_lock() {
    local lock="$CONFIG_DIR/.token.lock"
    if ! command -v flock >/dev/null 2>&1 || ! : 2>/dev/null >> "$lock"; then
        "$@"
        return
    fi
    (
        if ! flock -w "$OUTLOOK_TOKEN_LOCK_WAIT" 9; then
            echo "Error: another command has held the token lock for ${OUTLOOK_TOKEN_LOCK_WAIT}s. Try again." >&2
            exit 1
        fi
        "$@"
    ) 9>> "$lock"
}

# Refresh the access token and print it. Never leaves credentials.json worse
# than it found it:
#   - curl runs with --fail-with-body and timeouts, so an HTTP error or a
#     stalled connection is a failure rather than an empty "success";
#   - nothing is written unless the response is JSON with an access_token;
#   - the old refresh_token is kept when the response does not carry a new one;
#   - the new file is written beside the old one and renamed into place, so a
#     reader sees the old file or the new one, never a half-written one.
_outlook_refresh_unlocked() {
    local refresh_token client_id client_secret now response rc err tmp

    refresh_token=$(jq -r '.refresh_token // empty' "$CREDS_FILE" 2>/dev/null) || refresh_token=""
    if [ -z "$refresh_token" ]; then
        echo "Error: no refresh token in $CREDS_FILE." >&2
        echo "Sign in again: outlook-setup.sh --account $ACCOUNT" >&2
        return 1
    fi
    client_id=$(jq -r '.client_id // empty' "$CONFIG_FILE" 2>/dev/null) || client_id=""
    client_secret=$(jq -r '.client_secret // empty' "$CONFIG_FILE" 2>/dev/null) || client_secret=""

    # A public client (every setup since PKCE) has no secret and must not send
    # one: Microsoft refuses a secret from a public client. An install set up
    # before that still has one in config.json, and keeps sending it, so it
    # keeps working until it is set up again.
    local secret_arg=()
    [ -n "$client_secret" ] && secret_arg=(--data-urlencode "client_secret=$client_secret")

    now=$(date +%s)
    response=$(curl -sS --fail-with-body --connect-timeout 10 --max-time 60 \
        -X POST "$OUTLOOK_TOKEN_URL" \
        -H "Content-Type: application/x-www-form-urlencoded" \
        --data-urlencode "client_id=$client_id" \
        ${secret_arg[@]+"${secret_arg[@]}"} \
        --data-urlencode "refresh_token=$refresh_token" \
        --data-urlencode "grant_type=refresh_token" \
        --data-urlencode "scope=$OUTLOOK_SCOPE") && rc=0 || rc=$?

    if ! printf '%s' "$response" \
         | jq -e 'type == "object" and (.access_token | type == "string" and length > 0)' \
           >/dev/null 2>&1; then
        err=$(printf '%s' "$response" \
              | jq -r 'if type == "object" and has("error") then (.error_description // .error | tostring) else empty end' \
                2>/dev/null | head -1) || err=""
        if [ -n "$err" ]; then
            echo "Error: Microsoft refused the token refresh: $err" >&2
            case "$(printf '%s' "$response" | jq -r '.error // empty' 2>/dev/null)" in
                invalid_grant|interaction_required|invalid_client|unauthorized_client)
                    echo "Sign in again: outlook-setup.sh --account $ACCOUNT" >&2 ;;
            esac
        elif [ "$rc" -ne 0 ]; then
            echo "Error: could not refresh the access token (curl exit $rc: network error, timeout or HTTP error)." >&2
        else
            echo "Error: the token endpoint answered without an access token." >&2
        fi
        echo "credentials.json is unchanged. Try again, or run: outlook-token.sh --account $ACCOUNT refresh" >&2
        return 1
    fi

    if ! tmp=$(mktemp "$CONFIG_DIR/.credentials.XXXXXX" 2>/dev/null); then
        echo "Error: cannot write to $CONFIG_DIR. credentials.json is unchanged." >&2
        return 1
    fi
    chmod 600 "$tmp" 2>/dev/null || true
    if ! printf '%s' "$response" \
         | jq --argjson now "$now" --arg rt "$refresh_token" '
             . + {expires_at: ($now + ((.expires_in | tonumber?) // 3600))}
             | if ((.refresh_token // "") | length) == 0 then .refresh_token = $rt else . end
           ' > "$tmp" 2>/dev/null \
       || ! mv -f "$tmp" "$CREDS_FILE" 2>/dev/null; then
        rm -f "$tmp"
        echo "Error: could not save the refreshed token. credentials.json is unchanged." >&2
        return 1
    fi
    jq -r '.access_token' "$CREDS_FILE"
}

# Refresh only if the token is still stale once the lock is held: whoever held
# the lock before us may have just refreshed it.
_outlook_refresh_if_stale() {
    outlook_cached_token && return 0
    _outlook_refresh_unlocked
}

# Force a refresh and print the new access token. Used when Graph rejects a
# token that still looks unexpired (revoked, clock skew) and by
# `outlook-token.sh refresh`.
refresh_access_token() {
    _outlook_with_token_lock _outlook_refresh_unlocked
}

# Print a valid access token, refreshing over the network only when the stored
# one is missing, expired or within the margin of expiry. The common path makes
# no network call at all.
ensure_valid_token() {
    outlook_cached_token && return 0
    _outlook_with_token_lock _outlook_refresh_if_stale
}

# --- Sign-in: PKCE ----------------------------------------------------------------
# Setup signs in with the authorisation-code flow and PKCE (RFC 7636), as a
# public client, so there is no client secret to create, store or expire. The
# verifier stays on this machine; only its SHA-256 goes to the browser, and the
# code the browser brings back is useless without the verifier.

# outlook_random <length>: that many characters from the set RFC 7636 allows in
# a verifier (letters, digits and - . _ ~).
outlook_random() {
    local want="$1" out
    out=$(LC_ALL=C tr -dc 'A-Za-z0-9._~-' < /dev/urandom 2>/dev/null | head -c "$want") || true
    [ "${#out}" -eq "$want" ] || return 1
    printf '%s\n' "$out"
}

# A fresh code verifier: 64 characters (RFC 7636 allows 43 to 128).
outlook_pkce_verifier() {
    outlook_random 64
}

# outlook_pkce_challenge <verifier>: BASE64URL(SHA256(verifier)), no padding.
outlook_pkce_challenge() {
    printf '%s' "$1" | openssl dgst -sha256 -binary | openssl base64 -A | tr '+/' '-_' | tr -d '='
}

# outlook_authorize_url <client_id> <challenge> <state>: the sign-in URL.
outlook_authorize_url() {
    local scope redirect
    scope=$(printf '%s' "$OUTLOOK_SCOPE" | sed 's/ /%20/g')
    redirect=$(printf '%s' "$OUTLOOK_REDIRECT_URI" | sed 's/:/%3A/g; s#/#%2F#g')
    printf '%s?client_id=%s&response_type=code&redirect_uri=%s&response_mode=query&scope=%s&code_challenge=%s&code_challenge_method=S256&state=%s&prompt=select_account\n' \
        "$OUTLOOK_AUTHORIZE_URL" "$1" "$redirect" "$scope" "$2" "$3"
}

# outlook_url_param <url> <name>: the raw value of one query parameter, or
# nothing.
outlook_url_param() {
    printf '%s' "$1" | tr '?&#' '\n\n\n' | sed -n "s/^$2=//p" | head -1
}

# --- Requests: timeouts and throttling -------------------------------------------
# Every Graph request goes through outlook_curl, so every one of them has a
# timeout and handles throttling the same way.
#
# Without a timeout a stalled connection hangs the command for ever. Without
# throttling handling a busy mailbox fails: Graph answers with HTTP 429 and a
# Retry-After header, and Outlook allows only 4 concurrent requests per
# mailbox. See https://learn.microsoft.com/en-us/graph/throttling.
#
# The retry rule:
#   - 429 is retried on any method. Graph throttled the request, so it did
#     nothing, and sending it again is safe.
#   - 503 and 504 are retried only on GET and HEAD. A POST that timed out at a
#     gateway may still have run, and sending it twice could send a mail twice.
#   - The wait is Retry-After when Graph gives one, else 1s, 2s, 4s.
#   - At most OUTLOOK_MAX_RETRIES retries, and no single wait is longer than
#     OUTLOOK_MAX_RETRY_WAIT seconds, so a command cannot hang for minutes.
OUTLOOK_CONNECT_TIMEOUT="${OUTLOOK_CONNECT_TIMEOUT:-10}"
OUTLOOK_MAX_TIME="${OUTLOOK_MAX_TIME:-120}"
# An upload chunk, an attachment download or a whole message in MIME form can
# be large. They get longer.
OUTLOOK_TRANSFER_MAX_TIME="${OUTLOOK_TRANSFER_MAX_TIME:-600}"
OUTLOOK_MAX_RETRIES="${OUTLOOK_MAX_RETRIES:-3}"
OUTLOOK_MAX_RETRY_WAIT="${OUTLOOK_MAX_RETRY_WAIT:-60}"

# The status code of the last response in a curl -D header dump (the last one,
# because a 100 Continue comes first). Prints nothing if there is none.
_outlook_http_status() {
    awk '/^HTTP\//{s=$2} END{if (s != "") print s}' "$1" 2>/dev/null | tr -d '\r'
}

# outlook_should_retry <method> <status>: 0 if a response with this status to
# a request with this method should be sent again.
outlook_should_retry() {
    case "$2" in
        429) return 0 ;;
        503|504)
            case "$(printf '%s' "$1" | tr '[:lower:]' '[:upper:]')" in
                GET|HEAD) return 0 ;;
            esac ;;
    esac
    return 1
}

# outlook_retry_wait <retry-after value or ""> <attempt, from 0>: prints the
# seconds to wait before the next try. A Retry-After in seconds is honoured; a
# missing or unreadable one (an HTTP date, say) gives 1, 2, 4... Capped at
# OUTLOOK_MAX_RETRY_WAIT.
outlook_retry_wait() {
    local after="$1" attempt="$2" wait
    after=$(printf '%s' "$after" | tr -d '[:space:]')
    if [[ "$after" =~ ^[0-9]+$ ]]; then
        wait=$((10#$after))
    else
        wait=$((1 << attempt))
    fi
    [ "$wait" -gt "$OUTLOOK_MAX_RETRY_WAIT" ] && wait="$OUTLOOK_MAX_RETRY_WAIT"
    printf '%s\n' "$wait"
}

# --- Immutable IDs ----------------------------------------------------------------
# Graph's default message and event IDs change when an item moves to another
# folder, Deleted Items included, so an ID from a listing stopped working after
# a move and the user had to list the new folder again. With this header Graph
# returns immutable IDs instead, which keep their value for as long as the item
# stays in the mailbox. It applies only to the request it is sent with, so it
# goes on every request, here. Folder IDs never changed and are unaffected.
# See https://learn.microsoft.com/en-us/graph/outlook-immutable-id.
OUTLOOK_PREFER_IDS='Prefer: IdType="ImmutableId"'

# outlook_curl <method> <curl arguments...>
# Runs curl with the timeouts, retries as above, and prints the body of the
# last attempt. Returns curl's exit status. The caller passes everything but
# the timeouts and the immutable-ID header, including -X, so the method is
# given twice: once here, for the retry rule, and once to curl.
#
# OUTLOOK_CURL_MAX_TIME, set for one call, replaces the normal --max-time.
#
# When Graph answers 400 or above with no JSON error in the body (a gateway's
# HTML page, or nothing at all), a JSON error is printed in its place. Callers
# treat an empty body as success, so without this an empty 503 would read as a
# change that had been made.
outlook_curl() {
    local method="$1"; shift
    local hdr attempt=0 body rc status after wait
    local max_time="${OUTLOOK_CURL_MAX_TIME:-$OUTLOOK_MAX_TIME}"

    if ! hdr=$(mktemp "${TMPDIR:-/tmp}/outlook-headers.XXXXXX" 2>/dev/null); then
        curl -s --connect-timeout "$OUTLOOK_CONNECT_TIMEOUT" --max-time "$max_time" \
             -H "$OUTLOOK_PREFER_IDS" "$@"
        return
    fi
    while :; do
        : > "$hdr"
        body=$(curl -s --connect-timeout "$OUTLOOK_CONNECT_TIMEOUT" --max-time "$max_time" \
                    -H "$OUTLOOK_PREFER_IDS" -D "$hdr" "$@") && rc=0 || rc=$?
        status=$(_outlook_http_status "$hdr")
        if [ "$attempt" -lt "$OUTLOOK_MAX_RETRIES" ] && outlook_should_retry "$method" "$status"; then
            after=$(awk 'tolower($1) == "retry-after:" {v=$2} END{print v}' "$hdr" | tr -d '\r')
            wait=$(outlook_retry_wait "$after" "$attempt")
            attempt=$((attempt + 1))
            echo "Microsoft Graph answered HTTP $status. Retrying in ${wait}s (retry $attempt of $OUTLOOK_MAX_RETRIES)." >&2
            sleep "$wait"
            continue
        fi
        break
    done
    rm -f "$hdr"

    if [ -n "$status" ] && [ "$status" -ge 400 ] 2>/dev/null \
       && ! printf '%s' "$body" | jq -e 'objects | has("error")' >/dev/null 2>&1; then
        body=$(jq -cn --arg s "$status" \
            '{error: {code: ("HTTP" + $s), message: ("Microsoft Graph answered HTTP " + $s + " with no error details.")}}')
    fi
    printf '%s' "$body"
    return "$rc"
}

# --- Read-only mode ------------------------------------------------------------
# OUTLOOK_READ_ONLY=1 makes every command that could change the mailbox or send
# anything refuse, before a token is read or a request is made. It suits a
# triage session: the agent can list and read, and nothing else.
#
# It is an allow-list. Each script names the verbs that only read, and every
# other verb is refused. A verb added later is therefore refused in read-only
# mode until somebody decides it only reads, which is the safe way round.
#
# Any value other than empty, 0, false, no or off turns it on, so a user who
# writes OUTLOOK_READ_ONLY=true gets the protection they asked for.
outlook_read_only() {
    case "$(printf '%s' "${OUTLOOK_READ_ONLY:-}" | tr '[:upper:]' '[:lower:]')" in
        ""|0|false|no|off) return 1 ;;
        *) return 0 ;;
    esac
}

# outlook_read_only_gate <script> <verb> <read-only verb>...
# Returns 0 when the command may run, 1 (with the reason on stderr) when it may
# not. The usage text, printed for no verb or for help, reads nothing.
outlook_read_only_gate() {
    local script="$1" verb="$2" v
    shift 2
    outlook_read_only || return 0
    case "$verb" in ""|help|-h|--help) return 0 ;; esac
    for v in "$@"; do
        [ "$verb" = "$v" ] && return 0
    done
    echo "Refused: OUTLOOK_READ_ONLY is set, and '$script $verb' is not a read-only command." >&2
    echo "Nothing was sent to Microsoft Graph. Unset OUTLOOK_READ_ONLY to run it." >&2
    return 1
}
