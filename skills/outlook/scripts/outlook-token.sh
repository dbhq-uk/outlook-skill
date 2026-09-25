#!/bin/bash
# Outlook Token Management

set -e

BASE_DIR="$HOME/.dbhq/outlook"

# One-time migrations, oldest first. Each is guarded on the NEW directory not
# existing, so an install that has already moved is left alone and a second run
# does nothing. The skill has had three homes:
#
#   ~/.outlook-graph            before the ~/.dbhq rule (10 Sep 2026)
#   ~/.dbhq/outlook-graph       before the rename (17 Sep 2026)
#   ~/.dbhq/outlook             now
#
# The oldest path is checked against the OLD skill directory rather than the
# new one, because an install still sitting at ~/.outlook-graph never saw the
# middle step and has to make both hops.
if [ ! -e "$HOME/.dbhq/outlook-graph" ] && [ ! -e "$BASE_DIR" ] \
   && [ -d "$HOME/.outlook-graph" ]; then
    mkdir -p "$HOME/.dbhq"
    chmod 700 "$HOME/.dbhq"
    mv "$HOME/.outlook-graph" "$HOME/.dbhq/outlook-graph"
    chmod 700 "$HOME/.dbhq/outlook-graph"
fi

if [ ! -e "$BASE_DIR" ] && [ -d "$HOME/.dbhq/outlook-graph" ]; then
    mv "$HOME/.dbhq/outlook-graph" "$BASE_DIR"
    chmod 700 "$BASE_DIR"
fi

# Account resolution: --account/-a flag wins, else OUTLOOK_ACCOUNT env, else "default"
ACCOUNT="${OUTLOOK_ACCOUNT:-default}"
if [ "$1" = "--account" ] || [ "$1" = "-a" ]; then
    [ -n "$2" ] || { echo "Error: $1 requires an account name" >&2; exit 1; }
    ACCOUNT="$2"; shift 2
fi

# One-time migration: legacy flat config -> default/
if [ -f "$BASE_DIR/config.json" ] && [ ! -d "$BASE_DIR/default" ]; then
    mkdir -p "$BASE_DIR/default"
    chmod 700 "$BASE_DIR" "$BASE_DIR/default"
    mv "$BASE_DIR/config.json" "$BASE_DIR/credentials.json" "$BASE_DIR/id_cache.json" \
       "$BASE_DIR/default/" 2>/dev/null || true
fi

# `list` must work without a configured account, so handle it before the config check.
if [ "$1" = "list" ]; then
    echo "Configured accounts:"
    found=0
    for dir in "$BASE_DIR"/*/; do
        [ -f "$dir/credentials.json" ] || continue
        echo "  - $(basename "$dir")"
        found=1
    done
    [ "$found" = 0 ] && echo "  (none configured — run outlook-setup.sh)"
    exit 0
fi

CONFIG_DIR="$BASE_DIR/$ACCOUNT"
CONFIG_FILE="$CONFIG_DIR/config.json"
CREDS_FILE="$CONFIG_DIR/credentials.json"

# Check config exists
if [ ! -f "$CONFIG_FILE" ]; then
    echo "Error: Account '$ACCOUNT' not configured."
    echo "Run: outlook-setup.sh --account $ACCOUNT"
    exit 1
fi

# The token code is shared with the mail and calendar scripts in lib/graph.sh,
# found beside this script's real location (symlinks followed).
_self="${BASH_SOURCE[0]}"
while [ -L "$_self" ]; do
    _dir=$(cd -P "$(dirname "$_self")" && pwd)
    _self=$(readlink "$_self")
    case "$_self" in /*) ;; *) _self="$_dir/$_self" ;; esac
done
OUTLOOK_SCRIPT_DIR=$(cd -P "$(dirname "$_self")" && pwd)
unset _self _dir
# shellcheck source=lib/graph.sh
. "$OUTLOOK_SCRIPT_DIR/lib/graph.sh"

case "$1" in
    refresh)
        if [ ! -f "$CREDS_FILE" ]; then
            echo "Error: No credentials to refresh. Run outlook-setup.sh first."
            exit 1
        fi

        echo "Refreshing token..."
        # The shared refresh prints the new token on stdout; this command only
        # reports the outcome. On failure it has already said why on stderr and
        # left credentials.json unchanged.
        if ! refresh_access_token > /dev/null; then
            exit 1
        fi
        echo "Token refreshed successfully"
        ;;

    get)
        if [ ! -f "$CREDS_FILE" ]; then
            echo "Error: No credentials found."
            exit 1
        fi

        # Goes through the same check as every other command: a token that is
        # expired or within 60 seconds of expiry is refreshed first, so a
        # hand-written Graph call made with it does not fail with 401. Nothing
        # but the token reaches stdout, so T=$(outlook-token.sh get) stays clean.
        if ! ensure_valid_token; then
            exit 1
        fi
        ;;

    test)
        if [ ! -f "$CREDS_FILE" ]; then
            echo "Error: No credentials found. Run outlook-setup.sh first."
            exit 1
        fi

        # Refresh first if needed, as a real command would, so an expired
        # token is not reported as a broken connection.
        if ! ACCESS_TOKEN=$(ensure_valid_token); then
            exit 1
        fi

        echo "Testing connection..."

        RESPONSE=$(curl -s --connect-timeout "$OUTLOOK_CONNECT_TIMEOUT" --max-time "$OUTLOOK_MAX_TIME" \
            -X GET "https://graph.microsoft.com/v1.0/me/mailFolders/inbox" \
            -H "Authorization: Bearer $ACCESS_TOKEN")

        if echo "$RESPONSE" | jq -e '.error' > /dev/null 2>&1; then
            ERROR=$(echo "$RESPONSE" | jq -r '.error.code')
            if [ "$ERROR" = "InvalidAuthenticationToken" ]; then
                echo "Token expired. Run: outlook-token.sh refresh"
            else
                echo "Error:"
                echo "$RESPONSE" | jq -r '.error.message'
            fi
            exit 1
        fi

        TOTAL=$(echo "$RESPONSE" | jq -r '.totalItemCount')
        UNREAD=$(echo "$RESPONSE" | jq -r '.unreadItemCount')

        echo "Connection successful!"
        echo "Inbox: $TOTAL total, $UNREAD unread"
        ;;

    status)
        if [ ! -f "$CREDS_FILE" ]; then
            echo "Status: Not configured"
            echo "Run: outlook-setup.sh"
            exit 0
        fi

        ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE")

        # Quick test
        RESPONSE=$(curl -s --connect-timeout "$OUTLOOK_CONNECT_TIMEOUT" --max-time "$OUTLOOK_MAX_TIME" \
            -X GET "https://graph.microsoft.com/v1.0/me" \
            -H "Authorization: Bearer $ACCESS_TOKEN")

        if echo "$RESPONSE" | jq -e '.error' > /dev/null 2>&1; then
            echo "Status: Token expired"
            echo "Run: outlook-token.sh refresh"
        else
            NAME=$(echo "$RESPONSE" | jq -r '.displayName // .mail // "Unknown"')
            echo "Status: Connected"
            echo "Account: $NAME"
        fi
        ;;

    *)
        echo "Outlook Token Management"
        echo
        echo "Usage: outlook-token.sh <command>"
        echo
        echo "Commands:"
        echo "  refresh    Refresh the access token"
        echo "  get        Print a valid access token (refreshed first if needed)"
        echo "  test       Test connection to Outlook"
        echo "  status     Show connection status"
        echo "  list       List configured accounts"
        echo
        echo "Account selection: --account <name> | -a <name> | OUTLOOK_ACCOUNT env (default: default)"
        ;;
esac
