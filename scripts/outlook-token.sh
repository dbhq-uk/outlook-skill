#!/bin/bash
# Outlook Token Management

set -e

BASE_DIR="$HOME/.outlook"

# Account resolution: --account/-a flag wins, else OUTLOOK_ACCOUNT env, else "default"
ACCOUNT="${OUTLOOK_ACCOUNT:-default}"
if [ "$1" = "--account" ] || [ "$1" = "-a" ]; then
    [ -n "$2" ] || { echo "Error: $1 requires an account name" >&2; exit 1; }
    ACCOUNT="$2"; shift 2
fi

# One-time migration: legacy flat config -> default/
if [ -f "$BASE_DIR/config.json" ] && [ ! -d "$BASE_DIR/default" ]; then
    mkdir -p "$BASE_DIR/default"
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

CLIENT_ID=$(jq -r '.client_id' "$CONFIG_FILE")
CLIENT_SECRET=$(jq -r '.client_secret' "$CONFIG_FILE")
SCOPE="offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read"

case "$1" in
    refresh)
        if [ ! -f "$CREDS_FILE" ]; then
            echo "Error: No credentials to refresh. Run outlook-setup.sh first."
            exit 1
        fi

        REFRESH_TOKEN=$(jq -r '.refresh_token' "$CREDS_FILE")

        if [ -z "$REFRESH_TOKEN" ] || [ "$REFRESH_TOKEN" = "null" ]; then
            echo "Error: No refresh token found. Run outlook-setup.sh to re-authenticate."
            exit 1
        fi

        echo "Refreshing token..."

        RESPONSE=$(curl -s -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
            -H "Content-Type: application/x-www-form-urlencoded" \
            -d "client_id=$CLIENT_ID" \
            -d "client_secret=$CLIENT_SECRET" \
            -d "refresh_token=$REFRESH_TOKEN" \
            -d "grant_type=refresh_token" \
            -d "scope=$SCOPE")

        # Check for error
        if echo "$RESPONSE" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error refreshing token:"
            echo "$RESPONSE" | jq -r '.error_description'
            exit 1
        fi

        # Save new credentials
        echo "$RESPONSE" > "$CREDS_FILE"
        chmod 600 "$CREDS_FILE"

        echo "Token refreshed successfully"
        ;;

    get)
        if [ ! -f "$CREDS_FILE" ]; then
            echo "Error: No credentials found."
            exit 1
        fi

        jq -r '.access_token' "$CREDS_FILE"
        ;;

    test)
        if [ ! -f "$CREDS_FILE" ]; then
            echo "Error: No credentials found. Run outlook-setup.sh first."
            exit 1
        fi

        ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE")

        echo "Testing connection..."

        RESPONSE=$(curl -s -X GET "https://graph.microsoft.com/v1.0/me/mailFolders/inbox" \
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
        RESPONSE=$(curl -s -X GET "https://graph.microsoft.com/v1.0/me" \
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
        echo "  get        Output current access token"
        echo "  test       Test connection to Outlook"
        echo "  status     Show connection status"
        echo "  list       List configured accounts"
        echo
        echo "Account selection: --account <name> | -a <name> | OUTLOOK_ACCOUNT env (default: default)"
        ;;
esac
