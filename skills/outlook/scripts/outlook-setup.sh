#!/bin/bash
# Outlook OAuth Setup Script
# Registers (or reuses) an Azure app and signs a mailbox in, for M365 Outlook
# access. Needs a browser and a person: an agent should ask the user to run it.
#
# The app is a PUBLIC client and sign-in uses the authorisation-code flow with
# PKCE, so there is no client secret. Setup used to register a confidential web
# app with a two-year secret, shared by every account set up after it; when that
# secret expired, every account stopped at once, with no warning. An install
# from then still works (its secret is still sent on refresh), and running this
# setup again moves it over. See docs/guides/accounts.md.

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

BASE_DIR="$HOME/.dbhq/outlook"

# One-time migrations, oldest first. Each is guarded on the NEW directory not
# existing, so an install that has already moved is left alone and a second run
# does nothing. The skill has had three homes:
#
#   ~/.outlook-graph            before the ~/.dbhq rule (10 Sep 2026)
#   ~/.dbhq/outlook-graph       before the rename (17 Sep 2026)
#   ~/.dbhq/outlook             now
#
# EVERY ENTRY SCRIPT CARRIES THIS, and that is the point rather than
# duplication for its own sake: whichever one a user or an agent runs first has
# to be the one that moves the settings. A migration in only one script is a
# migration that has not run.
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

CONFIG_DIR="$BASE_DIR/$ACCOUNT"
CONFIG_FILE="$CONFIG_DIR/config.json"
CREDS_FILE="$CONFIG_DIR/credentials.json"

# The sign-in constants and the PKCE helpers live in lib/graph.sh, beside this
# script's real location (symlinks followed), shared with the other scripts.
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

# App name is suffixed per non-default account when a fresh app is created.
if [ "$ACCOUNT" = "default" ]; then
    APP_NAME="Claude-Outlook-Integration"
else
    APP_NAME="Claude-Outlook-Integration-$ACCOUNT"
fi

echo -e "${BLUE}=== Outlook OAuth Setup ===${NC}"
echo -e "Account: ${GREEN}$ACCOUNT${NC}"
echo

# Make an existing app registration a public client: the native-client
# redirect URI moves from the web platform to "Mobile and desktop
# applications", and public client flows are allowed. Other redirect URIs are
# left as they are. Existing client secrets are not touched, so any account
# still using one keeps working until it is set up again.
make_public_client() {
    local app_id="$1" web public body
    web=$(az ad app show --id "$app_id" --query "web.redirectUris" -o json 2>/dev/null) || web='[]'
    public=$(az ad app show --id "$app_id" --query "publicClient.redirectUris" -o json 2>/dev/null) || public='[]'
    body=$(jq -cn --arg uri "$OUTLOOK_REDIRECT_URI" --argjson web "${web:-[]}" --argjson public "${public:-[]}" '{
        isFallbackPublicClient: true,
        web: {redirectUris: (($web // []) - [$uri])},
        publicClient: {redirectUris: ((($public // []) - [$uri]) + [$uri])}
    }')
    az rest --method PATCH \
        --uri "https://graph.microsoft.com/v1.0/applications(appId='$app_id')" \
        --headers "Content-Type=application/json" \
        --body "$body" > /dev/null
}

# Detect a reusable app registration from an existing account (used later,
# but must be known before the dependency check so we don't block users
# who lack az but can reuse an existing app).
REUSE_CONFIG=""
for dir in "$BASE_DIR"/*/; do
    other="$dir/config.json"
    [ "$dir" = "$CONFIG_DIR/" ] && continue
    [ -f "$other" ] || continue
    REUSE_CONFIG="$other"
    break
done

# Check dependencies
echo -e "${YELLOW}Checking dependencies...${NC}"

if [ -z "$REUSE_CONFIG" ] && ! command -v az &> /dev/null; then
    echo -e "${RED}Error: Azure CLI (az) not found${NC}"
    echo "Install: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli"
    exit 1
fi

for tool in jq curl openssl; do
    if ! command -v "$tool" &> /dev/null; then
        echo -e "${RED}Error: $tool not found${NC}"
        echo "Install: brew install $tool (macOS) or apt install $tool (Linux)"
        exit 1
    fi
done

echo -e "${GREEN}All dependencies found${NC}"
echo

# Check for existing config
if [ -f "$CONFIG_FILE" ]; then
    echo -e "${YELLOW}Existing configuration found at $CONFIG_FILE${NC}"
    read -rp "Overwrite? (y/N): " overwrite
    if [[ ! "$overwrite" =~ ^[Yy]$ ]]; then
        echo "Setup cancelled"
        exit 0
    fi
fi

# Decide which app registration to use. Nothing in Azure changes yet.
CLIENT_ID=""
CLIENT_SECRET=""       # set only when the user keeps an old app's secret
CONVERT_APP=""         # set when an existing app must become a public client

# Offer to reuse an existing account's app registration when one was detected.
# The app is multi-tenant + personal-account, so one app can authorize many
# mailboxes, and additional mailboxes then need no Azure admin rights.
if [ -n "$REUSE_CONFIG" ]; then
    REUSE_NAME=$(basename "$(dirname "$REUSE_CONFIG")")
    echo -e "${BLUE}App registration${NC}"
    echo -e "${YELLOW}Found existing app registration from account '$REUSE_NAME'.${NC}"
    read -rp "Reuse it for '$ACCOUNT'? (recommended) (Y/n): " reuse_ans
    if [[ ! "$reuse_ans" =~ ^[Nn]$ ]]; then
        CLIENT_ID=$(jq -r '.client_id // empty' "$REUSE_CONFIG")
        reuse_secret=$(jq -r '.client_secret // empty' "$REUSE_CONFIG")
        if [ -n "$reuse_secret" ]; then
            echo
            echo -e "${YELLOW}That app was registered with a client secret. Secrets expire, and every${NC}"
            echo -e "${YELLOW}account sharing this one stops at once when it does. As a public client it${NC}"
            echo -e "${YELLOW}needs no secret. Converting it needs the Azure CLI and rights on the app.${NC}"
            read -rp "Convert it to a public client now? (Y/n): " convert_ans
            if [[ "$convert_ans" =~ ^[Nn]$ ]]; then
                CLIENT_SECRET="$reuse_secret"
                echo -e "${YELLOW}Keeping the secret. This account stops working when it expires.${NC}"
            elif ! command -v az &> /dev/null; then
                echo -e "${RED}Error: converting needs the Azure CLI (az), which is not installed.${NC}"
                echo "Install it, or answer n to keep the secret for now."
                exit 1
            else
                CONVERT_APP=1
            fi
        fi
        echo -e "${GREEN}Reusing app: $CLIENT_ID${NC}"
        echo
    fi
fi

if [ -z "$CLIENT_ID" ] || [ -n "$CONVERT_APP" ]; then
    echo -e "${BLUE}Azure login${NC}"
    if ! az account show &> /dev/null; then
        az login --use-device-code
    fi
    echo -e "${GREEN}Logged in to Azure${NC}"
    echo
fi

if [ -z "$CLIENT_ID" ]; then
    echo -e "${BLUE}App registration${NC}"
    EXISTING_APP=$(az ad app list --display-name "$APP_NAME" --query "[0].appId" -o tsv 2>/dev/null || echo "")

    if [ -n "$EXISTING_APP" ] && [ "$EXISTING_APP" != "None" ]; then
        echo -e "${YELLOW}Found existing app: $EXISTING_APP${NC}"
        read -rp "Use existing app? (Y/n): " use_existing
        if [[ ! "$use_existing" =~ ^[Nn]$ ]]; then
            CLIENT_ID="$EXISTING_APP"
            CONVERT_APP=1
        fi
        NEW_APP_NAME="$APP_NAME-$(date +%s)"
    else
        NEW_APP_NAME="$APP_NAME"
    fi

    if [ -z "$CLIENT_ID" ]; then
        echo "Creating app registration (public client, no secret)..."
        CLIENT_ID=$(az ad app create \
            --display-name "$NEW_APP_NAME" \
            --sign-in-audience "AzureADandPersonalMicrosoftAccount" \
            --public-client-redirect-uris "$OUTLOOK_REDIRECT_URI" \
            --is-fallback-public-client true \
            --query appId -o tsv)
    fi
    echo -e "${GREEN}App ID: $CLIENT_ID${NC}"
    echo
    ADD_PERMISSIONS=1
fi

if [ -z "$CLIENT_ID" ]; then
    echo -e "${RED}Error: no app registration to use.${NC}"
    exit 1
fi

# An app made before PKCE becomes a public client.
if [ -n "$CONVERT_APP" ]; then
    echo -e "${BLUE}Making the app a public client${NC}"
    if ! make_public_client "$CLIENT_ID"; then
        echo -e "${RED}Error: could not update app $CLIENT_ID.${NC}"
        echo "You may lack rights on it. In the Azure portal, under Authentication:"
        echo "  - add the platform 'Mobile and desktop applications' with the redirect URI"
        echo "    $OUTLOOK_REDIRECT_URI"
        echo "  - remove that URI from the Web platform"
        echo "  - set 'Allow public client flows' to Yes"
        echo "Then run this setup again."
        exit 1
    fi
    echo -e "${GREEN}App $CLIENT_ID is a public client. Its old secrets are untouched.${NC}"
    echo
fi

# Add API permissions to a new or found app
if [ -n "${ADD_PERMISSIONS:-}" ]; then
    echo -e "${BLUE}Configuring API permissions${NC}"

    # Microsoft Graph API ID
    GRAPH_API="00000003-0000-0000-c000-000000000000"

    # Permission IDs (delegated)
    MAIL_READ_WRITE="024d486e-b451-40bb-833d-3e66d98c5c73"    # Mail.ReadWrite
    MAIL_SEND="e383f46e-2787-4529-855e-0e479a3ffac0"          # Mail.Send
    CALENDARS_READ_WRITE="1ec239c2-d7c9-4623-a91a-a9775856bb36" # Calendars.ReadWrite
    OFFLINE_ACCESS="7427e0e9-2fba-42fe-b0c0-848c9e6a8182"     # offline_access
    USER_READ="e1fe6dd8-ba31-4d61-89e7-88639da4683d"          # User.Read

    echo "Adding Mail.ReadWrite..."
    az ad app permission add --id "$CLIENT_ID" --api "$GRAPH_API" --api-permissions "$MAIL_READ_WRITE=Scope" 2>/dev/null || true

    echo "Adding Mail.Send..."
    az ad app permission add --id "$CLIENT_ID" --api "$GRAPH_API" --api-permissions "$MAIL_SEND=Scope" 2>/dev/null || true

    echo "Adding Calendars.ReadWrite..."
    az ad app permission add --id "$CLIENT_ID" --api "$GRAPH_API" --api-permissions "$CALENDARS_READ_WRITE=Scope" 2>/dev/null || true

    echo "Adding offline_access..."
    az ad app permission add --id "$CLIENT_ID" --api "$GRAPH_API" --api-permissions "$OFFLINE_ACCESS=Scope" 2>/dev/null || true

    echo "Adding User.Read..."
    az ad app permission add --id "$CLIENT_ID" --api "$GRAPH_API" --api-permissions "$USER_READ=Scope" 2>/dev/null || true

    echo -e "${GREEN}Permissions configured${NC}"
    echo
fi

# Sign in, with PKCE
echo -e "${BLUE}Sign in${NC}"

if ! CODE_VERIFIER=$(outlook_pkce_verifier) || ! STATE=$(outlook_random 32); then
    echo -e "${RED}Error: could not read /dev/urandom for the sign-in.${NC}"
    exit 1
fi
CODE_CHALLENGE=$(outlook_pkce_challenge "$CODE_VERIFIER")
AUTH_URL=$(outlook_authorize_url "$CLIENT_ID" "$CODE_CHALLENGE" "$STATE")

echo -e "${YELLOW}Opening browser for Microsoft login...${NC}"
echo
echo "If browser doesn't open, visit this URL:"
echo -e "${BLUE}$AUTH_URL${NC}"
echo

# Try to open browser
if command -v xdg-open &> /dev/null; then
    xdg-open "$AUTH_URL" 2>/dev/null || true
elif command -v open &> /dev/null; then
    open "$AUTH_URL" 2>/dev/null || true
fi

echo -e "${YELLOW}After signing in, you'll be redirected to a blank page.${NC}"
echo -e "${YELLOW}Copy the ENTIRE URL from your browser's address bar and paste it here:${NC}"
echo
read -rp "Paste redirect URL: " REDIRECT_URL

AUTH_CODE=$(outlook_url_param "$REDIRECT_URL" code)
RETURNED_STATE=$(outlook_url_param "$REDIRECT_URL" state)

if [ -z "$AUTH_CODE" ]; then
    SIGNIN_ERROR=$(outlook_url_param "$REDIRECT_URL" error_description | sed 's/+/ /g; s/%20/ /g')
    echo -e "${RED}Error: Could not extract authorization code from URL${NC}"
    [ -n "$SIGNIN_ERROR" ] && echo "Microsoft said: $SIGNIN_ERROR"
    exit 1
fi

# The state ties the pasted URL to this sign-in. A URL from an earlier or
# different sign-in would not redeem against this verifier anyway, so refuse it
# with a clear reason rather than an opaque token error.
if [ "$RETURNED_STATE" != "$STATE" ]; then
    echo -e "${RED}Error: that URL is from a different sign-in (state does not match).${NC}"
    echo "Run setup again and paste the URL from the sign-in it opens."
    exit 1
fi

echo -e "${GREEN}Authorization code received${NC}"
echo

# Exchange code for tokens. No client secret: the code verifier proves this is
# the machine that started the sign-in. Only an app kept on its old secret
# (declined above) still sends one.
echo "Exchanging code for tokens..."

secret_arg=()
[ -n "$CLIENT_SECRET" ] && secret_arg=(--data-urlencode "client_secret=$CLIENT_SECRET")

NOW=$(date +%s)
TOKEN_RESPONSE=$(curl -s --connect-timeout 10 --max-time 60 \
    -X POST "$OUTLOOK_TOKEN_URL" \
    -H "Content-Type: application/x-www-form-urlencoded" \
    --data-urlencode "client_id=$CLIENT_ID" \
    ${secret_arg[@]+"${secret_arg[@]}"} \
    -d "code=$AUTH_CODE" \
    --data-urlencode "code_verifier=$CODE_VERIFIER" \
    --data-urlencode "redirect_uri=$OUTLOOK_REDIRECT_URI" \
    --data-urlencode "grant_type=authorization_code" \
    --data-urlencode "scope=$OUTLOOK_SCOPE")

# Check for error
if ! printf '%s' "$TOKEN_RESPONSE" | jq -e '.access_token | type == "string" and length > 0' > /dev/null 2>&1; then
    echo -e "${RED}Error getting tokens:${NC}"
    printf '%s' "$TOKEN_RESPONSE" | jq -r '.error_description // .error // .' 2>/dev/null || printf '%s\n' "$TOKEN_RESPONSE"
    exit 1
fi

# Save config and credentials together, only now that sign-in has worked, so a
# failed sign-in leaves an existing account as it was. The config carries no
# client_secret unless the user chose to keep an old app's secret.
mkdir -p "$CONFIG_DIR"
chmod 700 "$HOME/.dbhq" "$BASE_DIR" "$CONFIG_DIR"

jq -n --arg id "$CLIENT_ID" --arg secret "$CLIENT_SECRET" \
      --arg redirect "$OUTLOOK_REDIRECT_URI" --arg scope "$OUTLOOK_SCOPE" '
    {client_id: $id}
    + (if $secret != "" then {client_secret: $secret} else {} end)
    + {tenant: "common", redirect_uri: $redirect, scope: $scope}
' > "$CONFIG_FILE"
chmod 600 "$CONFIG_FILE"
echo -e "${GREEN}Configuration saved to $CONFIG_FILE${NC}"

# Stamp an absolute expiry so the mail/calendar scripts can skip their
# per-command token pre-flight.
EXPIRES_IN=$(printf '%s' "$TOKEN_RESPONSE" | jq -r '(.expires_in | tonumber?) // 3600')
printf '%s' "$TOKEN_RESPONSE" | jq --argjson at "$((NOW + EXPIRES_IN))" '. + {expires_at: $at}' > "$CREDS_FILE"
chmod 600 "$CREDS_FILE"

echo -e "${GREEN}Tokens saved to $CREDS_FILE${NC}"
echo

# Test the connection
echo -e "${BLUE}Testing the connection${NC}"

ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE")

TEST_RESPONSE=$(curl -s --connect-timeout 10 --max-time 60 \
    -X GET "https://graph.microsoft.com/v1.0/me/mailFolders/inbox" \
    -H "Authorization: Bearer $ACCESS_TOKEN")

if echo "$TEST_RESPONSE" | jq -e '.error' > /dev/null 2>&1; then
    echo -e "${RED}Connection test failed:${NC}"
    echo "$TEST_RESPONSE" | jq -r '.error.message'
    exit 1
fi

TOTAL=$(echo "$TEST_RESPONSE" | jq -r '.totalItemCount')
UNREAD=$(echo "$TEST_RESPONSE" | jq -r '.unreadItemCount')

echo -e "${GREEN}Connection successful!${NC}"
echo -e "Inbox: ${TOTAL} total, ${UNREAD} unread"
echo

# Other accounts still on an old secret for this same app: say so once.
for dir in "$BASE_DIR"/*/; do
    other="$dir/config.json"
    [ "$dir" = "$CONFIG_DIR/" ] && continue
    [ -f "$other" ] || continue
    if [ "$(jq -r '.client_id // empty' "$other")" = "$CLIENT_ID" ] \
       && [ -n "$(jq -r '.client_secret // empty' "$other")" ] && [ -z "$CLIENT_SECRET" ]; then
        echo -e "${YELLOW}Account '$(basename "$dir")' still uses this app's old secret. Run setup for it too:${NC}"
        echo "  outlook-setup.sh --account $(basename "$dir")"
    fi
done

echo -e "${GREEN}=== Setup Complete ===${NC}"
echo
echo "You can now use the Outlook skill in Claude Code."
echo "Try: 'check my email' or 'what's on my calendar today'"
