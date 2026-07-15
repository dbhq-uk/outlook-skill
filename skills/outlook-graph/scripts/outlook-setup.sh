#!/bin/bash
# Outlook OAuth Setup Script
# Automates Azure app registration and OAuth flow for M365 Outlook access

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

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

CONFIG_DIR="$BASE_DIR/$ACCOUNT"
CONFIG_FILE="$CONFIG_DIR/config.json"
CREDS_FILE="$CONFIG_DIR/credentials.json"

# App name is suffixed per non-default account when a fresh app is created.
if [ "$ACCOUNT" = "default" ]; then
    APP_NAME="Claude-Outlook-Integration"
else
    APP_NAME="Claude-Outlook-Integration-$ACCOUNT"
fi

echo -e "${BLUE}=== Outlook OAuth Setup ===${NC}"
echo -e "Account: ${GREEN}$ACCOUNT${NC}"
echo

# Detect a reusable app registration from an existing account (used later in
# Step 2, but must be known before the dependency check so we don't block users
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

if ! command -v jq &> /dev/null; then
    echo -e "${RED}Error: jq not found${NC}"
    echo "Install: brew install jq (macOS) or apt install jq (Linux)"
    exit 1
fi

if ! command -v curl &> /dev/null; then
    echo -e "${RED}Error: curl not found${NC}"
    exit 1
fi

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

# Step 1: Azure login
if [ -z "$SKIP_APP_CREATE" ]; then
    echo -e "${BLUE}Step 1/7: Azure Login${NC}"
    echo "Logging into Azure..."

    if ! az account show &> /dev/null; then
        az login --use-device-code
    fi

    echo -e "${GREEN}Logged in to Azure${NC}"
    echo
fi

# Step 2: Create or get app registration
echo -e "${BLUE}Step 2/7: App Registration${NC}"

# Offer to reuse an existing account's app registration when one was detected.
# The app is multi-tenant + personal-account, so one app can authorize many
# mailboxes, and additional mailboxes then need no Azure admin rights.
if [ -n "$REUSE_CONFIG" ]; then
    REUSE_NAME=$(basename "$(dirname "$REUSE_CONFIG")")
    echo -e "${YELLOW}Found existing app registration from account '$REUSE_NAME'.${NC}"
    read -rp "Reuse it for '$ACCOUNT'? (recommended) (Y/n): " reuse_ans
    if [[ ! "$reuse_ans" =~ ^[Nn]$ ]]; then
        CLIENT_ID=$(jq -r '.client_id' "$REUSE_CONFIG")
        CLIENT_SECRET=$(jq -r '.client_secret' "$REUSE_CONFIG")
        echo -e "${GREEN}Reusing app: $CLIENT_ID${NC}"
        SKIP_APP_CREATE=1
    fi
fi

if [ -z "$SKIP_APP_CREATE" ]; then
    # Check if app already exists
    EXISTING_APP=$(az ad app list --display-name "$APP_NAME" --query "[0].appId" -o tsv 2>/dev/null || echo "")

    if [ -n "$EXISTING_APP" ] && [ "$EXISTING_APP" != "None" ]; then
        echo -e "${YELLOW}Found existing app: $EXISTING_APP${NC}"
        read -rp "Use existing app? (Y/n): " use_existing
        if [[ "$use_existing" =~ ^[Nn]$ ]]; then
            echo "Creating new app..."
            CLIENT_ID=$(az ad app create \
                --display-name "$APP_NAME-$(date +%s)" \
                --sign-in-audience "AzureADandPersonalMicrosoftAccount" \
                --web-redirect-uris "https://login.microsoftonline.com/common/oauth2/nativeclient" \
                --query appId -o tsv)
        else
            CLIENT_ID="$EXISTING_APP"
        fi
    else
        echo "Creating app registration..."
        CLIENT_ID=$(az ad app create \
            --display-name "$APP_NAME" \
            --sign-in-audience "AzureADandPersonalMicrosoftAccount" \
            --web-redirect-uris "https://login.microsoftonline.com/common/oauth2/nativeclient" \
            --query appId -o tsv)
    fi

    echo -e "${GREEN}App ID: $CLIENT_ID${NC}"
    echo
fi

# Step 3: Create client secret
if [ -z "$SKIP_APP_CREATE" ]; then
    echo -e "${BLUE}Step 3/7: Creating Client Secret${NC}"

    SECRET_RESULT=$(az ad app credential reset \
        --id "$CLIENT_ID" \
        --append \
        --display-name "Claude Code Secret" \
        --years 2 \
        --query password -o tsv)

    CLIENT_SECRET="$SECRET_RESULT"
    echo -e "${GREEN}Client secret created (valid for 2 years)${NC}"
    echo
fi

# Step 4: Add API permissions
if [ -z "$SKIP_APP_CREATE" ]; then
    echo -e "${BLUE}Step 4/7: Configuring API Permissions${NC}"

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

# Step 5: Save config
mkdir -p "$CONFIG_DIR"
chmod 700 "$CONFIG_DIR"
echo -e "${BLUE}Step 5/7: Saving Configuration${NC}"

cat > "$CONFIG_FILE" << EOF
{
    "client_id": "$CLIENT_ID",
    "client_secret": "$CLIENT_SECRET",
    "tenant": "common",
    "redirect_uri": "https://login.microsoftonline.com/common/oauth2/nativeclient",
    "scope": "offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read"
}
EOF

chmod 600 "$CONFIG_FILE"
echo -e "${GREEN}Configuration saved to $CONFIG_FILE${NC}"
echo

# Step 6: OAuth authorization
echo -e "${BLUE}Step 6/7: OAuth Authorization${NC}"

SCOPE="offline_access%20Mail.ReadWrite%20Mail.Send%20Calendars.ReadWrite%20User.Read"
AUTH_URL="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=$CLIENT_ID&response_type=code&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&scope=$SCOPE&prompt=select_account"

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

# Extract authorization code
AUTH_CODE=$(echo "$REDIRECT_URL" | sed -n 's/.*code=\([^&]*\).*/\1/p')

if [ -z "$AUTH_CODE" ]; then
    echo -e "${RED}Error: Could not extract authorization code from URL${NC}"
    exit 1
fi

echo -e "${GREEN}Authorization code received${NC}"
echo

# Exchange code for tokens
echo "Exchanging code for tokens..."

NOW=$(date +%s)
TOKEN_RESPONSE=$(curl -s -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
    -H "Content-Type: application/x-www-form-urlencoded" \
    -d "client_id=$CLIENT_ID" \
    -d "client_secret=$CLIENT_SECRET" \
    -d "code=$AUTH_CODE" \
    -d "redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient" \
    -d "grant_type=authorization_code" \
    -d "scope=offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read")

# Check for error
if echo "$TOKEN_RESPONSE" | jq -e '.error' > /dev/null 2>&1; then
    echo -e "${RED}Error getting tokens:${NC}"
    echo "$TOKEN_RESPONSE" | jq -r '.error_description'
    exit 1
fi

# Save credentials, stamping an absolute expiry so the mail/calendar scripts can
# skip their per-command token pre-flight.
EXPIRES_IN=$(echo "$TOKEN_RESPONSE" | jq -r '.expires_in // 3600')
echo "$TOKEN_RESPONSE" | jq --argjson at "$((NOW + EXPIRES_IN))" '. + {expires_at: $at}' > "$CREDS_FILE"
chmod 600 "$CREDS_FILE"

echo -e "${GREEN}Tokens saved to $CREDS_FILE${NC}"
echo

# Step 7: Test connection
echo -e "${BLUE}Step 7/7: Testing Connection${NC}"

ACCESS_TOKEN=$(jq -r '.access_token' "$CREDS_FILE")

TEST_RESPONSE=$(curl -s -X GET "https://graph.microsoft.com/v1.0/me/mailFolders/inbox" \
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

echo -e "${GREEN}=== Setup Complete ===${NC}"
echo
echo "You can now use the Outlook skill in Claude Code."
echo "Try: 'check my email' or 'what's on my calendar today'"
