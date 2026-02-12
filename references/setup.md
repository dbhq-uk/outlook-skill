# Outlook Manual Setup Guide

If you prefer to set up the Azure app registration manually (instead of using `outlook-setup.sh`), follow these steps.

## Prerequisites

- Azure account (same account as your M365 subscription, or ability to create app registrations)
- `jq` and `curl` installed locally

## Step 1: Create Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure:
   - **Name:** `Claude-Outlook-Integration` (or your preferred name)
   - **Supported account types:** "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI:** Web → `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Click **Register**

## Step 2: Note Your Application ID

After registration, you'll see the **Application (client) ID** on the Overview page.

Copy this - you'll need it later.

## Step 3: Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Configure:
   - **Description:** `Claude Code Secret`
   - **Expires:** 24 months (recommended)
4. Click **Add**
5. **IMPORTANT:** Copy the secret **Value** immediately - you can only see it once!

## Step 4: Configure API Permissions

1. Go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph** → **Delegated permissions**
4. Add these permissions:
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `Calendars.ReadWrite`
   - `offline_access`
   - `User.Read`
5. Click **Add permissions**

Note: Admin consent is NOT required for delegated permissions with personal/org accounts.

## Step 5: Create Config Files

Create the config directory:

```bash
mkdir -p ~/.outlook
chmod 700 ~/.outlook
```

Create `~/.outlook/config.json`:

```json
{
    "client_id": "YOUR_APPLICATION_ID",
    "client_secret": "YOUR_CLIENT_SECRET",
    "tenant": "common",
    "redirect_uri": "https://login.microsoftonline.com/common/oauth2/nativeclient",
    "scope": "offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read"
}
```

Set permissions:

```bash
chmod 600 ~/.outlook/config.json
```

## Step 6: Get Authorization Code

Build the authorization URL (replace YOUR_CLIENT_ID):

```
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=YOUR_CLIENT_ID&response_type=code&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&scope=offline_access%20Mail.ReadWrite%20Mail.Send%20Calendars.ReadWrite%20User.Read
```

1. Open this URL in your browser
2. Sign in with your M365 account
3. Accept the permissions
4. You'll be redirected to a blank page
5. Copy the **entire URL** from your browser's address bar
6. Extract the `code` parameter value (everything between `code=` and `&`)

## Step 7: Exchange Code for Tokens

Run this command (replace placeholders):

```bash
curl -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "client_id=YOUR_CLIENT_ID" \
  -d "client_secret=YOUR_CLIENT_SECRET" \
  -d "code=YOUR_AUTHORIZATION_CODE" \
  -d "redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient" \
  -d "grant_type=authorization_code" \
  -d "scope=offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read" \
  > ~/.outlook/credentials.json

chmod 600 ~/.outlook/credentials.json
```

## Step 8: Verify Setup

Test the connection:

```bash
~/.claude/skills/outlook/scripts/outlook-token.sh test
```

You should see:
```
Connection successful!
Inbox: X total, Y unread
```

## Troubleshooting

### "Invalid client secret"
- Client secrets can only be viewed once when created
- Create a new secret if you lost the original

### "AADSTS50011: Reply URL does not match"
- Ensure redirect URI in Azure exactly matches: `https://login.microsoftonline.com/common/oauth2/nativeclient`

### "Token expired"
Tokens are automatically refreshed. If you see this error, it means the refresh also failed - likely due to expired refresh token.

### "Refresh token expired"
- Refresh tokens last ~90 days with activity
- If fully expired, re-run the authorization flow (Steps 6-7)

### "Insufficient privileges"
- Verify all permissions are added in Azure
- Try removing and re-adding the permissions
- Sign out and sign back in to re-consent

## File Locations

| File | Purpose |
|------|---------|
| `~/.outlook/config.json` | Azure app credentials |
| `~/.outlook/credentials.json` | OAuth tokens |
| `~/.claude/skills/outlook/` | Skill and scripts |
