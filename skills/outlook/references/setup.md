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
   - **Redirect URI:** Public client/native (mobile & desktop) → `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Click **Register**
6. Under **Authentication**, set **Allow public client flows** to **Yes**, and save

This makes the app a public client. It signs in with PKCE and has no client secret, so there
is nothing to expire. Do not add a Web platform redirect for the same URI: a Web redirect
makes Microsoft ask for a secret.

## Step 2: Note Your Application ID

After registration, you'll see the **Application (client) ID** on the Overview page.

Copy this - you'll need it later.

## Step 3: No client secret

A public client needs none. Skip **Certificates & secrets**.

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

Note: these are delegated permissions. A personal account consents for itself. On a work or
school account, whether you can consent yourself depends on the tenant's user-consent policy;
many tenants require an admin to approve an app that asks for `Mail.Send` or `Mail.ReadWrite`.

## Step 5: Create Config Files

> **Multi-account note:** This guide configures the `default` account. Credentials live
> under `~/.dbhq/outlook/<account>/`. The flat `~/.dbhq/outlook/*.json` files below are auto-migrated
> into `~/.dbhq/outlook/default/` the first time any script runs, so you can write them flat here.
> To set up an additional mailbox, prefer `outlook-setup.sh --account <name>`, which reuses
> this app registration.

Create the config directory:

```bash
mkdir -p ~/.dbhq/outlook
chmod 700 ~/.dbhq ~/.dbhq/outlook
```

Create `~/.dbhq/outlook/config.json`:

```json
{
    "client_id": "YOUR_APPLICATION_ID",
    "tenant": "common",
    "redirect_uri": "https://login.microsoftonline.com/common/oauth2/nativeclient",
    "scope": "offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read"
}
```

Set permissions:

```bash
chmod 600 ~/.dbhq/outlook/config.json
```

## Step 6: Get Authorization Code

Sign-in uses PKCE: make a random verifier, and send only its SHA-256 to the browser.

```bash
VERIFIER=$(LC_ALL=C tr -dc 'A-Za-z0-9._~-' < /dev/urandom | head -c 64)
CHALLENGE=$(printf '%s' "$VERIFIER" | openssl dgst -sha256 -binary | openssl base64 -A | tr '+/' '-_' | tr -d '=')
echo "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=YOUR_CLIENT_ID&response_type=code&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient&scope=offline_access%20Mail.ReadWrite%20Mail.Send%20Calendars.ReadWrite%20User.Read&code_challenge=$CHALLENGE&code_challenge_method=S256"
```

1. Open that URL in your browser, in the same shell session so `$VERIFIER` is kept
2. Sign in with your M365 account
3. Accept the permissions
4. You'll be redirected to a blank page
5. Copy the **entire URL** from your browser's address bar
6. Extract the `code` parameter value (everything between `code=` and `&`)

## Step 7: Exchange Code for Tokens

Run this in the same shell (replace placeholders). There is no client secret: the verifier
proves this is the machine that started the sign-in.

```bash
curl -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "client_id=YOUR_CLIENT_ID" \
  -d "code=YOUR_AUTHORIZATION_CODE" \
  --data-urlencode "code_verifier=$VERIFIER" \
  --data-urlencode "redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient" \
  -d "grant_type=authorization_code" \
  --data-urlencode "scope=offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read" \
  > ~/.dbhq/outlook/credentials.json

chmod 600 ~/.dbhq/outlook/credentials.json
```

## Step 8: Verify Setup

Test the connection:

```bash
${CLAUDE_SKILL_DIR}/scripts/outlook-token.sh test
```

You should see:
```
Connection successful!
Inbox: X total, Y unread
```

## Troubleshooting

### "AADSTS7000218: The request body must contain ... client_assertion or client_secret"
- The app is still a confidential (Web) client. Move the redirect URI to the
  **Mobile and desktop applications** platform and set **Allow public client flows** to Yes

### "Invalid client secret" (an install from before PKCE)
- An older `config.json` carries a `client_secret`, which expires. Run
  `outlook-setup.sh --account <name>` again: it converts the app to a public client and drops
  the secret

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
| `~/.dbhq/outlook/config.json` | Azure app client ID and sign-in settings |
| `~/.dbhq/outlook/credentials.json` | OAuth tokens |
| `${CLAUDE_SKILL_DIR}/` | Skill and scripts |
