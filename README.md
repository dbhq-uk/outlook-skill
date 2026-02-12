# Outlook Email & Calendar Integration

Access Microsoft 365 Outlook email and calendar from Claude Code via Microsoft Graph API.

## Features

**Email:**
- List inbox, unread, filtered by sender
- Search emails
- Read full message content
- Create drafts in Outlook (plain text or markdown-formatted)
- Send emails (with confirmation)
- Download and attach files (up to 150MB)
- Archive, delete, mark read/unread

**Calendar:**
- View today's events, week view, upcoming
- Check availability (free/busy)
- Create events (with confirmation)
- Update and delete events

## Quick Start

### 1. Install Dependencies

```bash
# macOS
brew install azure-cli jq curl pandoc

# Ubuntu/Debian
sudo apt install azure-cli jq curl pandoc

# Or install Azure CLI: https://docs.microsoft.com/en-us/cli/azure/install-azure-cli
```

**Note:** `pandoc` is optional - only needed for markdown-formatted emails (`mddraft` command).

### 2. Install the Skill

From the repo root:

```bash
./install.sh outlook
```

Or manually:

```bash
mkdir -p ~/.claude/skills/outlook
cp -r outlook/* ~/.claude/skills/outlook/
chmod +x ~/.claude/skills/outlook/scripts/*.sh
```

### 3. Run Setup

```bash
~/.claude/skills/outlook/scripts/outlook-setup.sh
```

This will:
1. Log you into Azure
2. Create an app registration with required permissions
3. Open browser for M365 consent
4. Store credentials securely in `~/.outlook/`
5. Test the connection

### 4. Verify

```bash
~/.claude/skills/outlook/scripts/outlook-token.sh test
```

Should show:
```
Connection successful!
Inbox: X total, Y unread
```

## Usage

### Email Commands

```bash
# List inbox
~/.claude/skills/outlook/scripts/outlook-mail.sh inbox [count]

# Unread only
~/.claude/skills/outlook/scripts/outlook-mail.sh unread [count]

# Filter by sender
~/.claude/skills/outlook/scripts/outlook-mail.sh from "john@example.com" [count]

# Search
~/.claude/skills/outlook/scripts/outlook-mail.sh search "project update" [count]

# Read full message
~/.claude/skills/outlook/scripts/outlook-mail.sh read <message-id>

# Create plain text draft
~/.claude/skills/outlook/scripts/outlook-mail.sh draft "to@example.com" "Subject" "Body"

# Create markdown-formatted draft (converts to HTML via pandoc)
~/.claude/skills/outlook/scripts/outlook-mail.sh mddraft "to@example.com" "Subject" "**Bold** and _italic_ text"

# Reply (creates draft)
~/.claude/skills/outlook/scripts/outlook-mail.sh reply <message-id> "Reply body"

# Send draft
~/.claude/skills/outlook/scripts/outlook-mail.sh send <draft-id>

# List drafts
~/.claude/skills/outlook/scripts/outlook-mail.sh drafts

# Management
~/.claude/skills/outlook/scripts/outlook-mail.sh markread <id>
~/.claude/skills/outlook/scripts/outlook-mail.sh markunread <id>
~/.claude/skills/outlook/scripts/outlook-mail.sh archive <id>
~/.claude/skills/outlook/scripts/outlook-mail.sh delete <id>

# Folders & stats
~/.claude/skills/outlook/scripts/outlook-mail.sh folders
~/.claude/skills/outlook/scripts/outlook-mail.sh stats
```

### Calendar Commands

```bash
# View events
~/.claude/skills/outlook/scripts/outlook-calendar.sh events [count]
~/.claude/skills/outlook/scripts/outlook-calendar.sh today
~/.claude/skills/outlook/scripts/outlook-calendar.sh week

# Read event details
~/.claude/skills/outlook/scripts/outlook-calendar.sh read <event-id>

# List calendars
~/.claude/skills/outlook/scripts/outlook-calendar.sh calendars

# Create event (times in YYYY-MM-DDTHH:MM format)
~/.claude/skills/outlook/scripts/outlook-calendar.sh create "Subject" "2025-02-05T14:00" "2025-02-05T15:00" "Location"

# Quick 1-hour event
~/.claude/skills/outlook/scripts/outlook-calendar.sh quick "Meeting" "2025-02-05T14:00"

# Check availability
~/.claude/skills/outlook/scripts/outlook-calendar.sh free "2025-02-05T09:00" "2025-02-05T17:00"

# Update event
~/.claude/skills/outlook/scripts/outlook-calendar.sh update <id> subject "New Subject"
~/.claude/skills/outlook/scripts/outlook-calendar.sh update <id> location "New Location"
~/.claude/skills/outlook/scripts/outlook-calendar.sh update <id> start "2025-02-05T15:00"

# Delete event
~/.claude/skills/outlook/scripts/outlook-calendar.sh delete <event-id>
```

### Token Management

```bash
# Test connection
~/.claude/skills/outlook/scripts/outlook-token.sh test

# Refresh expired token
~/.claude/skills/outlook/scripts/outlook-token.sh refresh

# Check status
~/.claude/skills/outlook/scripts/outlook-token.sh status

# Get raw token (for debugging)
~/.claude/skills/outlook/scripts/outlook-token.sh get
```

## Natural Language (via Claude)

Once installed, you can use natural language:

| You say | What happens |
|---------|--------------|
| "check email" | Lists recent inbox |
| "show unread" | Lists unread emails |
| "emails from John" | Filters by sender |
| "search email for invoice" | Searches content |
| "capture email 3" | Saves to brain's inbox/ |
| "draft email to X about Y" | Creates Outlook draft |
| "draft formatted email to X" | Creates markdown-formatted draft |
| "reply to that saying Z" | Creates reply draft |
| "send it" | Sends after confirmation |
| "what's on today" | Shows today's calendar |
| "my week" | Shows week view |
| "am I free Thursday 2pm" | Checks availability |
| "schedule call with X Tuesday 3pm" | Creates event (with confirm) |

## File Structure

```
~/.claude/skills/outlook/       # Skill (available everywhere)
├── SKILL.md                    # Skill definition
├── scripts/
│   ├── outlook-setup.sh        # One-time setup
│   ├── outlook-mail.sh         # Email operations
│   ├── outlook-calendar.sh     # Calendar operations
│   └── outlook-token.sh        # Token management
└── references/
    └── setup.md                # Manual setup guide

~/.outlook/                 # Credentials (created by setup)
├── config.json                 # Azure app ID + secret
└── credentials.json            # OAuth tokens
```

## Permissions

The Azure app requires these Microsoft Graph delegated permissions:

| Permission | Purpose |
|------------|---------|
| Mail.ReadWrite | Read inbox, manage drafts |
| Mail.Send | Send emails |
| Calendars.ReadWrite | Read/create events |
| offline_access | Refresh tokens without re-login |
| User.Read | Basic profile info |

## Token Expiry

- **Access tokens** expire after ~1 hour but are **automatically refreshed**
- **Refresh tokens** last ~90 days with activity
- No manual intervention needed for token refresh
- If refresh token expires, re-run setup

## Troubleshooting

### "Refresh token expired" / "Invalid grant"
Re-run setup:
```bash
~/.claude/skills/outlook/scripts/outlook-setup.sh
```

### "Insufficient privileges"
1. Go to Azure Portal → App registrations → Your app → API permissions
2. Verify all permissions are listed
3. Remove and re-add if needed
4. Re-run setup to re-consent

### "Message not found"
- Message IDs are the last 20 characters of full IDs
- IDs from `inbox` command work with `read`, `reply`, etc.
- Scripts search recent 100 messages when looking up short IDs

### Connection issues
```bash
# Check status
~/.claude/skills/outlook/scripts/outlook-token.sh status

# Test connection
~/.claude/skills/outlook/scripts/outlook-token.sh test
```

## Manual Setup

If automated setup fails, see `references/setup.md` for step-by-step Azure Portal instructions.

## Security Notes

- Credentials stored with 600 permissions (owner read/write only)
- Client secret valid for 2 years (configurable in Azure)
- Never commit `~/.outlook/` to version control
- Revoke app access: Azure Portal → Enterprise applications → Your app → Delete

## Uninstall

```bash
# Remove skill
rm -rf ~/.claude/skills/outlook

# Remove credentials
rm -rf ~/.outlook-mcp

# Optionally delete Azure app registration
az ad app delete --id <your-app-id>
```
