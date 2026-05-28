---
name: outlook
description: Use for email and calendar operations - checking inbox, sending emails, viewing calendar, scheduling events. Trigger on phrases like "check email", "draft email", "my calendar", "schedule", "am I free".
---

# Outlook Email & Calendar

Access Microsoft 365 Outlook email and calendar via Microsoft Graph API.

## CRITICAL: Replies preserve ALL original recipients (reply-all by default)

**`reply`, `mdreply`, and `followup` use Microsoft Graph's `createReplyAll` endpoint. The new draft includes every `To:` and `Cc:` recipient from the original message — not just the sender.**

Mandatory rules:

1. **Always read the original message's full `To:` and `Cc:` lists BEFORE creating a reply.** Use a direct API call if `read` truncates: `curl … "/me/messages/<id>?$select=toRecipients,ccRecipients"`. Knowing who's on the thread is part of "reading the full body end-to-end" — do not skip it.
2. **After creating any reply draft, confirm the displayed `To:`, `Cc:`, and `Bcc:` lines match what you intended.** All three reply commands now print every recipient (not just `To[0]`). If the list looks short, the original might have had CCs you missed — re-check before sending.
3. **If you genuinely want sender-only**, create the reply, then run `update to <sender-email>` and `update cc ""` (or manually edit) to trim recipients. Default is "everyone stays in the loop".
4. **Never assume a single-recipient `To:` means a single-recipient thread.** Estate agents, solicitors, accountants, and courts routinely Cc colleagues, assistants, and audit addresses. Dropping those CCs on reply is a real harm — they stop seeing the conversation.

This rule exists because a previous reply silently dropped two CCs (assistant addresses on an estate-agent thread); the recipients had to be looped back in via a follow-up email. Reply-all is now the default to make recipient loss impossible by accident.

## CRITICAL: Reading email content

**`preview` is a snippet (first ~200 chars of body), not the full message. NEVER use `preview` to analyse, summarise, respond to, or report on email content. A short preview does NOT mean a short message — the message can continue for many paragraphs and contain attachments, requests, deadlines, or substantive content not visible in the preview.**

Mandatory rules:

1. **Use `read <message-id>` for any substantive engagement** with an email — analysis, summary, reply, decision-making, documentation. Always.
2. **Use `preview` only for navigation** — finding the right message ID from a list, confirming a subject line, checking date/sender. Never for content.
3. **If the `read` output gets truncated** by terminal/tool limits, extract the body via grep or jq with a wide enough regex to capture the whole message. Do not stop at the first match.
4. **Before replying to or reporting on a message, confirm internally**: "Have I read the full body end-to-end, including any attachment list and request lines?"
5. **For long reply chains**: the `read` output includes the quoted prior chain. Identify the body of the *current* message (between the headers `---` separator and the start of the quoted chain) and ensure that body is fully captured before doing anything else.

This rule exists because trusting previews has led to missing critical content (attachments, action requests, deadlines, off-chain coordination signals). It is non-negotiable.

## Time and date awareness

Email correspondence routinely uses relative times — "today", "yesterday", "by tomorrow", "this morning", "by EOD Tuesday". Each of those is ambiguous without an anchor.

Mandatory rules:

1. **At the start of any email work, run `date` via Bash** to confirm the current date/time. Never assume the date from earlier in the conversation — the conversation may span days, and the system date can roll over.
2. **When computing deadlines or "ago" references**, anchor against actual `date` output, not memory. Example: an email timestamped `2026-05-06T16:10:58Z` is Wednesday 6 May at 17:10 BST (UTC+1 in summer), not yesterday.
3. **When an email is about to be sent**, confirm the date in the planned send is correct (the date you're embedding in the body must match the date the email will actually arrive).
4. **Track timezone explicitly** (BST vs UTC). UK summer time = UTC+1. Microsoft Graph timestamps are UTC.

If unsure of the date or time, run `date` and `date -u` (UTC) before responding.

## Email font and formatting preferences

The skill applies these inline styles to every markdown-converted email body, on every command (`mddraft`, `mdreply`, `followup`, `update mdbody`):

| Property | Value | Why inline |
|---|---|---|
| **Font family** | `'Aptos', 'Aptos Display', 'Segoe UI', Roboto, sans-serif` | Aptos is the Microsoft 365 default since 2024. Falls back to Segoe UI on older Outlook, Roboto / system sans on non-Microsoft clients. Inline `style=""` survives Outlook's `<style>`-block stripping. |
| **Font size** | `14px` | Readable, professional |
| **Line height** | `1.5` (mddraft / update mdbody) or `1.6` (mdreply / followup) | Comfortable spacing |
| **Colour** | `#333` | Soft black; avoids harsh `#000` |
| **Paragraph margin** | `0 0 14px 0` (inline on every `<p>` tag) | Outlook ignores `<p>` margins from `<style>` blocks but respects inline. Without this, paragraphs collapse together until Outlook re-renders the draft after an edit. |

These are set in `scripts/outlook-mail.sh` — search for `font-family` and `<p style=` to locate the four code paths.

To change font preferences globally, edit those four locations in the script.

## Prerequisites

- Credentials configured in `~/.outlook/` (run setup if not done)
- Azure CLI, jq, curl installed

**Note:** Tokens are automatically refreshed when needed. No manual intervention required.

## Email Operations

### Reading Email

```bash
# List inbox (default 10 messages)
~/.claude/skills/outlook/scripts/outlook-mail.sh inbox

# List more messages
~/.claude/skills/outlook/scripts/outlook-mail.sh inbox 25

# Unread only
~/.claude/skills/outlook/scripts/outlook-mail.sh unread

# Focused inbox only
~/.claude/skills/outlook/scripts/outlook-mail.sh focused

# List sent items (your sent emails)
~/.claude/skills/outlook/scripts/outlook-mail.sh sent
~/.claude/skills/outlook/scripts/outlook-mail.sh sent 25

# List messages from any folder by name (searches recursively)
~/.claude/skills/outlook/scripts/outlook-mail.sh folder "Chawton Hector" 20

# Filter by sender
~/.claude/skills/outlook/scripts/outlook-mail.sh from "john@example.com"

# Search emails
~/.claude/skills/outlook/scripts/outlook-mail.sh search "project update"

# Read full message (use ID from list)
~/.claude/skills/outlook/scripts/outlook-mail.sh read <message-id>

# Quick preview (subject, from, date, body preview)
~/.claude/skills/outlook/scripts/outlook-mail.sh preview <message-id>
```

### Sending Email

```bash
# Create plain text draft
~/.claude/skills/outlook/scripts/outlook-mail.sh draft "recipient@example.com" "Subject" "Body text"

# Create markdown-formatted draft (converts to HTML)
~/.claude/skills/outlook/scripts/outlook-mail.sh mddraft "recipient@example.com" "Subject" "**Bold** and _italic_ text"

# Send a draft (use draft ID)
~/.claude/skills/outlook/scripts/outlook-mail.sh send <draft-id>

# Reply to a message (plain text - creates draft, REPLY-ALL: includes original To: + Cc:)
~/.claude/skills/outlook/scripts/outlook-mail.sh reply <message-id> "Reply body"

# Reply with markdown formatting (converts to HTML - creates draft, REPLY-ALL: includes original To: + Cc:)
~/.claude/skills/outlook/scripts/outlook-mail.sh mdreply <message-id> "**Bold** reply with _formatting_"

# (For sender-only reply, create the draft then trim recipients via `update to`/`update cc`.)

# Send reply draft
~/.claude/skills/outlook/scripts/outlook-mail.sh send <reply-draft-id>

# Follow up on your own sent email (chaser)
~/.claude/skills/outlook/scripts/outlook-mail.sh followup <sent-message-id>
~/.claude/skills/outlook/scripts/outlook-mail.sh followup <sent-message-id> "Custom follow-up body in **markdown**"

# Update an existing draft
~/.claude/skills/outlook/scripts/outlook-mail.sh update <draft-id> subject "New subject line"
~/.claude/skills/outlook/scripts/outlook-mail.sh update <draft-id> body "Plain text body"
~/.claude/skills/outlook/scripts/outlook-mail.sh update <draft-id> mdbody "**Markdown** body"
~/.claude/skills/outlook/scripts/outlook-mail.sh update <draft-id> to "new-recipient@example.com"
~/.claude/skills/outlook/scripts/outlook-mail.sh update <draft-id> cc "cc@example.com"
~/.claude/skills/outlook/scripts/outlook-mail.sh update <draft-id> bcc "bcc@example.com"

# List drafts
~/.claude/skills/outlook/scripts/outlook-mail.sh drafts
```

**Note:** `mddraft`, `mdreply`, and `update mdbody` require `pandoc` for markdown conversion. Install with `brew install pandoc` (macOS) or `apt install pandoc` (Linux).

**IMPORTANT:** Always prefer `mdreply` over `reply` for professional emails - plain text replies look poorly formatted in Outlook.

**Reply-chain preservation:** `update mdbody` automatically preserves the quoted reply chain on drafts created via `mdreply` or `followup` (an invisible `<span data-mdreply-chain-start="1">` marker is injected when the reply is created, and `update mdbody` splits on it). The plain `update body` command does NOT preserve the chain - if you need to edit a reply draft body, use `update mdbody`.

### Attachments

**Reading attachments:**
```bash
# List attachments on a message
~/.claude/skills/outlook/scripts/outlook-mail.sh attachments <message-id>

# Download ALL attachments to ./inbox/
~/.claude/skills/outlook/scripts/outlook-mail.sh download <message-id>

# Download specific attachment
~/.claude/skills/outlook/scripts/outlook-mail.sh download <message-id> <attachment-id>
```

**Adding attachments to drafts:**
```bash
# Add attachment to a draft (supports files up to 150MB)
~/.claude/skills/outlook/scripts/outlook-mail.sh attach <draft-id> <file-path>
```

Upload method is automatic based on file size:
- **Small files (< 3MB):** Direct base64 upload - instant
- **Large files (3MB - 150MB):** Chunked upload with progress indicator

Multiple attachments can be added by calling `attach` multiple times on the same draft.

### Email Management

```bash
# Mark as read
~/.claude/skills/outlook/scripts/outlook-mail.sh markread <message-id>

# Mark as unread
~/.claude/skills/outlook/scripts/outlook-mail.sh markunread <message-id>

# Delete
~/.claude/skills/outlook/scripts/outlook-mail.sh delete <message-id>

# Archive
~/.claude/skills/outlook/scripts/outlook-mail.sh archive <message-id>

# Move to any folder (searches by name, supports nested folders)
~/.claude/skills/outlook/scripts/outlook-mail.sh move <message-id> "Projects"
~/.claude/skills/outlook/scripts/outlook-mail.sh move <message-id> "Clients/Acme"
```

### Folder Management

```bash
# List top-level folders
~/.claude/skills/outlook/scripts/outlook-mail.sh folders

# List subfolders of a folder (default: inbox)
~/.claude/skills/outlook/scripts/outlook-mail.sh subfolders
~/.claude/skills/outlook/scripts/outlook-mail.sh subfolders "Important"

# Create a new top-level folder
~/.claude/skills/outlook/scripts/outlook-mail.sh mkdir "Projects"

# Create a subfolder under an existing folder
~/.claude/skills/outlook/scripts/outlook-mail.sh mkdir "Acme" "Clients"
~/.claude/skills/outlook/scripts/outlook-mail.sh mkdir "Urgent" inbox

# Inbox statistics (total, unread counts)
~/.claude/skills/outlook/scripts/outlook-mail.sh stats
```

## Calendar Operations

### Viewing Calendar

```bash
# Upcoming events (default 10)
~/.claude/skills/outlook/scripts/outlook-calendar.sh events

# Today's events
~/.claude/skills/outlook/scripts/outlook-calendar.sh today

# This week
~/.claude/skills/outlook/scripts/outlook-calendar.sh week

# Read event details
~/.claude/skills/outlook/scripts/outlook-calendar.sh read <event-id>

# List calendars
~/.claude/skills/outlook/scripts/outlook-calendar.sh calendars
```

### Creating Events

```bash
# Create event (dates in YYYY-MM-DDTHH:MM format)
~/.claude/skills/outlook/scripts/outlook-calendar.sh create "Meeting subject" "2025-02-05T14:00" "2025-02-05T15:00" "Conference Room A"

# Quick 1-hour event
~/.claude/skills/outlook/scripts/outlook-calendar.sh quick "Team standup" "2025-02-05T09:00"
```

### Availability

```bash
# Check free/busy
~/.claude/skills/outlook/scripts/outlook-calendar.sh free "2025-02-05T09:00" "2025-02-05T17:00"
```

## Workflow: Capturing Email to Brain

When user wants to capture an email:

1. List emails to find the one to capture
2. Read the full message content
3. Check for attachments with `attachments` command
4. Download any attachments (goes to `./inbox/`)
5. Create markdown file in brain's `inbox/` directory:

```markdown
# Email: [Subject]

**From:** sender@example.com
**Date:** YYYY-MM-DD HH:MM
**Captured:** YYYY-MM-DD

## Content
[Email body]

## Attachments
- [[inbox/filename.pdf]] (captured)

## Notes
[User's annotations]
```

## Workflow: Processing Email Attachments

When user wants to grab attachments from an email:

1. Find the email: `inbox`, `search`, or `from` commands
2. List attachments: `attachments <message-id>`
3. Download: `download <message-id>` (all) or `download <message-id> <attachment-id>` (specific)
4. Files land in `./inbox/` for processing
5. User allocates files to appropriate areas during review

## Workflow: Sending Email

Always draft first, confirm, then send:

1. Create draft with `draft` or `mddraft` command
2. Show user the draft content
3. Wait for "send it" or change requests
4. Update draft if needed
5. Send with `send` command only after explicit approval

## Workflow: Sending Email with Attachments

1. Create draft with `draft` or `mddraft` command
2. Add attachments with `attach <draft-id> <file-path>` (repeat for multiple files)
3. Show user the draft details and attached files
4. Wait for confirmation
5. Send with `send` command only after explicit approval

**Example:**
```bash
# Create draft
~/.claude/skills/outlook/scripts/outlook-mail.sh draft "bob@example.com" "Q4 Report" "Please find the report attached."
# Output: Draft ID: xxxxxxxxxxxxxxxxxxxx

# Attach files (can be called multiple times)
~/.claude/skills/outlook/scripts/outlook-mail.sh attach xxxxxxxxxxxxxxxxxxxx /path/to/report.pdf
~/.claude/skills/outlook/scripts/outlook-mail.sh attach xxxxxxxxxxxxxxxxxxxx /path/to/data.xlsx

# Send after user confirms
~/.claude/skills/outlook/scripts/outlook-mail.sh send xxxxxxxxxxxxxxxxxxxx
```

## Workflow: Sending Follow-up / Chaser Emails

When user wants to follow up on an email they sent:

1. List sent items with `sent` command to find the original email
2. Create follow-up with `followup <sent-id>` (uses default message) or provide custom body
3. Show user the draft content
4. Wait for confirmation or changes
5. Send with `send` command only after explicit approval

**Example:**
```bash
# Find the original sent email
~/.claude/skills/outlook/scripts/outlook-mail.sh sent 20

# Create follow-up draft (default body)
~/.claude/skills/outlook/scripts/outlook-mail.sh followup abc123xyz

# Or with custom message
~/.claude/skills/outlook/scripts/outlook-mail.sh followup abc123xyz "Hi, just checking in on this. Would be great to get your thoughts when you have a moment."

# Send after user confirms
~/.claude/skills/outlook/scripts/outlook-mail.sh send <draft-id>
```

## Workflow: Creating Calendar Events

Always confirm before creating:

1. Parse user's request for: subject, start time, end time, location
2. Show proposed event details to user
3. Wait for confirmation or adjustments
4. Create event only after explicit "yes" / approval

## Error Handling

- **Token expired**: Automatically refreshed on next call
- **Permission denied**: Re-run setup to re-consent
- **Network error**: Check connectivity, retry

## Setup

If not configured, run:
```bash
~/.claude/skills/outlook/scripts/outlook-setup.sh
```

See `references/setup.md` for manual setup instructions.
