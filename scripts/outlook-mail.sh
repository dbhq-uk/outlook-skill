#!/bin/bash
# Outlook Mail Operations via Microsoft Graph API

set -e

CONFIG_DIR="$HOME/.outlook"
CONFIG_FILE="$CONFIG_DIR/config.json"
CREDS_FILE="$CONFIG_DIR/credentials.json"
GRAPH_URL="https://graph.microsoft.com/v1.0"

# Check credentials
if [ ! -f "$CREDS_FILE" ]; then
    echo "Error: Credentials not found. Run outlook-setup.sh first."
    exit 1
fi

# Auto-refresh token if needed
ensure_valid_token() {
    local access_token
    access_token=$(jq -r '.access_token' "$CREDS_FILE")

    # Quick test call
    local test_response
    test_response=$(curl -s -X GET "https://graph.microsoft.com/v1.0/me" \
        -H "Authorization: Bearer $access_token")

    if echo "$test_response" | jq -e '.error' > /dev/null 2>&1; then
        # Token invalid, try refresh
        local refresh_token client_id client_secret
        refresh_token=$(jq -r '.refresh_token' "$CREDS_FILE")
        client_id=$(jq -r '.client_id' "$CONFIG_FILE")
        client_secret=$(jq -r '.client_secret' "$CONFIG_FILE")

        if [ -z "$refresh_token" ] || [ "$refresh_token" = "null" ]; then
            echo "Error: No refresh token. Run outlook-setup.sh to re-authenticate."
            exit 1
        fi

        local scope="offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read"
        local refresh_response
        refresh_response=$(curl -s -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
            -H "Content-Type: application/x-www-form-urlencoded" \
            -d "client_id=$client_id" \
            -d "client_secret=$client_secret" \
            -d "refresh_token=$refresh_token" \
            -d "grant_type=refresh_token" \
            -d "scope=$scope")

        if echo "$refresh_response" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error refreshing token: $(echo "$refresh_response" | jq -r '.error_description')"
            exit 1
        fi

        echo "$refresh_response" > "$CREDS_FILE"
        chmod 600 "$CREDS_FILE"
        access_token=$(jq -r '.access_token' "$CREDS_FILE")
    fi

    echo "$access_token"
}

ACCESS_TOKEN=$(ensure_valid_token)

if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
    echo "Error: Invalid access token."
    exit 1
fi

# API call helper
api_call() {
    local method="$1"
    local endpoint="$2"
    local data="$3"

    if [ -n "$data" ]; then
        curl -s -X "$method" "${GRAPH_URL}${endpoint}" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -d "$data"
    else
        # Content-Length: 0 required for POST requests with no body
        curl -s -X "$method" "${GRAPH_URL}${endpoint}" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Length: 0"
    fi
}

# Format message for display
format_message() {
    jq -r '
        def short_id: .[-20:];
        "[\(.id | short_id)] \(.receivedDateTime | split("T")[0]) | \(.from.emailAddress.address // "unknown") | \(.subject // "(no subject)") | \(if .isRead then "read" else "UNREAD" end)"
    '
}

# Format message list
format_messages() {
    jq -r '
        if .error then
            "Error: \(.error.message // .error.code // "Unknown API error")"
        elif (.value | length) == 0 then
            "No messages found."
        else
            def short_id: .[-20:];
            .value | to_entries | .[] |
            "[\(.key + 1)] \(.value.id | short_id) | \(.value.receivedDateTime | split("T")[0]) | \(.value.from.emailAddress.address // "unknown") | \(.value.subject // "(no subject)")"
        end
    '
}

# Resolve short message ID to full ID
# Usage: full_id=$(resolve_message_id "short_id_or_full_id" "messages|drafts|sentitems")
resolve_message_id() {
    local msg_id="$1"
    local folder="${2:-messages}"  # "messages" for all mail, "drafts" for drafts folder, "sentitems" for sent

    # If it looks like a full ID (very long), return as-is
    if [ ${#msg_id} -gt 100 ]; then
        echo "$msg_id"
        return 0
    fi

    # Short ID - need to find the full ID
    local endpoint
    local search_limit=500

    if [ "$folder" = "drafts" ]; then
        endpoint="/me/mailFolders/drafts/messages?\$top=$search_limit&\$select=id"
    elif [ "$folder" = "sentitems" ]; then
        endpoint="/me/mailFolders/sentitems/messages?\$top=$search_limit&\$select=id&\$orderby=sentDateTime%20desc"
    else
        endpoint="/me/messages?\$top=$search_limit&\$select=id&\$orderby=receivedDateTime%20desc"
    fi

    local full_id
    full_id=$(api_call GET "$endpoint" | jq -r ".value[].id | select(endswith(\"$msg_id\"))" | head -1)

    if [ -z "$full_id" ]; then
        return 1
    fi

    echo "$full_id"
    return 0
}

# Commands
case "$1" in
    inbox)
        count="${2:-10}"
        echo "Fetching inbox ($count messages)..."
        api_call GET "/me/mailFolders/inbox/messages?\$top=$count&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" | format_messages
        ;;

    unread)
        count="${2:-10}"
        echo "Fetching unread messages..."
        api_call GET "/me/mailFolders/inbox/messages?\$filter=isRead%20eq%20false&\$top=$count&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" | format_messages
        ;;

    focused)
        count="${2:-10}"
        echo "Fetching focused inbox..."
        api_call GET "/me/mailFolders/inbox/messages?\$filter=inferenceClassification%20eq%20'focused'&\$top=$count&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" | format_messages
        ;;

    sent)
        count="${2:-10}"
        echo "Fetching sent items ($count messages)..."
        api_call GET "/me/mailFolders/sentitems/messages?\$top=$count&\$orderby=sentDateTime%20desc&\$select=id,subject,toRecipients,sentDateTime,bodyPreview" | jq -r '
            def short_id: .[-20:];
            .value | to_entries | .[] |
            "[\(.key + 1)] \(.value.id | short_id) | \(.value.sentDateTime | split("T")[0]) | To: \(.value.toRecipients[0].emailAddress.address // "unknown") | \(.value.subject // "(no subject)")"
        '
        ;;

    from)
        sender="$2"
        count="${3:-10}"
        if [ -z "$sender" ]; then
            echo "Usage: outlook-mail.sh from <sender-email> [count]"
            exit 1
        fi
        echo "Fetching emails from $sender..."
        api_call GET "/me/messages?\$search=%22from:$sender%22&\$top=$count&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" | format_messages
        ;;

    search)
        query="$2"
        count="${3:-10}"
        if [ -z "$query" ]; then
            echo "Usage: outlook-mail.sh search <query> [count]"
            exit 1
        fi
        # If query looks like an email address, use KQL from: search for precise matching
        if [[ "$query" =~ ^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$ ]]; then
            echo "Searching for emails from $query..."
            api_call GET "/me/messages?\$search=%22from:$query%22&\$top=$count&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" | format_messages
        else
            echo "Searching for: $query..."
            # Note: $search cannot be combined with $orderby in Microsoft Graph API
            api_call GET "/me/messages?\$search=\"$query\"&\$top=$count&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" | format_messages
        fi
        ;;

    read)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh read <message-id>"
            exit 1
        fi

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Reading message..."
        api_call GET "/me/messages/$msg_id" | jq -r '
            "Subject: \(.subject // "(no subject)")",
            "From: \(.from.emailAddress.name // "") <\(.from.emailAddress.address // "")>",
            "To: \([.toRecipients[].emailAddress | "\(.name // "") <\(.address)>"] | join(", "))",
            "Date: \(.receivedDateTime)",
            "---",
            (.body.content | gsub("<[^>]*>"; "") | gsub("&nbsp;"; " ") | gsub("\\s+"; " ") | ltrimstr(" ") | rtrimstr(" "))'
        ;;

    preview)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh preview <message-id>"
            exit 1
        fi

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        api_call GET "/me/messages/$msg_id?\$select=id,subject,from,receivedDateTime,bodyPreview" | jq -r '
            "Subject: \(.subject // "(no subject)")",
            "From: \(.from.emailAddress.address // "")",
            "Date: \(.receivedDateTime)",
            "Preview: \(.bodyPreview)"'
        ;;

    draft)
        to="$2"
        subject="$3"
        body="$4"
        if [ -z "$to" ] || [ -z "$subject" ]; then
            echo "Usage: outlook-mail.sh draft <to-email> <subject> <body>"
            exit 1
        fi

        echo "Creating draft..."
        payload=$(jq -n \
            --arg to "$to" \
            --arg subject "$subject" \
            --arg body "${body:-}" \
            '{
                subject: $subject,
                body: {
                    contentType: "Text",
                    content: $body
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: $to
                        }
                    }
                ]
            }')

        result=$(api_call POST "/me/messages" "$payload")
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating draft:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        echo "Draft created!"
        echo "Draft ID: ${draft_id: -20}"
        echo
        echo "$result" | jq -r '"To: \(.toRecipients[0].emailAddress.address)", "Subject: \(.subject)", "Body: \(.body.content)"'
        ;;

    mddraft)
        to="$2"
        subject="$3"
        body="$4"
        if [ -z "$to" ] || [ -z "$subject" ]; then
            echo "Usage: outlook-mail.sh mddraft <to-email> <subject> <markdown-body>"
            exit 1
        fi

        # Check for pandoc
        if ! command -v pandoc &> /dev/null; then
            echo "Error: pandoc is required for markdown conversion"
            echo "Install with: brew install pandoc (macOS) or apt install pandoc (Linux)"
            exit 1
        fi

        echo "Creating markdown draft..."

        # Convert markdown to HTML
        html_body=$(echo "${body:-}" | pandoc -f markdown -t html)

        # Wrap in basic email-friendly HTML structure
        html_body="<html><body style=\"font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 14px; line-height: 1.5; color: #333;\">
${html_body}
</body></html>"

        payload=$(jq -n \
            --arg to "$to" \
            --arg subject "$subject" \
            --arg body "$html_body" \
            '{
                subject: $subject,
                body: {
                    contentType: "HTML",
                    content: $body
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: $to
                        }
                    }
                ]
            }')

        result=$(api_call POST "/me/messages" "$payload")
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating draft:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        echo "Draft created (HTML from Markdown)!"
        echo "Draft ID: ${draft_id: -20}"
        echo
        echo "$result" | jq -r '"To: \(.toRecipients[0].emailAddress.address)", "Subject: \(.subject)"'
        ;;

    reply)
        msg_id="$2"
        body="$3"
        if [ -z "$msg_id" ] || [ -z "$body" ]; then
            echo "Usage: outlook-mail.sh reply <message-id> <body>"
            exit 1
        fi

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Creating reply draft..."
        payload=$(jq -n --arg body "$body" '{comment: $body}')

        result=$(api_call POST "/me/messages/$msg_id/createReply" "$payload")
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating reply:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        echo "Reply draft created!"
        echo "Draft ID: ${draft_id: -20}"
        echo
        echo "$result" | jq -r '"To: \(.toRecipients[0].emailAddress.address)", "Subject: \(.subject)"'
        ;;

    update)
        draft_id="$2"
        field="$3"
        value="$4"
        if [ -z "$draft_id" ] || [ -z "$field" ]; then
            echo "Usage: outlook-mail.sh update <draft-id> <field> <value>"
            echo ""
            echo "Fields:"
            echo "  subject <text>     Update subject line"
            echo "  body <text>        Replace body (plain text)"
            echo "  mdbody <markdown>  Replace body (markdown -> HTML)"
            echo "  to <email>         Change recipient"
            echo "  cc <email>         Add CC recipient"
            echo "  bcc <email>        Add BCC recipient"
            exit 1
        fi

        # Resolve short ID to full ID (search in drafts folder)
        if ! draft_id=$(resolve_message_id "$draft_id" "drafts"); then
            echo "Error: Draft not found with ID: $2"
            exit 1
        fi

        case "$field" in
            subject)
                if [ -z "$value" ]; then
                    echo "Error: Subject value required"
                    exit 1
                fi
                echo "Updating subject..."
                payload=$(jq -n --arg subject "$value" '{subject: $subject}')
                ;;
            body)
                if [ -z "$value" ]; then
                    echo "Error: Body value required"
                    exit 1
                fi
                echo "Updating body (plain text)..."
                payload=$(jq -n --arg body "$value" '{body: {contentType: "Text", content: $body}}')
                ;;
            mdbody)
                if [ -z "$value" ]; then
                    echo "Error: Body value required"
                    exit 1
                fi
                # Check for pandoc
                if ! command -v pandoc &> /dev/null; then
                    echo "Error: pandoc is required for markdown conversion"
                    echo "Install with: brew install pandoc (macOS) or apt install pandoc (Linux)"
                    exit 1
                fi
                echo "Updating body (markdown -> HTML)..."
                html_body=$(echo "$value" | pandoc -f markdown -t html)
                html_body="<div style=\"font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 14px; line-height: 1.5; color: #333;\">${html_body}</div>"
                payload=$(jq -n --arg body "$html_body" '{body: {contentType: "HTML", content: $body}}')
                ;;
            to)
                if [ -z "$value" ]; then
                    echo "Error: Email address required"
                    exit 1
                fi
                echo "Updating recipient..."
                payload=$(jq -n --arg email "$value" '{toRecipients: [{emailAddress: {address: $email}}]}')
                ;;
            cc)
                if [ -z "$value" ]; then
                    echo "Error: Email address required"
                    exit 1
                fi
                echo "Adding CC recipient..."
                # Get existing CC recipients and add new one
                existing=$(api_call GET "/me/messages/$draft_id?\$select=ccRecipients" | jq '.ccRecipients // []')
                payload=$(echo "$existing" | jq --arg email "$value" '. + [{emailAddress: {address: $email}}] | {ccRecipients: .}')
                ;;
            bcc)
                if [ -z "$value" ]; then
                    echo "Error: Email address required"
                    exit 1
                fi
                echo "Adding BCC recipient..."
                # Get existing BCC recipients and add new one
                existing=$(api_call GET "/me/messages/$draft_id?\$select=bccRecipients" | jq '.bccRecipients // []')
                payload=$(echo "$existing" | jq --arg email "$value" '. + [{emailAddress: {address: $email}}] | {bccRecipients: .}')
                ;;
            *)
                echo "Error: Unknown field '$field'"
                echo "Valid fields: subject, body, mdbody, to, cc, bcc"
                exit 1
                ;;
        esac

        result=$(api_call PATCH "/me/messages/$draft_id" "$payload")

        if echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error updating draft:"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi

        echo "Draft updated!"
        echo "$result" | jq -r '"To: \(.toRecipients[0].emailAddress.address // "none")", "Subject: \(.subject // "(no subject)")"'
        ;;

    mdreply)
        msg_id="$2"
        body="$3"
        if [ -z "$msg_id" ] || [ -z "$body" ]; then
            echo "Usage: outlook-mail.sh mdreply <message-id> <markdown-body>"
            exit 1
        fi

        # Check for pandoc
        if ! command -v pandoc &> /dev/null; then
            echo "Error: pandoc is required for markdown conversion"
            echo "Install with: brew install pandoc (macOS) or apt install pandoc (Linux)"
            exit 1
        fi

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Creating reply draft with markdown formatting..."

        # Step 1: Create reply draft (empty comment to get thread headers)
        result=$(api_call POST "/me/messages/$msg_id/createReply" '{}')
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating reply:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        # Step 2: Get the existing body (contains the quoted thread)
        existing_body=$(echo "$result" | jq -r '.body.content // ""')

        # Step 3: Convert markdown to HTML
        html_body=$(echo "$body" | pandoc -f markdown -t html)

        # Wrap reply in styled div and prepend to existing thread
        combined_body="<div style=\"font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 14px; line-height: 1.6; color: #333;\">
${html_body}
</div>
<br/>
${existing_body}"

        # Step 4: PATCH the draft to update body with combined HTML
        patch_payload=$(jq -n \
            --arg body "$combined_body" \
            '{
                body: {
                    contentType: "HTML",
                    content: $body
                }
            }')

        patch_result=$(api_call PATCH "/me/messages/$draft_id" "$patch_payload")

        if echo "$patch_result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error updating draft body:"
            echo "$patch_result" | jq -r '.error.message'
            exit 1
        fi

        echo "Reply draft created (HTML from Markdown)!"
        echo "Draft ID: ${draft_id: -20}"
        echo
        echo "$patch_result" | jq -r '"To: \(.toRecipients[0].emailAddress.address)", "Subject: \(.subject)"'
        ;;

    followup)
        msg_id="$2"
        body="$3"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh followup <sent-message-id> [markdown-body]"
            echo "       Creates a follow-up reply to your own sent email (chaser)"
            echo "       Body defaults to a standard follow-up message if not provided"
            exit 1
        fi

        # Check for pandoc
        if ! command -v pandoc &> /dev/null; then
            echo "Error: pandoc is required for markdown conversion"
            echo "Install with: brew install pandoc (macOS) or apt install pandoc (Linux)"
            exit 1
        fi

        # Resolve short ID from sent items folder
        if ! msg_id=$(resolve_message_id "$msg_id" "sentitems"); then
            echo "Error: Sent message not found with ID: $2"
            exit 1
        fi

        # Default follow-up body if not provided
        if [ -z "$body" ]; then
            body="Hi,

Just following up on my email below.

Please let me know if you have any questions or need any additional information."
        fi

        echo "Creating follow-up draft for sent message..."

        # Step 1: Create reply draft using replyAll to include all original recipients
        result=$(api_call POST "/me/messages/$msg_id/createReply" '{}')
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating follow-up:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        # Step 2: Get the existing body (contains the quoted thread)
        existing_body=$(echo "$result" | jq -r '.body.content // ""')

        # Step 3: Convert markdown to HTML
        html_body=$(echo "$body" | pandoc -f markdown -t html)

        # Wrap reply in styled div and prepend to existing thread
        combined_body="<div style=\"font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 14px; line-height: 1.6; color: #333;\">
${html_body}
</div>
<br/>
${existing_body}"

        # Step 4: PATCH the draft to update body with combined HTML
        patch_payload=$(jq -n \
            --arg body "$combined_body" \
            '{
                body: {
                    contentType: "HTML",
                    content: $body
                }
            }')

        patch_result=$(api_call PATCH "/me/messages/$draft_id" "$patch_payload")

        if echo "$patch_result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error updating draft body:"
            echo "$patch_result" | jq -r '.error.message'
            exit 1
        fi

        echo "Follow-up draft created!"
        echo "Draft ID: ${draft_id: -20}"
        echo
        echo "$patch_result" | jq -r '"To: \(.toRecipients[0].emailAddress.address)", "Subject: \(.subject)"'
        echo
        echo "Use 'outlook-mail.sh send ${draft_id: -20}' to send"
        ;;

    send)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh send <draft-id>"
            exit 1
        fi

        # Resolve short ID to full ID (search in drafts folder)
        if ! msg_id=$(resolve_message_id "$msg_id" "drafts"); then
            echo "Error: Draft not found with ID: $2"
            exit 1
        fi

        echo "Sending..."
        result=$(api_call POST "/me/messages/$msg_id/send")

        if [ -n "$result" ]; then
            echo "Error sending:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        echo "Email sent successfully!"
        ;;

    drafts)
        count="${2:-10}"
        echo "Fetching drafts..."
        api_call GET "/me/mailFolders/drafts/messages?\$top=$count&\$orderby=createdDateTime%20desc&\$select=id,subject,toRecipients,createdDateTime" | jq -r '
            if .error then
                "Error: \(.error.message // .error.code // "Unknown API error")"
            elif (.value | length) == 0 then
                "No drafts found."
            else
                def short_id: .[-20:];
                .value | to_entries | .[] |
                "[\(.key + 1)] \(.value.id | short_id) | \(.value.createdDateTime | split("T")[0]) | To: \(.value.toRecipients[0].emailAddress.address // "none") | \(.value.subject // "(no subject)")"
            end
        '
        ;;

    markread)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh markread <message-id>"
            exit 1
        fi

        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        api_call PATCH "/me/messages/$msg_id" '{"isRead": true}' > /dev/null
        echo "Marked as read"
        ;;

    markunread)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh markunread <message-id>"
            exit 1
        fi

        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        api_call PATCH "/me/messages/$msg_id" '{"isRead": false}' > /dev/null
        echo "Marked as unread"
        ;;

    delete)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh delete <message-id>"
            exit 1
        fi

        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        api_call DELETE "/me/messages/$msg_id" > /dev/null
        echo "Message deleted"
        ;;

    archive)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh archive <message-id>"
            exit 1
        fi

        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        # Get archive folder ID
        archive_id=$(api_call GET "/me/mailFolders/archive" | jq -r '.id')

        if [ -z "$archive_id" ] || [ "$archive_id" = "null" ]; then
            echo "Error: Archive folder not found"
            exit 1
        fi

        api_call POST "/me/messages/$msg_id/move" "{\"destinationId\": \"$archive_id\"}" > /dev/null
        echo "Message archived"
        ;;

    folders)
        echo "Mail folders:"
        api_call GET "/me/mailFolders?\$top=50" | jq -r '
            if .error then
                "Error: \(.error.message // .error.code // "Unknown API error")"
            elif (.value | length) == 0 then
                "No folders found."
            else
                .value[] | "[\(.displayName)] \(.totalItemCount) total, \(.unreadItemCount) unread"
            end
        '
        ;;

    subfolders)
        parent="${2:-inbox}"
        echo "Subfolders of $parent:"

        # Handle well-known folder names or folder IDs
        case "$parent" in
            inbox|drafts|sentitems|deleteditems|archive|junkemail)
                endpoint="/me/mailFolders/$parent/childFolders?\$top=100"
                ;;
            *)
                # Try to find folder by name (search all folders)
                folder_id=$(api_call GET "/me/mailFolders?\$top=100" | jq -r --arg name "$parent" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' | head -1)
                if [ -z "$folder_id" ]; then
                    # Search in inbox subfolders
                    folder_id=$(api_call GET "/me/mailFolders/inbox/childFolders?\$top=100" | jq -r --arg name "$parent" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' | head -1)
                fi
                if [ -z "$folder_id" ]; then
                    echo "Error: Folder '$parent' not found"
                    exit 1
                fi
                endpoint="/me/mailFolders/$folder_id/childFolders?\$top=100"
                ;;
        esac

        result=$(api_call GET "$endpoint")
        count=$(echo "$result" | jq -r '.value | length')

        if [ "$count" = "0" ]; then
            echo "No subfolders found"
        else
            echo "$result" | jq -r '.value[] | "[\(.displayName)] \(.totalItemCount) total, \(.unreadItemCount) unread"'
        fi
        ;;

    folder)
        folder_name="$2"
        count="${3:-10}"
        if [ -z "$folder_name" ]; then
            echo "Usage: outlook-mail.sh folder <folder-name> [count]"
            exit 1
        fi

        echo "Finding folder '$folder_name'..."

        # Search function to find folder ID by name (recursive, up to 4 levels deep)
        find_folder_id() {
            local search_name="$1"
            local folder_id=""

            # Check top-level folders first
            folder_id=$(api_call GET "/me/mailFolders?\$top=100" | jq -r --arg name "$search_name" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' | head -1)

            if [ -n "$folder_id" ]; then
                echo "$folder_id"
                return 0
            fi

            # Search in inbox subfolders (level 2)
            local inbox_children=$(api_call GET "/me/mailFolders/inbox/childFolders?\$top=100")
            folder_id=$(echo "$inbox_children" | jq -r --arg name "$search_name" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' | head -1)

            if [ -n "$folder_id" ]; then
                echo "$folder_id"
                return 0
            fi

            # Search level 3 - children of inbox subfolders
            local level2_ids=$(echo "$inbox_children" | jq -r '.value[].id')
            for parent_id in $level2_ids; do
                local level3_children=$(api_call GET "/me/mailFolders/$parent_id/childFolders?\$top=100" 2>/dev/null)
                folder_id=$(echo "$level3_children" | jq -r --arg name "$search_name" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' 2>/dev/null | head -1)
                if [ -n "$folder_id" ]; then
                    echo "$folder_id"
                    return 0
                fi

                # Search level 4 - one more level deep
                local level3_ids=$(echo "$level3_children" | jq -r '.value[].id' 2>/dev/null)
                for level3_id in $level3_ids; do
                    folder_id=$(api_call GET "/me/mailFolders/$level3_id/childFolders?\$top=100" 2>/dev/null | jq -r --arg name "$search_name" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' 2>/dev/null | head -1)
                    if [ -n "$folder_id" ]; then
                        echo "$folder_id"
                        return 0
                    fi
                done
            done

            # Also search other top-level folders (not just inbox)
            local top_folders=$(api_call GET "/me/mailFolders?\$top=50" | jq -r '.value[] | select(.displayName != "Inbox") | .id')
            for parent_id in $top_folders; do
                folder_id=$(api_call GET "/me/mailFolders/$parent_id/childFolders?\$top=100" 2>/dev/null | jq -r --arg name "$search_name" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' 2>/dev/null | head -1)
                if [ -n "$folder_id" ]; then
                    echo "$folder_id"
                    return 0
                fi
            done

            return 1
        }

        folder_id=$(find_folder_id "$folder_name")

        if [ -z "$folder_id" ]; then
            echo "Error: Folder '$folder_name' not found"
            exit 1
        fi

        echo "Fetching messages from '$folder_name' ($count messages)..."
        api_call GET "/me/mailFolders/$folder_id/messages?\$top=$count&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" | format_messages
        ;;

    stats)
        echo "Inbox statistics:"
        api_call GET "/me/mailFolders/inbox" | jq -r '
            if .error then
                "Error: \(.error.message // .error.code // "Unknown API error")"
            else
                "Total: \(.totalItemCount)", "Unread: \(.unreadItemCount)"
            end
        '
        ;;

    move)
        msg_id="$2"
        folder_name="$3"
        if [ -z "$msg_id" ] || [ -z "$folder_name" ]; then
            echo "Usage: outlook-mail.sh move <message-id> <folder-name>"
            exit 1
        fi

        # Resolve message ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Finding folder '$folder_name'..."

        # Find folder ID (reusing find_folder_id logic)
        find_folder_id() {
            local search_name="$1"
            local folder_id=""

            # Check well-known folder names first
            case "$search_name" in
                inbox|drafts|sentitems|deleteditems|archive|junkemail)
                    folder_id=$(api_call GET "/me/mailFolders/$search_name" | jq -r '.id // empty')
                    if [ -n "$folder_id" ]; then
                        echo "$folder_id"
                        return 0
                    fi
                    ;;
            esac

            # Check top-level folders
            folder_id=$(api_call GET "/me/mailFolders?\$top=100" | jq -r --arg name "$search_name" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' | head -1)
            if [ -n "$folder_id" ]; then
                echo "$folder_id"
                return 0
            fi

            # Search in inbox subfolders (level 2)
            local inbox_children=$(api_call GET "/me/mailFolders/inbox/childFolders?\$top=100")
            folder_id=$(echo "$inbox_children" | jq -r --arg name "$search_name" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' | head -1)
            if [ -n "$folder_id" ]; then
                echo "$folder_id"
                return 0
            fi

            # Search level 3 - children of inbox subfolders
            local level2_ids=$(echo "$inbox_children" | jq -r '.value[].id')
            for parent_id in $level2_ids; do
                local level3_children=$(api_call GET "/me/mailFolders/$parent_id/childFolders?\$top=100" 2>/dev/null)
                folder_id=$(echo "$level3_children" | jq -r --arg name "$search_name" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' 2>/dev/null | head -1)
                if [ -n "$folder_id" ]; then
                    echo "$folder_id"
                    return 0
                fi
            done

            return 1
        }

        dest_folder_id=$(find_folder_id "$folder_name")

        if [ -z "$dest_folder_id" ]; then
            echo "Error: Folder '$folder_name' not found"
            exit 1
        fi

        # Move the message
        result=$(api_call POST "/me/messages/$msg_id/move" "{\"destinationId\": \"$dest_folder_id\"}")

        if echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error moving message:"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi

        echo "Moved message to '$folder_name'"
        ;;

    mkdir)
        folder_name="$2"
        parent_folder="$3"
        if [ -z "$folder_name" ]; then
            echo "Usage: outlook-mail.sh mkdir <folder-name> [parent-folder]"
            echo "       Without parent-folder, creates a top-level folder"
            exit 1
        fi

        if [ -z "$parent_folder" ]; then
            # Create top-level folder
            echo "Creating top-level folder '$folder_name'..."
            payload=$(jq -n --arg name "$folder_name" '{"displayName": $name}')
            result=$(api_call POST "/me/mailFolders" "$payload")
        else
            # Find parent folder ID
            echo "Finding parent folder '$parent_folder'..."

            # Check well-known folder names
            case "$parent_folder" in
                inbox|drafts|sentitems|deleteditems|archive|junkemail)
                    parent_id=$(api_call GET "/me/mailFolders/$parent_folder" | jq -r '.id // empty')
                    ;;
                *)
                    # Search for folder by name
                    parent_id=$(api_call GET "/me/mailFolders?\$top=100" | jq -r --arg name "$parent_folder" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' | head -1)
                    if [ -z "$parent_id" ]; then
                        # Search in inbox subfolders
                        parent_id=$(api_call GET "/me/mailFolders/inbox/childFolders?\$top=100" | jq -r --arg name "$parent_folder" '.value[] | select(.displayName | ascii_downcase == ($name | ascii_downcase)) | .id' | head -1)
                    fi
                    ;;
            esac

            if [ -z "$parent_id" ]; then
                echo "Error: Parent folder '$parent_folder' not found"
                exit 1
            fi

            echo "Creating subfolder '$folder_name' under '$parent_folder'..."
            payload=$(jq -n --arg name "$folder_name" '{"displayName": $name}')
            result=$(api_call POST "/me/mailFolders/$parent_id/childFolders" "$payload")
        fi

        if echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error creating folder:"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi

        new_folder_id=$(echo "$result" | jq -r '.id')
        echo "Created folder '$folder_name'"
        echo "Folder ID: ${new_folder_id: -20}"
        ;;

    attachments)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh attachments <message-id>"
            exit 1
        fi

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Fetching attachments..."
        result=$(api_call GET "/me/messages/$msg_id/attachments?\$select=id,name,size,contentType")

        count=$(echo "$result" | jq -r '.value | length')
        if [ "$count" = "0" ]; then
            echo "No attachments on this message"
            exit 0
        fi

        echo "$result" | jq -r '
            def format_size:
                if . < 1024 then "\(.)B"
                elif . < 1048576 then "\((. / 1024 * 10 | floor) / 10)KB"
                else "\((. / 1048576 * 10 | floor) / 10)MB"
                end;
            .value | to_entries | .[] |
            "[\(.key + 1)] \(.value.id[-20:]) | \(.value.name) | \(.value.size | format_size) | \(.value.contentType)"
        '
        ;;

    download)
        msg_id="$2"
        attachment_id="$3"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-mail.sh download <message-id> [attachment-id]"
            echo "       Without attachment-id, downloads ALL attachments"
            exit 1
        fi

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        # Destination directory
        PROJECT_ROOT="${CLAUDE_PROJECT_DIR:-$(pwd)}"
        DOWNLOAD_DIR="$PROJECT_ROOT/inbox"
        mkdir -p "$DOWNLOAD_DIR"

        # Get attachments list
        attachments=$(api_call GET "/me/messages/$msg_id/attachments?\$select=id,name,size,contentType,contentBytes")

        count=$(echo "$attachments" | jq -r '.value | length')
        if [ "$count" = "0" ]; then
            echo "No attachments on this message"
            exit 0
        fi

        # Filter to specific attachment if provided
        if [ -n "$attachment_id" ]; then
            # Find full attachment ID if short
            if [ ${#attachment_id} -le 25 ]; then
                full_att_id=$(echo "$attachments" | jq -r ".value[].id | select(endswith(\"$attachment_id\"))" | head -1)
                if [ -z "$full_att_id" ]; then
                    echo "Error: Attachment not found with ID ending in: $attachment_id"
                    exit 1
                fi
                attachment_id="$full_att_id"
            fi
            attachments=$(echo "$attachments" | jq --arg id "$attachment_id" '{value: [.value[] | select(.id == $id)]}')
            count=1
        fi

        echo "Downloading $count attachment(s)..."

        # Download each attachment
        echo "$attachments" | jq -c '.value[]' | while read -r att; do
            att_id=$(echo "$att" | jq -r '.id')
            att_name=$(echo "$att" | jq -r '.name')
            att_size=$(echo "$att" | jq -r '.size')

            # Handle filename collisions
            base_name="${att_name%.*}"
            extension="${att_name##*.}"
            if [ "$base_name" = "$att_name" ]; then
                extension=""
            fi

            dest_path="$DOWNLOAD_DIR/$att_name"
            counter=1
            while [ -f "$dest_path" ]; do
                if [ -n "$extension" ] && [ "$extension" != "$base_name" ]; then
                    dest_path="$DOWNLOAD_DIR/${base_name}_${counter}.${extension}"
                else
                    dest_path="$DOWNLOAD_DIR/${att_name}_${counter}"
                fi
                counter=$((counter + 1))
            done

            # Check if contentBytes is included (small attachments)
            content_bytes=$(echo "$att" | jq -r '.contentBytes // empty')

            if [ -n "$content_bytes" ]; then
                # Decode base64 content directly
                echo "$content_bytes" | base64 -d > "$dest_path"
            else
                # Fetch raw content for larger attachments
                curl -s -X GET "${GRAPH_URL}/me/messages/$msg_id/attachments/$att_id/\$value" \
                    -H "Authorization: Bearer $ACCESS_TOKEN" \
                    -o "$dest_path"
            fi

            # Format size for display
            if [ "$att_size" -lt 1024 ]; then
                size_str="${att_size}B"
            elif [ "$att_size" -lt 1048576 ]; then
                size_str="$(echo "scale=1; $att_size / 1024" | bc)KB"
            else
                size_str="$(echo "scale=1; $att_size / 1048576" | bc)MB"
            fi

            echo "Saved: $dest_path ($size_str)"
        done
        ;;

    attach)
        draft_id="$2"
        file_path="$3"
        if [ -z "$draft_id" ] || [ -z "$file_path" ]; then
            echo "Usage: outlook-mail.sh attach <draft-id> <file-path>"
            exit 1
        fi

        # Verify file exists
        if [ ! -f "$file_path" ]; then
            echo "Error: File not found: $file_path"
            exit 1
        fi

        # Resolve short ID to full ID (search in drafts folder)
        if ! draft_id=$(resolve_message_id "$draft_id" "drafts"); then
            echo "Error: Draft not found with ID: $2"
            exit 1
        fi

        # Get file info
        file_name=$(basename "$file_path")
        file_size=$(stat -f%z "$file_path" 2>/dev/null || stat -c%s "$file_path" 2>/dev/null)

        # Detect content type
        content_type=$(file --mime-type -b "$file_path" 2>/dev/null || echo "application/octet-stream")

        # Size threshold: 3MB = 3145728 bytes
        SMALL_FILE_LIMIT=3145728

        if [ "$file_size" -lt "$SMALL_FILE_LIMIT" ]; then
            # Simple upload for small files
            echo "Attaching $file_name ($(echo "scale=1; $file_size / 1024" | bc)KB)..."

            # Base64 encode (handle Linux vs Mac)
            if base64 --help 2>&1 | grep -q GNU; then
                file_content=$(base64 -w0 "$file_path")
            else
                file_content=$(base64 -i "$file_path" | tr -d '\n')
            fi

            payload=$(jq -n \
                --arg name "$file_name" \
                --arg contentType "$content_type" \
                --arg contentBytes "$file_content" \
                '{
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": $name,
                    "contentType": $contentType,
                    "contentBytes": $contentBytes
                }')

            result=$(api_call POST "/me/messages/$draft_id/attachments" "$payload")

            if echo "$result" | jq -e '.error' > /dev/null 2>&1; then
                echo "Error attaching file:"
                echo "$result" | jq -r '.error.message'
                exit 1
            fi

            echo "Attached: $file_name to draft"
        else
            # Chunked upload for large files (3MB - 150MB)
            echo "Attaching $file_name ($(echo "scale=1; $file_size / 1048576" | bc)MB) via chunked upload..."

            # Create upload session
            session_payload=$(jq -n \
                --arg name "$file_name" \
                --argjson size "$file_size" \
                '{
                    "AttachmentItem": {
                        "attachmentType": "file",
                        "name": $name,
                        "size": $size
                    }
                }')

            session_result=$(api_call POST "/me/messages/$draft_id/attachments/createUploadSession" "$session_payload")

            upload_url=$(echo "$session_result" | jq -r '.uploadUrl // empty')
            if [ -z "$upload_url" ]; then
                echo "Error creating upload session:"
                echo "$session_result" | jq -r '.error.message // .'
                exit 1
            fi

            # Upload in 4MB chunks
            CHUNK_SIZE=4194304
            offset=0

            while [ "$offset" -lt "$file_size" ]; do
                # Calculate chunk end
                chunk_end=$((offset + CHUNK_SIZE - 1))
                if [ "$chunk_end" -ge "$file_size" ]; then
                    chunk_end=$((file_size - 1))
                fi
                chunk_length=$((chunk_end - offset + 1))

                # Progress indicator
                progress=$((offset * 100 / file_size))
                bar_filled=$((progress / 10))
                bar_empty=$((10 - bar_filled))
                printf "\rUploading: [%s%s] %d%%" "$(printf '#%.0s' $(seq 1 $bar_filled 2>/dev/null) || echo '')" "$(printf ' %.0s' $(seq 1 $bar_empty 2>/dev/null) || echo '')" "$progress"

                # Extract chunk efficiently (using large block size with byte-level positioning)
                # iflag=skip_bytes,count_bytes makes skip/count work in bytes regardless of bs
                chunk_result=$(dd if="$file_path" bs=1M iflag=skip_bytes,count_bytes skip="$offset" count="$chunk_length" 2>/dev/null | \
                curl -s -X PUT "$upload_url" \
                    -H "Content-Type: application/octet-stream" \
                    -H "Content-Length: $chunk_length" \
                    -H "Content-Range: bytes ${offset}-${chunk_end}/${file_size}" \
                    --data-binary @-)

                # Check for errors in chunk upload
                # Note: Successful uploads return empty body (HTTP 200) or JSON with nextExpectedRanges
                # Errors return JSON with .error object
                if [ -n "$chunk_result" ]; then
                    # Only check for errors if there's a response body
                    if echo "$chunk_result" | jq -e '.error' > /dev/null 2>&1; then
                        echo ""
                        echo "Error uploading chunk at offset $offset:"
                        echo "$chunk_result" | jq '.'
                        exit 1
                    fi
                fi

                offset=$((chunk_end + 1))
            done

            printf "\rUploading: [##########] 100%%\n"
            echo "Attached: $file_name to draft"
        fi
        ;;

    *)
        echo "Outlook Mail Operations"
        echo
        echo "Usage: outlook-mail.sh <command> [args]"
        echo
        echo "Reading:"
        echo "  inbox [count]              List inbox messages"
        echo "  unread [count]             List unread messages"
        echo "  focused [count]            List focused inbox"
        echo "  sent [count]               List sent items"
        echo "  folder <name> [count]      List messages in any folder by name"
        echo "  from <email> [count]       Filter by sender"
        echo "  search <query> [count]     Search emails"
        echo "  read <id>                  Read full message"
        echo "  preview <id>               Quick preview"
        echo
        echo "Sending:"
        echo "  draft <to> <subject> <body>    Create plain text draft"
        echo "  mddraft <to> <subject> <body>  Create draft with markdown formatting"
        echo "  reply <id> <body>              Create reply draft (plain text)"
        echo "  mdreply <id> <body>            Create reply draft with markdown formatting"
        echo "  followup <sent-id> [body]      Create chaser reply to your sent email"
        echo "  update <draft-id> <field> <value>  Update draft (subject/body/mdbody/to/cc/bcc)"
        echo "  send <draft-id>                Send draft"
        echo "  drafts [count]                 List drafts"
        echo
        echo "Attachments:"
        echo "  attachments <id>           List attachments on message"
        echo "  download <id> [att-id]     Download attachment(s) to ./inbox/"
        echo "  attach <draft-id> <file>   Add attachment to draft (up to 150MB)"
        echo
        echo "Management:"
        echo "  markread <id>              Mark as read"
        echo "  markunread <id>            Mark as unread"
        echo "  delete <id>                Delete message"
        echo "  archive <id>               Archive message"
        echo "  move <id> <folder>         Move message to folder"
        echo "  mkdir <name> [parent]      Create folder (subfolder if parent given)"
        echo "  folders                    List top-level mail folders"
        echo "  subfolders [parent]        List subfolders (default: inbox)"
        echo "  stats                      Inbox statistics"
        ;;
esac
