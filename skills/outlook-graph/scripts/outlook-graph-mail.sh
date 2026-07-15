#!/bin/bash
# Outlook Mail Operations via Microsoft Graph API

set -e

BASE_DIR="$HOME/.outlook-graph"

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
ID_CACHE_FILE="$CONFIG_DIR/id_cache.json"
GRAPH_URL="https://graph.microsoft.com/v1.0"

# Check credentials
if [ ! -f "$CREDS_FILE" ]; then
    echo "Error: Account '$ACCOUNT' not configured. Run: outlook-graph-setup.sh --account $ACCOUNT"
    exit 1
fi

# --- Token management -------------------------------------------------------
# The access token is resolved from a locally-stored absolute expiry
# (expires_at), so the common path makes NO network pre-flight call. The token
# is refreshed over the network only when it is missing/expired, or when Graph
# rejects it mid-run (handled reactively in api_call).

# Refresh the access token, stamp an absolute expiry into credentials.json, and
# print the new access token. Errors go to stderr so captured stdout stays a
# clean token (empty on failure -> the guard below catches it).
refresh_access_token() {
    local refresh_token client_id client_secret now response expires_in
    refresh_token=$(jq -r '.refresh_token // empty' "$CREDS_FILE")
    client_id=$(jq -r '.client_id // empty' "$CONFIG_FILE")
    client_secret=$(jq -r '.client_secret // empty' "$CONFIG_FILE")

    if [ -z "$refresh_token" ]; then
        echo "Error: No refresh token. Run outlook-graph-setup.sh to re-authenticate." >&2
        return 1
    fi

    now=$(date +%s)
    response=$(curl -s -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
        -H "Content-Type: application/x-www-form-urlencoded" \
        -d "client_id=$client_id" \
        -d "client_secret=$client_secret" \
        -d "refresh_token=$refresh_token" \
        -d "grant_type=refresh_token" \
        -d "scope=offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read")

    if echo "$response" | jq -e '.error' > /dev/null 2>&1; then
        echo "Error refreshing token: $(echo "$response" | jq -r '.error_description // .error')" >&2
        return 1
    fi

    expires_in=$(echo "$response" | jq -r '.expires_in // 3600')
    echo "$response" | jq --argjson at "$((now + expires_in))" '. + {expires_at: $at}' > "$CREDS_FILE"
    chmod 600 "$CREDS_FILE"
    jq -r '.access_token' "$CREDS_FILE"
}

# Print a valid access token, refreshing over the network only when needed.
ensure_valid_token() {
    local access_token expires_at now
    access_token=$(jq -r '.access_token // empty' "$CREDS_FILE")
    expires_at=$(jq -r '.expires_at // 0' "$CREDS_FILE")
    now=$(date +%s)

    # 60s safety margin. A missing/zero expires_at always falls through to refresh
    # (e.g. first run after upgrade, before an expiry has been stamped).
    if [ -n "$access_token" ] && [ "$now" -lt "$((expires_at - 60))" ]; then
        echo "$access_token"
        return 0
    fi
    refresh_access_token
}

ACCESS_TOKEN=$(ensure_valid_token) || true

if [ -z "$ACCESS_TOKEN" ] || [ "$ACCESS_TOKEN" = "null" ]; then
    echo "Error: Invalid access token. Run outlook-graph-setup.sh to re-authenticate."
    exit 1
fi

# Low-level Graph request using the current $ACCESS_TOKEN.
_graph_request() {
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

# As _graph_request, but streams the body from a FILE rather than an argv string.
# Required for any payload that can exceed Linux's ~128KB single-argument limit
# (MAX_ARG_STRLEN) - notably base64 file attachments, where `-d "$data"` fails
# with "Argument list too long" for files larger than ~96KB.
_graph_request_file() {
    local method="$1"
    local endpoint="$2"
    local body_file="$3"

    curl -s -X "$method" "${GRAPH_URL}${endpoint}" \
        -H "Authorization: Bearer $ACCESS_TOKEN" \
        -H "Content-Type: application/json" \
        --data-binary @"$body_file"
}

# Shared retry/error wrapper around a low-level request function ($1), so every
# Graph call gets the same behaviour regardless of how its body is sent.
# Transparently refreshes the token and retries once if Graph rejects it mid-run
# (revoked, clock skew, or a race with local expiry). A transport failure (curl
# non-zero) is turned into a JSON error so callers can surface it — but a
# legitimately empty body (HTTP 204/202 from DELETE/send) is left empty, since
# callers treat "no .error" as success.
_api_call() {
    local transport="$1"; shift
    local response rc
    response=$("$transport" "$@") && rc=0 || rc=$?

    if [ -z "${OUTLOOK_TOKEN_RETRIED:-}" ] && \
       printf '%s' "$response" | jq -e 'objects | .error.code == "InvalidAuthenticationToken"' >/dev/null 2>&1; then
        OUTLOOK_TOKEN_RETRIED=1
        ACCESS_TOKEN=$(refresh_access_token) || true
        response=$("$transport" "$@") && rc=0 || rc=$?
    fi

    if [ "$rc" -ne 0 ] && [ -z "$response" ]; then
        response='{"error":{"code":"NetworkError","message":"Request to Microsoft Graph failed (network error, timeout, or connectivity issue)."}}'
    fi
    printf '%s' "$response"
}

# API call helper: request body passed as an argv string.
api_call() {
    _api_call _graph_request "$@"
}

# As api_call, but streams the request body from a file. Used by the attachment
# upload path, whose base64 payload would otherwise overflow MAX_ARG_STRLEN.
api_call_file() {
    _api_call _graph_request_file "$@"
}

# Fail loudly on a Graph error. These write commands used to pipe their response
# to /dev/null and print "Marked as read" / "Message deleted" unconditionally, so
# a REJECTED write still reported success and the caller believed a change had
# been made that had not. Never claim a write succeeded without reading the
# response. (A legitimately empty 204 body is success and stays silent.)
die_on_error() {
    local response="$1" context="$2"
    if [ -n "$response" ] && printf '%s' "$response" | jq -e '.error' > /dev/null 2>&1; then
        echo "Error $context:" >&2
        printf '%s' "$response" | jq -r '.error.message // .error.code' >&2
        exit 1
    fi
}

# Cache message IDs from an API response for fast short-ID resolution.
# Called after every listing command so that resolve_message_id can find
# messages from any folder (inbox, subfolders, drafts, sent) without
# expensive API cascading.
# Written atomically: several agents/shells can share one ~/.outlook-graph/<account>/
# and a plain `>` redirect lets a concurrent writer be observed mid-write, so a
# reader can see a truncated (invalid) cache. Write to a temp file in the same
# directory, then rename — rename is atomic, so a reader sees either the old
# cache or the new one, never a half-written one.
cache_message_ids() {
    local response="$1" tmp
    tmp=$(mktemp "$CONFIG_DIR/.id_cache.XXXXXX" 2>/dev/null) || return 0
    if echo "$response" | jq -c '[.value[].id // empty]' > "$tmp" 2>/dev/null; then
        mv -f "$tmp" "$ID_CACHE_FILE" 2>/dev/null || rm -f "$tmp"
    else
        rm -f "$tmp"
    fi
}

# --- Folder resolution ------------------------------------------------------
# One resolver, shared by folder/move/batch-move/rename/rmdir/mkdir so a folder
# name resolves identically everywhere (the previous inline copies searched to
# different depths, so a deeply-nested folder could be listable but not movable).

# Breadth-first search for a folder whose displayName matches $1 (case-insensitive),
# starting at all top-level folders and descending through childFolders up to
# $2 levels (default 4). Ties resolve to the SHALLOWEST match; use a
# Parent/Child path to disambiguate. Prints the ID + returns 0, or returns 1.
#
# The Deleted Items subtree is NEVER descended into for a bare name. Deleting a
# folder in Outlook moves it (still named) into Deleted Items, so a mailbox that
# has ever had a folder deleted keeps a same-named ghost in the bin. Graph lists
# Deleted Items BEFORE Inbox, so without this prune a bare name could resolve to
# the ghost — silently filing live mail into the bin on `move`, or renaming /
# deleting the wrong folder. Target a binned folder deliberately with the path
# form ("Deleted Items/Old Project"), which still resolves.
_find_folder_by_name() {
    local target="$1" max_depth="${2:-4}"
    local resp match level_ids depth next_ids pid children deleted_id

    resp=$(api_call GET "/me/mailFolders?\$top=200")
    match=$(printf '%s' "$resp" | jq -r --arg n "$target" '.value[]? | select((.displayName|ascii_downcase)==($n|ascii_downcase)) | .id' | head -1)
    [ -n "$match" ] && { printf '%s' "$match"; return 0; }

    deleted_id=$(api_call GET "/me/mailFolders/deleteditems?\$select=id" | jq -r '.id // empty')
    level_ids=$(printf '%s' "$resp" | jq -r --arg del "$deleted_id" '.value[]? | select(.id != $del) | .id')

    depth=1
    while [ "$depth" -le "$max_depth" ] && [ -n "$level_ids" ]; do
        next_ids=""
        for pid in $level_ids; do
            children=$(api_call GET "/me/mailFolders/$pid/childFolders?\$top=200" 2>/dev/null)
            match=$(printf '%s' "$children" | jq -r --arg n "$target" '.value[]? | select((.displayName|ascii_downcase)==($n|ascii_downcase)) | .id' 2>/dev/null | head -1)
            [ -n "$match" ] && { printf '%s' "$match"; return 0; }
            next_ids="$next_ids $(printf '%s' "$children" | jq -r '.value[]?.id' 2>/dev/null)"
        done
        level_ids="$next_ids"
        depth=$((depth + 1))
    done
    return 1
}

# Walk a "Parent/Child/Grandchild" path, resolving each segment under the
# previous. Prints the leaf ID + returns 0, or returns 1 if any segment misses.
_resolve_folder_path() {
    local path="$1" parent_id="" first=1 seg match oldIFS
    oldIFS="$IFS"; IFS='/'; read -ra segs <<< "$path"; IFS="$oldIFS"
    for seg in "${segs[@]}"; do
        [ -z "$seg" ] && continue
        if [ "$first" = 1 ]; then
            match=$(resolve_folder_id "$seg") || match=""
            first=0
        else
            match=$(api_call GET "/me/mailFolders/$parent_id/childFolders?\$top=200" | jq -r --arg n "$seg" '.value[]? | select((.displayName|ascii_downcase)==($n|ascii_downcase)) | .id' | head -1)
        fi
        [ -z "$match" ] && return 1
        parent_id="$match"
    done
    [ -n "$parent_id" ] && { printf '%s' "$parent_id"; return 0; }
    return 1
}

# Resolve a folder name (or "Parent/Child" path) to its ID. Order: path form,
# then well-known aliases, then a case-insensitive BFS by name across the tree.
# Prints the ID + returns 0 on a match; returns 1 on no match.
resolve_folder_id() {
    local name="$1"

    if [[ "$name" == */* ]]; then
        _resolve_folder_path "$name"
        return $?
    fi

    local lc wid
    lc=$(printf '%s' "$name" | tr '[:upper:]' '[:lower:]')
    case "$lc" in
        inbox|drafts|sentitems|deleteditems|archive|junkemail)
            wid=$(api_call GET "/me/mailFolders/$lc" | jq -r '.id // empty')
            [ -n "$wid" ] && { printf '%s' "$wid"; return 0; }
            ;;
    esac

    _find_folder_by_name "$name"
}

# --- HTML -> readable plain text ---------------------------------------------
# A jq function, injected into the `read` filter. Block-level tags become line
# breaks BEFORE tags are stripped; otherwise "</p><p>" collapses and the text
# runs together ("Hello,This is..."), list items merge into one line, and a
# reader can misjudge where one point ends and the next begins. That matters
# here more than most places: this skill's core rule is to read the whole
# message end-to-end, so the rendering must not garble it. Entities are decoded
# so "&amp;" reads as "&".
HTML_TO_TEXT='
    def html_to_text:
        gsub("(?is)<(script|style)[^>]*>.*?</(script|style)>"; " ")
      | gsub("(?i)<br[^>]*>"; "\n")
      | gsub("(?i)<li[^>]*>"; "\n- ")
      | gsub("(?i)</(p|div|tr|h[1-6]|blockquote|ul|ol|table)>"; "\n")
      | gsub("(?i)</t[dh]>"; " ")
      | gsub("<[^>]*>"; "")
      | gsub("&nbsp;"; " ") | gsub("&amp;"; "&") | gsub("&lt;"; "<")
      | gsub("&gt;"; ">") | gsub("&quot;"; "\"") | gsub("&#39;"; "'"'"'")
      | gsub("&#8217;"; "'"'"'")
      | gsub("[ \t]+"; " ")
      | gsub(" *\n *"; "\n")
      | gsub("\n{3,}"; "\n\n")
      | sub("^\\s+"; "") | sub("\\s+$"; "");
'

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
# Strategy: cache-first (instant), then cascade through API endpoints.
#
# The cache is populated by every listing command (inbox, folder, drafts, sent,
# search, etc.), so if you just listed messages from any folder, their full IDs
# are available without any API call.
#
# If cache misses, we cascade through multiple API endpoints to find the
# message regardless of which folder it's in (inbox, custom subfolder, drafts,
# or sent items).
#
# Usage: full_id=$(resolve_message_id "short_id_or_full_id" "messages|drafts|sentitems")
resolve_message_id() {
    local msg_id="$1"
    local folder="${2:-messages}"  # "messages" for all mail, "drafts" for drafts folder, "sentitems" for sent

    # If it looks like a full ID (very long), return as-is
    if [ ${#msg_id} -gt 100 ]; then
        echo "$msg_id"
        return 0
    fi

    local full_id=""
    local search_limit=500

    # 1. Check cache first (instant, no API call)
    #    Cache is populated by listing commands (inbox, folder, drafts, sent, search, etc.)
    if [ -f "$ID_CACHE_FILE" ]; then
        full_id=$(jq -r ".[] | select(endswith(\"$msg_id\"))" "$ID_CACHE_FILE" 2>/dev/null | head -1)
        if [ -n "$full_id" ]; then
            echo "$full_id"
            return 0
        fi
    fi

    # 2. Search the hinted folder first
    case "$folder" in
        drafts)
            full_id=$(api_call GET "/me/mailFolders/drafts/messages?\$top=$search_limit&\$select=id" | jq -r ".value[].id | select(endswith(\"$msg_id\"))" | head -1)
            ;;
        sentitems)
            full_id=$(api_call GET "/me/mailFolders/sentitems/messages?\$top=$search_limit&\$select=id" | jq -r ".value[].id | select(endswith(\"$msg_id\"))" | head -1)
            ;;
        *)
            full_id=$(api_call GET "/me/messages?\$top=$search_limit&\$select=id" | jq -r ".value[].id | select(endswith(\"$msg_id\"))" | head -1)
            ;;
    esac

    if [ -n "$full_id" ]; then
        echo "$full_id"
        return 0
    fi

    # 3. Cascade: search other locations the hinted folder wouldn't cover
    #    /me/messages does NOT include drafts, so always cascade to drafts.
    #    Drafts/sentitems hints don't cover all-mail, so cascade to /me/messages.
    if [ "$folder" != "drafts" ]; then
        full_id=$(api_call GET "/me/mailFolders/drafts/messages?\$top=200&\$select=id" | jq -r ".value[].id | select(endswith(\"$msg_id\"))" | head -1)
        if [ -n "$full_id" ]; then echo "$full_id"; return 0; fi
    fi

    if [ "$folder" != "sentitems" ]; then
        full_id=$(api_call GET "/me/mailFolders/sentitems/messages?\$top=$search_limit&\$select=id" | jq -r ".value[].id | select(endswith(\"$msg_id\"))" | head -1)
        if [ -n "$full_id" ]; then echo "$full_id"; return 0; fi
    fi

    if [ "$folder" = "drafts" ] || [ "$folder" = "sentitems" ]; then
        full_id=$(api_call GET "/me/messages?\$top=$search_limit&\$select=id" | jq -r ".value[].id | select(endswith(\"$msg_id\"))" | head -1)
        if [ -n "$full_id" ]; then echo "$full_id"; return 0; fi
    fi

    return 1
}

# Convert a comma/semicolon-separated address string into a JSON array of
# Graph recipient objects. Trims surrounding whitespace and drops empties.
# Usage: arr=$(recipients_to_json "a@x.com, b@y.com; c@z.com")
recipients_to_json() {
    jq -n --arg raw "$1" '
        ($raw | gsub(";"; ",") | split(",")
            | map(gsub("^\\s+|\\s+$"; ""))
            | map(select(length > 0))
            | map({emailAddress: {address: .}}))'
}

# URL-encode a string for safe use as a query-string value (spaces, &, #, :, +,
# quotes, non-ASCII). Without this, a query like "Q3 & Q4" is silently truncated.
urlencode() { jq -rn --arg s "$1" '$s|@uri'; }

# --- Markdown -> Outlook-safe HTML -------------------------------------------
# Aptos is the Microsoft 365 default font since 2024; the stack falls back to
# Segoe UI on older Outlook and Roboto / system sans elsewhere. All styles are
# inline because Outlook strips <style> blocks, and every <p> gets an inline
# margin because Outlook ignores paragraph margins that are not inline.
FONT_STACK="'Aptos', 'Aptos Display', 'Segoe UI', Roboto, sans-serif"

require_pandoc() {
    command -v pandoc &> /dev/null && return 0
    echo "Error: pandoc is required for markdown conversion"
    echo "Install with: brew install pandoc (macOS) or apt install pandoc (Linux)"
    exit 1
}

# Convert markdown ($1) to Outlook-safe HTML wrapped in a styled div.
# $2 = line-height (default 1.5; replies use 1.6).
md_to_html() {
    local md="$1" lh="${2:-1.5}" html
    html=$(printf '%s\n' "$md" | pandoc -f markdown -t html \
        | sed 's/<p>/<p style="margin: 0 0 14px 0;">/g')
    printf '<div style="font-family: %s; font-size: 14px; line-height: %s; color: #333;">\n%s\n</div>' \
        "$FONT_STACK" "$lh" "$html"
}

# Run a Graph message $search and print the result object newest-first.
# The whole query is wrapped in the double quotes Graph requires around a
# $search value — this is the documented form for plain text AND for KQL field
# operators alike ($search="subject:pizza", $search="from:john AND subject:x").
# $search cannot be combined with $orderby server-side, so results (which come
# back ranked/date-mixed across folders) are sorted client-side by
# receivedDateTime desc. Follows @odata.nextLink to collect up to <count>
# messages (hard cap 1000). Usage: run_message_search "<text-or-KQL>" [count|all]
run_message_search() {
    local query="$1" max="${2:-10}"
    case "$max" in all|ALL) max=1000 ;; esac
    [[ "$max" =~ ^[0-9]+$ ]] || max=10
    [ "$max" -gt 1000 ] && max=1000
    [ "$max" -lt 1 ] && max=1

    # Strip any quotes the caller already wrapped the value in, then re-wrap so we
    # never emit $search=""x"" (which Graph rejects).
    query="${query%\"}"; query="${query#\"}"

    local encoded page_size url merged page collected next
    encoded=$(urlencode "\"$query\"")

    page_size=100
    [ "$max" -lt "$page_size" ] && page_size="$max"
    url="/me/messages?\$search=${encoded}&\$top=${page_size}&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview"
    merged='[]'
    while [ -n "$url" ]; do
        page=$(api_call GET "$url")
        if echo "$page" | jq -e '.error' >/dev/null 2>&1; then
            echo "$page"              # propagate the error object unchanged
            return 0
        fi
        merged=$(jq -n --argjson a "$merged" --argjson b "$(echo "$page" | jq '.value // []')" '$a + $b')
        collected=$(echo "$merged" | jq 'length')
        [ "$collected" -ge "$max" ] && break
        next=$(echo "$page" | jq -r '."@odata.nextLink" // empty')
        [ -z "$next" ] && break
        url="${next#"$GRAPH_URL"}"    # nextLink is absolute; strip base for api_call
    done
    echo "$merged" | jq --argjson max "$max" '{value: (sort_by(.receivedDateTime) | reverse | .[0:$max])}'
}

# Commands
case "$1" in
    inbox)
        count="${2:-10}"
        echo "Fetching inbox ($count messages)..."
        result=$(api_call GET "/me/mailFolders/inbox/messages?\$top=$count&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview")
        cache_message_ids "$result"
        echo "$result" | format_messages
        ;;

    unread)
        count="${2:-10}"
        echo "Fetching unread messages..."
        result=$(api_call GET "/me/mailFolders/inbox/messages?\$filter=isRead%20eq%20false&\$top=$count&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview")
        cache_message_ids "$result"
        echo "$result" | format_messages
        ;;

    focused)
        count="${2:-10}"
        echo "Fetching focused inbox..."
        # Graph rejects $orderby combined with an inferenceClassification $filter
        # ("The restriction or sort order is too complex for this operation"),
        # so sort newest-first client-side instead.
        result=$(api_call GET "/me/mailFolders/inbox/messages?\$filter=inferenceClassification%20eq%20'focused'&\$top=$count&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" \
            | jq 'if .error then . else {value: (.value | sort_by(.receivedDateTime) | reverse)} end')
        cache_message_ids "$result"
        echo "$result" | format_messages
        ;;

    sent)
        count="${2:-10}"
        echo "Fetching sent items ($count messages)..."
        result=$(api_call GET "/me/mailFolders/sentitems/messages?\$top=$count&\$orderby=sentDateTime%20desc&\$select=id,subject,toRecipients,sentDateTime,bodyPreview")
        cache_message_ids "$result"
        echo "$result" | jq -r '
            if .error then
                "Error: \(.error.message // .error.code // "Unknown API error")"
            elif (.value | length) == 0 then
                "No sent messages found."
            else
                def short_id: .[-20:];
                .value | to_entries | .[] |
                "[\(.key + 1)] \(.value.id | short_id) | \(.value.sentDateTime | split("T")[0]) | To: \(.value.toRecipients[0].emailAddress.address // "unknown") | \(.value.subject // "(no subject)")"
            end
        '
        ;;

    from)
        sender="$2"
        count="${3:-10}"
        if [ -z "$sender" ]; then
            echo "Usage: outlook-graph-mail.sh from <sender-email> [count]"
            exit 1
        fi
        echo "Fetching emails from $sender (newest first)..."
        result=$(run_message_search "from:$sender" "$count")
        cache_message_ids "$result"
        echo "$result" | format_messages
        ;;

    search)
        query="$2"
        count="${3:-10}"
        if [ -z "$query" ]; then
            echo "Usage: outlook-graph-mail.sh search <query> [count]"
            echo "  <query>  free text OR KQL field operators. Examples:"
            echo "             search \"invoice March\"                  free text"
            echo "             search 'subject:invoice AND from:jane@x.com'"
            echo "             search 'to:me@x.com AND body:contract' 50"
            echo "  [count]  results to return: default 10, max 1000, or 'all'"
            echo "  Results come back ranked by Graph, then sorted newest-first."
            exit 1
        fi
        # A bare email address -> precise from: match; otherwise free-text/KQL.
        if [[ "$query" =~ ^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$ ]]; then
            echo "Searching for emails from $query (newest first)..."
            result=$(run_message_search "from:$query" "$count")
        else
            echo "Searching for: $query (newest first)..."
            result=$(run_message_search "$query" "$count")
        fi
        cache_message_ids "$result"
        echo "$result" | format_messages
        ;;

    read)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh read <message-id>"
            exit 1
        fi

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Reading message..."
        # Expand attachments (metadata only) in the same call so image-only /
        # attachment-only messages are never rendered as a blank body. Inline
        # images (cid: refs) come back as attachments with isInline=true.
        api_call GET "/me/messages/$msg_id?\$expand=attachments(\$select=id,name,size,contentType,isInline)" | jq -r "$HTML_TO_TEXT"'
            def format_size:
                if . == null then "?"
                elif . < 1024 then "\(.)B"
                elif . < 1048576 then "\((. / 1024 * 10 | floor) / 10)KB"
                else "\((. / 1048576 * 10 | floor) / 10)MB"
                end;
            ( (.body.content // "") | html_to_text ) as $text
            | ( .subject // "" ) as $subj
            | ( .attachments // [] ) as $atts
            | "Subject: \(if ($subj | length) > 0 then $subj else "(no subject)" end)",
              "From: \(.from.emailAddress.name // "") <\(.from.emailAddress.address // "")>",
              "To: \([.toRecipients[]?.emailAddress | "\(.name // "") <\(.address)>"] | join(", "))",
              ( if (.ccRecipients // [] | length) > 0 then "Cc: \([.ccRecipients[]?.emailAddress | "\(.name // "") <\(.address)>"] | join(", "))" else empty end),
              "Date: \(.receivedDateTime)",
              "---",
              ( if ($text | length) > 0 then $text
                elif ($atts | length) > 0 then "(no text body - this message is \($atts | length) attachment(s)/inline image(s); listed below)"
                else "(no text body)" end ),
              ( if ($atts | length) > 0 then
                  "",
                  "--- Attachments (\($atts | length)) ---",
                  ( $atts[] | "- \(.name // "(unnamed)") | \(.contentType) | \(.size | format_size)\(if .isInline then " | inline" else "" end)" ),
                  "Download with: outlook-graph-mail.sh download <message-id>   (saves to ./inbox/)"
                else empty end )'
        ;;

    preview)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh preview <message-id>"
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
            echo "Usage: outlook-graph-mail.sh draft <to-email> <subject> <body>"
            exit 1
        fi

        echo "Creating draft..."
        from_address="${OUTLOOK_FROM_ADDRESS:-}"
        from_name="${OUTLOOK_FROM_NAME:-}"
        if [ -n "$from_address" ]; then
            payload=$(jq -n \
                --arg to "$to" \
                --arg subject "$subject" \
                --arg body "${body:-}" \
                --arg from_addr "$from_address" \
                --arg from_name "$from_name" \
                '{
                    subject: $subject,
                    body: {
                        contentType: "Text",
                        content: $body
                    },
                    from: {
                        emailAddress: {
                            address: $from_addr,
                            name: $from_name
                        }
                    },
                    toRecipients: [
                        {
                            emailAddress: {
                                address: $to
                            }
                        }
                    ]
                }')
        else
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
        fi

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
            echo "Usage: outlook-graph-mail.sh mddraft <to-email> <subject> <markdown-body>"
            exit 1
        fi

        require_pandoc

        echo "Creating markdown draft..."
        html_body=$(md_to_html "${body:-}")

        from_address="${OUTLOOK_FROM_ADDRESS:-}"
        from_name="${OUTLOOK_FROM_NAME:-}"
        if [ -n "$from_address" ]; then
            payload=$(jq -n \
                --arg to "$to" \
                --arg subject "$subject" \
                --arg body "$html_body" \
                --arg from_addr "$from_address" \
                --arg from_name "$from_name" \
                '{
                    subject: $subject,
                    body: {
                        contentType: "HTML",
                        content: $body
                    },
                    from: {
                        emailAddress: {
                            address: $from_addr,
                            name: $from_name
                        }
                    },
                    toRecipients: [
                        {
                            emailAddress: {
                                address: $to
                            }
                        }
                    ]
                }')
        else
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
        fi

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
            echo "Usage: outlook-graph-mail.sh reply <message-id> <body>"
            exit 1
        fi

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Creating reply-all draft..."
        payload=$(jq -n --arg body "$body" '{comment: $body}')

        result=$(api_call POST "/me/messages/$msg_id/createReplyAll" "$payload")
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating reply:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        echo "Reply draft created!"
        echo "Draft ID: ${draft_id: -20}"
        echo
        echo "$result" | jq -r '
            "To:      \(.toRecipients | map(.emailAddress.address) | join(", "))",
            (if (.ccRecipients | length) > 0 then "Cc:      \(.ccRecipients | map(.emailAddress.address) | join(", "))" else empty end),
            (if (.bccRecipients | length) > 0 then "Bcc:     \(.bccRecipients | map(.emailAddress.address) | join(", "))" else empty end),
            "Subject: \(.subject)"
        '
        ;;

    update)
        draft_id="$2"
        field="$3"
        value="$4"
        if [ -z "$draft_id" ] || [ -z "$field" ]; then
            echo "Usage: outlook-graph-mail.sh update <draft-id> <field> <value>"
            echo ""
            echo "Fields:"
            echo "  subject <text>     Update subject line"
            echo "  body <text>        Replace body (plain text)"
            echo "  mdbody <markdown>  Replace body (markdown -> HTML)"
            echo "  to <emails>        Set To recipient(s) - comma/semicolon-separated; replaces the To line"
            echo "  cc <emails>        Add CC recipient(s) - comma/semicolon-separated; dedups; empty string clears CC"
            echo "  bcc <emails>       Add BCC recipient(s) - comma/semicolon-separated; dedups; empty string clears BCC"
            echo "  importance <level> Set message importance: low, normal, or high"
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
                require_pandoc
                echo "Updating body (markdown -> HTML)..."
                html_body=$(md_to_html "$value")

                # Preserve reply chain if the draft was created via mdreply or followup.
                # Those commands inject a `<span data-mdreply-chain-start="1"></span>`
                # marker between the new message and the quoted history. If we find
                # that marker, keep everything from the marker onwards.
                chain_marker='<span data-mdreply-chain-start="1"></span>'
                existing_body=$(api_call GET "/me/messages/$draft_id?\$select=body" | jq -r '.body.content // ""')
                if [[ "$existing_body" == *"$chain_marker"* ]]; then
                    # Everything from the first occurrence of the marker onwards
                    chain_part="${chain_marker}${existing_body#*"$chain_marker"}"
                    full_body="${html_body}<br/>${chain_part}"
                else
                    full_body="${html_body}"
                fi

                payload=$(jq -n --arg body "$full_body" '{body: {contentType: "HTML", content: $body}}')
                ;;
            to)
                if [ -z "$value" ]; then
                    echo "Error: Email address required"
                    exit 1
                fi
                # Accepts a comma/semicolon-separated list; replaces the To line.
                new_recips=$(recipients_to_json "$value")
                if [ "$(echo "$new_recips" | jq 'length')" -eq 0 ]; then
                    echo "Error: No valid email address provided"
                    exit 1
                fi
                echo "Updating To recipient(s)..."
                payload=$(jq -n --argjson arr "$new_recips" '{toRecipients: $arr}')
                ;;
            cc)
                # Empty value clears the CC line; otherwise accepts a
                # comma/semicolon-separated list and appends to existing CC,
                # skipping addresses already present (case-insensitive) so re-adding
                # is a no-op rather than creating duplicate/malformed recipients.
                if [ -z "$value" ]; then
                    echo "Clearing CC recipients..."
                    payload='{"ccRecipients": []}'
                else
                    new_recips=$(recipients_to_json "$value")
                    if [ "$(echo "$new_recips" | jq 'length')" -eq 0 ]; then
                        echo "Error: No valid email address provided"
                        exit 1
                    fi
                    echo "Adding CC recipient(s)..."
                    existing=$(api_call GET "/me/messages/$draft_id?\$select=ccRecipients" | jq '.ccRecipients // []')
                    payload=$(jq -n --argjson ex "$existing" --argjson new "$new_recips" '
                        ($ex | map(.emailAddress.address // "" | ascii_downcase)) as $have
                        | {ccRecipients: ($ex + ($new | map(select((.emailAddress.address // "" | ascii_downcase) as $a | ($have | index($a)) | not))))}')
                fi
                ;;
            bcc)
                # Empty value clears the BCC line; otherwise accepts a
                # comma/semicolon-separated list and appends to existing BCC,
                # skipping addresses already present (case-insensitive).
                if [ -z "$value" ]; then
                    echo "Clearing BCC recipients..."
                    payload='{"bccRecipients": []}'
                else
                    new_recips=$(recipients_to_json "$value")
                    if [ "$(echo "$new_recips" | jq 'length')" -eq 0 ]; then
                        echo "Error: No valid email address provided"
                        exit 1
                    fi
                    echo "Adding BCC recipient(s)..."
                    existing=$(api_call GET "/me/messages/$draft_id?\$select=bccRecipients" | jq '.bccRecipients // []')
                    payload=$(jq -n --argjson ex "$existing" --argjson new "$new_recips" '
                        ($ex | map(.emailAddress.address // "" | ascii_downcase)) as $have
                        | {bccRecipients: ($ex + ($new | map(select((.emailAddress.address // "" | ascii_downcase) as $a | ($have | index($a)) | not))))}')
                fi
                ;;
            importance)
                lvl=$(printf '%s' "$value" | tr '[:upper:]' '[:lower:]')
                case "$lvl" in
                    low|normal|high) ;;
                    *) echo "Error: importance must be low, normal, or high"; exit 1 ;;
                esac
                echo "Setting importance to $lvl..."
                payload=$(jq -n --arg v "$lvl" '{importance: $v}')
                ;;
            *)
                echo "Error: Unknown field '$field'"
                echo "Valid fields: subject, body, mdbody, to, cc, bcc, importance"
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
        echo "$result" | jq -r '
            "To:      \((.toRecipients // []) | map(.emailAddress.address) | join(", ") | (if . == "" then "none" else . end))",
            (if ((.ccRecipients // []) | length) > 0 then "Cc:      \(.ccRecipients | map(.emailAddress.address) | join(", "))" else empty end),
            (if ((.bccRecipients // []) | length) > 0 then "Bcc:     \(.bccRecipients | map(.emailAddress.address) | join(", "))" else empty end),
            "Subject: \(.subject // "(no subject)")"
        '
        ;;

    mdreply)
        msg_id="$2"
        body="$3"
        if [ -z "$msg_id" ] || [ -z "$body" ]; then
            echo "Usage: outlook-graph-mail.sh mdreply <message-id> <markdown-body>"
            exit 1
        fi

        require_pandoc

        # Resolve short ID to full ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Creating reply-all draft with markdown formatting..."

        # Step 1: Create reply-all draft (empty comment to get thread headers)
        # createReplyAll preserves all original To: and Cc: recipients - this is the
        # correct default for litigation/business threads where dropping CCs is harmful.
        # To reply to the sender only, use mdreply and then `update to <email>` after.
        result=$(api_call POST "/me/messages/$msg_id/createReplyAll" '{}')
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating reply:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        # Step 2: Get the existing body (contains the quoted thread)
        existing_body=$(echo "$result" | jq -r '.body.content // ""')

        # Step 3: Convert markdown to HTML and prepend to the existing thread.
        # The `data-mdreply-chain-start` marker lets `update mdbody` find the
        # boundary between the new message and the quoted chain so subsequent
        # edits can replace the message without losing the chain.
        html_body=$(md_to_html "$body" 1.6)
        combined_body="${html_body}
<br/>
<span data-mdreply-chain-start=\"1\"></span>
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
        echo "$patch_result" | jq -r '
            "To:      \(.toRecipients | map(.emailAddress.address) | join(", "))",
            (if (.ccRecipients | length) > 0 then "Cc:      \(.ccRecipients | map(.emailAddress.address) | join(", "))" else empty end),
            (if (.bccRecipients | length) > 0 then "Bcc:     \(.bccRecipients | map(.emailAddress.address) | join(", "))" else empty end),
            "Subject: \(.subject)"
        '
        ;;

    followup)
        msg_id="$2"
        body="$3"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh followup <sent-message-id> [markdown-body]"
            echo "       Creates a follow-up reply to your own sent email (chaser)"
            echo "       Body defaults to a standard follow-up message if not provided"
            exit 1
        fi

        require_pandoc

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
        result=$(api_call POST "/me/messages/$msg_id/createReplyAll" '{}')
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating follow-up:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        # Step 2: Get the existing body (contains the quoted thread)
        existing_body=$(echo "$result" | jq -r '.body.content // ""')

        # Step 3: Convert markdown to HTML and prepend to the existing thread.
        # The `data-mdreply-chain-start` marker lets `update mdbody` find the
        # boundary between the new message and the quoted chain so subsequent
        # edits can replace the message without losing the chain.
        html_body=$(md_to_html "$body" 1.6)
        combined_body="${html_body}
<br/>
<span data-mdreply-chain-start=\"1\"></span>
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
        echo "$patch_result" | jq -r '
            "To:      \(.toRecipients | map(.emailAddress.address) | join(", "))",
            (if (.ccRecipients | length) > 0 then "Cc:      \(.ccRecipients | map(.emailAddress.address) | join(", "))" else empty end),
            (if (.bccRecipients | length) > 0 then "Bcc:     \(.bccRecipients | map(.emailAddress.address) | join(", "))" else empty end),
            "Subject: \(.subject)"
        '
        echo
        echo "Use 'outlook-graph-mail.sh send ${draft_id: -20}' to send"
        ;;

    forward)
        msg_id="$2"
        to="$3"
        body="$4"
        if [ -z "$msg_id" ] || [ -z "$to" ]; then
            echo "Usage: outlook-graph-mail.sh forward <message-id> <to-emails> [markdown-comment]"
            echo "       Creates a forward DRAFT with the full quoted message and any"
            echo "       attachments. <to-emails> is comma/semicolon-separated. The"
            echo "       optional comment is markdown, converted to styled HTML."
            exit 1
        fi

        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        recips=$(recipients_to_json "$to")
        if [ "$(echo "$recips" | jq 'length')" -eq 0 ]; then
            echo "Error: No valid recipient address provided"
            exit 1
        fi

        echo "Creating forward draft..."
        result=$(api_call POST "/me/messages/$msg_id/createForward" "$(jq -n --argjson r "$recips" '{toRecipients: $r}')")
        draft_id=$(echo "$result" | jq -r '.id')

        if [ -z "$draft_id" ] || [ "$draft_id" = "null" ]; then
            echo "Error creating forward:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        # Optional markdown comment above the forwarded message, with the same
        # chain marker mdreply uses so `update mdbody` can edit it safely.
        if [ -n "$body" ]; then
            require_pandoc
            existing_body=$(echo "$result" | jq -r '.body.content // ""')
            html_body=$(md_to_html "$body" 1.6)
            combined_body="${html_body}
<br/>
<span data-mdreply-chain-start=\"1\"></span>
${existing_body}"
            patch_result=$(api_call PATCH "/me/messages/$draft_id" \
                "$(jq -n --arg body "$combined_body" '{body: {contentType: "HTML", content: $body}}')")
            if echo "$patch_result" | jq -e '.error' > /dev/null 2>&1; then
                echo "Error adding comment to forward draft:"
                echo "$patch_result" | jq -r '.error.message'
                exit 1
            fi
            result="$patch_result"
        fi

        echo "Forward draft created!"
        echo "Draft ID: ${draft_id: -20}"
        echo
        echo "$result" | jq -r '
            "To:      \(.toRecipients | map(.emailAddress.address) | join(", "))",
            "Subject: \(.subject)"
        '
        echo
        echo "Use 'outlook-graph-mail.sh send ${draft_id: -20}' to send"
        ;;

    send)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh send <draft-id>"
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
        result=$(api_call GET "/me/mailFolders/drafts/messages?\$top=$count&\$orderby=createdDateTime%20desc&\$select=id,subject,toRecipients,createdDateTime")
        cache_message_ids "$result"
        echo "$result" | jq -r '
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
            echo "Usage: outlook-graph-mail.sh markread <message-id>"
            exit 1
        fi

        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        die_on_error "$(api_call PATCH "/me/messages/$msg_id" '{"isRead": true}')" "marking as read"
        echo "Marked as read"
        ;;

    markunread)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh markunread <message-id>"
            exit 1
        fi

        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        die_on_error "$(api_call PATCH "/me/messages/$msg_id" '{"isRead": false}')" "marking as unread"
        echo "Marked as unread"
        ;;

    flag)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh flag <message-id>"
            exit 1
        fi
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi
        die_on_error "$(api_call PATCH "/me/messages/$msg_id" '{"flag": {"flagStatus": "flagged"}}')" "flagging message"
        echo "Flagged for follow-up"
        ;;

    unflag)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh unflag <message-id>"
            exit 1
        fi
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi
        die_on_error "$(api_call PATCH "/me/messages/$msg_id" '{"flag": {"flagStatus": "notFlagged"}}')" "clearing flag"
        echo "Flag cleared"
        ;;

    flagged)
        count="${2:-10}"
        echo "Fetching flagged messages..."
        # $orderby cannot be combined with this $filter on all mailboxes, so
        # sort newest-first client-side.
        result=$(api_call GET "/me/messages?\$filter=flag/flagStatus%20eq%20'flagged'&\$top=$count&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" \
            | jq 'if .error then . else {value: (.value | sort_by(.receivedDateTime) | reverse)} end')
        cache_message_ids "$result"
        echo "$result" | format_messages
        ;;

    thread)
        msg_id="$2"
        count="${3:-25}"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh thread <message-id> [count]"
            echo "       Lists every message in the same conversation, oldest first."
            exit 1
        fi
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi
        conv_id=$(api_call GET "/me/messages/$msg_id?\$select=conversationId" | jq -r '.conversationId // empty')
        if [ -z "$conv_id" ]; then
            echo "Error: Could not determine conversation for message"
            exit 1
        fi
        echo "Fetching conversation (up to $count messages, oldest first)..."
        filter=$(urlencode "conversationId eq '$conv_id'")
        result=$(api_call GET "/me/messages?\$filter=$filter&\$top=$count&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview" \
            | jq 'if .error then . else {value: (.value | sort_by(.receivedDateTime))} end')
        cache_message_ids "$result"
        echo "$result" | format_messages
        ;;

    categories)
        echo "Master category list:"
        api_call GET "/me/outlook/masterCategories" | jq -r '
            if .error then
                "Error: \(.error.message // .error.code // "Unknown API error")"
            elif (.value | length) == 0 then
                "No categories defined."
            else
                .value[] | "- \(.displayName) (\(.color // "no colour"))"
            end
        '
        ;;

    categorize)
        msg_id="$2"
        cats="$3"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh categorize <message-id> <categories>"
            echo "       <categories> is comma-separated (must match names from"
            echo "       'categories'); an empty string clears all categories."
            exit 1
        fi
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi
        payload=$(jq -n --arg raw "$cats" \
            '{categories: ($raw | split(",") | map(gsub("^\\s+|\\s+$"; "")) | map(select(length > 0)))}')
        result=$(api_call PATCH "/me/messages/$msg_id" "$payload")
        if echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error setting categories:"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi
        echo "$result" | jq -r 'if (.categories | length) == 0 then "Categories cleared" else "Categories set: \(.categories | join(", "))" end'
        ;;

    junk)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh junk <message-id>"
            exit 1
        fi
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi
        junk_id=$(api_call GET "/me/mailFolders/junkemail" | jq -r '.id // empty')
        if [ -z "$junk_id" ]; then
            echo "Error: Junk Email folder not found"
            exit 1
        fi
        die_on_error "$(api_call POST "/me/messages/$msg_id/move" "{\"destinationId\": \"$junk_id\"}")" "moving to junk"
        echo "Moved to Junk Email"
        ;;

    notjunk)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh notjunk <message-id>"
            exit 1
        fi
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi
        inbox_id=$(api_call GET "/me/mailFolders/inbox" | jq -r '.id // empty')
        die_on_error "$(api_call POST "/me/messages/$msg_id/move" "{\"destinationId\": \"$inbox_id\"}")" "moving to inbox"
        echo "Moved back to Inbox"
        ;;

    delete)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh delete <message-id>"
            exit 1
        fi

        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        die_on_error "$(api_call DELETE "/me/messages/$msg_id")" "deleting message"
        echo "Message deleted"
        ;;

    archive)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh archive <message-id>"
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

        die_on_error "$(api_call POST "/me/messages/$msg_id/move" "{\"destinationId\": \"$archive_id\"}")" "archiving message"
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
            echo "Usage: outlook-graph-mail.sh folder <folder-name> [count]"
            exit 1
        fi

        echo "Finding folder '$folder_name'..."

        folder_id=$(resolve_folder_id "$folder_name") || true

        if [ -z "$folder_id" ]; then
            echo "Error: Folder '$folder_name' not found"
            exit 1
        fi

        echo "Fetching messages from '$folder_name' ($count messages)..."
        result=$(api_call GET "/me/mailFolders/$folder_id/messages?\$top=$count&\$orderby=receivedDateTime%20desc&\$select=id,subject,from,receivedDateTime,isRead,bodyPreview")
        cache_message_ids "$result"
        echo "$result" | format_messages
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
            echo "Usage: outlook-graph-mail.sh move <message-id> <folder-name>"
            exit 1
        fi

        # Resolve message ID
        if ! msg_id=$(resolve_message_id "$msg_id" "messages"); then
            echo "Error: Message not found with ID: $2"
            exit 1
        fi

        echo "Finding folder '$folder_name'..."

        dest_folder_id=$(resolve_folder_id "$folder_name") || true

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

    batch-move|bulk-move)
        folder_name="$2"
        if [ -z "$folder_name" ]; then
            echo "Usage: outlook-graph-mail.sh batch-move <folder-name> <id1> [id2 ...]"
            echo "       Message IDs may be given as arguments OR piped via stdin"
            echo "       (newline- or space-separated). The destination folder is"
            echo "       resolved once; moves run in batches of 20 via the Graph"
            echo "       \$batch endpoint."
            exit 1
        fi
        shift 2

        # Gather message IDs: from remaining args, else from stdin if piped
        ids=("$@")
        if [ ${#ids[@]} -eq 0 ] && [ ! -t 0 ]; then
            while read -r line; do
                for tok in $line; do
                    [ -n "$tok" ] && ids+=("$tok")
                done
            done
        fi
        if [ ${#ids[@]} -eq 0 ]; then
            echo "Error: no message IDs provided (as arguments or via stdin)"
            exit 1
        fi

        # Listings print SHORT (20-char) IDs, so resolve each to its full ID.
        # The resolver is cache-first, and every listing command populates the
        # cache, so the common list-then-move flow costs no extra API calls.
        resolved=()
        skipped=0
        for id in "${ids[@]}"; do
            if [ ${#id} -le 25 ]; then
                if full=$(resolve_message_id "$id" "messages"); then
                    resolved+=("$full")
                else
                    echo "  WARNING: could not resolve short ID '$id' - skipped"
                    skipped=$((skipped + 1))
                fi
            else
                resolved+=("$id")
            fi
        done
        ids=("${resolved[@]}")
        if [ ${#ids[@]} -eq 0 ]; then
            echo "Error: none of the provided IDs could be resolved"
            exit 1
        fi

        echo "Finding folder '$folder_name'..."
        dest_folder_id=$(resolve_folder_id "$folder_name") || true
        if [ -z "$dest_folder_id" ]; then
            echo "Error: Folder '$folder_name' not found"
            exit 1
        fi

        total=${#ids[@]}
        echo "Moving $total message(s) to '$folder_name' (batches of 20)..."
        moved=0
        failed=0
        i=0
        while [ $i -lt "$total" ]; do
            # Build a batch of up to 20 move sub-requests
            requests="[]"
            n=0
            while [ $n -lt 20 ] && [ $i -lt "$total" ]; do
                requests=$(echo "$requests" | jq \
                    --arg rid "$n" --arg mid "${ids[$i]}" --arg dest "$dest_folder_id" \
                    '. += [{id: $rid, method: "POST", url: ("/me/messages/" + $mid + "/move"), headers: {"Content-Type": "application/json"}, body: {destinationId: $dest}}]')
                n=$((n + 1))
                i=$((i + 1))
            done
            body=$(jq -n --argjson reqs "$requests" '{requests: $reqs}')
            resp=$(api_call POST "/\$batch" "$body")

            ok=$(echo "$resp" | jq '[.responses[]? | select(.status >= 200 and .status < 300)] | length' 2>/dev/null || echo 0)
            bad=$(echo "$resp" | jq '[.responses[]? | select(.status >= 300)] | length' 2>/dev/null || echo 0)
            moved=$((moved + ${ok:-0}))
            failed=$((failed + ${bad:-0}))

            # Surface any per-message failures
            echo "$resp" | jq -r '.responses[]? | select(.status >= 300) | "  FAILED [\(.status)] \(.body.error.message // "unknown error")"' 2>/dev/null || true
            echo "  progress: $moved/$total moved"
        done

        summary="Done: $moved moved, $failed failed."
        [ "$skipped" -gt 0 ] && summary="$summary $skipped skipped (unresolvable short ID)."
        echo "$summary"
        [ "$failed" -eq 0 ] && [ "$skipped" -eq 0 ] || exit 1
        ;;

    mkdir)
        folder_name="$2"
        parent_folder="$3"
        if [ -z "$folder_name" ]; then
            echo "Usage: outlook-graph-mail.sh mkdir <folder-name> [parent-folder]"
            echo "       Without parent-folder, creates a top-level folder"
            exit 1
        fi

        if [ -z "$parent_folder" ]; then
            # Create top-level folder
            echo "Creating top-level folder '$folder_name'..."
            payload=$(jq -n --arg name "$folder_name" '{"displayName": $name}')
            result=$(api_call POST "/me/mailFolders" "$payload")
        else
            # Find parent folder ID (shared resolver: aliases, name, or path)
            echo "Finding parent folder '$parent_folder'..."
            parent_id=$(resolve_folder_id "$parent_folder") || true

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

    rename)
        old_name="$2"
        new_name="$3"
        if [ -z "$old_name" ] || [ -z "$new_name" ]; then
            echo "Usage: outlook-graph-mail.sh rename <folder-name> <new-name>"
            exit 1
        fi
        case "$(echo "$old_name" | tr '[:upper:]' '[:lower:]')" in
            inbox|drafts|sentitems|sent\ items|deleteditems|deleted\ items|archive|junkemail|junk\ email|outbox)
                echo "Refusing to rename well-known system folder '$old_name'"
                exit 1
                ;;
        esac
        fid=$(resolve_folder_id "$old_name") || true
        if [ -z "$fid" ]; then
            echo "Error: Folder '$old_name' not found"
            exit 1
        fi
        result=$(api_call PATCH "/me/mailFolders/$fid" "$(jq -n --arg n "$new_name" '{displayName: $n}')")
        if echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error renaming folder:"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi
        echo "Renamed '$old_name' -> '$new_name'"
        ;;

    rmdir)
        target="$2"
        force="$3"
        if [ -z "$target" ]; then
            echo "Usage: outlook-graph-mail.sh rmdir <folder-name> [--force]"
            echo "       Refuses to delete a non-empty folder unless --force is given"
            echo "       (deleted folder contents are moved to Deleted Items)."
            exit 1
        fi
        case "$(echo "$target" | tr '[:upper:]' '[:lower:]')" in
            inbox|drafts|sentitems|sent\ items|deleteditems|deleted\ items|archive|junkemail|junk\ email|outbox)
                echo "Refusing to delete well-known system folder '$target'"
                exit 1
                ;;
        esac
        fid=$(resolve_folder_id "$target") || true
        if [ -z "$fid" ]; then
            echo "Error: Folder '$target' not found"
            exit 1
        fi
        count=$(api_call GET "/me/mailFolders/$fid?\$select=totalItemCount" | jq -r '.totalItemCount // 0')
        if [ "${count:-0}" -gt 0 ] && [ "$force" != "--force" ]; then
            echo "Refusing to delete '$target': it contains $count message(s)."
            echo "Re-run with --force to delete anyway (contents move to Deleted Items)."
            exit 1
        fi
        result=$(api_call DELETE "/me/mailFolders/$fid")
        if [ -n "$result" ] && echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error deleting folder:"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi
        echo "Deleted folder '$target'"
        ;;

    attachments)
        msg_id="$2"
        if [ -z "$msg_id" ]; then
            echo "Usage: outlook-graph-mail.sh attachments <message-id>"
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
            echo "Usage: outlook-graph-mail.sh download <message-id> [attachment-id]"
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

        # Get attachments list (don't request contentBytes in listing - causes some attachments to be omitted)
        attachments=$(api_call GET "/me/messages/$msg_id/attachments?\$select=id,name,size,contentType")

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

            # Sanitize server-supplied name: prevent path traversal out of $DOWNLOAD_DIR
            att_name=$(basename "$att_name" | sed 's/\.\.//g')
            if [ -z "$att_name" ] || [ "$att_name" = "." ]; then
                att_name="attachment"
            fi

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

            # Always fetch via raw content endpoint (contentBytes not requested
            # in listing). -f makes curl fail on an HTTP error instead of
            # silently saving the Graph error JSON as the attachment file.
            if ! curl -sf -X GET "${GRAPH_URL}/me/messages/$msg_id/attachments/$att_id/\$value" \
                -H "Authorization: Bearer $ACCESS_TOKEN" \
                -o "$dest_path"; then
                rm -f "$dest_path"
                echo "FAILED: $att_name (download error from Graph)"
                continue
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
            echo "Usage: outlook-graph-mail.sh attach <draft-id> <file-path>"
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

            # Base64 encode to a TEMP FILE, never a shell variable: a variable
            # would be passed to jq/curl as an argv string and blow Linux's
            # ~128KB MAX_ARG_STRLEN limit ("Argument list too long") for any
            # attachment over ~96KB. Keep the bytes off the command line.
            b64_file=$(mktemp) || { echo "Error: cannot create temp file"; exit 1; }
            payload_file=$(mktemp) || { rm -f "$b64_file"; echo "Error: cannot create temp file"; exit 1; }

            if base64 --help 2>&1 | grep -q GNU; then
                base64 -w0 "$file_path" > "$b64_file"
            else
                base64 -i "$file_path" | tr -d '\n' > "$b64_file"
            fi

            jq -n \
                --arg name "$file_name" \
                --arg contentType "$content_type" \
                --rawfile contentBytes "$b64_file" \
                '{
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": $name,
                    "contentType": $contentType,
                    "contentBytes": ($contentBytes | rtrimstr("\n"))
                }' > "$payload_file"

            result=$(api_call_file POST "/me/messages/$draft_id/attachments" "$payload_file")
            rm -f "$b64_file" "$payload_file"

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
        echo "Usage: outlook-graph-mail.sh <command> [args]"
        echo
        echo "Reading:"
        echo "  inbox [count]              List inbox messages"
        echo "  unread [count]             List unread messages"
        echo "  focused [count]            List focused inbox"
        echo "  sent [count]               List sent items"
        echo "  folder <name> [count]      List messages in any folder by name"
        echo "  from <email> [count]       Filter by sender"
        echo "  search <query> [count]     Search emails"
        echo "  flagged [count]            List messages flagged for follow-up"
        echo "  thread <id> [count]        List the whole conversation, oldest first"
        echo "  read <id>                  Read full message"
        echo "  preview <id>               Quick preview"
        echo
        echo "Sending:"
        echo "  draft <to> <subject> <body>    Create plain text draft"
        echo "  mddraft <to> <subject> <body>  Create draft with markdown formatting"
        echo "  reply <id> <body>              Create reply draft (plain text)"
        echo "  mdreply <id> <body>            Create reply draft with markdown formatting"
        echo "  forward <id> <to> [comment]    Create forward draft (markdown comment optional)"
        echo "  followup <sent-id> [body]      Create chaser reply to your sent email"
        echo "  update <draft-id> <field> <value>  Update draft (subject/body/mdbody/to/cc/bcc/importance)"
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
        echo "  flag <id>                  Flag for follow-up"
        echo "  unflag <id>                Clear follow-up flag"
        echo "  categorize <id> <cats>     Set categories (comma-separated; \"\" clears)"
        echo "  categories                 List the mailbox's master categories"
        echo "  junk <id>                  Move message to Junk Email"
        echo "  notjunk <id>               Move message back to Inbox"
        echo "  delete <id>                Delete message"
        echo "  archive <id>               Archive message"
        echo "  move <id> <folder>         Move message to folder"
        echo "  batch-move <folder> <ids>  Move many messages (args or stdin) via \$batch"
        echo "  mkdir <name> [parent]      Create folder (subfolder if parent given)"
        echo "  rename <folder> <new>      Rename a folder"
        echo "  rmdir <folder> [--force]   Delete a folder (--force if non-empty)"
        echo "  folders                    List top-level mail folders"
        echo "  subfolders [parent]        List subfolders (default: inbox)"
        echo "  stats                      Inbox statistics"
        ;;
esac
