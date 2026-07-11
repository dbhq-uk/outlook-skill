#!/bin/bash
# Outlook Calendar Operations via Microsoft Graph API

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

CONFIG_DIR="$BASE_DIR/$ACCOUNT"
CONFIG_FILE="$CONFIG_DIR/config.json"
CREDS_FILE="$CONFIG_DIR/credentials.json"
GRAPH_URL="https://graph.microsoft.com/v1.0"

# Timezone: OUTLOOK_TZ override, else system timezone, else Europe/London fallback
if [ -n "$OUTLOOK_TZ" ]; then
    DEFAULT_TIMEZONE="$OUTLOOK_TZ"
elif [ -f /etc/timezone ]; then
    DEFAULT_TIMEZONE=$(cat /etc/timezone)
elif command -v timedatectl &>/dev/null; then
    DEFAULT_TIMEZONE=$(timedatectl show -p Timezone --value 2>/dev/null)
elif [ -L /etc/localtime ]; then
    DEFAULT_TIMEZONE=$(readlink /etc/localtime | sed 's|.*/zoneinfo/||')
fi
[ -z "$DEFAULT_TIMEZONE" ] && DEFAULT_TIMEZONE="Europe/London"

# Check credentials
if [ ! -f "$CREDS_FILE" ]; then
    echo "Error: Account '$ACCOUNT' not configured. Run: outlook-setup.sh --account $ACCOUNT"
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
        echo "Error: No refresh token. Run outlook-setup.sh to re-authenticate." >&2
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
    echo "Error: Invalid access token. Run outlook-setup.sh to re-authenticate."
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
            -H "Prefer: outlook.timezone=\"$DEFAULT_TIMEZONE\"" \
            -d "$data"
    else
        curl -s -X "$method" "${GRAPH_URL}${endpoint}" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Prefer: outlook.timezone=\"$DEFAULT_TIMEZONE\""
    fi
}

# API call helper. Transparently refreshes the token and retries once if Graph
# rejects it mid-run (revoked, clock skew, or a race with local expiry). A
# transport failure (curl non-zero) is turned into a JSON error so callers can
# surface it — but a legitimately empty body (HTTP 204/202 from DELETE/send) is
# left empty, since callers treat "no .error" as success.
api_call() {
    local response rc
    response=$(_graph_request "$@") && rc=0 || rc=$?

    if [ -z "${OUTLOOK_TOKEN_RETRIED:-}" ] && \
       printf '%s' "$response" | jq -e 'objects | .error.code == "InvalidAuthenticationToken"' >/dev/null 2>&1; then
        OUTLOOK_TOKEN_RETRIED=1
        ACCESS_TOKEN=$(refresh_access_token) || true
        response=$(_graph_request "$@") && rc=0 || rc=$?
    fi

    if [ "$rc" -ne 0 ] && [ -z "$response" ]; then
        response='{"error":{"code":"NetworkError","message":"Request to Microsoft Graph failed (network error, timeout, or connectivity issue)."}}'
    fi
    printf '%s' "$response"
}

# Format event for display
format_event() {
    jq -r '
        def short_id: .[-20:];
        def format_time: split("T")[1] | split(":")[0:2] | join(":");
        "\(.start.dateTime | split("T")[0]) \(.start.dateTime | format_time)-\(.end.dateTime | format_time) | \(.subject // "(no subject)") | \(.location.displayName // "-") | [\(.id | short_id)]"
    '
}

# Format events list
format_events() {
    jq -r '
        def short_id: .[-20:];
        def format_time: split("T")[1] | split(":")[0:2] | join(":");
        .value | to_entries | .[] |
        "[\(.key + 1)] \(.value.start.dateTime | split("T")[0]) \(.value.start.dateTime | format_time)-\(.value.end.dateTime | format_time) | \(.value.subject // "(no subject)") | \(.value.location.displayName // "-")"
    '
}

# Get today's date in ISO format
today_start() {
    date -u +"%Y-%m-%dT00:00:00Z"
}

today_end() {
    date -u +"%Y-%m-%dT23:59:59Z"
}

week_end() {
    date -u -d "+7 days" +"%Y-%m-%dT23:59:59Z" 2>/dev/null || date -u -v+7d +"%Y-%m-%dT23:59:59Z"
}

# Commands
case "$1" in
    events)
        count="${2:-10}"
        echo "Upcoming events ($count)..."
        start=$(today_start)
        api_call GET "/me/calendar/events?\$filter=start/dateTime%20ge%20'$start'&\$top=$count&\$orderby=start/dateTime&\$select=id,subject,start,end,location,organizer,attendees" | format_events
        ;;

    today)
        echo "Today's events..."
        start=$(today_start)
        end=$(today_end)
        api_call GET "/me/calendar/calendarView?startDateTime=$start&endDateTime=$end&\$orderby=start/dateTime&\$select=id,subject,start,end,location" | format_events
        ;;

    week)
        echo "This week's events..."
        start=$(today_start)
        end=$(week_end)
        api_call GET "/me/calendar/calendarView?startDateTime=$start&endDateTime=$end&\$orderby=start/dateTime&\$select=id,subject,start,end,location" | format_events
        ;;

    read)
        event_id="$2"
        if [ -z "$event_id" ]; then
            echo "Usage: outlook-calendar.sh read <event-id>"
            exit 1
        fi

        # Find full ID if short ID provided
        if [ ${#event_id} -le 25 ]; then
            start=$(today_start)
            full_id=$(api_call GET "/me/calendar/events?\$filter=start/dateTime%20ge%20'$start'&\$top=100&\$select=id" | jq -r ".value[].id | select(endswith(\"$event_id\"))" | head -1)
            if [ -z "$full_id" ]; then
                echo "Error: Event not found with ID ending in: $event_id"
                exit 1
            fi
            event_id="$full_id"
        fi

        echo "Event details..."
        api_call GET "/me/calendar/events/$event_id" | jq -r '
            "Subject: \(.subject // "(no subject)")",
            "Start: \(.start.dateTime) (\(.start.timeZone))",
            "End: \(.end.dateTime) (\(.end.timeZone))",
            "Location: \(.location.displayName // "-")",
            "Organizer: \(.organizer.emailAddress.name // "") <\(.organizer.emailAddress.address // "")>",
            "Attendees: \([.attendees[]?.emailAddress | "\(.name // "") <\(.address)>"] | join(", ") | if . == "" then "-" else . end)",
            "Response: \(.responseStatus.response // "-")",
            "---",
            "Body: \(.body.content | gsub("<[^>]*>"; "") | gsub("&nbsp;"; " ") | gsub("\\s+"; " ") | ltrimstr(" ") | if . == "" then "(no description)" else . end)"
        '
        ;;

    calendars)
        echo "Available calendars..."
        api_call GET "/me/calendars?\$select=id,name,color,isDefaultCalendar" | jq -r '
            .value[] | "[\(if .isDefaultCalendar then "*" else " " end)] \(.name) (\(.color // "auto"))"
        '
        ;;

    create)
        subject="$2"
        start_time="$3"
        end_time="$4"
        location="${5:-}"

        if [ -z "$subject" ] || [ -z "$start_time" ] || [ -z "$end_time" ]; then
            echo "Usage: outlook-calendar.sh create <subject> <start-time> <end-time> [location]"
            echo "Times in format: YYYY-MM-DDTHH:MM"
            exit 1
        fi

        echo "Creating event..."
        payload=$(jq -n \
            --arg subject "$subject" \
            --arg start "$start_time" \
            --arg end "$end_time" \
            --arg location "$location" \
            --arg tz "$DEFAULT_TIMEZONE" \
            '{
                subject: $subject,
                start: {
                    dateTime: $start,
                    timeZone: $tz
                },
                end: {
                    dateTime: $end,
                    timeZone: $tz
                }
            } + (if $location != "" then {location: {displayName: $location}} else {} end)')

        result=$(api_call POST "/me/calendar/events" "$payload")
        event_id=$(echo "$result" | jq -r '.id')

        if [ -z "$event_id" ] || [ "$event_id" = "null" ]; then
            echo "Error creating event:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        echo "Event created!"
        echo "Event ID: ${event_id: -20}"
        echo
        echo "$result" | jq -r '"Subject: \(.subject)", "Start: \(.start.dateTime)", "End: \(.end.dateTime)", "Location: \(.location.displayName // "-")"'
        ;;

    quick)
        subject="$2"
        start_time="$3"

        if [ -z "$subject" ] || [ -z "$start_time" ]; then
            echo "Usage: outlook-calendar.sh quick <subject> <start-time>"
            echo "Creates a 1-hour event. Time format: YYYY-MM-DDTHH:MM"
            exit 1
        fi

        # Calculate end time (1 hour later)
        if command -v gdate &> /dev/null; then
            end_time=$(gdate -d "$start_time + 1 hour" +"%Y-%m-%dT%H:%M")
        else
            end_time=$(date -d "$start_time + 1 hour" +"%Y-%m-%dT%H:%M" 2>/dev/null || date -j -f "%Y-%m-%dT%H:%M" -v+1H "$start_time" +"%Y-%m-%dT%H:%M")
        fi

        echo "Creating 1-hour event..."
        payload=$(jq -n \
            --arg subject "$subject" \
            --arg start "$start_time" \
            --arg end "$end_time" \
            --arg tz "$DEFAULT_TIMEZONE" \
            '{
                subject: $subject,
                start: {
                    dateTime: $start,
                    timeZone: $tz
                },
                end: {
                    dateTime: $end,
                    timeZone: $tz
                }
            }')

        result=$(api_call POST "/me/calendar/events" "$payload")
        event_id=$(echo "$result" | jq -r '.id')

        if [ -z "$event_id" ] || [ "$event_id" = "null" ]; then
            echo "Error creating event:"
            echo "$result" | jq -r '.error.message // .'
            exit 1
        fi

        echo "Event created!"
        echo "Event ID: ${event_id: -20}"
        echo
        echo "$result" | jq -r '"Subject: \(.subject)", "Start: \(.start.dateTime)", "End: \(.end.dateTime)"'
        ;;

    update)
        event_id="$2"
        field="$3"
        value="$4"

        if [ -z "$event_id" ] || [ -z "$field" ] || [ -z "$value" ]; then
            echo "Usage: outlook-calendar.sh update <event-id> <field> <value>"
            echo "Fields: subject, location, start, end"
            exit 1
        fi

        # Find full ID if short ID provided
        if [ ${#event_id} -le 25 ]; then
            start=$(today_start)
            full_id=$(api_call GET "/me/calendar/events?\$filter=start/dateTime%20ge%20'$start'&\$top=100&\$select=id" | jq -r ".value[].id | select(endswith(\"$event_id\"))" | head -1)
            if [ -z "$full_id" ]; then
                echo "Error: Event not found"
                exit 1
            fi
            event_id="$full_id"
        fi

        echo "Updating event..."
        case "$field" in
            subject)
                payload=$(jq -n --arg v "$value" '{subject: $v}')
                ;;
            location)
                payload=$(jq -n --arg v "$value" '{location: {displayName: $v}}')
                ;;
            start)
                payload=$(jq -n --arg v "$value" --arg tz "$DEFAULT_TIMEZONE" '{start: {dateTime: $v, timeZone: $tz}}')
                ;;
            end)
                payload=$(jq -n --arg v "$value" --arg tz "$DEFAULT_TIMEZONE" '{end: {dateTime: $v, timeZone: $tz}}')
                ;;
            *)
                echo "Unknown field: $field"
                echo "Valid fields: subject, location, start, end"
                exit 1
                ;;
        esac

        api_call PATCH "/me/calendar/events/$event_id" "$payload" > /dev/null
        echo "Event updated"
        ;;

    delete)
        event_id="$2"
        if [ -z "$event_id" ]; then
            echo "Usage: outlook-calendar.sh delete <event-id>"
            exit 1
        fi

        # Find full ID if short ID provided
        if [ ${#event_id} -le 25 ]; then
            start=$(today_start)
            full_id=$(api_call GET "/me/calendar/events?\$filter=start/dateTime%20ge%20'$start'&\$top=100&\$select=id" | jq -r ".value[].id | select(endswith(\"$event_id\"))" | head -1)
            if [ -z "$full_id" ]; then
                echo "Error: Event not found"
                exit 1
            fi
            event_id="$full_id"
        fi

        api_call DELETE "/me/calendar/events/$event_id" > /dev/null
        echo "Event deleted"
        ;;

    free)
        start_time="$2"
        end_time="$3"

        if [ -z "$start_time" ] || [ -z "$end_time" ]; then
            echo "Usage: outlook-calendar.sh free <start-time> <end-time>"
            echo "Times in format: YYYY-MM-DDTHH:MM"
            exit 1
        fi

        echo "Checking availability from $start_time to $end_time..."

        # Get events in range
        events=$(api_call GET "/me/calendar/calendarView?startDateTime=${start_time}:00&endDateTime=${end_time}:00&\$orderby=start/dateTime&\$select=subject,start,end")

        event_count=$(echo "$events" | jq '.value | length')

        if [ "$event_count" -eq 0 ]; then
            echo "You are FREE during this time period."
        else
            echo "You have $event_count event(s) during this period:"
            echo "$events" | jq -r '
                def format_time: split("T")[1] | split(":")[0:2] | join(":");
                .value[] | "  \(.start.dateTime | format_time)-\(.end.dateTime | format_time): \(.subject)"
            '
        fi
        ;;

    *)
        echo "Outlook Calendar Operations"
        echo
        echo "Usage: outlook-calendar.sh <command> [args]"
        echo
        echo "Viewing:"
        echo "  events [count]             List upcoming events"
        echo "  today                      Today's events"
        echo "  week                       This week's events"
        echo "  read <id>                  Event details"
        echo "  calendars                  List calendars"
        echo
        echo "Creating:"
        echo "  create <subject> <start> <end> [location]"
        echo "                             Create event"
        echo "  quick <subject> <start>    Create 1-hour event"
        echo
        echo "Managing:"
        echo "  update <id> <field> <value>"
        echo "                             Update event field"
        echo "  delete <id>                Delete event"
        echo
        echo "Availability:"
        echo "  free <start> <end>         Check free/busy"
        echo
        echo "Times in format: YYYY-MM-DDTHH:MM"
        ;;
esac
