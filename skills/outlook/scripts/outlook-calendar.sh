#!/bin/bash
# Outlook Calendar Operations via Microsoft Graph API

set -e

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
# shellcheck disable=SC2034 # read by lib/graph.sh
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

# Every time shown or accepted by this script is wall-clock in $DEFAULT_TIMEZONE.
# Servers, containers and CI boxes almost always report UTC, while the mailbox
# owner lives somewhere else - and then "your 13:00 interview" is really at 14:00
# for them. Same instant, wrong wall-clock, missed meeting. We cannot read the
# mailbox's own timezone (that needs a MailboxSettings.Read scope this app
# deliberately does not request), so say so loudly rather than be quietly wrong.
case "$DEFAULT_TIMEZONE" in
    UTC|Etc/UTC|GMT|Etc/GMT|Universal)
        if [ -z "$OUTLOOK_TZ" ]; then
            echo "Note: times are in $DEFAULT_TIMEZONE (the system timezone). If your Outlook" >&2
            echo "      calendar is in another zone, set OUTLOOK_TZ (e.g. OUTLOOK_TZ=Europe/London)" >&2
            echo "      or every time below may be off by your UTC offset." >&2
        fi
        ;;
esac

# Check credentials
if [ ! -f "$CREDS_FILE" ]; then
    echo "Error: Account '$ACCOUNT' not configured. Run: outlook-setup.sh --account $ACCOUNT"
    exit 1
fi

# --- Token management -------------------------------------------------------
# The token code lives in lib/graph.sh, shared by every script, so a fix to it
# lands once. The access token is resolved from a locally-stored absolute expiry
# (expires_at), so the common path makes NO network pre-flight call. The token
# is refreshed over the network only when it is missing/expired, or when Graph
# rejects it mid-run (handled reactively in api_call).
#
# The library is found beside this script's real location, following symlinks,
# so it resolves however the script was reached (plugin, symlinked skill
# directory, or a symlink to the script itself).
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

# A failed refresh has already said why on stderr, and left credentials.json
# as it was.
if ! ACCESS_TOKEN=$(ensure_valid_token) || [ -z "$ACCESS_TOKEN" ]; then
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
    local response rc new_token
    response=$(_graph_request "$@") && rc=0 || rc=$?

    if [ -z "${OUTLOOK_TOKEN_RETRIED:-}" ] && \
       printf '%s' "$response" | jq -e 'objects | .error.code == "InvalidAuthenticationToken"' >/dev/null 2>&1; then
        OUTLOOK_TOKEN_RETRIED=1
        # Retry only with a new token. A failed refresh has explained itself on
        # stderr and left credentials.json alone; the original error stands.
        if new_token=$(refresh_access_token); then
            ACCESS_TOKEN="$new_token"
            response=$(_graph_request "$@") && rc=0 || rc=$?
        fi
    fi

    if [ "$rc" -ne 0 ] && [ -z "$response" ]; then
        response='{"error":{"code":"NetworkError","message":"Request to Microsoft Graph failed (network error, timeout, or connectivity issue)."}}'
    fi
    printf '%s' "$response"
}

# --- HTML -> readable plain text ---------------------------------------------
# Injected into the `read` filter. Block-level tags become line breaks BEFORE
# tags are stripped, so an event description with paragraphs, bullet points or a
# Teams join block stays readable instead of collapsing into one run-on line.
# Kept identical to the helper in outlook-mail.sh.
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

# Fail loudly on a Graph error. Several commands used to pipe their response to
# /dev/null and print "Event updated" unconditionally, so a REJECTED write (e.g.
# PATCHing a start that lands after the existing end) still reported success and
# the caller believed a change had been made that had not. Never claim a write
# succeeded without looking at what Graph said.
die_on_error() {
    local response="$1" context="$2"
    if [ -n "$response" ] && printf '%s' "$response" | jq -e '.error' > /dev/null 2>&1; then
        echo "Error $context:" >&2
        printf '%s' "$response" | jq -r '.error.message // .error.code' >&2
        exit 1
    fi
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

# --- Local-time window helpers ----------------------------------------------
# Graph interprets a calendarView startDateTime/endDateTime that carries no UTC
# offset as UTC. The user thinks in LOCAL time ("today", "am I free 9-5?"), so a
# naive UTC window is wrong wherever local time != UTC: in BST (UTC+1) a query
# for "the 16th" spans 16th 00:00Z-23:59Z, which is 01:00 on the 16th to 00:59
# on the 17th locally - so it misses the first hour of the day and picks up an
# all-day event that starts at local midnight on the 17th. Emitting an explicit
# offset (2026-07-16T00:00:00+01:00) removes the ambiguity entirely.

# URL-encode a query-string value. The offset in an ISO timestamp is a literal
# "+", which means SPACE in a query string - unencoded, Graph receives a broken
# datetime and the command dies. Everything below goes through this.
urlencode() { jq -rn --arg s "$1" '$s|@uri'; }

# Local wall-clock -> ISO 8601 with numeric offset, e.g. 2026-07-16T00:00:00+01:00
local_iso() {
    TZ="$DEFAULT_TIMEZONE" date -d "$1" +"%Y-%m-%dT%H:%M:%S%:z" 2>/dev/null \
        || TZ="$DEFAULT_TIMEZONE" date -j -f "%Y-%m-%d %H:%M:%S" "$1" +"%Y-%m-%dT%H:%M:%S%z"
}

# Local YYYY-MM-DD for a relative day expression ("today", "+7 days")
local_date() {
    TZ="$DEFAULT_TIMEZONE" date -d "$1" +"%Y-%m-%d" 2>/dev/null \
        || TZ="$DEFAULT_TIMEZONE" date -v"$1" +"%Y-%m-%d"
}

day_start() { urlencode "$(local_iso "$1 00:00:00")"; }
day_end()   { urlencode "$(local_iso "$1 23:59:59")"; }

today_start() { day_start "$(local_date today)"; }
today_end()   { day_end   "$(local_date today)"; }
week_end()    { day_end   "$(local_date '+7 days')"; }

# Resolve a short (20-char) event ID to its full ID. Searches upcoming events
# first, then falls back to the most recent 250 events so past events can be
# addressed too. Full-length IDs pass through untouched.
resolve_event_id() {
    local event_id="$1" start full_id
    if [ ${#event_id} -gt 25 ]; then
        printf '%s' "$event_id"
        return 0
    fi
    start=$(today_start)
    full_id=$(api_call GET "/me/calendar/events?\$filter=start/dateTime%20ge%20'$start'&\$top=100&\$select=id" \
        | jq -r ".value[].id | select(endswith(\"$event_id\"))" | head -1)
    if [ -z "$full_id" ]; then
        full_id=$(api_call GET "/me/events?\$top=250&\$orderby=start/dateTime%20desc&\$select=id" \
            | jq -r ".value[].id | select(endswith(\"$event_id\"))" | head -1)
    fi
    [ -n "$full_id" ] || return 1
    printf '%s' "$full_id"
}

# Convert a comma/semicolon-separated address list into Graph attendee objects.
# $2 = attendee type: required (default) or optional.
attendees_to_json() {
    jq -n --arg raw "$1" --arg type "${2:-required}" '
        ($raw | gsub(";"; ",") | split(",")
            | map(gsub("^\\s+|\\s+$"; ""))
            | map(select(length > 0))
            | map({emailAddress: {address: .}, type: $type}))'
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
        echo "Today's events ($DEFAULT_TIMEZONE)..."
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

        if ! event_id=$(resolve_event_id "$event_id"); then
            echo "Error: Event not found with ID ending in: $2"
            exit 1
        fi

        echo "Event details..."
        api_call GET "/me/calendar/events/$event_id" | jq -r "$HTML_TO_TEXT"'
            "Subject: \(.subject // "(no subject)")",
            "Start: \(.start.dateTime) (\(.start.timeZone))",
            "End: \(.end.dateTime) (\(.end.timeZone))",
            "Location: \(.location.displayName // "-")",
            "Organizer: \(.organizer.emailAddress.name // "") <\(.organizer.emailAddress.address // "")>",
            "Attendees: \([.attendees[]?.emailAddress | "\(.name // "") <\(.address)>"] | join(", ") | if . == "" then "-" else . end)",
            "Response: \(.responseStatus.response // "-")",
            "---",
            "Body:",
            ((.body.content // "") | html_to_text | if . == "" then "(no description)" else . end)
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
        attendees="${6:-}"

        if [ -z "$subject" ] || [ -z "$start_time" ] || [ -z "$end_time" ]; then
            echo "Usage: outlook-calendar.sh create <subject> <start-time> <end-time> [location] [attendees]"
            echo "Times in format: YYYY-MM-DDTHH:MM"
            echo "Attendees: comma/semicolon-separated emails; pass \"\" for location"
            echo "if you want attendees with no location. Invitations are sent."
            exit 1
        fi

        attendees_json=$(attendees_to_json "$attendees")

        echo "Creating event..."
        payload=$(jq -n \
            --arg subject "$subject" \
            --arg start "$start_time" \
            --arg end "$end_time" \
            --arg location "$location" \
            --arg tz "$DEFAULT_TIMEZONE" \
            --argjson attendees "$attendees_json" \
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
            }
            + (if $location != "" then {location: {displayName: $location}} else {} end)
            + (if ($attendees | length) > 0 then {attendees: $attendees} else {} end)')

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

    invite)
        event_id="$2"
        emails="$3"
        att_type="${4:-required}"
        if [ -z "$event_id" ] || [ -z "$emails" ]; then
            echo "Usage: outlook-calendar.sh invite <event-id> <emails> [required|optional]"
            echo "       Adds attendees to an existing event and SENDS them invitations."
            echo "       Two-step flow: 'create' the event first (no attendees, nothing"
            echo "       sent), confirm the details, then 'invite'. Emails are"
            echo "       comma/semicolon-separated; re-inviting an address is a no-op."
            exit 1
        fi
        case "$att_type" in
            required|optional) ;;
            *) echo "Error: attendee type must be 'required' or 'optional'"; exit 1 ;;
        esac
        if ! event_id=$(resolve_event_id "$event_id"); then
            echo "Error: Event not found"
            exit 1
        fi

        new_att=$(attendees_to_json "$emails" "$att_type")
        if [ "$(echo "$new_att" | jq 'length')" -eq 0 ]; then
            echo "Error: No valid attendee address provided"
            exit 1
        fi

        # Merge with existing attendees (deduped, case-insensitive) so a repeat
        # invite never duplicates anyone. Existing entries are stripped back to
        # emailAddress + type - read-only fields like status must not be PATCHed.
        existing=$(api_call GET "/me/events/$event_id?\$select=attendees" \
            | jq '[.attendees[]? | {emailAddress: {address: .emailAddress.address}, type: .type}]')
        payload=$(jq -n --argjson ex "$existing" --argjson new "$new_att" '
            ($ex | map(.emailAddress.address // "" | ascii_downcase)) as $have
            | {attendees: ($ex + ($new | map(select((.emailAddress.address // "" | ascii_downcase) as $a | ($have | index($a)) | not))))}')

        echo "Sending invitations..."
        result=$(api_call PATCH "/me/events/$event_id" "$payload")
        if echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error inviting attendees (only the organiser can invite):"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi
        echo "Invitations sent!"
        echo "$result" | jq -r '
            "Subject:   \(.subject)",
            "Start:     \(.start.dateTime)",
            "Attendees: \([.attendees[]?.emailAddress.address] | join(", "))"
        '
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

        if ! event_id=$(resolve_event_id "$event_id"); then
            echo "Error: Event not found"
            exit 1
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

        # Graph rejects a start later than the current end (and vice versa). When
        # moving an event, update the bound that keeps start < end first, or set
        # both. The error is now surfaced instead of being reported as success.
        result=$(api_call PATCH "/me/calendar/events/$event_id" "$payload")
        die_on_error "$result" "updating event"
        echo "Event updated"
        printf '%s' "$result" | jq -r '"  \(.subject) | \(.start.dateTime) -> \(.end.dateTime)"' 2>/dev/null || true
        ;;

    delete)
        event_id="$2"
        if [ -z "$event_id" ]; then
            echo "Usage: outlook-calendar.sh delete <event-id>"
            exit 1
        fi

        if ! event_id=$(resolve_event_id "$event_id"); then
            echo "Error: Event not found"
            exit 1
        fi

        result=$(api_call DELETE "/me/calendar/events/$event_id")
        die_on_error "$result" "deleting event"
        echo "Event deleted"
        ;;

    cancel)
        event_id="$2"
        comment="${3:-}"
        if [ -z "$event_id" ]; then
            echo "Usage: outlook-calendar.sh cancel <event-id> [comment]"
            echo "       Cancels a meeting YOU organise and notifies attendees."
            echo "       To decline someone else's invite, use: respond <id> decline"
            exit 1
        fi
        if ! event_id=$(resolve_event_id "$event_id"); then
            echo "Error: Event not found"
            exit 1
        fi
        echo "Cancelling event and notifying attendees..."
        result=$(api_call POST "/me/events/$event_id/cancel" "$(jq -n --arg c "$comment" '{comment: $c}')")
        if [ -n "$result" ] && echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error cancelling event (are you the organiser? if not, use 'respond <id> decline'):"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi
        echo "Event cancelled"
        ;;

    respond)
        event_id="$2"
        answer="$3"
        comment="${4:-}"
        if [ -z "$event_id" ] || [ -z "$answer" ]; then
            echo "Usage: outlook-calendar.sh respond <event-id> <accept|decline|tentative> [comment]"
            echo "       Responds to a meeting invitation and notifies the organiser."
            exit 1
        fi
        case "$(printf '%s' "$answer" | tr '[:upper:]' '[:lower:]')" in
            accept)    action="accept" ;;
            decline)   action="decline" ;;
            tentative) action="tentativelyAccept" ;;
            *) echo "Error: response must be accept, decline, or tentative"; exit 1 ;;
        esac
        if ! event_id=$(resolve_event_id "$event_id"); then
            echo "Error: Event not found"
            exit 1
        fi
        echo "Sending '$answer' response..."
        payload=$(jq -n --arg c "$comment" '{sendResponse: true} + (if $c != "" then {comment: $c} else {} end)')
        result=$(api_call POST "/me/events/$event_id/$action" "$payload")
        if [ -n "$result" ] && echo "$result" | jq -e '.error' > /dev/null 2>&1; then
            echo "Error responding to event:"
            echo "$result" | jq -r '.error.message'
            exit 1
        fi
        echo "Response sent: $answer"
        ;;

    day)
        day="$2"
        if [ -z "$day" ]; then
            echo "Usage: outlook-calendar.sh day <YYYY-MM-DD>"
            exit 1
        fi
        if ! [[ "$day" =~ ^[0-9]{4}-[0-9]{2}-[0-9]{2}$ ]]; then
            echo "Error: date must be in YYYY-MM-DD format"
            exit 1
        fi
        echo "Events on $day ($DEFAULT_TIMEZONE)..."
        api_call GET "/me/calendar/calendarView?startDateTime=$(day_start "$day")&endDateTime=$(day_end "$day")&\$orderby=start/dateTime&\$select=id,subject,start,end,location" | format_events
        ;;

    search)
        query="$2"
        days="${3:-90}"
        if [ -z "$query" ]; then
            echo "Usage: outlook-calendar.sh search <text> [days]"
            echo "       Case-insensitive match on subject/location over the next"
            echo "       <days> days (default 90)."
            exit 1
        fi
        [[ "$days" =~ ^[0-9]+$ ]] || days=90
        start=$(today_start)
        end=$(day_end "$(local_date "+$days days")")
        echo "Searching events for '$query' (next $days days)..."
        matches=$(api_call GET "/me/calendar/calendarView?startDateTime=$start&endDateTime=$end&\$orderby=start/dateTime&\$top=250&\$select=id,subject,start,end,location" \
            | jq --arg q "$query" '{value: [.value[]? | select(
                ((.subject // "") + " " + (.location.displayName // "")) | ascii_downcase | contains($q | ascii_downcase)
              )]}')
        if [ "$(echo "$matches" | jq '.value | length')" -eq 0 ]; then
            echo "No matching events found."
        else
            echo "$matches" | format_events
        fi
        ;;

    free)
        start_time="$2"
        end_time="$3"

        if [ -z "$start_time" ] || [ -z "$end_time" ]; then
            echo "Usage: outlook-calendar.sh free <start-time> <end-time>"
            echo "Times in format: YYYY-MM-DDTHH:MM"
            exit 1
        fi

        echo "Checking availability from $start_time to $end_time ($DEFAULT_TIMEZONE)..."

        # The user means LOCAL wall-clock time ("am I free 9-5?"). Convert both
        # bounds to an offset-qualified ISO string so Graph does not read them
        # as UTC and shift the window (an hour out in BST) - a silent wrong
        # answer here would have someone double-booked.
        win_start=$(urlencode "$(local_iso "$(printf '%s' "$start_time" | tr 'T' ' '):00")")
        win_end=$(urlencode "$(local_iso "$(printf '%s' "$end_time" | tr 'T' ' '):00")")

        # Get events in range
        events=$(api_call GET "/me/calendar/calendarView?startDateTime=${win_start}&endDateTime=${win_end}&\$orderby=start/dateTime&\$select=subject,start,end")

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
        echo "  day <YYYY-MM-DD>           Events on a specific date"
        echo "  search <text> [days]       Find events by subject/location (default: next 90 days)"
        echo "  read <id>                  Event details"
        echo "  calendars                  List calendars"
        echo
        echo "Creating:"
        echo "  create <subject> <start> <end> [location] [attendees]"
        echo "                             Create event. Without attendees nothing is sent"
        echo "                             (two-step: create, confirm, then 'invite')."
        echo "                             Passing attendees sends invitations immediately."
        echo "  invite <id> <emails> [required|optional]"
        echo "                             Add attendees to an event and send invitations"
        echo "  quick <subject> <start>    Create 1-hour event"
        echo
        echo "Managing:"
        echo "  update <id> <field> <value>"
        echo "                             Update event field (subject/location/start/end)"
        echo "  respond <id> <accept|decline|tentative> [comment]"
        echo "                             Respond to a meeting invitation"
        echo "  cancel <id> [comment]      Cancel a meeting you organise (notifies attendees)"
        echo "  delete <id>                Delete event (no notification)"
        echo
        echo "Availability:"
        echo "  free <start> <end>         Check free/busy"
        echo
        echo "Times in format: YYYY-MM-DDTHH:MM"
        ;;
esac
