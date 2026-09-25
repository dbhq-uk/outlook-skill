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
EVENT_ID_CACHE_FILE="$CONFIG_DIR/event_id_cache.json"
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

# Read-only mode: with OUTLOOK_READ_ONLY set, only these verbs run. Every
# other verb refuses here, before a token is read or a request is made. See
# outlook_read_only_gate in lib/graph.sh. A new verb that only reads belongs in
# this list; one that writes or sends must stay out of it.
READ_ONLY_VERBS=(events today week read calendars day search free)
outlook_read_only_gate outlook-calendar.sh "${1:-}" "${READ_ONLY_VERBS[@]}" || exit 1

# Check credentials
if [ ! -f "$CREDS_FILE" ]; then
    echo "Error: Account '$ACCOUNT' not configured. Run: outlook-setup.sh --account $ACCOUNT"
    exit 1
fi

# A failed refresh has already said why on stderr, and left credentials.json
# as it was.
if ! ACCESS_TOKEN=$(ensure_valid_token) || [ -z "$ACCESS_TOKEN" ]; then
    exit 1
fi

# Low-level Graph request using the current $ACCESS_TOKEN. outlook_curl (in
# lib/graph.sh) adds the timeouts and retries a throttled request.
_graph_request() {
    local method="$1"
    local endpoint="$2"
    local data="$3"

    if [ -n "$data" ]; then
        outlook_curl "$method" -X "$method" "${GRAPH_URL}${endpoint}" \
            -H "Authorization: Bearer $ACCESS_TOKEN" \
            -H "Content-Type: application/json" \
            -H "Prefer: outlook.timezone=\"$DEFAULT_TIMEZONE\"" \
            -d "$data"
    else
        outlook_curl "$method" -X "$method" "${GRAPH_URL}${endpoint}" \
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

# Format an event listing ({"value":[...]} from calendar_view). Every row carries
# the event's short ID - the last 20 characters, as in the mail listings - so
# read, update, respond, cancel and delete can act on it. Two events whose IDs
# end the same way (occurrences of one series can) are printed with their full
# ID instead, so a short ID on screen never resolves to the wrong event.
format_events() {
    jq -r '
        def short_id: .[-20:];
        def format_time: split("T")[1] | split(":")[0:2] | join(":");
        if .error then
            "Error: \(.error.message // .error.code // "Unknown API error")"
        elif (.value | length) == 0 then
            "No events found."
        else
            ([.value[].id | short_id] | group_by(.) | map(select(length > 1) | .[0])) as $clash
            | (.value | to_entries | .[] |
                (.value.id | short_id) as $s
                | "[\(.key + 1)] \(if ($clash | index($s)) then .value.id else $s end) | \(.value.start.dateTime | split("T")[0]) \(.value.start.dateTime | format_time)-\(.value.end.dateTime | format_time) | \(.value.subject // "(no subject)") | \(.value.location.displayName // "-")\(if .value.isCancelled then " | cancelled" else "" end)"),
              (if .more then "Stopped at \(.value | length) events. There are more in this window; narrow it." else empty end)
        end
    '
}

# The fields every listing asks for. showAs and isCancelled are what `free`
# needs to tell a real clash from a placeholder.
EVENT_SELECT="id,subject,start,end,location,showAs,isCancelled,type"

# Every event in a calendarView window, oldest first, following
# @odata.nextLink. calendarView expands a recurring series into its
# occurrences, which /me/calendar/events does not. Graph's default page is 10
# events, and without paging a busy week showed 10 and said nothing of the
# rest. $1/$2 are URL-encoded bounds; $3 stops after that many (default
# CALENDAR_VIEW_MAX). Prints {"value":[...], "more": bool} or the Graph error.
CALENDAR_VIEW_MAX=1000
calendar_view() {
    local start="$1" end="$2" max="${3:-$CALENDAR_VIEW_MAX}" page_size url merged page next collected more=false
    page_size=100
    [ "$max" -lt "$page_size" ] && page_size="$max"
    url="/me/calendar/calendarView?startDateTime=$start&endDateTime=$end&\$orderby=start/dateTime&\$top=$page_size&\$select=$EVENT_SELECT"
    merged='[]'
    while [ -n "$url" ]; do
        page=$(api_call GET "$url")
        if ! printf '%s' "$page" | jq -e 'type == "object" and has("value")' >/dev/null 2>&1; then
            # An error object passes through for the caller to print; anything
            # else (an empty or garbled body) becomes one.
            if printf '%s' "$page" | jq -e '.error' >/dev/null 2>&1; then
                printf '%s' "$page"
            else
                printf '%s' '{"error":{"code":"BadResponse","message":"Graph returned no event list."}}'
            fi
            return 0
        fi
        merged=$(jq -n --argjson a "$merged" --argjson b "$(printf '%s' "$page" | jq '.value')" '$a + $b')
        next=$(printf '%s' "$page" | jq -r '."@odata.nextLink" // empty')
        collected=$(printf '%s' "$merged" | jq 'length')
        if [ "$collected" -ge "$max" ]; then
            if [ -n "$next" ] || [ "$collected" -gt "$max" ]; then
                more=true
            fi
            break
        fi
        url="${next#"$GRAPH_URL"}"    # nextLink is absolute; strip base for api_call
    done
    printf '%s' "$merged" | jq --argjson max "$max" --argjson more "$more" '{value: .[0:$max], more: $more}'
}

# Remember the full IDs a listing printed, so a short ID copied off it resolves
# without another search. Written to a temp file and renamed into place, so a
# concurrent reader never sees half a file.
cache_event_ids() {
    local response="$1" tmp
    tmp=$(mktemp "$CONFIG_DIR/.event_id_cache.XXXXXX" 2>/dev/null) || return 0
    if printf '%s' "$response" | jq -c '[.value[]?.id // empty]' > "$tmp" 2>/dev/null; then
        mv -f "$tmp" "$EVENT_ID_CACHE_FILE" 2>/dev/null || rm -f "$tmp"
    else
        rm -f "$tmp"
    fi
}

# List a window, cache its IDs, print it. Returns 1 on a Graph error, after
# printing it, so the command exits non-zero.
list_window() {
    local result
    result=$(calendar_view "$1" "$2")
    printf '%s' "$result" | format_events
    if printf '%s' "$result" | jq -e '.error' >/dev/null 2>&1; then
        return 1
    fi
    cache_event_ids "$result"
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

# Resolve a short (20-char) event ID to its full ID. Full-length IDs pass
# through untouched. Looks first at the IDs the last listing printed, then at
# every occurrence from 30 days back to a year ahead (calendarView, so a single
# occurrence of a recurring meeting is found), then at the 250 most recent
# events and series. A short ID that matches two different events is refused
# rather than guessed. Errors go to stderr; returns 1 on a miss or a clash.
resolve_event_id() {
    local event_id="$1" matches n
    if [ ${#event_id} -gt 25 ]; then
        printf '%s' "$event_id"
        return 0
    fi
    matches=$(jq -r --arg s "$event_id" '.[]? | select(endswith($s))' "$EVENT_ID_CACHE_FILE" 2>/dev/null | sort -u)
    if [ -z "$matches" ]; then
        matches=$( {
            calendar_view "$(day_start "$(local_date '-30 days')")" "$(day_end "$(local_date '+365 days')")" \
                | jq -r '.value[]?.id'
            api_call GET "/me/events?\$top=250&\$orderby=start/dateTime%20desc&\$select=id" \
                | jq -r '.value[]?.id'
        } 2>/dev/null | jq -Rr --arg s "$event_id" 'select(endswith($s))' | sort -u)
    fi
    n=$(printf '%s' "$matches" | grep -c . || true)
    if [ "$n" -eq 0 ]; then
        echo "Error: no event found with an ID ending in: $event_id" >&2
        return 1
    fi
    if [ "$n" -gt 1 ]; then
        echo "Error: $n events have an ID ending in $event_id. Use the full ID from the listing." >&2
        return 1
    fi
    printf '%s' "$matches"
}

# If the event is a meeting the user organises and has attendees, print the
# attendees' addresses, comma-separated. Prints nothing for the user's own
# event or for someone else's meeting, where a change reaches nobody else.
# Returns 1, with the reason on stderr, when the event cannot be read, so a
# caller never takes "could not tell" for "no attendees".
organised_meeting_attendees() {
    local ev
    ev=$(api_call GET "/me/events/$1?\$select=isOrganizer,attendees")
    if ! printf '%s' "$ev" | jq -e 'type == "object" and (has("error") | not) and has("isOrganizer")' >/dev/null 2>&1; then
        echo "Error: could not read the event to check who it would notify:" >&2
        printf '%s' "$ev" | jq -r '.error.message // .error.code // "no response"' >&2 2>/dev/null || echo "no response" >&2
        return 1
    fi
    printf '%s' "$ev" | jq -r 'if .isOrganizer then [.attendees[]?.emailAddress.address // empty] | join(", ") else "" end'
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
        [[ "$count" =~ ^[0-9]+$ ]] && [ "$count" -ge 1 ] || count=10
        echo "Upcoming events ($count)..."
        # The next $count events in the coming year, each occurrence of a
        # recurring meeting on its own row. Asking for $count is the point, so
        # there is no "more" note.
        result=$(calendar_view "$(today_start)" "$(day_end "$(local_date '+365 days')")" "$count" | jq '.more = false')
        printf '%s' "$result" | format_events
        if printf '%s' "$result" | jq -e '.error' >/dev/null 2>&1; then
            exit 1
        fi
        cache_event_ids "$result"
        ;;

    today)
        echo "Today's events ($DEFAULT_TIMEZONE)..."
        list_window "$(today_start)" "$(today_end)"
        ;;

    week)
        echo "This week's events ($DEFAULT_TIMEZONE)..."
        list_window "$(today_start)" "$(week_end)"
        ;;

    read)
        event_id="$2"
        if [ -z "$event_id" ]; then
            echo "Usage: outlook-calendar.sh read <event-id>"
            exit 1
        fi

        if ! event_id=$(resolve_event_id "$event_id"); then
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
        # --send-invites may sit anywhere after the verb; the rest are positional.
        send_invites=0
        pos=()
        for arg in "${@:2}"; do
            if [ "$arg" = "--send-invites" ]; then send_invites=1; else pos+=("$arg"); fi
        done
        subject="${pos[0]:-}"
        start_time="${pos[1]:-}"
        end_time="${pos[2]:-}"
        location="${pos[3]:-}"
        attendees="${pos[4]:-}"

        if [ -z "$subject" ] || [ -z "$start_time" ] || [ -z "$end_time" ]; then
            echo "Usage: outlook-calendar.sh create <subject> <start-time> <end-time> [location] [attendees --send-invites]"
            echo "Times in format: YYYY-MM-DDTHH:MM"
            echo "Without attendees nothing is sent. Attendees are comma/semicolon-separated"
            echo "emails (pass \"\" for location if there is none), and they need --send-invites,"
            echo "because the invitations go out as soon as the event exists."
            exit 1
        fi

        attendees_json=$(attendees_to_json "$attendees")

        # A one-shot invite has to be asked for by name. Without the flag the
        # attendee list is refused before anything reaches Graph, so an agent
        # that passes attendees by habit gets an error rather than a sent invite.
        if [ "$(printf '%s' "$attendees_json" | jq 'length')" -gt 0 ] && [ "$send_invites" -ne 1 ]; then
            echo "Refused: create with attendees sends the invitations as soon as the event exists." >&2
            echo "  Create it without attendees, confirm the details, then run: invite <event-id> <emails>" >&2
            echo "  Or, if the exact attendee list is already approved, add --send-invites." >&2
            echo "Nothing was created and nothing was sent." >&2
            exit 1
        fi

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
        # --notify-attendees may sit anywhere after the verb.
        notify=0
        pos=()
        for arg in "${@:2}"; do
            if [ "$arg" = "--notify-attendees" ]; then notify=1; else pos+=("$arg"); fi
        done
        event_id="${pos[0]:-}"
        field="${pos[1]:-}"
        value="${pos[2]:-}"

        if [ -z "$event_id" ] || [ -z "$field" ] || [ -z "$value" ]; then
            echo "Usage: outlook-calendar.sh update <event-id> <field> <value> [--notify-attendees]"
            echo "Fields: subject, location, start, end"
            echo "Changing a meeting you organise sends every attendee an update, so it"
            echo "needs --notify-attendees. Your own events and other people's meetings do not."
            exit 1
        fi

        if ! event_id=$(resolve_event_id "$event_id"); then
            exit 1
        fi

        # Graph sends a meeting update to the attendees when the organiser
        # changes a meeting, so this is a send, and it has to be asked for.
        if ! attendees=$(organised_meeting_attendees "$event_id"); then
            exit 1
        fi
        if [ -n "$attendees" ] && [ "$notify" -ne 1 ]; then
            echo "Refused: this is a meeting you organise, and changing it sends an update to: $attendees" >&2
            echo "  Confirm the change with the user, then run it again with --notify-attendees." >&2
            echo "Nothing was changed and nothing was sent." >&2
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
            exit 1
        fi

        # Deleting a meeting on the organiser's calendar sends the attendees a
        # cancellation (Graph's documented behaviour), so it is not the silent
        # delete it looks like. `cancel` does the same thing and says so.
        if ! attendees=$(organised_meeting_attendees "$event_id"); then
            exit 1
        fi
        if [ -n "$attendees" ]; then
            echo "Refused: this is a meeting you organise, and deleting it sends a cancellation to: $attendees" >&2
            echo "  To cancel it and tell them, confirm with the user, then run: cancel <event-id> [comment]" >&2
            echo "Nothing was deleted and nothing was sent." >&2
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
        list_window "$(day_start "$day")" "$(day_end "$day")"
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
        matches=$(calendar_view "$start" "$end" \
            | jq --arg q "$query" 'if .error then . else {more, value: [.value[] | select(
                ((.subject // "") + " " + (.location.displayName // "")) | ascii_downcase | contains($q | ascii_downcase)
              )]} end')
        if printf '%s' "$matches" | jq -e '.error' >/dev/null 2>&1; then
            printf '%s' "$matches" | format_events
            exit 1
        elif [ "$(printf '%s' "$matches" | jq '.value | length')" -eq 0 ]; then
            echo "No matching events found."
        else
            cache_event_ids "$matches"
            printf '%s' "$matches" | format_events
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

        # Every event in the window, then drop the ones that do not block time:
        # an event shown as free (a reminder, an all-day marker) and one that
        # has been cancelled but still sits in the calendar. Counting those
        # would report a clash that is not there.
        events=$(calendar_view "$win_start" "$win_end")
        if printf '%s' "$events" | jq -e '.error' >/dev/null 2>&1; then
            printf '%s' "$events" | format_events
            exit 1
        fi
        busy=$(printf '%s' "$events" | jq '{value: [.value[] | select((.showAs // "busy") != "free" and (.isCancelled | not))]}')
        cache_event_ids "$busy"
        event_count=$(printf '%s' "$busy" | jq '.value | length')

        if [ "$event_count" -eq 0 ]; then
            echo "You are FREE during this time period."
        else
            echo "You have $event_count event(s) during this period:"
            printf '%s' "$busy" | format_events
        fi
        ;;

    *)
        echo "Outlook Calendar Operations"
        echo
        echo "Usage: outlook-calendar.sh <command> [args]"
        echo
        echo "Viewing (every listing prints a short ID the other commands accept):"
        echo "  events [count]             Next events, each occurrence of a series listed"
        echo "  today                      Today's events"
        echo "  week                       This week's events"
        echo "  day <YYYY-MM-DD>           Events on a specific date"
        echo "  search <text> [days]       Find events by subject/location (default: next 90 days)"
        echo "  read <id>                  Event details"
        echo "  calendars                  List calendars"
        echo
        echo "Creating:"
        echo "  create <subject> <start> <end> [location] [attendees --send-invites]"
        echo "                             Create event. Without attendees nothing is sent"
        echo "                             (two-step: create, confirm, then 'invite')."
        echo "                             Attendees need --send-invites: they are invited at once."
        echo "  invite <id> <emails> [required|optional]"
        echo "                             Add attendees to an event and send invitations"
        echo "  quick <subject> <start>    Create 1-hour event"
        echo
        echo "Managing:"
        echo "  update <id> <field> <value> [--notify-attendees]"
        echo "                             Update event field (subject/location/start/end)."
        echo "                             A meeting you organise needs --notify-attendees."
        echo "  respond <id> <accept|decline|tentative> [comment]"
        echo "                             Respond to a meeting invitation"
        echo "  cancel <id> [comment]      Cancel a meeting you organise (notifies attendees)"
        echo "  delete <id>                Delete an event that notifies nobody. A meeting you"
        echo "                             organise is refused: use cancel, which tells attendees"
        echo
        echo "Availability:"
        echo "  free <start> <end>         Check free/busy"
        echo
        echo "Times in format: YYYY-MM-DDTHH:MM"
        echo
        echo "OUTLOOK_READ_ONLY=1 refuses every command above that is not a viewing command."
        ;;
esac
