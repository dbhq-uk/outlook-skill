#!/bin/bash
# Outlook Contacts Operations via Microsoft Graph API

set -e

BASE_DIR="$HOME/.dbhq/outlook"

# Account resolution: --account/-a flag wins, else OUTLOOK_ACCOUNT env, else "default"
ACCOUNT="${OUTLOOK_ACCOUNT:-default}"
if [ "$1" = "--account" ] || [ "$1" = "-a" ]; then
    [ -n "$2" ] || { echo "Error: $1 requires an account name" >&2; exit 1; }
    ACCOUNT="$2"; shift 2
fi

CONFIG_DIR="$BASE_DIR/$ACCOUNT"
# shellcheck disable=SC2034 # read by lib/graph.sh
CONFIG_FILE="$CONFIG_DIR/config.json"
CREDS_FILE="$CONFIG_DIR/credentials.json"
GRAPH_URL="https://graph.microsoft.com/v1.0"

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
# this list; one that writes must stay out of it. Nothing here sends mail.
READ_ONLY_VERBS=(list search get folders duplicates)
outlook_read_only_gate outlook-contacts.sh "${1:-}" "${READ_ONLY_VERBS[@]}" || exit 1

usage() {
    echo "Outlook Contacts Operations"
    echo
    echo "Usage: outlook-contacts.sh [--account <name>] <command> [args]"
    echo
    echo "Viewing (every listing prints a short ID the other commands accept):"
    echo "  list [--folder <name|id>] [--top N] [--sort name|company|modified]"
    echo "                             List contacts (default: 50, from the main Contacts folder)"
    echo "  search <text>              Find contacts by name, email or company, in every folder"
    echo "  get <id>                   Everything Graph holds for one contact, as JSON"
    echo "  folders                    List contact folders"
    echo "  duplicates                 Report likely duplicates (same email, name or phone). Changes nothing"
    echo
    echo "Creating and changing:"
    echo "  create [fields]            Create a contact"
    echo "  update <id> [fields]       Change the fields given; the rest stay as they are"
    echo "  delete <id>                Delete a contact"
    echo "  move <id> --folder <name|id>"
    echo "                             Copy to the folder, then delete the original. Prints the new ID"
    echo "  folder-create <name> [--parent <name|id>]"
    echo "                             Create a contact folder"
    echo "  categorise <id> --add <category> | --remove <category>"
    echo "                             Add or remove an Outlook category on a contact"
    echo
    echo "Fields for create and update:"
    echo "  --name <display name>  --given <first>  --surname <last>"
    echo "  --email <address>      (repeat for more; on update they replace the list)"
    echo "  --mobile <number>  --phone <number> (same as --mobile)"
    echo "  --business-phone <number>  --home-phone <number>   (repeatable)"
    echo "  --other-phone <number> (Outlook's \"Other\" phone; on update \"\" clears it)"
    echo "  --company <text>  --title <text>  --notes <text>"
    echo "  --street <text>  --town <text>  --postcode <text>  --country <text>"
    echo "                         (the business address; on update, give every part to keep)"
    echo "  --category <name>      (repeatable; on update they replace the list)"
    echo "  --folder <name|id>     (create only: the folder to create it in)"
    echo "  On update, an empty value (--company \"\") clears that field."
    echo
    echo "Needs the Contacts.ReadWrite permission. An account set up before contacts"
    echo "existed signs in once more: outlook-setup.sh --account <name>"
    echo
    echo "OUTLOOK_READ_ONLY=1 refuses every command above that is not a viewing command."
}

case "${1:-}" in ""|help|-h|--help) usage; exit 0 ;; esac

# Check credentials
if [ ! -f "$CREDS_FILE" ]; then
    echo "Error: Account '$ACCOUNT' not configured. Run: outlook-setup.sh --account $ACCOUNT" >&2
    exit 1
fi

# Contacts.ReadWrite was added after the first release, so an account set up
# before then has not consented to it. Say so before calling Graph, which would
# answer with an opaque access-denied.
if ! outlook_scope_granted Contacts.ReadWrite; then
    echo "Error: account '$ACCOUNT' has not granted access to contacts (Contacts.ReadWrite)." >&2
    echo "Mail and calendar are not affected. Sign in once more to grant it:" >&2
    echo "  outlook-setup.sh --account $ACCOUNT" >&2
    echo "(answer y when it asks whether to overwrite the existing configuration)" >&2
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
            -d "$data"
    else
        outlook_curl "$method" -X "$method" "${GRAPH_URL}${endpoint}" \
            -H "Authorization: Bearer $ACCESS_TOKEN"
    fi
}

# API call helper. Transparently refreshes the token and retries once if Graph
# rejects it mid-run. A transport failure (curl non-zero) is turned into a JSON
# error so callers can surface it, but a legitimately empty body (HTTP 204 from
# DELETE) is left empty, since callers treat "no .error" as success.
api_call() {
    local response rc new_token
    response=$(_graph_request "$@") && rc=0 || rc=$?

    if [ -z "${OUTLOOK_TOKEN_RETRIED:-}" ] && \
       printf '%s' "$response" | jq -e 'objects | .error.code == "InvalidAuthenticationToken"' >/dev/null 2>&1; then
        OUTLOOK_TOKEN_RETRIED=1
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

# Fail loudly on a Graph error. Never claim a write succeeded without looking
# at what Graph said.
die_on_error() {
    local response="$1" context="$2"
    if [ -n "$response" ] && printf '%s' "$response" | jq -e '.error' > /dev/null 2>&1; then
        echo "Error $context:" >&2
        printf '%s' "$response" | jq -r '.error.message // .error.code' >&2
        exit 1
    fi
}

urlencode() { jq -rn --arg s "$1" '$s|@uri'; }

# The fields every listing asks for.
CONTACT_SELECT="id,displayName,givenName,surname,emailAddresses,mobilePhone,businessPhones,homePhones,companyName,jobTitle,categories,parentFolderId,lastModifiedDateTime"

# Outlook's "Other" phone is not a field of Graph's contact resource. It is the
# MAPI property PidTagOtherTelephoneNumber (String 0x3A1F), so every read of
# contacts expands it, and normalise_other turns it into a plain otherPhone.
OTHER_PROP_ID="String 0x3A1F"
OTHER_EXPAND="\$expand=singleValueExtendedProperties(\$filter=id%20eq%20'String%200x3A1F')"

# normalise_other: a contact or an array of contacts on stdin. Adds otherPhone
# (the number, or null when missing or empty) and drops the extended property.
# Graph returns the id in lower case, so it is matched ignoring case.
normalise_other() {
    jq -c '
        def one:
            ((.singleValueExtendedProperties // [])
             | map(select((.id | ascii_downcase) == "string 0x3a1f") | .value)
             | map(select(. != null and . != "")) | .[0]) as $o
            | del(.singleValueExtendedProperties) | . + {otherPhone: $o};
        if type == "array" then map(one) else one end'
}

# Most contacts any one command reads. A mailbox with more than this is
# truncated, and the command says so.
CONTACTS_MAX="${CONTACTS_MAX:-10000}"

# fetch_all <path-with-query> [max]: every item of a Graph collection,
# following @odata.nextLink, as one JSON array on stdout. On a Graph error,
# prints it to stderr and returns 1, so the caller stops rather than going on
# with half a list. Prints a note to stderr if it stopped at the maximum.
fetch_all() {
    local url="$1" max="${2:-$CONTACTS_MAX}" merged='[]' page next collected
    while [ -n "$url" ]; do
        page=$(api_call GET "$url")
        if ! printf '%s' "$page" | jq -e 'type == "object" and has("value")' >/dev/null 2>&1; then
            if printf '%s' "$page" | jq -e '.error' >/dev/null 2>&1; then
                echo "Error from Microsoft Graph:" >&2
                printf '%s' "$page" | jq -r '.error.message // .error.code' >&2
            else
                echo "Error: Graph returned no list." >&2
            fi
            return 1
        fi
        # Through stdin, not --argjson: one argument is capped at 128 KB on
        # Linux, and a few hundred contacts pass that.
        merged=$(printf '%s\n%s' "$merged" "$page" | jq -sc '.[0] + (.[1].value // [])')
        next=$(printf '%s' "$page" | jq -r '."@odata.nextLink" // empty')
        collected=$(printf '%s' "$merged" | jq 'length')
        if [ "$collected" -ge "$max" ]; then
            if [ -n "$next" ] || [ "$collected" -gt "$max" ]; then
                echo "Note: stopped at $max items. There are more." >&2
            fi
            break
        fi
        url="${next#"$GRAPH_URL"}"    # nextLink is absolute; strip base for api_call
    done
    printf '%s' "$merged" | jq -c --argjson max "$max" '.[0:$max]'
}

# --- Folders ----------------------------------------------------------------
# load_folders: sets FOLDERS to a JSON array of {id, displayName, path,
# parentFolderId} for every contact folder, children included. path is
# "Parent/Child". Loaded once per run.
#
# /me/contactFolders lists only the folders inside the main Contacts folder,
# never the main folder itself, so that one is fetched by its well-known name
# and put first. Without it a mailbox with no subfolders looks empty to search
# and duplicates, and nothing can be moved back into the main folder.
FOLDERS=""
load_folders() {
    [ -n "$FOLDERS" ] && return 0
    local main top queue='[]' all='[]' id kids path n=0
    main=$(api_call GET "/me/contactFolders/contacts") || return 1
    main=$(printf '%s' "$main" | jq -c '[{id, displayName, parentFolderId, path: .displayName}]') || return 1
    top=$(fetch_all "/me/contactFolders?\$top=100") || return 1
    queue=$(printf '%s' "$top" | jq -c 'map({id, displayName, parentFolderId, path: .displayName})')
    all=$(jq -nc --argjson a "$main" --argjson b "$queue" '$a + $b')
    # Breadth first through the children. The cap stops a runaway tree.
    while [ "$(printf '%s' "$queue" | jq 'length')" -gt 0 ] && [ "$n" -lt 200 ]; do
        id=$(printf '%s' "$queue" | jq -r '.[0].id')
        path=$(printf '%s' "$queue" | jq -r '.[0].path')
        queue=$(printf '%s' "$queue" | jq -c '.[1:]')
        n=$((n + 1))
        kids=$(fetch_all "/me/contactFolders/$id/childFolders?\$top=100") || return 1
        kids=$(printf '%s' "$kids" | jq -c --arg p "$path" 'map({id, displayName, parentFolderId, path: ($p + "/" + .displayName)})')
        all=$(printf '%s\n%s' "$all" "$kids" | jq -sc '.[0] + .[1]')
        queue=$(printf '%s\n%s' "$queue" "$kids" | jq -sc '.[0] + .[1]')
    done
    FOLDERS="$all"
}

# resolve_folder <name|id>: prints a folder's ID. A name matches the folder's
# name or its Parent/Child path, ignoring case, and must match one folder.
resolve_folder() {
    local want="$1" hits n
    load_folders || return 1
    hits=$(printf '%s' "$FOLDERS" | jq -r --arg w "$want" '
        ($w | ascii_downcase) as $l
        | (map(select(.id == $w)) ) as $byid
        | if ($byid | length) > 0 then $byid[].id
          else .[] | select((.path | ascii_downcase) == $l or (.displayName | ascii_downcase) == $l) | .id end')
    n=$(printf '%s' "$hits" | grep -c . || true)
    if [ "$n" -eq 0 ]; then
        if [ ${#want} -gt 60 ]; then   # a full folder ID that the listing did not show
            printf '%s' "$want"
            return 0
        fi
        echo "Error: no contact folder named: $want. Run: outlook-contacts.sh folders" >&2
        return 1
    fi
    if [ "$n" -gt 1 ]; then
        echo "Error: $n contact folders match '$want'. Use the Parent/Child path or the ID from: outlook-contacts.sh folders" >&2
        return 1
    fi
    printf '%s' "$hits"
}

# --- Contacts ---------------------------------------------------------------
# all_contacts: every contact in every folder, each with a .folder (its path),
# as one JSON array.
all_contacts() {
    local out='[]' id path rows
    load_folders || return 1
    while IFS=$'\t' read -r id path; do
        [ -n "$id" ] || continue
        rows=$(fetch_all "/me/contactFolders/$id/contacts?\$top=100&\$select=$CONTACT_SELECT&$OTHER_EXPAND") || return 1
        rows=$(printf '%s' "$rows" | normalise_other)
        # Through stdin: the list outgrows the 128 KB cap on one argument.
        out=$(printf '%s\n%s' "$out" "$rows" | jq -sc --arg f "$path" '.[0] + (.[1] | map(. + {folder: $f}))')
    done < <(printf '%s' "$FOLDERS" | jq -r '.[] | [.id, .path] | @tsv')
    printf '%s' "$out"
}

# Add each contact's folder path from its parentFolderId, when FOLDERS is known.
annotate_folder() {
    jq -c --argjson f "${FOLDERS:-[]}" '
        map(. + {folder: ((. as $c | $f | map(select(.id == $c.parentFolderId)) | .[0].path) // .folder // "-")})'
}

# Print a list of contacts (a JSON array on stdin) one per line, with a short
# ID. Two contacts whose short IDs match are printed with their full IDs.
format_contacts() {
    jq -r '
        def short_id: .[-20:];
        if length == 0 then "No contacts found."
        else
            ([.[].id | short_id] | group_by(.) | map(select(length > 1) | .[0])) as $clash
            | to_entries[] | .key as $i | .value as $c
            | ($c.id | short_id) as $s
            | "[\($i + 1)] \(if ($clash | index($s)) then $c.id else $s end) | \($c.displayName // "(no name)") | \([$c.emailAddresses[]?.address] | join(", ") | if . == "" then "-" else . end) | \([$c.mobilePhone, ($c.businessPhones // [])[], ($c.homePhones // [])[], $c.otherPhone] | map(select(. != null and . != "")) | join(", ") | if . == "" then "-" else . end) | \($c.companyName // "-") | \(($c.categories // []) | join(", ") | if . == "" then "-" else . end) | \($c.folder // "-")"
        end'
}

# resolve_contact_id <id>: a full ID passes through. A short ID (the last 20
# characters, as listings print) is matched against every contact. One that
# matches two is refused rather than guessed.
resolve_contact_id() {
    local want="$1" hits n
    if [ ${#want} -gt 60 ]; then
        printf '%s' "$want"
        return 0
    fi
    hits=$(all_contacts) || return 1
    hits=$(printf '%s' "$hits" | jq -r --arg s "$want" '.[] | select(.id | endswith($s)) | .id' | sort -u)
    n=$(printf '%s' "$hits" | grep -c . || true)
    if [ "$n" -eq 0 ]; then
        echo "Error: no contact found with an ID ending in: $want" >&2
        return 1
    fi
    if [ "$n" -gt 1 ]; then
        echo "Error: $n contacts have an ID ending in $want. Use the full ID from the listing." >&2
        return 1
    fi
    printf '%s' "$hits"
}

# --- Field flags for create and update ---------------------------------------
# parse_fields <args>: sets the F_* variables and HAVE_* flags. A flag that
# was given, even with an empty value, is "have"; update sends only those.
F_EMAILS="" F_BUSINESS="" F_HOME="" F_CATS=""
F_NAME="" F_GIVEN="" F_SURNAME="" F_MOBILE="" F_COMPANY="" F_TITLE="" F_NOTES="" F_FOLDER=""
F_STREET="" F_TOWN="" F_POSTCODE="" F_COUNTRY="" HAVE_ADDRESS="" F_OTHER="" HAVE_OTHER=""
HAVE_NAME="" HAVE_GIVEN="" HAVE_SURNAME="" HAVE_EMAILS="" HAVE_MOBILE="" HAVE_BUSINESS=""
HAVE_HOME="" HAVE_COMPANY="" HAVE_TITLE="" HAVE_NOTES="" HAVE_CATS=""
parse_fields() {
    local flag
    while [ $# -gt 0 ]; do
        flag="$1"
        case "$flag" in
            --name|--given|--surname|--email|--mobile|--phone|--business-phone|--home-phone|--other-phone|--company|--title|--notes|--category|--folder|--street|--town|--postcode|--country)
                if [ $# -lt 2 ]; then
                    echo "Error: $flag needs a value." >&2
                    return 1
                fi ;;
            *)
                echo "Error: unknown option: $flag" >&2
                return 1 ;;
        esac
        case "$flag" in
            --name)           F_NAME="$2"; HAVE_NAME=1 ;;
            --given)          F_GIVEN="$2"; HAVE_GIVEN=1 ;;
            --surname)        F_SURNAME="$2"; HAVE_SURNAME=1 ;;
            --email)          F_EMAILS="$F_EMAILS$2"$'\n'; HAVE_EMAILS=1 ;;
            --mobile|--phone) F_MOBILE="$2"; HAVE_MOBILE=1 ;;
            --business-phone) F_BUSINESS="$F_BUSINESS$2"$'\n'; HAVE_BUSINESS=1 ;;
            --home-phone)     F_HOME="$F_HOME$2"$'\n'; HAVE_HOME=1 ;;
            --other-phone)    F_OTHER="$2"; HAVE_OTHER=1 ;;
            --company)        F_COMPANY="$2"; HAVE_COMPANY=1 ;;
            --title)          F_TITLE="$2"; HAVE_TITLE=1 ;;
            --notes)          F_NOTES="$2"; HAVE_NOTES=1 ;;
            --category)       F_CATS="$F_CATS$2"$'\n'; HAVE_CATS=1 ;;
            --folder)         F_FOLDER="$2" ;;
            --street)         F_STREET="$2"; HAVE_ADDRESS=1 ;;
            --town)           F_TOWN="$2"; HAVE_ADDRESS=1 ;;
            --postcode)       F_POSTCODE="$2"; HAVE_ADDRESS=1 ;;
            --country)        F_COUNTRY="$2"; HAVE_ADDRESS=1 ;;
        esac
        shift 2
    done
}

# contact_body: the JSON for create or update from the F_* variables. Only the
# fields given are in it. An empty --email, --category or phone list entry is
# dropped, so `--email ""` on update clears the list. The address flags set the
# whole business address, so an update that gives one of them clears the parts
# it leaves out.
contact_body() {
    jq -n \
        --arg name "$F_NAME" --arg given "$F_GIVEN" --arg surname "$F_SURNAME" \
        --arg emails "$F_EMAILS" --arg mobile "$F_MOBILE" --arg business "$F_BUSINESS" \
        --arg home "$F_HOME" --arg company "$F_COMPANY" --arg title "$F_TITLE" \
        --arg notes "$F_NOTES" --arg cats "$F_CATS" \
        --arg street "$F_STREET" --arg town "$F_TOWN" --arg postcode "$F_POSTCODE" \
        --arg country "$F_COUNTRY" --arg haddress "$HAVE_ADDRESS" \
        --arg hname "$HAVE_NAME" --arg hgiven "$HAVE_GIVEN" --arg hsurname "$HAVE_SURNAME" \
        --arg hemails "$HAVE_EMAILS" --arg hmobile "$HAVE_MOBILE" --arg hbusiness "$HAVE_BUSINESS" \
        --arg hhome "$HAVE_HOME" --arg hcompany "$HAVE_COMPANY" --arg htitle "$HAVE_TITLE" \
        --arg hnotes "$HAVE_NOTES" --arg hcats "$HAVE_CATS" \
        --arg other "$F_OTHER" --arg hother "$HAVE_OTHER" --arg oid "$OTHER_PROP_ID" '
        def lines: split("\n") | map(gsub("^\\s+|\\s+$"; "")) | map(select(length > 0));
        {}
        | if $hname == "1" then .displayName = $name else . end
        | if $hgiven == "1" then .givenName = $given else . end
        | if $hsurname == "1" then .surname = $surname else . end
        | if $hemails == "1" then .emailAddresses = ($emails | lines | map({address: .})) else . end
        | if $hmobile == "1" then .mobilePhone = (if $mobile == "" then null else $mobile end) else . end
        | if $hbusiness == "1" then .businessPhones = ($business | lines) else . end
        | if $hhome == "1" then .homePhones = ($home | lines) else . end
        | if $hother == "1" then .singleValueExtendedProperties = [{id: $oid, value: $other}] else . end
        | if $hcompany == "1" then .companyName = $company else . end
        | if $htitle == "1" then .jobTitle = $title else . end
        | if $hnotes == "1" then .personalNotes = $notes else . end
        | if $hcats == "1" then .categories = ($cats | lines) else . end
        | if $haddress == "1" then .businessAddress = {street: $street, city: $town, postalCode: $postcode, countryOrRegion: $country} else . end'
}

# A one-line description of a contact (a JSON object on stdin).
describe_contact() {
    jq -r '"\(.displayName // "(no name)") <\([.emailAddresses[]?.address | select((. // "") != "")] | join(", "))>" | sub(" <>$"; "")'
}

# Remove what Graph sets itself, and the nulls, so a contact read back can be
# created again in another folder with everything else intact. otherPhone goes
# back to the extended property it was read from.
strip_for_copy() {
    jq -c --arg oid "$OTHER_PROP_ID" '
        (if (.otherPhone // "") != "" then .singleValueExtendedProperties = [{id: $oid, value: .otherPhone}] else . end)
        | del(.otherPhone)
        | with_entries(select((.key | startswith("@") or contains("@odata")) | not))
        | del(.id, .createdDateTime, .lastModifiedDateTime, .changeKey, .parentFolderId)
        | with_entries(select(.value != null))'
}

# Commands
case "$1" in
    list)
        shift
        folder="" top=50 sort=name
        while [ $# -gt 0 ]; do
            case "$1" in
                --folder) [ $# -ge 2 ] || { echo "Error: --folder needs a value." >&2; exit 1; }; folder="$2"; shift 2 ;;
                --top)    [ $# -ge 2 ] || { echo "Error: --top needs a value." >&2; exit 1; }; top="$2"; shift 2 ;;
                --sort)   [ $# -ge 2 ] || { echo "Error: --sort needs a value." >&2; exit 1; }; sort="$2"; shift 2 ;;
                *) echo "Error: unknown option: $1" >&2; exit 1 ;;
            esac
        done
        if ! [[ "$top" =~ ^[0-9]+$ ]] || [ "$top" -lt 1 ]; then
            echo "Error: --top must be a whole number of 1 or more." >&2
            exit 1
        fi
        case "$sort" in
            name)     order="displayName" ;;
            company)  order="companyName" ;;
            modified) order="lastModifiedDateTime%20desc" ;;
            *) echo "Error: --sort must be name, company or modified." >&2; exit 1 ;;
        esac
        page=100
        [ "$top" -lt "$page" ] && page="$top"
        query="\$top=$page&\$orderby=$order&\$select=$CONTACT_SELECT&$OTHER_EXPAND"
        if [ -n "$folder" ]; then
            load_folders || exit 1      # here, not in a subshell, so it is read once
            folder_id=$(resolve_folder "$folder") || exit 1
            path="/me/contactFolders/$folder_id/contacts?$query"
        else
            path="/me/contacts?$query"
        fi
        contacts=$(fetch_all "$path" "$top") || exit 1
        contacts=$(printf '%s' "$contacts" | normalise_other)
        load_folders || FOLDERS="[]"
        printf '%s' "$contacts" | annotate_folder | format_contacts
        ;;

    search)
        shift
        text="$*"
        if [ -z "$text" ]; then
            echo "Usage: outlook-contacts.sh search <text>" >&2
            exit 1
        fi
        contacts=$(all_contacts) || exit 1
        printf '%s' "$contacts" | jq -c --arg q "$text" '
            ($q | ascii_downcase) as $l
            | map(select(
                ([.displayName, .givenName, .surname, .companyName, (.emailAddresses[]?.address)]
                 | map(select(. != null) | ascii_downcase)
                 | any(contains($l)))))
            | sort_by((.displayName // "") | ascii_downcase)' | format_contacts
        ;;

    get)
        if [ -z "$2" ]; then
            echo "Usage: outlook-contacts.sh get <id>" >&2
            exit 1
        fi
        contact_id=$(resolve_contact_id "$2") || exit 1
        result=$(api_call GET "/me/contacts/$contact_id?$OTHER_EXPAND")
        die_on_error "$result" "reading the contact"
        printf '%s' "$result" | normalise_other | jq .
        ;;

    folders)
        load_folders || exit 1
        printf '%s' "$FOLDERS" | jq -r '
            if length == 0 then "No contact folders found."
            else to_entries[] | "[\(.key + 1)] \(.value.path) | \(.value.id)" end'
        ;;

    create)
        shift
        parse_fields "$@" || exit 1
        if [ -z "$HAVE_NAME$HAVE_GIVEN$HAVE_SURNAME$HAVE_EMAILS" ]; then
            echo "Error: give the contact at least a --name, --given, --surname or --email." >&2
            exit 1
        fi
        # A bare --name becomes a first and last name too, so Outlook files it
        # properly. --given and --surname, when given, are left alone.
        if [ -n "$HAVE_NAME" ] && [ -z "$HAVE_GIVEN$HAVE_SURNAME" ] && [ -n "$F_NAME" ]; then
            case "$F_NAME" in
                *" "*) F_GIVEN="${F_NAME%% *}"; F_SURNAME="${F_NAME#* }" ;;
                *)     F_GIVEN="$F_NAME" ;;
            esac
            HAVE_GIVEN=1; [ -n "$F_SURNAME" ] && HAVE_SURNAME=1
        fi
        body=$(contact_body)
        if [ -n "$F_FOLDER" ]; then
            load_folders || exit 1
            folder_id=$(resolve_folder "$F_FOLDER") || exit 1
            result=$(api_call POST "/me/contactFolders/$folder_id/contacts" "$body")
        else
            result=$(api_call POST "/me/contacts" "$body")
        fi
        die_on_error "$result" "creating the contact"
        echo "Created: $(printf '%s' "$result" | describe_contact)"
        echo "ID: $(printf '%s' "$result" | jq -r '.id // empty')"
        ;;

    update)
        contact_id="$2"
        if [ -z "$contact_id" ]; then
            echo "Usage: outlook-contacts.sh update <id> [fields]" >&2
            exit 1
        fi
        shift 2
        parse_fields "$@" || exit 1
        if [ -n "$F_FOLDER" ]; then
            echo "Error: --folder is for create. To change folder, use: outlook-contacts.sh move <id> --folder <name|id>" >&2
            exit 1
        fi
        body=$(contact_body)
        if [ "$body" = "{}" ]; then
            echo "Error: nothing to change. Give at least one field, for example --company \"...\"." >&2
            exit 1
        fi
        contact_id=$(resolve_contact_id "$contact_id") || exit 1
        result=$(api_call PATCH "/me/contacts/$contact_id" "$body")
        die_on_error "$result" "updating the contact"
        echo "Updated: $(printf '%s' "$result" | describe_contact)"
        ;;

    delete)
        if [ -z "$2" ]; then
            echo "Usage: outlook-contacts.sh delete <id>" >&2
            exit 1
        fi
        contact_id=$(resolve_contact_id "$2") || exit 1
        existing=$(api_call GET "/me/contacts/$contact_id?\$select=id,displayName,emailAddresses")
        die_on_error "$existing" "reading the contact before deleting it"
        label=$(printf '%s' "$existing" | describe_contact)
        result=$(api_call DELETE "/me/contacts/$contact_id")
        die_on_error "$result" "deleting the contact"
        echo "Deleted contact: $label"
        ;;

    folder-create)
        shift
        name="" parent=""
        while [ $# -gt 0 ]; do
            case "$1" in
                --parent) [ $# -ge 2 ] || { echo "Error: --parent needs a value." >&2; exit 1; }; parent="$2"; shift 2 ;;
                -*) echo "Error: unknown option: $1" >&2; exit 1 ;;
                *) if [ -z "$name" ]; then name="$1"; else echo "Error: one folder name only." >&2; exit 1; fi; shift ;;
            esac
        done
        if [ -z "$name" ]; then
            echo "Usage: outlook-contacts.sh folder-create <name> [--parent <name|id>]" >&2
            exit 1
        fi
        payload=$(jq -nc --arg n "$name" '{displayName: $n}')
        if [ -n "$parent" ]; then
            load_folders || exit 1
            parent_id=$(resolve_folder "$parent") || exit 1
            result=$(api_call POST "/me/contactFolders/$parent_id/childFolders" "$payload")
        else
            result=$(api_call POST "/me/contactFolders" "$payload")
        fi
        die_on_error "$result" "creating the folder"
        echo "Created contact folder: ${parent:+$parent/}$name"
        echo "ID: $(printf '%s' "$result" | jq -r '.id // empty')"
        ;;

    move)
        contact_id="$2"
        shift 2 2>/dev/null || true
        target=""
        while [ $# -gt 0 ]; do
            case "$1" in
                --folder) [ $# -ge 2 ] || { echo "Error: --folder needs a value." >&2; exit 1; }; target="$2"; shift 2 ;;
                *) echo "Error: unknown option: $1" >&2; exit 1 ;;
            esac
        done
        if [ -z "$contact_id" ] || [ -z "$target" ]; then
            echo "Usage: outlook-contacts.sh move <id> --folder <name|id>" >&2
            exit 1
        fi
        load_folders || exit 1
        folder_id=$(resolve_folder "$target") || exit 1
        contact_id=$(resolve_contact_id "$contact_id") || exit 1
        original=$(api_call GET "/me/contacts/$contact_id?$OTHER_EXPAND")
        die_on_error "$original" "reading the contact to move"
        original=$(printf '%s' "$original" | normalise_other)
        if [ "$(printf '%s' "$original" | jq -r '.parentFolderId // empty')" = "$folder_id" ]; then
            echo "Nothing to do: the contact is already in that folder."
            exit 0
        fi
        label=$(printf '%s' "$original" | describe_contact)
        # Graph cannot move a contact, so copy it and then delete the original.
        # The original is deleted only once the copy exists.
        copy=$(api_call POST "/me/contactFolders/$folder_id/contacts" "$(printf '%s' "$original" | strip_for_copy)")
        new_id=$(printf '%s' "$copy" | jq -r '.id // empty' 2>/dev/null) || new_id=""
        if [ -z "$new_id" ] || printf '%s' "$copy" | jq -e '.error' >/dev/null 2>&1; then
            echo "Error copying the contact to '$target'. The original is untouched." >&2
            printf '%s' "$copy" | jq -r '.error.message // .error.code // empty' >&2 2>/dev/null || true
            exit 1
        fi
        result=$(api_call DELETE "/me/contacts/$contact_id")
        if [ -n "$result" ] && printf '%s' "$result" | jq -e '.error' >/dev/null 2>&1; then
            echo "Copied $label to '$target', but could not delete the original, so it now exists twice:" >&2
            printf '%s' "$result" | jq -r '.error.message // .error.code' >&2
            echo "New ID: $new_id" >&2
            echo "Old ID: $contact_id" >&2
            exit 1
        fi
        echo "Moved $label to '$target'."
        echo "New ID: $new_id"
        echo "(A contact photo and any attachments are not copied.)"
        ;;

    categorise)
        contact_id="$2"
        shift 2 2>/dev/null || true
        add="" remove=""
        while [ $# -gt 0 ]; do
            case "$1" in
                --add)    [ $# -ge 2 ] || { echo "Error: --add needs a value." >&2; exit 1; }; add="$2"; shift 2 ;;
                --remove) [ $# -ge 2 ] || { echo "Error: --remove needs a value." >&2; exit 1; }; remove="$2"; shift 2 ;;
                *) echo "Error: unknown option: $1" >&2; exit 1 ;;
            esac
        done
        if [ -z "$contact_id" ] || { [ -z "$add" ] && [ -z "$remove" ]; }; then
            echo "Usage: outlook-contacts.sh categorise <id> --add <category> | --remove <category>" >&2
            exit 1
        fi
        contact_id=$(resolve_contact_id "$contact_id") || exit 1
        existing=$(api_call GET "/me/contacts/$contact_id?\$select=id,displayName,emailAddresses,categories")
        die_on_error "$existing" "reading the contact"
        cats=$(printf '%s' "$existing" | jq -c '.categories // []')
        if [ -n "$add" ]; then
            # Advisory only: a category outside the master list still works, but
            # Outlook shows it without a colour. The master list needs a
            # permission this app may not have, so a failed read is not an error.
            master=$(api_call GET "/me/outlook/masterCategories")
            if printf '%s' "$master" | jq -e 'has("value")' >/dev/null 2>&1; then
                if ! printf '%s' "$master" | jq -e --arg n "$add" 'any(.value[]; .displayName | ascii_downcase == ($n | ascii_downcase))' >/dev/null 2>&1; then
                    echo "Note: '$add' is not in the master category list, so it will show without a colour." >&2
                fi
            fi
            cats=$(printf '%s' "$cats" | jq -c --arg n "$add" 'if any(.[]; ascii_downcase == ($n | ascii_downcase)) then . else . + [$n] end')
        fi
        if [ -n "$remove" ]; then
            cats=$(printf '%s' "$cats" | jq -c --arg n "$remove" 'map(select(ascii_downcase != ($n | ascii_downcase)))')
        fi
        result=$(api_call PATCH "/me/contacts/$contact_id" "$(jq -nc --argjson c "$cats" '{categories: $c}')")
        die_on_error "$result" "setting the categories"
        echo "$(printf '%s' "$existing" | describe_contact): categories now $(printf '%s' "$cats" | jq -r 'if length == 0 then "none" else join(", ") end')"
        ;;

    duplicates)
        contacts=$(all_contacts) || exit 1
        report=$(printf '%s' "$contacts" | jq -r '
            def short_id: .[-20:];
            def norm_name: ascii_downcase | gsub("[^a-z0-9 ]"; " ") | split(" ") | map(select(length > 0)) | sort | join(" ");
            def phone_key: gsub("[^0-9]"; "") | if startswith("0044") then .[4:] elif startswith("44") then .[2:] elif startswith("0") then .[1:] else . end;
            def phones: [.mobilePhone, (.businessPhones // [])[], (.homePhones // [])[], .otherPhone] | map(select(. != null) | phone_key | select(length >= 7)) | unique;
            def row: "    \(.id | short_id) | \(.displayName // "(no name)") | \([.emailAddresses[]?.address | select((. // "") != "")] | join(", ") | if . == "" then "-" else . end) | \(.folder // "-")";
            . as $all
            | ([$all[] | . as $c | (.emailAddresses // [])[] | select((.address // "") != "") | {key: (.address | ascii_downcase), c: $c}]
               | group_by(.key) | map({kind: "email", key: .[0].key, members: (map(.c) | unique_by(.id))})) as $byemail
            | ([$all[] | select((.displayName // "") != "") | {key: (.displayName | norm_name), c: .}]
               | map(select(.key != "")) | group_by(.key)
               | map({kind: "name", key: .[0].key, members: (map(.c) | unique_by(.id))})) as $byname
            | ([$all[] | . as $c | phones[] | {key: ., c: $c}]
               | group_by(.key) | map({kind: "phone", key: .[0].key, members: (map(.c) | unique_by(.id))})) as $byphone
            | ($byemail + $byname + $byphone | map(select(.members | length > 1))) as $groups
            | if ($groups | length) == 0 then "No likely duplicates found among \($all | length) contacts."
              else
                "\($groups | length) likely duplicate group(s) among \($all | length) contacts:",
                ($groups[] | "  same \(.kind): \(.key)", (.members[] | row))
              end')
        printf '%s\n' "$report"
        ;;

    *)
        usage
        ;;
esac
