#!/bin/bash
# Offline tests for outlook-contacts.sh, run as a user would run it.
#
# A fake `curl` first on PATH answers Graph from fixture files and logs every
# request, and a throwaway HOME holds a fresh token, so no account or network
# is needed and nothing is changed for anyone. Nothing here touches ~/.dbhq.
#
#   bash skills/outlook/tests/contacts_test.sh
#
# What this pins: an account without Contacts.ReadWrite is refused before any
# request, listings page through @odata.nextLink, create and update send the
# right body, move is create then delete and never deletes when the create
# fails, duplicates finds a shared address, name or phone number, the Other phone is read, written and moved, and categorise edits the
# contact's own category list.
#
# The literal '$skip' and '$top' strings are Graph query parameters (SC2016).
# shellcheck disable=SC2016
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CON="$(cd "$TESTS_DIR/../scripts" && pwd)/outlook-contacts.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }
lacks() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" "does not contain: $2" "$3"; else eq "$1" ok ok; fi; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT

# --- Fake curl -----------------------------------------------------------------
# Two folders: FOLDER1 (the main "Contacts" folder) has two pages, FOLDER2
# ("Clients", with a child "Archive") has one. A GET of one contact is answered
# from $FAKE_CONTACT. A POST that creates a contact answers with the body plus
# an id, unless $FAKE_CREATE_FAIL is set. "METHOD URL" is logged, and on the next
# line "BODY <json>" for a request with a body.
mkdir -p "$TMP/bin"
cat > "$TMP/bin/curl" <<'FAKE'
#!/bin/bash
url="" method=GET data="" prev=""
for a in "$@"; do
  case "$prev" in -X) method="$a" ;; -d) data="$a" ;; esac
  case "$a" in https://*) url="$a" ;; esac
  prev="$a"
done
printf '%s %s\n' "$method" "$url" >> "$FAKE_CURL_LOG"
[ -n "$data" ] && printf 'BODY %s\n' "$(printf '%s' "$data" | jq -c . 2>/dev/null || printf '%s' "$data")" >> "$FAKE_CURL_LOG"
path="${url#https://graph.microsoft.com/v1.0}"
case "$method $path" in
  "GET /me/contactFolders?"*) cat "$FAKE_FOLDERS" ;;
  "GET /me/contactFolders/contacts") printf '{"id":"FOLDER1","displayName":"Contacts","parentFolderId":"ROOT"}' ;;
  "GET /me/contactFolders/FOLDER2/childFolders"*) printf '{"value":[{"id":"FOLDER3","displayName":"Archive","parentFolderId":"FOLDER2"}]}' ;;
  "GET /me/contactFolders/"*/childFolders*) printf '{"value":[]}' ;;
  "GET /me/contactFolders/FOLDER1/contacts"*'$skip'*) cat "$FAKE_F1_PAGE2" ;;
  "GET /me/contactFolders/FOLDER1/contacts"*) cat "$FAKE_F1_PAGE1" ;;
  "GET /me/contactFolders/FOLDER2/contacts"*) cat "$FAKE_F2" ;;
  "GET /me/contactFolders/"*/contacts*) printf '{"value":[]}' ;;
  "GET /me/contacts?"*'$skip'*) cat "$FAKE_F1_PAGE2" ;;
  "GET /me/contacts?"*) cat "$FAKE_F1_PAGE1" ;;
  "GET /me/contacts/"*) cat "$FAKE_CONTACT" ;;
  "GET /me/outlook/masterCategories") printf '{"value":[{"displayName":"Blue"}]}' ;;
  "POST /me/contactFolders/"*/contacts|"POST /me/contacts")
      if [ -n "${FAKE_CREATE_FAIL:-}" ]; then
          printf '{"error":{"code":"ErrorInvalidIdMalformed","message":"Create refused."}}'
      else
          printf '%s' "$data" | jq -c '. + {id: "NEWCONTACTIDxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"}'
      fi ;;
  "POST /me/contactFolders/"*/childFolders|"POST /me/contactFolders")
      printf '%s' "$data" | jq -c '. + {id: "NEWFOLDERID"}' ;;
  "PATCH "*) printf '%s' "$data" | jq -c '. + {id: "x", displayName: (.displayName // "Patched Person")}' ;;
  "DELETE "*) : ;;
  *) printf '{"value":[]}' ;;
esac
exit 0
FAKE
chmod +x "$TMP/bin/curl"
export PATH="$TMP/bin:$PATH"
export FAKE_CURL_LOG="$TMP/curl.log" FAKE_FOLDERS="$TMP/folders.json" FAKE_F1_PAGE1="$TMP/f1p1.json" \
       FAKE_F1_PAGE2="$TMP/f1p2.json" FAKE_F2="$TMP/f2.json" FAKE_CONTACT="$TMP/contact.json"

# --- Fixture HOME -----------------------------------------------------------------
FIX_HOME="$TMP/home"
ACC="$FIX_HOME/.dbhq/outlook/default"
mkdir -p "$ACC"
chmod 700 "$FIX_HOME/.dbhq" "$FIX_HOME/.dbhq/outlook" "$ACC"
printf '%s' '{"client_id":"c"}' > "$ACC/config.json"
EXP=$(( $(date +%s) + 3600 ))
granted()    { printf '{"access_token":"tok","refresh_token":"r","expires_at":%s,"scope":"Mail.ReadWrite Contacts.ReadWrite User.Read"}' "$EXP" > "$ACC/credentials.json"; }
granted_url() { printf '{"access_token":"tok","refresh_token":"r","expires_at":%s,"scope":"https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Contacts.ReadWrite"}' "$EXP" > "$ACC/credentials.json"; }
legacy()     { printf '{"access_token":"tok","refresh_token":"r","expires_at":%s,"scope":"offline_access Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read"}' "$EXP" > "$ACC/credentials.json"; }
no_scope()   { printf '{"access_token":"tok","refresh_token":"r","expires_at":%s}' "$EXP" > "$ACC/credentials.json"; }
granted

con() {
    : > "$FAKE_CURL_LOG"
    HOME="$FIX_HOME" OUTLOOK_ACCOUNT=default bash "$CON" "$@" 2>&1
}
requests() { grep -vc '^BODY ' "$FAKE_CURL_LOG" || true; }
body() { sed -n 's/^BODY //p' "$FAKE_CURL_LOG" | head -1; }

cid() { printf 'AAMkADkzNjdmZjRlLTc4ZWYtNDdiNC1hZDdiLWE3MjZmNTg0YjAwNgBGAAAAAAD%s' "$1"; }
contact() {  # id name email company [categories-json] [folder-id]
    printf '{"id":"%s","displayName":"%s","emailAddresses":[{"address":"%s"}],"companyName":"%s","categories":%s,"parentFolderId":"%s"}' \
        "$1" "$2" "$3" "$4" "${5:-[]}" "${6:-FOLDER1}"
}

# Like the real Graph, /me/contactFolders lists only the folders inside the
# main Contacts folder, never the main folder itself (FOLDER1 here), which
# comes from /me/contactFolders/contacts.
printf '%s' '{"value":[{"id":"FOLDER2","displayName":"Clients","parentFolderId":"FOLDER1"}]}' > "$FAKE_FOLDERS"
printf '{"@odata.nextLink":"https://graph.microsoft.com/v1.0/me/contacts?$top=100&$skip=2","value":[%s,%s]}' \
    "$(contact "$(cid LISTONE0000000001AA=)" "Ada Lovelace" ada@example.com "Analytical Co" '["Blue"]')" \
    "$(contact "$(cid LISTTWO0000000002AA=)" "Alan Turing" alan@example.com "Bletchley Ltd")" > "$FAKE_F1_PAGE1"
printf '{"value":[%s]}' \
    "$(contact "$(cid LISTTHR0000000003AA=)" "Grace Hopper" grace@example.com "Navy")" > "$FAKE_F1_PAGE2"
printf '{"value":[%s]}' \
    "$(contact "$(cid CLIENTONE00000004AA=)" "Katherine Johnson" kj@example.com "Nasa" '[]' FOLDER2)" > "$FAKE_F2"

########################################
# Missing scope: refused before any request, with the sign-in command.
########################################
for state in legacy no_scope; do
    "$state"
    out=$(con list); rc=$?
    eq "$state: list exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
    has "$state: says to sign in again" "outlook-setup.sh --account default" "$out"
    has "$state: names the permission" "Contacts.ReadWrite" "$out"
    eq "$state: makes no request" "0" "$(requests)"
done
out=$(con --account default create --name X); rc=$?
eq "an --account before the verb is refused too" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "a refused create sends nothing" "0" "$(requests)"
legacy; out=$(con --help)
has "help needs no permission" "Usage: outlook-contacts.sh" "$out"

granted_url
out=$(con list)
has "scopes returned as URLs count as granted" "Ada Lovelace" "$out"
granted

########################################
# list: pages through @odata.nextLink, short IDs, folder name, sort.
########################################
out=$(con list)
eq "list shows all 3 contacts across two pages" "3" "$(printf '%s\n' "$out" | grep -c '^\[[0-9]*\] ')"
has "list shows the row" "Ada Lovelace | ada@example.com | - | Analytical Co | Blue | Contacts" "$out"
has "list prints a short ID" "$(cid LISTONE0000000001AA= | tail -c 20)" "$out"
eq "list follows the next page" "1" "$(grep -c 'contacts?.*skip=2' "$FAKE_CURL_LOG")"
eq "list asks for the main folder by default" "1" "$(head -1 "$FAKE_CURL_LOG" | grep -c 'GET https://graph.microsoft.com/v1.0/me/contacts?')"

out=$(con list --top 1)
eq "list --top 1 shows one contact" "1" "$(printf '%s\n' "$out" | grep -c '^\[[0-9]*\] ')"
eq "list --top 1 asks for one" "1" "$(grep -c 'me/contacts?\$top=1&' "$FAKE_CURL_LOG")"

con list --sort modified >/dev/null
eq "list --sort modified orders by last modified" "1" "$(grep -c 'orderby=lastModifiedDateTime%20desc' "$FAKE_CURL_LOG")"
con list --sort company >/dev/null
eq "list --sort company orders by company" "1" "$(grep -c 'orderby=companyName' "$FAKE_CURL_LOG")"
out=$(con list --sort size); rc=$?
eq "list --sort with a bad key is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"

out=$(con list --folder clients)
has "list --folder resolves a name, ignoring case" "Katherine Johnson" "$out"
eq "list --folder reads that folder" "1" "$(grep -c 'GET https://graph.microsoft.com/v1.0/me/contactFolders/FOLDER2/contacts' "$FAKE_CURL_LOG")"
has "list --folder shows the folder name" "| Clients" "$out"
con list --folder Clients/Archive >/dev/null
eq "list --folder takes a Parent/Child path" "1" "$(grep -c 'me/contactFolders/FOLDER3/contacts' "$FAKE_CURL_LOG")"
out=$(con list --folder Nowhere); rc=$?
eq "list --folder with an unknown name fails" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "list --folder with an unknown name says so" "no contact folder named: Nowhere" "$out"

########################################
# search covers every folder, by name, email and company.
########################################
out=$(con search johnson)
has "search finds by name in another folder" "Katherine Johnson" "$out"
lacks "search leaves out the rest" "Ada Lovelace" "$out"
out=$(con search GRACE@EXAMPLE)
has "search finds by email, ignoring case" "Grace Hopper" "$out"
out=$(con search bletchley)
has "search finds by company" "Alan Turing" "$out"
out=$(con search zzzz)
has "search says when nothing matches" "No contacts found." "$out"

########################################
# get: a short ID from a listing resolves; the full contact comes back.
########################################
FULL=$(cid LISTTWO0000000002AA=)
contact "$FULL" "Alan Turing" alan@example.com "Bletchley Ltd" > "$FAKE_CONTACT"
out=$(con get "$(printf '%s' "$FULL" | tail -c 20)")
eq "get resolves a short ID" "1" "$(grep -c "GET https://graph.microsoft.com/v1.0/me/contacts/$FULL?" "$FAKE_CURL_LOG")"
eq "get prints the contact as JSON" "Alan Turing" "$(printf '%s' "$out" | jq -r .displayName)"
out=$(con get ZZZZnotanid); rc=$?
eq "get with an unknown short ID fails" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"

########################################
# create: body shape, and the folder it lands in.
########################################
out=$(con create --name "Mary Jackson" --email mj@example.com --email mary@example.org \
    --mobile "+44 7700 900000" --business-phone "01632 960000" --home-phone "01632 960001" \
    --company "Nasa" --title "Engineer" --notes "Met at a talk" --category Blue --category Clients)
has "create says what it made" "Created: Mary Jackson <mj@example.com, mary@example.org>" "$out"
has "create prints the new ID" "ID: NEWCONTACTID" "$out"
eq "create posts to the main folder" "1" "$(grep -c '^POST https://graph.microsoft.com/v1.0/me/contacts$' "$FAKE_CURL_LOG")"
eq "create body: names split from --name" "Mary Jackson|Mary|Jackson" \
   "$(body | jq -r '[.displayName, .givenName, .surname] | join("|")')"
eq "create body: both emails" "mj@example.com,mary@example.org" "$(body | jq -r '[.emailAddresses[].address] | join(",")')"
eq "create body: phones" "+44 7700 900000|01632 960000|01632 960001" \
   "$(body | jq -r '[.mobilePhone, .businessPhones[0], .homePhones[0]] | join("|")')"
eq "create body: company, title, notes" "Nasa|Engineer|Met at a talk" \
   "$(body | jq -r '[.companyName, .jobTitle, .personalNotes] | join("|")')"
eq "create body: categories" "Blue,Clients" "$(body | jq -r '.categories | join(",")')"
eq "create body: no address unless one is given" "false" "$(body | jq 'has("businessAddress")')"
out=$(con create --name "Corner Shop" --street "1 High Street" --town Exampletown --postcode "AB1 2CD" --country "United Kingdom")
eq "create body: the address flags fill the business address" "1 High Street|Exampletown|AB1 2CD|United Kingdom" \
   "$(body | jq -r '.businessAddress | [.street, .city, .postalCode, .countryOrRegion] | join("|")')"

out=$(con create --given Ada --surname King --email ak@example.com --folder Clients)
eq "create --folder posts into that folder" "1" "$(grep -c '^POST https://graph.microsoft.com/v1.0/me/contactFolders/FOLDER2/contacts$' "$FAKE_CURL_LOG")"
eq "create keeps --given and --surname as given" "Ada|King" "$(body | jq -r '[.givenName, .surname] | join("|")')"
eq "create does not send the folder as a field" "false" "$(body | jq 'has("folder")')"
out=$(con create); rc=$?
eq "create with no name or email is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "create with no name or email sends nothing" "0" "$(requests)"
out=$(con create --name X --colour red); rc=$?
eq "create with an unknown option is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "create with an unknown option sends nothing" "0" "$(requests)"
out=$(con create --name Zed --folder Nowhere); rc=$?
eq "create into an unknown folder is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "create into an unknown folder posts nothing" "0" "$(grep -c '^POST ' "$FAKE_CURL_LOG")"
out=$(FAKE_CREATE_FAIL=1 con create --name Zed); rc=$?
eq "create reports a Graph error and exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "create shows the Graph message" "Create refused." "$out"

########################################
# update: only the fields given are sent.
########################################
out=$(con update "$(printf '%s' "$FULL" | tail -c 20)" --company "New Co" --title "")
eq "update patches the resolved contact" "1" "$(grep -c "^PATCH https://graph.microsoft.com/v1.0/me/contacts/$FULL\$" "$FAKE_CURL_LOG")"
eq "update sends only the fields given" "companyName,jobTitle" "$(body | jq -r 'keys | join(",")')"
eq "update can clear a field" "" "$(body | jq -r '.jobTitle')"
out=$(con update "$FULL" --email a@example.com --email b@example.com)
eq "update replaces the email list" "a@example.com,b@example.com" "$(body | jq -r '[.emailAddresses[].address] | join(",")')"
out=$(con update "$FULL"); rc=$?
eq "update with nothing to change is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
out=$(con update "$FULL" --folder Clients); rc=$?
eq "update --folder is refused (use move)" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "update --folder points to move" "move" "$out"

########################################
# delete says what it deleted.
########################################
out=$(con delete "$FULL")
eq "delete deletes the contact" "1" "$(grep -c "^DELETE https://graph.microsoft.com/v1.0/me/contacts/$FULL\$" "$FAKE_CURL_LOG")"
has "delete says who was deleted" "Deleted contact: Alan Turing <alan@example.com>" "$out"

########################################
# move: create in the target, then delete the original.
########################################
contact "$FULL" "Alan Turing" alan@example.com "Bletchley Ltd" '["Blue"]' FOLDER1 > "$FAKE_CONTACT"
out=$(con move "$FULL" --folder Clients)
eq "move creates the copy first" "POST" "$(grep -v '^BODY ' "$FAKE_CURL_LOG" | grep -n 'POST\|DELETE' | head -1 | sed 's/^[0-9]*:\([A-Z]*\) .*/\1/')"
eq "move posts to the target folder" "1" "$(grep -c '^POST https://graph.microsoft.com/v1.0/me/contactFolders/FOLDER2/contacts$' "$FAKE_CURL_LOG")"
eq "move then deletes the original" "1" "$(grep -c "^DELETE https://graph.microsoft.com/v1.0/me/contacts/$FULL\$" "$FAKE_CURL_LOG")"
eq "move deletes after it creates" "POST DELETE" "$(grep -E '^(POST|DELETE) ' "$FAKE_CURL_LOG" | awk '{print $1}' | tr '\n' ' ' | sed 's/ $//')"
eq "move keeps the fields" "Alan Turing|Bletchley Ltd|Blue" "$(body | jq -r '[.displayName, .companyName, .categories[0]] | join("|")')"
eq "move does not send the old ID or folder" "false" "$(body | jq 'has("id") or has("parentFolderId")')"
has "move prints the new ID" "New ID: NEWCONTACTID" "$out"

out=$(FAKE_CREATE_FAIL=1 con move "$FULL" --folder Clients); rc=$?
eq "move exits non-zero when the create fails" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "move does not delete when the create fails" "0" "$(grep -c '^DELETE ' "$FAKE_CURL_LOG")"
has "move says the original is untouched" "original is untouched" "$out"

out=$(con move "$FULL" --folder Contacts)
eq "move to the folder it is already in does nothing" "0" "$(grep -c '^POST \|^DELETE ' "$FAKE_CURL_LOG")"
out=$(con move "$FULL" --folder Nowhere); rc=$?
eq "move to an unknown folder fails" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
eq "move to an unknown folder changes nothing" "0" "$(grep -c '^POST \|^DELETE ' "$FAKE_CURL_LOG")"
out=$(con move "$FULL"); rc=$?
eq "move without --folder is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"

########################################
# folders and folder-create.
########################################
out=$(con folders)
has "folders lists the top level" "Contacts | FOLDER1" "$out"
has "folders lists a child by its path" "Clients/Archive | FOLDER3" "$out"
out=$(con folder-create "Suppliers")
eq "folder-create posts to contactFolders" "1" "$(grep -c '^POST https://graph.microsoft.com/v1.0/me/contactFolders$' "$FAKE_CURL_LOG")"
eq "folder-create sends the name" "Suppliers" "$(body | jq -r .displayName)"
has "folder-create prints the new ID" "ID: NEWFOLDERID" "$out"
out=$(con folder-create "Old" --parent Clients)
eq "folder-create --parent posts to childFolders" "1" "$(grep -c '^POST https://graph.microsoft.com/v1.0/me/contactFolders/FOLDER2/childFolders$' "$FAKE_CURL_LOG")"

########################################
# categorise: edits the contact's own list, case-insensitively.
########################################
contact "$FULL" "Alan Turing" alan@example.com "Bletchley Ltd" '["Blue","Red"]' > "$FAKE_CONTACT"
out=$(con categorise "$FULL" --add Green)
eq "categorise --add appends" "Blue,Red,Green" "$(body | jq -r '.categories | join(",")')"
has "categorise warns about a category outside the master list" "not in the master category list" "$out"
out=$(con categorise "$FULL" --add blue)
eq "categorise --add does not duplicate (any case)" "Blue,Red" "$(body | jq -r '.categories | join(",")')"
lacks "categorise --add of a known category does not warn" "master category list" "$out"
out=$(con categorise "$FULL" --remove RED)
eq "categorise --remove removes (any case)" "Blue" "$(body | jq -r '.categories | join(",")')"
out=$(con categorise "$FULL"); rc=$?
eq "categorise with no --add or --remove is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"

########################################
# duplicates: same address or same name, across folders. Changes nothing.
########################################
printf '{"value":[%s,%s,%s]}' \
    "$(contact "$(cid DUPONE00000000001AA=)" "Ada Lovelace" "ADA@example.com" "A")" \
    "$(contact "$(cid DUPTWO00000000002AA=)" "Countess Lovelace" "ada@example.com" "B")" \
    "$(contact "$(cid DUPTHR00000000003AA=)" "Alan Turing" "alan@example.com" "C")" > "$FAKE_F1_PAGE1"
printf '{"value":[%s]}' "$(contact "$(cid DUPFOU00000000004AA=)" "turing, alan" "other@example.com" "D" '[]' FOLDER2)" > "$FAKE_F2"
out=$(con duplicates)
has "duplicates reports a shared address" "same email: ada@example.com" "$out"
has "duplicates lists both holders of the address" "Countess Lovelace" "$out"
has "duplicates reports a shared name across folders" "same name: alan turing" "$out"
has "duplicates shows the folder" "Clients" "$out"
eq "duplicates changes nothing" "0" "$(grep -cE '^(POST|PATCH|DELETE) ' "$FAKE_CURL_LOG")"
printf '{"value":[%s]}' "$(contact "$(cid DUPONE00000000001AA=)" "Ada Lovelace" "ada@example.com" "A")" > "$FAKE_F1_PAGE1"
printf '{"value":[]}' > "$FAKE_F2"
has "duplicates says when there are none" "No likely duplicates found" "$(con duplicates)"

# A mailbox with no subfolders at all keeps every contact in the main folder,
# which /me/contactFolders never lists. search and duplicates must still see them.
saved_folders=$(cat "$FAKE_FOLDERS")
printf '%s' '{"value":[]}' > "$FAKE_FOLDERS"
printf '{"value":[%s,%s]}' \
    "$(contact "$(cid MAINONE0000000001AA=)" "Ada Lovelace" "ada@example.com" "A")" \
    "$(contact "$(cid MAINTWO0000000002AA=)" "Countess Lovelace" "ada@example.com" "B")" > "$FAKE_F1_PAGE1"
has "search finds a contact in the main folder with no subfolders" "Countess Lovelace" "$(con search countess)"
has "duplicates sees the main folder with no subfolders" "same email: ada@example.com" "$(con duplicates)"

# Outlook keeps three email slots, so a contact with none set can come back as
# [{"address":""},...]. Two such contacts are not duplicates of each other.
printf '{"value":[%s,%s]}' \
    "$(contact "$(cid BLANKONE000000001AA=)" "Dan Fowler" "" "A")" \
    "$(contact "$(cid BLANKTWO000000002AA=)" "Bradley Jenkins" "" "B")" > "$FAKE_F1_PAGE1"
out=$(con duplicates)
has "blank email slots are not a shared address" "No likely duplicates found" "$out"
# Outlook returns all three slots, so three blank addresses must not print as ", ,".
printf '{"value":[{"id":"%s","displayName":"Three Blanks","emailAddresses":[{"address":""},{"address":""},{"address":""}],"categories":[],"parentFolderId":"FOLDER1"}]}' "$(cid THREEBLANK0000001AA=)" > "$FAKE_F1_PAGE1"
lacks "blank email slots are not listed as an address" ", ," "$(con list)"
has "a contact with only blank email slots shows no address" "Three Blanks | - |" "$(con list)"
# A mailbox big enough that its contacts pass the 128 KB cap on one argument.
big=$(jq -nc '[range(0; 600) | {id: ("AAMkBIG" + (tostring) + "xxxxxxxxxxxxxxxxxxxxxxxxxxxx"), displayName: ("Person \(.)"), emailAddresses: [{address: ("p\(.)@example.com")}], companyName: ("Company \(.) with a long enough name to add bulk"), personalNotes: ("x" * 200), categories: [], parentFolderId: "FOLDER1"}] | {value: .}')
printf '%s' "$big" > "$FAKE_F1_PAGE1"
out=$(con search "person 599")
has "search works on a mailbox past the 128 KB argument cap" "Person 599" "$out"
lacks "search past the cap raises no argument error" "Argument list too long" "$out"
printf '%s' "$saved_folders" > "$FAKE_FOLDERS"

########################################
# The Other phone: Graph has no field for it, so it is read and written as the
# extended property String 0x3A1F (returned with a lower-case id).
########################################
EXPAND_ENC="expand=singleValueExtendedProperties(\$filter=id%20eq%20'String%200x3A1F')"
OTHER_ID=$(cid OTHERONLY0000000001AA=)
printf '{"value":[%s,%s]}' \
    "{\"id\":\"$OTHER_ID\",\"displayName\":\"Otto Only\",\"emailAddresses\":[],\"parentFolderId\":\"FOLDER1\",\"singleValueExtendedProperties\":[{\"id\":\"String 0x3a1f\",\"value\":\"+44 7700 900111\"}]}" \
    "{\"id\":\"$(cid CLEARED00000000002AA=)\",\"displayName\":\"Clara Cleared\",\"emailAddresses\":[],\"mobilePhone\":\"07700 900222\",\"businessPhones\":[\"01632 960333\"],\"homePhones\":[\"01632 960444\"],\"parentFolderId\":\"FOLDER1\",\"singleValueExtendedProperties\":[{\"id\":\"String 0x3a1f\",\"value\":\"\"}]}" > "$FAKE_F1_PAGE1"
printf '{"value":[]}' > "$FAKE_F2"
out=$(con list)
has "list asks for the Other phone" "$EXPAND_ENC" "$(cat "$FAKE_CURL_LOG")"
has "list shows a contact whose only number is Other" "Otto Only |" "$out"
has "list prints the Other number in the phone column" "Otto Only | - | +44 7700 900111 | - | - | Contacts" "$out"
has "list prints mobile first, then business, then home" "Clara Cleared | - | 07700 900222, 01632 960333, 01632 960444 |" "$out"
lacks "list leaves out an emptied Other phone" "960444, " "$out"
out=$(con search otto)
has "search asks for the Other phone" "$EXPAND_ENC" "$(cat "$FAKE_CURL_LOG")"
has "search shows a contact whose only number is Other" "Otto Only | - | +44 7700 900111" "$out"

contact "$OTHER_ID" "Otto Only" otto@example.com "" > "$FAKE_CONTACT"
jq -c '. + {singleValueExtendedProperties: [{id: "String 0x3a1f", value: "+44 7700 900111"}]}' "$FAKE_CONTACT" > "$FAKE_CONTACT.new" && mv "$FAKE_CONTACT.new" "$FAKE_CONTACT"
out=$(con get "$OTHER_ID")
has "get asks for the Other phone" "$EXPAND_ENC" "$(cat "$FAKE_CURL_LOG")"
eq "get shows otherPhone" "+44 7700 900111" "$(printf '%s' "$out" | jq -r .otherPhone)"
eq "get hides the extended property" "false" "$(printf '%s' "$out" | jq 'has("singleValueExtendedProperties")')"
contact "$OTHER_ID" "Otto Only" otto@example.com "" > "$FAKE_CONTACT"
eq "get shows otherPhone as null when there is none" "null" "$(con get "$OTHER_ID" | jq -c .otherPhone)"

# Create and update write the extended property.
out=$(con create --name "Pat Other" --other-phone "01632 960555")
eq "create body: the Other phone is an extended property" "String 0x3A1F|01632 960555" \
   "$(body | jq -r '.singleValueExtendedProperties[0] | [.id, .value] | join("|")')"
out=$(con create --name "Pat Plain")
eq "create body: no Other phone unless given" "false" "$(body | jq 'has("singleValueExtendedProperties")')"
out=$(con update "$OTHER_ID" --other-phone "07700 900666")
eq "update sends only the Other phone" "singleValueExtendedProperties" "$(body | jq -r 'keys | join(",")')"
eq "update body: the Other phone value" "String 0x3A1F|07700 900666" \
   "$(body | jq -r '.singleValueExtendedProperties[0] | [.id, .value] | join("|")')"
out=$(con update "$OTHER_ID" --other-phone "")
eq "update --other-phone \"\" clears it with an empty value" "String 0x3A1F|" \
   "$(body | jq -r '.singleValueExtendedProperties[0] | [.id, .value] | join("|")')"
out=$(con update "$OTHER_ID" --other-phone); rc=$?
eq "--other-phone with no value is refused" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"

# Move carries it to the copy, and leaves no otherPhone field behind.
jq -c '. + {singleValueExtendedProperties: [{id: "String 0x3a1f", value: "+44 7700 900111"}]}' "$FAKE_CONTACT" > "$FAKE_CONTACT.new" && mv "$FAKE_CONTACT.new" "$FAKE_CONTACT"
out=$(con move "$OTHER_ID" --folder Clients)
has "move reads the Other phone" "$EXPAND_ENC" "$(cat "$FAKE_CURL_LOG")"
eq "move copies the Other phone to the new contact" "String 0x3A1F|+44 7700 900111" \
   "$(body | jq -r '.singleValueExtendedProperties[0] | [.id, .value] | join("|")')"
eq "move does not send otherPhone as a field" "false" "$(body | jq 'has("otherPhone")')"
contact "$OTHER_ID" "Otto Only" otto@example.com "" > "$FAKE_CONTACT"
out=$(con move "$OTHER_ID" --folder Clients)
eq "move sends no extended property when there is no Other phone" "false" "$(body | jq 'has("singleValueExtendedProperties")')"

# Duplicates: a shared number is found whichever way it is written and
# whichever slot holds it.
printf '{"value":[%s,%s,%s]}' \
    "{\"id\":\"$(cid PHONEONE000000001AA=)\",\"displayName\":\"Pia Phone\",\"emailAddresses\":[],\"mobilePhone\":\"07700 900777\",\"parentFolderId\":\"FOLDER1\"}" \
    "{\"id\":\"$(cid PHONETWO000000002AA=)\",\"displayName\":\"P Phoneman\",\"emailAddresses\":[],\"parentFolderId\":\"FOLDER1\",\"singleValueExtendedProperties\":[{\"id\":\"String 0x3a1f\",\"value\":\"+44 7700 900777\"}]}" \
    "{\"id\":\"$(cid PHONETHR000000003AA=)\",\"displayName\":\"Quentin Quiet\",\"emailAddresses\":[],\"mobilePhone\":\"07700 900888\",\"parentFolderId\":\"FOLDER1\"}" > "$FAKE_F1_PAGE1"
out=$(con duplicates)
has "duplicates asks for the Other phone" "$EXPAND_ENC" "$(cat "$FAKE_CURL_LOG")"
has "duplicates reports a shared number" "same phone: 7700900777" "$out"
has "duplicates lists both holders of the number" "P Phoneman" "$out"
lacks "duplicates leaves out a number held once" "Quentin Quiet" "$out"

########################################
# A Graph error is reported and the command exits non-zero.
########################################
printf '%s' '{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}' > "$FAKE_F1_PAGE1"
out=$(con list); rc=$?
eq "list on a Graph error exits non-zero" "1" "$([ "$rc" -ne 0 ] && echo 1 || echo 0)"
has "list on a Graph error says why" "Access is denied." "$out"

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
