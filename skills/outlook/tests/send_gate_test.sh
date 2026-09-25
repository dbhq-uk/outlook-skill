#!/bin/bash
# Offline tests for the two Claude Code layers of the send gap: the plugin's
# PreToolUse hook and the ask rules install.sh offers.
#
#   bash skills/outlook/tests/send_gate_test.sh
#
# What this pins:
#   - hooks/send-gate.sh answers "ask" for every command that sends: mail
#     `send`, calendar `invite`, `respond`, `cancel`, `create --send-invites`
#     and `update --notify-attendees`, however the script path is written.
#   - It stays silent, and exits 0, for everything else - drafts, reads, `sent`,
#     other tools, and input it cannot parse - so it never blocks other work.
#   - hooks/ask-rules.json has a rule for each of those sends and none that
#     matches a draft or a read.
#   - hooks/install-ask-rules.sh changes nothing without a yes, adds only the
#     rules that are missing, and keeps everything else in the settings file.
#   - install.sh passes --yes to it only when --ask-rules was given, and with
#     no terminal it never runs it at all.
#
# Every settings file here is under a throwaway directory. Nothing touches the
# real ~/.claude.
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$TESTS_DIR/../../.." && pwd)"
HOOK="$REPO/hooks/send-gate.sh"
RULES="$REPO/hooks/ask-rules.json"
HELPER="$REPO/hooks/install-ask-rules.sh"

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }
has() { if printf '%s' "$3" | grep -qF -- "$2"; then eq "$1" ok ok; else eq "$1" "contains: $2" "$3"; fi; }

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT
unset CLAUDE_CONFIG_DIR
export HOME="$TMP/home"
mkdir -p "$HOME"

M='${CLAUDE_SKILL_DIR}/scripts/outlook-mail.sh'
C='${CLAUDE_SKILL_DIR}/scripts/outlook-calendar.sh'
QM='"${CLAUDE_SKILL_DIR}/scripts/outlook-mail.sh"'
QC='"${CLAUDE_SKILL_DIR}/scripts/outlook-calendar.sh"'
ID="AAMkAGI2TG93AAA="

# Commands that send something. Each must be caught by the hook and by at
# least one ask rule.
SENDS=(
    "$M send $ID"
    "$M --account work send $ID"
    "$M -a work send $ID"
    "$QM send $ID"
    "$QM --account work send $ID"
    "OUTLOOK_ACCOUNT=work bash $M send $ID"
    "$M draft a@example.com Hi Body && $M send $ID"
    "$C invite $ID \"a@example.com, b@example.com\""
    "$C --account work invite $ID a@example.com optional"
    "$QC invite $ID a@example.com"
    "$C respond $ID accept"
    "$QC respond $ID decline \"Sorry, a clash\""
    "$C cancel $ID \"Postponed\""
    "$C -a work cancel $ID"
    "$C create Kickoff 2026-10-01T10:00 2026-10-01T11:00 \"\" \"a@example.com\" --send-invites"
    "$C create --send-invites Kickoff 2026-10-01T10:00 2026-10-01T11:00 Teams a@example.com"
    "$C update $ID start 2026-10-01T09:00 --notify-attendees"
)

# Commands that send nothing. The hook must stay silent and no ask rule may
# match, or the rules would nag on every draft and read.
QUIET=(
    "$M draft a@example.com Hi Body"
    "$M mddraft a@example.com Hi Body"
    "$M reply $ID \"Thanks\""
    "$M forward $ID a@example.com"
    "$M update $ID subject \"send it\""
    "$M sent 20"
    "$M read $ID"
    "$M inbox 10"
    "$M search \"please send the invoice\""
    "$C create Focus 2026-10-01T10:00 2026-10-01T11:00 Desk"
    "$C events"
    "$C read $ID"
    "$C update $ID subject \"New title\""
    "$C delete $ID"
    "$C search cancel"
    "ls -la"
    "git commit -m 'send the invites'"
)

hook() {  # hook <json> - prints the hook's stdout, then "rc=<exit code>"
    local out rc
    out=$(printf '%s' "$1" | bash "$HOOK" 2>/dev/null); rc=$?
    printf '%s\nrc=%s' "$out" "$rc"
}
# What hook prints when it stays out of the way: no output, exit 0.
SILENT=$'\nrc=0'
bash_input() { jq -cn --arg c "$1" '{tool_name: "Bash", tool_input: {command: $c}}'; }
decision() { printf '%s' "$1" | sed '$d' | jq -r '.hookSpecificOutput.permissionDecision // empty' 2>/dev/null; }

########################################
# The hook
########################################
for cmd in "${SENDS[@]}"; do
    out=$(hook "$(bash_input "$cmd")")
    eq "hook asks: $cmd" "ask" "$(decision "$out")"
    eq "hook exits 0: $cmd" "rc=0" "$(printf '%s' "$out" | tail -1)"
done
for cmd in "${QUIET[@]}"; do
    out=$(hook "$(bash_input "$cmd")")
    eq "hook is silent: $cmd" "$SILENT" "$out"
done

out=$(hook "$(bash_input "$M send $ID")")
eq "the ask names the hook event" "PreToolUse" \
   "$(printf '%s' "$out" | sed '$d' | jq -r '.hookSpecificOutput.hookEventName')"
has "the ask says what would be sent" "sends an email" "$(printf '%s' "$out" | sed '$d' | jq -r '.hookSpecificOutput.permissionDecisionReason')"
out=$(hook "$(bash_input "$C cancel $ID")")
has "a cancel says who hears about it" "every attendee" "$(printf '%s' "$out" | sed '$d' | jq -r '.hookSpecificOutput.permissionDecisionReason')"

# Not a Bash call, not JSON, or no command: silent, and never an error.
eq "hook ignores other tools" "$SILENT" \
   "$(hook "$(jq -cn --arg c "$M send $ID" '{tool_name: "Write", tool_input: {file_path: "x", content: $c}}')")"
eq "hook ignores input that is not JSON" "$SILENT" "$(hook "outlook-mail.sh send {not json")"
eq "hook ignores empty input" "$SILENT" "$(hook "")"
eq "hook ignores a Bash call with no command" "$SILENT" \
   "$(hook '{"tool_name":"Bash","tool_input":{"description":"outlook-mail.sh send"}}')"

# Without jq it cannot read the call, so it steps aside rather than fail.
mkdir -p "$TMP/nojq"
for t in bash cat; do ln -s "$(command -v "$t")" "$TMP/nojq/$t"; done
out=$(printf '%s' "$(bash_input "$M send $ID")" | PATH="$TMP/nojq" "$TMP/nojq/bash" "$HOOK" 2>&1); rc=$?
eq "hook without jq exits 0 and prints nothing" "$SILENT" "$out
rc=$rc"

# The hook file the plugin loads points at this script.
eq "hooks.json runs send-gate.sh on Bash" "Bash|bash \"\${CLAUDE_PLUGIN_ROOT}/hooks/send-gate.sh\"" \
   "$(jq -r '.hooks.PreToolUse[0] | "\(.matcher)|\(.hooks[0].command)"' "$REPO/hooks/hooks.json")"

########################################
# The ask rules
########################################
# Claude Code's `*` in a Bash rule matches any run of characters, spaces
# included. A bash glob does the same, so each rule is tried as a glob here.
# This approximates Claude Code's matcher; it does not run it.
rule_matches() {  # rule_matches <command> - prints the first rule that matches
    local cmd="$1" rule pat
    while IFS= read -r rule; do
        pat="${rule#Bash(}"; pat="${pat%)}"
        # shellcheck disable=SC2053 # the rule is a glob on purpose
        if [[ $cmd == $pat ]]; then printf '%s' "$rule"; return 0; fi
    done < <(jq -r '.permissions.ask[]' "$RULES")
    return 1
}
eq "every ask rule is a Bash rule" "0" "$(jq '[.permissions.ask[] | select(startswith("Bash(") | not)] | length' "$RULES")"
for cmd in "${SENDS[@]}"; do
    if rule=$(rule_matches "$cmd"); then eq "an ask rule matches: $cmd" ok ok
    else eq "an ask rule matches: $cmd" "a matching rule" "none"; fi
done
for cmd in "${QUIET[@]}"; do
    rule=$(rule_matches "$cmd") || rule=""
    eq "no ask rule matches: $cmd" "" "$rule"
done

########################################
# install-ask-rules.sh
########################################
S="$TMP/settings.json"
N=$(jq '.permissions.ask | length' "$RULES")

# No --yes and no terminal: nothing is created or changed.
out=$(bash "$HELPER" --settings "$S" < /dev/null 2>&1); rc=$?
eq "without consent it exits 0" "0" "$rc"
has "without consent it says nothing changed" "nothing was changed" "$out"
eq "without consent no settings file is created" "absent" "$([ -e "$S" ] && echo present || echo absent)"

printf '%s\n' '{"model":"x","permissions":{"allow":["Bash(ls:*)"],"ask":["Bash(rm *)"]},"hooks":{}}' > "$S"
before=$(cat "$S")
bash "$HELPER" --settings "$S" < /dev/null > /dev/null 2>&1
eq "without consent an existing file is untouched" "$before" "$(cat "$S")"

# A "no" at the prompt changes nothing either. `script` gives the helper a
# terminal to ask on where it is available.
if command -v script >/dev/null 2>&1 && script -qec true /dev/null >/dev/null 2>&1; then
    printf 'n\n' | script -qec "bash '$HELPER' --settings '$S'" /dev/null > "$TMP/tty.out" 2>&1
    eq "answering n at the prompt changes nothing" "$before" "$(cat "$S")"
    has "the prompt lists the rules first" "outlook-mail.sh send" "$(cat "$TMP/tty.out")"
fi

# --yes adds every missing rule and keeps everything else.
out=$(bash "$HELPER" --yes --settings "$S" 2>&1); rc=$?
eq "--yes exits 0" "0" "$rc"
eq "--yes adds every rule" "$N" "$(jq --slurpfile r "$RULES" '[.permissions.ask[] | select(. as $x | $r[0].permissions.ask | index($x))] | length' "$S")"
eq "--yes keeps the user's own ask rule first" "Bash(rm *)" "$(jq -r '.permissions.ask[0]' "$S")"
eq "--yes keeps the rest of the file" 'x|["Bash(ls:*)"]|{}' "$(jq -c -r '"\(.model)|\(.permissions.allow | tojson)|\(.hooks | tojson)"' "$S")"
eq "--yes saves the old file beside it" "$before" "$(cat "$S.before-outlook-ask-rules")"

after=$(cat "$S")
out=$(bash "$HELPER" --yes --settings "$S" 2>&1)
has "a second run has nothing to do" "Nothing to do" "$out"
eq "a second run changes nothing" "$after" "$(cat "$S")"
eq "a second run adds no duplicates" "$((N + 1))" "$(jq '.permissions.ask | length' "$S")"

# Only the missing rules are added.
jq --slurpfile r "$RULES" '.permissions.ask = [$r[0].permissions.ask[0]]' "$S" > "$S.tmp" && mv "$S.tmp" "$S"
bash "$HELPER" --yes --settings "$S" > /dev/null 2>&1
eq "a partial set is completed, not duplicated" "$N" "$(jq '.permissions.ask | length' "$S")"

# A missing file is created with just the rules.
S2="$TMP/new/settings.json"
bash "$HELPER" --yes --settings "$S2" > /dev/null 2>&1
eq "--yes creates a missing settings file" "$N" "$(jq '.permissions.ask | length' "$S2" 2>/dev/null)"

# A file it cannot safely merge into is left alone.
S3="$TMP/bad.json"
printf '%s' '{"permissions":{"ask":"Bash(x)"}}' > "$S3"
out=$(bash "$HELPER" --yes --settings "$S3" 2>&1); rc=$?
eq "a malformed settings file is refused" "1" "$rc"
eq "a malformed settings file is untouched" '{"permissions":{"ask":"Bash(x)"}}' "$(cat "$S3")"
printf 'not json' > "$S3"
bash "$HELPER" --yes --settings "$S3" > /dev/null 2>&1; rc=$?
eq "a settings file that is not JSON is refused" "1" "$rc"
eq "a settings file that is not JSON is untouched" "not json" "$(cat "$S3")"

# A symlinked settings file is written through, not replaced.
mkdir -p "$TMP/dotfiles"
printf '{}' > "$TMP/dotfiles/settings.json"
ln -s "$TMP/dotfiles/settings.json" "$TMP/link.json"
bash "$HELPER" --yes --settings "$TMP/link.json" > /dev/null 2>&1
eq "a symlinked settings file stays a symlink" "link" "$([ -L "$TMP/link.json" ] && echo link || echo file)"
eq "the rules land in the symlink's target" "$N" "$(jq '.permissions.ask | length' "$TMP/dotfiles/settings.json")"

# With no --settings it uses ~/.claude/settings.json - here the throwaway HOME.
bash "$HELPER" --yes > /dev/null 2>&1
eq "the default file is ~/.claude/settings.json" "$N" "$(jq '.permissions.ask | length' "$HOME/.claude/settings.json")"

########################################
# install.sh only passes consent it was given
########################################
# Pull offer_ask_rules out of install.sh and run it against a stub helper that
# logs its arguments.
fn=$(awk '/^offer_ask_rules\(\) \{/{f=1} f{print} f && /^}/{exit}' "$REPO/install.sh")
eq "install.sh defines offer_ask_rules" "1" "$([ -n "$fn" ] && echo 1 || echo 0)"
eval "$fn"
STUB="$TMP/stub-helper.sh"
printf '#!/bin/bash\nprintf "called:%%s\\n" "$*" >> "%s"\n' "$TMP/stub.log" > "$STUB"

: > "$TMP/stub.log"; offer_ask_rules yes "$STUB" > /dev/null < /dev/null
eq "--ask-rules runs the helper with --yes" "called:--yes" "$(cat "$TMP/stub.log")"
: > "$TMP/stub.log"; out=$(offer_ask_rules no "$STUB" < /dev/null)
eq "--no-ask-rules never runs the helper" "" "$(cat "$TMP/stub.log")"
: > "$TMP/stub.log"; out=$(offer_ask_rules offer "$STUB" < /dev/null)
eq "with no flag and no terminal the helper never runs" "" "$(cat "$TMP/stub.log")"
has "with no flag and no terminal it says how to add them" "$STUB" "$out"
if command -v script >/dev/null 2>&1 && script -qec true /dev/null >/dev/null 2>&1; then
    : > "$TMP/stub.log"
    script -qec "bash -c '$(declare -f offer_ask_rules); offer_ask_rules offer \"$STUB\"'" /dev/null > /dev/null 2>&1
    eq "with a terminal the helper runs without --yes, so it asks" "called:" "$(cat "$TMP/stub.log")"
fi

echo "-----------------------------"
printf 'PASS=%d FAIL=%d\n' "$PASS" "$FAIL"
[ "$FAIL" -eq 0 ]
