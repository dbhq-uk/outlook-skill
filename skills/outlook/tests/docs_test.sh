#!/bin/bash
# Checks that keep the agent-facing docs short, in house style and complete.
#
# SKILL.md loads in full every time the skill triggers, so it holds the rules
# and a one-line command index, and the full reference lives in
# references/commands.md, read only when needed. It had grown to over 4,000
# words, with em dashes against the house style and calendar verbs missing.
#
#   bash skills/outlook/tests/docs_test.sh
set -u

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_DIR="$(cd "$TESTS_DIR/.." && pwd)"
SKILLS_ROOT="$(cd "$SKILL_DIR/.." && pwd)"
SKILL="$SKILL_DIR/SKILL.md"
COMMANDS="$SKILL_DIR/references/commands.md"
MAX_WORDS=1500

PASS=0; FAIL=0
eq() { if [ "$2" = "$3" ]; then PASS=$((PASS+1)); printf 'ok   - %s\n' "$1";
       else FAIL=$((FAIL+1)); printf 'FAIL - %s\n   expected: %s\n   got:      %s\n' "$1" "$2" "$3"; fi; }

########################################
# Every SKILL.md in the repo: under the word limit, and no em or en dashes.
########################################
for f in "$SKILLS_ROOT"/*/SKILL.md; do
    name="$(basename "$(dirname "$f")")/SKILL.md"
    words=$(wc -w < "$f" | tr -d ' ')
    eq "$name is at most $MAX_WORDS words (has $words)" "1" "$([ "$words" -le "$MAX_WORDS" ] && echo 1 || echo 0)"
    dashes=$(grep -c $'\xe2\x80\x94\|\xe2\x80\x93' "$f" || true)
    eq "$name has no em or en dashes" "0" "$dashes"
done
eq "references/commands.md has no em or en dashes" "0" "$(grep -c $'\xe2\x80\x94\|\xe2\x80\x93' "$COMMANDS" || true)"

########################################
# Every skill folder stands alone: no path into another skill's folder.
########################################
# npx skills add installs each folder that holds a SKILL.md as a separate
# skill, and the installers skip a skill whose tools are missing. So the other
# skill may not be there, and ${CLAUDE_SKILL_DIR}/../<other>/... would point at
# nothing. Name the other skill instead, and say what to do without it.
for f in "$SKILLS_ROOT"/*/SKILL.md "$SKILLS_ROOT"/*/references/*.md; do
    [ -f "$f" ] || continue
    rel="${f#"$SKILLS_ROOT"/}"
    hits=$(grep -cE '\$\{?CLAUDE_SKILL_DIR\}?/\.\.' "$f" || true)
    eq "$rel has no path into another skill's folder" "0" "$hits"
done

########################################
# The description says what the skill is for, and what it is not for.
########################################
desc=$(awk '/^---$/ {n++; next} n == 1 && /^description:/ {sub(/^description:[[:space:]]*/, ""); print}' "$SKILL")
has_desc() { case "$desc" in *"$2"*) eq "$1" ok ok ;; *) eq "$1" "contains: $2" "$desc" ;; esac; }
has_desc "the description names Outlook" "Outlook"
has_desc "the description names Microsoft 365" "Microsoft 365"
has_desc "the description rules out Gmail" "Not for Gmail"
has_desc "the description sends PST files to outlook-to-md" "outlook-to-md"
eq "the description has no bare \"schedule\" trigger (the built-in schedule skill owns it)" "0" \
   "$(printf '%s' "$desc" | grep -ci '"schedule"' || true)"

########################################
# Every verb of every script is documented: in SKILL.md's index, or in
# references/commands.md. A new verb with no line in either fails here.
########################################
verbs_of() {  # the verbs in a script's top-level case: lines like "    verb)"
    grep -oE '^    [a-z][a-z|-]*\)' "$1" | tr -d ' )' | tr '|' '\n'
}
documented() {  # documented <script> <verb>
    local v="$2" pat
    pat="\`${v}[\` ]|[$][MCT] $v( |$)|$1 $v( |$)"
    grep -qE -- "$pat" "$SKILL" "$COMMANDS"
}
count=0
for script in outlook-mail outlook-calendar outlook-token; do
    verbs=$(verbs_of "$SKILL_DIR/scripts/$script.sh")
    [ "$script" = outlook-token ] && verbs="$verbs list"
    for v in $verbs; do
        count=$((count + 1))
        if documented "$script.sh" "$v"; then
            eq "$script $v is documented" ok ok
        else
            eq "$script $v is documented" "in SKILL.md or references/commands.md" "missing"
        fi
    done
done
eq "the verb check found the scripts' verbs" "1" "$([ "$count" -ge 60 ] && echo 1 || echo 0)"

# SKILL.md points at the reference, so the agent knows it exists.
eq "SKILL.md points to references/commands.md" "1" "$(grep -c 'references/commands.md' "$SKILL" | awk '{print ($1 > 0)}')"

echo
echo "PASS=$PASS FAIL=$FAIL"
[ "$FAIL" -eq 0 ]
