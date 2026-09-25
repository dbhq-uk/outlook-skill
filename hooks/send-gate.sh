#!/bin/bash
# PreToolUse hook: ask before a Bash command that sends mail, invitations or a
# meeting response through this skill.
#
# Claude Code runs it before every Bash tool call and passes the call as JSON on
# stdin. When the command runs one of the verbs below, it answers "ask", so
# Claude Code shows a permission prompt that names what would be sent. For any
# other command it prints nothing and exits 0, which leaves the decision to the
# normal permission rules.
#
#   outlook-mail.sh send
#   outlook-calendar.sh invite | respond | cancel
#   outlook-calendar.sh ... --send-invites        (create with attendees)
#   outlook-calendar.sh ... --notify-attendees    (update a meeting you organise)
#
# WHAT IT DOES NOT DO. It reads the command as written, so a script reached
# through a variable or a renamed copy is not caught. And Claude Code documents
# that an "ask" from a hook prompts in Manual and auto mode, but it does not
# document that one prompts in bypassPermissions mode. An explicit ask rule
# does prompt there, which is why hooks/ask-rules.json exists and install.sh
# offers it. The layers that hold without Claude Code are in the scripts: their
# refusals (create needs --send-invites, update needs --notify-attendees) and
# OUTLOOK_READ_ONLY=1. docs/architecture.md, "The send gap", has all four.
#
# A hook that errors must never block the user's work, so every failure path
# below is a silent exit 0.

input=$(cat)

# Cheap test first: most Bash calls have nothing to do with this skill.
case "$input" in
    *outlook-mail.sh*|*outlook-calendar.sh*) ;;
    *) exit 0 ;;
esac

command -v jq >/dev/null 2>&1 || exit 0
[ "$(printf '%s' "$input" | jq -r '.tool_name // empty' 2>/dev/null)" = "Bash" ] || exit 0
cmd=$(printf '%s' "$input" | jq -r '.tool_input.command // empty' 2>/dev/null) || exit 0
[ -n "$cmd" ] || exit 0

# The script path, an optional closing quote, an optional --account/-a <name>,
# then the verb as a whole word.
q="[\"']?"
sp='[[:space:]]+'
acct="((-a|--account)${sp}[^[:space:]]+${sp})?"
end="([[:space:]\"';&|)]|\$)"

outlook_send_reason() {
    local verb
    if [[ $1 =~ outlook-mail\.sh${q}${sp}${acct}${q}send${end} ]]; then
        echo "sends an email (outlook-mail.sh send)"
        return 0
    fi
    for verb in invite respond cancel; do
        if [[ $1 =~ outlook-calendar\.sh${q}${sp}${acct}${q}${verb}${end} ]]; then
            case "$verb" in
                invite)  echo "sends meeting invitations (outlook-calendar.sh invite)" ;;
                respond) echo "sends your response to the organiser (outlook-calendar.sh respond)" ;;
                cancel)  echo "sends a cancellation to every attendee (outlook-calendar.sh cancel)" ;;
            esac
            return 0
        fi
    done
    if [[ $1 =~ outlook-calendar\.sh.*--send-invites ]]; then
        echo "creates a meeting and invites its attendees at once (outlook-calendar.sh create --send-invites)"
        return 0
    fi
    if [[ $1 =~ outlook-calendar\.sh.*--notify-attendees ]]; then
        echo "changes a meeting and sends every attendee an update (outlook-calendar.sh update --notify-attendees)"
        return 0
    fi
    return 1
}

reason=$(outlook_send_reason "$cmd") || exit 0

jq -n --arg r "outlook: this command $reason. Check who it goes to before you approve it." '{
    hookSpecificOutput: {
        hookEventName: "PreToolUse",
        permissionDecision: "ask",
        permissionDecisionReason: $r
    }
}'
exit 0
