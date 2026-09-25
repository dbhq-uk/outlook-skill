# Running the calendar

Viewing a diary, booking a meeting without accidentally inviting anyone, answering
invitations, and getting the timezone right.

`calendar.sh` below is
`~/.claude/skills/outlook/scripts/outlook-calendar.sh`.

## Fix the timezone first

Every time this script prints or accepts is wall-clock in one zone, resolved in this order:
`OUTLOOK_TZ`, then `/etc/timezone`, then `timedatectl`, then the `/etc/localtime` symlink,
and `Europe/London` if none of those answer.

On a laptop that is usually right. On a server, container or CI box it is UTC while you are
not, and a 14:00 London meeting is reported as 13:00 - same instant, wrong wall-clock, missed
meeting. When the resolved zone is UTC and you have not set `OUTLOOK_TZ`, the script says so
on stderr before every command. Believe it:

```bash
export OUTLOOK_TZ=Europe/London
```

The script cannot read the mailbox's own timezone, because that needs a
`MailboxSettings.Read` scope the app deliberately does not ask for. So it warns instead of
guessing.

## Looking at the diary

```bash
calendar.sh today
calendar.sh week
calendar.sh day 2026-08-13
calendar.sh events 20                    # next N in the coming year, each occurrence listed
calendar.sh search "board meeting"       # next 90 days by default
calendar.sh search "dentist" 365
calendar.sh read <event-id>
```

Every line starts with a number and the event's short ID:

```
[1] QAAAElc2ZXJpZXMwOTA4 | 2026-09-08 09:00-10:00 | Weekly sync | Teams
```

That ID is what `read`, `respond`, `cancel`, `update` and `delete` take. A recurring meeting
appears once per occurrence, each with its own ID, so you can answer or cancel one occurrence
without touching the series. Listings follow Graph's pages to the end of the window rather
than stopping at 10.

`read` is the one that shows attendees, their responses, the body and the organiser - list
views do not.

## Booking a meeting

Do this in two steps, always, unless the exact attendee list has already been agreed.

**Step one, create the event with no attendees. Nothing is sent to anyone:**

```bash
calendar.sh create "Project kickoff" "2026-08-13T14:00" "2026-08-13T15:00" "Teams"
```

Times are `YYYY-MM-DDTHH:MM` in your configured zone. At this point the event exists in your
own calendar and no one else knows about it, which is the calendar equivalent of a draft.

**Step two, once the attendee list has been confirmed, send the invitations:**

```bash
calendar.sh invite <event-id> "a@example.com, b@example.com"
calendar.sh invite <event-id> "c@example.com" optional
```

That is the moment mail goes out. Re-inviting an address already on the event is a no-op,
deduped case-insensitively, so `invite` is safe to run again to add someone.

`create` does accept attendees as a sixth argument - `create <subject> <start> <end>
[location] [attendees]`, with `""` for an absent location. That form sends invitations
**immediately on creation**, with no gap in which to check the list. Use it only when the
exact addresses have already been approved.

A one-hour event with no location, for when the meeting is with yourself:

```bash
calendar.sh quick "Write the board pack" "2026-08-13T09:00"
```

## Changing and cancelling

```bash
calendar.sh update <event-id> subject "Project kickoff (rescheduled)"
calendar.sh update <event-id> start "2026-08-14T14:00"
```

Updatable fields are `subject`, `location`, `start` and `end`.

Moving an event to a different day takes two calls, and the order matters: Graph rejects a
`start` later than the event's current `end`, and an `end` earlier than its current `start`.
Set whichever bound keeps start before end first.

Cancelling and deleting are not the same operation:

```bash
calendar.sh cancel <event-id> "Postponed - new invite to follow"
calendar.sh delete <event-id>
```

`cancel` is for meetings you organise: it withdraws the meeting and tells every attendee, with
your comment. `delete` removes the event silently. Deleting a meeting other people have in
their calendars leaves it in theirs - use `cancel` whenever anyone else is involved.

## Answering an invitation

```bash
calendar.sh respond <event-id> accept
calendar.sh respond <event-id> tentative
calendar.sh respond <event-id> decline "Sorry, I have a clash"
```

The organiser is notified either way. The comment is optional and only worth adding when it
says something the organiser cannot infer.

## Finding a slot

```bash
calendar.sh free "2026-08-13T09:00" "2026-08-13T17:00"
```

Either "You are FREE during this time period" or a list of what is in the way, with times and
IDs. Events shown as free (reminders, markers) and cancelled events do not count.
Both bounds are converted to offset-qualified ISO before they reach Graph, so a window given
in local wall-clock is not silently read as UTC and shifted an hour in summer. Worth running
before you propose a time rather than after someone declines.
