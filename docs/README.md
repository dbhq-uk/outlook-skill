# Outlook skill pack documentation

Microsoft 365 mail, calendar and archives from the terminal. Start at the [project
README](../README.md) if you have not read it.

These pages are organised by what you are trying to do rather than by feature, and each one is
a single kind of document, so the heading tells you whether it will help.

## Learning

**[Getting started](getting-started.md)** &nbsp; A tutorial. Install, authenticate, read the
inbox and send one message, in about ten minutes. One happy path, no options.

## Doing

**[Handling a thread](guides/email.md)** &nbsp; Reading a message properly, replying without
losing a recipient, trimming to sender-only, aliases, attachments, forwards and chasers, and
filing what is done.

**[Running the calendar](guides/calendar.md)** &nbsp; Getting the timezone right, the two-step
invite flow, changing and cancelling, answering invitations, and finding a slot.

**[Building a markdown archive](guides/archiving.md)** &nbsp; Extracting a PST, verifying the
result, and keeping the archive current from live mail.

**[Accounts, tokens and things going wrong](guides/accounts.md)** &nbsp; Running more than one
mailbox, checking a connection, and every way authentication fails.

## Looking things up

**[Reference](reference.md)** &nbsp; Every command, argument, flag, environment variable and
file, for all four scripts and `outlook_to_md.py`.

**[Manual Azure setup](../skills/outlook/references/setup.md)** &nbsp; The app
registration and OAuth exchange step by step, for when you would rather not let `setup.sh` do
it. Ships inside the skill so it is available offline.

## Understanding

**[How the pack is built](architecture.md)** &nbsp; Why there are two skills, what the safety
rails are and why they are the defaults, the scope deliberately not requested, short IDs,
email HTML that survives Outlook, and why there is one PST reader.

## For contributors

- [dev-setup.md](dev-setup.md), running the pack from source with live edits
- [AGENTS.md](../AGENTS.md), the working brief for anyone changing the code, human or otherwise
- [CONTRIBUTING.md](../CONTRIBUTING.md)
