# Security

## Reporting a vulnerability

Email <dan@dbhq.uk> rather than opening a public issue. Include what you found,
how to reproduce it, and what an attacker could do with it. You will get a first
response within 48 hours.

## What this skill does

This skill reads and sends mail and manages calendar events through the
Microsoft Graph API, and archives PST exports to markdown. It holds OAuth tokens
for a real mailbox, which makes it the most security-sensitive skill in the
collection. Read this section before installing it.

### Network

**Two hosts, both Microsoft:**

| Host | Purpose |
|------|---------|
| `login.microsoftonline.com` | OAuth device-code sign-in and token refresh |
| `graph.microsoft.com` | Mail, calendar and category operations |

There is no DBHQ server in the path, no proxy and no telemetry. Your mail moves
between your machine and Microsoft directly.

### Credentials

Tokens live under `~/.dbhq/outlook/<account>/`, with permissions set explicitly
rather than left to the umask:

```
~/.dbhq/outlook/                            700   base directory
~/.dbhq/outlook/<account>/                  700   per-account directory
~/.dbhq/outlook/<account>/config.json       600   app registration details
~/.dbhq/outlook/<account>/credentials.json  600   access and refresh tokens
```

The credentials file is rewritten at `600` on every token refresh, not only at
setup, so the permissions cannot drift over the life of the install. The new
file is written beside the old one and renamed into place, and only when
Microsoft's response carries an access token. A refresh that fails, times out or
gets a non-JSON answer leaves the file unchanged.

Multiple accounts are isolated in separate directories and selected with
`--account` or `OUTLOOK_ACCOUNT`.

**Revoking access:** these are OAuth tokens against your own Azure app
registration. Revoke sessions in Entra ID (Azure AD) to cut access immediately.
Deleting `~/.dbhq/outlook` removes the local copy but does not revoke the
grant - do both.

### Scope of access

The skill can read, send and move mail, and create and delete calendar events.
That is the point of it, but it means an agent running this skill can send email
as you. `OUTLOOK_READ_ONLY=1` stops every command that writes or sends, and under
Claude Code the plugin's hook and the ask rules `install.sh` offers make it ask
you before a command that sends. See `docs/architecture.md`, "The send gap", for
what each of these does not cover. The permissions granted are those on the app registration you create
during setup: grant only the Graph scopes you actually need, and prefer
delegated over application permissions.

### On disk

- Installs into `~/.claude/skills/outlook` or `~/.codex`
- Reads and writes `~/.dbhq/outlook/` only
- `outlook-to-md` processes PST files **entirely locally** - a PST is never
  uploaded anywhere

### On the "sensitive file access" findings

Automated scanners flag the lines in `SKILL.md` that name
`~/.dbhq/outlook/<account>/` as "access to home directory dotfiles". Those are
sentences of documentation, not code - they describe where credentials live so
you can find, inspect and delete them.

The remediation such scanners advise, restricting the files to owner-only
access, is what the table above already does. Documenting the location is a
deliberate choice: a credential store you cannot find is not safer, only harder
to audit. The paths stay in the documentation.

## Note for automated scanners

`skills/outlook/tests/helpers_test.sh` uses `eval "$(extract_fn <name>)"`,
which static analysis flags as command injection. It is not.

The harness extracts individual named functions out of
`outlook-mail.sh` with `awk` and evaluates only those, so the pure logic
helpers can be unit-tested offline without a Microsoft account or a network. The
alternative - sourcing the script - would execute 2,500 lines of top-level
config loading and credential checks and fail without a configured mailbox.

The argument to `extract_fn` is a hardcoded function name in every one of its
call sites. No external, user-supplied or network-derived input reaches it, and
the file it reads is the skill's own script on local disk. The file is test-only
and is not on any runtime path.
