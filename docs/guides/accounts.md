# Accounts, tokens and things going wrong

Running more than one mailbox, and what to do when authentication stops working.

## More than one mailbox

Each account keeps its own credentials under `~/.dbhq/outlook/<account>/`. Add one:

```bash
~/.claude/skills/outlook/scripts/outlook-setup.sh --account work
```

That reuses the Azure app registration you already have if it finds one, so a second mailbox
is a sign-in rather than a fresh registration.

Every script resolves the active account the same way, in this order:

1. `--account <name>` or `-a <name>`, before the command
2. the `OUTLOOK_ACCOUNT` environment variable
3. `default`

```bash
mail.sh -a work inbox
OUTLOOK_ACCOUNT=work mail.sh inbox
~/.claude/skills/outlook/scripts/outlook-token.sh list
```

An install predating multi-account support keeps its files flat at
`~/.dbhq/outlook/{config,credentials,id_cache}.json`. The first run of any script moves them
into `~/.dbhq/outlook/default/`. Nothing to do, but do not be surprised.

## Checking the connection

```bash
token.sh test        # end-to-end: token, Graph call, inbox counts
token.sh status      # whether the token still works, and whose mailbox it is
token.sh refresh     # force a refresh
token.sh get         # print a valid access token, refreshed first if needed
token.sh list        # accounts on this machine
```

Tokens refresh automatically whenever a command needs one, so `refresh` is a diagnostic
rather than something to schedule.

`token.sh get` is how you make a Graph call the scripts do not cover. It goes through the same
expiry check as every other command, so the token it prints is good for at least a minute:

```bash
T=$(token.sh get)
curl -s -H "Authorization: Bearer $T" \
  "https://graph.microsoft.com/v1.0/me/messages/<id>?\$select=toRecipients,ccRecipients"
```

## When it stops working

**"Account 'x' not configured"** - there is no `credentials.json` under
`~/.dbhq/outlook/x/`. Run setup with `--account x`, or check you have not typo'd the name.

**"Could not refresh the access token" / "the token endpoint answered without an access token".**
The refresh request did not get a usable answer: no network, a timeout, or a proxy or gateway
error page. `credentials.json` is left exactly as it was, so run the command again once the
connection is back. Nothing needs re-authenticating.

**Token expired, and refresh also failed.** Refresh tokens last around 90 days of inactivity.
Past that, re-authenticate: `outlook-setup.sh`, or steps 6 and 7 of
[`references/setup.md`](../../skills/outlook/references/setup.md) by hand.

**"Invalid client secret", or a secret that has expired.** Setup no longer uses a secret: the
app is a public client and signs in with PKCE. An install from before that has a
`client_secret` in `config.json`, which it keeps sending, so it works until the secret expires.
To move it over, run setup again for that account:

```bash
~/.claude/skills/outlook/scripts/outlook-setup.sh --account <name>
```

When it offers to convert the app to a public client, say yes. That needs the Azure CLI and
rights on the app registration. It moves the redirect URI to the public client platform and
allows public client flows; old secrets are left in place, so other accounts on the same app
keep working. Setup then names any other account still on the old secret. Run setup for each
of those too. Once none is left, you can delete the secret in the Azure portal.

**"AADSTS50011: Reply URL does not match".** The redirect URI in the app registration must be
exactly `https://login.microsoftonline.com/common/oauth2/nativeclient`, on the **Mobile and
desktop applications** platform.

**"AADSTS7000218" (the request must contain a client secret).** The app is still a Web
(confidential) client. Run setup again and let it convert the app, or in the portal move the
redirect URI to **Mobile and desktop applications** and set **Allow public client flows** to
Yes.

**"Insufficient privileges".** The five delegated permissions - `Mail.ReadWrite`, `Mail.Send`,
`Calendars.ReadWrite`, `User.Read`, `offline_access` - are not all present or not all
consented. Add them, then sign out and in again to re-consent.

**Markdown commands failing.** `mddraft`, `mdreply`, `forward`, `followup` and `update mdbody`
shell out to `pandoc`. Install it.

**Calendar times an hour out.** Set `OUTLOOK_TZ` - see [running the
calendar](calendar.md#fix-the-timezone-first).

**A short ID that will not resolve.** Short IDs are cached by whichever listing produced them,
and a message that has been moved has a **new** ID in its new folder. Re-list from where the
message is now.

## What is stored, and where

| Path | Contents |
|---|---|
| `~/.dbhq/outlook/<account>/config.json` | Azure app client ID, tenant, scopes. A `client_secret` only on an install from before PKCE |
| `~/.dbhq/outlook/<account>/credentials.json` | OAuth access and refresh tokens |
| `~/.dbhq/outlook/<account>/id_cache.json` | Short ID to full Graph ID mapping |
| `~/.dbhq/outlook/<account>/event_id_cache.json` | Full event IDs from the last calendar listing |
| `~/.dbhq/outlook/<account>/.token.lock` | Empty lock file, held while a token refreshes |

The account directory is `700` and both credential files `600`, in your home directory and
nowhere else. Nothing is sent anywhere except
Microsoft Graph, directly from your machine. There are no secrets in this repository and
none in the skill.
