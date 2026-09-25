# Building a markdown archive

Turning a PST export into markdown you can grep, then keeping that archive current from live
mail. This is `outlook-to-md`, with `outlook` doing the fetching.

Two paths recur below:

```bash
MD=~/.claude/skills/outlook-to-md
MAIL=~/.claude/skills/outlook/scripts/outlook-mail.sh
```

## One-time setup

```bash
$MD/setup.sh
```

That builds a virtualenv next to the skill. Nothing else is needed - `outlook-to-md` makes no
network calls and holds no credentials. It reads files on disk.

To read a `.pst` you also need `readpst`: install `pst-utils` (Debian/Ubuntu) or `libpst`
(Homebrew). `setup.sh` says whether it found it. A directory of `.eml` files needs nothing extra.

## Extract a PST

```bash
$MD/.venv/bin/python $MD/scripts/outlook_to_md.py archive.pst ./out/ --verbose
```

Useful additions:

```bash
--timezone "Europe/London"          # dates render in this zone; default is the sender's offset
--owner-email "you@example.com"     # fixes MAILER-DAEMON senders in sent items
--include-deleted                   # deleted items too (readpst -D)
```

Reckon on roughly 5,000 emails an hour without attachments and 2,000 with; a 300 MB PST of
one to three thousand messages takes five to fifteen minutes.

What you get:

```
out/
├── emails/<Folder>/<date>_<time>_from-x_to-y_Subject/
│   ├── email.md              # YAML frontmatter + body as markdown
│   ├── email.eml             # the original, RFC 822
│   ├── attachment_001_*.pdf
│   └── checksums.sha256
├── index.csv                 # every email, machine-readable
├── index.md                  # timeline by year and month
├── extraction_log.txt
└── manifest.sha256
```

## Verify it

```bash
cd out && sha256sum -c manifest.sha256
```

Each email folder hashes its own files; the manifest hashes every one of those checksum files
plus the index, and records the source PST's own SHA256. So one command proves the whole
extraction is what came out of the PST, and the PST is the one you were given. That chain is
the point of the tool - it is built for archives someone will later ask you to stand behind.

The manifest lists **every** source the archive has ever been built or appended from, not
just the most recent one, with a hash and a timestamp each. Re-running an identical append
does not add a duplicate entry.

## Search it

There is no index and no embedding. It is files on disk, so use the search you already trust:

```bash
rg -i "settlement agreement" ./out/ -l
```

Grep is exact. Search on names, addresses and distinctive phrases rather than concepts.

## Keep it current

A PST is a snapshot of the day it was exported. To carry the archive forward, pull new mail
with `outlook` and append it - both skills write the same shape.

```bash
# 1. Export a folder as .eml
$MAIL export "Inbox/Clients" ./staging/ --since 2026-07-01

# 2. Append it
$MD/.venv/bin/python $MD/scripts/outlook_to_md.py ./staging/ ./archive/ --append
```

`--append` skips anything already in the archive, matched on `Message-ID`, so an overlapping
`--since` window costs bandwidth and nothing else. Take the date from the archive's newest
entry and overlap deliberately rather than trying to be exact.

**Pass `--append` when pointing at an existing archive.** Without it the run refuses and
changes nothing. To rebuild an archive from scratch, pass `--overwrite`: it deletes the
archive's `emails/` folder and its index, manifest and log files, then extracts afresh. Other
files in the output folder are left alone.

`export` also takes `--count N` to cap how many messages it writes, newest first, defaulting
to 1000.

The staging directory's layout becomes the archive's folder grouping, so `export
"Inbox/Clients"` lands under `emails/Inbox/Clients/`. Graph-sourced mail is recorded in the
index's `pst_folder` column like anything else - the column has always meant "the folder this
message came from".

**The one gap worth knowing.** Deduplication needs a `Message-ID` header. Received mail always
has one; some drafts and some malformed messages do not, and there is no content-hash
fallback. A message with no `Message-ID` is re-archived as a fresh entry on every overlapping
run. Narrow in practice, but real.

## Running it on a schedule

Nothing here needs a person. A weekly job that exports the last fortnight and appends it will
converge on the right archive regardless of overlap:

```bash
$MAIL export "Inbox/Clients" /tmp/staging --since "$(date -d '14 days ago' +%Y-%m-%d)"
$MD/.venv/bin/python $MD/scripts/outlook_to_md.py /tmp/staging ./archive/ --append
rm -rf /tmp/staging
```

Re-run `sha256sum -c manifest.sha256` after an append - the manifest is rewritten, and a
failure there is the signal that something outside the tool has touched the archive.
