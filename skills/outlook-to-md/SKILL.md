---
name: outlook-to-md
description: Turn Outlook mail into organised markdown archives - from a PST export, or from live mail exported by the outlook skill. Use when needing to convert PST files to markdown, extract email archives, process Outlook exports, create searchable email collections, or keep an existing archive current. Trigger on phrases like "extract pst", "convert pst", "pst to markdown", "outlook to markdown", "email archive", "extract outlook", "update my email archive".
---

# Outlook Email to Markdown

Turn Outlook mail into an organised, integrity-verified archive of markdown files, raw email backups, and attachments. Reads a PST export, or a directory of `.eml` files - which is how live mail arrives, via the sibling `outlook` skill's `export` verb. Supports full extraction and incremental append mode, so one archive can span both.

## Prerequisites

- Python virtual environment set up (run setup.sh if not done)
- At least one of: `libratom` (Python) or `readpst` (system tool from pst-utils)

### First-Time Setup

```bash
# Set up Python environment (one-time)
${CLAUDE_SKILL_DIR}/setup.sh
```

### System Dependencies (optional fallback)

If libratom installation fails, install readpst as a fallback:

```bash
# Ubuntu/Debian
sudo apt install pst-utils

# macOS
brew install libpst
```

## Extraction Operations

### Full Extraction

Extract all emails from a PST file into markdown:

```bash
# Basic extraction
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py /path/to/file.pst /path/to/output/

# Verbose output with progress
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py /path/to/file.pst /path/to/output/ --verbose

# Include deleted items
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py /path/to/file.pst /path/to/output/ --include-deleted --verbose

# Set timezone for date display
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py /path/to/file.pst /path/to/output/ --timezone "Europe/London"

# Fix MAILER-DAEMON sent items (provide the PST owner's email)
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py /path/to/file.pst /path/to/output/ --owner-email "user@example.com"
```

### Incremental Extraction (Append Mode)

Add only new emails (skips already-extracted messages by Message-ID):

```bash
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py /path/to/file.pst /path/to/output/ --append --verbose
```

### Keeping an Archive Current from Live Mail

A PST is a snapshot. To carry an archive forward, export new mail with the
sibling `outlook` skill and append it — the two produce the same shape.

```bash
# 1. Export live mail as .eml (needs outlook configured)
${CLAUDE_SKILL_DIR}/../outlook/scripts/outlook-mail.sh \
  export "Inbox/Clients" ./staging/ --since 2026-07-01

# --count N caps how many messages export writes, newest first (default 1000)

# 2. Append it to the existing archive
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py \
  ./staging/ ./archive/ --append
```

Deduplication is by `Message-ID`, so a `--since` window that overlaps what is
already archived costs bandwidth and nothing else. Graph-sourced mail is
recorded under the `pst_folder` index column like any other — the column means
"the folder this message came from", and always did.

This guarantee depends on the message actually carrying a `Message-ID`
header. Received mail always has one, but a message with none (some drafts,
some malformed mail) has no key to dedupe on and is re-archived as a fresh
entry on every overlapping run. Narrow in practice, but real — there is no
content-hash fallback.

### Extract from Pre-Extracted .eml Directory

If emails were already extracted with readpst elsewhere, point at the directory:

```bash
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py /path/to/eml-directory/ /path/to/output/
```

## Output Structure

```
output/
├── emails/
│   ├── FolderName/
│   │   ├── 2023-01-15_093042_from-john.smith_to-jane.doe_RE-Subject/
│   │   │   ├── email.md              # Formatted markdown with YAML frontmatter
│   │   │   ├── email.eml             # Raw original email (RFC 822)
│   │   │   ├── attachment_001_doc.pdf # Extracted attachments
│   │   │   └── checksums.sha256      # Per-email integrity hashes
│   │   └── .../
│   └── .../
├── index.csv                          # Machine-readable master index
├── index.md                           # Human-readable index with timeline
├── extraction_log.txt                 # Processing log with statistics
└── manifest.sha256                    # Master integrity manifest
```

### Email Markdown Format

Each `email.md` contains:
- **YAML frontmatter**: message_id, date, from, to, cc, subject, attachments with SHA256 hashes
- **Formatted body**: HTML converted to markdown, or plain text preserved
- **Attachment links**: Relative links to extracted files with sizes
- **Original headers**: Full RFC 822 headers in code block

### Index Files

- **index.csv**: All emails with date, sender, recipient, subject, folder, attachment count
- **index.md**: Timeline view grouped by year/month with links to each email

## CLI Reference

```
outlook_to_md.py [-h] [--include-deleted] [--timezone TZ] [--verbose] [--append] [--owner-email EMAIL] pst_file output_dir
```

| Argument | Description |
|----------|-------------|
| `pst_file` | Path to PST file, or directory of pre-extracted .eml files |
| `output_dir` | Output directory (created if needed) |
| `--include-deleted` | Include deleted items from PST |
| `--timezone TZ` | Target timezone for dates (default: UTC) |
| `--verbose`, `-v` | Verbose output with per-email logging |
| `--append` | Skip emails already in archive (by Message-ID) |
| `--owner-email EMAIL` | PST owner's email (fixes MAILER-DAEMON in sent items) |

## Extraction Backends

A directory input (pre-extracted `.eml` files) is always handled directly,
checked before any backend regardless of what is installed. A PST file falls
back in this order:

1. **libratom** (Python) - preferred, installed via requirements.txt
2. **readpst** (system CLI) - fallback, from pst-utils package

## Integrity Verification

Every extraction produces a verifiable chain of custody:

1. Each email folder has `checksums.sha256` (SHA256 of all its files)
2. `manifest.sha256` hashes all checksum files plus the index
3. Source PST SHA256 is recorded in the manifest

To verify: `sha256sum -c manifest.sha256`

## Workflow: Extract and Search

Extract, then search the markdown with ripgrep. There is no semantic index - the output is plain files on disk, so use whatever search you already trust.

```bash
# Step 1: Extract
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py archive.pst ./email-output/ --verbose

# Step 2: Search the output
rg -i "settlement agreement" ./email-output/ -l
```

Grep is exact, so search on names, addresses and distinctive phrases rather than concepts.

## Error Handling

- **"readpst not found"**: Install pst-utils or ensure libratom is installed via setup.sh
- **Corrupt emails**: Logged to extraction_log.txt, processing continues
- **Encoding issues**: Falls back through UTF-8 → latin-1 → raw bytes
- **Duplicate timestamps**: Appended with -001, -002 suffixes
- **Path too long**: Subject truncated, uniqueness preserved

## Performance

| Scenario | Approximate Speed |
|----------|-------------------|
| Emails without attachments | ~5,000/hour |
| Emails with attachments | ~2,000/hour |

A typical 300MB PST (~1,000-3,000 emails) processes in 5-15 minutes.
