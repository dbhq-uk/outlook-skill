#!/usr/bin/env python3
"""
PST Email Extraction Tool

Extract emails from Outlook PST files into an organized archive of markdown
files, raw email backups, and attachments with integrity verification.

Usage:
    python outlook_to_md.py <pst_file> <output_dir>              # New archive
    python outlook_to_md.py <pst_file> <output_dir> --append     # Append new emails only
    python outlook_to_md.py <pst_file> <output_dir> --overwrite  # Replace an archive

The --append flag enables incremental extraction: it loads the existing index.csv
and skips any emails whose Message-ID is already in the archive. This lets you
update the PST file and re-run extraction to add only new emails.

Without --append or --overwrite, an output directory that already holds an
archive is refused rather than half-replaced.
"""

import argparse
import csv
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from email.utils import getaddresses
from pathlib import Path
from typing import Optional
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

# Optional dependencies with fallbacks
try:
    from dateutil import parser as date_parser

    HAS_DATEUTIL = True
except ImportError:
    HAS_DATEUTIL = False

try:
    from tqdm import tqdm

    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

    # Simple fallback for tqdm
    def tqdm(iterable, desc=None, **kwargs):
        if desc:
            print(f"{desc}...")
        return iterable


try:
    import html2text

    HAS_HTML2TEXT = True
except ImportError:
    HAS_HTML2TEXT = False


def sanitize_filename(text: str, max_length: int = 50) -> str:
    """Sanitize text for use in filenames."""
    if not text:
        return "unknown"
    # Replace spaces with hyphens
    text = text.replace(" ", "-")
    # Remove problematic characters
    text = re.sub(r'[<>:"/\\|?*@\[\]]', '', text)
    # Collapse multiple hyphens
    text = re.sub(r'-+', '-', text)
    # Remove leading/trailing hyphens
    text = text.strip('-')
    # Truncate
    if len(text) > max_length:
        text = text[:max_length].rstrip('-')
    return text or "unknown"


def sanitize_email(email: str) -> str:
    """Sanitize email address for filename use."""
    if not email:
        return "unknown"
    # Extract just the email part if in "Name <email>" format
    match = re.search(r'<([^>]+)>', email)
    if match:
        email = match.group(1)
    # Remove @ and other special chars but keep dots
    email = re.sub(r'[<>:"/\\|?*\[\]@]', '', email)
    return sanitize_filename(email, max_length=40)


def compute_sha256(filepath: Path) -> str:
    """Compute SHA256 hash of a file."""
    sha256 = hashlib.sha256()
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            sha256.update(chunk)
    return sha256.hexdigest()


def format_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}" if unit != 'B' else f"{size_bytes} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def html_to_markdown(html_content: str) -> str:
    """Convert HTML to Markdown."""
    if not html_content:
        return ""
    if HAS_HTML2TEXT:
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.ignore_images = False
        h.body_width = 0  # Don't wrap lines
        return h.handle(html_content)
    else:
        # Simple fallback: strip HTML tags
        import re

        text = re.sub(r'<br\s*/?>', '\n', html_content, flags=re.IGNORECASE)
        text = re.sub(r'<p\s*/?>', '\n\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</p>', '', text, flags=re.IGNORECASE)
        text = re.sub(r'<[^>]+>', '', text)
        text = re.sub(r'&nbsp;', ' ', text)
        text = re.sub(r'&lt;', '<', text)
        text = re.sub(r'&gt;', '>', text)
        text = re.sub(r'&amp;', '&', text)
        text = re.sub(r'&quot;', '"', text)
        return text.strip()


def parse_email_address(addr_str: str) -> tuple[str, str]:
    """Parse email address into (name, email) tuple."""
    if not addr_str:
        return ("", "")
    match = re.match(r'^"?([^"<]*)"?\s*<?([^>]*)>?$', addr_str.strip())
    if match:
        name = match.group(1).strip().strip('"')
        email = match.group(2).strip() or addr_str.strip()
        return (name, email)
    return ("", addr_str.strip())


# Characters that make a display name need quoting in an address (RFC 5322
# "specials"). A name with a comma in it, unquoted, reads as two addresses.
_ADDRESS_SPECIALS = set('()<>[]:;@\\,."')


def format_address(name: str, addr: str) -> str:
    """One address as a string: 'Name <addr>', with the name quoted if it needs it.

    Unlike email.utils.formataddr this never RFC 2047-encodes a non-ASCII name,
    because the result goes into markdown for people to read.
    """
    name = (name or '').strip()
    addr = (addr or '').strip()
    if not name:
        return addr
    if not addr:
        return name
    if any(c in _ADDRESS_SPECIALS for c in name):
        name = '"' + name.replace('\\', '\\\\').replace('"', '\\"') + '"'
    return f'{name} <{addr}>'


def split_addresses(value) -> list:
    """Split an address header (To, Cc, Bcc) into one string per address.

    Splitting on commas broke any display name with a comma in it:
    '"Jones, Ann" <ann@example.com>' became '"Jones' and 'Ann" <ann@example.com>'.
    A header parsed with policy.default carries its addresses already parsed,
    before any encoded name is decoded, so those are used when present.
    Anything else goes through email.utils.getaddresses. If neither finds an
    address, the header is kept whole rather than lost.
    """
    if not value:
        return []
    pairs = None
    addresses = getattr(value, 'addresses', None)
    if addresses is not None:
        try:
            pairs = [(a.display_name, a.addr_spec) for a in addresses]
        except Exception:
            pairs = None
    if pairs is None:
        pairs = getaddresses([str(value)])
    out = [format_address(name, addr) for name, addr in pairs if (name or '').strip() or (addr or '').strip()]
    if not out and str(value).strip():
        out = [str(value).strip()]
    return out


# Characters JSON leaves as they are but YAML does not accept in a document
# (C1 controls, the Unicode line and paragraph separators, byte-order marks,
# non-characters and lone surrogates). They are written as \u escapes.
_YAML_UNSAFE = re.compile('[\x7f-\x9f\u2028\u2029\ufeff\ufffe\uffff\ud800-\udfff]')


def yaml_str(value) -> str:
    """A YAML double-quoted scalar for any string.

    A JSON string is valid YAML, so json.dumps escapes quotes, backslashes and
    control characters correctly. Writing f'"{value}"' did not: a subject such
    as 'Re: the "final" draft' made a frontmatter block no YAML parser accepts.
    Non-ASCII text is kept readable rather than escaped.
    """
    text = json.dumps('' if value is None else str(value), ensure_ascii=False)
    return _YAML_UNSAFE.sub(lambda m: '\\u%04x' % ord(m.group()), text)


def format_date_human(dt: datetime) -> str:
    """Format datetime for human-readable display."""
    return dt.strftime("%B %d, %Y at %I:%M %p")


class EmailExtractor:
    """Extract emails from PST file."""

    def __init__(
        self,
        pst_path: Path,
        output_dir: Path,
        include_deleted: bool = False,
        target_timezone: Optional[str] = None,
        verbose: bool = False,
        append: bool = False,
        owner_email: str = None,
        overwrite: bool = False,
    ):
        self.pst_path = pst_path
        self.output_dir = output_dir
        self.emails_dir = output_dir / "emails"
        self.include_deleted = include_deleted
        # None keeps each message's own offset, which is what the archive has
        # always done. A zone name converts every date to that zone.
        self.target_timezone = target_timezone
        self.tz = ZoneInfo(target_timezone) if target_timezone else None
        self.verbose = verbose
        self.append = append
        self.overwrite = overwrite
        self.owner_email = owner_email
        # 'NEW', 'APPEND' or 'OVERWRITE', settled in extract().
        self.mode = 'APPEND' if append else 'NEW'

        self.stats = {'total': 0, 'processed': 0, 'errors': 0, 'attachments': 0, 'skipped': 0}
        self.index_data = []
        self.existing_message_ids = set()  # For append mode
        self.error_log = []
        self.folder_counts = {}
        self.date_range = {'min': None, 'max': None}
        # Every source this archive has ever been built or appended from, each
        # with the hash it had at the time. An archive legitimately accrues
        # several sources over its life (the original PST, then one or more
        # --append runs), and none of the earlier ones may be forgotten just
        # because a later run regenerates the manifest.
        self.provenance = []

    def log(self, message: str):
        """Log message if verbose mode is on."""
        if self.verbose:
            print(message)

    def log_error(self, message: str):
        """Log error message."""
        self.error_log.append(f"{datetime.now().isoformat()} - {message}")
        print(f"ERROR: {message}", file=sys.stderr)

    # The files and the folder this tool writes at the top of an archive. These,
    # and nothing else in the output directory, are what --overwrite removes.
    ARCHIVE_FILES = ('index.csv', 'index.md', 'manifest.sha256', 'extraction_log.txt')

    def holds_archive(self) -> bool:
        """True if the output directory already holds an archive this tool wrote."""
        if any((self.output_dir / name).exists() for name in self.ARCHIVE_FILES):
            return True
        return self.emails_dir.is_dir() and any(self.emails_dir.iterdir())

    def clear_archive(self):
        """Remove the previous archive: emails/ and the top-level files above.

        Anything else in the output directory is left alone.
        """
        source = self.pst_path.resolve()
        emails = self.emails_dir.resolve()
        if source == emails or emails in source.parents:
            print(f"Error: the source {self.pst_path} is inside {self.emails_dir}, which --overwrite would delete.")
            sys.exit(1)
        if self.emails_dir.is_dir():
            shutil.rmtree(self.emails_dir)
        for name in self.ARCHIVE_FILES:
            path = self.output_dir / name
            if path.is_file():
                path.unlink()

    def setup_directories(self):
        """Create output directory structure."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.emails_dir.mkdir(exist_ok=True)

    def _load_existing_index(self):
        """Load existing index.csv to get already-extracted message IDs."""
        index_path = self.output_dir / "index.csv"
        if not index_path.exists():
            print("No existing index.csv found - will extract all emails")
            return

        print(f"Loading existing index from {index_path}...")
        try:
            with open(index_path, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    msg_id = row.get('message_id', '').strip()
                    if msg_id:
                        self.existing_message_ids.add(msg_id)
                    # Also load existing index data so we can merge later
                    self.index_data.append(row)

            print(f"Found {len(self.existing_message_ids)} existing emails (by message ID)")

            # Rebuild folder_counts and date_range from existing data so they
            # describe the whole archive rather than just this run. Note this
            # does NOT touch self.stats: that dict is deliberately run-only
            # (it drives the "N processed / N skipped" summary for THIS
            # invocation), so anything that must describe the archive as a
            # whole - like index.md's totals - has to be derived from
            # self.index_data/folder_counts, never from self.stats.
            for row in self.index_data:
                pst_folder = row.get('pst_folder', 'Unknown')
                self.folder_counts[pst_folder] = self.folder_counts.get(pst_folder, 0) + 1

                # Track date range (make timezone-aware for consistent comparisons)
                date_str = row.get('date', '')
                if date_str:
                    try:
                        date = datetime.strptime(date_str, '%Y-%m-%d').replace(tzinfo=timezone.utc)
                        if self.date_range['min'] is None or date < self.date_range['min']:
                            self.date_range['min'] = date
                        if self.date_range['max'] is None or date > self.date_range['max']:
                            self.date_range['max'] = date
                    except ValueError:
                        pass

        except Exception as e:
            self.log_error(f"Failed to load existing index: {e}")
            self.existing_message_ids.clear()
            self.index_data.clear()

    def _load_existing_provenance(self):
        """Read the sources already on record in an existing manifest.sha256.

        Returns a list of {'source', 'sha256', 'generated'} dicts, oldest first.
        Understands two formats:

        - the current multi-source block this method itself writes, and
        - the single-source header this tool wrote before provenance tracking
          existed ("# Source: x" / "# SHA256 of source: y"), so an archive
          built by an older version of this script still appends cleanly.
        """
        manifest_path = self.output_dir / "manifest.sha256"
        if not manifest_path.exists():
            return []

        try:
            lines = manifest_path.read_text(encoding='utf-8').splitlines()
        except Exception as e:
            self.log_error(f"Failed to read existing manifest for provenance: {e}")
            return []

        entry_re = re.compile(r'^#\s+sha256=(\S+)\s+generated=(\S+)\s+source=(.*)$')
        entries = []
        for line in lines:
            match = entry_re.match(line)
            if match:
                sha256, generated, source = match.groups()
                entries.append({'source': source, 'sha256': None if sha256 == 'N/A' else sha256, 'generated': generated})
        if entries:
            return entries

        # Fall back to the legacy single-source header.
        legacy_source = None
        legacy_sha256 = None
        legacy_generated = 'unknown'
        for line in lines:
            if line.startswith('# Generated:'):
                legacy_generated = line.split(':', 1)[1].strip()
            elif line.startswith('# Source:'):
                legacy_source = line.split(':', 1)[1].strip()
            elif line.startswith('# SHA256 of source:'):
                legacy_sha256 = line.split(':', 1)[1].strip()

        if legacy_source:
            return [{'source': legacy_source, 'sha256': legacy_sha256, 'generated': legacy_generated}]
        return []

    def _build_provenance(self):
        """Assemble this run's source list: every prior source plus this one.

        Append-safe by construction: prior entries come from the manifest
        already on disk, so they survive regardless of how many times this
        has run before. Re-running the same append (e.g. an empty --since
        window against the same staging directory) must not grow the list
        forever, so an identical (source, hash) pair is not repeated.
        """
        provenance = self._load_existing_provenance() if self.append else []

        current = {
            'source': self.pst_path.name or str(self.pst_path),
            'sha256': compute_sha256(self.pst_path) if self.pst_path.is_file() else None,
            'generated': datetime.now(timezone.utc).isoformat(),
        }
        already_recorded = any(
            entry['source'] == current['source'] and entry['sha256'] == current['sha256'] for entry in provenance
        )
        if not already_recorded:
            provenance.append(current)

        self.provenance = provenance

    def extract(self):
        """Main extraction method."""
        # A run without --append used to announce "OVERWRITE (replacing existing
        # emails)" and then replace nothing: the old email folders stayed on
        # disk, missing from the new index. Now an existing archive is either
        # appended to, replaced for real with --overwrite, or refused.
        if not self.append and self.holds_archive():
            if not self.overwrite:
                print(f"Error: {self.output_dir} already holds an archive.")
                print("  To add new mail to it:  --append")
                print("  To replace it:          --overwrite (deletes its emails/ folder and index files)")
                sys.exit(1)
            self.clear_archive()
            self.mode = 'OVERWRITE'

        self.setup_directories()

        print(f"Processing: {self.pst_path}")
        print(f"Output: {self.output_dir}")
        if self.append:
            print("Mode: APPEND (skipping existing emails)")
            self._load_existing_index()
        elif self.mode == 'OVERWRITE':
            print("Mode: OVERWRITE (the previous archive was removed)")
        else:
            print("Mode: NEW")
        self._build_provenance()
        print()

        # A directory of .eml is a first-class input (a Graph export, or a PST
        # already run through readpst elsewhere). Anything else is a PST, and
        # readpst is the only PST reader.
        if self.pst_path.is_dir():
            print(f"Processing pre-extracted emails from: {self.pst_path}")
            self._process_eml_directory(self.pst_path)
        else:
            self._extract_with_readpst()

        self._generate_index_files()
        self._generate_manifest()
        self._write_extraction_log()

        self._print_summary()

    def readpst_command(self, out_dir: Path) -> list:
        """The readpst command line for this extraction.

        -j 0 turns off readpst's parallel jobs. With -e, readpst may split one
        folder's messages across jobs, and then it sometimes never writes the
        last few, with no error and a zero exit. Measured on the same PST with
        readpst 0.6.76: 70 or 71 messages by default, 71 every time with -j 0.
        An archive that silently drops mail is worse than a slower one.

        -e writes each message as its own .eml file inside a folder tree that
        mirrors the PST, which is what _process_eml_directory reads. -8 asks
        for UTF-8 bodies where the PST holds them. -D includes deleted items,
        and is passed only when --include-deleted is given.
        """
        cmd = ['readpst', '-j', '0', '-e', '-8']
        if self.include_deleted:
            cmd.append('-D')
        cmd += ['-o', str(out_dir), str(self.pst_path)]
        return cmd

    def _extract_with_readpst(self):
        """Extract a PST with readpst, then process its .eml output."""
        if shutil.which('readpst') is None:
            print("ERROR: readpst is not installed, so this PST cannot be read.")
            print()
            print("Options:")
            print("  1. Install pst-utils, which provides readpst:")
            print("     Ubuntu/Debian: sudo apt install pst-utils")
            print("     macOS: brew install libpst")
            print()
            print("  2. Run readpst on another machine, then point this tool at its output:")
            print(f"     readpst -j 0 -e -8 -o extracted_emails/ {self.pst_path.name}")
            print("     python outlook_to_md.py extracted_emails/ <output_dir>")
            sys.exit(1)

        print("Using readpst for extraction...")
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)

            print("Extracting emails from PST (this may take a while)...")
            result = subprocess.run(
                self.readpst_command(tmppath),
                capture_output=True,
                text=True,
            )

            if result.returncode != 0:
                print(f"readpst failed: {result.stderr}")
                sys.exit(1)

            self._process_eml_directory(tmppath)

    def _process_eml_directory(self, eml_dir: Path):
        """Process a directory containing .eml files."""
        # Find all extracted .eml files
        eml_files = list(eml_dir.rglob('*.eml'))
        self.stats['total'] = len(eml_files)

        print(f"Found {len(eml_files)} emails")

        for eml_path in tqdm(eml_files, desc="Processing emails"):
            try:
                self._process_eml_file(eml_path, eml_dir)
            except Exception as e:
                self.stats['errors'] += 1
                self.log_error(f"Failed to process {eml_path.name}: {e}")

    def _process_eml_file(self, eml_path: Path, base_dir: Path):
        """Process a single .eml file."""
        import email
        from email import policy

        with open(eml_path, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=policy.default)

        # Extract date
        date_str = msg.get('Date', '')
        try:
            if HAS_DATEUTIL and date_str:
                sent_date = date_parser.parse(date_str)
            elif date_str:
                # Try standard library email.utils
                from email.utils import parsedate_to_datetime

                sent_date = parsedate_to_datetime(date_str)
            else:
                sent_date = datetime.now(timezone.utc)
        except Exception:
            sent_date = datetime.now(timezone.utc)

        # Ensure timezone aware
        if sent_date.tzinfo is None:
            sent_date = sent_date.replace(tzinfo=timezone.utc)
        if self.tz is not None:
            sent_date = sent_date.astimezone(self.tz)

        # Extract basic fields
        subject = msg.get('Subject', '(No Subject)')
        sender = msg.get('From', '')
        to_str = msg.get('To', '')
        cc_str = msg.get('Cc', '')
        bcc_str = msg.get('Bcc', '')

        # Fix MAILER-DAEMON sent emails - these are actually emails sent by the user
        # readpst/libpst exports sent emails with MAILER-DAEMON as the email address
        # but includes the real sender info in X-libpst-forensic-sender header
        if 'MAILER-DAEMON' in sender and self.owner_email:
            forensic_sender = msg.get('X-libpst-forensic-sender', '')
            # Extract display name from the From field
            sender_name_match = re.match(r'^"?([^"<]+)"?\s*<', sender)
            display_name = sender_name_match.group(1).strip() if sender_name_match else ''

            # Use owner email to reconstruct the real sender
            if forensic_sender:
                if display_name:
                    sender = f'"{display_name}" <{self.owner_email}>'
                else:
                    sender = self.owner_email

        to_list = split_addresses(to_str)
        cc_list = split_addresses(cc_str)
        bcc_list = split_addresses(bcc_str)

        message_id = msg.get('Message-ID', '')

        # Get folder path from relative path
        rel_path = eml_path.relative_to(base_dir)
        folder_path = str(rel_path.parent) if rel_path.parent != Path('.') else "Root"

        # Extract body
        body_text = ""
        body_html = ""

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == 'text/plain' and not body_text:
                    try:
                        body_text = part.get_content()
                    except Exception:
                        body_text = str(part.get_payload(decode=True) or b'', errors='replace')
                elif content_type == 'text/html' and not body_html:
                    try:
                        body_html = part.get_content()
                    except Exception:
                        body_html = str(part.get_payload(decode=True) or b'', errors='replace')
        else:
            content_type = msg.get_content_type()
            try:
                content = msg.get_content()
            except Exception:
                content = str(msg.get_payload(decode=True) or b'', errors='replace')
            if content_type == 'text/html':
                body_html = content
            else:
                body_text = content

        # Convert to markdown
        if body_html:
            body_md = html_to_markdown(body_html)
        else:
            body_md = body_text

        # Get headers
        headers = "\n".join(f"{k}: {v}" for k, v in msg.items())

        # Create email data
        email_data = {
            'sent_date': sent_date,
            'subject': subject,
            'sender': sender,
            'to_list': to_list,
            'cc_list': cc_list,
            'bcc_list': bcc_list,
            'body_md': body_md,
            'body_text': body_text,
            'headers': headers,
            'message_id': message_id,
            'folder_path': folder_path,
            'attachments': [],
            'raw_eml_path': eml_path,
        }

        # Extract attachments
        if msg.is_multipart():
            att_index = 1
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    att_data = part.get_payload(decode=True)
                    if att_data:
                        filename = part.get_filename() or f"attachment_{att_index}"
                        content_type = part.get_content_type()
                        email_data['attachments'].append(
                            {
                                'original_name': filename,
                                'data': att_data,
                                'content_type': content_type,
                                'index': att_index,
                            }
                        )
                        att_index += 1

        self._save_email(email_data)

    def _save_email(self, email_data: dict):
        """Save email to output directory."""
        # In append mode, skip emails we've already extracted
        message_id = email_data.get('message_id', '').strip()
        if self.append and message_id and message_id in self.existing_message_ids:
            self.stats['skipped'] += 1
            self.log(f"Skipping existing email: {email_data.get('subject', '(No Subject)')[:50]}")
            return

        sent_date = email_data['sent_date']
        subject = email_data['subject']
        sender = email_data['sender']
        to_list = email_data['to_list']

        # Parse sender
        sender_name, sender_email = parse_email_address(sender)

        # Parse primary recipient
        if to_list:
            to_name, to_email = parse_email_address(to_list[0])
        else:
            to_name, to_email = "", "unknown"

        # Create folder name - prefer name over email when available
        date_str = sent_date.strftime("%Y-%m-%d_%H%M%S")
        sender_sanitized = (
            sanitize_filename(sender_name, max_length=40) if sender_name else sanitize_email(sender_email)
        )
        to_sanitized = sanitize_filename(to_name, max_length=40) if to_name else sanitize_email(to_email)
        subject_sanitized = sanitize_filename(subject, max_length=50)

        folder_name = f"{date_str}_from-{sender_sanitized}_to-{to_sanitized}_{subject_sanitized}"

        # Ensure path isn't too long (Windows safety)
        if len(folder_name) > 150:
            folder_name = folder_name[:150]

        # Build the full path including PST folder structure
        pst_folder = email_data.get('folder_path', 'Unknown')
        if pst_folder and pst_folder not in ('Unknown', 'Root', '.'):
            # Sanitize each path component
            folder_parts = [sanitize_filename(part, max_length=50) for part in pst_folder.split('/') if part]
            subfolder_path = Path(*folder_parts) if folder_parts else Path()
            base_dir = self.emails_dir / subfolder_path
        else:
            base_dir = self.emails_dir

        # Handle duplicates
        email_folder = base_dir / folder_name
        counter = 1
        original_folder_name = folder_name
        while email_folder.exists():
            folder_name = f"{original_folder_name}-{counter:03d}"
            email_folder = base_dir / folder_name
            counter += 1

        email_folder.mkdir(parents=True, exist_ok=True)

        # Calculate relative path from emails_dir for index
        relative_folder_path = str(email_folder.relative_to(self.emails_dir))

        # Save attachments
        saved_attachments = []
        for att in email_data.get('attachments', []):
            att_info = self._write_attachment(email_folder, att)
            if att_info:
                saved_attachments.append(att_info)
                self.stats['attachments'] += 1

        # Save raw .eml file
        eml_path = email_folder / "email.eml"
        if 'raw_eml_path' in email_data:
            # Copy the original eml file
            shutil.copy2(email_data['raw_eml_path'], eml_path)
        else:
            # Generate .eml from message data
            self._generate_eml(email_folder, email_data)

        # Generate email.md
        self._generate_email_md(email_folder, email_data, saved_attachments)

        # Generate checksums
        self._generate_checksums(email_folder)

        # Update stats and index
        self.stats['processed'] += 1

        # Track date range (ensure timezone-aware for consistent comparisons)
        if sent_date.tzinfo is None:
            sent_date = sent_date.replace(tzinfo=timezone.utc)
            email_data['sent_date'] = sent_date  # Update for consistency
        if self.date_range['min'] is None or sent_date < self.date_range['min']:
            self.date_range['min'] = sent_date
        if self.date_range['max'] is None or sent_date > self.date_range['max']:
            self.date_range['max'] = sent_date

        # Track folder counts
        pst_folder = email_data.get('folder_path', 'Unknown')
        self.folder_counts[pst_folder] = self.folder_counts.get(pst_folder, 0) + 1

        # Add to index
        sender_name, sender_email = parse_email_address(email_data['sender'])
        to_name, to_email = parse_email_address(to_list[0]) if to_list else ("", "")

        self.index_data.append(
            {
                'folder_name': relative_folder_path,
                'date': sent_date.strftime("%Y-%m-%d"),
                'time': sent_date.strftime("%H:%M:%S"),
                'from_email': sender_email,
                'from_name': sender_name,
                'to_email': to_email,
                'to_name': to_name,
                'cc': ', '.join(email_data.get('cc_list', [])),
                'subject': subject,
                'attachment_count': len(saved_attachments),
                'has_body': bool(email_data.get('body_md', '').strip()),
                'pst_folder': pst_folder,
                'message_id': email_data.get('message_id', ''),
            }
        )

    def _write_attachment(self, email_folder: Path, att: dict) -> Optional[dict]:
        """Write attachment to disk."""
        try:
            original_name = att.get('original_name', 'attachment')
            index = att.get('index', 1)

            # Get file extension
            ext = Path(original_name).suffix or ''
            base_name = Path(original_name).stem

            # Sanitize and create filename
            sanitized_base = sanitize_filename(base_name, max_length=40)
            filename = f"attachment_{index:03d}_{sanitized_base}{ext}"

            filepath = email_folder / filename

            # Write data
            data = att.get('data', b'')
            with open(filepath, 'wb') as f:
                f.write(data)

            return {
                'filename': filename,
                'original_name': original_name,
                'size_bytes': len(data),
                'content_type': att.get('content_type', 'application/octet-stream'),
                'sha256': compute_sha256(filepath),
            }
        except Exception as e:
            self.log_error(f"Failed to save attachment {att.get('original_name', 'unknown')}: {e}")
            return None

    def _generate_eml(self, email_folder: Path, email_data: dict):
        """Generate .eml file from email data."""
        eml_path = email_folder / "email.eml"

        # Build a basic RFC 822 message
        lines = []

        if email_data.get('headers'):
            lines.append(email_data['headers'])
        else:
            lines.append(f"Date: {email_data['sent_date'].strftime('%a, %d %b %Y %H:%M:%S %z')}")
            lines.append(f"From: {email_data['sender']}")
            if email_data['to_list']:
                lines.append(f"To: {', '.join(email_data['to_list'])}")
            if email_data.get('cc_list'):
                lines.append(f"Cc: {', '.join(email_data['cc_list'])}")
            lines.append(f"Subject: {email_data['subject']}")
            if email_data.get('message_id'):
                lines.append(f"Message-ID: {email_data['message_id']}")
            lines.append("MIME-Version: 1.0")
            lines.append("Content-Type: text/plain; charset=UTF-8")

        lines.append("")
        lines.append(email_data.get('body_text', '') or email_data.get('body_md', ''))

        with open(eml_path, 'w', encoding='utf-8', errors='replace') as f:
            f.write('\n'.join(lines))

    def _generate_email_md(self, email_folder: Path, email_data: dict, attachments: list):
        """Generate email.md file."""
        sent_date = email_data['sent_date']
        subject = email_data['subject']

        sender_name, sender_email = parse_email_address(email_data['sender'])
        sender_display = f"{sender_name} <{sender_email}>" if sender_name else sender_email

        to_display = ', '.join(email_data.get('to_list', []))
        cc_display = ', '.join(email_data.get('cc_list', []))

        # Build YAML frontmatter. Every string goes through yaml_str, so a quote,
        # backslash or colon in a subject or a name cannot break the block.
        yaml_lines = [
            "---",
            f'message_id: {yaml_str(email_data.get("message_id", ""))}',
            f'date: {yaml_str(sent_date.isoformat())}',
            f'from: {yaml_str(email_data["sender"])}',
        ]
        for key in ('to', 'cc', 'bcc'):
            addresses = email_data.get(f'{key}_list', [])
            if addresses:
                yaml_lines.append(f"{key}:")
                yaml_lines.extend(f'  - {yaml_str(addr)}' for addr in addresses)
            else:
                yaml_lines.append(f"{key}: []")

        yaml_lines.append(f'subject: {yaml_str(subject)}')
        yaml_lines.append(f'has_attachments: {str(bool(attachments)).lower()}')
        yaml_lines.append(f'attachment_count: {len(attachments)}')

        if attachments:
            yaml_lines.append("attachments:")
            for att in attachments:
                yaml_lines.append(f'  - filename: {yaml_str(att["filename"])}')
                yaml_lines.append(f'    original_name: {yaml_str(att["original_name"])}')
                yaml_lines.append(f'    size_bytes: {att["size_bytes"]}')
                yaml_lines.append(f'    content_type: {yaml_str(att["content_type"])}')
                yaml_lines.append(f'    sha256: {yaml_str(att["sha256"])}')

        yaml_lines.append(f'pst_folder: {yaml_str(email_data.get("folder_path", "Unknown"))}')
        yaml_lines.append(f'extraction_date: {yaml_str(datetime.now(timezone.utc).isoformat())}')
        yaml_lines.append(f'source_file: {yaml_str(self.pst_path.name)}')
        yaml_lines.append("---")

        # Build markdown content
        md_lines = [
            "",
            f"# {subject}",
            "",
            f"**From:** {sender_display}",
            f"**To:** {to_display}",
        ]

        if cc_display:
            md_lines.append(f"**CC:** {cc_display}")

        md_lines.append(f"**Date:** {format_date_human(sent_date)}")
        md_lines.append(f"**Subject:** {subject}")
        md_lines.append("")
        md_lines.append("---")
        md_lines.append("")
        md_lines.append("## Body")
        md_lines.append("")
        md_lines.append(email_data.get('body_md', '') or "(No body content)")
        md_lines.append("")

        if attachments:
            md_lines.append("---")
            md_lines.append("")
            md_lines.append("## Attachments")
            md_lines.append("")
            for i, att in enumerate(attachments, 1):
                size_str = format_size(att['size_bytes'])
                md_lines.append(f"{i}. [{att['original_name']}](./{att['filename']}) ({size_str})")
            md_lines.append("")

        md_lines.append("---")
        md_lines.append("")
        md_lines.append("## Original Headers")
        md_lines.append("")
        md_lines.append("```")
        md_lines.append(email_data.get('headers', '(Headers not available)'))
        md_lines.append("```")

        # Write file
        md_path = email_folder / "email.md"
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(yaml_lines))
            f.write('\n'.join(md_lines))

    def _generate_checksums(self, email_folder: Path):
        """Generate checksums.sha256 for all files in folder."""
        checksums_path = email_folder / "checksums.sha256"
        lines = []

        for filepath in sorted(email_folder.iterdir()):
            if filepath.name != "checksums.sha256" and filepath.is_file():
                sha256 = compute_sha256(filepath)
                lines.append(f"{sha256}  {filepath.name}")

        with open(checksums_path, 'w') as f:
            f.write('\n'.join(lines) + '\n')

    def _generate_index_files(self):
        """Generate index.csv and index.md."""
        # Sort by date
        self.index_data.sort(key=lambda x: (x['date'], x['time']))

        # Generate CSV
        csv_path = self.output_dir / "index.csv"
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    'folder_name',
                    'date',
                    'time',
                    'from_email',
                    'from_name',
                    'to_email',
                    'to_name',
                    'cc',
                    'subject',
                    'attachment_count',
                    'has_body',
                    'pst_folder',
                    'message_id',
                ],
            )
            writer.writeheader()
            writer.writerows(self.index_data)

        # Generate MD
        md_path = self.output_dir / "index.md"

        date_range_str = "N/A"
        if self.date_range['min'] and self.date_range['max']:
            date_range_str = (
                f"{self.date_range['min'].strftime('%Y-%m-%d')} to {self.date_range['max'].strftime('%Y-%m-%d')}"
            )

        # Cumulative across the whole archive, not just this run: self.stats
        # only counts what THIS invocation processed, so on an --append run it
        # undercounts everything that was already there. self.index_data is
        # the merged (existing + new) row set, so it - not self.stats - is
        # the archive's true total. See _load_existing_index for how the
        # existing rows get merged in.
        total_emails = len(self.index_data)
        total_attachments = sum(int(row.get('attachment_count') or 0) for row in self.index_data)

        sources_str = ', '.join(
            f"{entry['source']} (sha256 {entry['sha256'][:12]}…)" if entry['sha256'] else f"{entry['source']} (N/A)"
            for entry in self.provenance
        )

        lines = [
            "# Email Archive Index",
            "",
            f"**Sources:** {sources_str}",
            f"**Extracted:** {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')} UTC",
            f"**Total Emails:** {total_emails:,}",
            f"**Total Attachments:** {total_attachments:,}",
            f"**Date Range:** {date_range_str}",
            "",
            "## Statistics",
            "",
            "| Folder | Count |",
            "|--------|-------|",
        ]

        for folder, count in sorted(self.folder_counts.items()):
            lines.append(f"| {folder} | {count} |")

        lines.append("")
        lines.append("## Timeline")
        lines.append("")

        # Group by year/month
        current_year = None
        current_month = None

        for item in self.index_data:
            date = datetime.strptime(item['date'], '%Y-%m-%d')
            year = date.year
            month = date.strftime('%B %Y')

            if year != current_year:
                lines.append(f"### {year}")
                lines.append("")
                current_year = year
                current_month = None

            if month != current_month:
                lines.append(f"#### {month}")
                lines.append("")
                current_month = month

            att_str = f" ({item['attachment_count']} attachments)" if item['attachment_count'] else ""
            from_display = item['from_name'] or item['from_email']
            to_display = item['to_name'] or item['to_email']

            lines.append(
                f"- [{item['date']} {item['time']}](./emails/{item['folder_name']}/) - "
                f"**{from_display}** \u2192 {to_display} - \"{item['subject']}\"{att_str}"
            )

        with open(md_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')

    def _generate_manifest(self):
        """Generate master manifest.sha256.

        Every run regenerates this file, but the source list is cumulative
        (see _build_provenance): appending to an archive must not erase the
        identity and hash of whatever it was originally built from.
        """
        manifest_path = self.output_dir / "manifest.sha256"
        lines = [
            f"# Generated: {datetime.now(timezone.utc).isoformat()}",
            "# Archive sources:",
        ]
        for entry in self.provenance:
            sha_display = entry['sha256'] or 'N/A'
            lines.append(f"#   sha256={sha_display} generated={entry['generated']} source={entry['source']}")

        lines.append("")

        # Hash all checksums files
        for checksums_file in sorted(self.emails_dir.rglob("checksums.sha256")):
            rel_path = checksums_file.relative_to(self.output_dir)
            sha256 = compute_sha256(checksums_file)
            lines.append(f"{sha256}  {rel_path}")

        # Hash index files
        for index_file in ['index.csv', 'index.md']:
            filepath = self.output_dir / index_file
            if filepath.exists():
                sha256 = compute_sha256(filepath)
                lines.append(f"{sha256}  {index_file}")

        with open(manifest_path, 'w') as f:
            f.write('\n'.join(lines) + '\n')

    def _write_extraction_log(self):
        """Write extraction log."""
        log_path = self.output_dir / "extraction_log.txt"

        lines = [
            "=" * 60,
            "PST Email Extraction Log",
            "=" * 60,
            "",
            f"Source: {self.pst_path}",
            f"Output: {self.output_dir}",
            f"Mode: {self.mode}",
            f"Started: {datetime.now().isoformat()}",
            "",
            "Statistics:",
            f"  Total messages found: {self.stats['total']}",
            f"  Successfully processed: {self.stats['processed']}",
            f"  Skipped (already exist): {self.stats['skipped']}",
            f"  Errors: {self.stats['errors']}",
            f"  Attachments extracted: {self.stats['attachments']}",
            "",
        ]

        if self.error_log:
            lines.append("Errors:")
            lines.extend(f"  {error}" for error in self.error_log)
            lines.append("")

        lines.append("=" * 60)

        with open(log_path, 'w') as f:
            f.write('\n'.join(lines) + '\n')

    def _print_summary(self):
        """Print extraction summary."""
        print()
        print("=" * 60)
        print("Extraction Complete")
        print("=" * 60)
        print(f"  Emails processed: {self.stats['processed']:,}")
        if self.append:
            print(f"  Emails skipped (already exist): {self.stats['skipped']:,}")
        print(f"  Attachments: {self.stats['attachments']:,}")
        print(f"  Errors: {self.stats['errors']:,}")
        print()
        print(f"Output: {self.output_dir}")
        print("=" * 60)


def timezone_name(value: str) -> str:
    """argparse type for --timezone: an IANA zone name this machine knows."""
    try:
        ZoneInfo(value)
    except (ZoneInfoNotFoundError, ValueError):
        raise argparse.ArgumentTypeError(f"unknown timezone {value!r}; use an IANA name such as Europe/London")
    return value


def main():
    parser = argparse.ArgumentParser(
        description="Extract emails from Outlook PST files into organized markdown archive"
    )
    parser.add_argument("pst_file", help="Path to PST file (or directory of .eml files)")
    parser.add_argument("output_dir", help="Output directory")
    parser.add_argument(
        "--include-deleted", action="store_true", help="Include deleted items (passes -D to readpst; PST input only)"
    )
    parser.add_argument(
        "--timezone",
        type=timezone_name,
        default=None,
        metavar="TZ",
        help="Render every date in this IANA zone, e.g. Europe/London (default: the offset each message was sent with)",
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--append", action="store_true", help="Append mode: skip emails already in the archive (by message ID)"
    )
    mode.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace an existing archive: delete its emails/ folder and index files first",
    )
    parser.add_argument("--owner-email", help="PST owner's email address (used to fix MAILER-DAEMON sent items)")

    args = parser.parse_args()

    pst_path = Path(args.pst_file)
    output_dir = Path(args.output_dir)

    if not pst_path.exists():
        print(f"Error: PST file not found: {pst_path}")
        sys.exit(1)

    extractor = EmailExtractor(
        pst_path=pst_path,
        output_dir=output_dir,
        include_deleted=args.include_deleted,
        target_timezone=args.timezone,
        verbose=args.verbose,
        append=args.append,
        owner_email=args.owner_email,
        overwrite=args.overwrite,
    )

    extractor.extract()


if __name__ == "__main__":
    main()
