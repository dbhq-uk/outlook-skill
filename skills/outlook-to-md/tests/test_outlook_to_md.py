#!/usr/bin/env python3
"""Tests for outlook_to_md.py.

This covers the pure helpers, the append-mode index loading, and the PST path
up to the readpst call: the command line it builds, and what comes out when
readpst leaves its usual tree of .eml files. readpst itself is stubbed, so the
suite needs no PST and no pst-utils. CI converts a real PST end to end in a
separate job.
"""

from __future__ import annotations

import csv
import hashlib
import io
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

import outlook_to_md  # noqa: E402

INDEX_COLUMNS = [
    "folder_name",
    "date",
    "time",
    "from_email",
    "from_name",
    "to_email",
    "to_name",
    "cc",
    "subject",
    "attachment_count",
    "has_body",
    "pst_folder",
    "message_id",
]


def index_row(message_id: str, **overrides) -> dict:
    row = {
        "folder_name": "Inbox/2026",
        "date": "2026-07-01",
        "time": "09:30:00",
        "from_email": "sender@example.com",
        "from_name": "A Sender",
        "to_email": "me@example.com",
        "to_name": "Me",
        "cc": "",
        "subject": "A subject",
        "attachment_count": "0",
        "has_body": "True",
        "pst_folder": "Inbox",
        "message_id": message_id,
    }
    row.update(overrides)
    return row


class TestSanitizeFilename(unittest.TestCase):
    def test_spaces_become_hyphens(self):
        self.assertEqual(outlook_to_md.sanitize_filename("quarterly report"), "quarterly-report")

    def test_path_separators_are_stripped(self):
        result = outlook_to_md.sanitize_filename("reports/2026/q3")
        self.assertNotIn("/", result)
        self.assertNotIn("\\", result)

    def test_windows_reserved_characters_are_stripped(self):
        result = outlook_to_md.sanitize_filename('a<b>c:d"e|f?g*h')
        for char in '<>:"|?*':
            self.assertNotIn(char, result)

    def test_repeated_hyphens_collapse(self):
        self.assertEqual(outlook_to_md.sanitize_filename("a   -  b"), "a-b")

    def test_leading_and_trailing_hyphens_are_dropped(self):
        self.assertEqual(outlook_to_md.sanitize_filename("  spaced  "), "spaced")

    def test_truncates_to_max_length(self):
        result = outlook_to_md.sanitize_filename("x" * 200)
        self.assertEqual(len(result), 50)

    def test_max_length_is_configurable(self):
        result = outlook_to_md.sanitize_filename("y" * 200, max_length=10)
        self.assertEqual(len(result), 10)

    def test_truncation_does_not_leave_a_trailing_hyphen(self):
        # "aaaa bbbb ..." truncated mid-gap would otherwise end in "-".
        result = outlook_to_md.sanitize_filename("a" * 49 + " tail", max_length=50)
        self.assertFalse(result.endswith("-"))

    def test_empty_input_returns_unknown(self):
        self.assertEqual(outlook_to_md.sanitize_filename(""), "unknown")

    def test_input_stripped_to_nothing_returns_unknown(self):
        self.assertEqual(outlook_to_md.sanitize_filename("///"), "unknown")

    def test_unicode_is_preserved(self):
        self.assertIn("café", outlook_to_md.sanitize_filename("café meeting"))


class TestSanitizeEmail(unittest.TestCase):
    def test_extracts_address_from_display_name_form(self):
        result = outlook_to_md.sanitize_email("A Sender <sender@example.com>")
        self.assertIn("sender", result)
        self.assertIn("example.com", result)
        self.assertNotIn("<", result)

    def test_at_sign_is_removed_but_dots_survive(self):
        result = outlook_to_md.sanitize_email("first.last@example.com")
        self.assertNotIn("@", result)
        self.assertIn(".", result)

    def test_empty_input_returns_unknown(self):
        self.assertEqual(outlook_to_md.sanitize_email(""), "unknown")

    def test_truncates_at_forty_characters(self):
        result = outlook_to_md.sanitize_email("x" * 100 + "@example.com")
        self.assertLessEqual(len(result), 40)


class TestParseEmailAddress(unittest.TestCase):
    def test_display_name_and_address(self):
        name, email = outlook_to_md.parse_email_address("A Sender <sender@example.com>")
        self.assertEqual(name, "A Sender")
        self.assertEqual(email, "sender@example.com")

    def test_quoted_display_name(self):
        name, email = outlook_to_md.parse_email_address('"Last, First" <a@b.com>')
        self.assertIn("Last", name)
        self.assertEqual(email, "a@b.com")

    def test_bare_address_yields_the_address(self):
        # Known quirk, asserted rather than wished away: the regex's greedy first
        # group also claims a bare address as the display name, so this returns
        # ("bare@example.com", "bare@example.com"). Harmless -- the index's
        # from_name column just repeats the address -- but it is the current
        # behaviour, and this test will fail loudly if anyone changes it.
        name, email = outlook_to_md.parse_email_address("bare@example.com")
        self.assertEqual(email, "bare@example.com")
        self.assertEqual(name, "bare@example.com")

    def test_empty_input_returns_empty_pair(self):
        self.assertEqual(outlook_to_md.parse_email_address(""), ("", ""))

    def test_surrounding_whitespace_is_trimmed(self):
        name, email = outlook_to_md.parse_email_address("  A Sender <s@e.com>  ")
        self.assertEqual(name, "A Sender")
        self.assertEqual(email, "s@e.com")


class TestComputeSha256(unittest.TestCase):
    def test_matches_hashlib(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "file.bin"
            payload = b"some attachment bytes"
            path.write_bytes(payload)
            self.assertEqual(
                outlook_to_md.compute_sha256(path),
                hashlib.sha256(payload).hexdigest(),
            )

    def test_empty_file_hashes_to_the_known_empty_digest(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "empty"
            path.write_bytes(b"")
            self.assertEqual(
                outlook_to_md.compute_sha256(path),
                "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
            )

    def test_reads_files_larger_than_the_chunk_size(self):
        # The implementation reads in 8 KiB chunks; this crosses several.
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "big.bin"
            payload = b"x" * (8192 * 3 + 17)
            path.write_bytes(payload)
            self.assertEqual(
                outlook_to_md.compute_sha256(path),
                hashlib.sha256(payload).hexdigest(),
            )


class TestFormatSize(unittest.TestCase):
    def test_bytes_render_without_a_decimal(self):
        self.assertEqual(outlook_to_md.format_size(512), "512 B")

    def test_kilobytes(self):
        self.assertEqual(outlook_to_md.format_size(1024), "1.0 KB")

    def test_megabytes(self):
        self.assertEqual(outlook_to_md.format_size(1024 * 1024), "1.0 MB")

    def test_gigabytes(self):
        self.assertEqual(outlook_to_md.format_size(1024**3), "1.0 GB")

    def test_terabytes_are_the_ceiling(self):
        self.assertTrue(outlook_to_md.format_size(1024**4).endswith("TB"))

    def test_zero(self):
        self.assertEqual(outlook_to_md.format_size(0), "0 B")


@unittest.skipUnless(outlook_to_md.HAS_HTML2TEXT, "html2text not installed")
class TestHtmlToMarkdown(unittest.TestCase):
    """The html2text path. CI installs requirements.txt, so this always runs there."""

    def test_empty_input_returns_empty(self):
        self.assertEqual(outlook_to_md.html_to_markdown(""), "")

    def test_link_urls_are_preserved(self):
        # ignore_links = False is the setting under test: losing it would
        # silently drop every URL out of an archived mailbox.
        result = outlook_to_md.html_to_markdown('<p>See <a href="https://example.com">this</a>.</p>')
        self.assertIn("example.com", result)
        self.assertIn("this", result)

    def test_long_lines_are_not_wrapped(self):
        # body_width = 0. Wrapping would corrupt quoted text and code blocks.
        sentence = " ".join(["word"] * 60)
        result = outlook_to_md.html_to_markdown(f"<p>{sentence}</p>")
        self.assertIn(sentence, result.replace("\n", " ").strip())

    def test_strips_tags(self):
        result = outlook_to_md.html_to_markdown("<p>Hello <b>world</b></p>")
        self.assertIn("Hello", result)
        self.assertIn("world", result)
        self.assertNotIn("<b>", result)


class TestHtmlToMarkdownFallback(unittest.TestCase):
    """The HAS_HTML2TEXT = False path, which runs whenever the optional dep is absent."""

    def setUp(self):
        self.patcher = patch.object(outlook_to_md, "HAS_HTML2TEXT", False)
        self.patcher.start()

    def tearDown(self):
        self.patcher.stop()

    def test_br_becomes_a_newline(self):
        self.assertIn("\n", outlook_to_md.html_to_markdown("one<br>two"))

    def test_tags_are_stripped(self):
        result = outlook_to_md.html_to_markdown("<p>Hello <b>world</b></p>")
        self.assertNotIn("<", result)
        self.assertIn("Hello", result)

    def test_entities_are_decoded(self):
        result = outlook_to_md.html_to_markdown("a &amp; b &lt;c&gt; &quot;d&quot;&nbsp;e")
        self.assertIn("&", result)
        self.assertIn("<c>", result)
        self.assertIn('"d"', result)

    def test_empty_input_still_returns_empty(self):
        self.assertEqual(outlook_to_md.html_to_markdown(""), "")


class TestAppendModeIndexLoading(unittest.TestCase):
    """--append dedupes against index.csv; getting this wrong duplicates an archive."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.output_dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def write_index(self, rows: list[dict]):
        path = self.output_dir / "index.csv"
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=INDEX_COLUMNS)
            writer.writeheader()
            for row in rows:
                writer.writerow(row)
        return path

    def extractor(self):
        return outlook_to_md.EmailExtractor(
            pst_path=self.output_dir / "archive.pst",
            output_dir=self.output_dir,
            append=True,
        )

    def test_no_index_leaves_the_id_set_empty(self):
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(ex.existing_message_ids, set())
        self.assertEqual(ex.index_data, [])

    def test_message_ids_are_loaded(self):
        self.write_index(
            [
                index_row("<a@example.com>"),
                index_row("<b@example.com>"),
            ]
        )
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(
            ex.existing_message_ids,
            {"<a@example.com>", "<b@example.com>"},
        )

    def test_existing_rows_are_kept_for_the_merged_index(self):
        self.write_index([index_row("<a@example.com>"), index_row("<b@example.com>")])
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(len(ex.index_data), 2)

    def test_blank_message_ids_are_not_added(self):
        # A row with no Message-ID must not make the empty string a "seen" id,
        # or every future header-less email would be skipped as a duplicate.
        self.write_index([index_row(""), index_row("<real@example.com>")])
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(ex.existing_message_ids, {"<real@example.com>"})

    def test_surrounding_whitespace_on_ids_is_trimmed(self):
        self.write_index([index_row("  <spaced@example.com>  ")])
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(ex.existing_message_ids, {"<spaced@example.com>"})

    def test_folder_counts_are_rebuilt_from_the_index(self):
        self.write_index(
            [
                index_row("<a@x>", pst_folder="Inbox"),
                index_row("<b@x>", pst_folder="Inbox"),
                index_row("<c@x>", pst_folder="Sent"),
            ]
        )
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(ex.folder_counts, {"Inbox": 2, "Sent": 1})

    def test_date_range_is_rebuilt_from_the_index(self):
        self.write_index(
            [
                index_row("<a@x>", date="2026-03-01"),
                index_row("<b@x>", date="2026-07-15"),
                index_row("<c@x>", date="2026-05-02"),
            ]
        )
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(ex.date_range["min"].strftime("%Y-%m-%d"), "2026-03-01")
        self.assertEqual(ex.date_range["max"].strftime("%Y-%m-%d"), "2026-07-15")

    def test_unparseable_dates_are_ignored_rather_than_fatal(self):
        self.write_index(
            [
                index_row("<a@x>", date="not-a-date"),
                index_row("<b@x>", date="2026-07-15"),
            ]
        )
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(ex.date_range["max"].strftime("%Y-%m-%d"), "2026-07-15")

    def test_an_unreadable_index_clears_state_instead_of_half_loading(self):
        # Half-loaded state is the dangerous outcome: it would look like a
        # successful dedupe set while silently missing entries.
        path = self.output_dir / "index.csv"
        path.write_bytes(b"\xff\xfe\x00 not valid utf-8 csv")
        ex = self.extractor()
        ex._load_existing_index()
        self.assertEqual(ex.existing_message_ids, set())
        self.assertEqual(ex.index_data, [])


class TestDirectoryDispatch(unittest.TestCase):
    """A directory input goes straight to directory mode and never runs readpst."""

    def test_directory_input_skips_readpst(self):
        with tempfile.TemporaryDirectory() as tmp:
            staging = Path(tmp) / "staging"
            staging.mkdir()
            (staging / "a.eml").write_text("Subject: x\n\nbody\n")
            out = Path(tmp) / "out"

            extractor = outlook_to_md.EmailExtractor(pst_path=staging, output_dir=out)

            called = {"dir": False, "readpst": False}

            def fake_dir(path):
                called["dir"] = True

            def fake_readpst():
                called["readpst"] = True

            with patch.object(extractor, "_process_eml_directory", fake_dir), patch.object(
                extractor, "_extract_with_readpst", fake_readpst
            ):
                extractor.extract()

            self.assertTrue(called["dir"], "directory input did not reach directory mode")
            self.assertFalse(called["readpst"], "directory input was sent to readpst")


# What readpst -e leaves behind: one .eml per message, in a folder tree that
# mirrors the PST, numbered from 1 within each folder.
READPST_TREE = {
    "Personal folders/Inbox/1.eml": (
        "Message-ID: <in-1@example.com>\n"
        "Date: Tue, 01 Sep 2026 09:00:00 +0000\n"
        "From: Alice Example <alice@example.com>\n"
        "To: Bob Example <bob@example.com>\n"
        "Subject: Plain inbox message\n"
        "Content-Type: text/plain; charset=utf-8\n\n"
        "Hello Bob.\n"
    ),
    "Personal folders/Inbox/2.eml": (
        "Message-ID: <in-2@example.com>\n"
        "Date: Wed, 02 Sep 2026 10:00:00 +0000\n"
        "From: Carol Example <carol@example.com>\n"
        "To: Bob Example <bob@example.com>\n"
        "Subject: With an attachment\n"
        "MIME-Version: 1.0\n"
        'Content-Type: multipart/mixed; boundary="XYZ"\n\n'
        "--XYZ\n"
        "Content-Type: text/plain; charset=utf-8\n\n"
        "See attached.\n"
        "--XYZ\n"
        'Content-Type: text/plain; name="notes.txt"\n'
        'Content-Disposition: attachment; filename="notes.txt"\n'
        "Content-Transfer-Encoding: base64\n\n"
        "YXR0YWNobWVudCBib2R5\n"
        "--XYZ--\n"
    ),
    "Personal folders/Sent Items/1.eml": (
        "Message-ID: <sent-1@example.com>\n"
        "Date: Thu, 03 Sep 2026 11:00:00 +0000\n"
        "From: Bob Example <bob@example.com>\n"
        "To: Alice Example <alice@example.com>\n"
        "Subject: Reply from sent items\n"
        "Content-Type: text/plain; charset=utf-8\n\n"
        "Thanks.\n"
    ),
}


class TestReadpstBackend(unittest.TestCase):
    """readpst is the only PST reader. Pin its command line and its hand-off."""

    def extractor(self, tmp, **kwargs):
        pst = Path(tmp) / "mailbox.pst"
        pst.write_bytes(b"stand-in PST bytes")
        return outlook_to_md.EmailExtractor(pst_path=pst, output_dir=Path(tmp) / "out", **kwargs), pst

    def test_command_without_include_deleted(self):
        with tempfile.TemporaryDirectory() as tmp:
            ex, pst = self.extractor(tmp)
            self.assertEqual(
                ex.readpst_command(Path("/staging")),
                ["readpst", "-e", "-8", "-o", "/staging", str(pst)],
            )

    def test_command_with_include_deleted_adds_D(self):
        with tempfile.TemporaryDirectory() as tmp:
            ex, pst = self.extractor(tmp, include_deleted=True)
            self.assertEqual(
                ex.readpst_command(Path("/staging")),
                ["readpst", "-e", "-8", "-D", "-o", "/staging", str(pst)],
            )

    def run_pst(self, tmp, **kwargs):
        """Run extract() on a .pst with readpst stubbed to write READPST_TREE."""
        ex, _ = self.extractor(tmp, **kwargs)
        calls = []

        def fake_run(cmd, **_):
            calls.append(list(cmd))
            out_dir = Path(cmd[cmd.index("-o") + 1])
            for rel, text in READPST_TREE.items():
                path = out_dir / rel
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text(text)

            class Result:
                returncode = 0
                stdout = ""
                stderr = ""

            return Result()

        with patch.object(outlook_to_md.shutil, "which", lambda name: "/usr/bin/readpst"), patch.object(
            outlook_to_md.subprocess, "run", fake_run
        ), redirect_stdout(io.StringIO()):
            ex.extract()
        return ex, calls

    def test_pst_goes_through_readpst_and_comes_out_whole(self):
        with tempfile.TemporaryDirectory() as tmp:
            ex, calls = self.run_pst(tmp)
            self.assertEqual(len(calls), 1, "readpst was not run exactly once")
            self.assertEqual(calls[0][:3], ["readpst", "-e", "-8"])
            self.assertNotIn("-D", calls[0])

            out = Path(tmp) / "out"
            with open(out / "index.csv", newline="", encoding="utf-8") as f:
                rows = list(csv.DictReader(f))
            self.assertEqual(ex.stats["errors"], 0)
            self.assertEqual(
                sorted(r["message_id"] for r in rows),
                ["<in-1@example.com>", "<in-2@example.com>", "<sent-1@example.com>"],
            )
            folders = {r["message_id"]: r["pst_folder"] for r in rows}
            self.assertEqual(folders["<in-1@example.com>"], "Personal folders/Inbox")
            self.assertEqual(folders["<sent-1@example.com>"], "Personal folders/Sent Items")

            with_att = next(r for r in rows if r["message_id"] == "<in-2@example.com>")
            self.assertEqual(with_att["attachment_count"], "1")
            saved = list((out / "emails").rglob("attachment_001_notes.txt"))
            self.assertEqual(len(saved), 1, "the attachment was not written")
            self.assertEqual(saved[0].read_bytes(), b"attachment body")

    def test_include_deleted_reaches_readpst(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, calls = self.run_pst(tmp, include_deleted=True)
            self.assertIn("-D", calls[0])

    def test_missing_readpst_stops_with_install_advice(self):
        with tempfile.TemporaryDirectory() as tmp:
            ex, _ = self.extractor(tmp)
            buf = io.StringIO()
            with patch.object(outlook_to_md.shutil, "which", lambda name: None), redirect_stdout(buf):
                with self.assertRaises(SystemExit):
                    ex.extract()
            self.assertIn("pst-utils", buf.getvalue())

    def test_readpst_failure_stops(self):
        with tempfile.TemporaryDirectory() as tmp:
            ex, _ = self.extractor(tmp)

            class Failed:
                returncode = 1
                stdout = ""
                stderr = "Error opening File"

            with patch.object(outlook_to_md.shutil, "which", lambda name: "/usr/bin/readpst"), patch.object(
                outlook_to_md.subprocess, "run", lambda cmd, **_: Failed()
            ), redirect_stdout(io.StringIO()):
                with self.assertRaises(SystemExit):
                    ex.extract()


class TestTimezoneFlag(unittest.TestCase):
    """--timezone converts every rendered date, and an unknown zone is refused.

    It was parsed and stored but never read, so every date came out in the
    offset the sender used whatever the flag said. These run the real command
    line, so a flag that is accepted and ignored fails here.
    """

    EML = (
        "Message-ID: <tz@example.com>\n"
        "Date: Mon, 01 Jul 2024 09:00:00 -0400\n"
        "From: Alice <alice@example.com>\n"
        "To: Bob <bob@example.com>\n"
        "Subject: Timezones\n"
        "Content-Type: text/plain; charset=utf-8\n"
        "\n"
        "Body.\n"
    )

    def run_cli(self, *extra):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        staging = Path(tmp.name) / "staging" / "Inbox"
        staging.mkdir(parents=True)
        (staging / "tz.eml").write_text(self.EML)
        out = Path(tmp.name) / "out"
        argv = ["outlook_to_md.py", str(staging.parent), str(out), *extra]
        with patch.object(sys, "argv", argv), redirect_stdout(io.StringIO()):
            outlook_to_md.main()
        md = next(out.rglob("email.md")).read_text()
        with open(out / "index.csv", newline="", encoding="utf-8") as f:
            row = next(csv.DictReader(f))
        return md, row

    def test_dates_render_in_the_named_zone(self):
        md, row = self.run_cli("--timezone", "Europe/London")
        self.assertIn('date: "2024-07-01T14:00:00+01:00"', md)
        self.assertEqual((row["date"], row["time"]), ("2024-07-01", "14:00:00"))

    def test_a_zone_across_midnight_moves_the_date(self):
        md, row = self.run_cli("--timezone", "Asia/Tokyo")
        self.assertIn('date: "2024-07-01T22:00:00+09:00"', md)
        md, row = self.run_cli("--timezone", "Pacific/Kiritimati")
        self.assertEqual(row["date"], "2024-07-02")

    def test_without_the_flag_the_sender_offset_is_kept(self):
        md, row = self.run_cli()
        self.assertIn('date: "2024-07-01T09:00:00-04:00"', md)
        self.assertEqual(row["time"], "09:00:00")

    def test_an_unknown_zone_is_refused(self):
        stderr = io.StringIO()
        with patch("sys.stderr", stderr), self.assertRaises(SystemExit) as caught:
            self.run_cli("--timezone", "Mars/Olympus_Mons")
        self.assertEqual(caught.exception.code, 2)
        self.assertIn("unknown timezone", stderr.getvalue())


class TestAppendRoundTrip(unittest.TestCase):
    """A re-run over the same staging directory must skip old mail and admit new mail.

    This is the property the outlook -> archive workflow relies on: a
    --since window that overlaps what is already archived costs bandwidth and
    nothing else, because Message-ID dedupe absorbs the overlap, while mail
    outside the overlap still lands. Pinning "adds nothing" alone would also
    pass under a broken append that discards every save, so a second test
    below checks that a genuinely new message still gets through.
    """

    EML = (
        "Message-ID: <roundtrip@example.com>\n"
        "Date: Tue, 29 Jul 2026 10:12:00 +0000\n"
        "From: Alice <alice@example.com>\n"
        "To: Bob <bob@example.com>\n"
        "Subject: Hello there\n"
        "Content-Type: text/plain; charset=utf-8\n"
        "\n"
        "Body text here.\n"
    )

    def test_second_append_run_is_a_noop(self):
        with tempfile.TemporaryDirectory() as tmp:
            staging = Path(tmp) / "staging" / "Inbox"
            staging.mkdir(parents=True)
            (staging / "20260729_101200_abc.eml").write_text(self.EML)
            out = Path(tmp) / "out"
            source = Path(tmp) / "staging"

            outlook_to_md.EmailExtractor(pst_path=source, output_dir=out).extract()

            first_rows = (out / "index.csv").read_text().splitlines()
            first_folders = sorted(p.parent.name for p in out.rglob("email.md"))
            self.assertEqual(len(first_folders), 1)

            outlook_to_md.EmailExtractor(pst_path=source, output_dir=out, append=True).extract()

            second_rows = (out / "index.csv").read_text().splitlines()
            second_folders = sorted(p.parent.name for p in out.rglob("email.md"))

            self.assertEqual(len(first_rows), len(second_rows), "append added an index row")
            self.assertEqual(first_folders, second_folders, "append added an email folder")

    NEW_EML = (
        "Message-ID: <second-message@example.com>\n"
        "Date: Wed, 30 Jul 2026 09:00:00 +0000\n"
        "From: Carol <carol@example.com>\n"
        "To: Bob <bob@example.com>\n"
        "Subject: A brand new thread\n"
        "Content-Type: text/plain; charset=utf-8\n"
        "\n"
        "Different body.\n"
    )

    def test_append_run_skips_duplicates_but_admits_new_mail(self):
        """The noop test alone cannot tell correct dedupe from "skip everything".

        A single-message round trip passes just as well under a broken
        implementation that discards every save once append=True. Pin the other
        half of the property: an already-archived message is skipped AND a
        genuinely new one - different Message-ID, sender, subject and date, so
        its derived folder name cannot collide with the first - still lands.
        """
        with tempfile.TemporaryDirectory() as tmp:
            staging = Path(tmp) / "staging" / "Inbox"
            staging.mkdir(parents=True)
            (staging / "20260729_101200_abc.eml").write_text(self.EML)
            out = Path(tmp) / "out"
            source = Path(tmp) / "staging"

            outlook_to_md.EmailExtractor(pst_path=source, output_dir=out).extract()

            first_rows = (out / "index.csv").read_text().splitlines()
            first_folders = sorted(p.parent.name for p in out.rglob("email.md"))
            self.assertEqual(len(first_folders), 1)

            # Simulate a --since window that overlaps the archive: the original
            # message is still on disk in staging, alongside one that is new.
            (staging / "20260730_090000_def.eml").write_text(self.NEW_EML)

            outlook_to_md.EmailExtractor(pst_path=source, output_dir=out, append=True).extract()

            second_rows = (out / "index.csv").read_text().splitlines()
            second_folders = sorted(p.parent.name for p in out.rglob("email.md"))

            self.assertEqual(len(second_folders), 2, "the new message was not added")
            self.assertEqual(
                len(second_rows),
                len(first_rows) + 1,
                "append did not add exactly one new index row",
            )
            self.assertTrue(
                set(second_folders) > set(first_folders),
                "the new folder does not extend the original one",
            )

    def test_staging_layout_becomes_archive_layout(self):
        with tempfile.TemporaryDirectory() as tmp:
            staging = Path(tmp) / "staging" / "Inbox" / "Clients"
            staging.mkdir(parents=True)
            (staging / "msg.eml").write_text(self.EML)
            out = Path(tmp) / "out"

            outlook_to_md.EmailExtractor(pst_path=Path(tmp) / "staging", output_dir=out).extract()

            found = list(out.rglob("email.md"))
            self.assertEqual(len(found), 1)
            # staging/Inbox/Clients/*.eml -> emails/Inbox/Clients/<folder>/email.md
            self.assertEqual(found[0].parent.parent.name, "Clients")
            self.assertEqual(found[0].parent.parent.parent.name, "Inbox")


class TestAppendWithoutMessageId(unittest.TestCase):
    """Documented limitation (see both SKILL.md caveats): dedupe needs a Message-ID.

    _save_email only checks self.existing_message_ids when message_id is
    truthy, so a message with no Message-ID header has no key to dedupe
    against and is re-archived on every overlapping append. This is not fixed
    here - a content-hash fallback is a separate design decision, out of
    scope for the caveat this pins - so the test exists to keep the
    behaviour honest: if it ever stops reproducing, the SKILL.md caveats
    this test backs are stale and must be revisited alongside it.
    """

    EML_NO_MESSAGE_ID = (
        "Date: Tue, 29 Jul 2026 10:12:00 +0000\n"
        "From: Alice <alice@example.com>\n"
        "To: Bob <bob@example.com>\n"
        "Subject: Draft with no Message-ID\n"
        "Content-Type: text/plain; charset=utf-8\n"
        "\n"
        "Body text here.\n"
    )

    def test_headerless_message_is_duplicated_on_append(self):
        with tempfile.TemporaryDirectory() as tmp:
            staging = Path(tmp) / "staging" / "Inbox"
            staging.mkdir(parents=True)
            (staging / "20260729_101200_abc.eml").write_text(self.EML_NO_MESSAGE_ID)
            out = Path(tmp) / "out"
            source = Path(tmp) / "staging"

            outlook_to_md.EmailExtractor(pst_path=source, output_dir=out).extract()
            first_rows = (out / "index.csv").read_text().splitlines()
            first_folders = sorted(p.parent.name for p in out.rglob("email.md"))
            self.assertEqual(len(first_folders), 1)

            # Same staging directory, re-appended - the overlap a --since
            # window would normally produce harmlessly.
            outlook_to_md.EmailExtractor(pst_path=source, output_dir=out, append=True).extract()
            second_rows = (out / "index.csv").read_text().splitlines()
            second_folders = sorted(p.parent.name for p in out.rglob("email.md"))

            self.assertEqual(
                len(second_folders),
                2,
                "a header-less message was not duplicated - if dedupe now covers "
                "this case, update the SKILL.md caveats and this test together",
            )
            self.assertTrue(any(f.endswith("-001") for f in second_folders))
            self.assertEqual(len(second_rows), len(first_rows) + 1, "expected one extra index row")


class TestManifestProvenance(unittest.TestCase):
    """--append must not erase the archive's original source hash.

    Regression pin for the chain-of-custody bug: _generate_manifest used to
    regenerate manifest.sha256 from scratch every run, keyed only on the
    CURRENT self.pst_path - so appending a staging directory into a
    PST-derived archive silently dropped the PST's own hash, even though
    README.md promises the manifest "records the source PST's own hash".
    """

    def _extract_stub_backend(self, pst_path, output_dir, **kwargs):
        """Run extract() with the real PST backend stubbed out.

        The provenance/manifest logic under test does not depend on what a
        real backend would parse out of pst_path - only on pst_path's own
        identity and hash - so stubbing avoids needing a real PST fixture
        (this suite deliberately carries none, see the module docstring) or
        an installed backend. Directory inputs are unaffected: they never
        reach readpst, so append-from-a-directory below exercises
        the real _process_eml_directory code path.
        """
        extractor = outlook_to_md.EmailExtractor(pst_path=pst_path, output_dir=output_dir, **kwargs)
        with patch.object(extractor, "_extract_with_readpst", lambda: None):
            extractor.extract()
        return extractor

    def test_append_preserves_the_original_source_hash(self):
        with tempfile.TemporaryDirectory() as tmp:
            pst_path = Path(tmp) / "archive.pst"
            pst_path.write_bytes(b"stand-in PST bytes; only the hash matters here")
            original_hash = outlook_to_md.compute_sha256(pst_path)

            out = Path(tmp) / "out"
            self._extract_stub_backend(pst_path, out)

            manifest_text = (out / "manifest.sha256").read_text()
            self.assertIn(original_hash, manifest_text, "sanity check: hash not even in the first manifest")

            # Append from an unrelated staging directory - a directory has no
            # hash of its own, which is exactly the case that lost the
            # original PST hash before the fix.
            staging = Path(tmp) / "staging"
            staging.mkdir()
            self._extract_stub_backend(staging, out, append=True)

            manifest_text = (out / "manifest.sha256").read_text()
            self.assertIn(original_hash, manifest_text, "the original PST's hash did not survive the append")

            # Stable across repeated runs: appending the same staging
            # directory again must not grow the source list without bound.
            self._extract_stub_backend(staging, out, append=True)
            manifest_text = (out / "manifest.sha256").read_text()
            self.assertEqual(
                manifest_text.count("source=archive.pst"),
                1,
                "repeating the same append duplicated the original source entry",
            )
            self.assertIn(original_hash, manifest_text)

    def test_index_md_totals_describe_the_archive_not_just_the_run(self):
        """The companion bug: index.md's totals used self.stats, a run-only counter.

        _load_existing_index rebuilds folder_counts/date_range from existing
        rows but never touches self.stats (deliberately - see its comment),
        so an append run's index.md showed only the new mail's count, not the
        archive's. This pins the fix: totals come from the merged
        self.index_data instead.
        """
        with tempfile.TemporaryDirectory() as tmp:
            staging1 = Path(tmp) / "staging1" / "Inbox"
            staging1.mkdir(parents=True)
            (staging1 / "a.eml").write_text(
                "Message-ID: <a@example.com>\n"
                "Date: Tue, 29 Jul 2026 10:12:00 +0000\n"
                "From: Alice <alice@example.com>\nTo: Bob <bob@example.com>\n"
                "Subject: First\nContent-Type: text/plain; charset=utf-8\n\nBody.\n"
            )
            out = Path(tmp) / "out"
            outlook_to_md.EmailExtractor(pst_path=Path(tmp) / "staging1", output_dir=out).extract()

            staging2 = Path(tmp) / "staging2" / "Inbox"
            staging2.mkdir(parents=True)
            (staging2 / "b.eml").write_text(
                "Message-ID: <b@example.com>\n"
                "Date: Wed, 30 Jul 2026 09:00:00 +0000\n"
                "From: Carol <carol@example.com>\nTo: Bob <bob@example.com>\n"
                "Subject: Second\nContent-Type: text/plain; charset=utf-8\n\nBody.\n"
            )
            outlook_to_md.EmailExtractor(pst_path=Path(tmp) / "staging2", output_dir=out, append=True).extract()

            index_md = (out / "index.md").read_text()
            self.assertIn("**Total Emails:** 2", index_md, "totals reflect this run only, not the whole archive")


class TestModuleContract(unittest.TestCase):
    """Guards against the optional-dependency wiring being removed."""

    def test_optional_dependency_flags_exist(self):
        for flag in ("HAS_DATEUTIL", "HAS_TQDM", "HAS_HTML2TEXT"):
            with self.subTest(flag=flag):
                self.assertIsInstance(getattr(outlook_to_md, flag), bool)

    def test_tqdm_fallback_is_iterable_when_absent(self):
        if outlook_to_md.HAS_TQDM:
            self.skipTest("tqdm is installed; the fallback is not in play")
        self.assertEqual(list(outlook_to_md.tqdm([1, 2, 3], desc="x")), [1, 2, 3])


if __name__ == "__main__":
    unittest.main()
