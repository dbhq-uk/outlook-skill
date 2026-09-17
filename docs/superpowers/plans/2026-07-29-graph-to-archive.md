# Keeping a PST archive current from live mail — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let `outlook` export live mail as `.eml` so `outlook-to-md` can append it to an existing archive in the identical shape.

**Architecture:** A new `export` verb in `outlook-mail.sh` writes raw MIME from Graph's `$value` endpoint into a staging directory whose layout mirrors the mail folder. `outlook_to_md.py` then consumes that directory in its existing directory mode with `--append`, which dedupes by `Message-ID`. The staging format is plain `.eml` and nothing more, so the archive format stays owned by `outlook-to-md` alone.

**Tech Stack:** Bash + curl + jq (outlook); Python 3.9+ stdlib + `unittest` (outlook-to-md).

**Spec:** `docs/superpowers/specs/2026-07-29-graph-to-archive-design.md`

## Global Constraints

- **Python floor is 3.9.** `parse_email_address` annotates `tuple[str, str]` at runtime with no `from __future__ import annotations` in `outlook_to_md.py`, so PEP 585 builtin generics must exist. CI runs 3.9, 3.11, 3.13.
- **Bash, not Python, for the outlook side.** Every script in that skill is bash + curl + jq; do not introduce a new runtime.
- **Never leave a Graph error body on disk named `.eml`.** Use `curl -sf` and delete the part-written file on failure, matching the attachment downloader at `outlook-mail.sh:1930`.
- **Test helpers must be top-level functions.** `helpers_test.sh` extracts functions by matching `^name() {` through the first line that is exactly `}` (`extract_fn`, line 31). Logic buried in a `case` branch cannot be tested offline.
- **British English in all prose and comments.** Match the surrounding files.
- **Comments explain why, not what.** The existing code comments justify non-obvious decisions; match that density and tone.

---

### Task 1: Filename and `--since` helpers

Two pure functions, no network. They go in `outlook-mail.sh` just above the `# Commands` marker (currently line 568), after `run_message_search`.

**Files:**
- Modify: `skills/outlook/scripts/outlook-mail.sh` (insert before `# Commands`, line 568)
- Test: `skills/outlook/tests/helpers_test.sh`

**Interfaces:**
- Consumes: `urlencode` (already defined, line 421)
- Produces:
  - `export_eml_filename <receivedDateTime> <message-id>` → prints `YYYYMMDD_HHMMSS_<short20>.eml`
  - `export_since_filter <YYYY-MM-DD>` → prints `receivedDateTime ge <date>T00:00:00Z`, or returns 1 with a message on stderr

- [ ] **Step 1: Write the failing tests**

Add to `skills/outlook/tests/helpers_test.sh`, immediately before the final `rm -f "$body_file" ...` cleanup line:

```bash
########################################
# export helpers: staging filename + --since validation.
# The filename only has to be unique and sortable - outlook_to_md.py derives the
# archive folder name from the message's own headers, not from this name.
########################################
eval "$(extract_fn export_eml_filename)"
eval "$(extract_fn export_since_filter)"

eq "export filename from Graph timestamp" "20260729_101200_AAAAAAAAAAAAAAAAAAAA.eml" \
   "$(export_eml_filename '2026-07-29T10:12:00Z' 'PREFIXAAAAAAAAAAAAAAAAAAAA')"
eq "export filename drops fractional seconds" "20260729_101200_abcdefghijklmnopqrst.eml" \
   "$(export_eml_filename '2026-07-29T10:12:00.5230000Z' 'abcdefghijklmnopqrst')"
# The id is server-supplied; a '/' in it would escape the output directory.
case "$(export_eml_filename '2026-07-29T10:12:00Z' 'aaaa/bbbb/cccc/dddd/eeee')" in
  */*) eq "export filename strips path separators" "no slash" "contains a slash";;
  *)   eq "export filename strips path separators" ok ok;;
esac

eq "since filter builds a Graph filter" "receivedDateTime ge 2026-07-01T00:00:00Z" \
   "$(export_since_filter '2026-07-01')"
# A malformed date must fail here. Sent to Graph it comes back as a generic
# BadRequest, or worse silently matches nothing and looks like "no new mail".
eq "since rejects a human-written date" "1" "$(export_since_filter '1 July 2026' 2>/dev/null; echo $?)"
eq "since rejects an empty value"       "1" "$(export_since_filter '' 2>/dev/null; echo $?)"
eq "since rejects a partial date"       "1" "$(export_since_filter '2026-07' 2>/dev/null; echo $?)"
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `bash skills/outlook/tests/helpers_test.sh`
Expected: FAIL — `extract_fn` returns nothing for the two undefined functions, so the `eval`s are no-ops and each `export_*` call reports "command not found" with empty output.

- [ ] **Step 3: Write the implementation**

Insert into `skills/outlook/scripts/outlook-mail.sh` immediately before the `# Commands` comment:

```bash
# --- .eml export ------------------------------------------------------------
# Staging filename for one exported message. outlook_to_md.py derives the archive
# folder name from the message's own headers, so this name only has to be unique
# and to sort chronologically in a directory listing. The short id is the same
# 20-character slice the listing commands print, so an id copied off a listing
# appears verbatim here. '/' is stripped because the id is server-supplied and
# would otherwise let a crafted id escape the output directory.
export_eml_filename() {
    local received="$1" msg_id="$2" stamp short
    stamp=$(printf '%s' "$received" | tr -d ':-' | sed 's/\.[0-9]*//; s/Z$//; s/T/_/')
    short=$(printf '%s' "${msg_id: -20}" | tr -d '/')
    printf '%s_%s.eml' "$stamp" "$short"
}

# Validate a --since value and print the Graph $filter fragment for it. The
# validation is the point: an unparseable date reaches Graph as a filter that
# either 400s or matches nothing, and "matches nothing" is indistinguishable
# from "no new mail" - so a typo would silently look like a clean sync.
export_since_filter() {
    local since="$1"
    case "$since" in
        [0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]) ;;
        *) echo "Error: --since expects YYYY-MM-DD, got '$since'" >&2; return 1 ;;
    esac
    printf 'receivedDateTime ge %sT00:00:00Z' "$since"
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `bash skills/outlook/tests/helpers_test.sh`
Expected: PASS, with `FAIL=0` on the summary line.

- [ ] **Step 5: Commit**

```bash
git add skills/outlook/scripts/outlook-mail.sh skills/outlook/tests/helpers_test.sh
git commit -m "feat(export): staging filename and --since validation helpers"
```

---

### Task 2: Message-listing helper with paging

**Files:**
- Modify: `skills/outlook/scripts/outlook-mail.sh` (after `export_since_filter` from Task 1)
- Test: `skills/outlook/tests/helpers_test.sh`

**Interfaces:**
- Consumes: `api_call` (line 157), `urlencode` (line 421), `export_since_filter` (Task 1), `GRAPH_URL`
- Produces: `export_list_messages <folder-id> <since-or-empty> <cap>` → prints `{"value":[{"id":…,"receivedDateTime":…}, …]}`, newest first, at most `cap` entries; returns 1 on a Graph error

- [ ] **Step 1: Write the failing tests**

Append to `skills/outlook/tests/helpers_test.sh`, after the Task 1 block:

```bash
########################################
# export_list_messages: paging + cap + --since encoding + error propagation.
# $top caps at 1000 per page and a mail folder can hold far more, so paging is
# required rather than optional.
########################################
eval "$(extract_fn export_list_messages)"

api_call() {
    local endpoint="$2"; printf '%s' "$endpoint" > /tmp/outlook_test_last_url
    if [[ "$endpoint" == *skiptoken* ]]; then
        echo '{"value":[{"id":"m3","receivedDateTime":"2026-04-01T00:00:00Z"}]}'
    else
        echo '{"@odata.nextLink":"'"$GRAPH_URL"'/me/messages?$skiptoken=ABC","value":[{"id":"m1","receivedDateTime":"2026-03-01T00:00:00Z"},{"id":"m2","receivedDateTime":"2026-02-01T00:00:00Z"}]}'
    fi
}
eq "export pages through nextLink" "3" "$(export_list_messages FID '' 100 | jq '.value|length')"
eq "export cap stops paging"       "2" "$(export_list_messages FID '' 2 | jq '.value|length')"
eq "export sorts newest-first" "m3,m1,m2" \
   "$(export_list_messages FID '' 100 | jq -r '[.value[].id]|join(",")')"

export_list_messages FID '2026-07-01' 5 >/dev/null
eq "export encodes --since into \$filter" "1" \
   "$(grep -c 'receivedDateTime%20ge%202026-07-01T00%3A00%3A00Z' /tmp/outlook_test_last_url)"
eq "export scopes the query to the folder" "1" \
   "$(grep -c '/me/mailFolders/FID/messages' /tmp/outlook_test_last_url)"

# A bad --since must stop before any request is issued.
eq "export rejects a bad --since without calling Graph" "1" \
   "$(export_list_messages FID 'last tuesday' 5 >/dev/null 2>&1; echo $?)"

api_call() { echo '{"error":{"code":"BadRequest","message":"nope"}}'; }
eq "export propagates a Graph error as rc1" "1" \
   "$(export_list_messages FID '' 10 >/dev/null 2>&1; echo $?)"
eq "export reports the Graph error message" "1" \
   "$(export_list_messages FID '' 10 2>&1 >/dev/null | grep -c 'nope')"
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `bash skills/outlook/tests/helpers_test.sh`
Expected: FAIL — `export_list_messages` is not yet defined.

- [ ] **Step 3: Write the implementation**

Insert after `export_since_filter`:

```bash
# List a folder's messages as {"value":[{id, receivedDateTime}, …]}, newest
# first, capped at $3. Paging is not optional: $top caps at 1000 per page and a
# real mail folder holds far more, so a single request would silently export a
# prefix of the folder and look complete.
#
# Returns 1 (rather than printing an error object, as run_message_search does)
# because the caller writes files: it must stop, not iterate over an error.
export_list_messages() {
    local folder_id="$1" since="$2" max="${3:-1000}"
    [[ "$max" =~ ^[0-9]+$ ]] || max=1000
    [ "$max" -lt 1 ] && max=1

    local page_size url filter merged page collected next
    page_size=200
    [ "$max" -lt "$page_size" ] && page_size="$max"

    url="/me/mailFolders/$folder_id/messages?\$top=${page_size}&\$orderby=receivedDateTime%20desc&\$select=id,receivedDateTime"
    if [ -n "$since" ]; then
        filter=$(export_since_filter "$since") || return 1
        url="${url}&\$filter=$(urlencode "$filter")"
    fi

    merged='[]'
    while [ -n "$url" ]; do
        page=$(api_call GET "$url")
        if echo "$page" | jq -e '.error' >/dev/null 2>&1; then
            echo "Error: $(echo "$page" | jq -r '.error.message // .error.code')" >&2
            return 1
        fi
        merged=$(jq -n --argjson a "$merged" --argjson b "$(echo "$page" | jq '.value // []')" '$a + $b')
        collected=$(echo "$merged" | jq 'length')
        [ "$collected" -ge "$max" ] && break
        next=$(echo "$page" | jq -r '."@odata.nextLink" // empty')
        [ -z "$next" ] && break
        url="${next#"$GRAPH_URL"}"    # nextLink is absolute; strip base for api_call
    done
    echo "$merged" | jq --argjson max "$max" '{value: (sort_by(.receivedDateTime) | reverse | .[0:$max])}'
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `bash skills/outlook/tests/helpers_test.sh`
Expected: PASS, `FAIL=0`.

- [ ] **Step 5: Commit**

```bash
git add skills/outlook/scripts/outlook-mail.sh skills/outlook/tests/helpers_test.sh
git commit -m "feat(export): folder message listing with paging, cap and --since"
```

---

### Task 3: The `export` verb

**Files:**
- Modify: `skills/outlook/scripts/outlook-mail.sh` (new `case` branch; help text near line 2103)

**Interfaces:**
- Consumes: `resolve_folder_id` (line 266), `export_list_messages` (Task 2), `export_eml_filename` (Task 1), `ACCESS_TOKEN`, `GRAPH_URL`
- Produces: `outlook-mail.sh export <folder> <output-dir> [--since YYYY-MM-DD] [--count N]`

- [ ] **Step 1: Add the `case` branch**

Insert into `skills/outlook/scripts/outlook-mail.sh` immediately after the `download)` branch ends (the `;;` following the attachment loop) and before `attach)`:

```bash
    export)
        folder_name="$2"
        out_dir="$3"
        if [ -z "$folder_name" ] || [ -z "$out_dir" ]; then
            echo "Usage: outlook-mail.sh export <folder> <output-dir> [--since YYYY-MM-DD] [--count N]"
            echo "       Writes each message as raw .eml for outlook-to-md to append."
            exit 1
        fi
        shift 3

        since=""
        cap=1000
        while [ $# -gt 0 ]; do
            case "$1" in
                --since) [ -n "$2" ] || { echo "Error: --since requires a date" >&2; exit 1; }; since="$2"; shift 2 ;;
                --count) [ -n "$2" ] || { echo "Error: --count requires a number" >&2; exit 1; }; cap="$2"; shift 2 ;;
                *) echo "Error: unknown option '$1'" >&2; exit 1 ;;
            esac
        done

        if ! folder_id=$(resolve_folder_id "$folder_name"); then
            echo "Error: Folder not found: $folder_name" >&2
            exit 1
        fi

        # The staging sub-path becomes the archive's folder grouping, so mirror
        # the folder the user named. '..' is stripped and leading slashes
        # trimmed so a folder name cannot climb out of $out_dir.
        rel_dir=$(printf '%s' "$folder_name" | sed 's#\.\.##g; s#^/*##; s#/*$##')
        dest_dir="$out_dir/$rel_dir"
        mkdir -p "$dest_dir"

        messages=$(export_list_messages "$folder_id" "$since" "$cap") || exit 1
        total=$(printf '%s' "$messages" | jq '.value | length')
        if [ "$total" = "0" ]; then
            echo "No messages to export from '$folder_name'."
            exit 0
        fi

        echo "Exporting $total message(s) from '$folder_name' to $dest_dir ..."
        written=0
        failed=0
        # Process substitution, NOT a pipe: a piped `while` runs in a subshell,
        # so $written and $failed would be discarded and the summary would
        # always report zero.
        while IFS=$'\t' read -r msg_id received; do
            [ -n "$msg_id" ] || continue
            fname=$(export_eml_filename "$received" "$msg_id")
            # -f so an HTTP error is a curl failure rather than a Graph error
            # body saved as an .eml, which the next step would parse as mail.
            if ! curl -sf -X GET "${GRAPH_URL}/me/messages/$msg_id/\$value" \
                -H "Authorization: Bearer $ACCESS_TOKEN" \
                -o "$dest_dir/$fname"; then
                rm -f "$dest_dir/$fname"
                echo "FAILED: $fname (MIME fetch error from Graph)"
                failed=$((failed + 1))
                continue
            fi
            written=$((written + 1))
        done < <(printf '%s' "$messages" | jq -r '.value[] | "\(.id)\t\(.receivedDateTime)"')

        echo "Wrote $written .eml file(s) to $dest_dir"
        if [ "$failed" -gt 0 ]; then
            echo "$failed message(s) failed and were not written."
        fi
        echo
        echo "Append to a markdown archive with:"
        echo "  outlook_to_md.py $out_dir <archive-dir> --append"
        ;;
```

- [ ] **Step 2: Add the help line**

In the `*)` usage block, add this line at the end of the "Reading:" group — directly after the `preview <id>              Quick preview` line (around line 2103) and before the blank `echo` that precedes "Sending:":

```bash
        echo "  export <folder> <dir>      Write folder's messages as .eml for archiving"
```

- [ ] **Step 3: Verify the script still parses and the helper tests still pass**

Run: `bash -n skills/outlook/scripts/outlook-mail.sh && bash skills/outlook/tests/helpers_test.sh`
Expected: no syntax output, and `FAIL=0`.

- [ ] **Step 4: Verify argument handling without a mailbox**

The credentials guard at the top of the script exits before any command runs when no account is configured, so run these against a configured account, or accept the credentials error as proof the branch was reached:

Run: `skills/outlook/scripts/outlook-mail.sh export 2>&1 | head -3`
Expected: the usage line `Usage: outlook-mail.sh export <folder> <output-dir> …` (or the "Account not configured" error if no account exists — in which case verify the branch by inspection instead).

- [ ] **Step 5: Commit**

```bash
git add skills/outlook/scripts/outlook-mail.sh
git commit -m "feat(export): add the export verb writing folder mail as .eml"
```

---

### Task 4: Fix directory-mode backend dispatch

A directory input is currently only honoured when both libratom and readpst are absent, because the directory branch is nested inside `_extract_with_readpst`. With libratom installed — what `setup.sh` aims for — the directory is passed to `PffArchive` and raises `OSError: … Is a directory`.

**Files:**
- Modify: `skills/outlook-to-md/scripts/outlook_to_md.py:257-260`
- Test: `skills/outlook-to-md/tests/test_outlook_to_md.py`

**Interfaces:**
- Consumes: `EmailExtractor.pst_path`, `EmailExtractor._process_eml_directory`, module-level `USE_LIBRATOM`
- Produces: no signature change; `extract()` gains a directory branch ahead of backend selection

- [ ] **Step 1: Write the failing test**

Add to `skills/outlook-to-md/tests/test_outlook_to_md.py`, after the `TestAppendModeIndexLoading` class:

```python
class TestDirectoryDispatch(unittest.TestCase):
    """A directory input must reach directory mode whatever backends exist.

    The directory branch used to sit inside the readpst path, reachable only
    when readpst was ALSO missing - so with libratom installed (what setup.sh
    aims for) a directory was handed to PffArchive and raised OSError.
    """

    def test_directory_input_bypasses_libratom(self):
        with tempfile.TemporaryDirectory() as tmp:
            staging = Path(tmp) / "staging"
            staging.mkdir()
            (staging / "a.eml").write_text("Subject: x\n\nbody\n")
            out = Path(tmp) / "out"

            extractor = outlook_to_md.EmailExtractor(pst_path=staging, output_dir=out)

            called = {"dir": False, "libratom": False}

            def fake_dir(path):
                called["dir"] = True

            def fake_libratom():
                called["libratom"] = True

            with patch.object(extractor, "_process_eml_directory", fake_dir), patch.object(
                extractor, "_extract_with_libratom", fake_libratom
            ), patch.object(outlook_to_md, "USE_LIBRATOM", True):
                extractor.extract()

            self.assertTrue(called["dir"], "directory input did not reach directory mode")
            self.assertFalse(called["libratom"], "directory input was sent to libratom")
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `cd skills/outlook-to-md && ./.venv/bin/python -m pytest tests/test_outlook_to_md.py::TestDirectoryDispatch -v -p no:cacheprovider`
Expected: FAIL — `directory input was sent to libratom` (or an assertion that `called["dir"]` is False).

- [ ] **Step 3: Write the implementation**

In `skills/outlook-to-md/scripts/outlook_to_md.py`, replace the dispatch inside `extract()`:

```python
        if USE_LIBRATOM:
            self._extract_with_libratom()
        else:
            self._extract_with_readpst()
```

with:

```python
        # Test the input before the backend. A directory of .eml is a documented
        # first-class input, but this branch used to live inside the readpst
        # path, so it was only reachable when readpst was ALSO missing - and a
        # directory handed to libratom raises "OSError: ... Is a directory".
        if self.pst_path.is_dir():
            print(f"Processing pre-extracted emails from: {self.pst_path}")
            self._process_eml_directory(self.pst_path)
        elif USE_LIBRATOM:
            self._extract_with_libratom()
        else:
            self._extract_with_readpst()
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `cd skills/outlook-to-md && ./.venv/bin/python -m pytest tests/ -v -p no:cacheprovider`
Expected: PASS — the new test plus the existing 47 passed, 1 skipped.

- [ ] **Step 5: Commit**

```bash
git add skills/outlook-to-md/scripts/outlook_to_md.py skills/outlook-to-md/tests/test_outlook_to_md.py
git commit -m "fix(pst): honour a directory input regardless of installed backend

Directory mode was nested inside the readpst path, so it was reachable only
when readpst was also absent. With libratom installed - what setup.sh works to
achieve - a directory was passed to PffArchive and died with 'Is a directory',
despite SKILL.md documenting directory input as first-class."
```

---

### Task 5: Append round-trip test

Proves the property the whole design rests on: re-running over the same staging directory adds nothing.

**Files:**
- Test: `skills/outlook-to-md/tests/test_outlook_to_md.py`

**Interfaces:**
- Consumes: `EmailExtractor(pst_path=…, output_dir=…, append=…)`, `extract()`; the dispatch fix from Task 4
- Produces: nothing consumed by later tasks

- [ ] **Step 1: Write the failing test**

Add to `skills/outlook-to-md/tests/test_outlook_to_md.py`, after `TestDirectoryDispatch`:

```python
class TestAppendRoundTrip(unittest.TestCase):
    """Second run over the same staging directory must add nothing.

    This is the property the outlook -> archive workflow relies on: a
    --since window that overlaps what is already archived costs bandwidth and
    nothing else, because Message-ID dedupe absorbs the overlap.
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
```

- [ ] **Step 2: Run the tests to verify they pass**

These should pass once Task 4 is in. Run: `cd skills/outlook-to-md && ./.venv/bin/python -m pytest tests/test_outlook_to_md.py::TestAppendRoundTrip -v -p no:cacheprovider`
Expected: PASS. If `test_second_append_run_is_a_noop` fails, the dedupe is genuinely broken — stop and investigate rather than adjusting the test.

- [ ] **Step 3: Run the whole suite**

Run: `cd skills/outlook-to-md && ./.venv/bin/python -m pytest tests/ -v -p no:cacheprovider`
Expected: PASS, no failures.

- [ ] **Step 4: Commit**

```bash
git add skills/outlook-to-md/tests/test_outlook_to_md.py
git commit -m "test(pst): pin append idempotence and staging-to-archive layout"
```

---

### Task 6: Document the workflow

**Files:**
- Modify: `skills/outlook/SKILL.md` (new section after "Attachments", before "Email Management")
- Modify: `skills/outlook-to-md/SKILL.md` (new section after "Incremental Extraction (Append Mode)")
- Modify: `README.md` (extend the "PST archives" section, around line 132)

- [ ] **Step 1: Add the outlook section**

In `skills/outlook/SKILL.md`, after the "### Attachments" block:

````markdown
### Exporting Mail to a Markdown Archive

Write a folder's messages out as raw `.eml`, then let `outlook-to-md` append
them to an archive. The PST backfills history; this keeps it current.

```bash
# Everything in a folder
${CLAUDE_SKILL_DIR}/scripts/outlook-mail.sh export "Inbox/Clients" ./staging/

# Only what arrived since a date (use the archive's newest entry)
${CLAUDE_SKILL_DIR}/scripts/outlook-mail.sh export "Inbox/Clients" ./staging/ --since 2026-07-01

# Then append into the archive - dedupes by Message-ID, so an overlapping
# --since window is harmless
${CLAUDE_SKILL_DIR}/../outlook-to-md/.venv/bin/python \
  ${CLAUDE_SKILL_DIR}/../outlook-to-md/scripts/outlook_to_md.py \
  ./staging/ ./archive/ --append
```

The staging directory's layout becomes the archive's folder grouping, so
`export "Inbox/Clients"` lands under `emails/Inbox/Clients/`.
````

- [ ] **Step 2: Add the outlook-to-md section**

In `skills/outlook-to-md/SKILL.md`, after the "### Incremental Extraction (Append Mode)" block:

````markdown
### Keeping an Archive Current from Live Mail

A PST is a snapshot. To carry an archive forward, export new mail with the
sibling `outlook` skill and append it — the two produce the same shape.

```bash
# 1. Export live mail as .eml (needs outlook configured)
${CLAUDE_SKILL_DIR}/../outlook/scripts/outlook-mail.sh \
  export "Inbox/Clients" ./staging/ --since 2026-07-01

# 2. Append it to the existing archive
${CLAUDE_SKILL_DIR}/.venv/bin/python ${CLAUDE_SKILL_DIR}/scripts/outlook_to_md.py \
  ./staging/ ./archive/ --append
```

Deduplication is by `Message-ID`, so a `--since` window that overlaps what is
already archived costs bandwidth and nothing else. Graph-sourced mail is
recorded under the `pst_folder` index column like any other — the column means
"the folder this message came from", and always did.
````

- [ ] **Step 3: Extend the README**

In `README.md`, after the paragraph ending "Nothing is uploaded and nothing is sent." (around line 132), add:

````markdown
A PST is a snapshot, so the two skills join up to carry an archive forward:
`outlook` exports new mail as `.eml` and `outlook-to-md` appends it in
the same shape, deduplicating by `Message-ID`.

```bash
outlook-mail.sh export "Inbox/Clients" ./staging/ --since 2026-07-01
outlook_to_md.py ./staging/ ./archive/ --append
```
````

- [ ] **Step 4: Verify the docs gates still pass**

The `validate` job checks each `SKILL.md`'s frontmatter and that its description stays under 1024 characters. Neither section touches frontmatter, but run the check:

Run: `bash -n skills/outlook/scripts/outlook-mail.sh && python3 -c "
import re,sys
for f in ['skills/outlook/SKILL.md','skills/outlook-to-md/SKILL.md']:
    t=open(f).read()
    m=re.match(r'^---\n(.*?)\n---\n', t, re.S)
    assert m, f+': missing frontmatter'
    d=re.search(r'^description:\s*(.+)$', m.group(1), re.M)
    assert d and len(d.group(1))<=1024, f+': description problem'
    print(f, 'OK')
"`
Expected: both files print `OK`.

- [ ] **Step 5: Commit**

```bash
git add skills/outlook/SKILL.md skills/outlook-to-md/SKILL.md README.md
git commit -m "docs: document exporting live mail into a markdown archive"
```

---

### Task 7: Full verification and PR

- [ ] **Step 1: Run every gate the CI runs**

```bash
fail=0
while IFS= read -r s; do bash -n "$s" || { echo "SYNTAX: $s"; fail=1; }; done \
  < <({ find . -maxdepth 1 -name '*.sh'; find skills -name '*.sh'; })
echo "shell parse fail=$fail"

find skills -name '*.py' -not -path '*/.venv/*' \
  -exec python3 -c "import ast,sys; ast.parse(open(sys.argv[1]).read())" {} \;
echo "python parse ok"

bash skills/outlook/tests/helpers_test.sh
(cd skills/outlook-to-md && ./.venv/bin/python -m pytest tests/ -q -p no:cacheprovider)
```

Expected: `fail=0`, `FAIL=0` from the bash suite, and no pytest failures.

- [ ] **Step 2: Prove the round trip on real data**

Only if an Outlook account is configured. Pick a small folder:

```bash
skills/outlook/scripts/outlook-mail.sh export "Inbox" /tmp/rt-staging --count 3
skills/outlook-to-md/.venv/bin/python skills/outlook-to-md/scripts/outlook_to_md.py \
  /tmp/rt-staging /tmp/rt-archive --append
find /tmp/rt-archive/emails -name email.md | head
# second run must add nothing
skills/outlook-to-md/.venv/bin/python skills/outlook-to-md/scripts/outlook_to_md.py \
  /tmp/rt-staging /tmp/rt-archive --append | grep -i skip
sha256sum -c /tmp/rt-archive/manifest.sha256 | tail -3
rm -rf /tmp/rt-staging /tmp/rt-archive
```

Expected: three email folders, the second run reporting 3 skipped and 0 processed, and the manifest verifying. If no account is configured, say so explicitly rather than claiming this step passed.

- [ ] **Step 3: Open the PR**

```bash
git push -u origin feat/graph-to-archive
gh pr create --base main --title "feat: keep a PST archive current from live mail" --body "$(cat <<'EOF'
Joins the pack's two halves: the PST backfills history, `outlook` appends
new mail into the identical archive shape.

## What changed

- `outlook-mail.sh export <folder> <dir> [--since] [--count]` writes a
  folder's messages as raw `.eml` via Graph's `$value` MIME endpoint, paging
  through `@odata.nextLink`.
- `outlook_to_md.py` now tests for a directory input *before* selecting a
  backend. Directory mode was nested inside the readpst path, so it only worked
  when readpst was also absent - with libratom installed (what `setup.sh` aims
  for) a directory went to `PffArchive` and died with "Is a directory", despite
  `SKILL.md` documenting it as first-class. Found while validating the design.
- Both `SKILL.md`s and the README document the two-step workflow.

Staging stays plain `.eml`: `export` knows nothing about `email.md`, checksums
or the index, so the archive format stays owned by `outlook-to-md`.

## Verified

- `helpers_test.sh` — filename construction, `--since` validation, paging, cap,
  error propagation
- `pytest tests/` — directory dispatch regression, append idempotence,
  staging-to-archive layout
- shell and Python parse gates
- Real-mailbox round trip: STATE THE RESULT HERE, or say plainly that no
  account was configured and this was not run.
EOF
)"
```

Replace the last bullet with what actually happened. If Step 2 did not run, say so — do not imply it passed.

- [ ] **Step 4: Merge once checks pass**

```bash
gh pr checks $(gh pr view --json number -q .number)
gh pr merge $(gh pr view --json number -q .number) --squash --delete-branch
```
