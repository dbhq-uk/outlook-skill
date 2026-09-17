#!/bin/bash
# Set up the Python virtual environment for the outlook-to-md skill.
#
# The awkward part is libratom, the preferred extraction backend. It pins
# numpy==1.23.5, whose newest wheel is cp311 - so on Python 3.12+ pip tries to
# build numpy from source and fails. Most current systems ship 3.12 or 3.13.
#
# Rather than fail the install or leave a venv that quietly lacks its main
# backend, this script:
#   1. prefers a libratom-capable interpreter (<= 3.11) if one is on PATH;
#   2. otherwise builds the venv on the default python3 WITHOUT libratom, and
#      says plainly that extraction then depends on readpst.
# outlook_to_md.py guards the libratom import and falls back to readpst, so the
# degraded mode is real and supported - it just must not be a silent surprise.
set -e

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$SKILL_DIR/.venv"

echo "==> Setting up outlook-to-md Python environment..."

# --- Choose an interpreter -------------------------------------------------
# Newest first, capped at 3.11: the cap is libratom's, not a preference.
PYTHON=""
for candidate in python3.11 python3.10 python3.9; do
    if command -v "$candidate" &>/dev/null; then PYTHON="$candidate"; break; fi
done

LIBRATOM_OK=1
if [ -z "$PYTHON" ]; then
    PYTHON="python3"
    if ! "$PYTHON" -c 'import sys; sys.exit(0 if sys.version_info < (3, 12) else 1)'; then
        LIBRATOM_OK=0
    fi
fi

if ! command -v "$PYTHON" &>/dev/null; then
    echo "Error: no python3 found on PATH" >&2
    exit 1
fi
echo "Interpreter: $PYTHON ($("$PYTHON" -V 2>&1))"

# --- Virtual environment ---------------------------------------------------
# A venv records the interpreter it was built with and ignores any later choice,
# so an existing one built on a different version must be replaced - otherwise
# picking a libratom-capable interpreter above silently achieves nothing.
WANT="$("$PYTHON" -c 'import sys; print("%d.%d" % sys.version_info[:2])')"
if [ -d "$VENV" ]; then
    HAVE="$("$VENV/bin/python" -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>/dev/null || echo none)"
    if [ "$HAVE" != "$WANT" ]; then
        echo "Existing venv is Python $HAVE, rebuilding on $WANT..."
        rm -rf "$VENV"
    fi
fi
if [ ! -d "$VENV" ]; then
    "$PYTHON" -m venv "$VENV"
    echo "Created virtual environment (Python $WANT)"
fi

echo "Installing dependencies..."
# setuptools and wheel first: libratom's transitive numpy==1.23.5 arrives as an
# sdist and needs a build backend present to configure.
#
# SETUPTOOLS IS PINNED BELOW 81 AND MUST STAY THERE. 81 removed pkg_resources,
# and libratom's spacy dependency imports it at module level - so an unpinned
# --upgrade installs 82, and every ratom call fails with ModuleNotFoundError:
# No module named 'pkg_resources'. That is not a build-time warning, it is the
# skill not working at all, and it broke this install silently (17 Sep 2026).
"$VENV/bin/pip" install --upgrade pip wheel -q
"$VENV/bin/pip" install --upgrade "setuptools<81" -q

if [ "$LIBRATOM_OK" = 1 ]; then
    "$VENV/bin/pip" install -r "$SKILL_DIR/requirements.txt" -q
else
    grep -v '^libratom' "$SKILL_DIR/requirements.txt" > "$VENV/requirements-no-libratom.txt"
    "$VENV/bin/pip" install -r "$VENV/requirements-no-libratom.txt" -q
fi

echo "==> outlook-to-md environment ready"
echo
echo "Dependencies installed:"
"$VENV/bin/pip" list --format=columns 2>/dev/null | grep -iE "libratom|html2text|dateutil|tqdm" || true
echo

# --- Backend availability --------------------------------------------------
echo "Extraction backends:"
if "$VENV/bin/python" -c 'from libratom.lib.pff import PffArchive' 2>/dev/null; then
    echo "  libratom: OK (preferred backend)"
else
    echo "  libratom: NOT INSTALLED"
    if [ "$LIBRATOM_OK" = 0 ]; then
        echo "    Reason: libratom pins numpy==1.23.5, which has no wheel above"
        echo "    Python 3.11 and cannot build on $("$PYTHON" -V 2>&1 | cut -d' ' -f2)."
        echo "    Fix by installing a 3.11 interpreter and re-running this script,"
        echo "    or rely on readpst below."
    fi
fi

if command -v readpst &>/dev/null; then
    echo "  readpst:  $(readpst -V 2>&1 | head -1)"
else
    echo "  readpst:  NOT FOUND"
    echo "    Ubuntu/Debian: sudo apt install pst-utils"
    echo "    macOS: brew install libpst"
fi

if ! "$VENV/bin/python" -c 'from libratom.lib.pff import PffArchive' 2>/dev/null \
   && ! command -v readpst &>/dev/null; then
    echo
    echo "WARNING: neither libratom nor readpst is available, so PST files cannot"
    echo "be read. Directory mode (a folder of pre-extracted .eml files) still works."
fi
