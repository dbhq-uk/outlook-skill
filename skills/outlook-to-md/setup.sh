#!/bin/bash
# Set up the Python virtual environment for the outlook-to-md skill.
#
# The Python side is small (html2text, python-dateutil, tqdm) and installs on
# any python3 from 3.9 up. Reading a .pst file needs readpst, from pst-utils,
# which is a system package rather than a Python one - so this script checks
# for it and says how to install it, but cannot install it for you. A folder of
# .eml files (a live-mail export, or readpst run elsewhere) needs no readpst.
set -e

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$SKILL_DIR/.venv"
PYTHON="python3"

echo "==> Setting up outlook-to-md Python environment..."

if ! command -v "$PYTHON" &>/dev/null; then
    echo "Error: no python3 found on PATH" >&2
    exit 1
fi
if ! "$PYTHON" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)'; then
    echo "Error: outlook-to-md needs Python 3.9 or newer, found $("$PYTHON" -V 2>&1)" >&2
    exit 1
fi
echo "Interpreter: $PYTHON ($("$PYTHON" -V 2>&1))"

# --- Virtual environment ---------------------------------------------------
# A venv records the interpreter it was built with. One built on a different
# version (an older install capped itself at 3.11) is rebuilt on this one.
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
"$VENV/bin/pip" install --upgrade pip -q
"$VENV/bin/pip" install -r "$SKILL_DIR/requirements.txt" -q

echo "==> outlook-to-md environment ready"
echo
echo "Dependencies installed:"
"$VENV/bin/pip" list --format=columns 2>/dev/null | grep -iE "html2text|dateutil|tqdm" || true
echo

# --- PST reader --------------------------------------------------------------
if command -v readpst &>/dev/null; then
    echo "readpst: $(readpst -V 2>&1 | head -1)"
else
    echo "readpst: NOT FOUND. It is needed to read .pst files. Install pst-utils:"
    echo "  Ubuntu/Debian: sudo apt install pst-utils"
    echo "  macOS:         brew install libpst"
    echo "A folder of .eml files (a live-mail export) works without it."
fi
