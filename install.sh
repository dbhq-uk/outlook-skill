#!/bin/bash
# Install Outlook skill to ~/.claude/skills/outlook
# Run from the tools/outlook directory

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TARGET_DIR="$HOME/.claude/skills/outlook"

echo "=== Outlook Skill Installer ==="
echo

# Check dependencies
echo "Checking dependencies..."
MISSING=""

if ! command -v az &> /dev/null; then
    MISSING="$MISSING azure-cli"
fi

if ! command -v jq &> /dev/null; then
    MISSING="$MISSING jq"
fi

if ! command -v curl &> /dev/null; then
    MISSING="$MISSING curl"
fi

if [ -n "$MISSING" ]; then
    echo "Missing required dependencies:$MISSING"
    echo
    echo "Install with:"
    echo "  macOS: brew install$MISSING"
    echo "  Ubuntu: sudo apt install$MISSING"
    exit 1
fi

echo "All required dependencies found"

# Optional dependency check
if ! command -v pandoc &> /dev/null; then
    echo "Optional: pandoc not found (needed for markdown-formatted emails)"
    echo "  Install with: brew install pandoc (macOS) or apt install pandoc (Linux)"
fi
echo

# Check for existing installation
if [ -d "$TARGET_DIR" ]; then
    echo "Existing installation found at $TARGET_DIR"
    read -p "Overwrite? (y/N): " overwrite
    if [[ ! "$overwrite" =~ ^[Yy]$ ]]; then
        echo "Installation cancelled"
        exit 0
    fi
    rm -rf "$TARGET_DIR"
fi

# Create target directory
echo "Installing to $TARGET_DIR..."
mkdir -p "$TARGET_DIR/scripts"
mkdir -p "$TARGET_DIR/references"

# Copy files
cp "$SCRIPT_DIR/SKILL.md" "$TARGET_DIR/"
cp "$SCRIPT_DIR/scripts/"* "$TARGET_DIR/scripts/"
cp "$SCRIPT_DIR/references/"* "$TARGET_DIR/references/"

# Make scripts executable
chmod +x "$TARGET_DIR/scripts/"*.sh

echo "Skill installed!"
echo

# Check if already configured
if [ -f "$HOME/.outlook/credentials.json" ]; then
    echo "Existing credentials found. Testing connection..."
    if "$TARGET_DIR/scripts/outlook-token.sh" test; then
        echo
        echo "=== Installation Complete ==="
        echo "Outlook skill is ready to use."
    else
        echo
        echo "Credentials exist but connection failed."
        read -p "Run setup to re-authenticate? (Y/n): " run_setup
        if [[ ! "$run_setup" =~ ^[Nn]$ ]]; then
            "$TARGET_DIR/scripts/outlook-setup.sh"
        fi
    fi
else
    echo "No credentials found. Running setup..."
    echo
    "$TARGET_DIR/scripts/outlook-setup.sh"
fi

echo
echo "=== Done ==="
echo
echo "Try these commands:"
echo "  Check email:    ~/.claude/skills/outlook/scripts/outlook-mail.sh inbox"
echo "  Today's calendar: ~/.claude/skills/outlook/scripts/outlook-calendar.sh today"
echo
echo "Or use natural language in Claude Code:"
echo "  'check my email'"
echo "  'what's on my calendar today'"
