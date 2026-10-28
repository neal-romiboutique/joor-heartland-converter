#!/bin/bash
# Double-clickable launcher for the Romi Boutique Order Converter (macOS).
# Handles first-time setup itself so nobody has to open Terminal by hand.
cd "$(dirname "$0")"

# A downloaded zip marks every file "from the internet"; clear that once so
# Pull_Updates.command and friends don't each trigger a security prompt.
xattr -dr com.apple.quarantine . 2>/dev/null

# Fresh Mac: python3 is a stub until Apple's Command Line Tools are installed.
if ! xcode-select -p >/dev/null 2>&1; then
  xcode-select --install 2>/dev/null
  echo ""
  echo "FIRST-TIME SETUP: a window just opened asking to install developer tools."
  echo "Click Install and wait for it to finish (10-30 minutes)."
  echo "Then double-click 'Order Converter' again."
  echo ""
  read -p "Press Enter to close this window."
  exit 0
fi

# The two libraries the converter needs; installed once, then skipped.
if ! python3 -c "import openpyxl, requests" 2>/dev/null; then
  echo "First-time setup: installing helper libraries (about a minute)..."
  if ! python3 -m pip install --user --quiet --disable-pip-version-check openpyxl requests; then
    echo ""
    echo "Setup failed. Check the internet connection and double-click again."
    echo "If it keeps failing, send Neal a photo of this window."
    read -p "Press Enter to close this window."
    exit 1
  fi
fi

python3 converter_app.py
