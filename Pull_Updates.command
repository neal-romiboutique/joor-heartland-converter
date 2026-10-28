#!/bin/bash
# Double-click to fetch the newest version of the converter.
cd "$(dirname "$0")"
xattr -dr com.apple.quarantine . 2>/dev/null
echo "Checking for updates..."
if git pull --ff-only 2>/dev/null; then
  echo ""
  echo "Done. If the converter is open, close its Terminal window and"
  echo "double-click 'Order Converter' again so the update takes effect."
else
  echo ""
  echo "Couldn't update automatically. Nothing is broken: your current"
  echo "version keeps working. Ask Neal for the newest zip."
fi
echo ""
read -p "Press Enter to close this window."
