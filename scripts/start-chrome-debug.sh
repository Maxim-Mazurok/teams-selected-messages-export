#!/bin/zsh

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
PROFILE_DIR="$ROOT_DIR/.runtime/chrome-profile"
DEBUG_PORT="${DEBUG_PORT:-9222}"
TARGET_URL="${1:-https://teams.microsoft.com}"
CHROME_BIN="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

if [[ ! -x "$CHROME_BIN" ]]; then
  echo "Google Chrome was not found at $CHROME_BIN" >&2
  exit 1
fi

mkdir -p "$PROFILE_DIR"

open -na "Google Chrome" --args \
  --remote-debugging-port="$DEBUG_PORT" \
  --user-data-dir="$PROFILE_DIR" \
  --new-window \
  "$TARGET_URL"

echo "Chrome launched with remote debugging on port $DEBUG_PORT"
echo "Profile directory: $PROFILE_DIR"
