#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
BRIDGE="$ROOT_DIR/skill/mac-powerpoint-live-builder/scripts/powerpoint_bridge.py"

if [ ! -f "$BRIDGE" ]; then
  echo "Bridge script not found: $BRIDGE" >&2
  exit 1
fi

echo "Starting PowerPoint live bridge..."
echo "Keep this window open while using WorkBuddy bridge mode."
echo
python3 "$BRIDGE"
