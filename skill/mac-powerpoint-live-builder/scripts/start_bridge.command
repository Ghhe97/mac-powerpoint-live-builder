#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
echo "Starting PowerPoint live bridge..."
echo "Keep this window open while using WorkBuddy bridge mode."
echo
python3 "$SCRIPT_DIR/powerpoint_bridge.py" --startup-self-test
