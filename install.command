#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILL_NAME="mac-powerpoint-live-builder"
SOURCE_SKILL="$ROOT_DIR/skill/$SKILL_NAME"
DEST_ROOT="$HOME/.codex/skills"
DEST_SKILL="$DEST_ROOT/$SKILL_NAME"

echo "Mac PowerPoint Live Builder installer"
echo

if [ ! -d "$SOURCE_SKILL" ]; then
  echo "Error: skill folder not found: $SOURCE_SKILL" >&2
  exit 1
fi

mkdir -p "$DEST_ROOT"

if [ -d "$DEST_SKILL" ]; then
  BACKUP="$DEST_SKILL.backup.$(date +%Y%m%d%H%M%S)"
  echo "Existing skill found. Moving it to:"
  echo "  $BACKUP"
  mv "$DEST_SKILL" "$BACKUP"
fi

echo "Installing skill to:"
echo "  $DEST_SKILL"
cp -R "$SOURCE_SKILL" "$DEST_SKILL"
chmod +x "$DEST_SKILL/scripts/install_mcp.py" \
  "$DEST_SKILL/scripts/check_pptx_mcp.py" \
  "$DEST_SKILL/scripts/run_pptx_sequence.py" \
  "$DEST_SKILL/scripts/powerpoint_bridge.py" \
  "$DEST_SKILL/scripts/start_bridge.command"

echo
echo "Installing bundled PowerPoint MCP server..."
"$DEST_SKILL/scripts/install_mcp.py" --write-codex-config

echo
echo "Done."
echo
echo "Next steps:"
echo "1. Restart Codex or your Agent app."
echo "2. Open Microsoft PowerPoint."
echo "3. On first use, allow macOS Automation permission to control PowerPoint."
echo "4. For WorkBuddy sandbox mode, start the bridge outside WorkBuddy:"
echo "   $DEST_SKILL/scripts/start_bridge.command"
echo "   Then run:"
echo "   $DEST_SKILL/scripts/install_mcp.py --write-workbuddy-config --bridge-mode"
echo "   Restart WorkBuddy after that config update."
echo "5. For another stdio MCP Agent, copy the JSON block printed above"
echo "   into that product's MCP settings, then restart the Agent."
echo "6. To verify real PowerPoint control after restart, run:"
echo "   $DEST_SKILL/scripts/install_mcp.py --doctor --smoke-powerpoint"
echo "   or, for WorkBuddy bridge mode:"
echo "   $DEST_SKILL/scripts/install_mcp.py --doctor --smoke-powerpoint --bridge-mode"
echo
echo "Try:"
echo "  Use \$mac-powerpoint-live-builder to create a three-slide research deck."
