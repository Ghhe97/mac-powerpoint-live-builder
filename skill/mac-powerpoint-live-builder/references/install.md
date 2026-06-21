# Installation and Self-Check

The skill includes a bundled local MCP server under `vendor/powerpoint-live-mcp`.
When the user asks to build a live PowerPoint deck, first ensure the MCP server is
available.

## Agent Procedure

1. Check whether live `pptx_*` MCP tools are already available in the current
   runtime. If they are, proceed to deck generation.
2. If tools are not available, check whether the MCP server is installed and can
   list tools:

   ```bash
   python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --check
   ```

3. Remember that this self-check only proves the server can start. It does not
   prove the current Agent session has mounted `pptx_*` tools. If the Agent only
   loads MCP at startup, the user must update MCP settings and restart the Agent.
4. If the check fails, ask for permission to run the installer when the runtime
   requires approval for writing outside the workspace or downloading packages.
5. Install from the bundled vendor package:

   ```bash
   python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-codex-config
   ```

6. Tell the user to restart their Agent app if the app only loads MCP servers at
   startup.
7. On first live use, macOS may ask whether the launching app may control Microsoft
   PowerPoint. The user must allow it.
8. If the Agent is WorkBuddy or another sandboxed runtime, and direct smoke tests
   fail with `-10004` even after Automation permission is enabled, start the
   bridge outside the Agent sandbox:

   ```bash
   ~/.codex/skills/mac-powerpoint-live-builder/scripts/start_bridge.command
   ```

   The bridge startup script runs a short PowerPoint Automation self-test. If it
   fails, the bridge still starts, but the terminal output explains which launcher
   permission must be enabled.

   Then write WorkBuddy config in bridge mode and restart WorkBuddy:

   ```bash
   ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-workbuddy-config --bridge-mode
   ```

   The app that starts the bridge also needs Automation permission. If the bridge
   is launched by `start_bridge.command`, macOS treats Terminal as the launcher:
   enable System Settings > Privacy & Security > Automation > Terminal >
   Microsoft PowerPoint.

9. For real PowerPoint control diagnostics, run:

   ```bash
   python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint
   ```

   In WorkBuddy bridge mode, run:

   ```bash
   python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint --bridge-mode
   ```

   This creates and closes a tiny presentation through MCP. Use it only when opening
   PowerPoint and triggering Automation prompts is acceptable.

## What the Installer Does

- Creates a virtualenv at `~/.local/share/powerpoint-live-mcp/.venv`.
- Installs Python runtime dependencies into that virtualenv and creates a wrapper
  that runs the bundled MCP server from `vendor/powerpoint-live-mcp`.
- Checks for Microsoft PowerPoint.
- Checks for `pdftoppm`; if missing, tells the user to install Homebrew `poppler`.
- Optionally writes a managed Codex MCP block to `~/.codex/config.toml`.
- Optionally writes or updates WorkBuddy MCP config with
  `--write-workbuddy-config`.
- Prints Codex TOML, generic stdio MCP JSON, and a WorkBuddy-style server block.
- Verifies the installed server exposes the expected `pptx_*` tools.
- With `--smoke-powerpoint`, verifies the MCP server can actually control
  PowerPoint, not just list tools.
- In `--bridge-mode`, configures the MCP server to proxy AppleScript through a
  localhost bridge process started outside the Agent sandbox.
- Provides `scripts/run_pptx_sequence.py`, a CLI helper for Agent products that
  can run local scripts but do not mount `pptx_*` tools directly.

## Other Agent Products

For non-Codex Agent clients, run the installer without `--write-codex-config` and
copy the printed stdio command or JSON block into that product's MCP settings:

```text
command: ~/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp
transport: stdio
```

If the product supports environment variables, include Homebrew paths in `PATH` so
`pdftoppm` can be found:

```text
PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin
```

For WorkBuddy-style `mcp.json`, use the printed server block shape:

```json
"powerpoint-live-mcp": {
  "command": "/Users/YOU/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp",
  "env": {
    "PATH": "/Users/YOU/.local/share/powerpoint-live-mcp/.venv/bin:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
  }
}
```

If direct WorkBuddy control fails with `-10004` after Automation permission is
enabled, use bridge mode. The bridge is a local HTTP server bound to
`127.0.0.1` and protected by a token file. It runs AppleScript outside the
WorkBuddy sandbox, while WorkBuddy's MCP process only talks to localhost.

Start the bridge:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/start_bridge.command
```

Write WorkBuddy config:

```bash
~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-workbuddy-config --bridge-mode
```

The resulting server env includes:

```text
POWERPOINT_LIVE_BRIDGE_URL=http://127.0.0.1:18765
POWERPOINT_LIVE_BRIDGE_TOKEN_FILE=~/.local/share/powerpoint-live-mcp/bridge_token
```

If the Agent cannot call MCP tools directly but can run scripts, use the sequence
runner. It still calls `pptx_run_live_sequence` through MCP:

```bash
~/.workbuddy/skills/mac-powerpoint-live-builder/scripts/run_pptx_sequence.py --bridge-mode sequence.json
```

Quick end-to-end smoke deck:

```bash
~/.workbuddy/skills/mac-powerpoint-live-builder/scripts/run_pptx_sequence.py --bridge-mode --demo-pptx ~/Desktop/powerpoint-live-smoke.pptx
```

For WorkBuddy-style command runners, keep diagnostics and smoke tests as one-line
commands. Avoid asking the Agent to generate multi-line heredocs or inline Python;
its shell wrapper may quote them incorrectly. For real decks, save the sequence as
JSON, then pass the JSON path to `run_pptx_sequence.py`.

If WorkBuddy's sandbox proxies localhost HTTP unreliably, prefer delegated bridge
runner mode:

```bash
~/.workbuddy/skills/mac-powerpoint-live-builder/scripts/run_pptx_sequence.py --delegate-to-bridge --demo-pptx ~/Desktop/powerpoint-live-smoke.pptx
```

Delegated mode sends one token-protected request to the bridge; the bridge runs
the full live sequence outside the WorkBuddy sandbox.

## Troubleshooting

- `install_mcp.py --check` passes but the Agent cannot call `pptx_*`: the server is
  installed, but the active Agent session has not mounted it. Add the MCP config
  and restart the Agent.
- Bridge `HTTP 400` on `/run-osascript`: the bridge received a request, but the
  request was not valid JSON or missed the required `script` field. Use the
  one-line runner/doctor commands instead of ad-hoc inline scripts, and inspect
  the bridge response for `content_length` and `body_preview`.
- WorkBuddy runner fails at random sequence steps with `timed out`: its localhost
  proxy may be dropping POST bodies. Use `run_pptx_sequence.py
  --delegate-to-bridge` so only one request crosses the WorkBuddy sandbox.
- `-1708` or `"activate" can't continue`: PowerPoint rejected foreground activation.
  Use the updated server, which wraps `activate` in `try/end try`.
- `-10004`, `not authorized`, or Automation permission errors: open macOS System
  Settings > Privacy & Security > Automation and allow the app that launched
  AppleScript to control Microsoft PowerPoint. If that checkbox is already enabled
  for the Agent but smoke still fails, the Agent is likely running MCP in a
  restricted sandbox; use bridge mode.
- `osascript timed out after 60s` in bridge mode: the bridge is reachable, but
  the app that launched the bridge may not be authorized for PowerPoint
  Automation. For `start_bridge.command`, enable Terminal > Microsoft PowerPoint,
  restart the bridge, then rerun the smoke test.
- If live control fails and the Agent uses `python-pptx` or another file-only
  method, label the result as a non-live fallback.
