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
8. For real PowerPoint control diagnostics, run:

   ```bash
   python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --doctor --smoke-powerpoint
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
- Prints Codex TOML, generic stdio MCP JSON, and a WorkBuddy-style server block.
- Verifies the installed server exposes the expected `pptx_*` tools.
- With `--smoke-powerpoint`, verifies the MCP server can actually control
  PowerPoint, not just list tools.

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

## Troubleshooting

- `install_mcp.py --check` passes but the Agent cannot call `pptx_*`: the server is
  installed, but the active Agent session has not mounted it. Add the MCP config
  and restart the Agent.
- `-1708` or `"activate" can't continue`: PowerPoint rejected foreground activation.
  Use the updated server, which wraps `activate` in `try/end try`.
- `-10004`, `not authorized`, or Automation permission errors: open macOS System
  Settings > Privacy & Security > Automation and allow the app that launched the
  MCP server to control Microsoft PowerPoint.
- If live control fails and the Agent uses `python-pptx` or another file-only
  method, label the result as a non-live fallback.
