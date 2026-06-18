# Installation and Self-Check

The skill includes a bundled local MCP server under `vendor/powerpoint-live-mcp`.
When the user asks to build a live PowerPoint deck, first ensure the MCP server is
available.

## Agent Procedure

1. Check whether live `pptx_*` MCP tools are already available in the current
   runtime. If they are, proceed to deck generation.
2. If tools are not available, run:

   ```bash
   python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --check
   ```

3. If the check fails, ask for permission to run the installer when the runtime
   requires approval for writing outside the workspace or downloading packages.
4. Install from the bundled vendor package:

   ```bash
   python ~/.codex/skills/mac-powerpoint-live-builder/scripts/install_mcp.py --write-codex-config
   ```

5. Tell the user to restart their Agent app if the app only loads MCP servers at
   startup.
6. On first live use, macOS may ask whether the launching app may control Microsoft
   PowerPoint. The user must allow it.

## What the Installer Does

- Creates a virtualenv at `~/.local/share/powerpoint-live-mcp/.venv`.
- Installs Python runtime dependencies into that virtualenv and creates a wrapper
  that runs the bundled MCP server from `vendor/powerpoint-live-mcp`.
- Checks for Microsoft PowerPoint.
- Checks for `pdftoppm`; if missing, tells the user to install Homebrew `poppler`.
- Optionally writes a managed Codex MCP block to `~/.codex/config.toml`.
- Verifies the installed server exposes the expected `pptx_*` tools.

## Other Agent Products

For non-Codex Agent clients, run the installer without `--write-codex-config` and
copy the printed stdio command into that product's MCP settings:

```text
command: ~/.local/share/powerpoint-live-mcp/.venv/bin/powerpoint-live-mcp
transport: stdio
```

If the product supports environment variables, include Homebrew paths in `PATH` so
`pdftoppm` can be found:

```text
PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin
```
