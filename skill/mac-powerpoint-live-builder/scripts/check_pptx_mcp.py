#!/usr/bin/env python3
"""Check that a local PowerPoint MCP server exposes live-build tools.

Usage:
    python scripts/check_pptx_mcp.py /absolute/path/to/pptx-mcp
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import anyio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


REQUIRED = {
    "pptx_create_presentation",
    "pptx_add_slide",
    "pptx_focus_slide",
    "pptx_add_shape",
    "pptx_add_text_box",
    "pptx_run_live_sequence",
    "pptx_save_presentation",
    "pptx_get_slide_thumbnail",
    "pptx_get_deck_overview",
}


async def main() -> int:
    if len(sys.argv) != 2:
        print("usage: check_pptx_mcp.py /absolute/path/to/pptx-mcp", file=sys.stderr)
        return 2

    server = Path(sys.argv[1]).expanduser().resolve()
    if not server.exists():
        print(f"server executable not found: {server}", file=sys.stderr)
        return 1

    env = dict(os.environ)
    params = StdioServerParameters(command=str(server), env=env)
    async with stdio_client(params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            result = await session.list_tools()
            names = sorted(tool.name for tool in result.tools)

    pptx_names = [name for name in names if name.startswith("pptx_")]
    missing = sorted(REQUIRED - set(names))
    print(f"tool_count={len(names)}")
    print(f"pptx_tool_count={len(pptx_names)}")
    for name in pptx_names:
        print(name)
    if missing:
        print("missing_required=" + ",".join(missing), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    try:
        sys.exit(anyio.run(main))
    except KeyboardInterrupt:
        sys.exit(130)
