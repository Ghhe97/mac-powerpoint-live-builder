#!/usr/bin/env python3
"""Check that a local PowerPoint MCP server exposes live-build tools.

Usage:
    python scripts/check_pptx_mcp.py /absolute/path/to/pptx-mcp
    python scripts/check_pptx_mcp.py /absolute/path/to/pptx-mcp --smoke-powerpoint
"""

from __future__ import annotations

import argparse
import json
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


def _result_text(result: object) -> str:
    parts: list[str] = []
    for item in getattr(result, "content", []) or []:
        text = getattr(item, "text", None)
        if text is not None:
            parts.append(text)
        else:
            parts.append(repr(item))
    return "\n".join(parts)


async def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("server", type=Path, help="Absolute path to the MCP server executable.")
    parser.add_argument(
        "--smoke-powerpoint",
        action="store_true",
        help="Also call pptx_create_presentation through MCP, then close it without saving.",
    )
    args = parser.parse_args()

    server = args.server.expanduser().resolve()
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

            if args.smoke_powerpoint:
                print("smoke_powerpoint=starting")
                create_result = await session.call_tool("pptx_create_presentation", {})
                if getattr(create_result, "isError", False):
                    print("pptx_create_presentation failed:", file=sys.stderr)
                    print(_result_text(create_result), file=sys.stderr)
                    return 1
                create_text = _result_text(create_result)
                print("pptx_create_presentation ok")
                print(create_text)
                presentation_name: str | None = None
                try:
                    payload = json.loads(create_text)
                    if isinstance(payload, dict):
                        presentation_name = payload.get("name")
                except json.JSONDecodeError:
                    presentation_name = None
                close_args = {"save_changes": False}
                if presentation_name:
                    close_args["presentation_name"] = presentation_name
                close_result = await session.call_tool("pptx_close_presentation", close_args)
                if getattr(close_result, "isError", False):
                    print("pptx_close_presentation failed after create smoke:", file=sys.stderr)
                    print(_result_text(close_result), file=sys.stderr)
                    return 1
                print("pptx_close_presentation ok")
                print(_result_text(close_result))

    return 0


if __name__ == "__main__":
    try:
        sys.exit(anyio.run(main))
    except KeyboardInterrupt:
        sys.exit(130)
