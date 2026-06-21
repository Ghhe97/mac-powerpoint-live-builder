#!/usr/bin/env python3
"""Run a PowerPoint live-build JSON sequence through the bundled MCP server.

This helper is useful for Agent products that can run local scripts but do not
directly expose MCP tools in the model runtime. The deck is still built live in
Microsoft PowerPoint via the same `pptx_run_live_sequence` MCP tool.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any


DEFAULT_HOME = Path.home() / ".local" / "share" / "powerpoint-live-mcp"
DEFAULT_SERVER = DEFAULT_HOME / ".venv" / "bin" / "powerpoint-live-mcp"
DEFAULT_PYTHON = DEFAULT_HOME / ".venv" / "bin" / "python"
DEFAULT_BRIDGE_URL = "http://127.0.0.1:18765"
DEFAULT_BRIDGE_TOKEN_FILE = DEFAULT_HOME / "bridge_token"


def ensure_mcp_runtime() -> None:
    try:
        import anyio  # noqa: F401
        import mcp  # noqa: F401
    except ModuleNotFoundError:
        if os.environ.get("POWERPOINT_LIVE_RUNNER_REEXEC") == "1":
            raise
        if DEFAULT_PYTHON.exists():
            env = dict(os.environ)
            env["POWERPOINT_LIVE_RUNNER_REEXEC"] = "1"
            os.execve(str(DEFAULT_PYTHON), [str(DEFAULT_PYTHON), __file__, *sys.argv[1:]], env)
        raise SystemExit(
            "MCP Python dependencies are not available. Run install_mcp.py first, "
            "or invoke this script with ~/.local/share/powerpoint-live-mcp/.venv/bin/python."
        )


def result_text(result: object) -> str:
    parts: list[str] = []
    for item in getattr(result, "content", []) or []:
        text = getattr(item, "text", None)
        parts.append(text if text is not None else repr(item))
    return "\n".join(parts)


def load_steps(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.expanduser().read_text(encoding="utf-8"))
    if isinstance(payload, dict):
        payload = payload.get("steps")
    if not isinstance(payload, list):
        raise SystemExit("Sequence JSON must be a list, or an object with a `steps` list.")
    steps: list[dict[str, Any]] = []
    for index, step in enumerate(payload, 1):
        if not isinstance(step, dict):
            raise SystemExit(f"Step {index} is not an object.")
        steps.append(step)
    return steps


def demo_steps(pptx_path: Path) -> list[dict[str, Any]]:
    pptx = str(pptx_path.expanduser().resolve())
    return [
        {"type": "create_presentation"},
        {"type": "add_slide", "layout": "blank"},
        {"type": "focus_slide", "slide_index": 1},
        {
            "type": "shape",
            "slide_index": 1,
            "shape_name": "demo_bg",
            "shape_type": "rectangle",
            "left": 0,
            "top": 0,
            "width": 960,
            "height": 540,
            "fill_color": "#101827",
            "line_color": "none",
            "line_weight": 0,
            "delay_seconds": 0.2,
        },
        {
            "type": "text",
            "slide_index": 1,
            "shape_name": "demo_title",
            "text": "PowerPoint Live Bridge: WorkBuddy Smoke Test",
            "left": 44,
            "top": 38,
            "width": 820,
            "height": 42,
            "font_size": 25,
            "font_color": "#FFFFFF",
            "bold": True,
            "fill_transparency": 100,
            "line_color": "none",
            "line_weight": 0,
        },
        {
            "type": "text",
            "slide_index": 1,
            "shape_name": "demo_subtitle",
            "text": "This slide was assembled live in Microsoft PowerPoint through MCP + localhost bridge.",
            "left": 46,
            "top": 85,
            "width": 810,
            "height": 30,
            "font_size": 11,
            "font_color": "#9FB0C8",
            "fill_transparency": 100,
            "line_color": "none",
            "line_weight": 0,
        },
        *[
            {
                "type": "shape",
                "slide_index": 1,
                "shape_name": f"demo_card_{i}",
                "shape_type": "rounded_rectangle",
                "left": x,
                "top": 155,
                "width": 260,
                "height": 132,
                "fill_color": color,
                "line_color": "#334155",
                "line_weight": 1,
                "delay_seconds": 0.15,
            }
            for i, (x, color) in enumerate([(44, "#172033"), (350, "#172033"), (656, "#172033")], 1)
        ],
        {
            "type": "text",
            "slide_index": 1,
            "shape_name": "demo_card_1_text",
            "text": "1. Bridge reachable\n127.0.0.1 token-protected HTTP bridge accepted MCP calls.",
            "left": 62,
            "top": 178,
            "width": 224,
            "height": 80,
            "font_size": 12,
            "font_color": "#D9E4F2",
            "bold": False,
            "fill_transparency": 100,
            "line_color": "none",
            "line_weight": 0,
        },
        {
            "type": "text",
            "slide_index": 1,
            "shape_name": "demo_card_2_text",
            "text": "2. PowerPoint controlled\nAppleScript created shapes and text in the visible app window.",
            "left": 368,
            "top": 178,
            "width": 224,
            "height": 80,
            "font_size": 12,
            "font_color": "#D9E4F2",
            "fill_transparency": 100,
            "line_color": "none",
            "line_weight": 0,
        },
        {
            "type": "text",
            "slide_index": 1,
            "shape_name": "demo_card_3_text",
            "text": "3. Editable PPTX saved\nThe output remains normal PowerPoint objects, not a flat image.",
            "left": 674,
            "top": 178,
            "width": 224,
            "height": 80,
            "font_size": 12,
            "font_color": "#D9E4F2",
            "fill_transparency": 100,
            "line_color": "none",
            "line_weight": 0,
        },
        {
            "type": "text",
            "slide_index": 1,
            "shape_name": "demo_footer",
            "text": "Generated by mac-powerpoint-live-builder live sequence runner.",
            "left": 46,
            "top": 500,
            "width": 840,
            "height": 18,
            "font_size": 8,
            "font_color": "#7C8BA1",
            "fill_transparency": 100,
            "line_color": "none",
            "line_weight": 0,
        },
        {"type": "save", "save_as_path": pptx},
    ]


def bridge_token(path: Path) -> str:
    try:
        return path.expanduser().read_text(encoding="utf-8").strip()
    except OSError:
        return ""


def run_delegated(args: argparse.Namespace) -> int:
    if args.demo_pptx:
        payload: dict[str, Any] = {"demo_pptx": str(args.demo_pptx.expanduser())}
    elif args.sequence_json:
        payload = {"steps": load_steps(args.sequence_json)}
    else:
        raise SystemExit("Provide a sequence JSON path or --demo-pptx.")

    payload["default_delay_seconds"] = args.default_delay_seconds
    payload["timeout"] = args.bridge_timeout
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    token = bridge_token(args.bridge_token_file)
    url = args.bridge_url.rstrip("/") + "/run-live-sequence"
    last_error = ""

    for attempt in range(1, 4):
        request = urllib.request.Request(
            url,
            data=body,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {token}",
                "Connection": "close",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=args.bridge_timeout + 10) as response:
                response_body = response.read().decode("utf-8")
            result = json.loads(response_body)
            stdout = str(result.get("stdout") or result.get("stdout_tail") or "")
            stderr = str(result.get("stderr") or result.get("stderr_tail") or "")
            if result.get("ok"):
                print("delegated_live_sequence ok")
                if result.get("output_path"):
                    print(f"output_path={result['output_path']}")
                print(stdout, end="" if stdout.endswith("\n") else "\n")
                if args.result_json:
                    args.result_json.expanduser().write_text(stdout + "\n", encoding="utf-8")
                return 0
            print("delegated_live_sequence failed:", file=sys.stderr)
            if stdout:
                print(stdout, file=sys.stderr)
            if stderr:
                print(stderr, file=sys.stderr)
            return int(result.get("returncode") or 1)
        except urllib.error.HTTPError as e:
            last_error = e.read().decode("utf-8", errors="replace")
            retryable_empty_body = (
                e.code == 400
                and ("empty JSON body" in last_error or '"body_preview": ""' in last_error)
                and attempt < 3
            )
            if retryable_empty_body:
                time.sleep(0.35 * attempt)
                continue
            print(f"delegated_live_sequence HTTP {e.code}: {last_error}", file=sys.stderr)
            return 1
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as e:
            last_error = str(e)
            if attempt < 3:
                time.sleep(0.35 * attempt)
                continue
            break

    print(f"delegated_live_sequence failed: {last_error}", file=sys.stderr)
    return 1


async def run_sequence(args: argparse.Namespace) -> int:
    from mcp import ClientSession, StdioServerParameters
    from mcp.client.stdio import stdio_client

    server = args.server.expanduser().resolve()
    if not server.exists():
        print(f"server executable not found: {server}", file=sys.stderr)
        return 1

    if args.demo_pptx:
        steps = demo_steps(args.demo_pptx)
    elif args.sequence_json:
        steps = load_steps(args.sequence_json)
    else:
        raise SystemExit("Provide a sequence JSON path or --demo-pptx.")

    env = dict(os.environ)
    if args.bridge_mode:
        env["POWERPOINT_LIVE_BRIDGE_URL"] = args.bridge_url
        env["POWERPOINT_LIVE_BRIDGE_TOKEN_FILE"] = str(args.bridge_token_file.expanduser())

    params = StdioServerParameters(command=str(server), env=env)
    async with stdio_client(params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            result = await session.call_tool(
                "pptx_run_live_sequence",
                {
                    "steps": steps,
                    "default_delay_seconds": args.default_delay_seconds,
                },
            )
            text = result_text(result)
            if getattr(result, "isError", False):
                print("pptx_run_live_sequence failed:", file=sys.stderr)
                print(text, file=sys.stderr)
                return 1
            print("pptx_run_live_sequence ok")
            print(text)
            if args.result_json:
                args.result_json.expanduser().write_text(text + "\n", encoding="utf-8")

    return 0


def main() -> int:
    ensure_mcp_runtime()
    import anyio

    parser = argparse.ArgumentParser()
    parser.add_argument("sequence_json", nargs="?", type=Path, help="JSON file containing live sequence steps.")
    parser.add_argument("--server", type=Path, default=DEFAULT_SERVER, help="MCP server executable.")
    parser.add_argument("--bridge-mode", action="store_true", help="Set bridge env for the MCP server.")
    parser.add_argument(
        "--delegate-to-bridge",
        action="store_true",
        help="Ask the localhost bridge to run the sequence outside the current Agent sandbox.",
    )
    parser.add_argument("--bridge-url", default=DEFAULT_BRIDGE_URL, help="PowerPoint bridge URL.")
    parser.add_argument("--bridge-token-file", type=Path, default=DEFAULT_BRIDGE_TOKEN_FILE, help="Bridge token file.")
    parser.add_argument("--bridge-timeout", type=int, default=600, help="Timeout for delegated bridge sequence runs.")
    parser.add_argument("--default-delay-seconds", type=float, default=0.35, help="Visible delay between live steps.")
    parser.add_argument("--demo-pptx", type=Path, help="Build and save a one-slide live smoke deck to this path.")
    parser.add_argument("--result-json", type=Path, help="Optional file path to write the MCP JSON result.")
    args = parser.parse_args()
    if args.delegate_to_bridge:
        return run_delegated(args)
    return anyio.run(run_sequence, args)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
