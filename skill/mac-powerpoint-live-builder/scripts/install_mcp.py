#!/usr/bin/env python3
"""Install or verify the bundled PowerPoint live MCP server.

This script is intentionally self-contained: it uses only the Python standard
library until the bundled MCP package is installed into its own virtualenv.
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional


SKILL_ROOT = Path(__file__).resolve().parents[1]
VENDOR = SKILL_ROOT / "vendor" / "powerpoint-live-mcp"
DEFAULT_HOME = Path.home() / ".local" / "share" / "powerpoint-live-mcp"
CODEX_CONFIG = Path.home() / ".codex" / "config.toml"
CODEX_RUNTIME = Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" / "dependencies"
MARKER_BEGIN = "# >>> powerpoint-live-mcp managed by mac-powerpoint-live-builder >>>"
MARKER_END = "# <<< powerpoint-live-mcp managed by mac-powerpoint-live-builder <<<"


def run(cmd: list[str], *, check: bool = True) -> subprocess.CompletedProcess[str]:
    print("+ " + " ".join(cmd))
    return subprocess.run(cmd, text=True, capture_output=False, check=check)


def python_candidates() -> list[str]:
    return [
        sys.executable,
        str(CODEX_RUNTIME / "python" / "bin" / "python3"),
        "python3.12",
        "python3.11",
        "python3.10",
        "python3",
    ]


def choose_python() -> str:
    seen: set[str] = set()
    for candidate in python_candidates():
        if candidate in seen:
            continue
        seen.add(candidate)
        found = shutil.which(candidate) if os.sep not in candidate else candidate
        if not found:
            continue
        probe = subprocess.run(
            [found, "-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"],
            text=True,
            capture_output=True,
            check=False,
        )
        if probe.returncode == 0:
            return found
    raise SystemExit("Python 3.10+ was not found. Install Python 3.10 or newer, then rerun this installer.")


def venv_paths(home: Path) -> tuple[Path, Path, Path]:
    venv = home / ".venv"
    if platform.system() == "Windows":
        py = venv / "Scripts" / "python.exe"
        exe = venv / "Scripts" / "powerpoint-live-mcp.exe"
    else:
        py = venv / "bin" / "python"
        exe = venv / "bin" / "powerpoint-live-mcp"
    return venv, py, exe


def check_powerpoint() -> bool:
    return Path("/Applications/Microsoft PowerPoint.app").exists()


def check_pdftoppm() -> Optional[str]:
    for candidate in (
        str(CODEX_RUNTIME / "bin" / "pdftoppm"),
        "/opt/homebrew/bin/pdftoppm",
        "/usr/local/bin/pdftoppm",
    ):
        if Path(candidate).exists():
            return candidate
    return shutil.which("pdftoppm")


def mcp_path_value(exe: Path) -> str:
    return ":".join(
        [
            str(exe.parent),
            str(CODEX_RUNTIME / "bin"),
            "/opt/homebrew/bin",
            "/usr/local/bin",
            "/usr/bin",
            "/bin",
            "/usr/sbin",
            "/sbin",
        ]
    )


def mcp_server_config(exe: Path) -> dict[str, object]:
    return {
        "command": str(exe),
        "env": {
            "PATH": mcp_path_value(exe),
        },
    }


def codex_snippet(exe: Path) -> str:
    path_value = mcp_path_value(exe)
    return f"""{MARKER_BEGIN}
[mcp_servers.powerpoint_live]
command = "{exe}"
startup_timeout_sec = 20
tool_timeout_sec = 240

[mcp_servers.powerpoint_live.env]
PATH = "{path_value}"
{MARKER_END}
"""


def generic_json_snippet(exe: Path) -> str:
    return json.dumps(
        {
            "mcpServers": {
                "powerpoint-live-mcp": mcp_server_config(exe),
            }
        },
        ensure_ascii=False,
        indent=2,
    )


def workbuddy_json_snippet(exe: Path) -> str:
    return json.dumps(
        {
            "powerpoint-live-mcp": mcp_server_config(exe),
        },
        ensure_ascii=False,
        indent=2,
    )


def print_config_snippets(exe: Path) -> None:
    print("\nCodex config.toml snippet:\n")
    print(codex_snippet(exe))
    print("\nGeneric stdio MCP JSON example:\n")
    print(generic_json_snippet(exe))
    print("\nWorkBuddy mcp.json server block example:\n")
    print(workbuddy_json_snippet(exe))
    print("\nNote: many Agent apps only load MCP servers at startup. Restart the Agent after updating config.")


def write_codex_config(exe: Path) -> None:
    CODEX_CONFIG.parent.mkdir(parents=True, exist_ok=True)
    existing = CODEX_CONFIG.read_text(encoding="utf-8") if CODEX_CONFIG.exists() else ""
    snippet = codex_snippet(exe)
    if MARKER_BEGIN in existing and MARKER_END in existing:
        before, rest = existing.split(MARKER_BEGIN, 1)
        _, after = rest.split(MARKER_END, 1)
        new_text = before.rstrip() + "\n\n" + snippet.rstrip() + "\n" + after
    else:
        new_text = existing.rstrip() + "\n\n" + snippet if existing.strip() else snippet
    CODEX_CONFIG.write_text(new_text, encoding="utf-8")
    print(f"Wrote Codex MCP config: {CODEX_CONFIG}")


def verify_tools(py: Path, exe: Path, *, smoke_powerpoint: bool = False) -> int:
    checker = SKILL_ROOT / "scripts" / "check_pptx_mcp.py"
    cmd = [str(py), str(checker), str(exe)]
    if smoke_powerpoint:
        cmd.append("--smoke-powerpoint")
    result = subprocess.run(cmd, text=True, check=False)
    return result.returncode


def doctor(home: Path, *, smoke_powerpoint: bool) -> int:
    print("PowerPoint live MCP doctor")
    print()
    ok = True
    if platform.system() != "Darwin":
        print("FAIL macOS: live PowerPoint control requires macOS.")
        ok = False
    else:
        print("OK macOS")

    if check_powerpoint():
        print("OK Microsoft PowerPoint found")
    else:
        print("FAIL Microsoft PowerPoint not found at /Applications/Microsoft PowerPoint.app")
        ok = False

    pdftoppm = check_pdftoppm()
    if pdftoppm:
        print(f"OK pdftoppm found: {pdftoppm}")
    else:
        print("WARN pdftoppm not found. Install Homebrew poppler for thumbnail export.")

    _, py, exe = venv_paths(home)
    if not py.exists() or not exe.exists():
        print(f"FAIL MCP server is not installed at {home}")
        print("Run this installer without --check/--doctor first.")
        print_config_snippets(exe)
        return 1

    print(f"OK MCP python: {py}")
    print(f"OK MCP command: {exe}")
    print()
    print("Checking MCP tool inventory" + (" and live PowerPoint smoke" if smoke_powerpoint else "") + "...")
    rc = verify_tools(py, exe, smoke_powerpoint=smoke_powerpoint)
    if rc != 0:
        ok = False
        if smoke_powerpoint:
            print()
            print("If the failure mentions -1708, PowerPoint rejected foreground activation.")
            print("If the failure mentions -10004 or not authorized, check macOS Automation permissions.")
    print_config_snippets(exe)
    return 0 if ok else 1


def install(home: Path) -> tuple[Path, Path]:
    if not VENDOR.exists():
        raise SystemExit(f"Bundled MCP package not found: {VENDOR}")
    home.mkdir(parents=True, exist_ok=True)
    venv, py, exe = venv_paths(home)
    base_python = choose_python()
    if not py.exists():
        run([base_python, "-m", "venv", str(venv)])
    run([str(py), "-m", "pip", "install", "mcp>=1.0.0", "Pillow>=10.0"])
    vendor_src = VENDOR / "src"
    if platform.system() == "Windows":
        exe.write_text(
            f'@echo off\r\nset "PYTHONPATH={vendor_src};%PYTHONPATH%"\r\n"{py}" -m pptx_mcp.server %*\r\n',
            encoding="utf-8",
        )
    else:
        exe.write_text(
            f'#!/bin/sh\nPYTHONPATH="{vendor_src}:$PYTHONPATH" exec "{py}" -m pptx_mcp.server "$@"\n',
            encoding="utf-8",
        )
        exe.chmod(0o755)
    if not exe.exists():
        raise SystemExit(f"Install finished but server executable is missing: {exe}")
    return py, exe


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--home", type=Path, default=DEFAULT_HOME, help="Install location for the MCP venv.")
    parser.add_argument("--check", action="store_true", help="Only check the default installation.")
    parser.add_argument("--doctor", action="store_true", help="Run environment diagnostics for the installed MCP server.")
    parser.add_argument(
        "--smoke-powerpoint",
        action="store_true",
        help="With --check or --doctor, create and close a tiny presentation through MCP.",
    )
    parser.add_argument("--write-codex-config", action="store_true", help="Write/update ~/.codex/config.toml.")
    parser.add_argument("--print-config", action="store_true", help="Print Codex, generic MCP JSON, and WorkBuddy config snippets.")
    args = parser.parse_args()

    if platform.system() != "Darwin":
        print("This MCP server controls Microsoft PowerPoint through macOS AppleScript.")
        print("Install can continue for packaging, but live PowerPoint control requires macOS.")

    if not check_powerpoint():
        print("Warning: /Applications/Microsoft PowerPoint.app was not found.")
        print("Install can continue, but live deck generation will not work until PowerPoint for Mac is installed.")

    pdftoppm = check_pdftoppm()
    if pdftoppm:
        print(f"pdftoppm found: {pdftoppm}")
    else:
        print("Warning: pdftoppm was not found. Install poppler for thumbnail/overview export.")
        print("Homebrew command: brew install poppler")

    _, py, exe = venv_paths(args.home)
    if args.doctor:
        return doctor(args.home, smoke_powerpoint=args.smoke_powerpoint)

    if args.check:
        if not exe.exists() or not py.exists():
            print(f"MCP not installed at {args.home}")
            return 1
        return verify_tools(py, exe, smoke_powerpoint=args.smoke_powerpoint)

    py, exe = install(args.home)
    if args.write_codex_config:
        write_codex_config(exe)
    if args.print_config or not args.write_codex_config:
        print_config_snippets(exe)
    print("Verifying MCP tool inventory...")
    rc = verify_tools(py, exe)
    if rc != 0:
        return rc
    print("\nInstallation complete. Restart the Agent app after adding/updating MCP config.")
    print("Use --doctor --smoke-powerpoint to verify real PowerPoint control after restarting/configuring your Agent.")
    print("On first use, allow Automation permission when macOS asks to control Microsoft PowerPoint.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
