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
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Optional


SKILL_ROOT = Path(__file__).resolve().parents[1]
VENDOR = SKILL_ROOT / "vendor" / "powerpoint-live-mcp"
DEFAULT_HOME = Path.home() / ".local" / "share" / "powerpoint-live-mcp"
DEFAULT_BRIDGE_URL = "http://127.0.0.1:18765"
DEFAULT_BRIDGE_TOKEN_FILE = DEFAULT_HOME / "bridge_token"
CODEX_CONFIG = Path.home() / ".codex" / "config.toml"
WORKBUDDY_CONFIG = Path.home() / ".workbuddy" / "mcp.json"
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


def bridge_env(bridge_url: str, bridge_token_file: Path) -> dict[str, str]:
    return {
        "POWERPOINT_LIVE_BRIDGE_URL": bridge_url,
        "POWERPOINT_LIVE_BRIDGE_TOKEN_FILE": str(bridge_token_file.expanduser()),
    }


def mcp_server_config(
    exe: Path,
    *,
    bridge_mode: bool = False,
    bridge_url: str = DEFAULT_BRIDGE_URL,
    bridge_token_file: Path = DEFAULT_BRIDGE_TOKEN_FILE,
) -> dict[str, object]:
    env = {
        "PATH": mcp_path_value(exe),
    }
    if bridge_mode:
        env.update(bridge_env(bridge_url, bridge_token_file))
    return {
        "command": str(exe),
        "env": env,
    }


def codex_snippet(
    exe: Path,
    *,
    bridge_mode: bool = False,
    bridge_url: str = DEFAULT_BRIDGE_URL,
    bridge_token_file: Path = DEFAULT_BRIDGE_TOKEN_FILE,
) -> str:
    path_value = mcp_path_value(exe)
    bridge_lines = ""
    if bridge_mode:
        bridge_lines = (
            f'POWERPOINT_LIVE_BRIDGE_URL = "{bridge_url}"\n'
            f'POWERPOINT_LIVE_BRIDGE_TOKEN_FILE = "{bridge_token_file.expanduser()}"\n'
        )
    return f"""{MARKER_BEGIN}
[mcp_servers.powerpoint_live]
command = "{exe}"
startup_timeout_sec = 20
tool_timeout_sec = 240

[mcp_servers.powerpoint_live.env]
PATH = "{path_value}"
{bridge_lines.rstrip()}
{MARKER_END}
"""


def generic_json_snippet(
    exe: Path,
    *,
    bridge_mode: bool = False,
    bridge_url: str = DEFAULT_BRIDGE_URL,
    bridge_token_file: Path = DEFAULT_BRIDGE_TOKEN_FILE,
) -> str:
    return json.dumps(
        {
            "mcpServers": {
                "powerpoint-live-mcp": mcp_server_config(
                    exe,
                    bridge_mode=bridge_mode,
                    bridge_url=bridge_url,
                    bridge_token_file=bridge_token_file,
                ),
            }
        },
        ensure_ascii=False,
        indent=2,
    )


def workbuddy_json_snippet(
    exe: Path,
    *,
    bridge_mode: bool = False,
    bridge_url: str = DEFAULT_BRIDGE_URL,
    bridge_token_file: Path = DEFAULT_BRIDGE_TOKEN_FILE,
) -> str:
    return json.dumps(
        {
            "powerpoint-live-mcp": mcp_server_config(
                exe,
                bridge_mode=bridge_mode,
                bridge_url=bridge_url,
                bridge_token_file=bridge_token_file,
            ),
        },
        ensure_ascii=False,
        indent=2,
    )


def print_config_snippets(
    exe: Path,
    *,
    bridge_mode: bool = False,
    bridge_url: str = DEFAULT_BRIDGE_URL,
    bridge_token_file: Path = DEFAULT_BRIDGE_TOKEN_FILE,
) -> None:
    print("\nCodex config.toml snippet:\n")
    print(
        codex_snippet(
            exe,
            bridge_mode=bridge_mode,
            bridge_url=bridge_url,
            bridge_token_file=bridge_token_file,
        )
    )
    title = "Generic stdio MCP JSON example"
    wb_title = "WorkBuddy mcp.json server block example"
    if bridge_mode:
        title += " (bridge mode)"
        wb_title += " (bridge mode)"
    print(f"\n{title}:\n")
    print(
        generic_json_snippet(
            exe,
            bridge_mode=bridge_mode,
            bridge_url=bridge_url,
            bridge_token_file=bridge_token_file,
        )
    )
    print(f"\n{wb_title}:\n")
    print(
        workbuddy_json_snippet(
            exe,
            bridge_mode=bridge_mode,
            bridge_url=bridge_url,
            bridge_token_file=bridge_token_file,
        )
    )
    if bridge_mode:
        print("\nBridge mode requires a bridge process running outside the Agent sandbox:")
        print(f"  python {SKILL_ROOT / 'scripts' / 'powerpoint_bridge.py'}")
        print(f"  bridge_url={bridge_url}")
        print(f"  bridge_token_file={bridge_token_file.expanduser()}")
    print("\nNote: many Agent apps only load MCP servers at startup. Restart the Agent after updating config.")


def write_codex_config(
    exe: Path,
    *,
    bridge_mode: bool = False,
    bridge_url: str = DEFAULT_BRIDGE_URL,
    bridge_token_file: Path = DEFAULT_BRIDGE_TOKEN_FILE,
) -> None:
    CODEX_CONFIG.parent.mkdir(parents=True, exist_ok=True)
    existing = CODEX_CONFIG.read_text(encoding="utf-8") if CODEX_CONFIG.exists() else ""
    snippet = codex_snippet(
        exe,
        bridge_mode=bridge_mode,
        bridge_url=bridge_url,
        bridge_token_file=bridge_token_file,
    )
    if MARKER_BEGIN in existing and MARKER_END in existing:
        before, rest = existing.split(MARKER_BEGIN, 1)
        _, after = rest.split(MARKER_END, 1)
        new_text = before.rstrip() + "\n\n" + snippet.rstrip() + "\n" + after
    else:
        new_text = existing.rstrip() + "\n\n" + snippet if existing.strip() else snippet
    CODEX_CONFIG.write_text(new_text, encoding="utf-8")
    print(f"Wrote Codex MCP config: {CODEX_CONFIG}")


def _load_json_config(path: Path) -> dict[str, object]:
    if not path.exists():
        return {"mcpServers": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise SystemExit(f"Could not parse JSON config {path}: {e}") from e
    if not isinstance(data, dict):
        raise SystemExit(f"JSON config must contain an object: {path}")
    return data


def _server_map(data: dict[str, object]) -> dict[str, object]:
    existing = data.get("mcpServers")
    if isinstance(existing, dict):
        return existing
    if "mcpServers" in data:
        raise SystemExit("mcpServers exists but is not an object; please fix the MCP config manually.")
    values = list(data.values())
    if values and all(isinstance(v, dict) for v in values):
        # Some clients accept the server map as the top-level JSON object.
        return data
    data["mcpServers"] = {}
    return data["mcpServers"]  # type: ignore[return-value]


def write_workbuddy_config(
    exe: Path,
    *,
    config_path: Path = WORKBUDDY_CONFIG,
    bridge_mode: bool = True,
    bridge_url: str = DEFAULT_BRIDGE_URL,
    bridge_token_file: Path = DEFAULT_BRIDGE_TOKEN_FILE,
) -> None:
    config_path = config_path.expanduser()
    config_path.parent.mkdir(parents=True, exist_ok=True)
    data = _load_json_config(config_path)
    servers = _server_map(data)
    new_config = mcp_server_config(
        exe,
        bridge_mode=bridge_mode,
        bridge_url=bridge_url,
        bridge_token_file=bridge_token_file,
    )
    previous = servers.get("powerpoint-live-mcp")
    if isinstance(previous, dict):
        merged = dict(previous)
        old_env = previous.get("env")
        new_env = new_config.get("env")
        if isinstance(old_env, dict) and isinstance(new_env, dict):
            env = dict(old_env)
            env.update(new_env)
            new_config["env"] = env
        merged.update(new_config)
        servers["powerpoint-live-mcp"] = merged
    else:
        servers["powerpoint-live-mcp"] = new_config
    if config_path.exists():
        backup = config_path.with_suffix(config_path.suffix + f".backup.{time.strftime('%Y%m%d%H%M%S')}")
        shutil.copy2(config_path, backup)
        print(f"Backed up WorkBuddy MCP config: {backup}")
    config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote WorkBuddy MCP config: {config_path}")
    if bridge_mode:
        print("WorkBuddy bridge mode is enabled for powerpoint-live-mcp.")


def verify_tools(
    py: Path,
    exe: Path,
    *,
    smoke_powerpoint: bool = False,
    extra_env: Optional[dict[str, str]] = None,
) -> int:
    checker = SKILL_ROOT / "scripts" / "check_pptx_mcp.py"
    cmd = [str(py), str(checker), str(exe)]
    if smoke_powerpoint:
        cmd.append("--smoke-powerpoint")
    env = dict(os.environ)
    if extra_env:
        env.update(extra_env)
    result = subprocess.run(cmd, text=True, check=False, env=env)
    return result.returncode


def _bridge_token(bridge_token_file: Path) -> str:
    try:
        return bridge_token_file.expanduser().read_text(encoding="utf-8").strip()
    except OSError:
        return ""


def bridge_get_json(bridge_url: str, bridge_token_file: Path, path: str, *, timeout: int = 10) -> dict[str, object]:
    token = _bridge_token(bridge_token_file)
    request = urllib.request.Request(
        bridge_url.rstrip("/") + path,
        headers={"Authorization": f"Bearer {token}"},
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        body = response.read().decode("utf-8")
    data = json.loads(body)
    if not isinstance(data, dict):
        raise RuntimeError(f"Bridge returned non-object JSON from {path}: {body[:200]}")
    return data


def check_bridge_preflight(bridge_url: str, bridge_token_file: Path, *, self_test: bool) -> bool:
    ok = True
    print("Checking PowerPoint bridge...")
    if not bridge_token_file.expanduser().exists():
        print(f"FAIL bridge token file missing: {bridge_token_file.expanduser()}")
        return False
    try:
        health = bridge_get_json(bridge_url, bridge_token_file, "/health", timeout=5)
    except urllib.error.HTTPError as e:
        print(f"FAIL bridge health HTTP {e.code}. Check bridge token/config.")
        return False
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        print(f"FAIL bridge is not reachable at {bridge_url}: {e}")
        print("Start the bridge outside the Agent sandbox, then rerun doctor.")
        return False
    except Exception as e:
        print(f"FAIL bridge health check failed: {e}")
        return False
    if health.get("ok"):
        print(f"OK bridge reachable: {bridge_url}")
    else:
        print(f"FAIL bridge health returned: {health}")
        return False

    if self_test:
        try:
            result = bridge_get_json(bridge_url, bridge_token_file, "/self-test", timeout=12)
        except Exception as e:
            print(f"FAIL bridge PowerPoint Automation self-test could not run: {e}")
            return False
        if result.get("ok"):
            print("OK bridge PowerPoint Automation self-test")
        else:
            ok = False
            detail = str(result.get("stderr") or result.get("stdout") or result.get("message") or "").strip()
            print("FAIL bridge PowerPoint Automation self-test")
            if detail:
                print(detail)
            print("If the bridge was launched by Terminal, enable:")
            print("System Settings > Privacy & Security > Automation > Terminal > Microsoft PowerPoint")
            print("Then restart the bridge and rerun doctor.")
    return ok


def doctor(
    home: Path,
    *,
    smoke_powerpoint: bool,
    bridge_mode: bool,
    bridge_url: str,
    bridge_token_file: Path,
) -> int:
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
    mode = "bridge" if bridge_mode else "direct"
    print(f"Checking MCP tool inventory in {mode} mode" + (" and live PowerPoint smoke" if smoke_powerpoint else "") + "...")
    extra_env = bridge_env(bridge_url, bridge_token_file) if bridge_mode else None
    if bridge_mode:
        bridge_ok = check_bridge_preflight(bridge_url, bridge_token_file, self_test=smoke_powerpoint)
        if not bridge_ok and smoke_powerpoint:
            ok = False
            print()
            print("Skipping MCP live smoke because bridge preflight failed.")
            print_config_snippets(
                exe,
                bridge_mode=bridge_mode,
                bridge_url=bridge_url,
                bridge_token_file=bridge_token_file,
            )
            return 1
    rc = verify_tools(py, exe, smoke_powerpoint=smoke_powerpoint, extra_env=extra_env)
    if rc != 0:
        ok = False
        if smoke_powerpoint:
            print()
            print("If the failure mentions -1708, PowerPoint rejected foreground activation.")
            print("If the failure mentions -10004 or not authorized, check macOS Automation permissions.")
    print_config_snippets(
        exe,
        bridge_mode=bridge_mode,
        bridge_url=bridge_url,
        bridge_token_file=bridge_token_file,
    )
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
    parser.add_argument(
        "--bridge-mode",
        action="store_true",
        help="Print/use MCP config that proxies AppleScript through the localhost PowerPoint bridge.",
    )
    parser.add_argument("--bridge-url", default=DEFAULT_BRIDGE_URL, help="PowerPoint bridge URL.")
    parser.add_argument(
        "--bridge-token-file",
        type=Path,
        default=DEFAULT_BRIDGE_TOKEN_FILE,
        help="PowerPoint bridge token file.",
    )
    parser.add_argument("--write-codex-config", action="store_true", help="Write/update ~/.codex/config.toml.")
    parser.add_argument(
        "--write-workbuddy-config",
        action="store_true",
        help="Write/update WorkBuddy mcp.json. Use --bridge-mode for sandbox-friendly PowerPoint control.",
    )
    parser.add_argument(
        "--workbuddy-config",
        type=Path,
        default=WORKBUDDY_CONFIG,
        help="Path to WorkBuddy mcp.json.",
    )
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
    if (
        args.print_config
        and not args.write_codex_config
        and not args.write_workbuddy_config
        and not args.check
        and not args.doctor
    ):
        print_config_snippets(
            exe,
            bridge_mode=args.bridge_mode,
            bridge_url=args.bridge_url,
            bridge_token_file=args.bridge_token_file,
        )
        return 0

    if args.doctor:
        return doctor(
            args.home,
            smoke_powerpoint=args.smoke_powerpoint,
            bridge_mode=args.bridge_mode,
            bridge_url=args.bridge_url,
            bridge_token_file=args.bridge_token_file,
        )

    if args.check:
        if not exe.exists() or not py.exists():
            print(f"MCP not installed at {args.home}")
            return 1
        extra_env = bridge_env(args.bridge_url, args.bridge_token_file) if args.bridge_mode else None
        return verify_tools(py, exe, smoke_powerpoint=args.smoke_powerpoint, extra_env=extra_env)

    py, exe = install(args.home)
    if args.write_codex_config:
        write_codex_config(
            exe,
            bridge_mode=args.bridge_mode,
            bridge_url=args.bridge_url,
            bridge_token_file=args.bridge_token_file,
        )
    if args.write_workbuddy_config:
        write_workbuddy_config(
            exe,
            config_path=args.workbuddy_config,
            bridge_mode=args.bridge_mode,
            bridge_url=args.bridge_url,
            bridge_token_file=args.bridge_token_file,
        )
    if args.print_config or not args.write_codex_config or not args.write_workbuddy_config:
        print_config_snippets(
            exe,
            bridge_mode=args.bridge_mode,
            bridge_url=args.bridge_url,
            bridge_token_file=args.bridge_token_file,
        )
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
