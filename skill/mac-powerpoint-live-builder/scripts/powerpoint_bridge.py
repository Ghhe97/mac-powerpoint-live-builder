#!/usr/bin/env python3
"""Local localhost bridge for PowerPoint AppleScript control.

Run this outside restrictive Agent sandboxes. The MCP server can then set
POWERPOINT_LIVE_BRIDGE_URL and POWERPOINT_LIVE_BRIDGE_TOKEN_FILE to proxy
AppleScript execution through this process.
"""

from __future__ import annotations

import argparse
import json
import secrets
import subprocess
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any


DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 18765
DEFAULT_HOME = Path.home() / ".local" / "share" / "powerpoint-live-mcp"
DEFAULT_TOKEN_FILE = DEFAULT_HOME / "bridge_token"


def ensure_token(path: Path) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        token = path.read_text(encoding="utf-8").strip()
        if token:
            return token
    token = secrets.token_urlsafe(32)
    path.write_text(token + "\n", encoding="utf-8")
    path.chmod(0o600)
    return token


def run_osascript(script: str, timeout: int) -> dict[str, Any]:
    try:
        result = subprocess.run(
            ["/usr/bin/osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return {
            "ok": False,
            "returncode": 124,
            "stdout": "",
            "stderr": f"osascript timed out after {timeout}s",
        }
    return {
        "ok": result.returncode == 0,
        "returncode": result.returncode,
        "stdout": result.stdout,
        "stderr": result.stderr,
    }


class BridgeHandler(BaseHTTPRequestHandler):
    server_version = "PowerPointLiveBridge/0.1"

    def _send_json(self, status: int, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _authorized(self) -> bool:
        expected = getattr(self.server, "bridge_token", "")
        if not expected:
            return True
        auth = self.headers.get("Authorization", "")
        bearer = "Bearer "
        return auth.startswith(bearer) and auth[len(bearer):].strip() == expected

    def do_GET(self) -> None:
        if self.path != "/health":
            self._send_json(404, {"ok": False, "error": "not found"})
            return
        if not self._authorized():
            self._send_json(401, {"ok": False, "error": "unauthorized"})
            return
        self._send_json(200, {"ok": True, "service": "powerpoint-live-bridge"})

    def do_POST(self) -> None:
        if self.path != "/run-osascript":
            self._send_json(404, {"ok": False, "error": "not found"})
            return
        if not self._authorized():
            self._send_json(401, {"ok": False, "error": "unauthorized"})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            script = str(payload["script"])
            timeout = max(1, min(int(payload.get("timeout", 60)), 600))
        except Exception as e:
            self._send_json(400, {"ok": False, "error": f"bad request: {e}"})
            return
        self._send_json(200, run_osascript(script, timeout))

    def log_message(self, format: str, *args: Any) -> None:
        print(f"{self.address_string()} - {format % args}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default=DEFAULT_HOST, help="Bind host. Defaults to 127.0.0.1.")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help="Bind port.")
    parser.add_argument("--token-file", type=Path, default=DEFAULT_TOKEN_FILE, help="Bridge token file.")
    args = parser.parse_args()

    token = ensure_token(args.token_file.expanduser())
    server = ThreadingHTTPServer((args.host, args.port), BridgeHandler)
    server.bridge_token = token  # type: ignore[attr-defined]
    print("PowerPoint live bridge started")
    print(f"url=http://{args.host}:{args.port}")
    print(f"token_file={args.token_file.expanduser()}")
    print("Set MCP env:")
    print(f"  POWERPOINT_LIVE_BRIDGE_URL=http://{args.host}:{args.port}")
    print(f"  POWERPOINT_LIVE_BRIDGE_TOKEN_FILE={args.token_file.expanduser()}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nPowerPoint live bridge stopped")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
