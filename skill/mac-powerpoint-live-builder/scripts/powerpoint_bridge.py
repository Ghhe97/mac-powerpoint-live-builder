#!/usr/bin/env python3
"""Local localhost bridge for PowerPoint AppleScript control.

Run this outside restrictive Agent sandboxes. The MCP server can then set
POWERPOINT_LIVE_BRIDGE_URL and POWERPOINT_LIVE_BRIDGE_TOKEN_FILE to proxy
AppleScript execution through this process.
"""

from __future__ import annotations

import argparse
import json
import os
import secrets
import subprocess
import sys
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit


DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 18765
DEFAULT_HOME = Path.home() / ".local" / "share" / "powerpoint-live-mcp"
DEFAULT_TOKEN_FILE = DEFAULT_HOME / "bridge_token"
DEFAULT_RUNNER_PYTHON = DEFAULT_HOME / ".venv" / "bin" / "python"
DEFAULT_JOB_DIR = DEFAULT_HOME / "bridge_jobs"
SELF_TEST_SCRIPT = '''
tell application "Microsoft PowerPoint"
    try
        activate
    end try
    return name
end tell
'''


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
            "stderr": (
                f"osascript timed out after {timeout}s. If this bridge was launched "
                "from Terminal, enable System Settings > Privacy & Security > "
                "Automation > Terminal > Microsoft PowerPoint, then restart the bridge."
            ),
        }
    return {
        "ok": result.returncode == 0,
        "returncode": result.returncode,
        "stdout": result.stdout,
        "stderr": result.stderr,
    }


def automation_self_test(timeout: int = 8) -> dict[str, Any]:
    result = run_osascript(SELF_TEST_SCRIPT, timeout)
    result["check"] = "powerpoint_automation"
    if result["ok"]:
        result["message"] = "PowerPoint Automation is available to the bridge launcher."
    else:
        result["message"] = (
            "The bridge is running, but its launcher cannot complete a PowerPoint "
            "Automation call. If launched from Terminal, enable System Settings > "
            "Privacy & Security > Automation > Terminal > Microsoft PowerPoint, "
            "then restart the bridge."
        )
    return result


def run_live_sequence(payload: dict[str, Any], timeout: int) -> dict[str, Any]:
    runner = Path(__file__).with_name("run_pptx_sequence.py")
    python = DEFAULT_RUNNER_PYTHON if DEFAULT_RUNNER_PYTHON.exists() else Path(sys.executable)
    default_delay = str(float(payload.get("default_delay_seconds", 0.35)))
    cmd = [str(python), str(runner), "--default-delay-seconds", default_delay]
    sequence_path: Path | None = None

    if payload.get("demo_pptx"):
        output_path = str(Path(str(payload["demo_pptx"])).expanduser())
        cmd.extend(["--demo-pptx", output_path])
    else:
        output_path = ""
        steps = payload.get("steps")
        if not isinstance(steps, list):
            raise ValueError("run-live-sequence requires `steps` or `demo_pptx`")
        DEFAULT_JOB_DIR.mkdir(parents=True, exist_ok=True)
        sequence_path = DEFAULT_JOB_DIR / f"sequence-{uuid.uuid4().hex}.json"
        sequence_path.write_text(json.dumps({"steps": steps}, ensure_ascii=False), encoding="utf-8")
        cmd.append(str(sequence_path))

    try:
        print(f"Delegated live sequence starting: {' '.join(cmd)}")
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
            env={**dict(os.environ), "POWERPOINT_LIVE_BRIDGE_DELEGATED": "1"},
        )
        print(f"Delegated live sequence finished: returncode={result.returncode} output={output_path}")
        return {
            "ok": result.returncode == 0,
            "returncode": result.returncode,
            "stdout_tail": result.stdout[-4000:],
            "stderr_tail": result.stderr[-4000:],
            "output_path": output_path,
        }
    except subprocess.TimeoutExpired as e:
        print(f"Delegated live sequence timed out: output={output_path}")
        return {
            "ok": False,
            "returncode": 124,
            "stdout_tail": str(e.stdout or "")[-4000:],
            "stderr_tail": str(e.stderr or f"delegated live sequence timed out after {timeout}s")[-4000:],
            "output_path": output_path,
        }
    finally:
        if sequence_path:
            try:
                sequence_path.unlink()
            except OSError:
                pass


class BridgeHandler(BaseHTTPRequestHandler):
    server_version = "PowerPointLiveBridge/0.1"

    def _request_path(self) -> str:
        # Some sandboxed Agent runtimes/proxy layers send HTTP absolute-form
        # request targets such as "http://127.0.0.1:18765/health" to a local
        # server. Accept both origin-form and absolute-form paths.
        if self.path.startswith("http://") or self.path.startswith("https://"):
            parsed = urlsplit(self.path)
            return parsed.path or "/"
        return self.path

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

    def _read_request_body(self) -> bytes:
        transfer_encoding = self.headers.get("Transfer-Encoding", "").lower()
        if "chunked" in transfer_encoding:
            chunks: list[bytes] = []
            while True:
                line = self.rfile.readline()
                if not line:
                    break
                size_text = line.split(b";", 1)[0].strip()
                size = int(size_text, 16)
                if size == 0:
                    # Consume optional trailing headers.
                    while True:
                        trailer = self.rfile.readline()
                        if trailer in (b"\r\n", b"\n", b""):
                            break
                    break
                chunks.append(self.rfile.read(size))
                self.rfile.read(2)
            return b"".join(chunks)
        length = int(self.headers.get("Content-Length", "0"))
        return self.rfile.read(length)

    def do_GET(self) -> None:
        path = self._request_path()
        if path == "/health":
            if not self._authorized():
                self._send_json(401, {"ok": False, "error": "unauthorized"})
                return
            self._send_json(200, {"ok": True, "service": "powerpoint-live-bridge"})
            return
        if path == "/self-test":
            if not self._authorized():
                self._send_json(401, {"ok": False, "error": "unauthorized"})
                return
            self._send_json(200, automation_self_test())
            return
        else:
            self._send_json(404, {"ok": False, "error": "not found"})
            return

    def do_POST(self) -> None:
        path = self._request_path()
        if path not in {"/run-osascript", "/run-live-sequence"}:
            self._send_json(404, {"ok": False, "error": "not found"})
            return
        if not self._authorized():
            self._send_json(401, {"ok": False, "error": "unauthorized"})
            return
        raw_body = b""
        try:
            raw_body = self._read_request_body()
            if not raw_body:
                raise ValueError("empty JSON body")
            payload = json.loads(raw_body.decode("utf-8"))
            timeout = max(1, min(int(payload.get("timeout", 60)), 600))
        except Exception as e:
            body_preview = raw_body[:200].decode("utf-8", errors="replace")
            print(
                "Bad /run-osascript request: "
                f"path={path!r} content_length={self.headers.get('Content-Length', '0')!r} "
                f"transfer_encoding={self.headers.get('Transfer-Encoding', '')!r} "
                f"body_preview={body_preview!r} error={e}"
            )
            self._send_json(
                400,
                {
                    "ok": False,
                    "error": f"bad request: {e}",
                    "path": path,
                    "content_length": self.headers.get("Content-Length", "0"),
                    "transfer_encoding": self.headers.get("Transfer-Encoding", ""),
                    "body_preview": body_preview,
                },
            )
            return
        if path == "/run-live-sequence":
            try:
                self._send_json(200, run_live_sequence(payload, timeout))
            except Exception as e:
                self._send_json(400, {"ok": False, "error": f"bad run-live-sequence request: {e}"})
            return
        script = str(payload["script"])
        self._send_json(200, run_osascript(script, timeout))

    def log_message(self, format: str, *args: Any) -> None:
        print(f"{self.address_string()} - {format % args}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default=DEFAULT_HOST, help="Bind host. Defaults to 127.0.0.1.")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help="Bind port.")
    parser.add_argument("--token-file", type=Path, default=DEFAULT_TOKEN_FILE, help="Bridge token file.")
    parser.add_argument(
        "--startup-self-test",
        action="store_true",
        help="Run a short PowerPoint Automation check before serving.",
    )
    args = parser.parse_args()

    token = ensure_token(args.token_file.expanduser())
    if args.startup_self_test:
        print("Running PowerPoint Automation self-test...")
        result = automation_self_test()
        if result["ok"]:
            print("OK PowerPoint Automation self-test")
        else:
            print("WARN PowerPoint Automation self-test failed")
            print(result.get("stderr") or result.get("message"))
            print("The bridge will still start so MCP clients can report this clearly.")

    server = ThreadingHTTPServer((args.host, args.port), BridgeHandler)
    server.bridge_token = token  # type: ignore[attr-defined]
    print("PowerPoint live bridge started")
    print(f"url=http://{args.host}:{args.port}")
    print(f"token_file={args.token_file.expanduser()}")
    print("Set MCP env:")
    print(f"  POWERPOINT_LIVE_BRIDGE_URL=http://{args.host}:{args.port}")
    print(f"  POWERPOINT_LIVE_BRIDGE_TOKEN_FILE={args.token_file.expanduser()}")
    print("Automation note:")
    print("  The app that starts this bridge must be allowed to control Microsoft PowerPoint.")
    print("  If launched from Terminal, enable Automation > Terminal > Microsoft PowerPoint.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nPowerPoint live bridge stopped")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
