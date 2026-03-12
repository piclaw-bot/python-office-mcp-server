"""Probe MCP stdio handshake behavior for Python and packaged executable."""

from __future__ import annotations

import json
import subprocess
import sys
import threading
from pathlib import Path
from queue import Empty, Queue


def probe(command: list[str], timeout_seconds: float = 8.0) -> tuple[bool, str, str, int | None]:
    """Start process, send initialize, and try to read one response line."""
    process = subprocess.Popen(
        command,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    queue: Queue[str] = Queue()

    def read_line() -> None:
        if process.stdout is None:
            queue.put("")
            return
        queue.put(process.stdout.readline())

    threading.Thread(target=read_line, daemon=True).start()

    request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {},
    }

    assert process.stdin is not None
    process.stdin.write(json.dumps(request) + "\n")
    process.stdin.flush()

    try:
        line = queue.get(timeout=timeout_seconds)
        ok = bool(line and line.strip())
    except Empty:
        line = ""
        ok = False

    if process.poll() is None:
        process.terminate()
        try:
            process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            process.kill()

    stderr_text = ""
    if process.stderr is not None:
        try:
            stderr_text = process.stderr.read()
        except Exception:
            stderr_text = ""

    return ok, line.strip(), stderr_text.strip(), process.returncode


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    py_cmd = [sys.executable, str(root / "office_server.py")]
    exe_cmd = [str(root / "dist" / "office-mcp-server.exe")]

    print("=== Python server stdio probe ===")
    py_ok, py_line, py_err, py_rc = probe(py_cmd)
    print(f"ok={py_ok} rc={py_rc}")
    print(f"line={py_line[:500]}")
    print(f"stderr={py_err[:500]}")

    print("=== EXE server stdio probe ===")
    exe_ok, exe_line, exe_err, exe_rc = probe(exe_cmd)
    print(f"ok={exe_ok} rc={exe_rc}")
    print(f"line={exe_line[:500]}")
    print(f"stderr={exe_err[:500]}")

    return 0 if py_ok and exe_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
