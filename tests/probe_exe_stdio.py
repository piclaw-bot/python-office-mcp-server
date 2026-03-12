"""Probe only the packaged EXE for interactive stdio initialize response."""

from __future__ import annotations

import json
import subprocess
import threading
import time
from pathlib import Path
from queue import Empty, Queue


def main() -> int:
    mcp_root = Path(__file__).resolve().parents[1]
    exe_path = mcp_root / "dist" / "office-mcp-server.exe"

    process = subprocess.Popen(
        [str(exe_path)],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    queue: Queue[str] = Queue()

    def read_stdout_line() -> None:
        assert process.stdout is not None
        queue.put(process.stdout.readline())

    threading.Thread(target=read_stdout_line, daemon=True).start()

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
        response_line = queue.get(timeout=10)
    except Empty:
        response_line = ""

    print(f"response_received={bool(response_line.strip())}")
    print(f"response={response_line.strip()[:1000]}")

    if process.poll() is None:
        process.terminate()
        time.sleep(0.5)
        if process.poll() is None:
            process.kill()

    print(f"returncode={process.returncode}")
    return 0 if response_line.strip() else 1


if __name__ == "__main__":
    raise SystemExit(main())
