"""Run a broad integration test suite against the packaged MCP binary."""

from __future__ import annotations

import json
import subprocess
import tempfile
from pathlib import Path


def _extract_json_response(stdout: str) -> dict:
    lines = [line.strip() for line in stdout.splitlines() if line.strip()]
    for line in reversed(lines):
        if line.startswith("{") and line.endswith("}"):
            return json.loads(line)
    raise ValueError(f"No JSON response found in stdout: {stdout[:500]}")


def _run_request(exe_path: Path, request: dict) -> dict:
    with tempfile.NamedTemporaryFile(mode="w", suffix=".json", encoding="utf-8", delete=False) as temp_file:
        json.dump(request, temp_file)
        request_path = Path(temp_file.name)

    try:
        result = subprocess.run(
            [str(exe_path), str(request_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"Binary exited with code {result.returncode}\n"
                f"STDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
            )
        return _extract_json_response(result.stdout)
    finally:
        request_path.unlink(missing_ok=True)


def _make_request(request_id: int, method: str, params: dict | None = None) -> dict:
    payload = {
        "jsonrpc": "2.0",
        "id": request_id,
        "method": method,
        "params": params or {},
    }
    return payload


def main() -> int:
    repo_root = Path(__file__).resolve().parents[3]
    mcp_root = repo_root / ".github" / "mcp"
    exe_path = mcp_root / "dist" / "office-mcp-server.exe"

    agile_docx = repo_root / ".github" / "skills" / "statement-of-work" / "templates" / "Agile.docx"
    sow_pptx = repo_root / ".github" / "skills" / "present-sow" / "templates" / "SOW-Presentation.pptx"

    with tempfile.TemporaryDirectory() as temp_dir_str:
        temp_dir = Path(temp_dir_str)
        copied_docx = temp_dir / "agile-copy.docx"
        generated_xlsx = temp_dir / "generated.xlsx"
        generated_pptx = temp_dir / "generated.pptx"

        markdown_table = "| Name | Value |\n|---|---|\n| A | 1 |\n| B | 2 |"
        markdown_slides = "# Binary Test Deck\n\n## Summary\n- Item 1\n- Item 2\n"

        cases = [
            ("initialize", _make_request(1, "initialize")),
            ("tools/list", _make_request(2, "tools/list")),
            (
                "tools/call:list_supported_formats",
                _make_request(
                    3,
                    "tools/call",
                    {"name": "list_supported_formats", "arguments": {}},
                ),
            ),
            (
                "tools/call:word_parse_sow_template",
                _make_request(
                    4,
                    "tools/call",
                    {
                        "name": "word_parse_sow_template",
                        "arguments": {"template_path": str(agile_docx)},
                    },
                ),
            ),
            (
                "tools/call:office_read_docx_json",
                _make_request(
                    5,
                    "tools/call",
                    {
                        "name": "office_read",
                        "arguments": {"file_path": str(agile_docx), "output_format": "json"},
                    },
                ),
            ),
            (
                "tools/call:office_read_docx_markdown",
                _make_request(
                    6,
                    "tools/call",
                    {
                        "name": "office_read",
                        "arguments": {
                            "file_path": str(agile_docx),
                            "output_format": "markdown",
                        },
                    },
                ),
            ),
            (
                "tools/call:office_read_pptx_json",
                _make_request(
                    7,
                    "tools/call",
                    {
                        "name": "office_read",
                        "arguments": {"file_path": str(sow_pptx), "output_format": "json"},
                    },
                ),
            ),
            (
                "tools/call:office_template_analyze",
                _make_request(
                    8,
                    "tools/call",
                    {
                        "name": "office_template",
                        "arguments": {
                            "source_path": str(agile_docx),
                            "destination_path": str(copied_docx),
                            "operation": "analyze",
                        },
                    },
                ),
            ),
            (
                "tools/call:office_template_copy",
                _make_request(
                    9,
                    "tools/call",
                    {
                        "name": "office_template",
                        "arguments": {
                            "source_path": str(agile_docx),
                            "destination_path": str(copied_docx),
                            "operation": "copy",
                        },
                    },
                ),
            ),
            (
                "tools/call:office_audit_placeholders",
                _make_request(
                    10,
                    "tools/call",
                    {
                        "name": "office_audit",
                        "arguments": {"file_path": str(agile_docx), "checks": ["placeholders"]},
                    },
                ),
            ),
            (
                "tools/call:excel_from_markdown",
                _make_request(
                    11,
                    "tools/call",
                    {
                        "name": "excel_from_markdown",
                        "arguments": {
                            "output_path": str(generated_xlsx),
                            "markdown": markdown_table,
                        },
                    },
                ),
            ),
            (
                "tools/call:office_read_generated_xlsx",
                _make_request(
                    12,
                    "tools/call",
                    {
                        "name": "office_read",
                        "arguments": {
                            "file_path": str(generated_xlsx),
                            "output_format": "json",
                        },
                    },
                ),
            ),
            (
                "tools/call:office_inspect_generated_xlsx",
                _make_request(
                    13,
                    "tools/call",
                    {
                        "name": "office_inspect",
                        "arguments": {"file_path": str(generated_xlsx), "what": "sheets"},
                    },
                ),
            ),
            (
                "tools/call:pptx_from_markdown",
                _make_request(
                    14,
                    "tools/call",
                    {
                        "name": "pptx_from_markdown",
                        "arguments": {
                            "output_path": str(generated_pptx),
                            "markdown": markdown_slides,
                        },
                    },
                ),
            ),
            (
                "tools/call:office_read_generated_pptx",
                _make_request(
                    15,
                    "tools/call",
                    {
                        "name": "office_read",
                        "arguments": {
                            "file_path": str(generated_pptx),
                            "output_format": "json",
                        },
                    },
                ),
            ),
            (
                "tools/call:pptx_recommend_layout",
                _make_request(
                    16,
                    "tools/call",
                    {
                        "name": "pptx_recommend_layout",
                        "arguments": {"file_path": str(sow_pptx), "content_type": "bullets"},
                    },
                ),
            ),
            ("prompts/list", _make_request(17, "prompts/list")),
        ]

        results: list[tuple[str, bool, str]] = []

        for case_name, request in cases:
            try:
                response = _run_request(exe_path, request)
                if "error" in response:
                    results.append((case_name, False, response["error"].get("message", "Unknown error")))
                else:
                    results.append((case_name, True, "ok"))
            except Exception as exc:  # noqa: BLE001
                results.append((case_name, False, str(exc)))

        passed = sum(1 for _, ok, _ in results if ok)
        failed = len(results) - passed

        print("Binary integration suite")
        print(f"Executable: {exe_path}")
        print(f"Total: {len(results)}  Passed: {passed}  Failed: {failed}")
        print()

        for case_name, ok, detail in results:
            status = "PASS" if ok else "FAIL"
            print(f"[{status}] {case_name}")
            if not ok:
                print(f"       {detail}")

        return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
