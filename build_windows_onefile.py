#!/usr/bin/env python3
"""Build the Office MCP server as a Windows directory bundle.

Produces a --onedir PyInstaller bundle (fast startup, no temp extraction).
The output is intended to be packaged into an Inno Setup installer.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from subprocess import run
from sys import executable


def _discover_hidden_tool_imports(server_dir: Path) -> list[str]:
    """Return fully qualified hidden imports for all modules in the tools package."""
    tools_dir = server_dir / "tools"
    if not tools_dir.exists():
        return []

    return [
        f"tools.{module_path.stem}"
        for module_path in sorted(tools_dir.glob("*.py"))
        if module_path.stem != "__init__"
    ]


def build_windows_onefile(name: str, output_dir: str, clean: bool) -> None:
    """Build a directory-mode executable using PyInstaller."""
    server_dir = Path(__file__).resolve().parent
    entrypoint = server_dir / "office_server.py"
    build_dir = server_dir / "build"
    dist_dir = server_dir / output_dir
    hidden_tool_imports = _discover_hidden_tool_imports(server_dir)

    args = [
        executable,
        "-m",
        "PyInstaller",
        str(entrypoint),
        "--name",
        name,
        "--onedir",
        "--console",
        "--collect-submodules",
        "tools",
        "--hidden-import",
        "aioumcp",
        # Core dependencies imported inside try/except blocks in tool modules;
        # PyInstaller may skip them during static analysis.
        "--hidden-import",
        "lxml",
        "--hidden-import",
        "lxml.etree",
        "--hidden-import",
        "lxml.html",
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "docx",
        "--hidden-import",
        "pptx",
        "--hidden-import",
        "markdown_it",
        "--hidden-import",
        "mdit_py_plugins",
        "--hidden-import",
        "PIL",
        # Optional web dependencies (try/except guarded in tool modules)
        "--hidden-import",
        "requests",
        "--hidden-import",
        "readability",
        "--hidden-import",
        "markdownify",
        "--hidden-import",
        "bs4",
        "--paths",
        str(server_dir),
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(build_dir),
        "--specpath",
        str(build_dir),
    ]

    for module_name in hidden_tool_imports:
        args.extend(["--hidden-import", module_name])

    if clean:
        args.append("--clean")

    completed = run(args, cwd=str(server_dir), check=False)
    if completed.returncode != 0:
        raise SystemExit(f"PyInstaller build failed with exit code {completed.returncode}")


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""
    parser = argparse.ArgumentParser(
        description="Build the Office MCP server as a single-file Windows executable."
    )
    parser.add_argument(
        "--name",
        default="office-mcp-server",
        help="Output executable base name (default: office-mcp-server)",
    )
    parser.add_argument(
        "--output-dir",
        default="dist",
        help="Output directory relative to .github/mcp (default: dist)",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Clean PyInstaller cache and temporary files before build",
    )
    return parser.parse_args()


def main() -> None:
    """CLI entry point."""
    args = parse_args()
    build_windows_onefile(name=args.name, output_dir=args.output_dir, clean=args.clean)


if __name__ == "__main__":
    main()
