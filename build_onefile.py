#!/usr/bin/env python3
"""Build the Office MCP server executable.

Uses uv + PyInstaller on Linux/macOS (--onefile), and direct PyInstaller
on Windows (--onedir for fast startup + Inno Setup installer packaging).

Usage:
    python build_onefile.py                      # auto-detect platform
    python build_onefile.py --target linux        # force Linux build (uv + PyInstaller)
    python build_onefile.py --target windows      # force Windows build (PyInstaller)
    python build_onefile.py --clean               # clean build artifacts first
    python build_onefile.py --name my-server      # custom output name
    python build_onefile.py --output-dir dist     # custom output directory

    python build_onefile.py --help
"""

from __future__ import annotations

import argparse
import platform
import shutil
from pathlib import Path
from shutil import which
from subprocess import run

SERVER_DIR = Path(__file__).resolve().parent
ENTRYPOINT = SERVER_DIR / "office_server.py"
BUILD_DIR = SERVER_DIR / "build"
DEFAULT_NAME = "office-mcp-server"


def detect_target() -> str:
    """Return 'linux' or 'windows' based on the current platform."""
    system = platform.system().lower()
    if system in ("linux", "darwin"):
        return "linux"
    if system == "windows":
        return "windows"
    raise SystemExit(f"Unsupported platform: {system}")


def _pyinstaller_args(name: str, dist_dir: Path, clean: bool, *, onedir: bool = False) -> list[str]:
    """Return the common PyInstaller argument list."""
    mode = "--onedir" if onedir else "--onefile"
    args = [
        "pyinstaller",
        str(ENTRYPOINT),
        "--name", name,
        mode,
        "--console",
        "--collect-submodules", "tools",
        "--hidden-import", "aioumcp",
        # Core deps imported inside try/except in tool modules
        "--hidden-import", "lxml",
        "--hidden-import", "lxml.etree",
        "--hidden-import", "lxml.html",
        "--hidden-import", "openpyxl",
        "--hidden-import", "docx",
        "--hidden-import", "pptx",
        "--hidden-import", "markdown_it",
        "--hidden-import", "mdit_py_plugins",
        "--hidden-import", "PIL",
        # Optional web deps (try/except guarded)
        "--hidden-import", "requests",
        "--hidden-import", "readability",
        "--hidden-import", "markdownify",
        "--hidden-import", "bs4",
        "--paths", str(SERVER_DIR),
        "--distpath", str(dist_dir),
        "--workpath", str(BUILD_DIR),
        "--specpath", str(BUILD_DIR),
    ]
    if clean:
        args.append("--clean")
    return args


# ---------------------------------------------------------------------------
# Linux / macOS build via uv + PyInstaller
# ---------------------------------------------------------------------------

def build_linux_onefile(name: str, output_dir: str, clean: bool) -> None:
    """Build a one-file executable using uv + PyInstaller (Linux / macOS).

    uv handles dependency resolution and environment creation.
    PyInstaller produces the single-file binary.
    """
    if which("uv") is None:
        raise SystemExit(
            "uv is not installed. Install it with: curl -LsSf https://astral.sh/uv/install.sh | sh"
        )

    dist_dir = SERVER_DIR / output_dir
    dist_dir.mkdir(parents=True, exist_ok=True)

    if clean and BUILD_DIR.exists():
        shutil.rmtree(BUILD_DIR)
        print(f"Cleaned {BUILD_DIR}")

    pyi_args = _pyinstaller_args(name, dist_dir, clean)

    # Use uv run to invoke PyInstaller inside a temporary environment that
    # includes the project dependencies plus the [build] extra.
    # --python 3.12 ensures uv uses a managed Python with shared library support
    # (required by PyInstaller), avoiding issues with system Python builds.
    args = [
        "uv", "run",
        "--python", "3.12",
        "--project", str(SERVER_DIR),
        "--extra", "build",
        "--extra", "web",
    ] + pyi_args

    print(f"Building Linux binary: {dist_dir / name}")
    print(f"  Command: {' '.join(args)}")

    completed = run(args, cwd=str(SERVER_DIR), check=False)
    if completed.returncode != 0:
        raise SystemExit(f"Build failed with exit code {completed.returncode}")

    binary = dist_dir / name
    if binary.exists():
        binary.chmod(0o755)
        size_mb = binary.stat().st_size / (1024 * 1024)
        print(f"Build succeeded: {binary} ({size_mb:.1f} MB)")
    else:
        raise SystemExit(f"Expected binary not found at {binary}")


# ---------------------------------------------------------------------------
# Windows build via PyInstaller (preserves existing workflow)
# ---------------------------------------------------------------------------

def build_windows_onefile(name: str, output_dir: str, clean: bool) -> None:
    """Build a one-file executable using PyInstaller (Windows).

    If uv is available, uses it to run PyInstaller. Otherwise falls back
    to a direct PyInstaller invocation (assumes pip-installed dependencies).
    """
    dist_dir = SERVER_DIR / output_dir
    dist_dir.mkdir(parents=True, exist_ok=True)

    # Windows uses --onedir for fast startup (no temp extraction)
    pyi_args = _pyinstaller_args(name, dist_dir, clean, onedir=True)

    if which("uv") is not None:
        # Prefer uv when available
        args = [
            "uv", "run",
            "--project", str(SERVER_DIR),
            "--extra", "build",
            "--extra", "web",
        ] + pyi_args
    elif which("pyinstaller") is not None:
        # Fallback: direct PyInstaller (original Windows workflow)
        args = pyi_args
    else:
        raise SystemExit(
            "Neither uv nor pyinstaller found.\n"
            "Install uv:          https://docs.astral.sh/uv/getting-started/installation/\n"
            "Or install directly:  pip install pyinstaller"
        )

    print(f"Building Windows binary: {dist_dir / name / name}.exe")
    print(f"  Command: {' '.join(args)}")

    completed = run(args, cwd=str(SERVER_DIR), check=False)
    if completed.returncode != 0:
        raise SystemExit(f"Build failed with exit code {completed.returncode}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""
    parser = argparse.ArgumentParser(
        description="Build the Office MCP server as a single-file executable.",
    )
    parser.add_argument(
        "--target",
        choices=["linux", "windows"],
        default=None,
        help="Target platform (default: auto-detect)",
    )
    parser.add_argument(
        "--name",
        default=DEFAULT_NAME,
        help=f"Output executable base name (default: {DEFAULT_NAME})",
    )
    parser.add_argument(
        "--output-dir",
        default="dist",
        help="Output directory relative to .github/mcp (default: dist)",
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Clean build artifacts before building",
    )
    return parser.parse_args()


def main() -> None:
    """CLI entry point."""
    args = parse_args()
    target = args.target or detect_target()

    print(f"Target platform: {target}")

    if target == "linux":
        build_linux_onefile(name=args.name, output_dir=args.output_dir, clean=args.clean)
    else:
        build_windows_onefile(name=args.name, output_dir=args.output_dir, clean=args.clean)


if __name__ == "__main__":
    main()
