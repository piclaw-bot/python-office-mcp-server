#!/usr/bin/env python3
"""
save_utils.py - Safe save utilities for Office documents.

When saving Office documents (Word, Excel, PowerPoint) to the same file that was
opened, the underlying libraries (python-docx, openpyxl, python-pptx) can create
duplicate entries in the ZIP archive. This causes Office applications to prompt
for repair when opening.

This module provides safe save functions that write to a temporary file first,
then atomically replace the original, avoiding these issues.
"""

import os
import shutil
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any


def safe_save_office(document: Any, output_path: str, suffix: str = None) -> None:
    """Safely save an Office document to avoid ZIP duplicate entry issues.

    Works with python-docx Document, openpyxl Workbook, and python-pptx Presentation
    objects. Saves to a temp file first, then replaces the original.

    Args:
        document: Document object with a .save() method (Document, Workbook, or Presentation)
        output_path: Path to save to
        suffix: File extension suffix (e.g., '.docx', '.xlsx', '.pptx').
                If not provided, inferred from output_path.
    """
    output_path = str(output_path)

    # Infer suffix from output path if not provided
    if suffix is None:
        suffix = Path(output_path).suffix or ".tmp"

    # Create a temp file in the same directory to ensure same filesystem
    output_dir = os.path.dirname(output_path) or "."
    fd, temp_path = tempfile.mkstemp(suffix=suffix, dir=output_dir)
    os.close(fd)

    try:
        document.save(temp_path)
        # Replace original with temp file (atomic on same filesystem)
        shutil.move(temp_path, output_path)
    except Exception:
        # Clean up temp file on error
        if os.path.exists(temp_path):
            os.unlink(temp_path)
        raise


def safe_save_docx(doc, output_path: str) -> None:
    """Safely save a Word document.

    Args:
        doc: python-docx Document object
        output_path: Path to save to
    """
    safe_save_office(doc, output_path, ".docx")


def safe_save_xlsx(wb, output_path: str) -> None:
    """Safely save an Excel workbook.

    Args:
        wb: openpyxl Workbook object
        output_path: Path to save to
    """
    safe_save_office(wb, output_path, ".xlsx")


def safe_save_pptx(prs, output_path: str) -> None:
    """Safely save a PowerPoint presentation.

    Args:
        prs: python-pptx Presentation object
        output_path: Path to save to
    """
    safe_save_office(prs, output_path, ".pptx")
    _dedupe_zip_entries(str(output_path))


def _dedupe_zip_entries(zip_path: str) -> None:
    """Rewrite a ZIP package to remove duplicate member names.

    Keeps the last occurrence of each member path, matching what most ZIP
    readers resolve at load time, and rewrites the package without duplicates.
    """
    if not zipfile.is_zipfile(zip_path):
        return

    with zipfile.ZipFile(zip_path, "r") as zin:
        infos = zin.infolist()
        if len({info.filename for info in infos}) == len(infos):
            return

        keep_names: list[str] = []
        seen: set[str] = set()
        for info in reversed(infos):
            if info.filename in seen:
                continue
            seen.add(info.filename)
            keep_names.append(info.filename)
        keep_names.reverse()

        fd, tmp_path = tempfile.mkstemp(suffix=Path(zip_path).suffix or ".zip", dir=Path(zip_path).parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for name in keep_names:
                    zout.writestr(name, zin.read(name))
            shutil.move(tmp_path, zip_path)
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)


def resolve_office_path(file_path: str) -> str:
    """Resolve a file path against common workspace roots.

    Resolution order:
    1. Path as provided (absolute or cwd-relative)
    2. MCP and CI workspace roots from environment variables
    3. Repository root inferred from this module location
    4. Inferred repo_root/workspace sandbox path

    Returns:
        First existing path, otherwise original file_path
    """
    p = Path(file_path)
    if p.exists():
        return str(p)

    candidate_roots: list[Path] = []
    for env_var in ("MCP_WORKSPACE_ROOT", "GITHUB_WORKSPACE", "WORKSPACE_FOLDER"):
        raw = os.environ.get(env_var)
        if raw:
            candidate_roots.append(Path(raw))

    repo_root = Path(__file__).resolve().parents[3]
    candidate_roots.extend([repo_root, repo_root / "workspace", Path.cwd()])

    seen: set[str] = set()
    for root in candidate_roots:
        root_str = str(root)
        if root_str in seen:
            continue
        seen.add(root_str)

        candidate = root / file_path
        if candidate.exists():
            return str(candidate)

    return file_path


def open_pptx_with_retries(file_path: str, retries: int = 2, retry_delay: float = 0.2) -> tuple[Any | None, str, str | None]:
    """Open a PowerPoint file with retries and diagnostics.

    Returns:
        (presentation, resolved_path, error_message)
    """
    resolved_path = resolve_office_path(file_path)
    path = Path(resolved_path)

    if not path.exists():
        return None, resolved_path, f"File not found: {file_path}"

    try:
        from pptx import Presentation
        from pptx.exc import PackageNotFoundError
    except ImportError:
        return None, resolved_path, "python-pptx not installed. Run: pip install python-pptx"

    last_error: Exception | None = None
    for attempt in range(retries + 1):
        try:
            return Presentation(resolved_path), resolved_path, None
        except PackageNotFoundError as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(retry_delay)
                continue
        except Exception as exc:
            return None, resolved_path, f"Failed to open PowerPoint package: {exc}"

    if zipfile.is_zipfile(resolved_path):
        try:
            with zipfile.ZipFile(resolved_path) as zf:
                names = set(zf.namelist())
            required = {"[Content_Types].xml", "ppt/presentation.xml"}
            if required.issubset(names):
                return (
                    None,
                    resolved_path,
                    "PackageNotFoundError while opening a valid PPTX package. "
                    "File exists and has expected OOXML entries; this can be caused by transient file locks. "
                    "Retry the operation or save to a new output_path.",
                )
        except Exception:
            pass

    return None, resolved_path, f"PackageNotFoundError: {last_error}"


def open_docx_with_retries(file_path: str, retries: int = 2, retry_delay: float = 0.2) -> tuple[Any | None, str, str | None]:
    """Open a Word file with retries and diagnostics.

    Returns:
        (document, resolved_path, error_message)
    """
    resolved_path = resolve_office_path(file_path)
    path = Path(resolved_path)

    if not path.exists():
        return None, resolved_path, f"File not found: {file_path}"

    try:
        from docx import Document
        from docx.opc.exceptions import PackageNotFoundError
    except ImportError:
        return None, resolved_path, "python-docx not installed. Run: pip install python-docx"

    last_error: Exception | None = None
    for attempt in range(retries + 1):
        try:
            return Document(resolved_path), resolved_path, None
        except PackageNotFoundError as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(retry_delay)
                continue
        except Exception as exc:
            return None, resolved_path, f"Failed to open Word package: {exc}"

    if zipfile.is_zipfile(resolved_path):
        try:
            with zipfile.ZipFile(resolved_path) as zf:
                names = set(zf.namelist())
            required = {"[Content_Types].xml", "word/document.xml"}
            if required.issubset(names):
                return (
                    None,
                    resolved_path,
                    "PackageNotFoundError while opening a valid DOCX package. "
                    "File exists and has expected OOXML entries; this can be caused by transient file locks. "
                    "Retry the operation or save to a new output_path.",
                )
        except Exception:
            pass

    return None, resolved_path, f"PackageNotFoundError: {last_error}"
