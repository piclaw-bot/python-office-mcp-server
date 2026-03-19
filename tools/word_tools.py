#!/usr/bin/env python3
"""
word_tools.py - MCP tools for Word document processing

Provides tools to parse, extract content from, and generate Word (.docx) files.
Supports full GitHub Flavored Markdown input with headings, bullets, bold/italic,
strikethrough, tables, code blocks, and task lists.
"""

import re
from pathlib import Path
from typing import Any

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

from .markdown_parser import (
    CodeBlock,
    HorizontalRule,
    Paragraph,
    Table,
    TextRun,
    parse_markdown_to_nodes,
)
from .save_utils import open_docx_with_retries, resolve_office_path, safe_save_docx


class WordTools:
    """MCP tool mixin for Word document processing.

    Note: The extract/to_markdown methods are kept for internal use by
    office_unified_tools but are not exposed as separate MCP tools.
    They are filtered out in office_server.py via DEPRECATED_TOOLS.

    Public MCP tools:
    - word_from_markdown: Convert Markdown to Word document

    Internal methods (used by unified tools):
    - tool_word_extract: Extract text and structure
    - tool_word_to_markdown: Convert to Markdown format

    See WordAdvancedTools for SOW-specific and track-changes tools.
    """

    @staticmethod
    def _get_heading_level(style_name: str) -> int:
        """Extract heading level from Word style names.

        Handles variants like "Heading 1", "Heading 1 NUM", and localized
        custom names that still include an integer heading level.
        """
        if not style_name.startswith("Heading"):
            return 1

        match = re.search(r"\b(\d+)\b", style_name)
        if not match:
            return 1

        try:
            return int(match.group(1))
        except ValueError:
            return 1

    def tool_word_extract(self, file_path: str) -> dict[str, Any]:
        """Extract text and structure from a Word document.

        NOTE: This method is internal - use office_read() instead.

        Args:
            file_path: Path to the .docx file

        Returns:
            Dictionary with paragraphs, headings, tables, and metadata
        """
        if not HAS_DOCX:
            return {"error": "python-docx not installed. Run: pip install python-docx"}

        resolved_path = resolve_office_path(file_path)
        path = Path(resolved_path)
        if not path.exists():
            return {"error": f"File not found: {file_path}"}
        if path.suffix.lower() != ".docx":
            return {"error": f"Expected .docx file, got: {path.suffix}"}

        doc, _resolved_path, open_error = open_docx_with_retries(file_path)
        if open_error:
            return {"error": f"Failed to open document: {open_error}"}

        paragraphs = []
        headings = []
        tables = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            style_name = para.style.name if para.style else "Normal"

            if style_name.startswith("Heading"):
                level = self._get_heading_level(style_name)
                headings.append({"level": level, "text": text})

            paragraphs.append({
                "text": text,
                "style": style_name
            })

        for table in doc.tables:
            rows = []
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                rows.append(cells)
            tables.append(rows)

        return {
            "file": path.name,
            "paragraphs": paragraphs,
            "headings": headings,
            "tables": tables,
            "paragraph_count": len(paragraphs),
            "table_count": len(tables)
        }

    def tool_word_to_markdown(self, file_path: str) -> str:
        """Convert a Word document to Markdown format.

        NOTE: This method is internal - use office_read(output_format="markdown") instead.

        Args:
            file_path: Path to the .docx file

        Returns:
            Markdown string representation of the document
        """
        if not HAS_DOCX:
            return "Error: python-docx not installed. Run: pip install python-docx"

        resolved_path = resolve_office_path(file_path)
        path = Path(resolved_path)
        if not path.exists():
            return f"Error: File not found: {file_path}"

        doc, _resolved_path, open_error = open_docx_with_retries(file_path)
        if open_error:
            return f"Error: Failed to open document: {open_error}"
        lines = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                lines.append("")
                continue

            style_name = para.style.name if para.style else "Normal"

            if style_name.startswith("Heading"):
                level = self._get_heading_level(style_name)
                lines.append(f"{'#' * level} {text}")
            elif style_name == "List Bullet":
                lines.append(f"- {text}")
            elif style_name == "List Number":
                lines.append(f"1. {text}")
            else:
                lines.append(text)
            lines.append("")

        # Process tables
        for table in doc.tables:
            if not table.rows:
                continue

            # Header row
            header = [cell.text.strip() for cell in table.rows[0].cells]
            lines.append("| " + " | ".join(header) + " |")
            lines.append("| " + " | ".join(["---"] * len(header)) + " |")

            # Data rows
            for row in table.rows[1:]:
                cells = [cell.text.strip() for cell in row.cells]
                lines.append("| " + " | ".join(cells) + " |")
            lines.append("")

        return "\n".join(lines)

    def tool_word_from_markdown(
        self,
        output_path: str,
        markdown: str | None = None,
        markdown_file: str | None = None,
    ) -> dict[str, Any]:
        """Convert full Markdown content to a Word document.

        This is the primary tool for creating Word documents from text content.

        Supports GitHub Flavored Markdown:
        - Headings (# ## ### ####)
        - Bullet lists (- or *)
        - Numbered lists (1. 2. 3.)
        - Task lists (- [ ] and - [x])
        - Bold (**text**) and italic (*text*) inline formatting
        - Strikethrough (~~text~~)
        - Inline code (`code`) rendered in Consolas font
        - Tables (| col | col |) with proper formatting
        - Code blocks (```) rendered in monospace with language hints
        - Horizontal rules (---)

        Example:
            word_from_markdown(
                output_path="04. Artifacts/report.docx",
                markdown='''
# Project Status Report

## Executive Summary
The project is **on track** for Q4 delivery with ~~no~~ minor delays.

## Key Metrics
| Metric | Value | Status |
|--------|-------|--------|
| Budget | $120,000 | On track |
| Timeline | Q4 2026 | Green |
| Quality | 95% | Exceeds |

## Next Steps
- [x] Complete UAT testing
- [ ] Finalize documentation
- [ ] Schedule go-live review

1. Phase 1 complete
2. Phase 2 in progress
3. Phase 3 planned
'''
            )

        Args:
            output_path: Path for the output .docx file
            markdown: Full GitHub Flavored Markdown content (inline)
            markdown_file: Optional path to a Markdown file. Use this for
                very large documents to avoid MCP argument-size limits.

        Returns:
            Status dictionary with file path
        """
        if not HAS_DOCX:
            return {"error": "python-docx not installed. Run: pip install python-docx"}

        markdown_text = markdown
        if markdown_file:
            resolved_md_path = resolve_office_path(markdown_file)
            md_path = Path(resolved_md_path)
            if not md_path.exists():
                return {"error": f"Markdown file not found: {markdown_file}"}
            if not md_path.is_file():
                return {"error": f"Markdown path is not a file: {markdown_file}"}
            markdown_text = md_path.read_text(encoding="utf-8")

        if markdown_text is None:
            return {"error": "Provide either 'markdown' or 'markdown_file'"}

        doc = Document()

        # Parse Markdown to structured nodes using GFM parser
        nodes = parse_markdown_to_nodes(markdown_text)

        def get_list_style(style_name: str, fallback: str = None) -> str | None:
            """Get a list style name, falling back to alternatives if not found."""
            # Try the requested style first
            try:
                doc.styles[style_name]
                return style_name
            except KeyError:
                pass

            # Try common alternatives
            alternatives = {
                "List Bullet": ["Bullet 1", "Bullet list", "List Paragraph"],
                "List Number": ["Number 1", "Numbered list", "List Paragraph"],
            }
            for alt in alternatives.get(style_name, []):
                try:
                    doc.styles[alt]
                    return alt
                except KeyError:
                    pass

            return fallback

        def add_runs_to_paragraph(para, runs: list[TextRun]):
            """Add formatted text runs to a paragraph."""
            for run_data in runs:
                run = para.add_run(run_data.text)
                if run_data.bold:
                    run.bold = True
                if run_data.italic:
                    run.italic = True
                if run_data.code:
                    run.font.name = "Consolas"
                if run_data.strikethrough:
                    run.font.strike = True

        def add_table(table: Table):
            """Add a table to the document."""
            if not table.rows:
                return

            num_rows = len(table.rows)
            num_cols = table.column_count

            doc_table = doc.add_table(rows=num_rows, cols=num_cols)
            doc_table.style = "Table Grid"

            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    if col_idx < num_cols:
                        doc_cell = doc_table.rows[row_idx].cells[col_idx]
                        para = doc_cell.paragraphs[0]
                        para.clear()
                        add_runs_to_paragraph(para, cell.runs)

            # Add spacing after table
            doc.add_paragraph()

        def add_code_block(code_block: CodeBlock):
            """Add a code block to the document."""
            para = doc.add_paragraph()
            run = para.add_run(code_block.code)
            run.font.name = "Consolas"

        # Get available list styles
        bullet_style = get_list_style("List Bullet")
        number_style = get_list_style("List Number")

        for node in nodes:
            if isinstance(node, Paragraph):
                # Heading
                if node.level > 0:
                    doc.add_heading(node.text, level=min(node.level, 4))
                # Bullet list
                elif node.list_type == "bullet":
                    # Handle task lists
                    if node.task_checked is not None:
                        prefix = "☑ " if node.task_checked else "☐ "
                        para = doc.add_paragraph(style=bullet_style) if bullet_style else doc.add_paragraph()
                        if not bullet_style:
                            para.add_run("• ")
                        para.add_run(prefix)
                        add_runs_to_paragraph(para, node.runs)
                    else:
                        para = doc.add_paragraph(style=bullet_style) if bullet_style else doc.add_paragraph()
                        if not bullet_style:
                            para.add_run("• ")
                        add_runs_to_paragraph(para, node.runs)
                # Numbered list
                elif node.list_type == "ordered":
                    para = doc.add_paragraph(style=number_style) if number_style else doc.add_paragraph()
                    add_runs_to_paragraph(para, node.runs)
                # Regular paragraph
                else:
                    if node.runs:
                        para = doc.add_paragraph()
                        add_runs_to_paragraph(para, node.runs)

            elif isinstance(node, Table):
                add_table(node)

            elif isinstance(node, CodeBlock):
                add_code_block(node)

            elif isinstance(node, HorizontalRule):
                doc.add_paragraph()

        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        safe_save_docx(doc, str(output))
        return {"success": True, "file": output_path, "message": "Created Word document from Markdown"}

    def tool_word_get_comments(self, file_path: str) -> dict[str, Any]:
        """Extract all comments from a Word document.

        Reads the comments.xml part of the docx package and returns
        structured comment data including author, date, and text.

        Args:
            file_path: Path to the .docx file

        Returns:
            Dictionary with comments list and count
        """
        import zipfile
        from datetime import datetime

        resolved_path = resolve_office_path(file_path)
        path = Path(resolved_path)
        if not path.exists():
            return {"error": f"File not found: {file_path}"}
        if path.suffix.lower() != ".docx":
            return {"error": f"Expected .docx file, got: {path.suffix}"}

        try:
            from lxml import etree
        except ImportError:
            return {"error": "lxml not installed. Run: pip install lxml"}

        W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        comments = []

        try:
            with zipfile.ZipFile(resolved_path, 'r') as zf:
                # Check if comments.xml exists
                if 'word/comments.xml' not in zf.namelist():
                    return {
                        "file": path.name,
                        "comments": [],
                        "comment_count": 0,
                        "note": "No comments in document"
                    }

                # Parse comments.xml
                comments_xml = zf.read('word/comments.xml')
                root = etree.fromstring(comments_xml)

                for comment_elem in root.findall(f'.//{{{W_NS}}}comment'):
                    comment_id = comment_elem.get(f'{{{W_NS}}}id')
                    author = comment_elem.get(f'{{{W_NS}}}author', 'Unknown')
                    date_str = comment_elem.get(f'{{{W_NS}}}date', '')
                    initials = comment_elem.get(f'{{{W_NS}}}initials', '')

                    # Extract comment text
                    text = ''.join(comment_elem.itertext()).strip()

                    # Parse date if present
                    date_formatted = date_str
                    if date_str:
                        try:
                            # Handle ISO format with Z suffix
                            dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                            date_formatted = dt.strftime('%Y-%m-%d %H:%M')
                        except (ValueError, TypeError):
                            pass

                    comments.append({
                        "id": comment_id,
                        "author": author,
                        "initials": initials,
                        "date": date_formatted,
                        "text": text
                    })

            return {
                "file": path.name,
                "comments": comments,
                "comment_count": len(comments)
            }

        except Exception as e:
            return {"error": f"Failed to extract comments: {e}"}

    def tool_word_delete_comment(
        self,
        file_path: str,
        comment_id: str,
        output_path: str | None = None,
    ) -> dict[str, Any]:
        """Delete a comment from a Word document by comment ID.

        Removes the comment from comments.xml and strips associated reference
        markers from document.xml.

        Args:
            file_path: Path to the .docx file
            comment_id: Comment ID from word_get_comments output
            output_path: Optional output path (defaults to overwriting input)

        Returns:
            Status dictionary with deletion details
        """
        import tempfile
        import zipfile

        resolved_path = resolve_office_path(file_path)
        path = Path(resolved_path)
        if not path.exists():
            return {"error": f"File not found: {file_path}"}
        if path.suffix.lower() != ".docx":
            return {"error": f"Expected .docx file, got: {path.suffix}"}

        try:
            from lxml import etree
        except ImportError:
            return {"error": "lxml not installed. Run: pip install lxml"}

        W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        comment_id = str(comment_id)
        try:
            with zipfile.ZipFile(resolved_path, 'r') as zf_in:
                if 'word/comments.xml' not in zf_in.namelist():
                    return {"error": "No comments.xml found in document"}

                comments_root = etree.fromstring(zf_in.read('word/comments.xml'))
                comment_nodes = comments_root.findall(f'.//{{{W_NS}}}comment')
                removed_nodes = [
                    node for node in comment_nodes
                    if node.get(f'{{{W_NS}}}id') == comment_id
                ]

                if not removed_nodes:
                    return {"error": f"Comment ID not found: {comment_id}"}

                for node in removed_nodes:
                    parent = node.getparent()
                    if parent is not None:
                        parent.remove(node)

                comments_bytes = etree.tostring(
                    comments_root,
                    xml_declaration=True,
                    encoding='UTF-8',
                    standalone='yes',
                )

                doc_root = etree.fromstring(zf_in.read('word/document.xml'))
                ref_tags = (
                    f'{{{W_NS}}}commentRangeStart',
                    f'{{{W_NS}}}commentRangeEnd',
                    f'{{{W_NS}}}commentReference',
                )
                removed_refs = 0
                for tag in ref_tags:
                    for node in list(doc_root.findall(f'.//{tag}')):
                        if node.get(f'{{{W_NS}}}id') == comment_id:
                            parent = node.getparent()
                            if parent is not None:
                                parent.remove(node)
                                removed_refs += 1

                doc_bytes = etree.tostring(
                    doc_root,
                    xml_declaration=True,
                    encoding='UTF-8',
                    standalone='yes',
                )

                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                    tmp_path = tmp.name

                with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    for item in zf_in.namelist():
                        if item == 'word/comments.xml':
                            zf_out.writestr(item, comments_bytes)
                        elif item == 'word/document.xml':
                            zf_out.writestr(item, doc_bytes)
                        else:
                            zf_out.writestr(item, zf_in.read(item))

            if output_path:
                from shutil import move
                move(tmp_path, output_path)
            else:
                from shutil import move
                move(tmp_path, resolved_path)

            return {
                "success": True,
                "file": output_path or resolved_path,
                "comment_id": comment_id,
                "removed_comment_nodes": len(removed_nodes),
                "removed_reference_nodes": removed_refs,
                "message": "Comment deleted",
            }
        except Exception as e:
            return {"error": f"Failed to delete comment: {e}"}
