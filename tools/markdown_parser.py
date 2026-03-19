#!/usr/bin/env python3
"""
markdown_parser.py - GitHub Flavored Markdown parsing utilities

Provides a shared GFM parser using markdown-it-py that renders Markdown to HTML,
and traversal utilities using lxml to extract structured content for document generation.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from markdown_it import MarkdownIt

# Lazy-loaded module references — populated on first use to avoid pulling
# in heavyweight third-party libraries at import time (critical for fast
# PyInstaller one-file startup on Windows).
_markdown_it = None   # markdown_it.MarkdownIt class
_tasklists_plugin = None
_etree = None         # lxml.etree module
_fragment_fromstring = None
_fragments_fromstring = None


def _ensure_markdown_it():
    """Lazily import markdown-it-py on first use."""
    global _markdown_it, _tasklists_plugin
    if _markdown_it is not None:
        return
    try:
        from markdown_it import MarkdownIt as _MarkdownIt
        from mdit_py_plugins.tasklists import tasklists_plugin as _tp
        _markdown_it = _MarkdownIt
        _tasklists_plugin = _tp
    except ImportError as exc:
        raise ImportError(
            "markdown-it-py not installed. Run: pip install markdown-it-py mdit-py-plugins"
        ) from exc


def _ensure_lxml():
    """Lazily import lxml on first use."""
    global _etree, _fragment_fromstring, _fragments_fromstring
    if _etree is not None:
        return
    try:
        from lxml import etree as _e
        from lxml.html import fragment_fromstring as _ff
        from lxml.html import fragments_fromstring as _ffs
        _etree = _e
        _fragment_fromstring = _ff
        _fragments_fromstring = _ffs
    except ImportError as exc:
        raise ImportError("lxml not installed. Run: pip install lxml") from exc


# Singleton parser instance configured for GFM
_parser: MarkdownIt | None = None


def get_parser() -> MarkdownIt:
    """Get the shared GFM-configured markdown-it-py parser."""
    global _parser
    if _parser is None:
        _ensure_markdown_it()
        # Use "commonmark" preset with GFM extensions enabled
        # (gfm-like requires linkify which needs an extra dependency)
        _parser = (
            _markdown_it("commonmark")
            .enable("table")
            .enable("strikethrough")
            .use(_tasklists_plugin)
        )
    return _parser


def render_markdown(markdown: str) -> str:
    """Render Markdown to HTML using GFM rules.

    Args:
        markdown: GitHub Flavored Markdown content

    Returns:
        HTML string
    """
    parser = get_parser()
    return parser.render(markdown)


def parse_html(html: str):
    """Parse HTML string to an lxml element tree.

    Wraps fragments in a <div> container for consistent traversal.

    Args:
        html: HTML string (can be a fragment)

    Returns:
        lxml Element (root container)
    """
    _ensure_lxml()

    # Wrap in a div to ensure we have a single root
    wrapped = f"<div>{html}</div>"
    return _etree.HTML(wrapped).find(".//body/div")


@dataclass
class TextRun:
    """A run of text with formatting attributes."""
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    strikethrough: bool = False
    link_url: str | None = None


@dataclass
class TableCell:
    """A table cell with its content and alignment."""
    runs: list[TextRun] = field(default_factory=list)
    align: str | None = None  # 'left', 'center', 'right', or None

    @property
    def text(self) -> str:
        """Get plain text content of the cell."""
        return "".join(run.text for run in self.runs)


@dataclass
class TableRow:
    """A row of table cells."""
    cells: list[TableCell] = field(default_factory=list)
    is_header: bool = False


@dataclass
class Table:
    """A complete table with rows."""
    rows: list[TableRow] = field(default_factory=list)

    @property
    def header_row(self) -> TableRow | None:
        """Get the header row if present."""
        return self.rows[0] if self.rows and self.rows[0].is_header else None

    @property
    def data_rows(self) -> list[TableRow]:
        """Get data rows (excluding header)."""
        if not self.rows:
            return []
        return self.rows[1:] if self.rows[0].is_header else self.rows

    @property
    def column_count(self) -> int:
        """Get the number of columns."""
        return max((len(row.cells) for row in self.rows), default=0)


@dataclass
class Paragraph:
    """A paragraph with formatted runs."""
    runs: list[TextRun] = field(default_factory=list)
    level: int = 0  # Heading level (0 = normal paragraph, 1-6 = h1-h6)
    list_type: str | None = None  # 'bullet', 'ordered', or None
    list_index: int | None = None  # For ordered lists
    task_checked: bool | None = None  # For task lists

    @property
    def text(self) -> str:
        """Get plain text content."""
        return "".join(run.text for run in self.runs)


@dataclass
class CodeBlock:
    """A fenced code block."""
    code: str
    language: str | None = None


@dataclass
class HorizontalRule:
    """A horizontal rule / thematic break."""
    pass


# Type alias for content nodes
ContentNode = Paragraph | Table | CodeBlock | HorizontalRule


def _extract_runs(element, inherited_bold: bool = False,
                  inherited_italic: bool = False, inherited_code: bool = False,
                  inherited_strike: bool = False, inherited_link: str | None = None) -> list[TextRun]:
    """Extract text runs with formatting from an element and its children.

    Args:
        element: lxml element to extract from
        inherited_*: Formatting inherited from parent elements

    Returns:
        List of TextRun objects
    """
    runs = []

    # Handle text before first child
    if element.text:
        runs.append(TextRun(
            text=element.text,
            bold=inherited_bold,
            italic=inherited_italic,
            code=inherited_code,
            strikethrough=inherited_strike,
            link_url=inherited_link
        ))

    # Process children
    for child in element:
        tag = child.tag.lower() if isinstance(child.tag, str) else ""

        # Determine formatting for this child
        child_bold = inherited_bold or tag in ("strong", "b")
        child_italic = inherited_italic or tag in ("em", "i")
        child_code = inherited_code or tag in ("code",)
        child_strike = inherited_strike or tag in ("del", "s", "strike")
        child_link = child.get("href") if tag == "a" else inherited_link

        # Recursively extract runs from child
        child_runs = _extract_runs(
            child, child_bold, child_italic, child_code, child_strike, child_link
        )
        runs.extend(child_runs)

        # Handle tail text (text after this child element)
        if child.tail:
            runs.append(TextRun(
                text=child.tail,
                bold=inherited_bold,
                italic=inherited_italic,
                code=inherited_code,
                strikethrough=inherited_strike,
                link_url=inherited_link
            ))

    return runs


def _parse_table(table_elem) -> Table:
    """Parse an HTML table element into a Table object.

    Args:
        table_elem: lxml <table> element

    Returns:
        Table object with rows and cells
    """
    table = Table()

    # Process thead rows
    thead = table_elem.find(".//thead")
    if thead is not None:
        for tr in thead.findall(".//tr"):
            row = TableRow(is_header=True)
            # Find th and td elements
            for cell_elem in list(tr):
                if cell_elem.tag in ("th", "td"):
                    cell = TableCell(
                        runs=_extract_runs(cell_elem),
                        align=cell_elem.get("align") or cell_elem.get("style", "").split("text-align:")[-1].split(";")[0].strip() or None
                    )
                    row.cells.append(cell)
            if row.cells:
                table.rows.append(row)

    # Process tbody rows
    tbody = table_elem.find(".//tbody")
    rows_container = tbody if tbody is not None else table_elem

    for tr in rows_container.findall(".//tr"):
        # Skip if this row was already processed in thead
        if thead is not None and tr.getparent() == thead:
            continue

        row = TableRow(is_header=False)
        # Find th and td elements
        for cell_elem in list(tr):
            if cell_elem.tag in ("th", "td"):
                cell = TableCell(
                    runs=_extract_runs(cell_elem),
                    align=cell_elem.get("align") or cell_elem.get("style", "").split("text-align:")[-1].split(";")[0].strip() or None
                )
                row.cells.append(cell)
        if row.cells:
            table.rows.append(row)

    # If no thead, mark first row as header (common in Markdown tables)
    if thead is None and table.rows and not table.rows[0].is_header:
        table.rows[0].is_header = True

    return table


def _extract_list_item_runs(li_elem) -> list[TextRun]:
    """Extract only the direct content of a list item.

    This intentionally excludes nested ``<ul>/<ol>`` blocks and task-list
    checkbox ``<input>`` elements to avoid duplicate text when nested lists are
    rendered as separate list paragraphs.
    """
    runs: list[TextRun] = []

    if li_elem.text:
        runs.append(TextRun(text=li_elem.text))

    for child in li_elem:
        tag = child.tag.lower() if isinstance(child.tag, str) else ""

        # Nested lists are parsed separately by _parse_list.
        if tag in ("ul", "ol"):
            continue

        # Task-list checkbox metadata is handled separately.
        if tag == "input" and (child.get("type") or "").lower() == "checkbox":
            continue

        runs.extend(_extract_runs(child))

        if child.tail:
            runs.append(TextRun(text=child.tail))

    return runs


def _parse_list(list_elem, list_type: str) -> list[Paragraph]:
    """Parse an HTML list element into Paragraph objects.

    Args:
        list_elem: lxml <ul> or <ol> element
        list_type: 'bullet' or 'ordered'

    Returns:
        List of Paragraph objects
    """
    paragraphs = []

    for idx, li in enumerate(list_elem.findall("./li"), start=1):
        # Check for task list checkbox
        task_checked = None
        checkbox = li.find(".//input[@type='checkbox']")
        if checkbox is not None:
            task_checked = checkbox.get("checked") is not None

        para = Paragraph(
            runs=_extract_list_item_runs(li),
            list_type=list_type,
            list_index=idx if list_type == "ordered" else None,
            task_checked=task_checked
        )
        paragraphs.append(para)

        # Handle nested lists
        for nested_ul in li.findall("./ul"):
            paragraphs.extend(_parse_list(nested_ul, "bullet"))
        for nested_ol in li.findall("./ol"):
            paragraphs.extend(_parse_list(nested_ol, "ordered"))

    return paragraphs


def parse_markdown_to_nodes(markdown: str) -> list[ContentNode]:
    """Parse Markdown to a list of structured content nodes.

    This is the main entry point for document generation. It renders Markdown
    to HTML using GFM rules, then traverses the HTML to extract structured
    content suitable for generating Word, Excel, or PowerPoint documents.

    Args:
        markdown: GitHub Flavored Markdown content

    Returns:
        List of ContentNode objects (Paragraph, Table, CodeBlock, HorizontalRule)
    """
    html = render_markdown(markdown)
    root = parse_html(html)

    if root is None:
        return []

    nodes: list[ContentNode] = []

    for elem in root:
        tag = elem.tag.lower() if isinstance(elem.tag, str) else ""

        # Headings
        if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(tag[1])
            nodes.append(Paragraph(runs=_extract_runs(elem), level=level))

        # Paragraphs
        elif tag == "p":
            nodes.append(Paragraph(runs=_extract_runs(elem)))

        # Lists
        elif tag == "ul":
            nodes.extend(_parse_list(elem, "bullet"))
        elif tag == "ol":
            nodes.extend(_parse_list(elem, "ordered"))

        # Tables
        elif tag == "table":
            nodes.append(_parse_table(elem))

        # Code blocks
        elif tag == "pre":
            code_elem = elem.find(".//code")
            if code_elem is not None:
                # Extract language from class (e.g., "language-python")
                classes = code_elem.get("class", "")
                lang = None
                for cls in classes.split():
                    if cls.startswith("language-"):
                        lang = cls[9:]
                        break
                nodes.append(CodeBlock(code=code_elem.text or "", language=lang))
            else:
                nodes.append(CodeBlock(code=elem.text or ""))

        # Horizontal rules
        elif tag == "hr":
            nodes.append(HorizontalRule())

        # Blockquotes - extract content as paragraphs
        elif tag == "blockquote":
            for child in elem:
                child_tag = child.tag.lower() if isinstance(child.tag, str) else ""
                if child_tag == "p":
                    nodes.append(Paragraph(runs=_extract_runs(child)))

    return nodes


def extract_tables_from_markdown(markdown: str) -> list[Table]:
    """Extract only tables from Markdown content.

    Convenience function for Excel generation.

    Args:
        markdown: GitHub Flavored Markdown content

    Returns:
        List of Table objects
    """
    nodes = parse_markdown_to_nodes(markdown)
    return [node for node in nodes if isinstance(node, Table)]
