# Office Document MCP Server

This is a sanitized version of a prototype Model Context Protocol (MCP) server providing tools to extract, convert, and generate Microsoft Office documents (Word, Excel, PowerPoint) that I developed in parallel with [`go-ooxml`](https://github.com/rcarmo/go-ooxml). 

Since the server itself is pretty generic code and does not support enterprise features like Information Rights Management (which I will eventually tackle with a brand new implementation atop [`go-ooxml`](https://github.com/rcarmo/go-ooxml)), I have decided to carve it out into a standalone repository to have its own CI/CD workflows and making it easier to install via `uv`.

The code is considered _stable_, so it will not be maintained other than patches/hotfixes and there is _zero_ support or issue tracking.

## Available Tools

### Unified Tools (Primary Interface)

These 9 tools auto-detect document format from file extension or provide cross-format workflow guidance:

| Tool | Description |
|------|-------------|
| `office_help` | Structured workflow help and recommendations for consulting/architecture document workflows |
| `office_read` | Read content from Word/Excel/PowerPoint as JSON or Markdown |
| `office_inspect` | Get document structure (sheets, slides, sections, tables, comments) |
| `office_patch` | Edit cells, shapes, sections, or replace placeholders |
| `office_comment` | Add or get comments from any document type |
| `office_table` | Table operations: add rows, create tables, add bullets |
| `office_template` | Copy templates or analyze template structure |
| `office_audit` | Audit for placeholders, completion, or tracking status |
| `office_image` | Insert images into Word, Excel, or PowerPoint documents |

### Specialized Tools

#### Word SOW Generation

These were a proof-of-concept approach fod managing and updating specific document templates - all the tools marked `sow` are deprecated and kept only for historical interest.

| Tool | Description |
|------|-------------|
| `word_generate_sow` | Fill SOW template with structured data |
| `word_cleanup_sow` | Remove template artifacts and guidance (tracked) |
| `word_get_section_guidance` | Extract template instructions from a section |
| `word_parse_sow_template` | Analyze SOW template structure |
| `word_create_sow_from_markdown` | Create SOW from Markdown content |
| `word_extract_sow_structure` | Extract structured data from existing SOW |
| `word_insert_at_anchor` | Insert paragraphs before/after an anchor paragraph or paragraph index |
| `word_enable_track_changes` | Enable Word's track changes mode |
| `word_patch_with_track_changes` | Replace text with revision marks |

#### PowerPoint Slide Management

| Tool | Description |
|------|-------------|
| `pptx_add_slide` | Add new slide with specified layout |
| `pptx_delete_slide` | Remove a slide |
| `pptx_duplicate_slide` | Copy a slide |
| `pptx_reorder_slides` | Change slide order |
| `pptx_hide_slide` | Hide/unhide a slide |
| `pptx_set_notes` | Set speaker notes |
| `pptx_recommend_layout` | Get best layout for content type |
| `pptx_log_changes` | Add change log slide |

#### Document Conversion

| Tool | Description |
|------|-------------|
| `word_from_markdown` | Create Word document from Markdown (supports inline text or `markdown_file` path for large inputs) |
| `excel_from_markdown` | Create Excel workbook from Markdown tables (supports inline text or `markdown_file` for large inputs) |
| `pptx_from_markdown` | Create PowerPoint from Markdown slides (supports inline text or `markdown_file` for large inputs) |

#### Utility

| Tool | Description |
|------|-------------|
| `restart_server` | Hot-reload the server after code changes |
| `list_supported_formats` | Show available document formats |

## Quick Examples

### Workflow Discovery

```python
# Find the best workflow for filling a consulting SOW from markdown
office_help(
  goal="fill_sow_from_markdown",
  document_type="word",
  constraints=["preserve_template_structure"],
  format="summary"
)

# Map a common consulting request onto a deterministic workflow
office_help(
  task="Patch an Excel estimate workbook safely and verify the result",
  format="detailed"
)

# Discover the safest path for a stakeholder review deck
office_help(
  goal="create_review_deck",
  document_type="powerpoint",
  format="summary"
)
```

### Template Analysis Cache

Word template metadata is cached on disk as JSON to avoid re-scanning the same template on every analysis call.

- default cache location: `.office-metadata-cache/`
- override with: `OFFICE_MCP_METADATA_CACHE_DIR`
- invalidation uses: resolved path + file size + mtime
- current cache-backed flow: `word_parse_sow_template` and `office_template(operation="analyze")`

### Reading Documents

```python
# Read Excel as Markdown
office_read(file_path="data.xlsx", output_format="markdown")

# Read specific range
office_read(file_path="data.xlsx", scope="Sheet1!A1:D10")

# Read a single worksheet
office_read(file_path="data.xlsx", scope="Sheet1")

# Read Word document
office_read(file_path="report.docx", output_format="markdown")
```

### Inspecting Structure

```python
# List Excel sheets
office_inspect(file_path="data.xlsx", what="sheets")

# List Word tables
office_inspect(file_path="report.docx", what="tables")

# List PowerPoint slides
office_inspect(file_path="deck.pptx", what="slides")

# Analyze a Word template and reuse cached metadata on later runs
office_template(
  source_path="templates/sow.docx",
  destination_path="",
  operation="analyze"
)
```

### Editing Content

```python
# Patch Excel cell
office_patch(
  file_path="data.xlsx",
  changes=[{"target": "A1", "value": "New Value"}]
)

# Patch Word placeholder
office_patch(
  file_path="report.docx",
  changes=[{"target": "<Customer>", "value": "Contoso"}]
)

# Patch PowerPoint shape
office_patch(
  file_path="deck.pptx",
  changes=[{"target": "slide:1/Title 1", "value": "New Title"}]
)

# PowerPoint soft return in a single text box
office_patch(
  file_path="deck.pptx",
  changes=[{"target": "slide:1/Title 2", "value": "Contoso{br}Project"}]
)
```

Mutation tools now expose a common diagnostics shape for covered Word/Excel workflows:

- `success`
- `status` (`success`, `partial_success`, `failed`, `skipped`)
- `warnings`
- `matched_targets`
- `unmatched_targets`
- `skipped_targets`
- `diagnostics`
- `next_tools`

That makes partial success and recovery paths explicit instead of relying on generic success messages.

### Table Operations

```python
# Add row to Word table
office_table(
  file_path="report.docx",
  operation="add_row",
  table_id="staffing",
  data={"Role": "PM", "Count": "1", "Notes": "Lead"}
)

# Create table in Word
office_table(
  file_path="report.docx",
  operation="create",
  data={
    "headers": ["Phase", "Owner", "Target Date"],
    "rows": [{"Phase": "Discovery", "Owner": "PM", "Target Date": "2026-04-01"}],
    "insert_after_section": "Delivery Plan"
  }
)

# Add table to PowerPoint
office_table(
  file_path="deck.pptx",
  operation="create",
  table_id="3",
  data={
    "headers": ["Phase", "Duration"],
    "rows": [["Discovery", "2 weeks"]]
  }
)
```

### Large Markdown Inputs

```python
# Avoid MCP argument-size limits by passing a markdown_file path
word_from_markdown(
  output_path="report.docx",
  markdown_file="inputs/large-report.md"
)

excel_from_markdown(
  output_path="budget.xlsx",
  markdown_file="inputs/budget-tables.md"
)

pptx_from_markdown(
  output_path="deck.pptx",
  markdown_file="inputs/deck.md"
)

word_create_sow_from_markdown(
  output_path="sow.docx",
  template_path="templates/Agile.docx",
  markdown_file="inputs/sow.md"
)
```

### Auditing

```python
# Audit for placeholders
office_audit(file_path="report.docx", checks=["placeholders"])

# Audit for completion
office_audit(file_path="report.docx", checks=["completion"])
```

## Word Review Workflow

The recommended workflow for reviewing documents is this:

```
1. office_help(goal="fill_sow_from_markdown") → Choose the workflow and recovery path first
2. office_template(operation="copy")          → Create working document from template
3. office_template(operation="analyze")       → Understand what to preserve vs fill
4. office_inspect(what="tables")              → Get EXACT column names for all tables
5. word_generate_sow                           → Fill placeholders and tables with data
6. office_patch(operation="section")          → Add prose to Introduction, Business Context
7. office_table(operation="insert_row")       → Add engagement-specific rows to tables
8. office_patch(operation="fix_split")        → Replace any remaining split placeholders
9. office_comment(operation="add")            → Add review comments for stakeholders
10. word_cleanup_sow                           → Remove template guidance (tracked)
11. office_audit(checks=["completion"])       → Verify completion score ≥ 80%
```

> **Quality Bar:** All review tools preserve document structure by editing templates rather than creating new documents from scratch. All changes are tracked for stakeholder review.

## Setup

### Install from repository (recommended)

Using [uv](https://docs.astral.sh/uv/):

```bash
uv pip install "git+https://github.com/rcarmo/python-office-mcp-server.git"
```

Using pip:

```bash
pip install "git+https://github.com/rcarmo/python-office-mcp-server.git"
```

This installs the `office-mcp-server` command. Requires Python ≥3.10 (tested with 3.12).

### Install from local clone

```bash
git clone https://github.com/rcarmo/python-office-mcp-server.git
cd python-office-mcp-server
uv pip install .
# or: pip install .
# or for development: pip install -e .
```

### Run without installing

```bash
git clone https://github.com/rcarmo/python-office-mcp-server.git
cd python-office-mcp-server
pip install -r requirements.txt
python office_server.py
```

### MCP client configuration

**VS Code** (`.vscode/mcp.json`):

```json
{
  "servers": {
    "office": {
      "command": "office-mcp-server"
    }
  }
}
```

**Claude Desktop** (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "office": {
      "command": "office-mcp-server"
    }
  }
}
```

**mcp-cli** (`mcp_servers.json`):

```json
{
  "mcpServers": {
    "office": {
      "command": "office-mcp-server"
    }
  }
}
```

If using `uvx` or `bunx` instead of a pre-installed binary:

```json
{
  "command": "uvx",
  "args": ["--from", "git+https://github.com/rcarmo/python-office-mcp-server.git", "office-mcp-server"]
}
```

## Tool Activation Note

Some MCP clients can start with subsets of tools disabled by policy/session settings.
If a call returns a disabled-tool error, enable the corresponding MCP tools in the client first, then retry.
This enablement behavior is controlled by the MCP client/host, not by this server.

### VS Code (Automatic)

The server can be set to be auto-discovered from `.vscode/mcp.json`. That is left as an exercise to the reader, but to verify: Open Command Palette → **MCP: List Servers** → confirm `officeServer` is listed.

### GitHub Copilot CLI

Add the server to your Copilot CLI configuration:

```bash
# Open config file
code ~/.config/github-copilot/config.json

# Add this to the mcpServers section:
{
  "mcpServers": {
    "officeServer": {
      "command": "python",
      "args": ["/path/to/.github/mcp/office_server.py"]
    }
  }
}
```

### Running Manually

```bash
cd .github/mcp
pip install -r requirements.txt
python office_server.py
```

## Windows Single-File Distribution

Build a standalone `.exe` using PyInstaller.

### Build on Windows

```powershell
cd .github/mcp
python -m pip install -r requirements.txt
python -m pip install -r requirements-build.txt
python build_windows_onefile.py --clean
```

Output artifact:

- `dist/office-mcp-server.exe`

### Custom output name

```powershell
python build_windows_onefile.py --name office-server-prod
```

### Run executable

```powershell
dist\office-mcp-server.exe
```

Use the generated executable in MCP client configuration by pointing `command` to the `.exe` path.

## Dependencies

- `python-docx` — Word document handling
- `openpyxl` — Excel workbook handling  
- `python-pptx` — PowerPoint presentation handling
- `aioumcp` — Async MCP server framework
- `pyinstaller` — Build-time dependency for one-file Windows executable

## Architecture

The server dynamically loads tool modules from `tools/`:

- `office_unified_tools.py` — Unified interface (7 tools)
- `word_tools.py` — Word conversion tools
- `word_advanced_tools.py` — SOW-specific tools
- `excel_tools.py` — Excel conversion tools
- `excel_advanced_tools.py` — Excel advanced operations (internal)
- `pptx_tools.py` — PowerPoint conversion tools
- `pptx_advanced_tools.py` — Slide management tools

Tools are discovered automatically by class name pattern (`*Tools`).
