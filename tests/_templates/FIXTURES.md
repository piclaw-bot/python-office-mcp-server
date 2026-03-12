# Test Fixtures Checklist

Test fixtures created programmatically using python-docx, openpyxl, and python-pptx libraries with content derived from Mary Shelley's "Frankenstein" (1818, public domain). Author set to "Test Author".

> **Note**: For maximum fidelity testing, these could be replaced with files created in Microsoft Office applications (Word, Excel, PowerPoint), as OOXML files created by Office have subtle differences from programmatically generated ones.

## Online References for Validation

Use these resources to validate our implementation:

- **[officeopenxml.com](http://officeopenxml.com/)** - User-friendly element reference
  - [Document structure](http://officeopenxml.com/WPdocument.php)
  - [Paragraphs](http://officeopenxml.com/WPparagraph.php)
  - [Tables](http://officeopenxml.com/WPtable.php)
  - [Styles](http://officeopenxml.com/WPstyles.php)
- **[ECMA-376 Spec](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)** - Official standard (free PDF)
- **[Microsoft Open XML SDK docs](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)** - API reference with examples

## Word Documents (`testdata/word/`)

- [x] `minimal.docx` - Empty document with just body element
- [x] `single_paragraph.docx` - One paragraph, no formatting
- [x] `formatted_text.docx` - Bold, italic, underline, colors, font sizes
- [x] `headings.docx` - All heading levels 1-9
- [x] `simple_table.docx` - 3x3 table, no merged cells
- [x] `complex_table.docx` - Merged cells, nested tables
- [x] `track_changes.docx` - Document with insertions and deletions tracked
- [x] `comments.docx` - Multiple comments with replies
- [x] `styles.docx` - Custom styles applied
- [x] `headers_footers.docx` - Different first page, odd/even headers/footers
- [x] `sdt_content_controls.docx` - Content controls/placeholders
- [x] `numbered_list.docx` - Numbered list items
- [x] `bullet_list.docx` - Bullet list items

## Excel Workbooks (`testdata/excel/`)

- [x] `minimal.xlsx` - Empty workbook with one sheet
- [x] `single_cell.xlsx` - One cell with a value
- [x] `data_types.xlsx` - String, number, date, boolean, formula cells
- [x] `formatting.xlsx` - Colors, fonts, borders
- [x] `multiple_sheets.xlsx` - Three sheets with data
- [x] `tables.xlsx` - Excel tables (not just cell ranges)
- [x] `merged_cells.xlsx` - Merged cell regions
- [x] `named_ranges.xlsx` - Named ranges defined
- [x] `comments.xlsx` - Cell comments
- [x] `formulas.xlsx` - Various formulas (SUM, VLOOKUP, etc.)
- [x] `conditional_format.xlsx` - Conditional formatting rules

## PowerPoint Presentations (`testdata/pptx/`)

- [x] `minimal.pptx` - Single blank slide
- [x] `title_slide.pptx` - Title layout slide with content
- [x] `bullet_points.pptx` - Slide with bullet points
- [x] `shapes.pptx` - Various shape types (rectangles, arrows, etc.)
- [x] `tables.pptx` - Table on slide
- [x] `images.pptx` - Embedded images (placeholder shapes)
- [x] `notes.pptx` - Slides with speaker notes
- [x] `comments.pptx` - Slide with author metadata
- [x] `hidden_slides.pptx` - Mix of visible and hidden slides
- [x] `multiple_masters.pptx` - Multiple slide layouts used
- [x] `layouts.pptx` - All standard layouts used

---

## Fixture Creation Guidelines

1. **Created using** python-docx, openpyxl, python-pptx
2. **Save as .docx/.xlsx/.pptx** (not .doc/.xls/.ppt)
3. **Keep files minimal** - Include only what's needed to test the specific feature
4. **Content source** - Mary Shelley's "Frankenstein" (public domain)
5. **Author** - "Test Author"

## Total: 35 fixtures created

| Format | Count | Status |
|--------|-------|--------|
| Word   | 13    | 13/13 ✓ |
| Excel  | 11    | 11/11 ✓ |
| PowerPoint | 11 | 11/11 ✓ |
