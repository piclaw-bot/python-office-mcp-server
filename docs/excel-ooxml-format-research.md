# Excel OOXML Format Research (XLSX/XLSM)

## Executive Summary

This document provides comprehensive research on Excel Open XML formats (XLSX and XLSM), based on the ECMA-376 standard and practical implementation with openpyxl. This research supports the development of advanced Excel tooling for the MCP Office Server.

**Key Findings:**
- XLSX/XLSM files are ZIP archives containing XML parts following the SpreadsheetML vocabulary
- XLSM differs from XLSX only by including a VBA binary part for macros
- Workbook structure uses sheets, named ranges, tables (ListObjects), and defined names
- openpyxl provides comprehensive support for most features except VBA and some chart types
- Excel does NOT support OOXML-level track changes like Word; change tracking is application-specific

---

## 1. Package Structure

### 1.1 File Extensions

| Extension | Description | Content Type |
|-----------|-------------|--------------|
| `.xlsx` | Standard workbook (no macros) | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| `.xlsm` | Macro-enabled workbook | `application/vnd.ms-excel.sheet.macroEnabled.12` |
| `.xltx` | Template (no macros) | `application/vnd.openxmlformats-officedocument.spreadsheetml.template` |
| `.xltm` | Macro-enabled template | `application/vnd.ms-excel.template.macroEnabled.12` |

### 1.2 Package Contents (ZIP Archive)

```
myworkbook.xlsx
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── docProps/
│   ├── app.xml
│   └── core.xml
└── xl/
    ├── _rels/
    │   └── workbook.xml.rels
    ├── workbook.xml
    ├── styles.xml
    ├── sharedStrings.xml
    ├── theme/
    │   └── theme1.xml
    ├── worksheets/
    │   ├── sheet1.xml
    │   ├── sheet2.xml
    │   └── _rels/
    │       └── sheet1.xml.rels
    ├── tables/
    │   └── table1.xml
    ├── drawings/
    │   └── drawing1.xml
    ├── charts/
    │   └── chart1.xml
    ├── comments1.xml
    ├── printerSettings/
    └── calcChain.xml
```

### 1.3 XLSM Additional Parts

XLSM adds:
```
xl/
├── vbaProject.bin          # Binary VBA project
└── _rels/
    └── vbaProject.bin.rels
```

---

## 2. Core XML Parts

### 2.1 Workbook (xl/workbook.xml)

The central part defining sheets, names, and workbook-level settings:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27231"/>
  <workbookPr defaultThemeVersion="166925"/>
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"
                  activeTab="0"/>
  </bookViews>
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Data" sheetId="2" r:id="rId2"/>
    <sheet name="Hidden" sheetId="3" state="hidden" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="ProjectName">Summary!$B$2</definedName>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Summary!$A$1:$G$50</definedName>
    <definedName name="DataRange">Data!$A$1:$D$100</definedName>
  </definedNames>
  <calcPr calcId="191029"/>
</workbook>
```

#### Element: `sheet`

| Attribute | Required | Type | Description |
|-----------|----------|------|-------------|
| `name` | Yes | string | Sheet tab name (max 31 chars) |
| `sheetId` | Yes | integer | Unique internal ID |
| `r:id` | Yes | string | Relationship ID to sheet part |
| `state` | No | enum | `visible`, `hidden`, `veryHidden` |

#### Element: `definedName`

| Attribute | Required | Type | Description |
|-----------|----------|------|-------------|
| `name` | Yes | string | Name identifier |
| `localSheetId` | No | integer | Scope to specific sheet |
| `hidden` | No | boolean | Hide from Name Manager |
| `comment` | No | string | Description text |

**Built-in Names (prefixed with `_xlnm.`):**
- `_xlnm.Print_Area` — Print area
- `_xlnm.Print_Titles` — Print titles (repeating rows/columns)
- `_xlnm._FilterDatabase` — AutoFilter range
- `_xlnm.Extract` — Advanced filter output range

---

### 2.2 Worksheet (xl/worksheets/sheetN.xml)

Individual sheet content including cells, formatting, and features:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr>
    <tabColor rgb="FF0070C0"/>
    <pageSetUpPr fitToPage="1"/>
  </sheetPr>
  <dimension ref="A1:Z100"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="B5" sqref="B5"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43"/>
  <cols>
    <col min="1" max="1" width="25" customWidth="1"/>
    <col min="2" max="4" width="15" style="1" customWidth="1"/>
  </cols>
  <sheetData>
    <!-- Row and cell data -->
  </sheetData>
  <sheetProtection password="ABCD" sheet="1" objects="1" scenarios="1"/>
  <autoFilter ref="A1:D100"/>
  <mergeCells count="2">
    <mergeCell ref="A1:D1"/>
    <mergeCell ref="A10:B10"/>
  </mergeCells>
  <conditionalFormatting sqref="C2:C100">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>0</formula>
    </cfRule>
  </conditionalFormatting>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" showDropDown="0" sqref="B2:B100">
      <formula1>"Option1,Option2,Option3"</formula1>
    </dataValidation>
  </dataValidations>
  <hyperlinks>
    <hyperlink ref="A5" r:id="rId1" display="Click here"/>
  </hyperlinks>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup paperSize="9" orientation="landscape" r:id="rId2"/>
  <tableParts count="1">
    <tablePart r:id="rId3"/>
  </tableParts>
</worksheet>
```

---

### 2.3 Sheet Data (Cells)

Cell content within `<sheetData>`:

```xml
<sheetData>
  <row r="1" spans="1:4" ht="20" customHeight="1">
    <c r="A1" s="1" t="s">
      <v>0</v>
    </c>
    <c r="B1" s="2">
      <v>1234.56</v>
    </c>
    <c r="C1" s="3">
      <f>B1*1.1</f>
      <v>1358.016</v>
    </c>
    <c r="D1" t="inlineStr">
      <is><t>Inline text</t></is>
    </c>
  </row>
  <row r="2" spans="1:4" hidden="1">
    <!-- Hidden row -->
  </row>
</sheetData>
```

#### Element: `c` (Cell)

| Attribute | Required | Type | Description |
|-----------|----------|------|-------------|
| `r` | Yes | string | Cell reference (e.g., `A1`) |
| `s` | No | integer | Style index (from styles.xml) |
| `t` | No | enum | Cell type |

#### Cell Types (`t` attribute)

| Value | Description | Value Element |
|-------|-------------|---------------|
| `b` | Boolean | `<v>0</v>` or `<v>1</v>` |
| `d` | Date (ISO 8601) | `<v>2026-01-22</v>` |
| `e` | Error | `<v>#REF!</v>` |
| `inlineStr` | Inline string | `<is><t>text</t></is>` |
| `n` | Number (default) | `<v>123.45</v>` |
| `s` | Shared string | `<v>5</v>` (index) |
| `str` | Formula string | `<f>A1&B1</f>` |

#### Element: `f` (Formula)

| Attribute | Description |
|-----------|-------------|
| `t` | Formula type: `normal`, `array`, `dataTable`, `shared` |
| `ref` | Array/shared formula range |
| `si` | Shared formula index |
| `ca` | Calculate always (volatile) |

```xml
<!-- Normal formula -->
<c r="C1"><f>SUM(A1:B1)</f><v>100</v></c>

<!-- Array formula (CSE) -->
<c r="D1"><f t="array" ref="D1:D10">SUM(A1:A10*B1:B10)</f><v>500</v></c>

<!-- Shared formula (master) -->
<c r="E1"><f t="shared" ref="E1:E10" si="0">A1+B1</f><v>10</v></c>
<!-- Shared formula (reference) -->
<c r="E2"><f t="shared" si="0"/><v>20</v></c>
```

---

### 2.4 Shared Strings (xl/sharedStrings.xml)

String deduplication for text values:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="150" uniqueCount="45">
  <si><t>Customer Name</t></si>
  <si><t>Order Date</t></si>
  <si>
    <r>
      <t>Regular text </t>
    </r>
    <r>
      <rPr><b/><sz val="12"/><color rgb="FFFF0000"/></rPr>
      <t>with bold red</t>
    </r>
  </si>
</sst>
```

| Element | Description |
|---------|-------------|
| `<si>` | String item |
| `<t>` | Plain text |
| `<r>` | Rich text run |
| `<rPr>` | Run properties (formatting) |

---

### 2.5 Styles (xl/styles.xml)

Centralized formatting definitions:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00\ &quot;USD&quot;"/>
  </numFmts>
  <fonts count="2">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
  </cellXfs>
  <dxfs count="1">
    <!-- Differential formats for conditional formatting -->
    <dxf>
      <fill><patternFill><bgColor rgb="FF92D050"/></patternFill></fill>
    </dxf>
  </dxfs>
</styleSheet>
```

#### Style Index Resolution

Cell `s="1"` → `cellXfs[1]` → applies `numFmtId="164"`, `fontId="1"`

---

## 3. Tables (ListObjects)

### 3.1 Table Definition (xl/tables/tableN.xml)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="StaffingTable" displayName="StaffingTable"
       ref="A1:E20" totalsRowShown="0" headerRowCount="1">
  <autoFilter ref="A1:E20"/>
  <tableColumns count="5">
    <tableColumn id="1" name="Role"/>
    <tableColumn id="2" name="Rate" dataDxfId="1"/>
    <tableColumn id="3" name="Hours"/>
    <tableColumn id="4" name="Total" dataDxfId="0">
      <calculatedColumnFormula>[@Rate]*[@Hours]</calculatedColumnFormula>
    </tableColumn>
    <tableColumn id="5" name="Notes"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0"
                  showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>
```

#### Element: `table`

| Attribute | Description |
|-----------|-------------|
| `id` | Unique table ID in workbook |
| `name` | Internal name |
| `displayName` | Name shown to users |
| `ref` | Table range including headers |
| `headerRowCount` | Number of header rows (usually 1) |
| `totalsRowShown` | Show totals row |

#### Structured References

Tables support structured references in formulas:
- `[@ColumnName]` — Current row, specific column
- `[ColumnName]` — Entire column (data only)
- `[#Headers]` — Header row
- `[#Totals]` — Totals row
- `[#All]` — Entire table including headers

---

## 4. Comments

### 4.1 Legacy Comments (Notes)

**Location:** `/xl/comments1.xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>John Doe</author>
    <author>Jane Smith</author>
  </authors>
  <commentList>
    <comment ref="B5" authorId="0" shapeId="0">
      <text>
        <r>
          <rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr>
          <t>John Doe:</t>
        </r>
        <r>
          <rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr>
          <t xml:space="preserve">
This is a comment about cell B5.
It can span multiple lines.</t>
        </r>
      </text>
    </comment>
  </commentList>
</comments>
```

### 4.2 Threaded Comments (Modern Excel)

**Location:** `/xl/threadedComments/threadedComment1.xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="D10" id="{guid-1}" personId="{person-guid}"
                   dT="2026-01-22T10:30:00.000">
    <text>What's the status of this item?</text>
  </threadedComment>
  <threadedComment ref="D10" id="{guid-2}" personId="{person-guid-2}"
                   dT="2026-01-22T11:00:00.000" parentId="{guid-1}">
    <text>Still waiting on customer feedback.</text>
  </threadedComment>
</ThreadedComments>
```

#### Persons Part

**Location:** `/xl/persons/person.xml`

```xml
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="{person-guid}" displayName="John Doe" userId="user@contoso.com"/>
  <person id="{person-guid-2}" displayName="Jane Smith" userId="jane@contoso.com"/>
</personList>
```

---

## 5. Data Validation

### 5.1 Validation Types

```xml
<dataValidations count="3">
  <!-- Dropdown list from explicit values -->
  <dataValidation type="list" allowBlank="1" sqref="B2:B100">
    <formula1>"Yes,No,N/A"</formula1>
  </dataValidation>

  <!-- Dropdown from named range -->
  <dataValidation type="list" allowBlank="1" sqref="C2:C100">
    <formula1>StatusList</formula1>
  </dataValidation>

  <!-- Numeric range -->
  <dataValidation type="decimal" operator="between" sqref="D2:D100">
    <formula1>0</formula1>
    <formula2>100</formula2>
  </dataValidation>

  <!-- Date validation -->
  <dataValidation type="date" operator="greaterThan" sqref="E2:E100">
    <formula1>TODAY()</formula1>
  </dataValidation>

  <!-- Custom formula -->
  <dataValidation type="custom" sqref="F2:F100">
    <formula1>LEN(F2)&lt;=50</formula1>
  </dataValidation>
</dataValidations>
```

| Attribute | Values | Description |
|-----------|--------|-------------|
| `type` | `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom` | Validation type |
| `operator` | `between`, `notBetween`, `equal`, `notEqual`, `lessThan`, `lessThanOrEqual`, `greaterThan`, `greaterThanOrEqual` | Comparison |
| `allowBlank` | `0`, `1` | Allow empty cells |
| `showDropDown` | `0`, `1` | Show dropdown (confusingly, `0` = show, `1` = hide) |
| `showErrorMessage` | `0`, `1` | Show error on invalid |
| `showInputMessage` | `0`, `1` | Show prompt when selected |

---

## 6. Conditional Formatting

### 6.1 Rule Types

```xml
<conditionalFormatting sqref="C2:C100">
  <!-- Cell value comparison -->
  <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
    <formula>0</formula>
  </cfRule>

  <!-- Color scale (gradient) -->
  <cfRule type="colorScale" priority="2">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="percentile" val="50"/>
      <cfvo type="max"/>
      <color rgb="FFF8696B"/>
      <color rgb="FFFFEB84"/>
      <color rgb="FF63BE7B"/>
    </colorScale>
  </cfRule>

  <!-- Data bar -->
  <cfRule type="dataBar" priority="3">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF638EC6"/>
    </dataBar>
  </cfRule>

  <!-- Icon set -->
  <cfRule type="iconSet" priority="4">
    <iconSet iconSet="3TrafficLights1">
      <cfvo type="percent" val="0"/>
      <cfvo type="percent" val="33"/>
      <cfvo type="percent" val="67"/>
    </iconSet>
  </cfRule>

  <!-- Formula-based -->
  <cfRule type="expression" dxfId="1" priority="5">
    <formula>$A2="Important"</formula>
  </cfRule>

  <!-- Duplicate values -->
  <cfRule type="duplicateValues" dxfId="2" priority="6"/>

  <!-- Top N -->
  <cfRule type="top10" dxfId="3" priority="7" rank="10"/>
</conditionalFormatting>
```

---

## 7. Protection

### 7.1 Sheet Protection

```xml
<sheetProtection algorithmName="SHA-512" hashValue="base64hash..."
                 saltValue="base64salt..." spinCount="100000"
                 sheet="1" objects="1" scenarios="1"
                 selectLockedCells="0" selectUnlockedCells="0"
                 formatCells="0" formatColumns="0" formatRows="0"
                 insertColumns="0" insertRows="0" insertHyperlinks="0"
                 deleteColumns="0" deleteRows="0"
                 sort="0" autoFilter="0" pivotTables="0"/>
```

| Attribute | Description |
|-----------|-------------|
| `sheet` | Protect sheet structure |
| `objects` | Protect drawing objects |
| `scenarios` | Protect scenarios |
| `formatCells` | Allow cell formatting (1=allow, 0=protect) |
| `insertRows` | Allow row insertion |

### 7.2 Workbook Protection

In `workbook.xml`:

```xml
<workbookProtection workbookAlgorithmName="SHA-512"
                    workbookHashValue="base64hash..."
                    workbookSaltValue="base64salt..."
                    workbookSpinCount="100000"
                    lockStructure="1" lockWindows="0"/>
```

### 7.3 Cell-Level Protection

Via cell format (xf):

```xml
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"
    applyProtection="1">
  <protection locked="0" hidden="0"/>
</xf>
```

---

## 8. Key Constants and URIs

### Namespace URIs

| Prefix | Namespace URI |
|--------|---------------|
| (default) | `http://schemas.openxmlformats.org/spreadsheetml/2006/main` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| `mc` | `http://schemas.openxmlformats.org/markup-compatibility/2006` |
| `x14` | `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main` |
| `x14ac` | `http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac` |
| `xr` | `http://schemas.microsoft.com/office/spreadsheetml/2014/revision` |

### Content Types

```python
CT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_WORKBOOK_MACRO = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
CT_SHARED_STRINGS = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
CT_VBA = "application/vnd.ms-office.vbaProject"
```

### Relationship Types

```python
RT_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
RT_SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
RT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
RT_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
RT_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
```

---

## 9. Python Implementation with openpyxl

### 9.1 Basic Operations

```python
from openpyxl import Workbook, load_workbook

# Load workbook
wb = load_workbook("template.xlsx", data_only=False)  # Keep formulas
# wb = load_workbook("template.xlsx", data_only=True)  # Computed values only

# Access sheets
ws = wb.active
ws = wb["SheetName"]
ws = wb.worksheets[0]

# List all sheets
for name in wb.sheetnames:
    print(name)

# Cell operations
cell = ws["A1"]
cell = ws.cell(row=1, column=1)
cell.value = "New value"
cell.value = 123.45
cell.value = "=SUM(B1:B10)"

# Read value vs formula
print(cell.value)  # Formula: "=SUM(B1:B10)"
# Use data_only=True to get computed value

# Range operations
for row in ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=5):
    for cell in row:
        print(cell.value)

# Save
wb.save("output.xlsx")
```

### 9.2 Working with Named Ranges

```python
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName

wb = load_workbook("workbook.xlsx")

# List all defined names
for name in wb.defined_names.values():
    print(f"{name.name}: {name.attr_text}")

# Get specific named range
if "ProjectName" in wb.defined_names:
    defn = wb.defined_names["ProjectName"]
    print(f"Refers to: {defn.attr_text}")  # e.g., "Summary!$B$2"

# Create named range
new_name = DefinedName("DataRange", attr_text="Data!$A$1:$D$100")
wb.defined_names.add(new_name)

# Workbook-scoped vs sheet-scoped
sheet_scope_name = DefinedName("LocalName", attr_text="Sheet1!$A$1", localSheetId=0)
wb.defined_names.add(sheet_scope_name)
```

### 9.3 Working with Tables

```python
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = load_workbook("workbook.xlsx")
ws = wb.active

# List tables in sheet
for table in ws.tables.values():
    print(f"{table.name}: {table.ref}")

# Get table by name
table = ws.tables["StaffingTable"]
print(f"Range: {table.ref}")
print(f"Headers: {[col.name for col in table.tableColumns]}")

# Create new table
new_table = Table(displayName="NewTable", ref="A1:D10")
style = TableStyleInfo(
    name="TableStyleMedium2",
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False
)
new_table.tableStyleInfo = style
ws.add_table(new_table)

# Expand table (add rows)
# Note: Must update table.ref manually
```

### 9.4 Comments

```python
from openpyxl import load_workbook
from openpyxl.comments import Comment

wb = load_workbook("workbook.xlsx")
ws = wb.active

# Add comment
comment = Comment("This is a note", "John Doe")
comment.width = 200  # pixels
comment.height = 100
ws["A1"].comment = comment

# Read comments
for row in ws.iter_rows():
    for cell in row:
        if cell.comment:
            print(f"{cell.coordinate}: {cell.comment.text} by {cell.comment.author}")

# Remove comment
ws["A1"].comment = None
```

### 9.5 Data Validation

```python
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

wb = load_workbook("workbook.xlsx")
ws = wb.active

# Dropdown from list
dv = DataValidation(
    type="list",
    formula1='"Yes,No,N/A"',
    allow_blank=True
)
ws.add_data_validation(dv)
dv.add("B2:B100")

# Dropdown from named range
dv2 = DataValidation(type="list", formula1="StatusList")
ws.add_data_validation(dv2)
dv2.add("C2:C100")

# Number range
dv3 = DataValidation(
    type="decimal",
    operator="between",
    formula1="0",
    formula2="100"
)
ws.add_data_validation(dv3)
dv3.add("D2:D100")

# With error message
dv.error = "Please select a valid option"
dv.errorTitle = "Invalid Input"
dv.showErrorMessage = True
```

### 9.6 Conditional Formatting

```python
from openpyxl import load_workbook
from openpyxl.formatting.rule import (
    ColorScaleRule, DataBarRule, IconSetRule,
    FormulaRule, CellIsRule
)

wb = load_workbook("workbook.xlsx")
ws = wb.active

# Cell value rule
rule = CellIsRule(
    operator='greaterThan',
    formula=['0'],
    fill=PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
)
ws.conditional_formatting.add("C2:C100", rule)

# Color scale (red-yellow-green)
rule2 = ColorScaleRule(
    start_type='min', start_color='F8696B',
    mid_type='percentile', mid_value=50, mid_color='FFEB84',
    end_type='max', end_color='63BE7B'
)
ws.conditional_formatting.add("D2:D100", rule2)

# Formula-based rule
rule3 = FormulaRule(
    formula=['$A2="Important"'],
    fill=PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
)
ws.conditional_formatting.add("A2:E100", rule3)
```

### 9.7 Sheet Operations

```python
from openpyxl import load_workbook
import copy

wb = load_workbook("workbook.xlsx")

# Copy sheet
source = wb["Template"]
target = wb.copy_worksheet(source)
target.title = "NewSheet"

# Hide/show sheet
ws = wb["HiddenSheet"]
ws.sheet_state = 'hidden'  # 'visible', 'hidden', 'veryHidden'

# Reorder sheets
wb.move_sheet("NewSheet", offset=-1)

# Delete sheet
del wb["OldSheet"]
```

### 9.8 Protection

```python
from openpyxl import load_workbook

wb = load_workbook("workbook.xlsx")
ws = wb.active

# Protect sheet
ws.protection.sheet = True
ws.protection.password = "secret"
ws.protection.formatCells = False  # Allow formatting
ws.protection.insertRows = False   # Allow row insertion

# Unlock specific cells (before protecting)
for row in ws.iter_rows(min_row=2, max_row=100, min_col=2, max_col=4):
    for cell in row:
        cell.protection = Protection(locked=False)

# Protect workbook structure
wb.security.lockStructure = True
wb.security.workbookPassword = "secret"
```

---

## 10. Comparison: XLSX vs DOCX vs PPTX

| Feature | XLSX | DOCX | PPTX |
|---------|------|------|------|
| Primary unit | Cell | Paragraph | Shape |
| Track changes | ❌ Not in OOXML | ✅ `w:ins`/`w:del` | ❌ No |
| Comments | ✅ Cell-anchored | ✅ Range-anchored | ✅ Position-anchored |
| Threaded comments | ✅ Modern Excel | ✅ `commentsExtended` | ✅ Extensions |
| Formulas | ✅ Rich formula engine | ❌ Fields only | ❌ No |
| Tables | ✅ ListObjects | ✅ Tables | ✅ Slide tables |
| Named ranges | ✅ Yes | ❌ Bookmarks | ❌ No |
| Data validation | ✅ Rich validation | ❌ Limited | ❌ No |
| Conditional formatting | ✅ Extensive | ❌ No | ❌ No |
| Protection | ✅ Cell/sheet/workbook | ✅ Document | ✅ Presentation |
| Macros | ✅ VBA (XLSM) | ✅ VBA (DOCM) | ✅ VBA (PPTM) |

---

## 11. Track Changes Alternative

Excel OOXML does not support track changes at the XML level like Word. Instead, alternatives include:

### 11.1 Change Tracking via Comments

Add comments documenting changes:

```python
def add_change_comment(cell, old_value, new_value, author="Agent"):
    """Document a change via cell comment."""
    from openpyxl.comments import Comment
    from datetime import datetime
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    text = f"Changed by {author} at {timestamp}\nOld: {old_value}\nNew: {new_value}"
    
    if cell.comment:
        cell.comment.text += f"\n\n{text}"
    else:
        cell.comment = Comment(text, author)
```

### 11.2 Change Log Sheet

Maintain a separate sheet with change history:

```python
def log_change(wb, sheet_name, cell_ref, old_value, new_value, author="Agent"):
    """Log change to a dedicated sheet."""
    from datetime import datetime
    
    # Create or get log sheet
    if "_ChangeLog" not in wb.sheetnames:
        log_ws = wb.create_sheet("_ChangeLog")
        log_ws.append(["Timestamp", "Sheet", "Cell", "Old Value", "New Value", "Author"])
    else:
        log_ws = wb["_ChangeLog"]
    
    log_ws.append([
        datetime.now().isoformat(),
        sheet_name,
        cell_ref,
        str(old_value),
        str(new_value),
        author
    ])
```

### 11.3 Cell Highlighting

Use conditional formatting or fill to highlight changes:

```python
def highlight_changed_cell(cell):
    """Apply visual indicator for changed cell."""
    from openpyxl.styles import PatternFill
    
    cell.fill = PatternFill(
        start_color="FFFF99",  # Light yellow
        end_color="FFFF99",
        fill_type="solid"
    )
```

---

## 12. XLSM Considerations

### 12.1 VBA Preservation

openpyxl preserves VBA code when using `keep_vba=True`:

```python
from openpyxl import load_workbook

# Load with VBA preservation
wb = load_workbook("template.xlsm", keep_vba=True)

# Make changes...

# Save as .xlsm (extension matters!)
wb.save("output.xlsm")
```

**Important:** 
- Must save as `.xlsm` to retain macros
- Cannot modify VBA code with openpyxl
- VBA is stored as binary blob (`vbaProject.bin`)

### 12.2 Security Considerations

- XLSM files can execute arbitrary code
- Consider extracting VBA for review before executing
- Tool should warn when processing XLSM files

---

## Assumptions

- Research based on ECMA-376 5th Edition and ISO/IEC 29500
- Tested with openpyxl 3.1.x and Microsoft 365 Excel
- XLSM VBA handling limited to preservation only
- Modern threaded comments require Excel 2016+

## Next Steps

- [ ] Implement Phase 1 tools in `excel_advanced_tools.py`
- [ ] Add `excel_list_sheets` for workbook introspection
- [ ] Add `excel_get_range` for targeted data extraction
- [ ] Add `excel_patch_cell` with change logging
- [ ] Add `excel_replace_placeholders` for template filling
- [ ] Add `excel_audit_placeholders` for QA
- [ ] Test with ECIF Request Work Scope template
