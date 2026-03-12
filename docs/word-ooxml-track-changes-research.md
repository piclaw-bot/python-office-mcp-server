# Word OOXML Track Changes and Comments Research

## Executive Summary

This document provides comprehensive research on Word OOXML format for track changes (revisions) and comments, based on the ECMA-376 standard and practical implementation testing.

**Key Findings:**
- Word supports full track changes via `w:ins` and `w:del` elements for insertions and deletions
- Comments are supported through `w:comments` and `w:commentRangeStart`/`w:commentRangeEnd` markers
- python-docx has partial support; full tracked changes require manual XML manipulation
- Track changes preserve revision history with author, date, and revision ID

---

## 1. Track Changes XML Structure

### 1.1 Revision Settings (settings.xml)

**Location:** `/word/settings.xml`

To enable track changes, set the `trackRevisions` element:

```xml
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:trackRevisions/>
  <w:revisionView w:markup="1" w:insDel="1" w:formatting="1"/>
</w:settings>
```

| Element | Description |
|---------|-------------|
| `w:trackRevisions` | Enables revision tracking |
| `w:revisionView` | Controls what revisions are visible |

---

### 1.2 Insertions (w:ins)

Inserted content is wrapped in `w:ins` elements:

```xml
<w:p>
  <w:r>
    <w:t>Original text</w:t>
  </w:r>
  <w:ins w:id="0" w:author="John Doe" w:date="2026-01-21T10:30:00Z">
    <w:r>
      <w:t> with inserted content</w:t>
    </w:r>
  </w:ins>
</w:p>
```

#### Attributes

| Attribute | Required | Type | Description |
|-----------|----------|------|-------------|
| `w:id` | Yes | integer | Unique revision ID in the document |
| `w:author` | Yes | string | Name of the person who made the change |
| `w:date` | No | dateTime | When the change was made (ISO 8601) |

---

### 1.3 Deletions (w:del)

Deleted content is wrapped in `w:del` elements but preserved in the document:

```xml
<w:p>
  <w:del w:id="1" w:author="Jane Smith" w:date="2026-01-21T11:00:00Z">
    <w:r>
      <w:delText>This text was deleted</w:delText>
    </w:r>
  </w:del>
  <w:r>
    <w:t>Remaining text</w:t>
  </w:r>
</w:p>
```

**Note:** Deleted text uses `w:delText` instead of `w:t` to distinguish it from active content.

---

### 1.4 Move Operations (w:moveFrom / w:moveTo)

Word tracks moved content with paired elements:

```xml
<!-- Source location -->
<w:moveFrom w:id="2" w:author="John Doe" w:date="2026-01-21T12:00:00Z" w:name="move1">
  <w:r>
    <w:t>Moved paragraph</w:t>
  </w:r>
</w:moveFrom>

<!-- Destination location -->
<w:moveTo w:id="3" w:author="John Doe" w:date="2026-01-21T12:00:00Z" w:name="move1">
  <w:r>
    <w:t>Moved paragraph</w:t>
  </w:r>
</w:moveTo>
```

| Attribute | Description |
|-----------|-------------|
| `w:name` | Links the moveFrom and moveTo pair |

---

### 1.5 Formatting Changes (w:rPrChange / w:pPrChange)

Track changes to character and paragraph formatting:

```xml
<w:r>
  <w:rPr>
    <w:b/>
    <w:rPrChange w:id="4" w:author="John Doe" w:date="2026-01-21T13:00:00Z">
      <w:rPr>
        <!-- Previous formatting - no bold -->
      </w:rPr>
    </w:rPrChange>
  </w:rPr>
  <w:t>Now bold text</w:t>
</w:r>
```

---

## 2. Comments XML Structure

### 2.1 Comments Part (comments.xml)

**Location:** `/word/comments.xml`  
**Content-Type:** `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml`

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="John Doe" w:date="2026-01-21T10:30:00Z" w:initials="JD">
    <w:p>
      <w:r>
        <w:t>This is my comment text</w:t>
      </w:r>
    </w:p>
  </w:comment>
  <w:comment w:id="1" w:author="Jane Smith" w:date="2026-01-21T11:00:00Z" w:initials="JS">
    <w:p>
      <w:r>
        <w:t>Another comment with </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>rich formatting</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>
```

#### Element: `w:comment`

| Attribute | Required | Type | Description |
|-----------|----------|------|-------------|
| `w:id` | Yes | integer | Unique comment ID |
| `w:author` | Yes | string | Comment author name |
| `w:date` | No | dateTime | Creation timestamp |
| `w:initials` | No | string | Author initials |

**Note:** Comments support full paragraph content including rich text formatting.

---

### 2.2 Comment Range Markers (in document.xml)

Comments reference text ranges using start/end markers in the document body:

```xml
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r>
    <w:t>This text has a comment attached</w:t>
  </w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r>
    <w:commentReference w:id="0"/>
  </w:r>
</w:p>
```

| Element | Description |
|---------|-------------|
| `w:commentRangeStart` | Marks where commented text begins |
| `w:commentRangeEnd` | Marks where commented text ends |
| `w:commentReference` | Displays the comment marker (superscript number) |

---

### 2.3 Extended Comments (commentsExtended.xml)

**Location:** `/word/commentsExtended.xml`  
**Content-Type:** `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml`

Stores additional comment metadata for threading and resolution:

```xml
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:commentEx w15:paraId="00000001" w15:done="0"/>
  <w15:commentEx w15:paraId="00000002" w15:done="1" w15:paraIdParent="00000001"/>
</w15:commentsEx>
```

| Attribute | Description |
|-----------|-------------|
| `w15:paraId` | Links to the comment's paragraph |
| `w15:done` | 0 = open, 1 = resolved |
| `w15:paraIdParent` | Parent comment for replies (threading) |

---

## 3. Package Relationships

### 3.1 Document to Comments

Add to `/word/_rels/document.xml.rels`:

```xml
<Relationship 
    Id="rIdN" 
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" 
    Target="comments.xml"/>
```

### 3.2 Content Types

Add to `[Content_Types].xml`:

```xml
<Override 
    PartName="/word/comments.xml" 
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
<Override 
    PartName="/word/commentsExtended.xml" 
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
```

---

## 4. Revision ID Management

### 4.1 Document-Level Revision Tracking

`settings.xml` contains revision tracking configuration:

```xml
<w:settings>
  <w:trackRevisions/>
  <w:rsids>
    <w:rsidRoot w:val="00A77B3E"/>
    <w:rsid w:val="00A77B3E"/>
    <w:rsid w:val="00B88C4F"/>
  </w:rsids>
</w:settings>
```

### 4.2 Run-Level RSIDs

Each run can have revision session IDs:

```xml
<w:r w:rsidR="00A77B3E" w:rsidRPr="00B88C4F">
  <w:t>Text content</w:t>
</w:r>
```

| Attribute | Description |
|-----------|-------------|
| `w:rsidR` | Revision ID for run content |
| `w:rsidRPr` | Revision ID for run properties |
| `w:rsidDel` | Revision ID for deletion |

---

## 5. Key Constants and URIs

### Namespace URIs

| Prefix | Namespace URI |
|--------|---------------|
| `w` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` |
| `w14` | `http://schemas.microsoft.com/office/word/2010/wordml` |
| `w15` | `http://schemas.microsoft.com/office/word/2012/wordml` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |

### Content Types

```python
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_COMMENTS_EXTENDED = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
CT_DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
```

### Relationship Types

```python
RT_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
RT_COMMENTS_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
```

### DateTime Format

ISO 8601 format with timezone: `YYYY-MM-DDTHH:MM:SSZ`

Example: `2026-01-21T10:30:00Z`

---

## 6. XSD Schema Excerpts (ECMA-376)

### CT_TrackChange (Base for ins/del)

```xml
<xsd:complexType name="CT_TrackChange">
  <xsd:attribute name="author" type="s:ST_String" use="required"/>
  <xsd:attribute name="date" type="xsd:dateTime"/>
  <xsd:attribute name="id" type="ST_DecimalNumber" use="required"/>
</xsd:complexType>
```

### CT_RunTrackChange (w:ins, w:del)

```xml
<xsd:complexType name="CT_RunTrackChange">
  <xsd:complexContent>
    <xsd:extension base="CT_TrackChange">
      <xsd:sequence>
        <xsd:group ref="EG_ContentRunContent" minOccurs="0" maxOccurs="unbounded"/>
      </xsd:sequence>
    </xsd:extension>
  </xsd:complexContent>
</xsd:complexType>
```

### CT_Comment

```xml
<xsd:complexType name="CT_Comment">
  <xsd:complexContent>
    <xsd:extension base="CT_TrackChange">
      <xsd:sequence>
        <xsd:group ref="EG_BlockLevelElts" minOccurs="0" maxOccurs="unbounded"/>
      </xsd:sequence>
      <xsd:attribute name="initials" type="s:ST_String"/>
    </xsd:extension>
  </xsd:complexContent>
</xsd:complexType>
```

---

## 7. Python Implementation

### 7.1 Using python-docx for Basic Operations

python-docx has limited built-in support. For simple tracked changes:

```python
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_tracked_insertion(paragraph, text, author="Author"):
    """Add text as a tracked insertion."""
    from datetime import datetime
    
    # Create w:ins element
    ins = OxmlElement('w:ins')
    ins.set(qn('w:id'), '0')
    ins.set(qn('w:author'), author)
    ins.set(qn('w:date'), datetime.now().isoformat() + 'Z')
    
    # Create run inside
    run = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    ins.append(run)
    
    paragraph._p.append(ins)

def add_tracked_deletion(paragraph, text, author="Author"):
    """Add text as a tracked deletion."""
    from datetime import datetime
    
    # Create w:del element
    del_elem = OxmlElement('w:del')
    del_elem.set(qn('w:id'), '1')
    del_elem.set(qn('w:author'), author)
    del_elem.set(qn('w:date'), datetime.now().isoformat() + 'Z')
    
    # Create run with delText
    run = OxmlElement('w:r')
    del_text = OxmlElement('w:delText')
    del_text.text = text
    run.append(del_text)
    del_elem.append(run)
    
    paragraph._p.append(del_elem)
```

### 7.2 Working with Comments

```python
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime

def add_comment(document, paragraph, text, comment_text, author="Author", initials="A"):
    """Add a comment to specific text in a paragraph."""
    
    # Get the comments part (create if needed)
    comments_part = document.part.comments_part
    
    # Create comment range markers in the document
    comment_id = "0"  # Should be unique
    
    # commentRangeStart
    range_start = OxmlElement('w:commentRangeStart')
    range_start.set(qn('w:id'), comment_id)
    
    # commentRangeEnd
    range_end = OxmlElement('w:commentRangeEnd')
    range_end.set(qn('w:id'), comment_id)
    
    # commentReference (shows the marker)
    ref_run = OxmlElement('w:r')
    comment_ref = OxmlElement('w:commentReference')
    comment_ref.set(qn('w:id'), comment_id)
    ref_run.append(comment_ref)
    
    # Insert markers around target text
    paragraph._p.insert(0, range_start)
    paragraph._p.append(range_end)
    paragraph._p.append(ref_run)
    
    # Add comment to comments.xml
    # (This requires access to the comments part)
```

### 7.3 Enabling Track Changes

```python
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def enable_track_changes(document):
    """Enable track changes in the document."""
    settings = document.settings.element
    
    # Check if trackRevisions already exists
    track_revisions = settings.find(qn('w:trackRevisions'))
    if track_revisions is None:
        track_revisions = OxmlElement('w:trackRevisions')
        settings.append(track_revisions)
    
    return document
```

---

## 8. Accepting/Rejecting Changes

### 8.1 Accept Insertion

To accept an insertion, move the content out of `w:ins` and remove the wrapper:

```xml
<!-- Before (tracked) -->
<w:ins w:id="0" w:author="John Doe">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>

<!-- After (accepted) -->
<w:r><w:t>inserted text</w:t></w:r>
```

### 8.2 Accept Deletion

To accept a deletion, remove the entire `w:del` element and its contents.

### 8.3 Reject Insertion

To reject an insertion, remove the entire `w:ins` element and its contents.

### 8.4 Reject Deletion

To reject a deletion, convert `w:delText` back to `w:t` and remove the `w:del` wrapper.

---

## 9. Comparison: Word vs PowerPoint

| Feature | Word | PowerPoint |
|---------|------|------------|
| Track Changes | ✅ Full support (`w:ins`, `w:del`) | ❌ No text tracking |
| Move Tracking | ✅ Yes (`w:moveFrom`, `w:moveTo`) | ❌ No |
| Format Changes | ✅ Yes (`w:rPrChange`) | ❌ No |
| Comments | ✅ Rich text, ranges | ✅ Plain text, coordinates |
| Comment Threading | ✅ Yes (commentsExtended) | ✅ Yes (extensions) |
| Comment Resolution | ✅ Yes (`w15:done`) | ❌ No |
| Revision IDs | ✅ RSIDs per run | ❌ N/A |

---

## 10. Best Practices

### For Track Changes

1. **Use unique revision IDs** - Maintain a counter for `w:id` attributes
2. **Include timestamps** - Always set `w:date` for audit trails
3. **Preserve RSIDs** - Don't remove existing rsid attributes
4. **Handle nested changes** - A deletion can contain formatted runs

### For Comments

1. **Match IDs across parts** - `w:comment/@w:id` must match `w:commentRangeStart/@w:id`
2. **Place reference correctly** - `w:commentReference` should follow `w:commentRangeEnd`
3. **Support threading** - Use commentsExtended.xml for replies
4. **Handle resolution** - Track `w15:done` for workflow status

---

## Assumptions

- Research based on ECMA-376 5th Edition specification
- Tested with python-docx 1.x and Microsoft 365 Word
- Extended features (threading, resolution) require Word 2013 or later

## Next Steps

- [x] Implement track changes support in the MCP Office Server
- [x] Add `word_patch_with_track_changes` tool
- [x] Add `word_add_comment` tool
- [ ] Implement accept/reject changes tools
- [ ] Add comment reply/threading support

