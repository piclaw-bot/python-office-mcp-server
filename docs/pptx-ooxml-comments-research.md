# PowerPoint OOXML Comments and Track Changes Research

## Executive Summary

This document provides comprehensive research on PowerPoint OOXML format for comments and track changes (revisions), based on the ECMA-376 standard and practical implementation testing.

**Key Findings:**
- PowerPoint supports comments through `p:cmAuthorLst` and `p:cmLst` XML structures
- **PowerPoint does NOT support track changes like Word does** - there is no equivalent to `w:ins` and `w:del`
- Comments can be added programmatically via manual XML manipulation since python-pptx lacks native support

---

## 1. PowerPoint Comments XML Structure

### 1.1 Comment Authors Part (commentAuthors.xml)

**Location:** `/ppt/commentAuthors.xml`  
**Content-Type:** `application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml`

This part stores information about all comment authors in the presentation. It must be linked from `presentation.xml` via a relationship.

#### XML Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="0" name="John Doe" initials="JD" lastIdx="1" clrIdx="0"/>
  <p:cmAuthor id="1" name="Jane Smith" initials="JS" lastIdx="2" clrIdx="1"/>
</p:cmAuthorLst>
```

#### Element: `p:cmAuthor` (Comment Author)

| Attribute | Required | Type | Description |
|-----------|----------|------|-------------|
| `id` | Yes | unsignedInt | Unique identifier for the author |
| `name` | Yes | string | Full name of the author |
| `initials` | No | string | Author's initials |
| `lastIdx` | Yes | unsignedInt | Last comment index used by this author |
| `clrIdx` | Yes | unsignedInt | Color index for distinguishing authors (0-9) |

---

### 1.2 Slide Comments Part (commentN.xml)

**Location:** `/ppt/comments/comment{n}.xml` (one per slide with comments)  
**Content-Type:** `application/vnd.openxmlformats-officedocument.presentationml.comments+xml`

Each slide that has comments gets its own comments part, linked from the slide via a relationship.

#### XML Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2026-01-21T10:30:00.000" idx="1">
    <p:pos x="1524000" y="914400"/>
    <p:text>This is my comment on the slide</p:text>
  </p:cm>
  <p:cm authorId="1" dt="2026-01-21T11:00:00.000" idx="2">
    <p:pos x="3048000" y="1828800"/>
    <p:text>Another comment by a different author</p:text>
    <p:extLst>
      <!-- Extensions for modern comment features like threads -->
    </p:extLst>
  </p:cm>
</p:cmLst>
```

#### Element: `p:cm` (Comment)

| Attribute | Required | Type | Description |
|-----------|----------|------|-------------|
| `authorId` | Yes | unsignedInt | References the author ID from commentAuthors.xml |
| `dt` | No | dateTime | DateTime when comment was created (ISO 8601) |
| `idx` | Yes | unsignedInt | Unique index for this comment within the author's comments |

#### Child Elements

| Element | Required | Description |
|---------|----------|-------------|
| `p:pos` | Yes | Position of the comment marker (x, y in EMUs) |
| `p:text` | Yes | The comment text content |
| `p:extLst` | No | Extension list for additional features |

---

## 2. Package Relationships

### 2.1 Presentation to Comment Authors

Add to `/ppt/_rels/presentation.xml.rels`:

```xml
<Relationship 
    Id="rIdN" 
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" 
    Target="commentAuthors.xml"/>
```

### 2.2 Slide to Comments

Add to `/ppt/slides/_rels/slideN.xml.rels`:

```xml
<Relationship 
    Id="rIdN" 
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" 
    Target="../comments/comment1.xml"/>
```

### 2.3 Content Types

Add to `[Content_Types].xml`:

```xml
<Override 
    PartName="/ppt/commentAuthors.xml" 
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>
<Override 
    PartName="/ppt/comments/comment1.xml" 
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>
```

---

## 3. Track Changes / Revisions in PowerPoint

### ⚠️ IMPORTANT: PowerPoint Does NOT Support Track Changes Like Word

Unlike Word which has `w:ins` and `w:del` elements for tracking insertions and deletions, **PowerPoint has NO equivalent mechanism for tracking text content changes**.

### What PowerPoint DOES Have

PowerPoint tracks **slide-level operations only** through revision-related parts:

| Part | Content Type |
|------|--------------|
| `/ppt/revisionInfo.xml` | `application/vnd.ms-powerpoint.revisioninfo+xml` |

**Relationship Type:** `http://schemas.microsoft.com/office/2015/10/relationships/revisionInfo`

### What Revisions Track

- ✅ Slide additions/deletions
- ✅ Slide reordering  
- ❌ Text content changes
- ❌ Shape modifications
- ❌ Formatting changes

### Alternative for "Track Changes"

For tracking text changes in presentations, **comments are the standard approach**. You can add comments to indicate:
- What was changed
- Why it was changed
- Previous values (in comment text)

---

## 4. Modern Comments (Office 2019+)

Microsoft introduced "Modern Comments" in newer Office versions with threading capabilities:

```xml
<p:cm authorId="0" dt="2026-01-21T10:30:00.000" idx="1">
  <p:pos x="1524000" y="914400"/>
  <p:text>Original comment</p:text>
  <p:extLst>
    <p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}">
      <p15:threadingInfo 
          xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2015/main" 
          timeZoneBias="0"/>
    </p:ext>
  </p:extLst>
</p:cm>
```

**Modern Comment Namespace:** `http://schemas.microsoft.com/office/powerpoint/2015/main` (prefix: `p15`)

---

## 5. Key Constants and URIs

### Namespace URIs

| Prefix | Namespace URI |
|--------|---------------|
| `p` | `http://schemas.openxmlformats.org/presentationml/2006/main` |
| `a` | `http://schemas.openxmlformats.org/drawingml/2006/main` |
| `r` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| `pr` | `http://schemas.openxmlformats.org/package/2006/relationships` |
| `ct` | `http://schemas.openxmlformats.org/package/2006/content-types` |

### Content Types

```python
CT_COMMENT_AUTHORS = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
```

### Relationship Types

```python
RT_COMMENT_AUTHORS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
RT_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
```

### Position Units (EMUs)

| Conversion | Value |
|------------|-------|
| 1 inch | 914,400 EMUs |
| 1 point | 12,700 EMUs |
| 1 cm | 360,000 EMUs |

### DateTime Format

ISO 8601 format: `YYYY-MM-DDTHH:MM:SS.sss`

Example: `2026-01-21T10:30:00.000`

---

## 6. XSD Schema Excerpts (ECMA-376)

### CT_CommentAuthor

```xml
<xsd:complexType name="CT_CommentAuthor">
  <xsd:sequence>
    <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
  </xsd:sequence>
  <xsd:attribute name="id" type="xsd:unsignedInt" use="required"/>
  <xsd:attribute name="name" type="xsd:string" use="required"/>
  <xsd:attribute name="initials" type="xsd:string"/>
  <xsd:attribute name="lastIdx" type="xsd:unsignedInt" use="required"/>
  <xsd:attribute name="clrIdx" type="xsd:unsignedInt" use="required"/>
</xsd:complexType>
```

### CT_CommentAuthorList

```xml
<xsd:complexType name="CT_CommentAuthorList">
  <xsd:sequence>
    <xsd:element name="cmAuthor" type="CT_CommentAuthor" minOccurs="0" maxOccurs="unbounded"/>
  </xsd:sequence>
</xsd:complexType>
```

### CT_Comment

```xml
<xsd:complexType name="CT_Comment">
  <xsd:sequence>
    <xsd:element name="pos" type="a:CT_Point2D"/>
    <xsd:element name="text" type="xsd:string"/>
    <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0"/>
  </xsd:sequence>
  <xsd:attribute name="authorId" type="xsd:unsignedInt" use="required"/>
  <xsd:attribute name="dt" type="xsd:dateTime"/>
  <xsd:attribute name="idx" type="xsd:unsignedInt" use="required"/>
</xsd:complexType>
```

### CT_CommentList

```xml
<xsd:complexType name="CT_CommentList">
  <xsd:sequence>
    <xsd:element name="cm" type="CT_Comment" minOccurs="0" maxOccurs="unbounded"/>
  </xsd:sequence>
</xsd:complexType>
```

### a:CT_Point2D

```xml
<xsd:complexType name="CT_Point2D">
  <xsd:attribute name="x" type="ST_Coordinate" use="required"/>
  <xsd:attribute name="y" type="ST_Coordinate" use="required"/>
</xsd:complexType>
```

---

## 7. Python Implementation

Since python-pptx doesn't natively support comments, here's how to add them manually:

### Implementation Steps

1. **Create the presentation** using python-pptx
2. **Save** the presentation to a file
3. **Re-open as ZIP** and modify:
   - Add `ppt/commentAuthors.xml`
   - Add `ppt/comments/commentN.xml` for each slide with comments
   - Modify `[Content_Types].xml` to include new parts
   - Modify `ppt/_rels/presentation.xml.rels` for comment authors
   - Modify `ppt/slides/_rels/slideN.xml.rels` for each slide's comments

### Working Python Code

```python
import zipfile
from datetime import datetime
from lxml import etree

# Namespaces
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

# Constants
CT_COMMENT_AUTHORS = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml"
RT_COMMENT_AUTHORS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
RT_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
EMU_PER_INCH = 914400


def create_comment_authors_xml(authors):
    """Create commentAuthors.xml content."""
    nsmap = {None: NS_P}
    root = etree.Element("{%s}cmAuthorLst" % NS_P, nsmap=nsmap)
    
    for author in authors:
        elem = etree.SubElement(root, "{%s}cmAuthor" % NS_P)
        elem.set("id", str(author["id"]))
        elem.set("name", author["name"])
        if author.get("initials"):
            elem.set("initials", author["initials"])
        elem.set("lastIdx", str(author["lastIdx"]))
        elem.set("clrIdx", str(author["clrIdx"]))
    
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
           etree.tostring(root, encoding="unicode").encode("utf-8")


def create_comments_xml(comments):
    """Create commentN.xml content for a slide."""
    nsmap = {None: NS_P}
    root = etree.Element("{%s}cmLst" % NS_P, nsmap=nsmap)
    
    for cm in comments:
        elem = etree.SubElement(root, "{%s}cm" % NS_P)
        elem.set("authorId", str(cm["authorId"]))
        if cm.get("dt"):
            if isinstance(cm["dt"], datetime):
                elem.set("dt", cm["dt"].strftime("%Y-%m-%dT%H:%M:%S.000"))
            else:
                elem.set("dt", cm["dt"])
        elem.set("idx", str(cm["idx"]))
        
        pos = etree.SubElement(elem, "{%s}pos" % NS_P)
        pos.set("x", str(int(cm["x"] * EMU_PER_INCH)))
        pos.set("y", str(int(cm["y"] * EMU_PER_INCH)))
        
        text = etree.SubElement(elem, "{%s}text" % NS_P)
        text.text = cm["text"]
    
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + \
           etree.tostring(root, encoding="unicode").encode("utf-8")


# Usage example
authors = [
    {"id": 0, "name": "John Doe", "initials": "JD", "lastIdx": 1, "clrIdx": 0},
]

slide_comments = {
    1: [
        {
            "authorId": 0,
            "dt": datetime.now(),
            "idx": 1,
            "x": 1.5,  # inches
            "y": 1.5,  # inches
            "text": "This is a comment"
        }
    ]
}
```

---

## 8. Comparison: Word vs PowerPoint Comments

| Feature | Word | PowerPoint |
|---------|------|------------|
| Comments | ✅ Yes (`w:comments`) | ✅ Yes (`p:cmLst`) |
| Comment Authors | Part of comment | Separate part (`p:cmAuthorLst`) |
| Track Changes | ✅ Yes (`w:ins`, `w:del`) | ❌ No |
| Position | Range markers | X/Y coordinates (EMUs) |
| Threading | Via replies | Via extensions (Office 2019+) |
| Rich Text | ✅ Yes | ❌ Plain text only |

---

## Assumptions

- Research based on ECMA-376 5th Edition specification
- Tested with python-pptx 0.6.x and Microsoft 365 PowerPoint
- Modern comments (threading) require Office 2019 or later

## Next Steps

- [ ] Implement comment support in the MCP Office Server
- [ ] Add `pptx_add_comment` tool
- [ ] Test with various PowerPoint versions
- [ ] Consider implementing a visual comment indicator approach
