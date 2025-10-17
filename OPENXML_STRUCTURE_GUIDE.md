# OpenXML Document Structure Guide

## Table of Contents
1. [Executive Summary](#executive-summary)
2. [DOCX File Architecture](#docx-file-architecture)
3. [XML Namespaces and Prefixes](#xml-namespaces-and-prefixes)
4. [Relationship System](#relationship-system)
5. [Document Part Structure](#document-part-structure)
6. [Nesting Rules and Element Hierarchy](#nesting-rules-and-element-hierarchy)
7. [Accuracy and Validation](#accuracy-and-validation)
8. [Implementation Patterns](#implementation-patterns)
9. [Hyperlink Best Practices](#hyperlink-best-practices)
10. [References](#references)

---

## Executive Summary

A Microsoft Word `.docx` file is **not a single file**—it's a ZIP archive containing multiple XML files, relationships, and binary resources. Understanding how these parts interconnect is critical for programmatic document manipulation.

**Key Concepts:**
- **ZIP Container**: The outer structure that holds everything
- **Parts**: Individual XML files that define content, styles, numbering, etc.
- **Relationships**: Links between parts (similar to URLs in a website)
- **Namespaces**: XML prefixes that identify the schema being used
- **Content Types**: MIME type declarations for all parts

**This document explains:**
- How XML namespaces map to functionality
- When and how to use relationships
- The nesting hierarchy that Word expects
- How to ensure accuracy when generating documents

---

## DOCX File Architecture

### High-Level Structure

```
document.docx (ZIP archive)
├── [Content_Types].xml          ← MIME type registry
├── _rels/
│   └── .rels                    ← Package-level relationships
├── word/
│   ├── document.xml             ← MAIN CONTENT (paragraphs, tables, text)
│   ├── _rels/
│   │   └── document.xml.rels    ← Document-level relationships
│   ├── styles.xml               ← Style definitions
│   ├── numbering.xml            ← List numbering definitions
│   ├── settings.xml             ← Document settings
│   ├── fontTable.xml            ← Font declarations
│   ├── theme/
│   │   └── theme1.xml           ← Color scheme and fonts
│   ├── media/
│   │   ├── image1.png           ← Embedded images
│   │   └── image2.jpg
│   ├── header1.xml              ← Header content (optional)
│   ├── footer1.xml              ← Footer content (optional)
│   └── comments.xml             ← Comments (optional)
└── docProps/
    ├── core.xml                 ← Metadata (title, author, dates)
    └── app.xml                  ← Application properties
```

### Critical File Paths (Canonical)

Per ECMA-376 Part 2 Section 1, these paths are **fixed by convention**:

| Path | Purpose | Required |
|------|---------|----------|
| `[Content_Types].xml` | Declares MIME types for all parts | ✅ YES |
| `_rels/.rels` | Package relationships (links to document.xml, properties) | ✅ YES |
| `word/document.xml` | Main document content | ✅ YES |
| `word/_rels/document.xml.rels` | Links from document to styles, images, etc. | ⚠️ If document references other parts |
| `word/styles.xml` | Style definitions | ⚠️ If styles are used |
| `word/numbering.xml` | List numbering | ⚠️ If lists exist |
| `word/media/*` | Binary resources (images, fonts) | ⚠️ If embedded |
| `docProps/core.xml` | Document properties | ⚠️ Recommended |

**Implementation Note:** DocXML defines these in `src/zip/types.ts` as the `DOCX_PATHS` constant.

---

## XML Namespaces and Prefixes

### What Are XML Namespaces?

XML namespaces prevent naming collisions when combining schemas. Think of them like TypeScript imports:

```typescript
// TypeScript analogy
import * as React from 'react';           // "React" is the prefix
import * as lodash from 'lodash';         // "lodash" is the prefix

// React.Component vs lodash.Component are distinct
```

In XML:
```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <!-- 'w:' prefix means this is WordprocessingML -->
  </w:body>
</w:document>
```

### Primary Namespace Prefixes

#### 1. **`w:` - WordprocessingML (Main Document Content)**

**Namespace URI:** `http://schemas.openxmlformats.org/wordprocessingml/2006/main`

**Usage:** Core document structure—paragraphs, runs, text, formatting.

**Spec Reference:** ECMA-376 Part 1 Section 17

**Common Elements:**
```xml
<w:document>    <!-- Root element of word/document.xml -->
<w:body>        <!-- Contains all content -->
<w:p>           <!-- Paragraph -->
<w:r>           <!-- Run (text span with uniform formatting) -->
<w:t>           <!-- Text content -->
<w:pPr>         <!-- Paragraph properties (alignment, spacing) -->
<w:rPr>         <!-- Run properties (bold, italic, font) -->
<w:tbl>         <!-- Table -->
<w:tr>          <!-- Table row -->
<w:tc>          <!-- Table cell -->
```

**When to Use:**
- Generating `word/document.xml`
- Creating headers (`word/header1.xml`) and footers (`word/footer1.xml`)
- Defining content in comments (`word/comments.xml`)

**Implementation:** `src/xml/XMLBuilder.ts` provides `XMLBuilder.w('element')` helper.

---

#### 2. **`r:` - Relationships**

**Namespace URI:** `http://schemas.openxmlformats.org/officeDocument/2006/relationships`

**Usage:** Linking document parts together (like hyperlinks in HTML).

**Spec Reference:** ECMA-376 Part 2 Section 9

**Common Use Cases:**
```xml
<!-- In word/document.xml, referencing an image -->
<w:drawing>
  <wp:inline>
    <a:graphic>
      <a:graphicData>
        <pic:pic>
          <pic:blipFill>
            <a:blip r:embed="rId5"/>  <!-- r:embed points to relationship -->
          </pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>

<!-- In word/_rels/document.xml.rels -->
<Relationship Id="rId5"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

**Relationship Attributes:**
- `r:embed` - Embedded resource (internal to package)
- `r:link` - External resource (URL)
- `r:id` - Relationship ID reference

**When to Use:**
- Images embedded in document
- Hyperlinks to external URLs
- Headers/footers attached to sections
- Styles and numbering references

**Implementation:** `src/core/Relationship.ts` and `src/core/RelationshipManager.ts`

---

#### 3. **`wp:` - DrawingML WordprocessingDrawing**

**Namespace URI:** `http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`

**Usage:** Positioning and wrapping for drawings (images, shapes) in Word documents.

**Spec Reference:** ECMA-376 Part 1 Section 20.4

**Common Elements:**
```xml
<wp:inline>     <!-- Inline with text (like a character) -->
<wp:anchor>     <!-- Floating (positioned relative to page/paragraph) -->
<wp:extent>     <!-- Size of drawing (in EMUs) -->
<wp:docPr>      <!-- Drawing properties (name, description) -->
```

**When to Use:**
- Embedding images in `word/document.xml`
- Specifying whether image is inline or floating
- Setting image wrapping (square, tight, through, top/bottom)

**Implementation:** `src/elements/Image.ts`

---

#### 4. **`a:` - DrawingML (Core Graphics)**

**Namespace URI:** `http://schemas.openxmlformats.org/drawingml/2006/main`

**Usage:** Drawing primitives—shapes, images, effects, transforms.

**Spec Reference:** ECMA-376 Part 1 Section 20.1

**Common Elements:**
```xml
<a:graphic>         <!-- Container for graphic data -->
<a:graphicData>     <!-- Specifies type (picture, chart, SmartArt) -->
<a:blip>            <!-- Binary Large Image or Picture -->
<a:stretch>         <!-- Fill mode for image -->
<a:fillRect>        <!-- Rectangle to fill -->
```

**When to Use:**
- Rendering images within `<wp:inline>` or `<wp:anchor>`
- Applying effects (shadows, reflections)
- Defining shapes and SmartArt

**Implementation:** Used in `src/elements/Image.ts` for image rendering.

---

#### 5. **`pic:` - PictureML**

**Namespace URI:** `http://schemas.openxmlformats.org/drawingml/2006/picture`

**Usage:** Specifically for pictures (not charts or shapes).

**Spec Reference:** ECMA-376 Part 1 Section 20.2

**Common Elements:**
```xml
<pic:pic>           <!-- Picture container -->
<pic:nvPicPr>       <!-- Non-visual picture properties (ID, name) -->
<pic:blipFill>      <!-- How image fills the frame -->
<pic:spPr>          <!-- Shape properties (transforms, geometry) -->
```

**When to Use:**
- Wrapping image blips in DrawingML graphics
- Always paired with `a:` and `wp:` namespaces

**Implementation:** `src/elements/Image.ts`

---

#### 6. **`cp:`, `dc:`, `dcterms:` - Document Properties**

**Namespace URIs:**
- `cp:` → `http://schemas.openxmlformats.org/package/2006/metadata/core-properties`
- `dc:` → `http://purl.org/dc/elements/1.1/` (Dublin Core)
- `dcterms:` → `http://purl.org/dc/terms/` (Dublin Core Terms)

**Usage:** Document metadata (title, author, created date, modified date).

**Spec Reference:** ECMA-376 Part 2 Section 11 (Open Packaging Conventions)

**Common Elements:**
```xml
<!-- In docProps/core.xml -->
<cp:coreProperties xmlns:cp="..." xmlns:dc="..." xmlns:dcterms="...">
  <dc:title>Document Title</dc:title>
  <dc:creator>Author Name</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">2025-10-16T10:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2025-10-16T12:45:00Z</dcterms:modified>
</cp:coreProperties>
```

**When to Use:**
- Generating `docProps/core.xml`
- Setting document metadata visible in File > Properties

**Implementation:** `src/core/DocumentGenerator.ts` - `generateCoreProps()`

---

#### 7. **`vt:` - Variant Types**

**Namespace URI:** `http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes`

**Usage:** Extended property types (vectors, arrays, variants).

**Spec Reference:** ECMA-376 Part 1 Section 22.4

**When to Use:**
- Generating `docProps/app.xml`
- Rarely needed in basic document generation

---

### Namespace Summary Table

| Prefix | Schema | Primary Use | File Locations |
|--------|--------|-------------|----------------|
| `w:` | WordprocessingML | Document content | `word/document.xml`, `word/header*.xml`, `word/footer*.xml`, `word/comments.xml` |
| `r:` | Relationships | Linking parts | `word/_rels/*.rels`, used in `w:` content |
| `wp:` | WP Drawing | Image positioning | `word/document.xml` (within `<w:drawing>`) |
| `a:` | DrawingML | Graphics primitives | `word/document.xml` (within `<w:drawing>`) |
| `pic:` | PictureML | Picture-specific | `word/document.xml` (within `<w:drawing>`) |
| `cp:` | Core Properties | Metadata | `docProps/core.xml` |
| `dc:` | Dublin Core | Metadata fields | `docProps/core.xml` |
| `dcterms:` | DC Terms | Metadata dates | `docProps/core.xml` |
| `vt:` | Variant Types | Extended properties | `docProps/app.xml` |

---

## Relationship System

### What Are Relationships?

**Analogy:** Think of a website:
- `index.html` (main page) links to `style.css` and `image.png`
- The browser resolves these links relative to `index.html`

In DOCX:
- `word/document.xml` (main document) links to `word/media/image1.png`
- The link is declared in `word/_rels/document.xml.rels`

### Relationship Structure

A relationship has four components:

```xml
<Relationship Id="rId5"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

1. **Id**: Unique identifier (must start with `rId`, e.g., `rId1`, `rId2`)
2. **Type**: Relationship type (URL defining what kind of link this is)
3. **Target**: Relative path to the linked resource
4. **TargetMode** (optional): `Internal` (default) or `External`

### Relationship Types (RelationshipType enum)

Per ECMA-376 Part 2 Section 8, these are the canonical relationship type URLs:

```typescript
// From src/core/Relationship.ts
export enum RelationshipType {
  STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  NUMBERING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  // ... more types
}
```

### Two-Level Relationship Hierarchy

#### Level 1: Package-Level Relationships (`_rels/.rels`)

These link the package root to core document parts.

**Location:** `_rels/.rels`

**Example:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                Target="docProps/core.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                Target="docProps/app.xml"/>
</Relationships>
```

**Purpose:**
- `rId1` → Points to the main document
- `rId2` → Points to document properties
- `rId3` → Points to application properties

**Generation:** `src/core/DocumentGenerator.ts` - `generateRels()`

#### Level 2: Part-Level Relationships (`word/_rels/document.xml.rels`)

These link the document to its dependencies (styles, images, headers, etc.).

**Location:** `word/_rels/document.xml.rels`

**Example:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
                Target="numbering.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"/>
  <Relationship Id="rId4"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="https://example.com"
                TargetMode="External"/>
</Relationships>
```

**Purpose:**
- `rId1` → Document uses styles from `word/styles.xml`
- `rId2` → Document uses numbering from `word/numbering.xml`
- `rId3` → Document embeds image from `word/media/image1.png`
- `rId4` → Document contains hyperlink to external URL

**Generation:** `src/core/RelationshipManager.ts` - `generateXml()`

### Relationship ID Generation Rules

**CRITICAL:** Relationship IDs must be unique within their scope.

```typescript
// BAD - Duplicate IDs
<Relationship Id="rId1" Type="...styles..." Target="styles.xml"/>
<Relationship Id="rId1" Type="...image..." Target="media/image1.png"/>  // ERROR!

// GOOD - Sequential IDs
<Relationship Id="rId1" Type="...styles..." Target="styles.xml"/>
<Relationship Id="rId2" Type="...image..." Target="media/image1.png"/>
```

**Implementation Pattern:**
```typescript
// From src/core/RelationshipManager.ts
class RelationshipManager {
  private nextId: number = 1;

  generateId(): string {
    return `rId${this.nextId++}`;  // rId1, rId2, rId3...
  }
}
```

### Referencing Relationships from Content

Once a relationship is declared in `.rels`, reference it in content:

#### Example 1: Embedded Image

**Step 1: Add relationship** (`word/_rels/document.xml.rels`)
```xml
<Relationship Id="rId5"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              Target="media/image1.png"/>
```

**Step 2: Reference in document** (`word/document.xml`)
```xml
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:blipFill>
                <a:blip r:embed="rId5"/>  <!-- ← Reference relationship ID -->
              </a:blip>
            </pic:blipFill>
          </pic:pic>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
```

#### Example 2: External Hyperlink

**Step 1: Add relationship** (`word/_rels/document.xml.rels`)
```xml
<Relationship Id="rId7"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="https://example.com"
              TargetMode="External"/>
```

**Step 2: Reference in document** (`word/document.xml`)
```xml
<w:p>
  <w:hyperlink r:id="rId7">  <!-- ← Reference relationship ID -->
    <w:r>
      <w:rPr>
        <w:rStyle w:val="Hyperlink"/>
      </w:rPr>
      <w:t>Click here</w:t>
    </w:r>
  </w:hyperlink>
</w:p>
```

**Implementation:** `src/elements/Hyperlink.ts` and `src/core/DocumentGenerator.ts`

### Internal vs External Relationships

| Type | Target | TargetMode | Example |
|------|--------|------------|---------|
| **Internal** | Relative path within package | `Internal` (default) | `media/image1.png`, `styles.xml` |
| **External** | Absolute URL outside package | `External` | `https://example.com`, `mailto:user@example.com` |

**Rule:** Hyperlinks to websites MUST use `TargetMode="External"`. Images/styles MUST NOT.

---

## Document Part Structure

### Part 1: [Content_Types].xml

**Purpose:** Registry of MIME types for all files in the package.

**Location:** `[Content_Types].xml` (root of ZIP)

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- Default types (by file extension) -->
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>

  <!-- Override types (by file path) -->
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml"
            ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
```

**Rules:**
- Every file in the ZIP must have a content type
- `Default` applies to all files with that extension
- `Override` applies to a specific file path
- PartName must start with `/`

**Implementation:** `src/core/DocumentGenerator.ts` - `generateContentTypes()`

### Part 2: word/document.xml

**Purpose:** The main document content—all paragraphs, tables, text.

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <!-- All content goes here -->
    <w:p>  <!-- Paragraph -->
      <w:r>  <!-- Run -->
        <w:t>Hello World</w:t>  <!-- Text -->
      </w:r>
    </w:p>

    <!-- Section properties at end -->
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>  <!-- Page size: 8.5" × 11" -->
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>  <!-- Margins: 1" -->
    </w:sectPr>
  </w:body>
</w:document>
```

**Implementation:** `src/core/DocumentGenerator.ts` - `generateDocumentXml()`

### Part 3: word/styles.xml

**Purpose:** Style definitions (Heading 1, Normal, etc.).

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="480" w:after="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>  <!-- 16pt font -->
    </w:rPr>
  </w:style>
</w:styles>
```

**Implementation:** `src/formatting/StylesManager.ts`

### Part 4: word/numbering.xml

**Purpose:** List numbering definitions (bullet points, numbered lists).

**Structure:**
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>
```

**Implementation:** `src/formatting/NumberingManager.ts`

---

## Nesting Rules and Element Hierarchy

### WordprocessingML Nesting (w: namespace)

Per ECMA-376 Part 1 Section 17, the content structure is strictly hierarchical:

```
w:document (root)
└── w:body (body container)
    ├── w:p (paragraph) [1..n]
    │   ├── w:pPr (paragraph properties) [0..1]
    │   │   ├── w:pStyle (style reference)
    │   │   ├── w:jc (justification/alignment)
    │   │   ├── w:ind (indentation)
    │   │   └── w:spacing (spacing before/after)
    │   ├── w:r (run) [1..n]
    │   │   ├── w:rPr (run properties) [0..1]
    │   │   │   ├── w:b (bold)
    │   │   │   ├── w:i (italic)
    │   │   │   ├── w:u (underline)
    │   │   │   ├── w:sz (font size)
    │   │   │   └── w:color (text color)
    │   │   └── w:t (text content) [1]
    │   ├── w:hyperlink (hyperlink) [0..n]
    │   │   └── w:r (run with hyperlink text)
    │   └── w:drawing (embedded image/shape) [0..n]
    ├── w:tbl (table) [1..n]
    │   ├── w:tblPr (table properties)
    │   └── w:tr (table row) [1..n]
    │       └── w:tc (table cell) [1..n]
    │           └── w:p (paragraph in cell)
    └── w:sectPr (section properties) [1] (must be last child of w:body)
```

### Critical Nesting Rules

#### Rule 1: Paragraph Properties MUST Come First

```xml
<!-- CORRECT -->
<w:p>
  <w:pPr>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r><w:t>Text</w:t></w:r>
</w:p>

<!-- WRONG - Word will reject or ignore -->
<w:p>
  <w:r><w:t>Text</w:t></w:r>
  <w:pPr>  <!-- ← Too late! Must come before runs -->
    <w:jc w:val="center"/>
  </w:pPr>
</w:p>
```

**Implementation:** `src/elements/Paragraph.ts:438-552` generates `<w:pPr>` before runs.

#### Rule 2: Run Properties MUST Come Before Text

```xml
<!-- CORRECT -->
<w:r>
  <w:rPr>
    <w:b/>
  </w:rPr>
  <w:t>Bold text</w:t>
</w:r>

<!-- WRONG -->
<w:r>
  <w:t>Bold text</w:t>
  <w:rPr>  <!-- ← Too late! -->
    <w:b/>
  </w:rPr>
</w:r>
```

**Implementation:** `src/elements/Run.ts:201-278` generates `<w:rPr>` before `<w:t>`.

#### Rule 3: Text Elements Cannot Be Self-Closing

```xml
<!-- CORRECT -->
<w:t xml:space="preserve">Hello</w:t>
<w:t xml:space="preserve"></w:t>  <!-- Empty but valid -->

<!-- WRONG - Word silently ignores, text vanishes -->
<w:t xml:space="preserve"/>  <!-- CRITICAL ERROR -->
```

**Implementation:** `src/xml/XMLBuilder.ts:121-130` throws error if `<w:t>` is self-closing.

#### Rule 4: Section Properties Must Be Last in Body

```xml
<!-- CORRECT -->
<w:body>
  <w:p><w:r><w:t>Content</w:t></w:r></w:p>
  <w:sectPr>  <!-- ← Last child -->
    <w:pgSz w:w="12240" w:h="15840"/>
  </w:sectPr>
</w:body>

<!-- WRONG -->
<w:body>
  <w:sectPr>  <!-- ← Too early -->
    <w:pgSz w:w="12240" w:h="15840"/>
  </w:sectPr>
  <w:p><w:r><w:t>Content</w:t></w:r></w:p>
</w:body>
```

**Implementation:** `src/core/DocumentGenerator.ts:58-77` adds section properties at end.

### DrawingML Nesting (wp:, a:, pic: namespaces)

Images have a deeply nested structure:

```
w:p (paragraph)
└── w:r (run)
    └── w:drawing (drawing container)
        └── wp:inline OR wp:anchor (positioning)
            ├── wp:extent (size in EMUs)
            ├── wp:docPr (properties: id, name, description)
            └── a:graphic (graphic container)
                └── a:graphicData (data type declaration)
                    └── pic:pic (picture element)
                        ├── pic:nvPicPr (non-visual properties)
                        ├── pic:blipFill (fill properties)
                        │   ├── a:blip (image reference)
                        │   │   └── @r:embed="rId5" (relationship ID)
                        │   └── a:stretch (fill mode)
                        └── pic:spPr (shape properties)
                            └── a:xfrm (transform: position, size)
```

**Key Point:** You cannot skip levels. Every element is required for Word to render the image.

**Implementation:** `src/elements/Image.ts:445-755` generates this structure.

### Property Order Within Elements

Per ECMA-376, child elements within `<w:pPr>` and `<w:rPr>` have a defined order:

#### Paragraph Properties Order (`<w:pPr>`)

```xml
<w:pPr>
  <w:pStyle w:val="Heading1"/>         <!-- 1. Style reference -->
  <w:keepNext/>                        <!-- 2. Keep with next -->
  <w:keepLines/>                       <!-- 3. Keep lines together -->
  <w:pageBreakBefore/>                 <!-- 4. Page break before -->
  <w:numPr>                            <!-- 5. Numbering properties -->
    <w:ilvl w:val="0"/>
    <w:numId w:val="1"/>
  </w:numPr>
  <w:spacing w:before="240"/>          <!-- 6. Spacing -->
  <w:ind w:left="720"/>                <!-- 7. Indentation -->
  <w:jc w:val="center"/>               <!-- 8. Justification -->
</w:pPr>
```

**Current Implementation Issue:** `src/elements/Paragraph.ts:442-495` generates properties out-of-order (see previous analysis). This is low-severity (Word accepts it) but violates the spec.

#### Run Properties Order (`<w:rPr>`)

```xml
<w:rPr>
  <w:rFonts w:ascii="Arial"/>          <!-- 1. Fonts -->
  <w:b/>                               <!-- 2. Bold -->
  <w:i/>                               <!-- 3. Italic -->
  <w:u w:val="single"/>                <!-- 4. Underline -->
  <w:strike/>                          <!-- 5. Strikethrough -->
  <w:sz w:val="24"/>                   <!-- 6. Size -->
  <w:color w:val="FF0000"/>            <!-- 7. Color -->
  <w:highlight w:val="yellow"/>        <!-- 8. Highlight -->
  <w:vertAlign w:val="subscript"/>     <!-- 9. Vertical alignment -->
</w:rPr>
```

**Current Implementation Issue:** `src/elements/Run.ts:211-262` also generates out-of-order. Should be reordered for strict compliance.

---

## Accuracy and Validation

### Ensuring Correct Document Generation

#### 1. Validate Relationship IDs

**Problem:** Duplicate relationship IDs cause unpredictable behavior.

**Solution:** Use `RelationshipManager` to auto-generate unique IDs.

```typescript
const relManager = new RelationshipManager();
const imgRel = relManager.addImage('media/image1.png');  // Returns rId1
const linkRel = relManager.addHyperlink('https://example.com');  // Returns rId2
```

**Implementation:** `src/core/RelationshipManager.ts:105-107`

#### 2. Validate Required Files

**Problem:** Missing `[Content_Types].xml` or `word/document.xml` → Word refuses to open.

**Solution:** Use `REQUIRED_DOCX_FILES` constant and validate before save.

```typescript
import { REQUIRED_DOCX_FILES } from './zip/types';

function validateDocxStructure(zipHandler: ZipHandler): void {
  for (const requiredFile of REQUIRED_DOCX_FILES) {
    if (!zipHandler.hasFile(requiredFile)) {
      throw new Error(`Missing required file: ${requiredFile}`);
    }
  }
}
```

**Implementation:** `src/zip/ZipWriter.ts:136-155`

#### 3. Prevent Self-Closing Text Tags

**Problem:** `<w:t/>` causes silent text loss.

**Solution:** XMLBuilder throws error if `<w:t>` is marked self-closing.

```typescript
// In src/xml/XMLBuilder.ts:121-130
if (element.selfClosing && element.name === 'w:t') {
  throw new Error('CRITICAL: Text elements (w:t) cannot be self-closing.');
}
```

This prevents accidental corruption at the source.

#### 4. Validate XML Structure After Parse

**Problem:** Corrupted documents may have missing text nodes.

**Solution:** `DocumentParser` tracks parse errors and validates loaded content.

```typescript
// In src/core/DocumentParser.ts:147-202
private validateLoadedContent(bodyElements: BodyElement[]): void {
  const paragraphs = bodyElements.filter(el => el instanceof Paragraph);
  let emptyRuns = 0;
  let totalRuns = 0;

  for (const para of paragraphs) {
    for (const run of para.getRuns()) {
      totalRuns++;
      if (run.getText().length === 0) emptyRuns++;
    }
  }

  if (totalRuns > 0 && (emptyRuns / totalRuns) > 0.9) {
    console.warn('Document appears corrupted: 90% of runs have no text');
  }
}
```

#### 5. Round-Trip Testing

**Best Practice:** Always test load → modify → save → load.

```typescript
// Test template
test('preserves text through round-trip', async () => {
  // Create document
  const doc1 = Document.create();
  doc1.createParagraph('Test text');
  const buffer = await doc1.toBuffer();

  // Load and verify
  const doc2 = await Document.loadFromBuffer(buffer);
  const text = doc2.getParagraphs()[0].getText();
  expect(text).toBe('Test text');  // Must match exactly
});
```

**Implementation:** `tests/integration/` contains round-trip tests.

---

## Implementation Patterns

### Pattern 1: Creating a Document with Images

```typescript
import { Document, Paragraph, Run, Image } from 'docxml';

async function createDocumentWithImage() {
  const doc = Document.create();

  // Add title paragraph
  const title = doc.createParagraph();
  title.setAlignment('center');
  title.addText('Report Title', { bold: true, size: 18 });

  // Add image paragraph
  const imgPara = doc.createParagraph();
  const img = await Image.fromFile('./chart.png');
  img.setSize(400, 300);  // 400px × 300px
  imgPara.addImage(img);

  // Save
  await doc.save('report.docx');
}
```

**What happens under the hood:**
1. `Document.create()` initializes ZIP structure and relationship manager
2. `createParagraph()` adds `<w:p>` to body elements
3. `addImage()` triggers:
   - Copy image to `word/media/image1.png`
   - Add relationship: `rId3 → media/image1.png`
   - Generate DrawingML XML with `r:embed="rId3"`
4. `save()` calls:
   - `generateDocumentXml()` → `word/document.xml`
   - `generateXml()` on RelationshipManager → `word/_rels/document.xml.rels`
   - `generateContentTypes()` → `[Content_Types].xml`
   - ZIP all files into `report.docx`

### Pattern 2: Loading and Modifying

```typescript
async function modifyExistingDocument() {
  // Load
  const doc = await Document.load('input.docx');

  // Modify
  const paragraphs = doc.getParagraphs();
  for (const para of paragraphs) {
    const runs = para.getRuns();
    for (const run of runs) {
      if (run.getText().includes('PLACEHOLDER')) {
        run.setText('ACTUAL VALUE');
      }
    }
  }

  // Save
  await doc.save('output.docx');
}
```

**What happens under the hood:**
1. `load()` calls:
   - `ZipHandler.load()` → Extract all files
   - `DocumentParser.parseDocument()` → Parse `word/document.xml`
   - `XMLParser.extractElements()` → Find all `<w:p>` tags
   - `parseParagraph()` → Create `Paragraph` objects
   - `parseRun()` → Create `Run` objects with text
   - `parseRelationships()` → Load existing `word/_rels/document.xml.rels`
2. Modifications update in-memory objects
3. `save()` regenerates XML from objects (same as Pattern 1)

### Pattern 3: Preserving Relationships During Modification

**Critical:** When modifying a document, preserve existing relationships.

```typescript
async function addImageToExistingDoc() {
  const doc = await Document.load('existing.docx');

  // RelationshipManager was populated during load
  // It already has rId1 (styles), rId2 (numbering), etc.

  // Add new image - automatically gets next available ID
  const img = await Image.fromFile('./new-image.png');
  doc.getParagraphs()[0].addImage(img);
  // → RelationshipManager generates rId3 (or higher if rId3 exists)

  await doc.save('existing.docx');  // Overwrites with new relationships
}
```

**Implementation:** `src/core/DocumentParser.ts:569-584` parses existing relationships before adding new ones.

---

## Hyperlink Best Practices

### Overview

Hyperlinks in OpenXML documents require careful handling to ensure ECMA-376 compliance and prevent document corruption. This section provides guidelines for creating, validating, and troubleshooting hyperlinks in DocXML.

### Hyperlink Types

#### External Hyperlinks

External hyperlinks point to resources outside the document (websites, email addresses, files).

**XML Structure:**
```xml
<!-- In word/document.xml -->
<w:p>
  <w:hyperlink r:id="rId7">  <!-- ← Must reference a relationship -->
    <w:r>
      <w:rPr>
        <w:color w:val="0563C1"/>  <!-- Blue link color -->
        <w:u w:val="single"/>      <!-- Underlined -->
      </w:rPr>
      <w:t>Click here</w:t>
    </w:r>
  </w:hyperlink>
</w:p>

<!-- In word/_rels/document.xml.rels -->
<Relationship Id="rId7"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
              Target="https://example.com"
              TargetMode="External"/>  <!-- ← REQUIRED for external links -->
```

**CRITICAL Requirements (ECMA-376 Part 1 §17.16.22):**
1. External hyperlinks **MUST** have a `r:id` attribute referencing a relationship
2. The relationship **MUST** have `TargetMode="External"`
3. The relationship **MUST** use `Type=".../hyperlink"`

**Failure Mode:** If `r:id` is missing or points to a non-existent relationship:
- Microsoft Word will refuse to open the document, showing corruption error
- LibreOffice may open but display broken links

#### Internal Hyperlinks (Bookmarks)

Internal hyperlinks navigate to bookmarks within the same document.

**XML Structure:**
```xml
<!-- In word/document.xml -->
<w:p>
  <w:hyperlink w:anchor="Section1">  <!-- ← Uses w:anchor, not r:id -->
    <w:r>
      <w:rPr>
        <w:color w:val="0563C1"/>
        <w:u w:val="single"/>
      </w:rPr>
      <w:t>Go to Section 1</w:t>
    </w:r>
  </w:hyperlink>
</w:p>

<!-- Bookmark definition (target) -->
<w:p>
  <w:bookmarkStart w:id="0" w:name="Section1"/>
  <w:r><w:t>Section 1 Content</w:t></w:r>
  <w:bookmarkEnd w:id="0"/>
</w:p>
```

**Key Differences from External Links:**
- Uses `w:anchor` attribute instead of `r:id`
- **No relationship required** in `document.xml.rels`
- Target must be a valid bookmark name in the document

### DocXML API Patterns

#### Recommended Pattern: Use Document.save()

**Always use `Document.save()` or `Document.toBuffer()`** — they automatically handle relationship registration.

```typescript
import { Document, Hyperlink } from 'docxml';

// ✅ CORRECT: Document handles relationships automatically
const doc = Document.create();
const para = doc.createParagraph();

// External link - relationship automatically created
para.addHyperlink(Hyperlink.createExternal('https://example.com', 'Visit us'));

// Internal link - no relationship needed
para.addHyperlink(Hyperlink.createInternal('Section1', 'Jump to section'));

await doc.save('document.docx');  // ← Relationships auto-registered
```

**What happens internally:**
1. `createExternal()` creates `Hyperlink` object with URL
2. `addHyperlink()` adds it to paragraph's content
3. `save()` calls `DocumentGenerator.processHyperlinks()`:
   - Finds all external hyperlinks
   - Calls `relationshipManager.addHyperlink(url)` for each
   - Calls `hyperlink.setRelationshipId(rId)` to link them
4. XML generation succeeds because all links have relationship IDs

#### Anti-Pattern: Manual toXML() Without Relationships

**DO NOT** call `toXML()` on external hyperlinks without setting relationship ID first.

```typescript
import { Hyperlink } from 'docxml';

// ❌ WRONG: Will throw error
const link = Hyperlink.createExternal('https://example.com', 'Link');
const xml = link.toXML();  // ERROR: Missing relationship ID

// Error message:
// "CRITICAL: External hyperlink to "https://example.com" is missing relationship ID.
//  This would create an invalid OpenXML document per ECMA-376 §17.16.22.
//  Solution: Use Document.save() which automatically registers relationships,
//  or manually call relationshipManager.addHyperlink(url) and set the relationship ID."
```

**Why this validation exists:**
- Prevents silent document corruption
- Forces developers to use correct API patterns
- Fails fast with clear error message

#### Advanced Pattern: Manual Relationship Management

For advanced use cases where you need control over relationship registration:

```typescript
import { Document, Hyperlink, RelationshipManager } from 'docxml';

const doc = Document.create();
const para = doc.createParagraph();
const relManager = doc.getRelationshipManager();

// Create hyperlink
const link = Hyperlink.createExternal('https://example.com', 'Link');

// Manually register relationship
const relationship = relManager.addHyperlink('https://example.com');

// Set relationship ID on hyperlink
link.setRelationshipId(relationship.getId());  // e.g., "rId5"

// Now toXML() will succeed
const xml = link.toXML();  // ✅ Valid XML
```

**When to use this pattern:**
- Building custom XML generators
- Debugging relationship issues
- Testing relationship logic

### Validation Rules (Enforced by DocXML)

#### Rule 1: External Links Require Relationship ID

```typescript
// From src/elements/Hyperlink.ts:238-245
toXML(): XMLElement {
  if (this.url && !this.relationshipId) {
    throw new Error(
      `CRITICAL: External hyperlink to "${this.url}" is missing relationship ID. ` +
      `This would create an invalid OpenXML document per ECMA-376 §17.16.22. ` +
      `Solution: Use Document.save() which automatically registers relationships, ` +
      `or manually call relationshipManager.addHyperlink(url) and set the relationship ID.`
    );
  }
  // ... generate XML
}
```

**Test Case:**
```typescript
// From tests/core/HyperlinkParsing.test.ts:265-272
it('should throw error if external link toXML() called without relationship ID', () => {
  const link = Hyperlink.createExternal('https://example.com', 'Link');
  expect(() => link.toXML()).toThrow(/CRITICAL: External hyperlink/);
  expect(() => link.toXML()).toThrow(/missing relationship ID/);
  expect(() => link.toXML()).toThrow(/ECMA-376 §17.16.22/);
});
```

#### Rule 2: Hyperlinks Must Have URL or Anchor

Empty hyperlinks are invalid per spec.

```typescript
// From src/elements/Hyperlink.ts:229-234
toXML(): XMLElement {
  if (!this.url && !this.anchor) {
    throw new Error(
      'CRITICAL: Hyperlink must have either a URL (external link) or anchor (internal link). ' +
      'Cannot generate valid XML for empty hyperlink.'
    );
  }
  // ... generate XML
}
```

**Test Case:**
```typescript
// From tests/core/HyperlinkParsing.test.ts:286-292
it('should throw error if hyperlink has neither url nor anchor', () => {
  const link = new Hyperlink({ text: 'Empty Link' });
  expect(() => link.toXML()).toThrow(/CRITICAL: Hyperlink must have either a URL/);
  expect(() => link.toXML()).toThrow(/or anchor/);
});
```

#### Rule 3: Warn on Hybrid Links (URL + Anchor)

Per ECMA-376, hyperlinks should have **either** URL **or** anchor, not both.

```typescript
// From src/elements/Hyperlink.ts:98-104
constructor(properties: HyperlinkProperties) {
  if (this.url && this.anchor) {
    console.warn(
      `DocXML Warning: Hyperlink has both URL ("${this.url}") and anchor ("${this.anchor}"). ` +
      `This is ambiguous per ECMA-376 spec. URL will take precedence. ` +
      `Use Hyperlink.createExternal() or Hyperlink.createInternal() to avoid ambiguity.`
    );
  }
}
```

**Test Case:**
```typescript
// From tests/core/HyperlinkParsing.test.ts:306-320
it('should warn when hyperlink has both url and anchor (hybrid link)', () => {
  const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation();

  new Hyperlink({ url: 'https://example.com', anchor: 'Section1', text: 'Hybrid' });

  expect(consoleWarnSpy).toHaveBeenCalledWith(
    expect.stringContaining('DocXML Warning: Hyperlink has both URL')
  );
  expect(consoleWarnSpy).toHaveBeenCalledWith(
    expect.stringContaining('ambiguous per ECMA-376 spec')
  );

  consoleWarnSpy.mockRestore();
});
```

### Improved Text Fallback

When hyperlink text is empty, DocXML uses an improved fallback chain:

```
text → url → anchor → "Link"
```

**Before (v0.2.x):**
```typescript
// Empty text always defaulted to generic "Link"
const link = Hyperlink.createExternal('https://example.com', '');
link.getText();  // "Link" (not helpful)
```

**After (v0.3.x):**
```typescript
// Empty text uses URL as fallback (more user-friendly)
const link = Hyperlink.createExternal('https://example.com', '');
link.getText();  // "https://example.com" (shows actual target)

const link2 = Hyperlink.createInternal('Section1', '');
link2.getText();  // "Section1" (shows bookmark name)

const link3 = new Hyperlink({ text: '' });
link3.getText();  // "Link" (only when nothing else available)
```

**Implementation:**
```typescript
// From src/elements/Hyperlink.ts:108
this.text = properties.text || this.url || this.anchor || 'Link';
```

### Troubleshooting

#### Problem: "Document is corrupted" error in Microsoft Word

**Symptom:** Word shows error: "Word found unreadable content in document.docx. Do you want to recover?"

**Cause:** External hyperlink missing relationship ID or relationship has wrong `TargetMode`.

**Solution:**
1. Always use `Document.save()` API (recommended)
2. If manually creating XML, ensure:
   - `<w:hyperlink>` has `r:id` attribute
   - Relationship exists in `word/_rels/document.xml.rels`
   - Relationship has `TargetMode="External"`

**Debug Steps:**
```bash
# Extract DOCX to inspect
unzip document.docx -d extracted/

# Check if hyperlink has r:id
grep -n "w:hyperlink" extracted/word/document.xml

# Check if relationship exists
grep -n "rId7" extracted/word/_rels/document.xml.rels

# Verify TargetMode="External"
grep "TargetMode" extracted/word/_rels/document.xml.rels
```

#### Problem: Hyperlink appears as plain text

**Symptom:** Link is not clickable in Word, shows as plain blue underlined text.

**Cause:** Hyperlink styling applied but no `<w:hyperlink>` element, just formatted `<w:r>`.

**Solution:** Use `Hyperlink` class, not just styled `Run`:

```typescript
// ❌ WRONG: This creates styled text, not a hyperlink
const run = new Run('Click here', { color: '0563C1', underline: 'single' });
para.addRun(run);

// ✅ CORRECT: This creates an actual hyperlink
const link = Hyperlink.createExternal('https://example.com', 'Click here');
para.addHyperlink(link);
```

#### Problem: Internal hyperlink doesn't jump to target

**Symptom:** Clicking bookmark link does nothing or shows error.

**Cause:** Bookmark doesn't exist or name doesn't match.

**Solution:** Ensure bookmark is defined:

```typescript
// Create target bookmark
const targetPara = doc.createParagraph();
targetPara.addBookmark('Section1');  // ← Must match anchor name
targetPara.addText('Section 1 Content');

// Create link to bookmark
const linkPara = doc.createParagraph();
linkPara.addHyperlink(Hyperlink.createInternal('Section1', 'Go to Section 1'));
```

**Validation:** DocXML doesn't currently validate bookmark existence (planned for Phase 5).

### Special Characters in Hyperlinks

#### Tooltip Attribute Escaping

Tooltip text is automatically escaped by `XMLBuilder`:

```typescript
const link = Hyperlink.createExternal('https://example.com', 'Link');
link.setTooltip('This is a "tooltip" with <special> & characters');

// XMLBuilder.escapeXmlAttribute() handles escaping:
// Output: w:tooltip="This is a &quot;tooltip&quot; with &lt;special&gt; &amp; characters"
```

**Implementation:** `src/xml/XMLBuilder.ts:escapeXmlAttribute()`

#### URL Encoding

DocXML does **not** automatically URL-encode hyperlink targets. If your URL contains special characters, encode them before passing to DocXML:

```typescript
// ❌ WRONG: Spaces and special characters not encoded
const badUrl = 'https://example.com/path with spaces?query=<value>';
const link1 = Hyperlink.createExternal(badUrl, 'Link');

// ✅ CORRECT: URL-encode special characters
const goodUrl = encodeURI('https://example.com/path with spaces?query=<value>');
// → 'https://example.com/path%20with%20spaces?query=%3Cvalue%3E'
const link2 = Hyperlink.createExternal(goodUrl, 'Link');
```

### Factory Methods

Use factory methods to avoid ambiguity:

```typescript
// External links
Hyperlink.createExternal(url, text, formatting?);
Hyperlink.createWebLink(url, text?, formatting?);  // Same as createExternal
Hyperlink.createEmail(email, text?, formatting?);  // Adds "mailto:" prefix

// Internal links
Hyperlink.createInternal(anchor, text, formatting?);

// Generic (not recommended)
Hyperlink.create(properties);  // Use when you need full control
new Hyperlink(properties);     // Same as above
```

**Examples:**
```typescript
// Web link
Hyperlink.createWebLink('https://example.com');  // Text defaults to URL

// Email link (automatically adds mailto:)
Hyperlink.createEmail('user@example.com');  // Text defaults to email

// Internal bookmark link
Hyperlink.createInternal('Conclusion', 'Jump to conclusion');

// Custom formatting
Hyperlink.createExternal('https://example.com', 'Link', {
  color: 'FF0000',    // Red instead of default blue
  bold: true,
  underline: 'double'
});
```

### Testing Hyperlinks

DocXML includes comprehensive hyperlink validation tests:

**Test Coverage:**
- ✅ External hyperlinks with relationship IDs
- ✅ Internal hyperlinks (bookmarks)
- ✅ Tooltip parsing and escaping
- ✅ Formatted hyperlink text
- ✅ Round-trip fidelity (load → save → load)
- ✅ Validation errors (missing relationship ID)
- ✅ Empty hyperlink rejection
- ✅ Hybrid link warnings
- ✅ Text fallback chain
- ✅ Document.save() workflow

**Run hyperlink tests:**
```bash
npm test -- HyperlinkParsing
```

**Test file:** `tests/core/HyperlinkParsing.test.ts` (19 tests, 520+ lines)

### Migration Guide

#### Upgrading from v0.2.x to v0.3.x

**Breaking Change:** `toXML()` now throws error if external link missing relationship ID.

**Old Code (v0.2.x):**
```typescript
// This created invalid XML silently
const link = Hyperlink.createExternal('https://example.com', 'Link');
const xml = link.toXML();  // Generated XML without r:id attribute
```

**New Code (v0.3.x):**
```typescript
// Option 1: Use Document.save() (recommended)
const doc = Document.create();
const para = doc.createParagraph();
para.addHyperlink(Hyperlink.createExternal('https://example.com', 'Link'));
await doc.save('document.docx');  // ✅ Works automatically

// Option 2: Manual relationship management (advanced)
const link = Hyperlink.createExternal('https://example.com', 'Link');
const rel = relationshipManager.addHyperlink('https://example.com');
link.setRelationshipId(rel.getId());
const xml = link.toXML();  // ✅ Now valid
```

**Impact:** If you were manually calling `toXML()` on external hyperlinks (rare), you'll get a clear error message guiding you to the solution.

### Related Sections

- **Relationship System** (earlier in this guide) - How relationships work
- **Element Hierarchy** - Where hyperlinks fit in document structure
- **Implementation Patterns** - Examples of document generation workflows

**Source Files:**
- `src/elements/Hyperlink.ts` - Hyperlink class implementation
- `src/core/DocumentGenerator.ts` - Relationship registration (processHyperlinks method)
- `src/core/DocumentParser.ts` - Hyperlink parsing from XML
- `tests/core/HyperlinkParsing.test.ts` - Comprehensive test suite

---

## References

### Official Specifications

1. **ECMA-376 Office Open XML File Formats**
   - Part 1: Fundamentals and Markup Language Reference
   - Part 2: Open Packaging Conventions
   - Part 4: Transitional Migration Features
   - Download: https://www.ecma-international.org/publications-and-standards/standards/ecma-376/

2. **ISO/IEC 29500** (identical to ECMA-376)
   - ISO standardized version of Office Open XML

### Key Sections Referenced

| Spec Section | Topic | Relevance |
|--------------|-------|-----------|
| Part 1 §17 | WordprocessingML | `w:` namespace, document structure |
| Part 1 §20.1 | DrawingML (Main) | `a:` namespace, graphics |
| Part 1 §20.2 | PictureML | `pic:` namespace, images |
| Part 1 §20.4 | WP Drawing | `wp:` namespace, image positioning |
| Part 2 §8 | Relationships | Relationship types and structure |
| Part 2 §9 | Relationship Reference | `r:` namespace attributes |
| Part 2 §10 | Package Structure | ZIP layout, required files |
| Part 2 §11 | Core Properties | `cp:`, `dc:`, `dcterms:` namespaces |

### Microsoft Documentation

1. **Open XML SDK Documentation**
   - https://learn.microsoft.com/en-us/office/open-xml/structure-of-a-wordprocessingml-document

2. **WordprocessingML Reference**
   - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing

### Tools for Inspection

1. **Open XML SDK Productivity Tool** (Windows)
   - Download: https://github.com/OfficeDev/Open-XML-SDK/releases
   - View document structure, validate schema, generate code

2. **7-Zip or WinZip**
   - Extract `.docx` files to inspect raw XML

3. **VS Code with XML Extensions**
   - Format and validate XML files

### DocXML Implementation References

| File | Purpose | Spec Section |
|------|---------|--------------|
| `src/xml/XMLBuilder.ts` | Namespace definitions | Part 1 §17, §20 |
| `src/core/Relationship.ts` | Relationship types | Part 2 §8 |
| `src/core/RelationshipManager.ts` | Relationship management | Part 2 §9 |
| `src/core/DocumentGenerator.ts` | XML generation | Part 1 §17, Part 2 §11 |
| `src/core/DocumentParser.ts` | XML parsing | Part 1 §17 |
| `src/elements/Paragraph.ts` | Paragraph structure | Part 1 §17.3.1.22 |
| `src/elements/Run.ts` | Run structure | Part 1 §17.3.2.25 |
| `src/elements/Image.ts` | Image embedding | Part 1 §20.1, §20.2, §20.4 |
| `src/zip/types.ts` | File paths | Part 2 §10 |

---

## Conclusion

Understanding OpenXML structure is essential for programmatic document manipulation. Key takeaways:

1. **DOCX is a ZIP of XML files** — not a single file format
2. **Relationships link parts together** — like URLs in a website
3. **Namespaces prevent collisions** — `w:`, `r:`, `a:`, `pic:`, etc. have distinct purposes
4. **Element order matters** — properties before content, section at end
5. **Never self-close text tags** — `<w:t/>` causes data loss
6. **Validate relationships** — duplicate IDs break documents
7. **Test round-trips** — ensure load → save preserves data

By following these patterns and understanding the underlying structure, you can build robust document generation and manipulation tools.

---

**Document Version:** 1.0
**Last Updated:** 2025-10-16
**Maintained By:** DocXML Project
**Spec Reference:** ECMA-376 4th Edition (ISO/IEC 29500)
