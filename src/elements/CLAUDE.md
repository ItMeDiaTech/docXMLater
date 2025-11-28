# Elements Module Documentation

The `elements` module contains all document element classes that represent the building blocks of a DOCX document.

## Module Overview

This module provides classes for paragraphs, runs (text spans), tables, images, and other document elements. Each class knows how to serialize itself to OOXML and can be parsed from existing documents.

**Location:** `src/elements/`

**Total Files:** 36 TypeScript files

## Architecture

### Element Hierarchy

```
Body Elements (can be added directly to document body)
├── Paragraph
├── Table
├── TableOfContents / TableOfContentsElement
└── StructuredDocumentTag

Inline Elements (contained within paragraphs)
├── Run (formatted text span)
├── ImageRun (inline image)
├── Hyperlink
├── Bookmark / RangeMarker
├── Field
├── Comment (anchor)
├── Footnote/Endnote (reference)
└── StructuredDocumentTag (inline)

Header/Footer Elements
├── Header
├── Footer
└── Section (contains header/footer references)

Manager Classes (collection management)
├── RevisionManager - Track changes
├── BookmarkManager - Bookmarks
├── CommentManager - Comments
├── ImageManager - Images and media
├── HeaderFooterManager - Headers/footers
├── FontManager - Font declarations
├── FootnoteManager - Footnotes
└── EndnoteManager - Endnotes
```

## Core Element Classes

### Paragraph (`Paragraph.ts`)

The primary container for document text and inline content.

**Key Features:**
- Text content via Run elements
- Paragraph formatting (alignment, spacing, indentation)
- List numbering support
- Style application
- Revision tracking (w:pPrChange)
- Section breaks (inline sectPr)

**Usage:**
```typescript
const para = new Paragraph();
para.addText('Hello ').setBold(true);
para.addText('World').setItalic(true);
para.setAlignment('center');
para.setSpacing({ before: 240, after: 120 });
```

**Important Properties:**
- `formatting.alignment` - 'left' | 'center' | 'right' | 'justify'
- `formatting.spacing` - { before, after, line, lineRule }
- `formatting.indentation` - { left, right, firstLine, hanging }
- `formatting.pPrChange` - Paragraph property change tracking
- `formatting.keepNext` / `formatting.keepLines` - Pagination control
- `formatting.pageBreakBefore` - Force page break

**Note:** Setting `keepNext` or `keepLines` automatically clears `pageBreakBefore` to prevent Word rendering issues (massive whitespace).

### Run (`Run.ts`)

A formatted text span within a paragraph.

**Key Features:**
- Character formatting (bold, italic, underline, strikethrough)
- Font properties (family, size, color, highlight)
- Special characters (tabs, breaks, symbols)
- Revision tracking for formatting changes
- Complex script support (bidirectional text)

**Usage:**
```typescript
const run = new Run('Text content');
run.setBold(true)
   .setFont('Arial', 12)
   .setColor('FF0000')
   .setHighlight('yellow');
```

**Color Handling:**
- All colors normalized to uppercase 6-character hex
- Supports 3-char shorthand: '#F00' becomes 'FF0000'
- Theme colors supported via themeColor property

### Table (`Table.ts`)

Table container with rows and cells.

**Related Files:**
- `TableRow.ts` - Table row element
- `TableCell.ts` - Table cell element
- `TableGridChange.ts` - Grid structure change tracking

**Key Features:**
- Row/column management
- Cell merging (horizontal and vertical)
- Border styling per cell/row/table
- Table-wide formatting
- Property change tracking (w:tblPrChange)

**Usage:**
```typescript
const table = new Table(3, 4); // 3 rows, 4 columns
table.getRow(0).getCell(0).addParagraph('Header');
table.setColumnWidths([2000, 2000, 2000, 2000]);
```

### Image (`Image.ts`, `ImageRun.ts`)

Image handling for the document.

**Key Features:**
- Supports PNG, JPEG, GIF, SVG
- Inline and floating positioning
- Size and aspect ratio control
- Text wrapping options
- Relationship management

**Usage:**
```typescript
const image = await doc.addImage(imageBuffer, {
  width: 100,   // points
  height: 100,  // points
  type: 'png',
});
```

## Revision System Classes

### Revision (`Revision.ts`)

Represents a single tracked change.

**Supported Types:**
- Content: `insert`, `delete`, `moveFrom`, `moveTo`
- Properties: `runPropertiesChange`, `paragraphPropertiesChange`, `tablePropertiesChange`, etc.
- Table: `tableCellInsert`, `tableCellDelete`, `tableCellMerge`
- Other: `numberingChange`, `sectionPropertiesChange`

**Usage:**
```typescript
// Create insertion
const insertion = Revision.createInsertion('Author', new Run('new text'));

// Create deletion
const deletion = Revision.createDeletion('Author', new Run('removed text'));

// Create from text
const revision = Revision.fromText('insert', 'Author', 'text content');

// Property change
const propChange = Revision.createRunPropertiesChange(
  'Author',
  new Run('formatted'),
  { bold: true }  // previous properties
);
```

**Key Methods:**
- `getType()` - Revision type
- `getAuthor()` - Who made the change
- `getDate()` - When change was made
- `getRuns()` - Affected content
- `toXML()` - Generate OOXML

### RevisionManager (`RevisionManager.ts`)

Manages all revisions in a document.

**Key Features:**
- Unique ID assignment
- Revision registration and retrieval
- Category-based filtering (content, formatting, structural, table)
- Author and date filtering
- Location-aware queries
- Statistics and summaries
- Validation (duplicate IDs, orphaned moves)

**Usage:**
```typescript
const manager = doc.getRevisionManager();

// Register revisions
manager.register(revision);

// Query revisions
const insertions = manager.getAllInsertions();
const byAuthor = manager.getRevisionsByAuthor('Alice');
const recent = manager.getRecentRevisions(10);

// Get statistics
const stats = manager.getStats();
const summary = manager.getSummary();

// Validation
const idValidation = manager.validateRevisionIds();
const moveValidation = manager.validateMovePairs();
```

### RevisionContent (`RevisionContent.ts`)

Type definitions for content inside revisions.

### PropertyChangeTypes (`PropertyChangeTypes.ts`)

Type definitions for property change tracking, including:
- `RevisionLocation` - Position in document structure
- Property type definitions for run, paragraph, table changes

## Header/Footer Classes

### Header / Footer (`Header.ts`, `Footer.ts`)

Container elements for header and footer content.

### HeaderFooterManager (`HeaderFooterManager.ts`)

Manages different header/footer types:
- Default headers/footers
- First page headers/footers
- Odd/even page headers/footers

### Section (`Section.ts`)

Section properties including:
- Page size and margins
- Orientation (portrait/landscape)
- Column layout
- Header/footer references
- Page numbering

## Note Elements

### Footnote / Endnote (`Footnote.ts`, `Endnote.ts`)

Note content elements.

### FootnoteManager / EndnoteManager

Manage footnote and endnote collections.

## Other Elements

### Hyperlink (`Hyperlink.ts`)

Internal and external hyperlinks with:
- URL targets
- Anchor targets (bookmarks)
- Tooltip text
- Relationship management

### Bookmark (`Bookmark.ts`)

Named locations in document for:
- Cross-references
- Hyperlink anchors
- Table of contents entries

### BookmarkManager (`BookmarkManager.ts`)

Manages bookmark collection with unique ID assignment.

### Field (`Field.ts`, `FieldHelpers.ts`)

Complex field support:
- Merge fields (MERGEFIELD)
- Date/time fields (DATE, TIME)
- Page numbers (PAGE, NUMPAGES)
- Table of contents (TOC)
- Cross-references (REF)

### Comment (`Comment.ts`)

Document annotation with:
- Author and date
- Reply threading
- Anchor positioning

### CommentManager (`CommentManager.ts`)

Manages comment collection.

### TableOfContents / TableOfContentsElement

Table of contents generation and elements.

### StructuredDocumentTag (`StructuredDocumentTag.ts`)

Content controls (SDTs) for:
- Rich text content
- Plain text content
- Date pickers
- Dropdown lists
- Checkboxes

### Shape / TextBox (`Shape.ts`, `TextBox.ts`)

Drawing elements:
- Basic shapes
- Text boxes with content
- Positioning and sizing

### FontManager (`FontManager.ts`)

Font table management and declarations.

### ImageManager (`ImageManager.ts`)

Image collection management with relationship handling.

### RangeMarker (`RangeMarker.ts`)

Generic range start/end markers for bookmarks, comments, etc.

### CommonTypes (`CommonTypes.ts`)

Shared type definitions used across elements.

## Testing

The elements module has comprehensive test coverage:

**Test Files:**
- `tests/elements/Paragraph.test.ts` - Paragraph formatting
- `tests/elements/Table.test.ts` - Table operations
- `tests/elements/Revision.test.ts` - Revision creation and serialization (35+ tests)
- `tests/elements/RevisionManager.test.ts` - Revision management (65+ tests)
- `tests/elements/Field.test.ts` - Field support
- `tests/elements/ContentControls.test.ts` - SDT support
- `tests/elements/ImageProperties.test.ts` - Image handling
- Plus 20+ additional test files for specific features

**Total: 400+ element-related tests**

## Best Practices

### Creating Elements

```typescript
// Prefer using Document factory methods
const para = doc.createParagraph('text');
const revision = doc.createInsertion('Author', new Run('text'));

// Or create directly and add
const para = new Paragraph();
para.addText('content');
doc.addParagraph(para);
```

### Element Serialization

Each element has a `toXML()` method that generates OOXML:

```typescript
const xml = paragraph.toXML();
// Returns XMLElement { name, attributes, children }
```

### Memory Management

Elements don't hold references to their parent document. Store references externally if needed.

## See Also

- `src/core/CLAUDE.md` - Document class and parsing
- `src/formatting/CLAUDE.md` - Styles and numbering
- `src/xml/CLAUDE.md` - XML generation
- `docs/guides/using-track-changes.md` - Track changes guide
