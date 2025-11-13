# DOCX Editing Framework - Project Specification

## Project Overview

Build a comprehensive, production-ready DOCX editing framework from scratch that can create, read, modify, and manipulate Microsoft Word documents programmatically.

## Current Status (Updated: November 2025)

**Phases Completed: 5 of 5 - Production Ready**

| Phase                            | Status   | Tests      | Features                                          |
| -------------------------------- | -------- | ---------- | ------------------------------------------------- |
| **Phase 1: Foundation**          | Complete | 80 tests   | ZIP handling, XML generation, validation          |
| **Phase 2: Core Elements**       | Complete | 46 tests   | Paragraph, Run, formatting                        |
| **Phase 3: Advanced Formatting** | Complete | 100+ tests | Styles, tables, sections, lists                   |
| **Phase 4: Rich Content**        | Complete | 500+ tests | Images, headers, footers, hyperlinks, bookmarks   |
| **Phase 5: Polish**              | Complete | 800+ tests | Track changes, comments, TOC, fields, footnotes   |

**Total: 2073+ tests passing | 65 source files | ~25,000+ lines of code**

### What Works Now

**Core Document Operations:**
- Create DOCX files from scratch
- Read and modify existing DOCX files
- Buffer-based operations (load/save from memory)
- Document properties (core, extended, custom)
- Memory management with dispose pattern

**Text & Paragraph Formatting:**
- Format text (bold, italic, underline, colors, fonts, highlight)
- Format paragraphs (alignment, indentation, spacing, borders, shading)
- Text search and replace with regex support
- Custom styles (paragraph, character, table styles)
- Multi-level numbered and bulleted lists

**Tables & Sections:**
- Tables with formatting, borders, shading, and cell spanning
- Advanced table features (vertical merge, complex borders)
- Section configuration (page size, margins, orientation, columns)

**Rich Content (Phase 4):**
- Images (PNG, JPEG, GIF, SVG) with positioning and wrapping
- Headers & footers (different first page, odd/even pages)
- Hyperlinks (external, internal) with defragmentation utility
- Bookmarks and cross-references
- Shapes and text boxes

**Advanced Features (Phase 5):**
- Track changes (revisions for insertions, deletions, formatting)
- Comments and annotations
- Table of contents generation
- Fields (merge fields, date/time, page numbers, TOC)
- Footnotes and endnotes
- Content controls (Structured Document Tags)
- Font management

**Developer Tools:**
- Complete XML generation and parsing
- 40+ unit conversion functions (twips, EMUs, points, pixels)
- Validation utilities and corruption detection
- Full TypeScript support with type definitions
- Comprehensive error handling

### Module Documentation

Each module has its own CLAUDE.md file:

- `src/zip/CLAUDE.md` - ZIP archive handling
- `src/xml/CLAUDE.md` - XML generation and parsing
- `src/elements/CLAUDE.md` - Document elements (Paragraph, Run, Table, etc.)
- `src/utils/CLAUDE.md` - Validation utilities
- `src/core/CLAUDE.md` - Core document classes (Document, Parser, Generator)
- `src/formatting/CLAUDE.md` - Style and numbering systems
- `src/managers/CLAUDE.md` - Manager pattern classes

## Core Requirements

### 1. DOCX Format Understanding

- **File Structure**: DOCX files are ZIP archives containing XML files and resources
- **Key Components**:
  - `[Content_Types].xml` - MIME types for all parts
  - `_rels/.rels` - Package-level relationships
  - `word/document.xml` - Main document content
  - `word/_rels/document.xml.rels` - Document relationships
  - `word/styles.xml` - Style definitions
  - `word/numbering.xml` - List numbering definitions
  - `word/settings.xml` - Document settings
  - `word/fontTable.xml` - Font declarations
  - `word/theme/theme1.xml` - Theme colors and fonts
  - `word/media/` - Embedded images and media
  - `docProps/core.xml` - Core document properties
  - `docProps/app.xml` - Application-specific properties

### 2. Core Features to Implement

#### Text Manipulation

- Insert, delete, and replace text
- Find and replace with regex support
- Text extraction
- Preserve formatting during edits

#### Formatting

- **Character formatting**: Bold, italic, underline, strikethrough, subscript, superscript
- **Font properties**: Font family, size, color (RGB and theme colors)
- **Highlight colors**: Background highlighting
- **Text effects**: Small caps, all caps, shadow, emboss, engrave

#### Paragraph Formatting

- Alignment (left, center, right, justify)
- Indentation (first line, hanging, left, right)
- Line spacing and spacing before/after
- Borders and shading
- Keep with next, keep lines together, page break before

#### Styles

- Read and apply existing styles
- Create custom styles (paragraph, character, table, list)
- Modify style definitions
- Style inheritance and cascading

#### Lists

- Numbered lists (decimal, roman, alpha)
- Bulleted lists (various bullet styles)
- Multi-level lists
- Custom numbering formats

#### Tables

- Create tables with specified rows/columns
- Add/delete rows and columns
- Merge and split cells
- Cell formatting (borders, shading, alignment)
- Column widths and row heights
- Table styles

#### Images

- Insert images (PNG, JPEG, GIF, SVG)
- Position and size images
- Wrap text around images
- Image relationships and part management

#### Sections

- Multiple sections with different properties
- Page orientation (portrait/landscape)
- Page size and margins
- Headers and footers per section
- Page numbering

#### Headers and Footers

- Different first page
- Different odd/even pages
- Page numbers with formatting
- Dynamic fields (date, time, filename)

#### Advanced Features

- Track changes (insertions, deletions, formatting)
- Comments and annotations
- Hyperlinks (internal and external)
- Bookmarks and cross-references
- Table of contents generation
- Fields and field codes
- Document properties (author, title, keywords, etc.)

### 3. API Design

#### Core Classes

```
Document
  - load(filepath) / loadFromBuffer(buffer)
  - save(filepath) / saveToBuffer()
  - addParagraph(text, options)
  - addTable(rows, cols, options)
  - addImage(source, options)
  - getBody()
  - getSections()
  - getStyles()

Paragraph
  - addRun(text, formatting)
  - getRuns()
  - setAlignment()
  - setIndentation()
  - setSpacing()
  - setNumbering()
  - applyStyle()

Run (formatted text span)
  - setText()
  - setBold() / setItalic() / setUnderline()
  - setFont(name, size)
  - setColor()
  - setHighlight()

Table
  - addRow()
  - getRow(index)
  - mergeCells(startRow, startCol, endRow, endCol)
  - setCellContent()
  - setBorders()
  - setColumnWidths()

Section
  - setPageSize()
  - setPageMargins()
  - setOrientation()
  - addHeader() / addFooter()

Style
  - setName()
  - setBasedOn()
  - setCharacterFormatting()
  - setParagraphFormatting()
```

#### Usage Examples

```javascript
// Create document
const doc = new Document();

// Add styled paragraph
const para = doc.addParagraph();
para.addRun("Hello ", { bold: true });
para.addRun("World", { italic: true, color: "FF0000" });
para.setAlignment("center");

// Add table
const table = doc.addTable(3, 4);
table.getRow(0).getCell(0).addParagraph("Header 1");

// Save
doc.save("output.docx");
```

### 4. Technical Architecture

#### Dependencies

- **JSZip** or **AdmZip**: ZIP archive handling
- **xml2js** or **fast-xml-parser**: XML parsing and generation
- **Sharp** (optional): Image processing
- **TypeScript**: Type safety (recommended)

#### Module Structure

```
src/
  ├── core/
  │   ├── Document.ts
  │   ├── Part.ts (base class for document parts)
  │   └── Relationship.ts
  ├── elements/
  │   ├── Paragraph.ts
  │   ├── Run.ts
  │   ├── Table.ts
  │   ├── Image.ts
  │   └── Section.ts
  ├── formatting/
  │   ├── Style.ts
  │   ├── Numbering.ts
  │   └── Theme.ts
  ├── xml/
  │   ├── XMLBuilder.ts
  │   ├── XMLParser.ts
  │   └── namespaces.ts
  ├── utils/
  │   ├── units.ts (EMUs, twips, points conversions)
  │   ├── colors.ts
  │   └── validation.ts
  └── index.ts
```

### 5. Character Encoding (UTF-8)

**Critical Requirement**: All text content in DOCX files must be UTF-8 encoded per OpenXML (ECMA-376) specification.

#### Implementation Details

**File I/O:**

- All XML files include `encoding="UTF-8"` in their XML declaration
- String content is explicitly converted to UTF-8 Buffers before being added to the ZIP archive
- When reading, text files are decoded as UTF-8 strings
- This ensures consistent encoding regardless of system locale or platform

**Code Locations:**

- `src/zip/ZipWriter.ts` - Converts string content to UTF-8 Buffer in `addFile()`
- `src/zip/ZipReader.ts` - Extracts text as UTF-8 strings via `async('string')`
- `src/zip/ZipHandler.ts` - Wrapper methods explicitly document UTF-8 handling

**Character Support:**
The framework correctly handles:

- ASCII text (a-z, A-Z, 0-9)
- Latin characters with diacritics (à, é, ñ, ü, etc.)
- Greek letters (α, β, γ, δ, etc.)
- Cyrillic (а, б, в, г, etc.)
- Arabic (ا, ب, ت, ث, etc.)
- Hebrew (א, ב, ג, ד, etc.)
- Devanagari (अ, आ, इ, ई, etc.)
- CJK characters (Chinese 中, Japanese 日, Korean 한)
- Emoji (😀, 🎉, ❤️, 🚀, etc.)
- Right-to-left text (Arabic, Hebrew)
- Complex multi-byte sequences

**Testing:**

- 11 comprehensive UTF-8 encoding tests in `tests/zip/ZipHandler.test.ts`
- Tests cover emoji, mixed scripts, RTL text, and round-trip verification
- All 62 tests in ZipHandler suite pass with 100% success rate

**Best Practices:**

1. All string input is automatically UTF-8 encoded - no explicit encoding needed
2. All text output is UTF-8 decoded - use as standard JavaScript strings
3. Binary files (images) are preserved as-is without encoding conversion
4. XML declarations always specify `encoding="UTF-8"` and `standalone="yes"`

### 6. XML Parsing (parseToObject)

**New Feature (v0.11.0)**: XMLParser now includes `parseToObject()` method for converting XML to JavaScript objects.

**Compatible with fast-xml-parser format:**

- Attributes → `@_` prefix (e.g., `@_Id`, `@_Type`)
- Text content → `#text` property
- Multiple child elements → Array `[]`
- Single child element → Object `{}`
- Namespaces → Preserved in keys (e.g., `w:p`, `w:r`)
- Self-closing tags → Empty object `{}`

**Usage Example:**

```typescript
import { XMLParser } from "docxmlater";

const xml = `
  <Relationships xmlns="http://...">
    <Relationship Id="rId1" Type="http://..." Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://..." Target="numbering.xml"/>
  </Relationships>
`;

const result = XMLParser.parseToObject(xml);
// Result: { Relationships: { Relationship: [{ '@_Id': 'rId1', ... }, { '@_Id': 'rId2', ... }] } }
```

**Parsing Options:**

- `ignoreAttributes`: Ignore all attributes (default: false)
- `attributeNamePrefix`: Custom attribute prefix (default: '@\_')
- `textNodeName`: Custom text property name (default: '#text')
- `parseAttributeValue`: Parse numbers/booleans (default: true)
- `trimValues`: Trim whitespace (default: true)
- `alwaysArray`: Always return arrays for elements (default: false)

**Key Features:**

- Position-based parsing prevents ReDoS attacks
- Automatic array coalescing for duplicate element names
- Type conversion for numeric/boolean attribute values
- Handles all OOXML structures (Relationships, Styles, Document XML)
- Safe for large documents (size validation)
- Full namespace handling (including ignoreNamespace option)
- **39 comprehensive tests - 100% passing**

### 7. XML Namespaces

Must handle these OpenXML namespaces:

```xml
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
xmlns:v="urn:schemas-microsoft-com:vml"
```

### 8. Unit Conversions

Handle multiple measurement units:

- **Twips**: 1/20th of a point (used for most measurements)
- **EMUs**: English Metric Units (used for images, 914400 EMUs = 1 inch)
- **Points**: Typography unit (72 points = 1 inch)
- **Pixels**: Screen measurement (depends on DPI)
- **Inches/Centimeters**: Human-readable units

### 8. Testing Requirements

- Unit tests for each component
- Integration tests for complex documents
- Test against Microsoft Word for compatibility
- Test edge cases:
  - Empty documents
  - Large documents (1000+ pages)
  - Complex tables with merged cells
  - Multiple sections with different headers
  - Documents with images and embedded objects
  - Unicode and special characters
  - Right-to-left text

### 9. Performance Considerations

- Stream processing for large files
- Lazy loading of document parts
- Efficient XML parsing and generation
- Memory management for images
- Caching of frequently accessed data

### 10. Error Handling

- Validate DOCX structure on load
- Graceful degradation for unsupported features
- Clear error messages with context
- Recovery from corrupted files where possible

### 11. Documentation

- API reference with JSDoc comments
- Architecture documentation
- Examples for common use cases
- Migration guide from other libraries
- Contributing guidelines

## Implementation Phases

### Phase 1: Foundation (COMPLETED)

- ZIP archive handling (ZipHandler, ZipReader, ZipWriter)
- Basic XML generation (XMLBuilder)
- Document structure validation
- Helper methods (14 convenience functions)
- 80 tests passing
- **Status:** Production-ready

### Phase 2: Core Elements (COMPLETED)

- Paragraph class with formatting
- Run class for formatted text spans
- Character formatting (bold, italic, font, color, etc.)
- Paragraph formatting (alignment, indentation, spacing)
- Document creation from scratch
- XML generation for elements
- 46 additional tests
- **Status:** Production-ready with 126 total tests

### Phase 3: Advanced Formatting (COMPLETED)

- Styles implementation (Style, StylesManager)
- Lists and numbering (NumberingLevel, AbstractNumbering, NumberingManager)
- Tables with advanced formatting (Table, TableRow, TableCell)
- Sections (Section with page setup, margins, columns)
- 100+ additional tests
- **Status:** Production-ready with 226+ total tests

### Phase 4: Rich Content (COMPLETED)

- [x] Images and media (PNG, JPEG, GIF, SVG)
- [x] Headers and footers (different first page, odd/even)
- [x] Advanced table features (vertical merge, complex borders)
- [x] Hyperlinks and bookmarks (with defragmentation utility)
- [x] Shapes and text boxes

### Phase 5: Polish (COMPLETED)

- [x] Color normalization (uppercase hex per Microsoft convention)
- [x] ECMA-376 compliance validation (RSIDs, properties order)
- [x] Cell margins support (table formatting)
- [x] Contextual spacing support (paragraph formatting)
- [x] TOC field validation (prevents corruption)
- [x] Track changes support (Revision, RevisionManager)
- [x] Comments (Comment, CommentManager)
- [x] Comprehensive field support (merge, date, page numbers, TOC)
- [x] Footnotes and endnotes
- [x] Content controls (Structured Document Tags)
- [x] Font management
- [x] Drawing manager for shapes and graphics

## Architecture & Design Philosophy

### XML Generation Philosophy: KISS Principle

**The framework follows a lean approach to XML generation:**

1. **No Optimizer Needed**

   - Properties are only serialized if explicitly set
   - Generator checks `if (property)` before adding elements
   - Empty attributes objects aren't included in output
   - Default values are naturally omitted - no special logic needed

   Example:

   ```typescript
   // Paragraph.ts - Already optimal
   if (this.formatting.spacing) {
     // Only build spacing if it has attributes
     if (Object.keys(attributes).length > 0) {
       pPrChildren.push(XMLBuilder.wSelf("spacing", attributes));
     }
   }
   // Result: Lean XML without explicit optimization
   ```

2. **Why Complexity Was Avoided**

   - XMLOptimizer class would add 100+ lines
   - Solves problem that doesn't exist
   - Adds maintenance burden for zero benefit
   - Framework envy / resume-driven development trap
   - **Better to write code you don't need than code you don't use**

3. **RSID Handling (Revision Session IDs)**

   - Framework correctly omits RSIDs for programmatic generation
   - Word regenerates RSIDs automatically on first edit
   - RSIDs only matter for collaborative editing / change tracking
   - Not needed for document generation use case
   - Per ECMA-376: RSIDs are OPTIONAL and may be omitted

4. **Color Handling**

   - All colors normalized to uppercase 6-character hex
   - Supports both 3-char (#F00) and 6-char (#FF0000) formats
   - Automatic expansion and normalization on set/load
   - Aligns with Microsoft conventions
   - See: `src/elements/Run.ts - normalizeColor() method`

5. **Paragraph Property Conflict Resolution** (Added v0.28.2)

   - Automatically prevents `pageBreakBefore` + `keepNext`/`keepLines` conflicts
   - The `pageBreakBefore` property causes massive whitespace in Word when combined with keep properties
   - **Design Decision**: "Keep together" properties take priority over page breaks
   - Rationale: Keep properties represent explicit user intent to keep content together; page breaks are layout hints
   - Implementation: `Paragraph.setKeepNext(true)` and `setKeepLines(true)` automatically clear `pageBreakBefore`
   - Applies during both: creation (API calls) and parsing (loading documents)
   - Files affected:
     - `src/elements/Paragraph.ts:474-507` - Automatic conflict resolution in setKeepNext/setKeepLines
     - `src/core/DocumentParser.ts:646-665` - Parse pageBreakBefore first, then resolve conflicts
   - Test coverage:
     - `tests/elements/Paragraph.test.ts` - 7 tests for setter behavior (all passing)
     - `tests/core/Document.test.ts` - 4 tests for parsing round-trips (all passing)
   - User documentation: `README.md` - Troubleshooting section

   **Why This Matters**:

   - Word's layout engine creates massive whitespace when `pageBreakBefore` conflicts with keep properties
   - The `pageBreakBefore` property is what causes the whitespace, not the keep properties
   - Removing `pageBreakBefore` eliminates whitespace while preserving user's intention to keep content together
   - Common issue when processing documents with complex layouts (discovered via Test4.docx analysis)

   **Discovery Process**:

   - Initial implementation had priority backwards (page breaks cleared keep properties)
   - XML analysis showed `pageBreakBefore` present after "fix" → whitespace persisted
   - Reversed priority: keep properties now clear `pageBreakBefore` → whitespace eliminated
   - Confirmed via Test4.docx: Element 18 now has only keepNext/keepLines (no pageBreakBefore)

   **Philosophy Alignment**:

   - Defensive: Prevents common mistakes automatically
   - Predictable: Clear priority (keep properties win)
   - Evidence-based: Reversed priority based on XML analysis of actual problem documents
   - Documented: Users understand the behavior and rationale
   - Testable: Comprehensive test coverage ensures reliability (596 tests passing)

### Senior Development Principle

**"The best code is the code you don't write"**

Decision framework:

- ✅ Optimize WITH measurement (find actual problems first)
- ✅ Use proven patterns that solve real problems
- ✅ KISS: Simplest solution that works
- ❌ Don't optimize without evidence
- ❌ Don't build frameworks for non-existent problems
- ❌ Don't add complexity "just in case"

**This is how the framework stays lean and maintainable.**

## Resources and References

### Official Specifications

1. **ECMA-376 Office Open XML** - The official standard

   - Part 1: Fundamentals and Markup Language Reference
   - Part 4: Transitional Migration Features

2. **Microsoft Open XML SDK Documentation**
   - Understanding document structure
   - Element reference

### Existing Libraries (for reference)

1. **docx (by dolanmiu)** - Modern JavaScript library
2. **python-docx** - Python implementation (good design patterns)
3. **Open XML SDK** - Official .NET library

### Tools

1. **Open XML SDK Productivity Tool** - Inspect DOCX structure
2. **7-Zip** - View DOCX contents
3. **XML Tree Viewer** - Visualize XML structure

### Key Learning Resources

1. Office Open XML specifications (ISO/IEC 29500)
2. WordprocessingML reference
3. DrawingML for images and graphics
4. Relationship handling in Open XML

## Success Criteria

- [x] Can create Word documents from scratch (Phase 2)
- [x] Can read and modify existing documents (Phase 1)
- [ ] Preserves document structure and formatting (Partial - basic formatting done)
- [x] Compatible with Microsoft Word 2016+ (OpenXML compliant)
- [x] Handles edge cases gracefully (Comprehensive error handling)
- [x] Well-documented API (Complete documentation)
- [x] > 90% test coverage (126 tests covering all modules)
- [ ] Performance: Process 100-page document in <1 second (Not yet tested at scale)

## Common User Mistakes & Troubleshooting

### XML Corruption Issue (October 2025 Analysis)

**Issue**: Users reported documents displaying text like "Important Information<w:t xml:space="preserve">1" in Word, appearing as corrupted XML.

**Root Cause Analysis**: This was determined to be **USER ERROR**, not a framework bug.

#### How It Happens

Users pass XML-like strings to text methods:

```javascript
// WRONG - User code that causes corruption
paragraph.addText('Important Information<w:t xml:space="preserve">1</w:t>');

// What Word displays:
// "Important Information<w:t xml:space="preserve">1"
```

#### Why This Is NOT A Bug

1. **Proper XML Escaping**: Framework correctly escapes special characters per XML spec

   - `<` becomes `&lt;`, `>` becomes `&gt;`, `"` becomes `&quot;`
   - This is REQUIRED by XML standards (ECMA-376)

2. **DOM-Based Generation**: Uses XMLBuilder to create proper element structure

   - Never uses string concatenation
   - All text goes through `escapeXmlText()` function
   - See: `src/xml/XMLBuilder.ts:161-166`

3. **Existing Protection**: Framework already has detection and cleaning:
   - `validateRunText()` - Detects XML patterns and warns
   - `cleanXmlFromText()` - Removes XML patterns
   - `detectCorruptionInDocument()` - Full document scanning
   - See: `src/utils/validation.ts` and `src/utils/corruptionDetection.ts`

#### The Correct Approach

```javascript
// CORRECT - Separate runs
paragraph.addText("Important Information");
paragraph.addText("1");

// Or combined
paragraph.addText("Important Information 1");

// With formatting
paragraph.addText("Important Information", { bold: true });
paragraph.addText("1", { italic: true });
```

#### Detection & Fixing Tools

**Detection Utility** (`src/utils/corruptionDetection.ts`):

```javascript
import { detectCorruptionInDocument } from "docxmlater";

const doc = await Document.load("file.docx");
const report = detectCorruptionInDocument(doc);

if (report.isCorrupted) {
  console.log(report.summary);
  report.locations.forEach((loc) => {
    console.log(`Fix: "${loc.suggestedFix}"`);
  });
}
```

**Auto-Cleaning Option**:

```javascript
// Clean XML patterns automatically
paragraph.addText("Text<w:t>value</w:t>", { cleanXmlFromText: true });
// Result: "Textvalue"
```

#### Files Added for This Issue

- `src/utils/corruptionDetection.ts` - Detection utility (300 lines)
- `tests/utils/corruptionDetection.test.ts` - Comprehensive tests (15 test cases)
- `examples/troubleshooting/xml-corruption.ts` - Common mistake demo
- `examples/troubleshooting/fix-corrupted-document.ts` - Fix utility demo
- `README.md` - Added troubleshooting section

#### Key Takeaway

**This is a user education issue, not a framework bug.** The framework:

1. ✅ Works correctly per XML specifications
2. ✅ Already has detection and cleaning capabilities
3. ✅ Warns users about potential issues
4. ✅ Provides auto-clean option

The solution was better documentation and tooling to help users avoid and fix this common mistake.

## Anti-Goals

- Not a complete Word replacement
- No support for VBA macros
- No support for legacy .doc format
- No GUI/editor component (API only)

## Next Steps

1. Set up project structure and build system
2. Implement ZIP handling and basic document loading
3. Create XML parsing utilities
4. Build core Document class
5. Implement Paragraph and Run classes
6. Add formatting support incrementally
7. Write tests continuously
8. Document as you go

---

## Documentation Guidelines

**Important Rules for All Documentation:**

1. **Never include emojis in any outward-facing documentation** - This includes README.md, release notes, CHANGELOG.md, and any public-facing materials. Keep documentation professional and plain.

2. **Never include talk of AI or Claude in any outward-facing documentation** - Don't mention that code was "generated with Claude" or reference AI tools in public documentation. Keep focus on the actual product and its capabilities.

3. **Internal Development Documentation** - CLAUDE.md files in the project (like this one) can be more informal and include implementation notes, but public-facing materials must remain professional.

**Note**: Focus on clean, maintainable code over feature completeness. Better to have 80% of features working perfectly than 100% working poorly.
