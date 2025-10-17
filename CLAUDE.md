# DOCX Editing Framework - Project Specification

## Project Overview
Build a comprehensive, production-ready DOCX editing framework from scratch that can create, read, modify, and manipulate Microsoft Word documents programmatically.

## Current Status (Updated: October 2025)

**Phases Completed: 3 of 5**

| Phase | Status | Tests | Features |
|-------|--------|-------|----------|
| **Phase 1: Foundation** | Complete | 80 tests | ZIP handling, XML generation, validation |
| **Phase 2: Core Elements** | Complete | 46 tests | Paragraph, Run, formatting |
| **Phase 3: Advanced Formatting** | Complete | 100+ tests | Styles, tables, sections, lists |
| **Phase 4: Rich Content** | Next | - | Images, headers, footers |
| **Phase 5: Polish** | Planned | - | Track changes, comments, TOC |

**Total: 226+ tests passing | 48 source files | ~10,000+ lines of code**

### What Works Now
- Create DOCX files from scratch
- Read and modify existing DOCX files
- Format text (bold, italic, underline, colors, fonts)
- Format paragraphs (alignment, indentation, spacing)
- Custom styles (paragraph, character, table styles)
- Tables with formatting, borders, shading, and cell spanning
- Section configuration (page size, margins, orientation)
- Multi-level numbered and bulleted lists
- 14 helper methods for file operations
- Complete XML generation
- Full TypeScript support

### Module Documentation
Each module has its own CLAUDE.md file:
- `src/zip/CLAUDE.md` - ZIP archive handling
- `src/xml/CLAUDE.md` - XML generation
- `src/elements/CLAUDE.md` - Paragraph and Run classes
- `src/utils/CLAUDE.md` - Validation utilities

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

### 5. XML Namespaces
Must handle these OpenXML namespaces:
```xml
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
xmlns:v="urn:schemas-microsoft-com:vml"
```

### 6. Unit Conversions
Handle multiple measurement units:
- **Twips**: 1/20th of a point (used for most measurements)
- **EMUs**: English Metric Units (used for images, 914400 EMUs = 1 inch)
- **Points**: Typography unit (72 points = 1 inch)
- **Pixels**: Screen measurement (depends on DPI)
- **Inches/Centimeters**: Human-readable units

### 7. Testing Requirements
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

### 8. Performance Considerations
- Stream processing for large files
- Lazy loading of document parts
- Efficient XML parsing and generation
- Memory management for images
- Caching of frequently accessed data

### 9. Error Handling
- Validate DOCX structure on load
- Graceful degradation for unsupported features
- Clear error messages with context
- Recovery from corrupted files where possible

### 10. Documentation
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

### Phase 4: Rich Content (PLANNED)
- [ ] Images and media
- [ ] Headers and footers
- [ ] Advanced table features
- [ ] Hyperlinks and bookmarks

### Phase 5: Polish (PLANNED)
- [ ] Track changes support
- [ ] Comments
- [ ] Fields and TOC
- [ ] Performance optimization
- [ ] Comprehensive testing

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
- [x] >90% test coverage (126 tests covering all modules)
- [ ] Performance: Process 100-page document in <1 second (Not yet tested at scale)

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

**Note**: Focus on clean, maintainable code over feature completeness. Better to have 80% of features working perfectly than 100% working poorly.
