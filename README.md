# docXMLater

A comprehensive, production-ready TypeScript/JavaScript framework for creating, reading, and manipulating Microsoft Word (.docx) documents programmatically.

## Features

### Core Document Operations

- Create DOCX files from scratch
- Read and modify existing DOCX files
- Buffer-based operations (load/save from memory)
- Document properties (core, extended, custom)
- Memory management with dispose pattern

### Text & Paragraph Formatting

- Character formatting: bold, italic, underline, strikethrough, subscript, superscript
- Font properties: family, size, color (RGB and theme colors), highlight
- Text effects: small caps, all caps, shadow, emboss, engrave
- Paragraph alignment, indentation, spacing, borders, shading
- Text search and replace with regex support
- Custom styles (paragraph, character, table)

### Lists & Tables

- Numbered lists (decimal, roman, alpha)
- Bulleted lists with various bullet styles
- Multi-level lists with custom numbering
- Tables with formatting, borders, shading
- Cell spanning (merge cells horizontally and vertically)
- Advanced table properties (margins, widths, alignment)

### Rich Content

- Images (PNG, JPEG, GIF, SVG) with positioning and text wrapping
- Headers & footers (different first page, odd/even pages)
- Hyperlinks (external URLs, internal bookmarks)
- Hyperlink defragmentation utility (fixes fragmented links from Google Docs)
- Bookmarks and cross-references
- Shapes and text boxes

### Advanced Features

- Track changes (revisions for insertions, deletions, formatting)
- Comments and annotations
- Table of contents generation with customizable heading levels
- Fields: merge fields, date/time, page numbers, TOC fields
- Footnotes and endnotes
- Content controls (Structured Document Tags)
- Multiple sections with different page layouts
- Page orientation, size, and margins

### Developer Tools

- Complete XML generation and parsing (ReDoS-safe, position-based parser)
- 40+ unit conversion functions (twips, EMUs, points, pixels, inches, cm)
- Validation utilities and corruption detection
- Full TypeScript support with comprehensive type definitions
- Error handling utilities
- Logging infrastructure with multiple log levels

## Installation

```bash
npm install docxmlater
```

## Quick Start

### Creating a New Document

```typescript
import { Document } from "docxmlater";

// Create a new document
const doc = Document.create();

// Add a paragraph
const para = doc.createParagraph();
para.addText("Hello, World!", { bold: true, fontSize: 24 });

// Save to file
await doc.save("hello.docx");

// Don't forget to dispose
doc.dispose();
```

### Loading and Modifying Documents

```typescript
import { Document } from "docxmlater";

// Load existing document
const doc = await Document.load("input.docx");

// Find and replace text
doc.replaceText(/old text/g, "new text");

// Add a new paragraph
const para = doc.createParagraph();
para.addText("Added paragraph", { italic: true });

// Save modifications
await doc.save("output.docx");
doc.dispose();
```

### Working with Tables

```typescript
import { Document } from "docxmlater";

const doc = Document.create();

// Create a 3x4 table
const table = doc.createTable(3, 4);

// Set header row
const headerRow = table.getRow(0);
headerRow.getCell(0).addParagraph().addText("Column 1", { bold: true });
headerRow.getCell(1).addParagraph().addText("Column 2", { bold: true });
headerRow.getCell(2).addParagraph().addText("Column 3", { bold: true });
headerRow.getCell(3).addParagraph().addText("Column 4", { bold: true });

// Add data
table.getRow(1).getCell(0).addParagraph().addText("Data 1");
table.getRow(1).getCell(1).addParagraph().addText("Data 2");

// Apply borders
table.setBorders({
  top: { style: "single", size: 4, color: "000000" },
  bottom: { style: "single", size: 4, color: "000000" },
  left: { style: "single", size: 4, color: "000000" },
  right: { style: "single", size: 4, color: "000000" },
  insideH: { style: "single", size: 4, color: "000000" },
  insideV: { style: "single", size: 4, color: "000000" },
});

await doc.save("table.docx");
doc.dispose();
```

### Adding Images

```typescript
import { Document } from "docxmlater";
import { readFileSync } from "fs";

const doc = Document.create();

// Load image from file
const imageBuffer = readFileSync("photo.jpg");

// Add image to document
const para = doc.createParagraph();
await para.addImage(imageBuffer, {
  width: 400,
  height: 300,
  format: "jpg",
});

await doc.save("with-image.docx");
doc.dispose();
```

### Hyperlink Management

```typescript
import { Document } from "docxmlater";

const doc = await Document.load("document.docx");

// Get all hyperlinks
const hyperlinks = doc.getHyperlinks();
console.log(`Found ${hyperlinks.length} hyperlinks`);

// Update URLs in batch (30-50% faster than manual iteration)
doc.updateHyperlinkUrls("http://old-domain.com", "https://new-domain.com");

// Fix fragmented hyperlinks from Google Docs
const mergedCount = doc.defragmentHyperlinks({
  resetFormatting: true, // Fix corrupted fonts
});
console.log(`Merged ${mergedCount} fragmented hyperlinks`);

await doc.save("updated.docx");
doc.dispose();
```

### Custom Styles

```typescript
import { Document, Style } from "docxmlater";

const doc = Document.create();

// Create custom paragraph style
const customStyle = new Style("CustomHeading", "paragraph");
customStyle.setName("Custom Heading");
customStyle.setRunFormatting({
  bold: true,
  fontSize: 32,
  color: "0070C0",
});
customStyle.setParagraphFormatting({
  alignment: "center",
  spacingAfter: 240,
});

// Add style to document
doc.getStylesManager().addStyle(customStyle);

// Apply style to paragraph
const para = doc.createParagraph();
para.addText("Styled Heading");
para.applyStyle("CustomHeading");

await doc.save("styled.docx");
doc.dispose();
```

## API Overview

### Document Class

**Creation & Loading:**

- `Document.create(options?)` - Create new document
- `Document.load(filepath)` - Load from file
- `Document.loadFromBuffer(buffer)` - Load from memory

**Content Management:**

- `createParagraph()` - Add paragraph
- `createTable(rows, cols)` - Add table
- `createSection()` - Add section
- `getBodyElements()` - Get all body content

**Search & Replace:**

- `findText(pattern)` - Find text matches
- `replaceText(pattern, replacement)` - Replace text

**Hyperlinks:**

- `getHyperlinks()` - Get all hyperlinks
- `updateHyperlinkUrls(oldUrl, newUrl)` - Batch URL update
- `defragmentHyperlinks(options?)` - Fix fragmented links

**Statistics:**

- `getWordCount()` - Count words
- `getCharacterCount(includeSpaces?)` - Count characters
- `estimateSize()` - Estimate file size

**Saving:**

- `save(filepath)` - Save to file
- `toBuffer()` - Save to Buffer
- `dispose()` - Free resources (important!)

### Paragraph Class

**Content:**

- `addText(text, formatting?)` - Add text run
- `addRun(run)` - Add custom run
- `addHyperlink(hyperlink)` - Add hyperlink
- `addImage(buffer, options)` - Add image

**Formatting:**

- `setAlignment(alignment)` - Left, center, right, justify
- `setIndentation(options)` - First line, hanging, left, right
- `setSpacing(options)` - Line spacing, before/after
- `setBorders(borders)` - Paragraph borders
- `setShading(shading)` - Background color
- `applyStyle(styleId)` - Apply paragraph style

**Properties:**

- `setKeepNext(value)` - Keep with next paragraph
- `setKeepLines(value)` - Keep lines together
- `setPageBreakBefore(value)` - Page break before

**Numbering:**

- `setNumbering(numId, level)` - Apply list numbering

### Run Class

**Text:**

- `setText(text)` - Set run text
- `getText()` - Get run text

**Character Formatting:**

- `setBold(value)` - Bold text
- `setItalic(value)` - Italic text
- `setUnderline(style?)` - Underline
- `setStrikethrough(value)` - Strikethrough
- `setFont(name)` - Font family
- `setFontSize(size)` - Font size in points
- `setColor(color)` - Text color (hex)
- `setHighlight(color)` - Highlight color

**Advanced:**

- `setSubscript(value)` - Subscript
- `setSuperscript(value)` - Superscript
- `setSmallCaps(value)` - Small capitals
- `setAllCaps(value)` - All capitals

### Table Class

**Structure:**

- `addRow()` - Add row
- `getRow(index)` - Get row by index
- `getCell(row, col)` - Get specific cell

**Formatting:**

- `setBorders(borders)` - Table borders
- `setAlignment(alignment)` - Table alignment
- `setWidth(width)` - Table width
- `setLayout(layout)` - Fixed or auto layout

**Style:**

- `applyStyle(styleId)` - Apply table style

### TableCell Class

**Content:**

- `addParagraph()` - Add paragraph to cell
- `getParagraphs()` - Get all paragraphs

**Formatting:**

- `setBorders(borders)` - Cell borders
- `setShading(color)` - Cell background
- `setVerticalAlignment(alignment)` - Top, center, bottom
- `setWidth(width)` - Cell width

**Spanning:**

- `setHorizontalMerge(mergeType)` - Horizontal merge
- `setVerticalMerge(mergeType)` - Vertical merge

### Utilities

**Unit Conversions:**

```typescript
import { twipsToPoints, inchesToTwips, emusToPixels } from "docxmlater";

const points = twipsToPoints(240); // 240 twips = 12 points
const twips = inchesToTwips(1); // 1 inch = 1440 twips
const pixels = emusToPixels(914400, 96); // 914400 EMUs = 96 pixels at 96 DPI
```

**Validation:**

```typescript
import { validateRunText, detectXmlInText, cleanXmlFromText } from "docxmlater";

// Detect XML patterns in text
const result = validateRunText("Some <w:t>text</w:t>");
if (result.hasXml) {
  console.warn(result.message);
  const cleaned = cleanXmlFromText(result.text);
}
```

**Corruption Detection:**

```typescript
import { detectCorruptionInDocument } from "docxmlater";

const doc = await Document.load("suspect.docx");
const report = detectCorruptionInDocument(doc);

if (report.isCorrupted) {
  console.log(`Found ${report.locations.length} corruption issues`);
  report.locations.forEach((loc) => {
    console.log(`Line ${loc.lineNumber}: ${loc.issue}`);
    console.log(`Suggested fix: ${loc.suggestedFix}`);
  });
}
```

## TypeScript Support

Full TypeScript definitions included:

```typescript
import {
  Document,
  Paragraph,
  Run,
  Table,
  RunFormatting,
  ParagraphFormatting,
  DocumentProperties,
} from "docxmlater";

// Type-safe formatting
const formatting: RunFormatting = {
  bold: true,
  fontSize: 12,
  color: "FF0000",
};

// Type-safe document properties
const properties: DocumentProperties = {
  title: "My Document",
  author: "John Doe",
  created: new Date(),
};
```

## Version History

**Current Version: 5.0.0**

See [CHANGELOG.md](CHANGELOG.md) for detailed version history.

## RAG-CLI Integration (Development Only)

This project includes MCP (Model Context Protocol) configuration to allow Claude Code to access docXMLater documentation from Documentation_Hub during development.

**Note:** RAG-CLI uses `python-docx` for DOCX indexing, not docXMLater. These are complementary tools:

- **RAG-CLI**: Index DOCX files for search/retrieval (read-only)
- **docXMLater**: Create, modify, format DOCX files (read-write)

The `.mcp.json` configuration is for development assistance only and does not represent a runtime integration between the two projects.

## Testing

The framework includes comprehensive test coverage:

- **2073+ test cases** across 59 test files
- Tests cover all phases of implementation
- Integration tests for complex scenarios
- Performance benchmarks
- Edge case validation

Run tests:

```bash
npm test              # Run all tests
npm run test:watch   # Watch mode
npm run test:coverage # Coverage report
```

## Performance Considerations

- Use `dispose()` to free resources after document operations
- Buffer-based operations are faster than file I/O
- Batch hyperlink updates are 30-50% faster than manual iteration
- Large documents (1000+ pages) supported with memory management
- Streaming support for very large files

## Architecture

The framework follows a modular architecture:

```
src/
├── core/          # Document, Parser, Generator, Validator
├── elements/      # Paragraph, Run, Table, Image, etc.
├── formatting/    # Style, Numbering managers
├── managers/      # Drawing, Image, Relationship managers
├── xml/           # XML generation and parsing
├── zip/           # ZIP archive handling
└── utils/         # Validation, units, error handling
```

Key design principles:

- KISS (Keep It Simple, Stupid) - no over-engineering
- Position-based XML parsing (ReDoS-safe)
- Defensive programming with comprehensive validation
- Memory-efficient with explicit disposal pattern
- Full ECMA-376 (OpenXML) compliance

## Requirements

- Node.js 18.0.0 or higher
- TypeScript 5.0+ (for development)

## Dependencies

- `jszip` - ZIP archive handling

## License

MIT

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for new features
4. Ensure all tests pass
5. Submit a pull request

## Support

- GitHub Issues: https://github.com/ItMeDiaTech/docXMLater/issues
- Documentation: See CLAUDE.md for detailed implementation notes

## Acknowledgments

Built with careful attention to the ECMA-376 Office Open XML specification. Special thanks to the OpenXML community for comprehensive documentation and examples.
