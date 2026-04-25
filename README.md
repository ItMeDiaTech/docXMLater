# docXMLater

A comprehensive, production-ready TypeScript/JavaScript framework for creating, reading, and manipulating Microsoft Word (.docx) documents programmatically.

## When to Use docxmlater

docxmlater is designed for **editing existing Word documents** with full round-trip XML fidelity. It excels at:

- Loading a .docx, making targeted modifications, and saving without losing formatting or structure
- Working with tracked changes, comments, and revision history
- Preserving complex elements (math equations, charts, SmartArt) through raw XML passthrough
- Programmatic batch processing of corporate documents

If you only need to **generate documents from scratch** and don't need to load/edit existing files, consider the [docx](https://www.npmjs.com/package/docx) package which has a declarative builder API optimized for document creation.

## Features

### Core Document Operations

- Create DOCX files from scratch
- Read and modify existing DOCX files
- Buffer-based operations (load/save from memory)
- Document properties (core, extended, custom)
- Memory management with dispose pattern
- Bookmark pair validation and auto-repair (`validateBookmarkPairs()`)
- App.xml metadata preservation (HeadingPairs, TotalTime, etc.)
- Document background color/theme support

### Text & Paragraph Formatting

- Character formatting: bold, italic, underline, strikethrough, subscript, superscript
- Font properties: family, size, color (RGB and theme colors), highlight
- Text effects: small caps, all caps, shadow, emboss, engrave
- Paragraph alignment, indentation, spacing, borders, shading
- Text search and replace with regex support
- Custom styles (paragraph, character, table)
- CJK/East Asian paragraph properties (kinsoku, wordWrap, overflowPunct, topLinePunct)
- Underline color and theme color attributes
- Theme font references (asciiTheme, hAnsiTheme, eastAsiaTheme, csTheme)

### Lists & Tables

- Numbered lists (decimal, roman, alpha)
- Bulleted lists with various bullet styles
- Multi-level lists with custom numbering and restart control
- Tables with formatting, borders, shading
- Cell spanning (merge cells horizontally and vertically)
- Advanced table properties (margins, widths, alignment)
- Table navigation helpers (`getFirstParagraph()`, `getLastParagraph()`)
- Legacy horizontal merge (`hMerge`) support
- Table layout parsing (`fixed`/`auto`)
- Table style shading updates (modify styles.xml colors)
- Cell content management (trailing blank removal with structure preservation)

### Rich Content

- Images (PNG, JPEG, GIF, SVG, EMF, WMF) with positioning, text wrapping, and full ECMA-376 DrawingML attribute coverage
- Headers & footers (different first page, odd/even pages)
- Hyperlinks (external URLs, internal bookmarks)
- Hyperlink defragmentation utility (fixes fragmented links from Google Docs)
- Hyperlink URL sanitization (strips browser extension prefixes from corrupted URLs)
- Bookmarks and cross-references
- Body-level bookmark support (bookmarks between block elements)
- Shapes and text boxes

### Advanced Features

- Track changes (revisions for insertions, deletions, formatting)
- Granular character-level tracked changes (text diff-based)
- Comments and annotations
- Compatibility mode detection and upgrade (Word 2003/2007/2010/2013+ modes)
- Table of contents generation with customizable heading levels and relative indentation
- Fields: merge fields, date/time, page numbers, TOC fields
- Footnotes and endnotes (full round-trip with save pipeline, parsing, and clear API)
- Content controls (Structured Document Tags)
- Form field data preservation (text input, checkbox, dropdown per ECMA-376 §17.16)
- w14 run effects passthrough (Word 2010+ ligatures, numForm, textOutline, etc.)
- Expanded document settings (evenAndOddHeaders, mirrorMargins, autoHyphenation, decimalSymbol)
- People.xml auto-registration for tracked changes authors
- Style default attribute preservation (`w:default="1"`)
- Namespace order preservation in generated XML
- Multiple sections with different page layouts
- Page orientation, size, and margins
- Preserved element round-trip (math equations, alternate content, custom XML)
- Unified shading model with theme color support and inheritance resolution
- Lossless image optimization (PNG re-compression, BMP-to-PNG conversion)
- Run property change tracking (w:rPrChange) with direct API access
- Paragraph mark revision tracking (w:del/w:ins in w:pPr/w:rPr) for full tracked-changes fidelity
- Normal/NormalWeb style linking with preservation flags

### Developer Tools

- Complete XML generation and parsing (ReDoS-safe, position-based parser)
- 40+ unit conversion functions (twips, EMUs, points, pixels, inches, cm)
- Validation utilities and corruption detection
- Text diff utility for character-level comparisons
- webSettings.xml auto-generation
- Safe OOXML parsing helpers (zero-value handling, boolean parsing)
- Full TypeScript support with comprehensive type definitions
- Error handling utilities with custom error types (`DocxError`, `InvalidDocxError`, `CorruptedArchiveError`)
- Logging infrastructure with multiple log levels (`DOCXMLATER_LOG_LEVEL=debug|info|warn|error`)
- Plain text extraction (`doc.toPlainText()`) and heading hierarchy (`doc.getHeadingHierarchy()`)
- Accessibility auditing (`doc.findImagesWithoutAltText()`)

### Unsupported OOXML Features

The following features are preserved as raw XML on round-trip but have no editing API:

- **Charts** (c:chartSpace) -- preserved but not editable
- **SmartArt** -- preserved as raw XML passthrough
- **OLE embedded objects** (`<w:object>`) -- preserved, no API
- **Glossary document** (glossary.xml) -- not handled
- **DrawingML advanced features** -- gradient fills, pattern fills, group shapes, 3D effects, shape effects (shadow, reflection, glow)

## Installation

```bash
npm install docxmlater
```

## Quick Start

### Creating a New Document

```typescript
import { Document } from 'docxmlater';

// Create a new document
const doc = Document.create();

// Add a paragraph
const para = doc.createParagraph();
para.addText('Hello, World!', { bold: true, fontSize: 24 });

// Save to file
await doc.save('hello.docx');

// Don't forget to dispose
doc.dispose();
```

### Loading and Modifying Documents

```typescript
import { Document } from 'docxmlater';

// Load existing document
const doc = await Document.load('input.docx');

// Find and replace text
doc.replaceText(/old text/g, 'new text');

// Add a new paragraph
const para = doc.createParagraph();
para.addText('Added paragraph', { italic: true });

// Save modifications
await doc.save('output.docx');
doc.dispose();
```

### Working with Tables

```typescript
import { Document } from 'docxmlater';

const doc = Document.create();

// Create a 3x4 table
const table = doc.createTable(3, 4);

// Set header row
const headerRow = table.getRow(0);
headerRow.getCell(0).addParagraph().addText('Column 1', { bold: true });
headerRow.getCell(1).addParagraph().addText('Column 2', { bold: true });
headerRow.getCell(2).addParagraph().addText('Column 3', { bold: true });
headerRow.getCell(3).addParagraph().addText('Column 4', { bold: true });

// Add data
table.getRow(1).getCell(0).addParagraph().addText('Data 1');
table.getRow(1).getCell(1).addParagraph().addText('Data 2');

// Apply borders
table.setBorders({
  top: { style: 'single', size: 4, color: '000000' },
  bottom: { style: 'single', size: 4, color: '000000' },
  left: { style: 'single', size: 4, color: '000000' },
  right: { style: 'single', size: 4, color: '000000' },
  insideH: { style: 'single', size: 4, color: '000000' },
  insideV: { style: 'single', size: 4, color: '000000' },
});

await doc.save('table.docx');
doc.dispose();
```

### Adding Images

```typescript
import { Document } from 'docxmlater';
import { readFileSync } from 'fs';

const doc = Document.create();

// Load image from file
const imageBuffer = readFileSync('photo.jpg');

// Add image to document
const para = doc.createParagraph();
await para.addImage(imageBuffer, {
  width: 400,
  height: 300,
  format: 'jpg',
});

await doc.save('with-image.docx');
doc.dispose();
```

### Hyperlink Management

```typescript
import { Document } from 'docxmlater';

const doc = await Document.load('document.docx');

// Get all hyperlinks
const hyperlinks = doc.getHyperlinks();
console.log(`Found ${hyperlinks.length} hyperlinks`);

// Update URLs in batch (30-50% faster than manual iteration)
doc.updateHyperlinkUrls('http://old-domain.com', 'https://new-domain.com');

// Fix fragmented hyperlinks from Google Docs
const mergedCount = doc.defragmentHyperlinks({
  resetFormatting: true, // Fix corrupted fonts
});
console.log(`Merged ${mergedCount} fragmented hyperlinks`);

await doc.save('updated.docx');
doc.dispose();
```

### Custom Styles

```typescript
import { Document, Style } from 'docxmlater';

const doc = Document.create();

// Create custom paragraph style
const customStyle = new Style('CustomHeading', 'paragraph');
customStyle.setName('Custom Heading');
customStyle.setRunFormatting({
  bold: true,
  fontSize: 32,
  color: '0070C0',
});
customStyle.setParagraphFormatting({
  alignment: 'center',
  spacingAfter: 240,
});

// Add style to document
doc.getStylesManager().addStyle(customStyle);

// Apply style to paragraph
const para = doc.createParagraph();
para.addText('Styled Heading');
para.applyStyle('CustomHeading');

await doc.save('styled.docx');
doc.dispose();
```

### Compatibility Mode Detection and Upgrade

```typescript
import { Document, CompatibilityMode } from 'docxmlater';

const doc = await Document.load('legacy.docx');

// Check compatibility mode
console.log(`Mode: ${doc.getCompatibilityMode()}`); // e.g., 12 (Word 2007)

if (doc.isCompatibilityMode()) {
  // Get detailed compatibility info
  const info = doc.getCompatibilityInfo();
  console.log(`Legacy flags: ${info.legacyFlags.length}`);

  // Upgrade to Word 2013+ mode (equivalent to File > Info > Convert)
  const report = doc.upgradeToModernFormat();
  console.log(`Removed ${report.removedFlags.length} legacy flags`);
  console.log(`Added ${report.addedSettings.length} modern settings`);
}

await doc.save('modern.docx');
doc.dispose();
```

## API Overview

### Document Class

**Creation & Loading:**

- `Document.create(options?)` - Create new document
- `Document.load(filepath, options?)` - Load from file
- `Document.loadFromBuffer(buffer, options?)` - Load from memory

**Handling Tracked Changes:**

By default, docXMLater accepts all tracked changes during document loading to prevent corruption:

```typescript
// Default: Accepts all changes (recommended)
const doc = await Document.load('document.docx');

// Explicit control
const doc = await Document.load('document.docx', {
  revisionHandling: 'accept'  // Accept all changes (default)
  // OR
  revisionHandling: 'strip'   // Remove all revision markup
  // OR
  revisionHandling: 'preserve' // Keep tracked changes (may cause corruption, but should not do so - report errors if found)
});
```

**Revision Handling Options:**

- `'accept'` (default): Removes revision markup, keeps inserted content, removes deleted content
- `'strip'`: Removes all revision markup completely
- `'preserve'`: Keeps tracked changes as-is (may cause Word "unreadable content" errors)

**Why Accept By Default?**

Documents with tracked changes can cause Word corruption errors during round-trip processing due to revision ID conflicts. Accepting changes automatically prevents this issue while preserving document content.

**Content Management:**

- `createParagraph()` - Add paragraph
- `createTable(rows, cols)` - Add table
- `createSection()` - Add section
- `getBodyElements()` - Get all body content

**Search & Replace:**

- `findText(pattern)` - Find text matches
- `replaceText(pattern, replacement)` - Replace text
- `findParagraphsByText(pattern)` - Find paragraphs containing text/regex
- `getParagraphsByStyle(styleId)` - Get paragraphs with specific style
- `getRunsByFont(fontName)` - Get runs using a specific font
- `getRunsByColor(color)` - Get runs with a specific color

**Bulk Formatting:**

- `setAllRunsFont(fontName)` - Apply font to all text
- `setAllRunsSize(size)` - Apply font size to all text
- `setAllRunsColor(color)` - Apply color to all text
- `getFormattingReport()` - Get document formatting statistics

**Hyperlinks:**

- `getHyperlinks()` - Get all hyperlinks
- `updateHyperlinkUrls(oldUrl, newUrl)` - Batch URL update
- `defragmentHyperlinks(options?)` - Fix fragmented links
- `collectAllReferencedHyperlinkIds()` - Comprehensive scan of all hyperlink relationship IDs (includes nested tables, headers/footers, footnotes/endnotes)

**Statistics:**

- `getWordCount()` - Count words
- `getCharacterCount(includeSpaces?)` - Count characters
- `estimateSize()` - Estimate file size

**Compatibility Mode:**

- `getCompatibilityMode()` - Get document's Word version mode (11/12/14/15)
- `isCompatibilityMode()` - Check if document targets a legacy Word version
- `getCompatibilityInfo()` - Get full parsed compat settings
- `upgradeToModernFormat()` - Upgrade to Word 2013+ mode (removes legacy flags)

**Footnotes & Endnotes:**

- `createFootnote(paragraph, text)` - Add footnote
- `createEndnote(paragraph, text)` - Add endnote
- `clearFootnotes()` / `clearEndnotes()` - Remove all notes
- `getFootnoteManager()` / `getEndnoteManager()` - Access note managers

**Numbering:**

- `restartNumbering(numId, level?, startValue?)` - Restart list numbering (creates new instance with startOverride)
- `cleanupUnusedNumbering()` - Remove unused numbering definitions (scans body, headers, footers, footnotes, endnotes)
- `consolidateNumbering(options?)` - Merge duplicate abstract numbering definitions
- `validateNumberingReferences()` - Fix orphaned numId references

**Shading:**

- `getComputedCellShading(table, row, col)` - Resolve effective cell shading with inheritance

**Document Sanitization:**

- `flattenFieldCodes()` - Strip INCLUDEPICTURE field markup, preserving embedded images
- `stripOrphanRSIDs()` - Remove orphan RSIDs from settings.xml
- `clearDirectSpacingForStyles(styleIds)` - Remove direct spacing overrides from styled paragraphs

**Image Optimization:**

- `optimizeImages()` - Lossless PNG re-compression and BMP-to-PNG conversion (zero dependencies)

**Document Convenience:**

- `addHeading(text, level?)` - Add heading paragraph (H1-H9)
- `addPageBreak()` - Insert page break
- `addHorizontalRule(color?, size?)` - Insert horizontal line
- `setDefaultFont(name, size?)` - Set document default font via Normal style
- `setDefaultFontSize(size)` - Set document default font size
- `clear()` - Remove all body content (preserves styles/settings)
- `clone()` - Deep copy document for template batch generation
- `addBulletListFromArray(items)` - Create bullet list from string array
- `addNumberedListFromArray(items)` - Create numbered list from string array
- `createTableFromCSV(csv, delimiter?)` - Create table from CSV data

**Template Engine:**

- `fillTemplate(data, options?)` - Replace `{{key}}` placeholders across runs
- `findAndHighlight(text, color?)` - Highlight all text occurrences
- `findAndFormat(text, formatting)` - Apply formatting to all text occurrences

**Document Conversion:**

- `toMarkdown()` - Export as Markdown
- `toHTML(options?)` - Export as HTML (fragment or full page)
- `toBase64()` - Export as base64 string
- `toDataUri()` - Export as data URI
- `fromMarkdown(md)` - Create document from Markdown (static)
- `loadFromBase64(base64)` - Load document from base64 (static)

**Content Structure:**

- `insertAfter(reference, element)` - Insert element after reference
- `insertBefore(reference, element)` - Insert element before reference
- `replaceElement(old, new)` - Replace body element in-place
- `removeElement(element)` - Remove body element by reference
- `extractByHeading(maxLevel?)` - Group content by heading sections
- `getElementsBetween(start, end)` - Get elements between two references
- `forEachParagraph(callback)` - Iterate top-level paragraphs
- `forEachTable(callback)` - Iterate top-level tables
- `getStatistics()` - Comprehensive document metrics

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
- `clearSpacing()` - Remove direct spacing (inherit from style)

**Text Manipulation:**

- `applyFormattingToRange(start, end, formatting)` - Apply formatting to character range
- `deleteRange(start, end)` - Delete character range
- `truncate(maxLength, suffix?)` - Truncate text with ellipsis
- `wrap(prefix, suffix, formatting?)` - Wrap content with prefix/suffix
- `splitAt(offset)` - Split paragraph into two at character position
- `consolidateRuns()` - Merge adjacent runs with identical formatting
- `replaceAll(find, replace)` - Cross-run find and replace
- `findTextCrossRun(find)` - Cross-run text search with offsets
- `getRunAtOffset(offset)` - Get run at character position
- `getFormattingAtOffset(offset)` - Get formatting at character position
- `contains(text, caseSensitive?)` - Check if paragraph contains text
- `toJSON()` / `fromJSON(data)` - Serialize/deserialize paragraph

**Numbering:**

- `setNumbering(numId, level)` - Apply list numbering

### Run Class

**Text:**

- `setText(text)` - Set run text
- `getText()` - Get run text
- `getPlainText()` - Get text only (no tabs/breaks)
- `splitAt(offset)` - Split run at character position

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
- `clearMatchingFormatting(styleFormatting)` - Remove formatting matching a style (for inheritance)
- `equals(other)` - Compare text and formatting equality
- `hasSameFormatting(other)` - Compare formatting only
- `clone()` - Deep copy run

### Table Class

**Structure:**

- `addRow()` - Add row
- `addRowFromArray(cells)` - Add row from string array
- `getRow(index)` - Get row by index
- `getCell(row, col)` - Get specific cell
- `setCell(row, col, text)` - Set cell text by coordinates
- `duplicateRow(index, count?)` - Clone a row in-place
- `addSummaryRow(options?)` - Add computed totals row

**Data Conversion:**

- `fromArray(data)` / `toArray()` - 2D string array I/O
- `fromCSV(csv, delimiter?)` / `toCSV(delimiter?)` - CSV round-trip
- `toPlainText(colSep?, rowSep?)` - Delimited text export
- `transpose()` - Swap rows and columns
- `clone()` - Deep copy table

**Queries:**

- `getColumnCells(colIndex)` - Get cells in a column
- `getColumnTexts(colIndex)` - Get text values in a column
- `findCell(predicate)` - Find first matching cell with coordinates
- `filterRows(predicate)` - Get indices of matching rows
- `forEachCell(callback)` - Iterate all cells with row/col
- `mapColumn(colIndex, transform)` - Transform column values

**Cleanup:**

- `removeEmptyRows()` - Remove rows with no text
- `removeEmptyColumns()` - Remove columns with no text

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
- `setShading(shading)` - Cell shading/background
- `setBackgroundColor(hex)` / `getBackgroundColor()` - Simple color shortcut
- `setVerticalAlignment(alignment)` - Top, center, bottom
- `setWidth(width)` - Cell width

**Spanning:**

- `setHorizontalMerge(mergeType)` - Horizontal merge
- `setVerticalMerge(mergeType)` - Vertical merge

**Convenience Methods:**

- `setTextAlignment(alignment)` - Set alignment for all paragraphs
- `setAllParagraphsStyle(styleId)` - Apply style to all paragraphs
- `setAllRunsFont(fontName)` - Apply font to all runs
- `setAllRunsSize(size)` - Apply font size to all runs
- `setAllRunsColor(color)` - Apply color to all runs

**Content Management:**

- `removeTrailingBlankParagraphs(options?)` - Remove trailing blank paragraphs from cell
- `removeParagraph(index)` - Remove paragraph at index (updates nested content positions)
- `addParagraphAt(index, paragraph)` - Insert paragraph at index (updates nested content positions)

### Document Class

**Table Style Shading:**

- `updateTableStyleShading(oldColor, newColor)` - Update shading colors in styles.xml
- `updateTableStyleShadingBulk(settings)` - Bulk update table style shading
- `removeTrailingBlanksInTableCells(options?)` - Remove trailing blanks from all table cells

### Table Class

**Sorting:**

- `sortRows(columnIndex, options?)` - Sort table rows by column

### Section Class

**Line Numbering:**

- `setLineNumbering(options)` - Enable line numbering
- `getLineNumbering()` - Get line numbering settings
- `clearLineNumbering()` - Disable line numbering

### Comment Class

**Resolution:**

- `resolve()` - Mark comment as resolved
- `unresolve()` - Mark comment as unresolved
- `isResolved()` - Check if comment is resolved

### CommentManager Class

**Filtering:**

- `getResolvedComments()` - Get all resolved comments
- `getUnresolvedComments()` - Get all unresolved comments

### Utilities

**Unit Conversions:**

```typescript
import { twipsToPoints, inchesToTwips, emusToPixels } from 'docxmlater';

const points = twipsToPoints(240); // 240 twips = 12 points
const twips = inchesToTwips(1); // 1 inch = 1440 twips
const pixels = emusToPixels(914400, 96); // 914400 EMUs = 96 pixels at 96 DPI
```

**Validation:**

```typescript
import { validateRunText, detectXmlInText, cleanXmlFromText } from 'docxmlater';

// Detect XML patterns in text
const result = validateRunText('Some <w:t>text</w:t>');
if (result.hasXml) {
  console.warn(result.message);
  const cleaned = cleanXmlFromText(result.text);
}
```

**Corruption Detection:**

```typescript
import { detectCorruptionInDocument } from 'docxmlater';

const doc = await Document.load('suspect.docx');
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
} from 'docxmlater';

// Type-safe formatting
const formatting: RunFormatting = {
  bold: true,
  fontSize: 12,
  color: 'FF0000',
};

// Type-safe document properties
const properties: DocumentProperties = {
  title: 'My Document',
  author: 'John Doe',
  created: new Date(),
};
```

## Version History

**Current Version: 11.0.2**

See [CHANGELOG.md](CHANGELOG.md) for detailed version history.

## Testing

The framework includes comprehensive test coverage:

- **4,134 test cases** across 195 test suites
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

## Error Handling

All document operations should be wrapped in try/finally to ensure proper cleanup:

```typescript
import { Document } from 'docxmlater';

let doc: Document | undefined;
try {
  doc = await Document.load('input.docx');
  doc.replaceText(/draft/gi, 'final');
  await doc.save('output.docx');
} catch (error) {
  console.error('Document operation failed:', error);
} finally {
  doc?.dispose();
}
```

For buffer-based workflows (common in web servers):

```typescript
async function processDocument(inputBuffer: Buffer): Promise<Buffer> {
  const doc = await Document.loadFromBuffer(inputBuffer);
  try {
    doc.replaceText(/placeholder/g, 'actual value');
    return await doc.toBuffer();
  } finally {
    doc.dispose();
  }
}
```

Custom error types are available from `docxmlater/internal` — `InvalidDocxError`, `CorruptedArchiveError`, and `FileOperationError` all extend `DocxError`.

## Working with Large Documents

- Use buffer operations (`loadFromBuffer`/`toBuffer`) for 20-30% faster I/O
- Call `dispose()` promptly to release ZIP handles and image buffers
- Size limits default to warning at 50MB and error at 150MB (configurable via `LoadOptions.sizeLimits`)
- Memory usage: ~2MB base per Document, ~2 bytes/char, full buffer per embedded image, ~200 bytes/cell
- For repeated paragraph access, cache the result of `getAllParagraphs()` rather than calling it in a loop

## Architecture

The framework follows a modular architecture:

```
src/
├── core/          # Document, Parser, Generator, Validator
├── elements/      # Paragraph, Run, Table, Image, etc.
├── formatting/    # Style, Numbering managers
├── managers/      # Drawing, Image, Relationship managers
├── constants/     # Compatibility mode constants, limits
├── types/         # Type definitions (compatibility, formatting, lists)
├── tracking/      # Change tracking context
├── validation/    # Revision validation rules
├── helpers/       # Cleanup utilities
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

## Security

docXMLater includes multiple security measures to protect against common attack vectors:

### ReDoS Prevention

The XML parser uses position-based parsing instead of regular expressions, preventing catastrophic backtracking attacks that can cause denial of service.

### Input Validation

**Size Limits:**

- Default document size limit: 150 MB (configurable)
- Warning threshold: 50 MB
- XML content size validation before parsing

```typescript
// Configure size limits
const doc = await Document.load('large.docx', {
  sizeLimits: {
    warningSizeMB: 100,
    maxSizeMB: 500,
  },
});
```

**Nesting Depth:**

- Maximum XML nesting depth: 256 (configurable)
- Prevents stack overflow attacks

```typescript
import { XMLParser } from 'docxmlater/internal';

// Parse with custom depth limit
const obj = XMLParser.parseToObject(xml, {
  maxNestingDepth: 512, // Increase if needed
});
```

### Path Traversal Prevention

File paths within DOCX archives are validated to prevent directory traversal attacks:

- Blocks `../` path sequences
- Blocks absolute paths
- Validates URL-encoded path components

### XML Injection Prevention

All text content is properly escaped using:

- `XMLBuilder.escapeXmlText()` for element content
- `XMLBuilder.escapeXmlAttribute()` for attribute values

This prevents injection of malicious XML elements through user-provided text content.

### UTF-8 Encoding

All text files are explicitly UTF-8 encoded per ECMA-376 specification, preventing encoding-related vulnerabilities.

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

## Acknowledgments

Built with careful attention to the ECMA-376 Office Open XML specification. Special thanks to the OpenXML community for comprehensive documentation and examples.
