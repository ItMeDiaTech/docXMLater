# docxmlater

A production-ready TypeScript framework for creating, reading, editing, and manipulating Microsoft Word `.docx` documents with full ECMA-376 (Office Open XML) fidelity.

---

## About the Project

docxmlater began in early 2025 as a personal effort to build a TypeScript framework capable of full programmatic interaction with `.docx` files. What started as a focused side project grew into a much larger undertaking as the depth of the OOXML specification revealed itself. The work is implemented directly against the 6,000+ page ECMA-376 standard, with attention paid to round-trip fidelity, schema correctness, and the practical edge cases real-world Word documents introduce.

The library is in active production use on a small team for day-to-day document formatting workflows. The aim is to provide a free, capable alternative to commercial DOCX engines that charge thousands of dollars per year per seat. If you need a TypeScript library to read, edit, or manipulate Word documents, docxmlater is designed to be a complete solution rather than a thin wrapper.

What distinguishes docxmlater from existing libraries is its first-class support for revision workflows. Tracked changes, comments, and bookmarks are fully integrated. Documents that already contain tracked changes can be processed without corruption, preserving the existing revision history where required while still applying new formatting on top.

If you encounter a use case that is not yet implemented and would be broadly useful, please open an issue.

---

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Feature Overview](#feature-overview)
- [API Reference](#api-reference)
  - [Document](#document)
  - [Paragraph](#paragraph)
  - [Run](#run)
  - [Table](#table)
  - [TableCell](#tablecell)
  - [Section](#section)
  - [Comment & CommentManager](#comment--commentmanager)
  - [Utilities](#utilities)
- [Advanced Topics](#advanced-topics)
  - [Tracked Changes](#tracked-changes)
  - [Custom Styles](#custom-styles)
  - [Hyperlink Management](#hyperlink-management)
  - [Compatibility Mode](#compatibility-mode)
  - [Templates](#templates)
  - [Document Conversion](#document-conversion)
- [Performance & Memory Management](#performance--memory-management)
- [Architecture](#architecture)
- [Security](#security)
- [TypeScript Support](#typescript-support)
- [Requirements](#requirements)
- [Contributing](#contributing)
- [License](#license)

---

## Installation

```bash
npm install docxmlater
```

Requires Node.js **18.0.0** or higher. TypeScript 5.0+ is recommended for development.

The only runtime dependency is `jszip` for ZIP archive handling.

---

## Quick Start

### Create a new document

```typescript
import { Document } from 'docxmlater';

const doc = Document.create();
const para = doc.createParagraph();
para.addText('Hello, World!', { bold: true, fontSize: 24 });

await doc.save('hello.docx');
doc.dispose();
```

### Load and modify an existing document

```typescript
import { Document } from 'docxmlater';

const doc = await Document.load('input.docx');

doc.replaceText(/old text/g, 'new text');
doc.createParagraph().addText('Added paragraph', { italic: true });

await doc.save('output.docx');
doc.dispose();
```

### Tables

```typescript
const doc = Document.create();
const table = doc.createTable(3, 4);

const header = table.getRow(0);
header.getCell(0).addParagraph().addText('Column 1', { bold: true });
header.getCell(1).addParagraph().addText('Column 2', { bold: true });

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

### Images

```typescript
import { Document } from 'docxmlater';
import { readFileSync } from 'fs';

const doc = Document.create();
const imageBuffer = readFileSync('photo.jpg');

const para = doc.createParagraph();
await para.addImage(imageBuffer, { width: 400, height: 300, format: 'jpg' });

await doc.save('with-image.docx');
doc.dispose();
```

---

## Feature Overview

### Document Operations

Create, load, and save documents from files or buffers. Manage core, extended, and custom document properties. Validate and auto-repair bookmark pairs. Preserve `app.xml` metadata (HeadingPairs, TotalTime, etc.). Configurable document background color and theme support.

### Text & Paragraph Formatting

- Character formatting: bold, italic, underline, strikethrough, sub/superscript, small caps, all caps, shadow, emboss, engrave
- Font properties: family, size, color (RGB and theme), highlight, underline color
- Theme font references (`asciiTheme`, `hAnsiTheme`, `eastAsiaTheme`, `csTheme`)
- Paragraph alignment, indentation, spacing, borders, shading
- CJK / East Asian properties (kinsoku, wordWrap, overflowPunct, topLinePunct)
- Cross-run text search and replace, regex supported

### Lists & Tables

- Numbered (decimal, roman, alpha) and bulleted lists
- Multi-level lists with custom numbering and restart control
- Tables with borders, shading, alignment, and width control
- Horizontal and vertical cell merging, including legacy `hMerge`
- Fixed and auto table layouts
- Cell content management with structure preservation

### Rich Content

- Images: PNG, JPEG, GIF, SVG, EMF, WMF - with positioning, text wrapping, and full DrawingML attribute coverage
- Headers and footers with first-page and odd/even variants
- Hyperlinks (external and internal), with defragmentation and URL sanitization utilities
- Bookmarks (block and inline level) and cross-references
- Shapes and text boxes

### Revisions & Collaboration

- Track changes (insertions, deletions, formatting)
- Character-level granular revisions via text diffing
- Comments with resolve/unresolve workflow
- Run property change tracking (`w:rPrChange`)
- Paragraph mark revision tracking (`w:del`/`w:ins` in `w:pPr`/`w:rPr`)
- People.xml auto-registration for revision authors
- Full round-trip preservation of pre-existing tracked changes

### Advanced Features

- Compatibility mode detection and upgrade (Word 2003 / 2007 / 2010 / 2013+)
- Table of contents generation with customizable heading levels
- Fields: merge fields, date/time, page numbers, TOC fields
- Footnotes and endnotes (full round-trip and dedicated API)
- Content controls (Structured Document Tags)
- Form field data preservation (text, checkbox, dropdown per ECMA-376 §17.16)
- `w14` run effects passthrough (Word 2010+ ligatures, numForm, textOutline)
- Multiple sections with independent page layouts and orientations
- Lossless image optimization (PNG re-compression, BMP-to-PNG conversion)
- Unified shading model with theme color support and inheritance resolution

### Document Conversion

Export to Markdown, HTML (fragment or full page), Base64, or Data URI. Create documents from Markdown.

### Preserved (round-trip only)

The following features round-trip safely as raw XML but have no editing API:

- Charts (`c:chartSpace`)
- SmartArt
- OLE embedded objects (`w:object`)
- Math equations
- Glossary documents (`glossary.xml`)
- Advanced DrawingML (gradient/pattern fills, group shapes, 3D effects)

---

## API Reference

### Document

**Creation & Loading**

| Method                                      | Description                 |
| ------------------------------------------- | --------------------------- |
| `Document.create(options?)`                 | Create a new document       |
| `Document.load(path, options?)`             | Load from a file path       |
| `Document.loadFromBuffer(buffer, options?)` | Load from a `Buffer`        |
| `Document.fromMarkdown(md)`                 | Create from Markdown source |
| `Document.loadFromBase64(b64)`              | Load from a Base64 string   |

**Content Management**

- `createParagraph()`, `createTable(rows, cols)`, `createSection()`
- `addHeading(text, level?)`, `addPageBreak()`, `addHorizontalRule(color?, size?)`
- `addBulletListFromArray(items)`, `addNumberedListFromArray(items)`
- `createTableFromCSV(csv, delimiter?)`
- `getBodyElements()`, `clear()`, `clone()`
- `insertAfter(ref, el)`, `insertBefore(ref, el)`, `replaceElement(old, new)`, `removeElement(el)`
- `forEachParagraph(cb)`, `forEachTable(cb)`, `extractByHeading(maxLevel?)`, `getElementsBetween(start, end)`

**Search & Replace**

- `findText(pattern)`, `replaceText(pattern, replacement)`
- `findParagraphsByText(pattern)`, `getParagraphsByStyle(styleId)`
- `getRunsByFont(name)`, `getRunsByColor(color)`

**Bulk Formatting**

- `setAllRunsFont(name)`, `setAllRunsSize(size)`, `setAllRunsColor(color)`
- `setDefaultFont(name, size?)`, `setDefaultFontSize(size)`
- `getFormattingReport()`

**Hyperlinks**

- `getHyperlinks()`, `updateHyperlinkUrls(oldUrl, newUrl)`
- `defragmentHyperlinks(options?)`, `collectAllReferencedHyperlinkIds()`

**Statistics**

- `getWordCount()`, `getCharacterCount(includeSpaces?)`
- `estimateSize()`, `getStatistics()`

**Compatibility**

- `getCompatibilityMode()`, `isCompatibilityMode()`
- `getCompatibilityInfo()`, `upgradeToModernFormat()`

**Footnotes & Endnotes**

- `createFootnote(paragraph, text)`, `createEndnote(paragraph, text)`
- `clearFootnotes()`, `clearEndnotes()`
- `getFootnoteManager()`, `getEndnoteManager()`

**Numbering**

- `restartNumbering(numId, level?, startValue?)`
- `cleanupUnusedNumbering()`, `consolidateNumbering(options?)`
- `validateNumberingReferences()`

**Sanitization & Optimization**

- `flattenFieldCodes()` - strip INCLUDEPICTURE markup, keep images
- `stripOrphanRSIDs()` - remove unused RSIDs from `settings.xml`
- `clearDirectSpacingForStyles(ids)` - remove direct spacing on styled paragraphs
- `optimizeImages()` - lossless PNG re-compression, BMP-to-PNG

**Templates & Highlighting**

- `fillTemplate(data, options?)` - replace `{{key}}` placeholders across runs
- `findAndHighlight(text, color?)`, `findAndFormat(text, formatting)`

**Conversion**

- `toMarkdown()`, `toHTML(options?)`, `toPlainText()`
- `toBase64()`, `toDataUri()`, `getHeadingHierarchy()`
- `findImagesWithoutAltText()` (accessibility audit)

**Saving**

- `save(path)`, `toBuffer()`, `dispose()` - _always call `dispose()` when finished_

### Paragraph

**Content**: `addText(text, formatting?)`, `addRun(run)`, `addHyperlink(link)`, `addImage(buffer, options)`

**Formatting**: `setAlignment`, `setIndentation`, `setSpacing`, `setBorders`, `setShading`, `applyStyle`, `setKeepNext`, `setKeepLines`, `setPageBreakBefore`, `clearSpacing`

**Text manipulation**: `applyFormattingToRange`, `deleteRange`, `truncate`, `wrap`, `splitAt`, `consolidateRuns`, `replaceAll`, `findTextCrossRun`, `getRunAtOffset`, `getFormattingAtOffset`, `contains`, `toJSON` / `fromJSON`

**Numbering**: `setNumbering(numId, level)`

### Run

**Text**: `setText`, `getText`, `getPlainText`, `splitAt`

**Character formatting**: `setBold`, `setItalic`, `setUnderline`, `setStrikethrough`, `setFont`, `setFontSize`, `setColor`, `setHighlight`

**Advanced**: `setSubscript`, `setSuperscript`, `setSmallCaps`, `setAllCaps`, `clearMatchingFormatting`, `equals`, `hasSameFormatting`, `clone`

### Table

**Structure**: `addRow`, `addRowFromArray`, `getRow`, `getCell`, `setCell`, `duplicateRow`, `addSummaryRow`

**Data**: `fromArray` / `toArray`, `fromCSV` / `toCSV`, `toPlainText`, `transpose`, `clone`, `sortRows`

**Queries**: `getColumnCells`, `getColumnTexts`, `findCell`, `filterRows`, `forEachCell`, `mapColumn`

**Formatting**: `setBorders`, `setAlignment`, `setWidth`, `setLayout`, `applyStyle`

**Cleanup**: `removeEmptyRows`, `removeEmptyColumns`

### TableCell

**Content**: `addParagraph`, `getParagraphs`, `removeTrailingBlankParagraphs`, `removeParagraph`, `addParagraphAt`

**Formatting**: `setBorders`, `setShading`, `setBackgroundColor` / `getBackgroundColor`, `setVerticalAlignment`, `setWidth`

**Spanning**: `setHorizontalMerge`, `setVerticalMerge`

**Convenience**: `setTextAlignment`, `setAllParagraphsStyle`, `setAllRunsFont`, `setAllRunsSize`, `setAllRunsColor`

### Section

**Line numbering**: `setLineNumbering(options)`, `getLineNumbering()`, `clearLineNumbering()`

### Comment & CommentManager

**Comment**: `resolve()`, `unresolve()`, `isResolved()`

**CommentManager**: `getResolvedComments()`, `getUnresolvedComments()`

### Utilities

**Unit conversion**

```typescript
import { twipsToPoints, inchesToTwips, emusToPixels } from 'docxmlater';

twipsToPoints(240); // 12 points
inchesToTwips(1); // 1440 twips
emusToPixels(914400, 96); // 96 pixels at 96 DPI
```

40+ conversion helpers across twips, EMUs, points, pixels, inches, and centimeters.

**Validation**

```typescript
import { validateRunText, cleanXmlFromText } from 'docxmlater';

const result = validateRunText('Some <w:t>text</w:t>');
if (result.hasXml) {
  const cleaned = cleanXmlFromText(result.text);
}
```

**Corruption detection**

```typescript
import { detectCorruptionInDocument } from 'docxmlater';

const doc = await Document.load('suspect.docx');
const report = detectCorruptionInDocument(doc);

if (report.isCorrupted) {
  report.locations.forEach((loc) => {
    console.log(`Line ${loc.lineNumber}: ${loc.issue}`);
  });
}
```

---

## Advanced Topics

### Tracked Changes

By default, `Document.load()` accepts all tracked changes during loading. This prevents revision-ID conflicts that can cause Word to report "unreadable content" on round-trip.

```typescript
const doc = await Document.load('document.docx', {
  revisionHandling: 'accept', // default - keep insertions, drop deletions
  // revisionHandling: 'strip',    - remove all revision markup entirely
  // revisionHandling: 'preserve', - keep tracked changes verbatim (advanced)
});
```

| Mode               | Behavior                                                                 |
| ------------------ | ------------------------------------------------------------------------ |
| `accept` (default) | Removes revision markup, keeps inserted content, removes deleted content |
| `strip`            | Removes all revision markup completely                                   |
| `preserve`         | Keeps tracked changes intact for advanced workflows                      |

### Custom Styles

```typescript
import { Document, Style } from 'docxmlater';

const doc = Document.create();

const heading = new Style('CustomHeading', 'paragraph');
heading.setName('Custom Heading');
heading.setRunFormatting({ bold: true, fontSize: 32, color: '0070C0' });
heading.setParagraphFormatting({ alignment: 'center', spacingAfter: 240 });

doc.getStylesManager().addStyle(heading);

const para = doc.createParagraph();
para.addText('Styled Heading');
para.applyStyle('CustomHeading');

await doc.save('styled.docx');
doc.dispose();
```

### Hyperlink Management

```typescript
const doc = await Document.load('document.docx');

const links = doc.getHyperlinks();
console.log(`Found ${links.length} hyperlinks`);

doc.updateHyperlinkUrls('http://old-domain.com', 'https://new-domain.com');

const merged = doc.defragmentHyperlinks({ resetFormatting: true });
console.log(`Merged ${merged} fragmented hyperlinks`);

await doc.save('updated.docx');
doc.dispose();
```

`defragmentHyperlinks` repairs fragmented links commonly produced by Google Docs exports. Batch URL updates run 30-50% faster than manual iteration.

### Compatibility Mode

```typescript
const doc = await Document.load('legacy.docx');

console.log(`Mode: ${doc.getCompatibilityMode()}`); // e.g. 12 (Word 2007)

if (doc.isCompatibilityMode()) {
  const info = doc.getCompatibilityInfo();
  console.log(`Legacy flags: ${info.legacyFlags.length}`);

  const report = doc.upgradeToModernFormat();
  console.log(`Removed ${report.removedFlags.length} legacy flags`);
  console.log(`Added ${report.addedSettings.length} modern settings`);
}

await doc.save('modern.docx');
doc.dispose();
```

`upgradeToModernFormat()` is the programmatic equivalent of _File → Info → Convert_ in Word.

### Templates

```typescript
const doc = await Document.load('template.docx');

doc.fillTemplate({
  customer: 'Acme Corp',
  date: '2025-04-25',
  total: '$12,400.00',
});

await doc.save('invoice-acme.docx');
doc.dispose();
```

Placeholders use `{{key}}` syntax and are replaced safely across run boundaries.

### Document Conversion

```typescript
const doc = await Document.load('report.docx');

const md = doc.toMarkdown();
const html = doc.toHTML({ fullPage: true });
const base64 = doc.toBase64();

doc.dispose();
```

---

## Performance & Memory Management

- **Always call `dispose()`** to release ZIP handles and image buffers
- Buffer-based I/O (`loadFromBuffer` / `toBuffer`) is 20-30% faster than file-path I/O
- Default size limits: warn at 50 MB, error at 150 MB (configurable via `LoadOptions.sizeLimits`)
- Memory footprint: ~2 MB per `Document`, ~2 bytes/character, full buffer per embedded image, ~200 bytes/cell
- For repeated paragraph access, cache `getAllParagraphs()` rather than calling it inside a loop
- Large documents (1,000+ pages) are supported

### Recommended Pattern

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

For server-side buffer workflows:

```typescript
async function processDocument(input: Buffer): Promise<Buffer> {
  const doc = await Document.loadFromBuffer(input);
  try {
    doc.replaceText(/placeholder/g, 'actual value');
    return await doc.toBuffer();
  } finally {
    doc.dispose();
  }
}
```

Custom error types are available from `docxmlater/internal`. These include `DocxError`, `InvalidDocxError`, `CorruptedArchiveError`, and `FileOperationError`.

Logging is configurable via `DOCXMLATER_LOG_LEVEL=debug|info|warn|error`.

---

## Architecture

```
src/
├── core/          Document, Parser, Generator, Validator
├── elements/      Paragraph, Run, Table, Image, Section, ...
├── formatting/    Style and Numbering managers
├── managers/      Drawing, Image, Relationship managers
├── tracking/      Revision tracking context
├── validation/    Revision and structural validation
├── helpers/       Cleanup utilities
├── xml/           XML generation and parsing (ReDoS-safe)
├── zip/           ZIP archive handling
├── constants/     Compatibility flags, limits, schema constants
├── types/         TypeScript type definitions
└── utils/         Units, validation, error handling
```

**Design principles**

- Strict adherence to ECMA-376 (Office Open XML)
- Position-based XML parsing (not regex) to prevent ReDoS
- Round-trip XML fidelity through `_originalXml` preservation and dirty-flag regeneration
- Explicit memory management via the `dispose()` pattern
- Defensive validation with comprehensive type coverage

---

## Security

- **ReDoS protection** - position-based XML parsing eliminates catastrophic backtracking
- **Path traversal prevention** - DOCX archive entries are validated against `../`, absolute paths, and URL-encoded traversal
- **XML injection prevention** - all text and attribute content is escaped via `XMLBuilder.escapeXmlText()` and `XMLBuilder.escapeXmlAttribute()`
- **Size limits** - configurable warning (50 MB) and hard cap (150 MB) on document size
- **Nesting limits** - XML parser caps nesting depth at 256 levels (configurable) to prevent stack overflow
- **UTF-8 enforcement** - all text content is explicitly UTF-8 encoded per ECMA-376

```typescript
const doc = await Document.load('large.docx', {
  sizeLimits: { warningSizeMB: 100, maxSizeMB: 500 },
});
```

```typescript
import { XMLParser } from 'docxmlater/internal';

const obj = XMLParser.parseToObject(xml, { maxNestingDepth: 512 });
```

---

## TypeScript Support

Full type definitions are bundled with the package:

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

const formatting: RunFormatting = {
  bold: true,
  fontSize: 12,
  color: 'FF0000',
};

const properties: DocumentProperties = {
  title: 'My Document',
  author: 'Jane Doe',
  created: new Date(),
};
```

---

## Requirements

- Node.js 18.0.0 or higher
- TypeScript 5.0+ (for development)

Single runtime dependency: `jszip`.

---

## Contributing

Contributions are welcome. Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for any new functionality
4. Ensure the full test suite passes (`npm test`)
5. Open a pull request

If you have a use case that is not yet supported, opening an issue first is the best way to discuss design before code.

---

## License

MIT
