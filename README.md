# docXMLater - Professional DOCX Framework

[![npm version](https://img.shields.io/npm/v/docxmlater.svg)](https://www.npmjs.com/package/docxmlater)
[![Tests](https://img.shields.io/badge/tests-474%20passing-brightgreen)](https://github.com/ItMeDiaTech/docXMLater)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.3-blue)](https://www.typescriptlang.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A comprehensive, production-ready TypeScript/JavaScript library for creating, reading, and manipulating Microsoft Word (.docx) documents programmatically. Full OpenXML compliance with extensive API coverage.

I do a lot of professional documentation work. From the solutions that exist out there for working with .docx files and therefore .xml files, they are not amazing. Most of the frameworks that exist that do give you everything you want... charge thousands a year. I decided to make my own framework to interact with these filetypes and focus on ease of usability. I think most if not all functionality works right now with helper functions to interact wiht all aspects of a .docx / xml document.

## Quick Start

```bash
npm install docxmlater
```

```typescript
import { Document } from "docxmlater";

// Create document
const doc = Document.create();
doc.createParagraph("Hello World").setStyle("Title");

// Save document
await doc.save("output.docx");
```

## Complete API Reference

### Document Operations

| Method                            | Description             | Example                                          |
| --------------------------------- | ----------------------- | ------------------------------------------------ |
| `Document.create(options?)`       | Create new document     | `const doc = Document.create()`                  |
| `Document.createEmpty()`          | Create minimal document | `const doc = Document.createEmpty()`             |
| `Document.load(path)`             | Load from file          | `const doc = await Document.load('file.docx')`   |
| `Document.loadFromBuffer(buffer)` | Load from buffer        | `const doc = await Document.loadFromBuffer(buf)` |
| `save(path)`                      | Save to file            | `await doc.save('output.docx')`                  |
| `toBuffer()`                      | Export as buffer        | `const buffer = await doc.toBuffer()`            |
| `dispose()`                       | Clean up resources      | `doc.dispose()`                                  |

### Content Creation

| Method                           | Description            | Example                          |
| -------------------------------- | ---------------------- | -------------------------------- |
| `createParagraph(text?)`         | Add paragraph          | `doc.createParagraph('Text')`    |
| `createTable(rows, cols)`        | Add table              | `doc.createTable(3, 4)`          |
| `addParagraph(para)`             | Add existing paragraph | `doc.addParagraph(myPara)`       |
| `addTable(table)`                | Add existing table     | `doc.addTable(myTable)`          |
| `addImage(image)`                | Add image              | `doc.addImage(myImage)`          |
| `addTableOfContents(toc?)`       | Add TOC                | `doc.addTableOfContents()`       |
| `insertParagraphAt(index, para)` | Insert at position     | `doc.insertParagraphAt(0, para)` |

### Content Retrieval

| Method                | Description           | Returns                                    |
| --------------------- | --------------------- | ------------------------------------------ |
| `getParagraphs()`     | Get all paragraphs    | `Paragraph[]`                              |
| `getTables()`         | Get all tables        | `Table[]`                                  |
| `getBodyElements()`   | Get all body elements | `BodyElement[]`                            |
| `getParagraphCount()` | Count paragraphs      | `number`                                   |
| `getTableCount()`     | Count tables          | `number`                                   |
| `getHyperlinks()`     | Get all links         | `Array<{hyperlink, paragraph}>`            |
| `getBookmarks()`      | Get all bookmarks     | `Array<{bookmark, paragraph}>`             |
| `getImages()`         | Get all images        | `Array<{image, relationshipId, filename}>` |

### Content Removal

| Method                         | Description        | Returns   |
| ------------------------------ | ------------------ | --------- |
| `removeParagraph(paraOrIndex)` | Remove paragraph   | `boolean` |
| `removeTable(tableOrIndex)`    | Remove table       | `boolean` |
| `clearParagraphs()`            | Remove all content | `this`    |

### Search & Replace

| Method                                 | Description           | Options                        |
| -------------------------------------- | --------------------- | ------------------------------ |
| `findText(text, options?)`             | Find text occurrences | `{caseSensitive?, wholeWord?}` |
| `replaceText(find, replace, options?)` | Replace all text      | `{caseSensitive?, wholeWord?}` |
| `updateHyperlinkUrls(urlMap)`          | Update hyperlink URLs | `Map<oldUrl, newUrl>`          |

### Document Statistics

| Method                              | Description         | Returns                        |
| ----------------------------------- | ------------------- | ------------------------------ |
| `getWordCount()`                    | Total word count    | `number`                       |
| `getCharacterCount(includeSpaces?)` | Character count     | `number`                       |
| `estimateSize()`                    | Size estimation     | `{totalEstimatedMB, warning?}` |
| `getSizeStats()`                    | Detailed size stats | `{elements, size, warnings}`   |

### Text Formatting

| Property      | Values                           | Example                 |
| ------------- | -------------------------------- | ----------------------- |
| `bold`        | `true/false`                     | `{bold: true}`          |
| `italic`      | `true/false`                     | `{italic: true}`        |
| `underline`   | `'single'/'double'/'dotted'/etc` | `{underline: 'single'}` |
| `strike`      | `true/false`                     | `{strike: true}`        |
| `font`        | Font name                        | `{font: 'Arial'}`       |
| `size`        | Points                           | `{size: 12}`            |
| `color`       | Hex color                        | `{color: 'FF0000'}`     |
| `highlight`   | Color name                       | `{highlight: 'yellow'}` |
| `subscript`   | `true/false`                     | `{subscript: true}`     |
| `superscript` | `true/false`                     | `{superscript: true}`   |
| `smallCaps`   | `true/false`                     | `{smallCaps: true}`     |
| `allCaps`     | `true/false`                     | `{allCaps: true}`       |

### Paragraph Formatting

| Method                         | Description         | Values                              |
| ------------------------------ | ------------------- | ----------------------------------- |
| `setAlignment(align)`          | Text alignment      | `'left'/'center'/'right'/'justify'` |
| `setLeftIndent(twips)`         | Left indentation    | Twips value                         |
| `setRightIndent(twips)`        | Right indentation   | Twips value                         |
| `setFirstLineIndent(twips)`    | First line indent   | Twips value                         |
| `setSpaceBefore(twips)`        | Space before        | Twips value                         |
| `setSpaceAfter(twips)`         | Space after         | Twips value                         |
| `setLineSpacing(twips, rule?)` | Line spacing        | Twips + rule                        |
| `setStyle(styleId)`            | Apply style         | Style ID                            |
| `setKeepNext()`                | Keep with next      | -                                   |
| `setKeepLines()`               | Keep lines together | -                                   |
| `setPageBreakBefore()`         | Page break before   | -                                   |

### Table Operations

| Method                  | Description          | Example                                  |
| ----------------------- | -------------------- | ---------------------------------------- |
| `getRow(index)`         | Get table row        | `table.getRow(0)`                        |
| `getCell(row, col)`     | Get table cell       | `table.getCell(0, 1)`                    |
| `addRow()`              | Add new row          | `table.addRow()`                         |
| `removeRow(index)`      | Remove row           | `table.removeRow(2)`                     |
| `insertColumn(index)`   | Insert column        | `table.insertColumn(1)`                  |
| `removeColumn(index)`   | Remove column        | `table.removeColumn(3)`                  |
| `setWidth(twips)`       | Set table width      | `table.setWidth(8640)`                   |
| `setAlignment(align)`   | Table alignment      | `table.setAlignment('center')`           |
| `setAllBorders(border)` | Set all borders      | `table.setAllBorders({style: 'single'})` |
| `setBorders(borders)`   | Set specific borders | `table.setBorders({top: {...}})`         |

### Table Cell Operations

| Method                        | Description           | Example                               |
| ----------------------------- | --------------------- | ------------------------------------- |
| `createParagraph(text?)`      | Add paragraph to cell | `cell.createParagraph('Text')`        |
| `setShading(shading)`         | Cell background       | `cell.setShading({fill: 'E0E0E0'})`   |
| `setVerticalAlignment(align)` | Vertical align        | `cell.setVerticalAlignment('center')` |
| `setColumnSpan(cols)`         | Merge columns         | `cell.setColumnSpan(3)`               |
| `setRowSpan(rows)`            | Merge rows            | `cell.setRowSpan(2)`                  |
| `setBorders(borders)`         | Cell borders          | `cell.setBorders({top: {...}})`       |
| `setWidth(width, type?)`      | Cell width            | `cell.setWidth(2000, 'dxa')`          |

### Style Management

| Method                        | Description        | Example                            |
| ----------------------------- | ------------------ | ---------------------------------- |
| `addStyle(style)`             | Add custom style   | `doc.addStyle(myStyle)`            |
| `getStyle(styleId)`           | Get style by ID    | `doc.getStyle('Heading1')`         |
| `hasStyle(styleId)`           | Check style exists | `doc.hasStyle('CustomStyle')`      |
| `getStyles()`                 | Get all styles     | `doc.getStyles()`                  |
| `removeStyle(styleId)`        | Remove style       | `doc.removeStyle('OldStyle')`      |
| `updateStyle(styleId, props)` | Update style       | `doc.updateStyle('Normal', {...})` |

#### Built-in Styles

- `Normal` - Default paragraph
- `Title` - Document title
- `Subtitle` - Document subtitle
- `Heading1` through `Heading9` - Section headings
- `ListParagraph` - List items

### List Management

| Method                                  | Description             | Returns |
| --------------------------------------- | ----------------------- | ------- |
| `createBulletList(levels?, bullets?)`   | Create bullet list      | `numId` |
| `createNumberedList(levels?, formats?)` | Create numbered list    | `numId` |
| `createMultiLevelList()`                | Create multi-level list | `numId` |

### Image Handling

| Method                                  | Description        | Example                          |
| --------------------------------------- | ------------------ | -------------------------------- |
| `Image.fromFile(path, width?, height?)` | Load from file     | `Image.fromFile('pic.jpg')`      |
| `Image.fromBuffer(buffer, ext, w?, h?)` | Load from buffer   | `Image.fromBuffer(buf, 'png')`   |
| `setWidth(emus, maintainRatio?)`        | Set width          | `img.setWidth(inchesToEmus(3))`  |
| `setHeight(emus, maintainRatio?)`       | Set height         | `img.setHeight(inchesToEmus(2))` |
| `setSize(width, height)`                | Set dimensions     | `img.setSize(w, h)`              |
| `setRotation(degrees)`                  | Rotate image       | `img.setRotation(90)`            |
| `setAltText(text)`                      | Accessibility text | `img.setAltText('Description')`  |

### Hyperlinks

| Method                                            | Description      | Example                                                    |
| ------------------------------------------------- | ---------------- | ---------------------------------------------------------- |
| `Hyperlink.createExternal(url, text, format?)`    | Web link         | `Hyperlink.createExternal('https://example.com', 'Click')` |
| `Hyperlink.createEmail(email, text?, format?)`    | Email link       | `Hyperlink.createEmail('user@example.com')`                |
| `Hyperlink.createInternal(anchor, text, format?)` | Internal link    | `Hyperlink.createInternal('Section1', 'Go to')`            |
| `para.addHyperlink(hyperlink)`                    | Add to paragraph | `para.addHyperlink(link)`                                  |

### Headers & Footers

| Method                       | Description        | Example                          |
| ---------------------------- | ------------------ | -------------------------------- |
| `setHeader(header)`          | Set default header | `doc.setHeader(myHeader)`        |
| `setFooter(footer)`          | Set default footer | `doc.setFooter(myFooter)`        |
| `setFirstPageHeader(header)` | First page header  | `doc.setFirstPageHeader(header)` |
| `setFirstPageFooter(footer)` | First page footer  | `doc.setFirstPageFooter(footer)` |
| `setEvenPageHeader(header)`  | Even page header   | `doc.setEvenPageHeader(header)`  |
| `setEvenPageFooter(footer)`  | Even page footer   | `doc.setEvenPageFooter(footer)`  |

### Page Setup

| Method                                | Description       | Example                               |
| ------------------------------------- | ----------------- | ------------------------------------- |
| `setPageSize(width, height, orient?)` | Page dimensions   | `doc.setPageSize(12240, 15840)`       |
| `setPageOrientation(orientation)`     | Page orientation  | `doc.setPageOrientation('landscape')` |
| `setMargins(margins)`                 | Page margins      | `doc.setMargins({top: 1440, ...})`    |
| `setLanguage(language)`               | Document language | `doc.setLanguage('en-US')`            |

### Document Properties

| Method                 | Description  | Properties                            |
| ---------------------- | ------------ | ------------------------------------- |
| `setProperties(props)` | Set metadata | `{title, subject, creator, keywords}` |
| `getProperties()`      | Get metadata | Returns all properties                |

### Advanced Features

#### Bookmarks

| Method                                   | Description         |
| ---------------------------------------- | ------------------- |
| `createBookmark(name)`                   | Create bookmark     |
| `createHeadingBookmark(text)`            | Auto-named bookmark |
| `getBookmark(name)`                      | Get by name         |
| `hasBookmark(name)`                      | Check existence     |
| `addBookmarkToParagraph(para, bookmark)` | Add to paragraph    |

#### Comments

| Method                                      | Description       |
| ------------------------------------------- | ----------------- |
| `createComment(author, content, initials?)` | Add comment       |
| `createReply(parentId, author, content)`    | Reply to comment  |
| `getComment(id)`                            | Get by ID         |
| `getAllComments()`                          | Get all top-level |
| `addCommentToParagraph(para, comment)`      | Add to paragraph  |

#### Track Changes

| Method                               | Description       |
| ------------------------------------ | ----------------- |
| `trackInsertion(para, author, text)` | Track insertion   |
| `trackDeletion(para, author, text)`  | Track deletion    |
| `isTrackingChanges()`                | Check if tracking |
| `getRevisionStats()`                 | Get statistics    |

#### Footnotes & Endnotes

| Method                     | Description      |
| -------------------------- | ---------------- |
| `FootnoteManager.create()` | Manage footnotes |
| `EndnoteManager.create()`  | Manage endnotes  |

### Low-Level Document Parts

| Method                       | Description           | Example                                           |
| ---------------------------- | --------------------- | ------------------------------------------------- |
| `getPart(partName)`          | Get document part     | `doc.getPart('word/document.xml')`                |
| `setPart(partName, content)` | Set document part     | `doc.setPart('custom.xml', data)`                 |
| `removePart(partName)`       | Remove part           | `doc.removePart('custom.xml')`                    |
| `listParts()`                | List all parts        | `const parts = await doc.listParts()`             |
| `partExists(partName)`       | Check part exists     | `if (await doc.partExists('...'))`                |
| `getContentTypes()`          | Get content types     | `const types = await doc.getContentTypes()`       |
| `addContentType(part, type)` | Register content type | `doc.addContentType('.json', 'application/json')` |

### Unit Conversion Utilities

| Function                     | Description          | Example                     |
| ---------------------------- | -------------------- | --------------------------- |
| `inchesToTwips(inches)`      | Inches to twips      | `inchesToTwips(1)` // 1440  |
| `inchesToEmus(inches)`       | Inches to EMUs       | `inchesToEmus(1)` // 914400 |
| `cmToTwips(cm)`              | Centimeters to twips | `cmToTwips(2.54)` // 1440   |
| `pointsToTwips(points)`      | Points to twips      | `pointsToTwips(12)` // 240  |
| `pixelsToEmus(pixels, dpi?)` | Pixels to EMUs       | `pixelsToEmus(96)`          |

## Common Recipes

### Create a Simple Document

```typescript
const doc = Document.create();
doc.createParagraph("Title").setStyle("Title");
doc.createParagraph("This is a simple document.");
await doc.save("simple.docx");
```

### Add Formatted Text

```typescript
const para = doc.createParagraph();
para.addText("Bold", { bold: true });
para.addText(" and ");
para.addText("Colored", { color: "FF0000" });
```

### Create a Table with Borders

```typescript
const table = doc.createTable(3, 3);
table.setAllBorders({ style: "single", size: 8, color: "000000" });
table.getCell(0, 0)?.createParagraph("Header 1");
table.getRow(0)?.getCell(0)?.setShading({ fill: "4472C4" });
```

### Insert an Image

```typescript
import { Image, inchesToEmus } from "docxmlater";

const image = Image.fromFile("./photo.jpg");
image.setWidth(inchesToEmus(4), true); // 4 inches, maintain ratio
doc.addImage(image);
```

### Add a Hyperlink

```typescript
const para = doc.createParagraph();
para.addText("Visit ");
para.addHyperlink(
  Hyperlink.createExternal("https://example.com", "our website")
);
```

### Search and Replace Text

```typescript
// Find all occurrences
const results = doc.findText("old text", { caseSensitive: true });
console.log(`Found ${results.length} occurrences`);

// Replace all
const count = doc.replaceText("old text", "new text", { wholeWord: true });
console.log(`Replaced ${count} occurrences`);
```

### Load and Modify Existing Document

```typescript
const doc = await Document.load("existing.docx");
doc.createParagraph("Added paragraph");

// Update all hyperlinks
const urlMap = new Map([["https://old-site.com", "https://new-site.com"]]);
doc.updateHyperlinkUrls(urlMap);

await doc.save("modified.docx");
```

### Create Lists

```typescript
// Bullet list
const bulletId = doc.createBulletList(3);
doc.createParagraph("First item").setNumbering(bulletId, 0);
doc.createParagraph("Second item").setNumbering(bulletId, 0);

// Numbered list
const numberId = doc.createNumberedList(3);
doc.createParagraph("Step 1").setNumbering(numberId, 0);
doc.createParagraph("Step 2").setNumbering(numberId, 0);
```

### Apply Custom Styles

```typescript
import { Style } from "docxmlater";

const customStyle = Style.create({
  styleId: "CustomHeading",
  name: "Custom Heading",
  basedOn: "Normal",
  runFormatting: { bold: true, size: 14, color: "2E74B5" },
  paragraphFormatting: { alignment: "center", spaceAfter: 240 },
});

doc.addStyle(customStyle);
doc.createParagraph("Custom Styled Text").setStyle("CustomHeading");
```

### Add Headers and Footers

```typescript
import { Header, Footer, Field } from "docxmlater";

// Header with page numbers
const header = Header.create();
header.addParagraph("Document Title").setAlignment("center");

// Footer with page numbers
const footer = Footer.create();
const footerPara = footer.addParagraph();
footerPara.addText("Page ");
footerPara.addField(Field.create({ type: "PAGE" }));
footerPara.addText(" of ");
footerPara.addField(Field.create({ type: "NUMPAGES" }));

doc.setHeader(header);
doc.setFooter(footer);
```

### Work with Document Statistics

```typescript
// Get word and character counts
console.log("Words:", doc.getWordCount());
console.log("Characters:", doc.getCharacterCount());
console.log("Characters (no spaces):", doc.getCharacterCount(false));

// Check document size
const size = doc.estimateSize();
if (size.warning) {
  console.warn(size.warning);
}
console.log(`Estimated size: ${size.totalEstimatedMB} MB`);
```

### Handle Large Documents Efficiently

```typescript
const doc = Document.create({
  maxMemoryUsagePercent: 80,
  maxRssMB: 2048,
  maxImageCount: 50,
  maxTotalImageSizeMB: 100,
});

// Process document...

// Clean up resources after saving
await doc.save("large-document.docx");
doc.dispose(); // Free memory
```

### Direct XML Access (Advanced)

```typescript
// Get raw XML
const documentXml = await doc.getPart("word/document.xml");
console.log(documentXml?.content);

// Modify raw XML (use with caution)
await doc.setPart("word/custom.xml", "<custom>data</custom>");
await doc.addContentType("/word/custom.xml", "application/xml");

// List all parts
const parts = await doc.listParts();
console.log("Document contains:", parts.length, "parts");
```

## Features

- **Full OpenXML Compliance** - Follows ECMA-376 standard
- **TypeScript First** - Complete type definitions
- **Memory Efficient** - Handles large documents with streaming
- **Atomic Saves** - Prevents corruption with temp file pattern
- **Rich Formatting** - Complete text and paragraph formatting
- **Tables** - Full support with borders, shading, merging
- **Images** - PNG, JPEG, GIF with sizing and positioning
- **Hyperlinks** - External, internal, and email links
- **Styles** - 13 built-in styles + custom style creation
- **Lists** - Bullets, numbering, multi-level
- **Headers/Footers** - Different first/even/odd pages
- **Search & Replace** - With case and whole word options
- **Document Stats** - Word count, character count, size estimation
- **Track Changes** - Insertions and deletions with authors
- **Comments** - With replies and threading
- **Bookmarks** - For internal navigation
- **Low-level Access** - Direct ZIP and XML manipulation

## Performance

- Process 100+ page documents efficiently
- Atomic save pattern prevents corruption
- Memory management for large files
- Lazy loading of document parts
- Resource cleanup with `dispose()`

## Testing

```bash
npm test                 # Run all tests
npm run test:watch      # Watch mode
npm run test:coverage   # Coverage report
```

**Current:** 474 tests passing | 98.1% pass rate | 100% core functionality covered

## Development

```bash
# Install dependencies
npm install

# Build TypeScript
npm run build

# Run examples
npx ts-node examples/simple-document.ts
```

## Project Structure

```text
src/
├── core/           # Document, Parser, Generator, Validator
├── elements/       # Paragraph, Run, Table, Image, Hyperlink
├── formatting/     # Style, NumberingManager
├── xml/           # XMLBuilder, XMLParser
├── zip/           # ZipHandler for DOCX manipulation
└── utils/         # Validation, Units conversion

examples/
├── 01-basic/      # Simple document creation
├── 02-text/       # Text formatting examples
├── 03-tables/     # Table examples
├── 04-styles/     # Style examples
├── 05-images/     # Image handling
├── 06-complete/   # Full document examples
└── 07-hyperlinks/ # Link examples
```

## Hierarchy

```text
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

## Requirements

- Node.js 16+
- TypeScript 5.0+ (for development)

## Installation Options

```bash
# NPM
npm install docxmlater

# Yarn
yarn add docxmlater

# PNPM
pnpm add docxmlater
```

## Troubleshooting

### XML Corruption in Text

**Problem**: Text displays with XML tags like `Important Information<w:t xml:space="preserve">1` in Word.

**Cause**: Passing XML-like strings to text methods instead of using the API properly.

```typescript
// WRONG - Will display escaped XML as literal text
paragraph.addText("Important Information<w:t>1</w:t>");
// Displays as: "Important Information<w:t>1</w:t>"

// CORRECT - Use separate text runs
paragraph.addText("Important Information");
paragraph.addText("1");
// Displays as: "Important Information1"

// Or combine in one call
paragraph.addText("Important Information 1");
```

**Detection**: Use the corruption detection utility to find issues:

```typescript
import { detectCorruptionInDocument } from "docxmlater";

const doc = await Document.load("file.docx");
const report = detectCorruptionInDocument(doc);

if (report.isCorrupted) {
  console.log(report.summary);
  report.locations.forEach((loc) => {
    console.log(`Paragraph ${loc.paragraphIndex}, Run ${loc.runIndex}:`);
    console.log(`  Original: ${loc.text}`);
    console.log(`  Fixed:    ${loc.suggestedFix}`);
  });
}
```

**Auto-Cleaning**: XML patterns are automatically removed by default for defensive data handling:

```typescript
// Default behavior - auto-clean enabled
const run = new Run("Text<w:t>value</w:t>");
// Result: "Textvalue" (XML tags removed automatically)

// Disable auto-cleaning (for debugging)
const run = new Run("Text<w:t>value</w:t>", { cleanXmlFromText: false });
// Result: "Text<w:t>value</w:t>" (XML tags preserved, will display in Word)
```

**Why This Happens**: The framework correctly escapes XML special characters per the XML specification. When you pass XML tags as text, they are properly escaped (`<` becomes `&lt;`) and Word displays them as literal text, not as markup.

**The Right Approach**: Use the framework's API methods instead of embedding XML:

- ✅ Use `paragraph.addText()` multiple times for separate text runs
- ✅ Use formatting options: `{bold: true}`, `{italic: true}`, etc.
- ✅ Use `paragraph.addHyperlink()` for links
- ❌ Don't pass XML strings to text methods
- ❌ Don't try to embed `<w:t>` or other XML tags in your text

For more details, see the [corruption detection examples](examples/troubleshooting/).

## Contributing

Contributions are welcome! Please read our [Contributing Guide](CONTRIBUTING.md).

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Recent Updates (v0.20.1)

**Critical Bug Fix Release:**

- ✅ **Fixed Paragraph.getText()** - Now includes hyperlink text content (critical data loss bug)
- ✅ **Added hyperlink integration tests** - 6 new comprehensive test cases
- ✅ **Enhanced test suite** - 474/478 tests passing (98.1% pass rate)
- ✅ **Fixed type safety** - XMLElement handling improvements across test files
- ✅ **Improved StylesManager** - XML corruption detection moved before parser
- ✅ **Hyperlink management** - Proper relationship ID clearing on URL updates

**What This Fixes:**
When using `para.addText('foo') + para.addHyperlink(link)`, the hyperlink text is now properly included in `paragraph.getText()`, preventing silent text loss.

## License

MIT © DiaTech

## Acknowledgments

- Built with [JSZip](https://stuk.github.io/jszip/) for ZIP handling
- Follows [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) Office Open XML standard
- Inspired by [python-docx](https://python-docx.readthedocs.io/) and [docx](https://github.com/dolanmiu/docx)

## Support

- **Documentation**: [Full Docs](https://github.com/ItMeDiaTech/docXMLater/tree/main/docs)
- **Examples**: [Example Code](https://github.com/ItMeDiaTech/docXMLater/tree/main/examples)
- **Issues**: [GitHub Issues](https://github.com/ItMeDiaTech/docXMLater/issues)
- **Discussions**: [GitHub Discussions](https://github.com/ItMeDiaTech/docXMLater/discussions)

## Quick Links

- [NPM Package](https://www.npmjs.com/package/docxmlater)
- [GitHub Repository](https://github.com/ItMeDiaTech/docXMLater)
- [API Reference](https://github.com/ItMeDiaTech/docXMLater/tree/main/docs/api)
- [Change Log](https://github.com/ItMeDiaTech/docXMLater/blob/main/CHANGELOG.md)

---

**Ready to create amazing Word documents?** Start with our [examples](https://github.com/ItMeDiaTech/docXMLater/tree/main/examples) or dive into the [API Reference](#complete-api-reference) above!
