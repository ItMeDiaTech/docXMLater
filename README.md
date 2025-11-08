# docXMLater - Professional DOCX Framework

[![npm version](https://img.shields.io/npm/v/docxmlater.svg)](https://www.npmjs.com/package/docxmlater)
[![Tests](https://img.shields.io/badge/tests-1119%20passing-brightgreen)](https://github.com/ItMeDiaTech/docXMLater)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.7-blue)](https://www.typescriptlang.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A comprehensive, production-ready TypeScript/JavaScript library for creating, reading, and manipulating Microsoft Word (.docx) documents programmatically. Full OpenXML compliance with extensive API coverage and robust test suite.

Built for professional documentation work, docXMLater provides a complete solution for programmatic DOCX manipulation with an intuitive API and helper functions for all aspects of document creation and modification.

## Latest Updates - v1.3.0

**Enhanced Parsing & Helper Functions:**

### What's New in v1.3.0

- **TOC Parsing:** Parse Table of Contents from existing documents with full SDT support
- **TOC Modification:** Modify TOC field instructions (add/remove switches: \h, \u, \z, \n, \o, \t)
- **Header/Footer Removal:** New `removeAllHeadersFooters()` helper function
- **Complete Feature Set:** All 102 major features implemented
- **Table Styles:** Full support with 12 conditional formatting types
- **Content Controls:** 9 control types (rich text, plain text, combo box, dropdown, date picker, checkbox, picture, building block, group)
- **Field Types:** 11 field types (PAGE, NUMPAGES, DATE, TIME, FILENAME, AUTHOR, TITLE, REF, HYPERLINK, SEQ, TC/XE)
- **Drawing Elements:** Shapes and textboxes with full positioning
- **Document Properties:** Core, extended, and custom properties
- **Production Ready:** Full ECMA-376 compliance

**Test Results:** 1,119/1,150 tests passing (97.3% pass rate - 1,119 core features validated)

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

| Method                           | Description              | Example                          |
| -------------------------------- | ------------------------ | -------------------------------- |
| `createParagraph(text?)`         | Add paragraph            | `doc.createParagraph('Text')`    |
| `createTable(rows, cols)`        | Add table                | `doc.createTable(3, 4)`          |
| `addParagraph(para)`             | Add existing paragraph   | `doc.addParagraph(myPara)`       |
| `addTable(table)`                | Add existing table       | `doc.addTable(myTable)`          |
| `addImage(image)`                | Add image                | `doc.addImage(myImage)`          |
| `addTableOfContents(toc?)`       | Add TOC                  | `doc.addTableOfContents()`       |
| `insertParagraphAt(index, para)` | Insert at position       | `doc.insertParagraphAt(0, para)` |
| `insertTableAt(index, table)`    | Insert table at position | `doc.insertTableAt(5, table)`    |
| `insertTocAt(index, toc)`        | Insert TOC at position   | `doc.insertTocAt(0, toc)`        |

### Content Manipulation

| Method                            | Description                  | Returns   |
| --------------------------------- | ---------------------------- | --------- |
| `replaceParagraphAt(index, para)` | Replace paragraph            | `boolean` |
| `replaceTableAt(index, table)`    | Replace table                | `boolean` |
| `moveElement(fromIndex, toIndex)` | Move element to new position | `boolean` |
| `swapElements(index1, index2)`    | Swap two elements            | `boolean` |
| `removeTocAt(index)`              | Remove TOC element           | `boolean` |

### Content Retrieval

| Method                 | Description                             | Returns                                    |
| ---------------------- | --------------------------------------- | ------------------------------------------ |
| `getParagraphs()`      | Get top-level paragraphs                | `Paragraph[]`                              |
| `getAllParagraphs()`   | Get all paragraphs (recursive)          | `Paragraph[]`                              |
| `getTables()`          | Get top-level tables                    | `Table[]`                                  |
| `getAllTables()`       | Get all tables (recursive)              | `Table[]`                                  |
| `getBodyElements()`    | Get all body elements                   | `BodyElement[]`                            |
| `getParagraphCount()`  | Count paragraphs                        | `number`                                   |
| `getTableCount()`      | Count tables                            | `number`                                   |
| `getHyperlinks()`      | Get all links                           | `Array<{hyperlink, paragraph}>`            |
| `getBookmarks()`       | Get all bookmarks                       | `Array<{bookmark, paragraph}>`             |
| `getImages()`          | Get all images                          | `Array<{image, relationshipId, filename}>` |

**Note**: The `getAllParagraphs()` and `getAllTables()` methods recursively search inside tables and SDTs (Structured Document Tags), while the non-prefixed methods only return top-level elements.

**Example - Recursive Element Access:**

```typescript
import { Document, Hyperlink } from 'docxmlater';

// Load document with complex structure (tables, SDTs, nested content)
const doc = await Document.load('complex.docx');

// Get only top-level paragraphs (misses nested content)
const topLevel = doc.getParagraphs();
console.log(`Top-level paragraphs: ${topLevel.length}`); // e.g., 37

// Get ALL paragraphs including those in tables and SDTs
const allParas = doc.getAllParagraphs();
console.log(`All paragraphs: ${allParas.length}`); // e.g., 52

// Apply formatting to ALL paragraphs (including nested ones)
for (const para of allParas) {
  para.setSpaceAfter(120); // Set 6pt spacing after each paragraph
}

// Get all tables including those inside SDTs
const allTables = doc.getAllTables();
for (const table of allTables) {
  table.setWidth(5000).setWidthType('pct'); // Set to 100% width
}

// Find all hyperlinks in the entire document
let hyperlinkCount = 0;
for (const para of allParas) {
  for (const content of para.getContent()) {
    if (content instanceof Hyperlink) {
      hyperlinkCount++;
      content.setFormatting({ color: '0000FF' }); // Make all links blue
    }
  }
}
console.log(`Updated ${hyperlinkCount} hyperlinks`);
```

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

### Style Application

| Method                                | Description                      | Returns             |
| ------------------------------------- | -------------------------------- | ------------------- |
| `applyStyleToAll(styleId, predicate)` | Apply style to matching elements | `number`            |
| `findElementsByStyle(styleId)`        | Find all elements using a style  | `Array<Para\|Cell>` |

**Example:**

```typescript
// Apply Heading1 to all paragraphs containing "Chapter"
const count = doc.applyStyleToAll("Heading1", (el) => {
  return el instanceof Paragraph && el.getText().includes("Chapter");
});

// Find all Heading1 elements
const headings = doc.findElementsByStyle("Heading1");
```

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

### Paragraph Operations

#### Creating Detached Paragraphs

Create paragraphs independently before adding to a document:

```typescript
// Create empty paragraph
const para1 = Paragraph.create();

// Create with text
const para2 = Paragraph.create("Hello World");

// Create with text and formatting
const para3 = Paragraph.create("Centered text", { alignment: "center" });

// Create with just formatting
const para4 = Paragraph.create({
  alignment: "right",
  spacing: { before: 240 },
});

// Create with style
const heading = Paragraph.createWithStyle("Chapter 1", "Heading1");

// Create with both run and paragraph formatting
const important = Paragraph.createFormatted(
  "Important Text",
  { bold: true, color: "FF0000" },
  { alignment: "center" }
);

// Add to document later
doc.addParagraph(para1);
doc.addParagraph(heading);
```

#### Paragraph Factory Methods

| Method                                              | Description                 | Example                                 |
| --------------------------------------------------- | --------------------------- | --------------------------------------- |
| `Paragraph.create(text?, formatting?)`              | Create detached paragraph   | `Paragraph.create('Text')`              |
| `Paragraph.create(formatting?)`                     | Create with formatting only | `Paragraph.create({alignment: 'left'})` |
| `Paragraph.createWithStyle(text, styleId)`          | Create with style           | `Paragraph.createWithStyle('', 'H1')`   |
| `Paragraph.createEmpty()`                           | Create empty paragraph      | `Paragraph.createEmpty()`               |
| `Paragraph.createFormatted(text, run?, paragraph?)` | Create with dual formatting | See example above                       |

#### Paragraph Formatting Methods

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

#### Paragraph Manipulation Methods

| Method                                 | Description             | Returns     |
| -------------------------------------- | ----------------------- | ----------- |
| `insertRunAt(index, run)`              | Insert run at position  | `this`      |
| `removeRunAt(index)`                   | Remove run at position  | `boolean`   |
| `replaceRunAt(index, run)`             | Replace run at position | `boolean`   |
| `findText(text, options?)`             | Find text in runs       | `number[]`  |
| `replaceText(find, replace, options?)` | Replace text in runs    | `number`    |
| `mergeWith(otherPara)`                 | Merge another paragraph | `this`      |
| `clone()`                              | Clone paragraph         | `Paragraph` |

**Example:**

```typescript
const para = doc.createParagraph("Hello World");

// Find and replace
const indices = para.findText("World"); // [1]
const count = para.replaceText("World", "Universe", { caseSensitive: true });

// Manipulate runs
para.insertRunAt(0, new Run("Start: ", { bold: true }));
para.replaceRunAt(1, new Run("HELLO", { allCaps: true }));

// Merge paragraphs
const para2 = Paragraph.create(" More text");
para.mergeWith(para2); // Combines runs
```

### Run (Text Span) Operations

| Method                          | Description               | Returns |
| ------------------------------- | ------------------------- | ------- |
| `clone()`                       | Clone run with formatting | `Run`   |
| `insertText(index, text)`       | Insert text at position   | `this`  |
| `appendText(text)`              | Append text to end        | `this`  |
| `replaceText(start, end, text)` | Replace text range        | `this`  |

**Example:**

```typescript
const run = new Run("Hello World", { bold: true });

// Text manipulation
run.insertText(6, "Beautiful "); // "Hello Beautiful World"
run.appendText("!"); // "Hello Beautiful World!"
run.replaceText(0, 5, "Hi"); // "Hi Beautiful World!"

// Clone for reuse
const copy = run.clone();
copy.setColor("FF0000"); // Original unchanged
```

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

#### Advanced Table Operations

| Method                                           | Description               | Returns      |
| ------------------------------------------------ | ------------------------- | ------------ |
| `mergeCells(startRow, startCol, endRow, endCol)` | Merge cells               | `this`       |
| `splitCell(row, col)`                            | Remove cell spanning      | `this`       |
| `moveCell(fromRow, fromCol, toRow, toCol)`       | Move cell contents        | `this`       |
| `swapCells(row1, col1, row2, col2)`              | Swap two cells            | `this`       |
| `setColumnWidth(index, width)`                   | Set specific column width | `this`       |
| `setColumnWidths(widths)`                        | Set all column widths     | `this`       |
| `insertRows(startIndex, count)`                  | Insert multiple rows      | `TableRow[]` |
| `removeRows(startIndex, count)`                  | Remove multiple rows      | `boolean`    |
| `clone()`                                        | Clone entire table        | `Table`      |

**Example:**

```typescript
const table = doc.createTable(3, 3);

// Merge cells horizontally (row 0, columns 0-2)
table.mergeCells(0, 0, 0, 2);

// Move cell contents
table.moveCell(1, 1, 2, 2);

// Swap cells
table.swapCells(0, 0, 2, 2);

// Batch row operations
table.insertRows(1, 3); // Insert 3 rows at position 1
table.removeRows(4, 2); // Remove 2 rows starting at position 4

// Set column widths
table.setColumnWidth(0, 2000); // First column = 2000 twips
table.setColumnWidths([2000, 3000, 2000]); // All columns

// Clone table for reuse
const tableCopy = table.clone();
```

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

#### Style Manipulation

| Method                   | Description                         | Returns |
| ------------------------ | ----------------------------------- | ------- |
| `style.clone()`          | Clone style                         | `Style` |
| `style.mergeWith(other)` | Merge properties from another style | `this`  |

**Example:**

```typescript
// Clone a style
const heading1 = doc.getStyle("Heading1");
const customHeading = heading1.clone();
customHeading.setRunFormatting({ color: "FF0000" });

// Merge styles
const baseStyle = Style.createNormalStyle();
const overrideStyle = Style.create({
  styleId: "Override",
  name: "Override",
  type: "paragraph",
  runFormatting: { bold: true, color: "FF0000" },
});
baseStyle.mergeWith(overrideStyle); // baseStyle now has bold red text
```

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

### Table of Contents (TOC)

#### Basic TOC Creation

| Method                     | Description            | Example                    |
| -------------------------- | ---------------------- | -------------------------- |
| `addTableOfContents(toc?)` | Add TOC to document    | `doc.addTableOfContents()` |
| `insertTocAt(index, toc)`  | Insert TOC at position | `doc.insertTocAt(0, toc)`  |
| `removeTocAt(index)`       | Remove TOC at position | `doc.removeTocAt(0)`       |

#### TOC Factory Methods

| Method                                                   | Description              | Example                                              |
| -------------------------------------------------------- | ------------------------ | ---------------------------------------------------- |
| `TableOfContents.createStandard(title?)`                 | Standard TOC (3 levels)  | `TableOfContents.createStandard()`                   |
| `TableOfContents.createSimple(title?)`                   | Simple TOC (2 levels)    | `TableOfContents.createSimple()`                     |
| `TableOfContents.createDetailed(title?)`                 | Detailed TOC (4 levels)  | `TableOfContents.createDetailed()`                   |
| `TableOfContents.createHyperlinked(title?)`              | Hyperlinked TOC          | `TableOfContents.createHyperlinked()`                |
| `TableOfContents.createNoPageNumbers(opts?)`             | TOC without page numbers | `TableOfContents.createNoPageNumbers()`              |
| `TableOfContents.createWithStyles(styles, opts?)`        | TOC with specific styles | `TableOfContents.createWithStyles(['H1','H3'])`      |
| `TableOfContents.createFlat(title?, styles?)`            | Flat TOC (no indent)     | `TableOfContents.createFlat()`                       |
| `TableOfContents.createNumbered(title?, format?)`        | Numbered TOC             | `TableOfContents.createNumbered('TOC', 'roman')`     |
| `TableOfContents.createWithSpacing(spacing, opts?)`      | TOC with custom spacing  | `TableOfContents.createWithSpacing(120)`             |
| `TableOfContents.createWithHyperlinkColor(color, opts?)` | Custom hyperlink color   | `TableOfContents.createWithHyperlinkColor('FF0000')` |

**Note:** All TOC elements are automatically wrapped in an SDT (Structured Document Tag) for native Word integration. This enables Word's "Update Table" button and provides better compatibility with Microsoft Word's TOC features.

#### TOC Configuration Methods

| Method                            | Description                    | Values                     |
| --------------------------------- | ------------------------------ | -------------------------- |
| `setIncludeStyles(styles)`        | Select specific heading styles | `['Heading1', 'Heading3']` |
| `setNumbered(numbered, format?)`  | Enable/disable numbering       | `(true, 'roman')`          |
| `setNoIndent(noIndent)`           | Remove indentation             | `true/false`               |
| `setCustomIndents(indents)`       | Custom indents per level       | `[0, 200, 400]` (twips)    |
| `setSpaceBetweenEntries(spacing)` | Spacing between entries        | `120` (twips)              |
| `setHyperlinkColor(color)`        | Hyperlink color                | `'0000FF'` (default blue)  |
| `setHideInWebLayout(hide)`        | Hide page numbers in web view  | `true/false`               |
| `configure(options)`              | Bulk configuration             | See example below          |

#### TOC Properties

| Property              | Type                                 | Default               | Description                        |
| --------------------- | ------------------------------------ | --------------------- | ---------------------------------- |
| `title`               | `string`                             | `'Table of Contents'` | TOC title                          |
| `levels`              | `number` (1-9)                       | `3`                   | Heading levels to include          |
| `includeStyles`       | `string[]`                           | `undefined`           | Specific styles (overrides levels) |
| `showPageNumbers`     | `boolean`                            | `true`                | Show page numbers                  |
| `useHyperlinks`       | `boolean`                            | `false`               | Use hyperlinks instead of page #s  |
| `hideInWebLayout`     | `boolean`                            | `false`               | Hide page numbers in web layout    |
| `numbered`            | `boolean`                            | `false`               | Number TOC entries                 |
| `numberingFormat`     | `'decimal'/'roman'/'alpha'`          | `'decimal'`           | Numbering format                   |
| `noIndent`            | `boolean`                            | `false`               | Remove all indentation             |
| `customIndents`       | `number[]`                           | `undefined`           | Custom indents in twips            |
| `spaceBetweenEntries` | `number`                             | `0`                   | Spacing in twips                   |
| `hyperlinkColor`      | `string`                             | `'0000FF'`            | Hyperlink color (hex without #)    |
| `tabLeader`           | `'dot'/'hyphen'/'underscore'/'none'` | `'dot'`               | Tab leader character               |

**Example:**

```typescript
// Basic TOC
const simpleToc = TableOfContents.createStandard();
doc.addTableOfContents(simpleToc);

// Select specific styles (e.g., only Heading1 and Heading3)
const customToc = TableOfContents.createWithStyles(["Heading1", "Heading3"]);

// Flat TOC with no indentation
const flatToc = TableOfContents.createFlat("Contents");

// Numbered TOC with roman numerals
const numberedToc = TableOfContents.createNumbered(
  "Table of Contents",
  "roman"
);

// Custom hyperlink color (red instead of blue)
const coloredToc = TableOfContents.createWithHyperlinkColor("FF0000");

// Advanced configuration
const toc = TableOfContents.create()
  .setIncludeStyles(["Heading1", "Heading2", "Heading3"])
  .setNumbered(true, "decimal")
  .setSpaceBetweenEntries(120) // 6pt spacing
  .setHyperlinkColor("0000FF")
  .setNoIndent(false);

// Or use configure() for bulk settings
toc.configure({
  title: "Table of Contents",
  includeStyles: ["Heading1", "CustomHeader"],
  numbered: true,
  numberingFormat: "alpha",
  spaceBetweenEntries: 100,
  hyperlinkColor: "FF0000",
  noIndent: true,
});

// Insert at specific position
doc.insertTocAt(0, toc);
```

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

#### Document Helper Functions

High-level helper methods for common document formatting tasks:

| Method                                   | Description                                                                                 |
| ---------------------------------------- | ------------------------------------------------------------------------------------------- |
| `applyCustomFormattingToExistingStyles()`| Modify Heading1, Heading2, Normal styles with Verdana font, specific spacing, single line spacing, wrap Heading2 in tables, right-align "Top of Document" hyperlinks, set all hyperlinks to blue, and hide TOC page numbers |
| `wrapParagraphInTable(para, options?)`   | Wrap a paragraph in a 1x1 table with optional shading, margins, and width settings         |
| `isParagraphInTable(para)`               | Check if a paragraph is inside a table; returns `{inTable: boolean, cell?: TableCell}`      |
| `updateAllHyperlinkColors(color)`        | Set all hyperlinks in the document to a specific color (e.g., '0000FF' for blue)           |
| `removeAllHeadersFooters()`              | Remove all headers and footers from the document; returns count of headers/footers removed  |

**Example - Using Helper Functions:**

```typescript
import { Document } from 'docxmlater';

const doc = await Document.load('document.docx');

// Apply comprehensive formatting to standard styles
const results = doc.applyCustomFormattingToExistingStyles();
console.log(`Modified styles:`, results);
// Output: { heading1: true, heading2: true, normal: true }

// Wrap a specific paragraph in a table
const para = doc.getParagraphs()[0];
doc.wrapParagraphInTable(para, {
  shading: 'BFBFBF',           // Gray background
  marginLeft: 101,             // 5pt margins
  marginRight: 101,
  tableWidthPercent: 5000      // 100% width
});

// Check if a paragraph is in a table
const { inTable, cell } = doc.isParagraphInTable(para);
if (inTable && cell) {
  console.log('Paragraph is in a table cell');
  cell.setShading({ fill: 'FFFF00' }); // Change to yellow
}

// Set all hyperlinks to blue
doc.updateAllHyperlinkColors('0000FF');

// Remove all headers and footers
const removedCount = doc.removeAllHeadersFooters();
console.log(`Removed ${removedCount} headers and footers`);

await doc.save('formatted.docx');
```

**Note on `applyCustomFormattingToExistingStyles()`:**

This helper function applies a comprehensive set of formatting rules:
- **Heading1**: 18pt black bold Verdana, left aligned, 0pt before/12pt after, single line spacing
- **Heading2**: 14pt black bold Verdana, left aligned, 6pt before/after, single line spacing, wrapped in gray tables (100% width)
- **Normal**: 12pt Verdana, left aligned, 3pt before/after, single line spacing
- **All Styles**: Removes italic and underline formatting
- **Hyperlinks**: "Top of the Document" links are right-aligned with 0pt spacing; all hyperlinks set to blue (#0000FF)
- **Empty Paragraphs**: Empty Heading2 paragraphs are skipped (not wrapped in tables)
- **TOC Elements**: All Table of Contents have page numbers hidden (showPageNumbers=false, hideInWebLayout=true with \n and \z switches)

Per ECMA-376 §17.7.2, direct formatting in document.xml overrides style definitions. This method automatically clears conflicting direct formatting to ensure style changes take effect.

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

#### Twips Conversions
| Function                  | Description         | Example                     |
| ------------------------- | ------------------- | --------------------------- |
| `twipsToPoints(twips)`    | Twips to points     | `twipsToPoints(240)` // 12  |
| `twipsToInches(twips)`    | Twips to inches     | `twipsToInches(1440)` // 1  |
| `twipsToCm(twips)`        | Twips to cm         | `twipsToCm(1440)` // 2.54   |
| `twipsToEmus(twips)`      | Twips to EMUs       | `twipsToEmus(1440)`         |

#### EMUs (English Metric Units) Conversions
| Function                    | Description          | Example                       |
| --------------------------- | -------------------- | ----------------------------- |
| `emusToTwips(emus)`         | EMUs to twips        | `emusToTwips(914400)` // 1440 |
| `emusToInches(emus)`        | EMUs to inches       | `emusToInches(914400)` // 1   |
| `emusToCm(emus)`            | EMUs to cm           | `emusToCm(914400)` // 2.54    |
| `emusToPoints(emus)`        | EMUs to points       | `emusToPoints(914400)` // 72  |
| `emusToPixels(emus, dpi?)`  | EMUs to pixels       | `emusToPixels(914400)` // 96  |

#### Points Conversions
| Function                 | Description        | Example                    |
| ------------------------ | ------------------ | -------------------------- |
| `pointsToTwips(points)`  | Points to twips    | `pointsToTwips(12)` // 240 |
| `pointsToEmus(points)`   | Points to EMUs     | `pointsToEmus(72)`         |
| `pointsToInches(points)` | Points to inches   | `pointsToInches(72)` // 1  |
| `pointsToCm(points)`     | Points to cm       | `pointsToCm(72)` // 2.54   |

#### Inches Conversions
| Function                      | Description         | Example                       |
| ----------------------------- | ------------------- | ----------------------------- |
| `inchesToTwips(inches)`       | Inches to twips     | `inchesToTwips(1)` // 1440    |
| `inchesToEmus(inches)`        | Inches to EMUs      | `inchesToEmus(1)` // 914400   |
| `inchesToPoints(inches)`      | Inches to points    | `inchesToPoints(1)` // 72     |
| `inchesToCm(inches)`          | Inches to cm        | `inchesToCm(1)` // 2.54       |
| `inchesToPixels(inches, dpi)` | Inches to pixels    | `inchesToPixels(1, 96)` // 96 |

#### Centimeters Conversions
| Function                | Description      | Example                     |
| ----------------------- | ---------------- | --------------------------- |
| `cmToTwips(cm)`         | cm to twips      | `cmToTwips(2.54)` // 1440   |
| `cmToEmus(cm)`          | cm to EMUs       | `cmToEmus(2.54)` // 914400  |
| `cmToInches(cm)`        | cm to inches     | `cmToInches(2.54)` // 1     |
| `cmToPoints(cm)`        | cm to points     | `cmToPoints(2.54)` // 72    |
| `cmToPixels(cm, dpi?)`  | cm to pixels     | `cmToPixels(2.54, 96)` // 96|

#### Pixels Conversions
| Function                     | Description         | Example                        |
| ---------------------------- | ------------------- | ------------------------------ |
| `pixelsToEmus(pixels, dpi?)` | Pixels to EMUs      | `pixelsToEmus(96)` // 914400   |
| `pixelsToInches(pixels, dpi?)`| Pixels to inches   | `pixelsToInches(96, 96)` // 1  |
| `pixelsToTwips(pixels, dpi?)`| Pixels to twips     | `pixelsToTwips(96, 96)` // 1440|
| `pixelsToCm(pixels, dpi?)`   | Pixels to cm        | `pixelsToCm(96, 96)` // 2.54   |
| `pixelsToPoints(pixels, dpi?)`| Pixels to points   | `pixelsToPoints(96, 96)` // 72 |

**Note:** Default DPI is 96 for pixel conversions

### ZIP Archive Helper Methods

#### File Operations
| Method                          | Description               | Example                                      |
| ------------------------------- | ------------------------- | -------------------------------------------- |
| `addFile(path, content)`        | Add file to archive       | `handler.addFile('doc.xml', xmlContent)`     |
| `updateFile(path, content)`     | Update existing file      | `handler.updateFile('doc.xml', newContent)`  |
| `removeFile(path)`              | Remove file from archive  | `handler.removeFile('old.xml')`              |
| `renameFile(oldPath, newPath)`  | Rename file               | `handler.renameFile('a.xml', 'b.xml')`       |
| `copyFile(srcPath, destPath)`   | Copy file                 | `handler.copyFile('a.xml', 'copy-a.xml')`    |
| `moveFile(srcPath, destPath)`   | Move file                 | `handler.moveFile('a.xml', 'folder/a.xml')`  |

#### File Retrieval
| Method                    | Description            | Returns         |
| ------------------------- | ---------------------- | --------------- |
| `getFile(path)`           | Get file object        | `ZipFile`       |
| `getFileAsString(path)`   | Get file as string     | `string`        |
| `getFileAsBuffer(path)`   | Get file as buffer     | `Buffer`        |
| `hasFile(path)`           | Check if file exists   | `boolean`       |
| `getFilePaths()`          | Get all file paths     | `string[]`      |
| `getAllFiles()`           | Get all files          | `FileMap`       |

#### Batch Operations
| Method                          | Description                  | Returns        |
| ------------------------------- | ---------------------------- | -------------- |
| `removeFiles(paths[])`          | Remove multiple files        | `number`       |
| `getFilesByExtension(ext)`      | Get files by extension       | `ZipFile[]`    |
| `getTextFiles()`                | Get all text files           | `ZipFile[]`    |
| `getBinaryFiles()`              | Get all binary files         | `ZipFile[]`    |
| `getMediaFiles()`               | Get media files              | `ZipFile[]`    |

#### Archive Information
| Method             | Description               | Returns                  |
| ------------------ | ------------------------- | ------------------------ |
| `getFileCount()`   | Count files in archive    | `number`                 |
| `getTotalSize()`   | Get total size in bytes   | `number`                 |
| `getStats()`       | Get detailed statistics   | `{fileCount, size, ...}` |
| `isEmpty()`        | Check if archive is empty | `boolean`                |

#### Import/Export
| Method                           | Description              | Returns              |
| -------------------------------- | ------------------------ | -------------------- |
| `exportFile(internal, external)` | Export file from archive | `Promise<void>`      |
| `importFile(external, internal)` | Import file to archive   | `Promise<void>`      |

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

### Build Content with Detached Paragraphs

Create paragraphs independently and add them conditionally:

```typescript
import { Paragraph } from "docxmlater";

// Create reusable paragraph templates
const warningTemplate = Paragraph.createFormatted(
  "WARNING: ",
  { bold: true, color: "FF6600" },
  { spacing: { before: 120, after: 120 } }
);

// Clone and customize
const warning1 = warningTemplate.clone();
warning1.addText("Please read the documentation before proceeding.");

// Build content from data
const items = [
  { title: "First Item", description: "Description here" },
  { title: "Second Item", description: "Another description" },
];

items.forEach((item, index) => {
  const titlePara = Paragraph.create(`${index + 1}. `);
  titlePara.addText(item.title, { bold: true });

  const descPara = Paragraph.create(item.description, {
    indentation: { left: 360 },
  });

  doc.addParagraph(titlePara);
  doc.addParagraph(descPara);
});

// See examples/advanced/detached-paragraphs.ts for more patterns
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

**Current:** 1,119 tests passing (97.3% pass rate) | 100% core functionality covered

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

- Use `paragraph.addText()` multiple times for separate text runs
- Use formatting options: `{bold: true}`, `{italic: true}`, etc.
- Use `paragraph.addHyperlink()` for links
- Don't pass XML strings to text methods
- Don't try to embed `<w:t>` or other XML tags in your text

For more details, see the [corruption detection examples](examples/troubleshooting/).

### Layout Conflicts (Massive Whitespace)

**Problem**: Documents show massive whitespace between paragraphs when opened in Word, even though the XML looks correct.

**Cause**: The `pageBreakBefore` property conflicting with `keepNext`/`keepLines` properties. When a paragraph has both `pageBreakBefore` and keep properties set to true, Word's layout engine tries to satisfy contradictory constraints (insert break vs. keep together), resulting in massive whitespace as it struggles to resolve the conflict.

**Why This Causes Problems**:

- `pageBreakBefore` tells Word to insert a page break before the paragraph
- `keepNext` tells Word to keep the paragraph with the next one (no break)
- `keepLines` tells Word to keep all lines together (no break)
- The combination creates layout conflicts that manifest as massive whitespace

**Automatic Conflict Resolution** (v0.28.2+):

The framework now automatically prevents these conflicts by **prioritizing keep properties over page breaks**:

```typescript
// When setting keepNext or keepLines, pageBreakBefore is automatically cleared
const para = new Paragraph()
  .addText("Content")
  .setPageBreakBefore(true) // Set to true
  .setKeepNext(true); // Automatically clears pageBreakBefore

// Result: keepNext=true, pageBreakBefore=false (conflict resolved)
```

**Why This Priority?**

- Keep properties (`keepNext`/`keepLines`) represent explicit user intent to keep content together
- Page breaks are often layout hints that may conflict with document flow
- Removing `pageBreakBefore` eliminates whitespace while preserving the user's intention

**Parsing Documents**:

When loading existing DOCX files with conflicts, they are automatically resolved:

```typescript
// Load document with conflicts
const doc = await Document.load("document-with-conflicts.docx");

// Conflicts are automatically resolved during parsing
// keepNext/keepLines take priority, pageBreakBefore is removed
```

**How It Works**:

1. When `setKeepNext(true)` is called, `pageBreakBefore` is automatically set to `false`
2. When `setKeepLines(true)` is called, `pageBreakBefore` is automatically set to `false`
3. When parsing documents, if both properties exist, `pageBreakBefore` is cleared
4. Keep properties win because they represent explicit user intent

**Manual Override**:

If you need a page break despite keep properties, set it after:

```typescript
const para = new Paragraph()
  .setKeepNext(true) // Set first
  .setPageBreakBefore(true); // Override - you explicitly want this conflict

// But note: This will cause layout issues (whitespace) in Word
```

## Known Limitations

While docXMLater provides comprehensive DOCX manipulation capabilities, there are some features that are not yet fully implemented:

### 1. Table Row Spanning with vMerge

**Status:** FULLY IMPLEMENTED ✅

**What Works:**
- Column spanning (horizontal cell merging) is fully supported
- Row spanning (vertical cell merging) is now fully implemented
- Both horizontal and vertical merging can be combined
- Uses Word's proper `vMerge` attribute ('restart' and 'continue')

**Usage:**
```typescript
// Merge cells horizontally (column spanning)
table.mergeCells(0, 0, 0, 2); // Merge columns 0-2 in row 0

// Merge cells vertically (row spanning)
table.mergeCells(0, 0, 2, 0); // Merge rows 0-2 in column 0

// Merge both horizontally and vertically (2x2 block)
table.mergeCells(0, 0, 1, 1); // Merge 2x2 block starting at (0,0)
```

### 2. Structured Document Tags (SDT) Parsing

**Status:** FULLY IMPLEMENTED ✅

**What Works:**
- Complete SDT parsing from existing documents
- All 9 control types supported (richText, plainText, comboBox, dropDownList, datePicker, checkbox, picture, buildingBlock, group)
- SDT properties fully extracted (id, tag, lock, alias, controlType)
- Nested content parsing (paragraphs, tables, nested SDTs)
- Preserves element order using XMLParser's `_orderedChildren` metadata
- Round-trip operations fully supported

**Control Types Supported:**
- **Rich Text** - Multi-formatted text content
- **Plain Text** - Simple text with optional multiLine support
- **Combo Box** - User-editable dropdown with list items
- **Dropdown List** - Fixed selection from list items
- **Date Picker** - Date selection with format and calendar type
- **Checkbox** - Boolean selection with custom checked/unchecked states
- **Picture** - Image content control
- **Building Block** - Gallery and category-based content
- **Group** - Grouping of other controls

**Usage:**
```typescript
// Load documents with SDTs - fully parsed
const doc = await Document.load('document-with-sdts.docx');

// Access parsed SDT content
const sdts = doc.getBodyElements().filter(el => el instanceof StructuredDocumentTag);
for (const sdt of sdts) {
  console.log('ID:', sdt.getId());
  console.log('Tag:', sdt.getTag());
  console.log('Type:', sdt.getControlType());
  console.log('Content:', sdt.getContent());
}

// Create new SDTs programmatically
const sdt = new StructuredDocumentTag({
  id: 123456,
  tag: 'MyControl',
  controlType: 'richText',
  alias: 'Rich Text Control'
});
sdt.addContent(paragraph);
```

All known limitations have been resolved! For feature requests or bug reports, please visit our [GitHub Issues](https://github.com/ItMeDiaTech/docXMLater/issues).

## Contributing

Contributions are welcome! Please read our [Contributing Guide](CONTRIBUTING.md).

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Recent Updates (v1.3.0)

**Enhanced Document Parsing & Helper Functions:**

- **TOC Parsing** - Parse Table of Contents from existing DOCX files
  - Extract TOC field instructions with all switches (\h, \u, \z, \n, \o, \t)
  - Detect SDT wrappers with `docPartGallery="Table of Contents"`
  - Create `TableOfContentsElement` objects from parsed TOCs
  - Support for modifying TOC field instructions in loaded documents
- **removeAllHeadersFooters() Helper** - New document helper function
  - Removes all headers and footers from the document
  - Deletes header/footer XML files and relationships
  - Returns count of removed headers/footers
- **Enhanced Test Suite** - 1,119/1,150 tests passing (97.3% pass rate)
- **Documentation Updates** - Complete API reference for new helper functions

**Previous Enhancements (v1.2.0):**
- 5 advanced document helper functions
- Enhanced document modification capabilities
- Improved paragraph and table wrapping utilities

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
