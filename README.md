# docxmlater - Professional DOCX Editing Framework

[![Tests](https://img.shields.io/badge/tests-205%20passing-brightgreen)](https://github.com/ItMeDiaTech/docXMLater)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.3-blue)](https://www.typescriptlang.org/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A comprehensive, production-ready TypeScript/JavaScript library for creating, reading, and manipulating Microsoft Word (.docx) documents programmatically.

## Features

- **High-Level Document API** - Create documents with a simple, intuitive interface
- **Text Formatting** - Bold, italic, fonts, colors, and 15+ formatting options
- **Paragraph Formatting** - Alignment, indentation, spacing, keep-with-next
- **Tables** - Full table support with borders, shading, cell merging
- **Images** - Embed PNG, JPEG, GIF images with sizing and positioning
- **Hyperlinks** - Internal and external links with full formatting support
- **Styles System** - 13 built-in styles + custom style creation
- **ZIP Archive Handling** - Low-level DOCX manipulation
- **TypeScript First** - Full type safety with comprehensive definitions
- **Well Tested** - 205 tests, all passing
- **Production Ready** - Used in real-world applications

## Installation

```bash
npm install docxmlater
```

## Quick Start

### Create Your First Document

```typescript
import { Document } from "docxmlater";

const doc = Document.create();

// Add content with styles
doc.createParagraph("My Document").setStyle("Title");
doc.createParagraph("Introduction").setStyle("Heading1");
doc.createParagraph("This is body text with the Normal style.");

// Save
await doc.save("my-document.docx");
```

### Add Formatted Text

```typescript
const doc = Document.create();

// Create a paragraph with mixed formatting
const para = doc.createParagraph();
para.addText("Bold text", { bold: true });
para.addText(" and ");
para.addText("colored text", { color: "FF0000" });

await doc.save("formatted.docx");
```

### Create Tables

```typescript
const doc = Document.create();

// Create a 3x3 table
const table = doc.createTable(3, 3);

// Add borders
table.setAllBorders({ style: "single", size: 8, color: "000000" });

// Populate cells
table.getCell(0, 0)?.createParagraph("Header 1");
table.getCell(0, 1)?.createParagraph("Header 2");

// Add shading to header row
table.getRow(0)?.getCell(0)?.setShading({ fill: "4472C4" });

await doc.save("table.docx");
```

### Add Images

```typescript
import { Document, Image, inchesToEmus } from "docxmlater";

const doc = Document.create();

// Add title
doc.createParagraph("Document with Image").setStyle("Title");

// Create image from file
const image = Image.fromFile("./photo.png");
image.setWidth(inchesToEmus(4), true); // 4 inches wide, maintain aspect ratio

// Add image to document
doc.addImage(image);

// Add caption
doc
  .createParagraph("Figure 1: Sample image")
  .setAlignment("center")
  .addText("Figure 1: Sample image", { italic: true, size: 10 });

await doc.save("with-image.docx");
```

### Use Custom Styles

```typescript
import { Document, Style } from "docxmlater";

const doc = Document.create();

// Create a custom style
const alertStyle = Style.create({
  styleId: "Alert",
  name: "Alert",
  type: "paragraph",
  basedOn: "Normal",
  runFormatting: {
    bold: true,
    color: "FF0000",
    size: 12,
  },
  paragraphFormatting: {
    alignment: "center",
  },
});

doc.addStyle(alertStyle);
doc.createParagraph("Important Warning").setStyle("Alert");

await doc.save("custom-styles.docx");
```

## Core Concepts

### Document Class

The high-level API for creating Word documents:

```typescript
// Create new document
const doc = Document.create({
  properties: {
    title: "My Document",
    creator: "DocXML",
    subject: "Example",
  },
});

// Add content
doc.createParagraph("Content here");
doc.createTable(5, 3);

// Save
await doc.save("output.docx");
```

### Paragraphs and Runs

Paragraphs contain runs of formatted text:

```typescript
const para = doc.createParagraph();

// Add text with different formatting
para.addText("Normal ");
para.addText("Bold", { bold: true });
para.addText(" Italic", { italic: true });

// Set paragraph formatting
para.setAlignment("center");
para.setSpaceAfter(240);
```

### Tables

Create tables with full formatting control:

```typescript
const table = doc.createTable(4, 3);

// Format table
table.setWidth(8640); // Full page width
table.setAllBorders({ style: "single", size: 6 });

// Access cells
const cell = table.getCell(0, 0);
cell?.createParagraph("Cell content");
cell?.setShading({ fill: "D9E1F2" });
cell?.setVerticalAlignment("center");

// Merge cells
cell?.setColumnSpan(3); // Span 3 columns
```

### Styles

13 built-in styles ready to use:

```typescript
// Built-in styles
doc.createParagraph("Title").setStyle("Title");
doc.createParagraph("Subtitle").setStyle("Subtitle");
doc.createParagraph("Chapter 1").setStyle("Heading1");
doc.createParagraph("Section 1.1").setStyle("Heading2");
doc.createParagraph("Body text").setStyle("Normal");

// Check available styles
console.log(doc.getStylesManager().getStyleCount()); // 13

// Create custom styles (see examples)
```

## API Overview

### Document

```typescript
// Creating
Document.create(options?)
Document.load(filePath)
Document.loadFromBuffer(buffer)

// Content
doc.createParagraph(text?)
doc.createTable(rows, columns)
doc.addParagraph(paragraph)
doc.addTable(table)

// Styles
doc.addStyle(style)
doc.getStyle(styleId)
doc.hasStyle(styleId)
doc.getStylesManager()

// Saving
doc.save(filePath)
doc.toBuffer()

// Access
doc.getParagraphs()
doc.getTables()
doc.getImageManager()
doc.getProperties()
doc.setProperties(props)

// Images
doc.addImage(image)
```

### Paragraph

```typescript
// Content
para.addText(text, formatting?)
para.addRun(run)
para.setText(text, formatting?)

// Formatting
para.setAlignment('left' | 'center' | 'right' | 'justify')
para.setLeftIndent(twips)
para.setRightIndent(twips)
para.setSpaceBefore(twips)
para.setSpaceAfter(twips)
para.setLineSpacing(twips, rule?)

// Styles
para.setStyle(styleId)

// Special
para.setKeepNext()
para.setKeepLines()
para.setPageBreakBefore()
```

### Run (Text Formatting)

```typescript
const formatting = {
  bold: true,
  italic: true,
  underline: "single",
  font: "Arial",
  size: 12, // points
  color: "FF0000", // hex without #
  highlight: "yellow",
  strike: true,
  subscript: true,
  superscript: true,
  smallCaps: true,
  allCaps: true,
};
```

### Table

```typescript
// Creation
const table = doc.createTable(rows, cols);

// Access
table.getRow(index);
table.getCell(rowIndex, colIndex);

// Formatting
table.setWidth(twips);
table.setAlignment("left" | "center" | "right");
table.setAllBorders(border);
table.setBorders(borders);
table.setLayout("auto" | "fixed");
```

### Style

```typescript
Style.create(properties);
Style.createNormalStyle();
Style.createHeadingStyle(1 - 9);
Style.createTitleStyle();
Style.createSubtitleStyle();

// Properties
style.setBasedOn(styleId);
style.setParagraphFormatting(formatting);
style.setRunFormatting(formatting);
```

### Image

```typescript
// Creation
Image.fromFile(filePath, width?, height?)
Image.fromBuffer(buffer, extension, width?, height?)
Image.create(properties)

// Sizing (all in EMUs - use inchesToEmus() helper)
image.setWidth(emus, maintainAspectRatio?)
image.setHeight(emus, maintainAspectRatio?)
image.setSize(width, height)

// Properties
image.getWidth()
image.getHeight()
image.getExtension()
image.getImageData()
```

## Built-in Styles

All documents include 13 ready-to-use styles:

| Style             | Description         | Font          | Size    | Color           |
| ----------------- | ------------------- | ------------- | ------- | --------------- |
| **Normal**        | Default paragraph   | Calibri       | 11pt    | Black           |
| **Heading1**      | Major headings      | Calibri Light | 16pt    | Blue (#2E74B5)  |
| **Heading2**      | Section headings    | Calibri Light | 13pt    | Blue (#1F4D78)  |
| **Heading3-9**    | Subsection headings | Calibri Light | 12-11pt | Blue (#1F4D78)  |
| **Title**         | Document title      | Calibri Light | 28pt    | Blue (#2E74B5)  |
| **Subtitle**      | Document subtitle   | Calibri Light | 14pt    | Gray, Italic    |
| **ListParagraph** | List items          | Calibri       | 11pt    | Black, Indented |

See [Using Styles Guide](docs/guides/using-styles.md) for complete documentation.

## Examples

The `examples/` directory contains comprehensive examples:

### Basic Examples (`examples/01-basic/`)

- Creating simple documents
- Reading and modifying DOCX files
- Working with ZIP archives

### Text Formatting (`examples/02-text/`)

- Paragraph formatting examples
- Text formatting (bold, italic, colors)
- Advanced formatting techniques

### Tables (`examples/03-tables/`)

- Simple tables
- Tables with borders and shading
- Complex tables with merged cells

### Styles (`examples/04-styles/`)

- Using built-in styles
- Creating custom styles
- Style inheritance

### Images (`examples/05-images/`)

- Adding images from files and buffers
- Sizing and resizing images
- Multiple images in documents
- Images with text content

### Complete Examples (`examples/06-complete/`)

- Professional reports
- Invoice templates
- Styled documents

Run any example:

```bash
npx ts-node examples/02-text/paragraph-basics.ts
```

## Documentation

- **[Getting Started Guide](docs/guides/getting-started.md)** - Your first document
- **[Using Styles](docs/guides/using-styles.md)** - Complete styles guide
- **[Working with Tables](docs/guides/working-with-tables.md)** - Table guide
- **[API Reference](docs/api/)** - Complete API documentation
- **[Architecture](docs/architecture/)** - System architecture

## Advanced Usage

### Load and Modify Existing Documents

```typescript
// Load existing DOCX
const doc = await Document.load("existing.docx");

// Add content
doc.createParagraph("New paragraph added");

// Save
await doc.save("modified.docx");
```

### Work with Buffers

```typescript
// Create document
const doc = Document.create();
doc.createParagraph("Content");

// Save to buffer
const buffer = await doc.toBuffer();

// Load from buffer
const doc2 = await Document.loadFromBuffer(buffer);
```

### Low-Level ZIP Access

For advanced users, direct ZIP manipulation is available:

```typescript
import { ZipHandler, DOCX_PATHS } from "docxmlater";

const handler = new ZipHandler();
await handler.load("document.docx");

// Direct XML access
const xml = handler.getFileAsString(DOCX_PATHS.DOCUMENT);
handler.updateFile(DOCX_PATHS.DOCUMENT, modifiedXml);

await handler.save("output.docx");
```

## Development

### Setup

```bash
# Install dependencies
npm install

# Build
npm run build

# Run tests
npm test

# Run tests with coverage
npm run test:coverage
```

### Project Structure

```
src/
├── core/          # Document class
├── elements/      # Paragraph, Run, Table
├── formatting/    # Style, StylesManager
├── xml/           # XML generation
├── zip/           # ZIP archive handling
└── utils/         # Validation utilities

tests/
├── core/          # Document tests
├── elements/      # Element tests
├── formatting/    # Style tests (pending)
├── zip/           # ZIP tests
└── utils/         # Utility tests

docs/
├── api/           # API reference
├── guides/        # User guides
└── architecture/  # Architecture docs

examples/
├── 01-basic/      # Basic examples
├── 02-text/       # Text examples
├── 03-tables/     # Table examples
├── 04-styles/     # Style examples
├── 05-images/     # Image examples
└── 06-complete/   # Complete examples
```

## Phase Implementation Status

| Phase                            | Status      | Features                                                |
| -------------------------------- | ----------- | ------------------------------------------------------- |
| **Phase 1: Foundation**          | ✅ Complete | ZIP handling, XML generation, validation                |
| **Phase 2: Core Elements**       | ✅ Complete | Paragraph, Run, text formatting                         |
| **Phase 3: Advanced Formatting** | ✅ Complete | Document API, Tables, Styles, Lists                     |
| **Phase 4: Rich Content**        | ✅ Complete | Images, Headers, Footers, Hyperlinks (NEW!)             |
| **Phase 5: Polish**              | 🚧 Planned  | Track changes, comments, TOC                            |

**Current: 205 tests passing | 20+ source files | ~5,000+ lines of code**

## Requirements

- Node.js 14+
- TypeScript 5.0+ (for development)

## Dependencies

- **jszip** (^3.10.1) - ZIP archive handling

## Browser Support

DocXML works in Node.js environments. Browser support requires bundling with webpack/rollup and may need buffer polyfills.

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Testing

All features are comprehensively tested:

```bash
# Run all tests
npm test

# Run specific test suite
npm test -- tests/core/Document.test.ts

# Watch mode
npm run test:watch

# Coverage report
npm run test:coverage
```

**Test Statistics:**

- 159 tests passing
- 4 test suites
- High code coverage
- Integration tests included

## License

MIT © DiaTech

## Acknowledgments

- Built with [JSZip](https://stuk.github.io/jszip/) for ZIP archive handling
- Follows [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) Office Open XML standard
- Inspired by [python-docx](https://python-docx.readthedocs.io/) and [docx](https://github.com/dolanmiu/docx)

## Support

- **Documentation**: [docs/](docs/)
- **Examples**: [examples/](examples/)
- **Issues**: [GitHub Issues](https://github.com/ItMeDiaTech/docXMLater/issues)

## Roadmap

**Phase 3 (Complete):**

- [x] Document API
- [x] Tables with formatting
- [x] Styles system
- [x] Lists and numbering

**Phase 4 (Complete):**

- [x] Images and media
- [x] Headers and footers
- [x] Page sections
- [x] Hyperlinks (**NEW in v0.2.0!**)

**Phase 5 (Future):**

- [ ] Track changes
- [ ] Comments
- [ ] Table of contents
- [ ] Fields

## Related Projects

- **[python-docx](https://python-docx.readthedocs.io/)** - Python DOCX library
- **[docx](https://github.com/dolanmiu/docx)** - JavaScript DOCX library
- **[mammoth.js](https://github.com/mwilliamson/mammoth.js)** - Convert DOCX to HTML

---

**Ready to get started?** Check out the [Quick Start Guide](docs/guides/getting-started.md) or explore the [examples](examples/).
