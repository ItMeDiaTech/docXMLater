# docXMLater Framework Skill

## Overview

docXMLater is a custom TypeScript library (npm: `docxml`) for DOCX manipulation built by ItMeDiaTech. Used in Documentation Hub Electron app for document processing workflows.

## Architecture

### Two-Tier Design

1. **High-Level API**: Document-centric operations (`Document`, `Paragraph`, `Table`, `Style`)
2. **Low-Level API**: Direct ZIP/XML manipulation (`ZipHandler`, XML access)

### Core Structure

```
src/
├── core/           # Document class
├── elements/       # Paragraph, Run, Table
├── formatting/     # Style, StylesManager
├── xml/            # XML generation
├── zip/            # ZipHandler (JSZip wrapper)
└── utils/          # Validation
```

### Dependencies

- `jszip` (^3.10.1) - ZIP handling only
- No other DOCX libraries (mammoth, docx.js, etc.)

## Key Patterns

### Document Lifecycle

```typescript
// Create new
const doc = Document.create();
doc.createParagraph("Content");
await doc.save("output.docx");

// Load existing
const doc = await Document.load("existing.docx");
doc.createParagraph("New content");
await doc.save("modified.docx");

// Buffer operations
const buffer = await doc.toBuffer();
const doc2 = await Document.loadFromBuffer(buffer);
```

### Element Hierarchy

- **Document** → contains Paragraphs, Tables, Images
- **Paragraph** → contains Runs (formatted text segments)
- **Run** → text with formatting (bold, color, font, etc.)
- **Table** → Rows → Cells → Paragraphs

### Built-in Styles

13 ready-to-use styles: Normal, Heading1-9, Title, Subtitle, ListParagraph

### Low-Level Access

```typescript
import { ZipHandler, DOCX_PATHS } from "docxml";

const handler = new ZipHandler();
await handler.load("document.docx");

// Direct XML manipulation
const xml = handler.getFileAsString(DOCX_PATHS.DOCUMENT);
const modified = modifyXml(xml);
handler.updateFile(DOCX_PATHS.DOCUMENT, modified);

await handler.save("output.docx");
```

## Common Issues & Solutions

### 1. Document Not Opening After Edit

**Cause**: Invalid XML, broken relationships, or ZIP corruption
**Fix**:

- Use high-level API when possible (handles XML generation)
- If using ZipHandler: validate XML before `updateFile()`
- Never modify XML strings directly without parsing
- Ensure compression settings match: `{type: 'nodebuffer', compression: 'DEFLATE'}`

### 2. Formatting Not Applied

**Cause**: Style doesn't exist or Run formatting incorrect
**Fix**:

```typescript
// Check style exists
if (!doc.hasStyle("CustomStyle")) {
  doc.addStyle(customStyleDefinition);
}

// Apply to paragraph
para.setStyle("CustomStyle");

// Or format runs directly
para.addText("Bold text", { bold: true, color: "FF0000" });
```

### 3. Tables Not Rendering

**Cause**: Missing borders, invalid dimensions, or cell formatting
**Fix**:

```typescript
const table = doc.createTable(rows, cols);
table.setAllBorders({ style: "single", size: 8, color: "000000" });
table.setWidth(8640); // Full page width in twips

// Access cells safely
const cell = table.getCell(row, col);
if (cell) {
  cell.createParagraph("Content");
  cell.setShading({ fill: "D9E1F2" });
}
```

### 4. Images Not Displaying

**Cause**: Incorrect EMU sizing, missing relationships, or invalid buffer
**Fix**:

```typescript
import { Image, inchesToEmus } from "docxml";

const image = Image.fromFile("photo.png");
image.setWidth(inchesToEmus(4), true); // 4 inches, maintain aspect
doc.addImage(image);
```

### 5. Performance with Large Documents

**Symptoms**: Slow processing, high memory usage
**Fix**:

- Use buffer operations instead of repeated file I/O
- Process tables/paragraphs in batches
- Consider streaming for files >10MB
- Don't repeatedly call `save()` - build then save once

## Best Practices

### In Electron Main Process

```typescript
import { Document } from "docxml";

ipcMain.handle("process-document", async (event, filePath) => {
  try {
    const doc = await Document.load(filePath);

    // Modifications
    doc.createParagraph("Processed");

    await doc.save(filePath);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});
```

### Backup Strategy

```typescript
const backupPath = filePath.replace(".docx", `_backup_${Date.now()}.docx`);
await fs.copyFile(filePath, backupPath);

// Then modify original
const doc = await Document.load(filePath);
// ... modifications
await doc.save(filePath);
```

### Style Management

```typescript
// Check before use
if (!doc.hasStyle("Alert")) {
  const alertStyle = Style.create({
    styleId: "Alert",
    name: "Alert",
    type: "paragraph",
    basedOn: "Normal",
    runFormatting: { bold: true, color: "FF0000" },
    paragraphFormatting: { alignment: "center" },
  });
  doc.addStyle(alertStyle);
}

doc.createParagraph("Warning!").setStyle("Alert");
```

### Error Handling

```typescript
try {
  const doc = await Document.load(filePath);
  // operations
  await doc.save(outputPath);
} catch (error) {
  if (error.message.includes("ZIP")) {
    // Corrupted file
  } else if (error.message.includes("XML")) {
    // Invalid structure
  }
  // Restore from backup if needed
}
```

## Integration with Documentation Hub

### WordDocumentProcessor Usage

The app uses a `WordDocumentProcessor` service that wraps docXMLater:

```typescript
import { Document } from "docxml";

class WordDocumentProcessor {
  async processDocument(filePath: string, options: Options) {
    const doc = await Document.load(filePath);

    // Use ZipHandler for hyperlink processing (low-level)
    const handler = new ZipHandler();
    await handler.load(filePath);

    // Modify relationships in document.xml.rels
    // Modify hyperlinks in document.xml

    await handler.save(filePath);

    return result;
  }
}
```

### When to Use Each API

**High-Level API** (Document, Paragraph, Table):

- Creating new content
- Applying styles
- Standard document operations
- Adding images

**Low-Level API** (ZipHandler):

- Modifying hyperlink relationships
- Editing `_rels` files
- Batch XML transformations
- Custom XML that high-level API doesn't support

## Debugging Checklist

1. **File won't open in Word**:

   - Extract ZIP: `unzip document.docx -d extracted/`
   - Check `[Content_Types].xml` is present
   - Validate `word/document.xml` structure
   - Verify `word/_rels/document.xml.rels` exists

2. **Content missing after edit**:

   - Check if using correct API (high vs low-level)
   - Verify `save()` was called
   - Confirm no errors were swallowed
   - Check backup file to compare

3. **Styles not working**:

   - Verify style exists: `doc.hasStyle(styleId)`
   - Check `word/styles.xml` in ZIP
   - Ensure styleId matches exactly (case-sensitive)

4. **TypeScript errors**:
   - Import from 'docxml', not 'docXMLater'
   - Check types: `Document`, `Paragraph`, `Style`, etc.
   - Verify async/await usage on file operations

## Testing Patterns

```typescript
import { Document } from "docxml";
import * as fs from "fs";

describe("Document Processing", () => {
  it("should modify document", async () => {
    const doc = Document.create();
    doc.createParagraph("Test");

    const buffer = await doc.toBuffer();
    expect(buffer).toBeInstanceOf(Buffer);

    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.getParagraphs()).toHaveLength(1);
  });
});
```

## Common TypeScript Issues

### Import Errors

```typescript
// ✅ Correct
import { Document, Style, Image } from "docxml";

// ❌ Wrong
import Document from "docxml";
import { Document } from "docXMLater";
```

### Async/Await

```typescript
// ✅ All file operations are async
await Document.load(path);
await doc.save(path);
await doc.toBuffer();

// ❌ Forgetting await causes issues
Document.load(path); // Returns Promise, not Document!
```

### Type Safety

```typescript
// ✅ Null checks when accessing elements
const cell = table.getCell(0, 0);
if (cell) {
  cell.createParagraph("Safe");
}

// ❌ Assuming element exists
table.getCell(0, 0).createParagraph("Crash!"); // TypeError if null
```

## Performance Tips

1. **Batch Operations**: Group multiple edits, save once
2. **Buffer Usage**: Load once, process, save once
3. **Avoid Redundant Loads**: Cache Document object if processing multiple times
4. **Style Reuse**: Create styles once, apply to multiple elements
5. **Large Files**: Consider low-level API with streaming for >50MB

## Validation Utilities

The framework includes validators:

```typescript
import { validateDocument } from "docxml/utils";

const isValid = validateDocument(xmlString);
if (!isValid) {
  throw new Error("Invalid DOCX structure");
}
```

## Key Differences from Other Libraries

**vs python-docx**: More low-level control, TypeScript native
**vs docx.js**: Lighter, custom implementation, two-tier architecture
**vs mammoth**: Can create/edit, not just convert to HTML

## Resources

- GitHub: https://github.com/ItMeDiaTech/docXMLater
- npm: `docxml`
- Examples: `/examples` directory in repo
- Docs: `/docs` directory in repo
- Tests: 159 passing tests in `/tests`
