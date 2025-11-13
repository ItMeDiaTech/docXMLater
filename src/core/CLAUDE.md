# Core Module Documentation

The `core` module contains the foundational classes for document management, parsing, generation, and validation.

## Module Overview

The core module provides the high-level API for working with DOCX documents. It handles document lifecycle, XML parsing/generation, and validation.

**Location:** `src/core/`

**Key Classes:**
- `Document` - Main document interface
- `DocumentParser` - Parses XML to document structure
- `DocumentGenerator` - Generates XML from document structure
- `DocumentValidator` - Validates document size and structure
- `RelationshipManager` - Manages document relationships
- `Relationship` - Represents individual relationships
- `BaseManager` - Base class for manager pattern

## Architecture

### Document Class

**File:** `src/core/Document.ts`

The Document class is the primary entry point for all document operations. It provides a high-level API that abstracts away ZIP and XML complexities.

**Core Responsibilities:**
1. Document lifecycle (create, load, save, dispose)
2. Content management (paragraphs, tables, sections)
3. Style and numbering management
4. Relationship management (images, hyperlinks)
5. Search and replace operations
6. Document statistics (word count, size estimation)

**Key Design Patterns:**
- **Factory Pattern**: `Document.create()`, `Document.load()`
- **Disposal Pattern**: `dispose()` for resource cleanup
- **Manager Pattern**: Delegates to specialized managers (StylesManager, NumberingManager, etc.)
- **Builder Pattern**: Fluent API for content creation

**Memory Management:**
The Document class implements explicit memory management via the `dispose()` method:
- Clears body elements array
- Disposes managers (images, bookmarks, etc.)
- Clears relationship caches
- Resets internal state

Users MUST call `dispose()` when done with a document to prevent memory leaks:
```typescript
const doc = await Document.load('file.docx');
try {
  // Work with document
  await doc.save('output.docx');
} finally {
  doc.dispose(); // Always dispose
}
```

### DocumentParser Class

**File:** `src/core/DocumentParser.ts`

Converts XML from DOCX files into JavaScript object models.

**Core Responsibilities:**
1. Parse `word/document.xml` into Document structure
2. Parse paragraphs, runs, tables, images, etc.
3. Parse complex elements (fields, hyperlinks, bookmarks)
4. Handle special characters and formatting
5. Merge fragmented hyperlinks
6. Preserve unknown elements

**Parsing Strategy:**
- **Position-based parsing** (ReDoS-safe) using XMLParser
- **Element-by-element traversal** of XML tree
- **Defensive parsing** with error recovery
- **Preservation mode** for unknown elements

**Special Handling:**
- **Hyperlink Defragmentation**: Merges consecutive hyperlinks with same URL
- **Special Characters**: Converts `<w:tab/>`, `<w:br/>` to string equivalents
- **Complex Fields**: Parses field instructions and content
- **Content Controls**: Handles SDT (Structured Document Tags)

**Error Handling:**
- Non-critical errors are logged but don't halt parsing
- Unknown elements are preserved for round-trip fidelity
- Malformed XML generates clear error messages

### DocumentGenerator Class

**File:** `src/core/DocumentGenerator.ts`

Converts JavaScript object models into DOCX XML files.

**Core Responsibilities:**
1. Generate `word/document.xml` from Document structure
2. Generate all required XML files (styles, numbering, settings, etc.)
3. Generate relationships for images, hyperlinks, etc.
4. Ensure ECMA-376 compliance
5. Handle special cases (empty documents, large documents)

**Generation Strategy:**
- **Template-based**: Uses minimal XML templates
- **On-demand generation**: Only generates what's needed
- **Validation**: Checks structure before generation
- **Optimization**: Omits default values, empty elements

**XML Files Generated:**
- `[Content_Types].xml` - MIME type declarations
- `_rels/.rels` - Package relationships
- `word/document.xml` - Main document content
- `word/styles.xml` - Style definitions
- `word/numbering.xml` - List numbering
- `word/settings.xml` - Document settings
- `word/fontTable.xml` - Font declarations
- `word/_rels/document.xml.rels` - Document relationships
- `docProps/core.xml` - Core properties
- `docProps/app.xml` - Application properties

**ECMA-376 Compliance:**
- Proper namespace declarations
- Correct element ordering per schema
- Valid attribute values
- Relationship ID management
- UTF-8 encoding

### DocumentValidator Class

**File:** `src/core/DocumentValidator.ts`

Validates document structure and estimates file size.

**Core Responsibilities:**
1. Estimate final DOCX file size
2. Validate memory usage
3. Check for size limits
4. Warn about potential issues

**Size Estimation Algorithm:**
```
baseSize = minimal DOCX structure (~5KB)
+ textSize = character count * 1.5 (accounts for XML markup)
+ imageSize = sum of image buffer sizes
+ tableSize = cell count * 200 (estimate)
+ compressionFactor = 0.3 (ZIP compression)

estimatedSize = (baseSize + textSize + imageSize + tableSize) * compressionFactor
```

**Memory Validation:**
- Checks current Node.js RSS (Resident Set Size)
- Warns if memory usage > threshold (default 80%)
- Prevents out-of-memory errors during large document operations

### RelationshipManager Class

**File:** `src/core/RelationshipManager.ts`

Manages relationships between document parts.

**Core Responsibilities:**
1. Track relationships (images, hyperlinks, headers, footers)
2. Generate unique relationship IDs
3. Generate relationship XML
4. Prevent duplicate relationships

**Relationship Types:**
- `image` - Images in media folder
- `hyperlink` - External URLs (mode: External)
- `styles` - Styles definition
- `numbering` - Numbering definition
- `header` - Header part
- `footer` - Footer part
- `fontTable` - Font table
- `settings` - Document settings

**ID Generation:**
- Sequential IDs: `rId1`, `rId2`, `rId3`, etc.
- Thread-safe counter
- No ID reuse within document

### Relationship Class

**File:** `src/core/Relationship.ts`

Represents a single relationship.

**Properties:**
- `id` - Unique relationship ID
- `type` - Relationship type URL
- `target` - Target path/URL
- `targetMode` - Internal or External

**Type URLs (per ECMA-376):**
```typescript
const RELATIONSHIP_TYPES = {
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  // ... etc
};
```

### BaseManager Class

**File:** `src/core/BaseManager.ts`

Abstract base class for manager pattern implementation.

**Purpose:**
Provides common functionality for all manager classes:
- State management
- Lifecycle methods
- Common utilities

**Subclasses:**
- `StylesManager`
- `NumberingManager`
- `ImageManager`
- `BookmarkManager`
- `CommentManager`
- `RevisionManager`
- `FootnoteManager`
- `EndnoteManager`
- `DrawingManager`
- `HeaderFooterManager`

## Data Flow

### Loading a Document
```
1. Document.load(filepath)
   ↓
2. ZipHandler.load(filepath)
   ↓
3. Extract XML files from ZIP
   ↓
4. DocumentParser.parse(documentXML)
   ↓
5. Parse styles, numbering, relationships
   ↓
6. Build Document object model
   ↓
7. Return Document instance
```

### Saving a Document
```
1. Document.save(filepath)
   ↓
2. DocumentValidator.estimateSize()
   ↓
3. DocumentGenerator.generate(document)
   ↓
4. Generate all XML files
   ↓
5. RelationshipManager.generateXML()
   ↓
6. ZipWriter.create()
   ↓
7. Add files to ZIP archive
   ↓
8. Write to filesystem
```

### Search and Replace
```
1. Document.replaceText(pattern, replacement)
   ↓
2. Iterate body elements
   ↓
3. For each Paragraph:
   - Get all runs
   - Search run text
   - Split run at match boundaries
   - Replace matched runs
   ↓
4. Return match count
```

## Best Practices

### Always Dispose Documents
```typescript
// ✓ Correct
const doc = await Document.load('file.docx');
try {
  await doc.save('output.docx');
} finally {
  doc.dispose();
}

// ✗ Wrong - memory leak!
const doc = await Document.load('file.docx');
await doc.save('output.docx');
// Missing dispose()
```

### Use Buffer Operations for Performance
```typescript
// Faster for multiple operations
const buffer = fs.readFileSync('input.docx');
const doc = await Document.loadFromBuffer(buffer);
// ... modifications ...
const outputBuffer = await doc.toBuffer();
fs.writeFileSync('output.docx', outputBuffer);
doc.dispose();
```

### Check Size Before Saving
```typescript
const sizeInfo = doc.estimateSize();
console.log(`Estimated size: ${sizeInfo.estimatedBytes} bytes`);

if (sizeInfo.estimatedBytes > 10 * 1024 * 1024) {
  console.warn('Document is large (>10MB)');
}
```

### Handle Errors Gracefully
```typescript
try {
  const doc = await Document.load('file.docx');
  try {
    await doc.save('output.docx');
  } finally {
    doc.dispose();
  }
} catch (error) {
  if (error instanceof InvalidDocxError) {
    console.error('Invalid DOCX file:', error.message);
  } else {
    throw error;
  }
}
```

## Testing

The core module has comprehensive test coverage:

**File:** `tests/core/Document.test.ts`
- Document creation and loading (50+ tests)
- Content management (30+ tests)
- Search and replace (20+ tests)
- Statistics and validation (15+ tests)

**File:** `tests/core/DocumentParser.test.ts`
- XML parsing (40+ tests)
- Special character handling (15+ tests)
- Error recovery (10+ tests)

**File:** `tests/core/DocumentProperties.test.ts`
- Property management (20+ tests)

**File:** `tests/core/HyperlinkParsing.test.ts`
- Hyperlink defragmentation (15+ tests)

**Total: 200+ tests covering core functionality**

## Performance Characteristics

### Load Performance
- Small documents (<100KB): <50ms
- Medium documents (100KB-1MB): 100-500ms
- Large documents (>1MB): 500ms-2s

### Save Performance
- Similar to load performance
- Buffer operations 20-30% faster than file I/O

### Memory Usage
- Base overhead: ~2MB per Document instance
- Text: ~2 bytes per character (includes XML markup)
- Images: Full buffer size (no compression in memory)
- Tables: ~200 bytes per cell

### Optimization Tips
1. Use `dispose()` to free memory
2. Process documents in batches
3. Use buffer operations for speed
4. Avoid loading entire documents for small edits
5. Stream large operations when possible

## Future Enhancements

Potential improvements for core module:

1. **Streaming Parser**: Parse documents without loading entire XML into memory
2. **Lazy Loading**: Load document parts on-demand
3. **Incremental Saving**: Save only modified parts
4. **Parallel Processing**: Parse independent elements concurrently
5. **Schema Validation**: Validate against ECMA-376 schema
6. **Format Conversion**: Convert to/from other formats (HTML, PDF)

## See Also

- `src/zip/CLAUDE.md` - ZIP archive handling
- `src/xml/CLAUDE.md` - XML generation and parsing
- `src/elements/CLAUDE.md` - Document elements
- `src/formatting/CLAUDE.md` - Style and numbering systems
