# Managers Module Documentation

The `managers` module contains specialized manager classes that handle specific document features and resources.

## Module Overview

Manager classes follow the **Manager Pattern** - they encapsulate complex operations for specific document features (images, drawings, bookmarks, etc.) and provide a clean API for the Document class.

**Location:** `src/managers/`

**Key Classes:**
- `DrawingManager` - Manages shapes, text boxes, and drawing elements
- (Other managers are in their respective module folders)

**Related Managers in Other Modules:**
- `ImageManager` (`src/elements/`) - Image handling
- `HeaderFooterManager` (`src/elements/`) - Headers and footers
- `BookmarkManager` (`src/elements/`) - Bookmarks
- `CommentManager` (`src/elements/`) - Comments
- `RevisionManager` (`src/elements/`) - Track changes
- `FootnoteManager` (`src/elements/`) - Footnotes
- `EndnoteManager` (`src/elements/`) - Endnotes
- `FontManager` (`src/elements/`) - Fonts
- `StylesManager` (`src/formatting/`) - Styles
- `NumberingManager` (`src/formatting/`) - Lists
- `RelationshipManager` (`src/core/`) - Relationships

## Manager Pattern

### Design Philosophy

The Manager Pattern in docXMLater provides:
1. **Separation of Concerns** - Each manager handles one feature area
2. **Resource Management** - Centralized tracking of resources (IDs, relationships)
3. **Validation** - Ensures data integrity before use
4. **Lifecycle Management** - Handles initialization, usage, and cleanup

### Common Manager Interface

All managers implement a common pattern:

```typescript
class ExampleManager {
  private items: Map<string, Item>;

  constructor() {
    this.items = new Map();
  }

  // Add item
  addItem(item: Item): void {
    this.validateItem(item);
    this.items.set(item.id, item);
  }

  // Get item
  getItem(id: string): Item | undefined {
    return this.items.get(id);
  }

  // Get all items
  getAllItems(): Item[] {
    return Array.from(this.items.values());
  }

  // Remove item
  removeItem(id: string): boolean {
    return this.items.delete(id);
  }

  // Validation
  private validateItem(item: Item): void {
    // Validation logic
  }

  // Cleanup
  dispose(): void {
    this.items.clear();
  }
}
```

## DrawingManager

**File:** `src/managers/DrawingManager.ts`

Manages complex drawing elements including shapes, text boxes, and preserved drawings from parsed documents.

### Core Responsibilities

1. **Preserve Unknown Drawings** - Keeps drawings during round-trip parsing
2. **Track Drawing Elements** - Maintains registry of all drawings
3. **Generate Drawing XML** - Creates proper WordprocessingML for drawings
4. **Manage Drawing IDs** - Assigns unique IDs to drawing elements

### Drawing Types

```typescript
type DrawingType =
  | 'shape'           // Basic shapes (rectangle, ellipse, etc.)
  | 'picture'         // Images and pictures
  | 'textBox'         // Text boxes
  | 'diagram'         // SmartArt diagrams
  | 'chart'           // Charts and graphs
  | 'preserved';      // Unknown/preserved drawings
```

### PreservedDrawing

When parsing documents, unknown drawing elements are preserved as-is:

```typescript
interface PreservedDrawing {
  type: 'preserved';
  xml: string;        // Original XML
  location: string;   // Where it appeared (paragraph ID, etc.)
}
```

This ensures round-trip fidelity - drawings not yet supported by docXMLater are preserved without corruption.

### Usage

```typescript
// Get drawing manager from document
const drawingManager = doc.getDrawingManager();

// Add a shape
const shape = new Shape('rectangle', {
  width: 400,
  height: 300,
  fill: { color: 'FF0000' },
  outline: { color: '000000', width: 2 }
});
drawingManager.addDrawing(shape);

// Get all drawings
const drawings = drawingManager.getAllDrawings();

// Dispose when done
drawingManager.dispose();
```

### XML Generation

The DrawingManager generates proper WordprocessingML drawing elements:

```xml
<w:p>
  <w:r>
    <w:drawing>
      <wp:anchor>
        <wp:simplePos x="0" y="0"/>
        <wp:extent cx="400000" cy="300000"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/shape">
            <wps:wsp>
              <!-- Shape definition -->
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>
  </w:r>
</w:p>
```

## Related Managers

### ImageManager

**Location:** `src/elements/ImageManager.ts`

Manages embedded images in the document.

**Core Features:**
- Image buffering and storage
- Relationship management for images
- Format detection (PNG, JPEG, GIF, SVG)
- Size calculation and validation
- Media folder organization

**Usage:**
```typescript
const imageManager = doc.getImageManager();

// Add image
const imageBuffer = fs.readFileSync('photo.jpg');
const image = await imageManager.addImage(imageBuffer, 'jpg');

// Get relationship ID for embedding
const relId = image.getRelationshipId();

// Estimate total image size
const totalSize = imageManager.getTotalImageSize();
```

### HeaderFooterManager

**Location:** `src/elements/HeaderFooterManager.ts`

Manages headers and footers for document sections.

**Core Features:**
- Header types: default, first page, even pages
- Footer types: default, first page, even pages
- Relationship management
- XML generation for header/footer parts

**Usage:**
```typescript
const hfManager = doc.getHeaderFooterManager();

// Create header
const header = new Header('default');
const headerPara = header.createParagraph();
headerPara.addText('Document Title', { bold: true });

hfManager.addHeader(header);

// Create footer with page numbers
const footer = new Footer('default');
const footerPara = footer.createParagraph();
footerPara.addText('Page ');
footerPara.addField('PAGE');

hfManager.addFooter(footer);
```

### BookmarkManager

**Location:** `src/elements/BookmarkManager.ts`

Manages bookmarks for internal document navigation.

**Core Features:**
- Bookmark registration and lookup
- Unique ID generation
- Bookmark start/end pairing
- Cross-reference support

**Usage:**
```typescript
const bookmarkManager = doc.getBookmarkManager();

// Create bookmark
const bookmark = new Bookmark('Section1', 'Introduction Section');
bookmarkManager.addBookmark(bookmark);

// Get bookmark by name
const found = bookmarkManager.getBookmarkByName('Section1');

// Create hyperlink to bookmark
const link = Hyperlink.createInternal('Section1', 'Go to Introduction');
para.addHyperlink(link);
```

### CommentManager

**Location:** `src/elements/CommentManager.ts`

Manages document comments and annotations.

**Core Features:**
- Comment storage and retrieval
- Author tracking
- Date/time stamps
- Reply threading
- Comment range management

**Usage:**
```typescript
const commentManager = doc.getCommentManager();

// Create comment
const comment = new Comment({
  author: 'John Doe',
  date: new Date(),
  text: 'This needs revision'
});
commentManager.addComment(comment);

// Add reply to comment
const reply = new Comment({
  author: 'Jane Smith',
  date: new Date(),
  text: 'Agreed, will update',
  parentId: comment.getId()
});
commentManager.addComment(reply);
```

### RevisionManager

**Location:** `src/elements/RevisionManager.ts`

Manages track changes and document revisions.

**Core Features:**
- Insertion tracking
- Deletion tracking
- Formatting change tracking
- Author and date metadata
- RSID (Revision Save ID) management

**Usage:**
```typescript
const revisionManager = doc.getRevisionManager();

// Track insertion
const insertion = new Revision('insert', {
  author: 'John Doe',
  date: new Date(),
  content: run
});
revisionManager.addRevision(insertion);

// Track deletion
const deletion = new Revision('delete', {
  author: 'Jane Smith',
  date: new Date(),
  content: deletedRun
});
revisionManager.addRevision(deletion);
```

### FootnoteManager & EndnoteManager

**Location:** `src/elements/FootnoteManager.ts`, `src/elements/EndnoteManager.ts`

Manage footnotes and endnotes for citations and references.

**Core Features:**
- Footnote/endnote numbering
- Reference management
- Separator and continuation separator
- Custom formatting

**Usage:**
```typescript
const footnoteManager = doc.getFootnoteManager();

// Create footnote
const footnote = new Footnote('normal', {
  id: 1,
  content: [
    para.addText('Reference text goes here')
  ]
});
footnoteManager.addFootnote(footnote);

// Reference footnote in text
para.addFootnoteReference(1);
```

### FontManager

**Location:** `src/elements/FontManager.ts`

Manages font table and font declarations.

**Core Features:**
- Font registration
- Font embedding (subset or full)
- Fallback font specification
- Font family declarations

**Usage:**
```typescript
const fontManager = doc.getFontManager();

// Register font
fontManager.addFont({
  name: 'Calibri',
  family: 'swiss',
  pitch: 'variable',
  charset: '0'
});

// Generate font table XML
const fontTableXml = fontManager.generateFontTableXml();
```

## Manager Lifecycle

### Initialization

Managers are created when a Document is instantiated:

```typescript
class Document {
  private stylesManager: StylesManager;
  private numberingManager: NumberingManager;
  private imageManager: ImageManager;
  // ... other managers

  constructor() {
    this.stylesManager = new StylesManager();
    this.numberingManager = new NumberingManager();
    this.imageManager = new ImageManager();
    // ... initialize other managers
  }
}
```

### Usage Phase

During document editing, managers are accessed via getters:

```typescript
const doc = Document.create();

// Access managers
const styles = doc.getStylesManager();
const numbering = doc.getNumberingManager();
const images = doc.getImageManager();

// Use managers
styles.addStyle(customStyle);
images.addImage(buffer, 'jpg');
```

### Cleanup

When disposing a document, all managers are cleaned up:

```typescript
dispose(): void {
  this.stylesManager.dispose();
  this.numberingManager.dispose();
  this.imageManager.dispose();
  this.bookmarkManager.dispose();
  this.commentManager.dispose();
  this.revisionManager.dispose();
  this.footnoteManager.dispose();
  this.endnoteManager.dispose();
  this.drawingManager.dispose();
  this.headerFooterManager.dispose();
  this.fontManager.dispose();
}
```

## Best Practices

### 1. Always Access Managers Through Document

```typescript
// ✓ Correct
const styles = doc.getStylesManager();
styles.addStyle(customStyle);

// ✗ Wrong - don't create managers directly
const styles = new StylesManager();  // Not linked to document!
```

### 2. Validate Before Adding

```typescript
// ✓ Correct
if (!bookmarkManager.hasBookmark('Section1')) {
  bookmarkManager.addBookmark(bookmark);
}

// ✗ Wrong - may cause duplicate errors
bookmarkManager.addBookmark(bookmark);
```

### 3. Use Manager-Specific Methods

```typescript
// ✓ Correct - uses manager method
const totalImageSize = imageManager.getTotalImageSize();

// ✗ Wrong - manual calculation error-prone
let total = 0;
imageManager.getAllImages().forEach(img => {
  total += img.getBuffer().length;
});
```

### 4. Cleanup When Done

```typescript
// ✓ Correct
const doc = Document.create();
try {
  // Use document and managers
} finally {
  doc.dispose();  // Cleans up all managers
}

// ✗ Wrong - memory leak
const doc = Document.create();
// ... use document ...
// Missing dispose()
```

## Testing

Manager classes have dedicated test suites:

**Files:**
- `tests/elements/ImageProperties.test.ts` - Image manager tests
- `tests/elements/Field.test.ts` - Field manager tests
- `tests/elements/DrawingElements.test.ts` - Drawing manager tests
- `tests/elements/ContentControls.test.ts` - Content control tests
- `tests/core/CommentParsing.test.ts` - Comment manager tests
- `tests/core/BookmarkParsing.test.ts` - Bookmark manager tests

**Coverage:**
- Manager creation and initialization
- Resource addition and retrieval
- Validation logic
- ID generation and uniqueness
- Cleanup and disposal

## Performance Considerations

### Memory Usage

- **Managers maintain in-memory maps** of their resources
- **Large documents** with many images/comments can consume significant memory
- **Use dispose()** to free memory when done

### Lookup Performance

- **Map-based storage** provides O(1) lookup by ID
- **Array methods** (getAllItems) are O(n)
- **Cache computed values** when possible

### Optimization Tips

1. **Batch operations** when adding multiple items
2. **Validate once** on add, not on every access
3. **Use weak references** for large binary data when possible
4. **Dispose managers** as soon as document operations complete

## Extending the Framework

### Creating a Custom Manager

To add a new manager:

1. **Extend BaseManager** (if it exists) or create standalone
2. **Implement core methods**: add, get, remove, dispose
3. **Add validation logic**
4. **Generate appropriate XML**
5. **Add to Document class**
6. **Write comprehensive tests**

Example:

```typescript
class CustomManager {
  private items: Map<string, CustomItem>;

  constructor() {
    this.items = new Map();
  }

  addItem(item: CustomItem): void {
    this.validateItem(item);
    this.items.set(item.id, item);
  }

  getItem(id: string): CustomItem | undefined {
    return this.items.get(id);
  }

  private validateItem(item: CustomItem): void {
    if (!item.id) {
      throw new Error('Item must have an ID');
    }
    if (this.items.has(item.id)) {
      throw new Error(`Item with ID ${item.id} already exists`);
    }
  }

  generateXml(): string {
    // Generate XML for all items
    return '...';
  }

  dispose(): void {
    this.items.clear();
  }
}
```

## Troubleshooting

### Issue: Manager Not Initialized

**Problem:** Accessing manager before document creation
**Solution:** Always create/load document first
**Fix:**
```typescript
const doc = Document.create();
const manager = doc.getImageManager();  // Now safe
```

### Issue: Resource Not Found

**Problem:** Accessing resource by ID that doesn't exist
**Solution:** Check existence first
**Fix:**
```typescript
const bookmark = bookmarkManager.getBookmarkByName('Section1');
if (!bookmark) {
  console.error('Bookmark not found');
  return;
}
```

### Issue: Memory Leak

**Problem:** Managers not disposed properly
**Solution:** Always call dispose on document
**Fix:**
```typescript
try {
  const doc = Document.create();
  // ... use document ...
} finally {
  doc.dispose();  // Ensures cleanup
}
```

## See Also

- `src/core/CLAUDE.md` - Core document classes
- `src/elements/CLAUDE.md` - Document elements
- `src/formatting/CLAUDE.md` - Style and numbering systems
