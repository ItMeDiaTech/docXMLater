# TOC Pre-Population: Action Plan

## Objective

Create TOC that displays actual heading entries when document is first opened in Word, while maintaining field structure for "Update Field" functionality.

## Current State

✅ [`replaceTableOfContents(filePath)`](src/core/Document.ts:5108-5184) already implements pre-population
❌ Requires two-step process: save, then replace, then save again
❌ Not integrated into main save flow

## Solution Overview

### Integrate existing population logic into save() process

```typescript
// Before (current):
await doc.save("output.docx");
await doc.replaceTableOfContents("output.docx"); // Separate step!

// After (proposed):
doc.setAutoPopulateTOCs(true);
await doc.save("output.docx"); // TOC populated automatically!
```

## Implementation Tasks

### Task 1: Add Auto-Population Flag

**File**: [`src/core/Document.ts`](src/core/Document.ts)
**Location**: Class properties (around line 166)

```typescript
private autoPopulateTOCs: boolean = false;

public setAutoPopulateTOCs(enabled: boolean): this {
  this.autoPopulateTOCs = enabled;
  return this;
}

public isAutoPopulateTOCsEnabled(): boolean {
  return this.autoPopulateTOCs;
}
```

### Task 2: Integrate into Save Process

**File**: [`src/core/Document.ts`](src/core/Document.ts)
**Location**: Modify `save()` method (line 858)

```typescript
async save(filePath: string): Promise<void> {
  const tempPath = `${filePath}.tmp.${Date.now()}`;

  try {
    // ... existing validation and processing ...

    // Save to temp file FIRST
    await this.zipHandler.save(tempPath);

    // NEW: Auto-populate TOCs if enabled
    if (this.autoPopulateTOCs) {
      const handler = new ZipHandler();
      await handler.load(tempPath);

      const docXml = handler.getFileAsString('word/document.xml');
      if (docXml) {
        const populatedXml = this.populateAllTOCsInXML(docXml);
        handler.updateFile('word/document.xml', populatedXml);
        await handler.save(tempPath);
      }
    }

    // Atomic rename
    const { promises: fs } = await import('fs');
    await fs.rename(tempPath, filePath);
  } catch (error) {
    // ... error handling ...
  }
}
```

### Task 3: Extract Population Logic

**File**: [`src/core/Document.ts`](src/core/Document.ts)
**Location**: New private methods

```typescript
/**
 * Populates all TOCs in document XML
 * Extracted from replaceTableOfContents for reuse
 */
private populateAllTOCsInXML(docXml: string): string {
  const tocRegex = /<w:sdt>[\s\S]*?<w:docPartGallery w:val="Table of Contents"[\s\S]*?<\/w:sdt>/g;
  const tocMatches = Array.from(docXml.matchAll(tocRegex));

  if (tocMatches.length === 0) return docXml;

  let modifiedXml = docXml;

  for (const match of tocMatches) {
    const tocXml = match[0];

    // Extract and decode field instruction
    const instrMatch = tocXml.match(/<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/);
    if (!instrMatch?.[1]) continue;

    let fieldInstruction = instrMatch[1]
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");

    // Parse levels and find headings
    const levels = this.parseTOCFieldInstruction(fieldInstruction);
    const headings = this.findHeadingsForTOCFromXML(docXml, levels);

    if (headings.length === 0) continue;

    // Generate populated TOC
    const newTocXml = this.generateTOCXML(headings, fieldInstruction);
    modifiedXml = modifiedXml.replace(tocXml, newTocXml);
  }

  return modifiedXml;
}
```

### Task 4: Add Convenience Method

**File**: [`src/core/Document.ts`](src/core/Document.ts)
**Location**: After `createTableOfContents()` (around line 532)

```typescript
/**
 * Creates and adds a pre-populated Table of Contents
 * The TOC will display actual heading entries when document is opened
 * Field structure is preserved for "Update Field" functionality
 *
 * @param title - Optional TOC title
 * @param options - TOC options
 * @returns This document for chaining
 *
 * @example
 * const doc = Document.create();
 * doc.createParagraph('Chapter 1').setStyle('Heading1');
 * doc.createParagraph('Section 1.1').setStyle('Heading2');
 *
 * doc.createPrePopulatedTableOfContents('Contents');
 *
 * await doc.save('output.docx');
 * // TOC entries visible immediately when opened!
 */
public createPrePopulatedTableOfContents(
  title?: string,
  options?: Partial<TOCProperties>
): this {
  this.createTableOfContents(title);
  this.setAutoPopulateTOCs(true);
  return this;
}
```

### Task 5: Update toBuffer() Method

**File**: [`src/core/Document.ts`](src/core/Document.ts)
**Location**: `toBuffer()` method (line 929)

Apply same logic as save() for buffer generation.

## Implementation Order

1. ✅ **Complete**: Architectural analysis and documentation
2. **Next**: Add `autoPopulateTOCs` flag and setter (5 minutes)
3. **Next**: Extract `populateAllTOCsInXML()` method (10 minutes)
4. **Next**: Integrate into `save()` and `toBuffer()` (15 minutes)
5. **Next**: Add `createPrePopulatedTableOfContents()` (5 minutes)
6. **Next**: Create example file (15 minutes)
7. **Next**: Create tests (30 minutes)

**Total Estimated Time**: ~90 minutes

## Files to Modify

| File                                                | Modifications                                                  | Est. Lines |
| --------------------------------------------------- | -------------------------------------------------------------- | ---------- |
| [`src/core/Document.ts`](src/core/Document.ts)      | Add flag, integrate into save/toBuffer, add convenience method | +80        |
| `examples/08-table-of-contents/toc-prepopulated.ts` | New example file                                               | +150       |
| `src/__tests__/toc-prepopulation.test.ts`           | New test file                                                  | +200       |
| [`src/index.ts`](src/index.ts)                      | Export new types (if any)                                      | +5         |

## Key Design Decisions

### ✅ Decision 1: Opt-In Auto-Population

Auto-population is **disabled by default** to maintain backward compatibility.

```typescript
// Explicit opt-in required
doc.setAutoPopulateTOCs(true);
// OR
doc.createPrePopulatedTableOfContents();
```

### ✅ Decision 2: Population Happens During Save

Population integrated into save flow, not during `toXML()` generation.

**Rationale**:

- Headings may be added after TOC creation
- Avoids circular dependencies
- Reuses existing robust implementation

### ✅ Decision 3: Preserve Field Structure

Generated XML maintains complete field structure:

- Field BEGIN marker
- Field INSTRUCTION with switches
- Field SEPARATOR
- Populated entries (instead of placeholder)
- Field END marker

**Result**: Users can still right-click "Update Field" in Word!

## Success Criteria

### Must Have ✅

- [ ] TOC shows actual entries when document first opened
- [ ] Field structure preserved (begin/instruction/separate/content/end)
- [ ] "Update Field" still works in Word
- [ ] Backward compatible (auto-populate is opt-in)
- [ ] Works with all TOC types (standard, detailed, hyperlinked, etc.)

### Should Have

- [ ] Convenience method `createPrePopulatedTableOfContents()`
- [ ] Works in both `save()` and `toBuffer()` flows
- [ ] Comprehensive examples
- [ ] Full test coverage

## Example Usage

```typescript
import { Document } from "docxmlater";

// Create document
const doc = Document.create();

// Add content with headings
doc.createParagraph("Introduction").setStyle("Heading1");
doc.createParagraph("Background").setStyle("Heading2");
doc.createParagraph("Methodology").setStyle("Heading1");
doc.createParagraph("Data Collection").setStyle("Heading2");
doc.createParagraph("Results").setStyle("Heading1");

// Create pre-populated TOC
doc.createPrePopulatedTableOfContents("Table of Contents", {
  levels: 3,
  useHyperlinks: true,
});

// Save once - TOC is populated!
await doc.save("output.docx");

// When opened in Word:
// - TOC shows: Introduction, Background, Methodology, Data Collection, Results
// - Each entry is clickable (hyperlinks to heading)
// - Users can still right-click "Update Field" if they add more headings
```

## Ready for Implementation

All architectural decisions are complete. The plan leverages existing code (`replaceTableOfContents`) and integrates it cleanly into the save flow.

**Next Step**: Switch to Code mode to implement these changes.
