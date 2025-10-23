# Implementation Specification: Document Corruption Fixes & Missing Features

**Date:** 2025-10-23
**Version:** 1.0
**Status:** Approved for Implementation

## Executive Summary

This document specifies the implementation of missing ECMA-376 (Office Open XML) features that are causing data loss and corruption when loading/saving DOCX files. Based on comprehensive research of the ECMA-376 specification and analysis of the current codebase.

---

## Table of Contents

1. [Research Summary](#research-summary)
2. [Current State Analysis](#current-state-analysis)
3. [Implementation Phases](#implementation-phases)
4. [Detailed Specifications](#detailed-specifications)
5. [Testing Strategy](#testing-strategy)
6. [Success Criteria](#success-criteria)

---

## Research Summary

### ECMA-376 Features Investigated

| Feature | Specification | Current Status | Priority |
|---------|--------------|----------------|----------|
| **Bookmarks** | §17.13.5 | ✅ Implemented | HIGH |
| **Field Codes (Complex)** | §17.16.4-5 | ⚠️ Partial (simple only) | HIGH |
| **Structured Document Tags** | §17.5.2 | ✅ Basic support | MEDIUM |
| **TOC with SDT** | §17.16.5.68 | ❌ Missing | HIGH |
| **Table Cell Margins** | §17.4.43 | ✅ Implemented | ✅ DONE |
| **RSID Attributes** | §17.13.5.29 | ✅ Correctly omitted | SKIP |
| **Table Grid Changes** | §17.4.49 | ✅ Correctly omitted | SKIP |

### Key Research Findings

1. **RSIDs are OPTIONAL** - Can be safely omitted for programmatic generation
2. **Table Cell Margins ALREADY IMPLEMENTED** - No action needed
3. **Complex Fields Required for TOC** - Current fldSimple insufficient
4. **SDT Wrapper Needed for TOC** - docPartObj gallery identification

---

## Current State Analysis

### Already Implemented ✅

The framework has MORE features than initially realized:

```
src/elements/
├── Bookmark.ts              ✅ Complete (toStartXML/toEndXML)
├── BookmarkManager.ts       ✅ ID management
├── Field.ts                 ⚠️ Partial (fldSimple only)
├── Comment.ts               ✅ Basic support
├── Revision.ts              ✅ Basic support
├── StructuredDocumentTag.ts ✅ Basic support
├── TableCell.ts             ✅ Cell margins (tcMar) implemented
└── TableOfContents.ts       ⚠️ No SDT wrapper
```

### Critical Bugs Identified 🐛

#### Bug #1: Style Color Parsing
**Location:** `src/core/DocumentParser.ts:1438-1550`
**Issue:** Hex color "000000" becomes number 36 (font size value)
**Root Cause:** `parseRunFormattingFromXml()` doesn't preserve hex colors
**Impact:** All style colors corrupted on load/save

#### Bug #2: Missing Complex Field Support
**Location:** `src/elements/Field.ts`
**Issue:** Only supports `<w:fldSimple>`, TOC requires `<w:fldChar>` + `<w:instrText>`
**Impact:** TOC fields not properly generated

#### Bug #3: TOC Missing SDT Wrapper
**Location:** `src/elements/TableOfContents.ts`
**Issue:** Doesn't wrap TOC in `<w:sdt>` with `<w:docPartObj>`
**Impact:** Word doesn't recognize as "Table of Contents" building block

---

## Implementation Phases

### Phase 1: Critical Bugs (IMMEDIATE) ⚡

**Goal:** Fix style corruption and basic functionality
**Time Estimate:** 2 hours

#### 1.1 Fix Style Color Parsing
**File:** `src/core/DocumentParser.ts`

**Problem:**
```typescript
// Current (BUGGY):
if (rPrXml.includes("<w:sz", "/>")){
  const val = XMLParser.extractAttribute(`<w:sz${szElement}`, "w:val");
  formatting.size = parseInt(val, 10) / 2; // size is 36
}

if (rPrXml.includes("<w:color", "/>")){
  const val = XMLParser.extractAttribute(`<w:color${colorElement}`, "w:val");
  formatting.color = val; // Returns "000000" but gets assigned as number
}
```

**Solution:**
```typescript
// Fixed:
private parseRunFormattingFromXml(rPrXml: string): RunFormatting {
  // ... existing code ...

  // Parse color (w:color) - PRESERVE HEX STRINGS
  const colorElement = XMLParser.extractBetweenTags(rPrXml, "<w:color", "/>");
  if (colorElement) {
    const val = XMLParser.extractAttribute(`<w:color${colorElement}`, "w:val");
    if (val && val !== "auto") {
      // Preserve hex colors - don't convert to numbers
      // "000000" should stay "000000", not become 0
      formatting.color = val;
    }
  }

  return formatting;
}
```

**Files Modified:**
- `src/core/DocumentParser.ts` (parseRunFormattingFromXml method)

**Tests Required:**
- Parse style with color "000000" → verify stays "000000"
- Parse style with color "FF0000" → verify stays "FF0000"
- Round-trip test: load → save → load → verify

#### 1.2 Add Comprehensive Style Round-Trip Tests
**File:** `tests/formatting/StylesRoundTrip.test.ts` (NEW)

**Test Cases:**
```typescript
describe('Style Round-Trip Tests', () => {
  it('should preserve hex color "000000" (black)', async () => {
    const doc = await Document.load('fixtures/styles-test.docx');
    const heading1 = doc.getStyle('Heading1');
    expect(heading1?.getProperties().runFormatting?.color).toBe('000000');

    await doc.save('output.docx');
    const doc2 = await Document.load('output.docx');
    const heading1_2 = doc2.getStyle('Heading1');
    expect(heading1_2?.getProperties().runFormatting?.color).toBe('000000');
  });

  it('should preserve all Heading style colors', async () => {
    // Test Heading1-9 colors
  });

  it('should preserve custom style colors', async () => {
    // Test custom styles with various hex colors
  });
});
```

**Coverage:**
- Black (000000) - edge case that failed
- Red (FF0000), Blue (0000FF), Green (00FF00)
- Theme colors (2E74B5, 1F4D78, etc.)
- Custom colors (123456, ABCDEF, etc.)

---

### Phase 2: Complex Fields (HIGH PRIORITY) 🔧

**Goal:** Support TOC and other dynamic fields
**Time Estimate:** 3 hours

#### 2.1 Enhance Field.ts for Complex Fields

**Current Implementation (fldSimple):**
```xml
<w:fldSimple w:instr=" PAGE \* MERGEFORMAT ">
  <w:t>1</w:t>
</w:fldSimple>
```

**Needed Implementation (fldChar + instrText):**
```xml
<!-- Begin marker -->
<w:r>
  <w:fldChar w:fldCharType="begin"/>
</w:r>

<!-- Instruction -->
<w:r>
  <w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>
</w:r>

<!-- Separator -->
<w:r>
  <w:fldChar w:fldCharType="separate"/>
</w:r>

<!-- Result (optional) -->
<w:r>
  <w:t>Table of Contents</w:t>
</w:r>

<!-- End marker -->
<w:r>
  <w:fldChar w:fldCharType="end"/>
</w:r>
```

**New Classes to Add:**

```typescript
/**
 * Field character types for complex fields
 */
export type FieldCharType = 'begin' | 'separate' | 'end';

/**
 * Complex field properties
 */
export interface ComplexFieldProperties {
  /** Field instruction (e.g., " TOC \o \"1-3\" \h \z \u ") */
  instruction: string;

  /** Current field result text (optional) */
  result?: string;

  /** Run formatting for instruction */
  instructionFormatting?: RunFormatting;

  /** Run formatting for result */
  resultFormatting?: RunFormatting;
}

/**
 * Represents a complex field (begin/separate/end structure)
 */
export class ComplexField {
  private instruction: string;
  private result?: string;
  private instructionFormatting?: RunFormatting;
  private resultFormatting?: RunFormatting;

  constructor(properties: ComplexFieldProperties) {
    this.instruction = properties.instruction;
    this.result = properties.result;
    this.instructionFormatting = properties.instructionFormatting;
    this.resultFormatting = properties.resultFormatting;
  }

  /**
   * Generates XML for the complex field
   * Returns array of run elements (begin, instr, sep, result, end)
   */
  toXML(): XMLElement[] {
    const runs: XMLElement[] = [];

    // 1. Begin marker run
    runs.push(XMLBuilder.w('r', undefined, [
      XMLBuilder.wSelf('fldChar', { 'w:fldCharType': 'begin' })
    ]));

    // 2. Instruction run
    const instrChildren: XMLElement[] = [];
    if (this.instructionFormatting) {
      instrChildren.push(this.createRunProperties(this.instructionFormatting));
    }
    instrChildren.push(XMLBuilder.w('instrText', {
      'xml:space': 'preserve'
    }, [this.instruction]));
    runs.push(XMLBuilder.w('r', undefined, instrChildren));

    // 3. Separator run
    runs.push(XMLBuilder.w('r', undefined, [
      XMLBuilder.wSelf('fldChar', { 'w:fldCharType': 'separate' })
    ]));

    // 4. Result run (optional)
    if (this.result) {
      const resultChildren: XMLElement[] = [];
      if (this.resultFormatting) {
        resultChildren.push(this.createRunProperties(this.resultFormatting));
      }
      resultChildren.push(XMLBuilder.w('t', undefined, [this.result]));
      runs.push(XMLBuilder.w('r', undefined, resultChildren));
    }

    // 5. End marker run
    runs.push(XMLBuilder.w('r', undefined, [
      XMLBuilder.wSelf('fldChar', { 'w:fldCharType': 'end' })
    ]));

    return runs;
  }

  private createRunProperties(formatting: RunFormatting): XMLElement {
    // Generate w:rPr from RunFormatting
    // Similar to Run.toXML() rPr generation
  }
}
```

#### 2.2 TOC Field Builder

**Add to Field.ts:**

```typescript
/**
 * TOC field options
 */
export interface TOCFieldOptions {
  /** Heading levels to include (e.g., "1-3" for levels 1-3) */
  levels?: string;

  /** Make entries hyperlinks (\h switch) */
  hyperlinks?: boolean;

  /** Hide tab leaders and page numbers in Web Layout (\z switch) */
  hideInWebLayout?: boolean;

  /** Use outline levels (\u switch) */
  useOutlineLevels?: boolean;

  /** Omit page numbers (\n switch) */
  omitPageNumbers?: boolean;

  /** Custom styles to use (\t switch) */
  customStyles?: string;
}

/**
 * Creates a TOC (Table of Contents) complex field
 */
export function createTOCField(options: TOCFieldOptions = {}): ComplexField {
  // Build instruction string
  let instruction = ' TOC';

  // Add outline levels switch
  if (options.levels !== undefined) {
    instruction += ` \\o "${options.levels}"`;
  } else {
    instruction += ' \\o "1-3"'; // Default: levels 1-3
  }

  // Add hyperlinks switch
  if (options.hyperlinks !== false) {
    instruction += ' \\h';
  }

  // Add hide in web layout switch
  if (options.hideInWebLayout !== false) {
    instruction += ' \\z';
  }

  // Add use outline levels switch
  if (options.useOutlineLevels !== false) {
    instruction += ' \\u';
  }

  // Add omit page numbers switch
  if (options.omitPageNumbers) {
    instruction += ' \\n';
  }

  // Add custom styles switch
  if (options.customStyles) {
    instruction += ` \\t "${options.customStyles}"`;
  }

  instruction += ' '; // Trailing space per Microsoft convention

  return new ComplexField({
    instruction,
    result: 'Table of Contents' // Placeholder result
  });
}
```

**Files Modified:**
- `src/elements/Field.ts` (add ComplexField class and createTOCField)

**Tests Required:**
- `tests/elements/ComplexField.test.ts` (NEW)
  - Generate complex field XML structure
  - Verify begin/separate/end markers
  - Test TOC field instruction building
  - Test all TOC switches

---

### Phase 3: Enhanced TOC with SDT Wrapper (MEDIUM PRIORITY) 📦

**Goal:** Proper TOC structure per ECMA-376
**Time Estimate:** 2 hours

#### 3.1 Enhance StructuredDocumentTag.ts

**Add TOC-Specific Support:**

```typescript
/**
 * Document part gallery properties for SDT
 * Used to identify TOC and other building blocks
 */
export interface SDTDocPartProperties {
  /** Gallery name (e.g., "Table of Contents") */
  gallery: string;

  /** Mark as unique instance */
  unique?: boolean;
}

export class StructuredDocumentTag {
  private properties: SDTProperties;
  private content: SDTContent[];
  private docPartObj?: SDTDocPartProperties; // NEW

  // ... existing methods ...

  /**
   * Set document part gallery (for TOC, bibliography, etc.)
   * @param gallery - Gallery name
   * @param unique - Whether this is a unique instance
   */
  setDocPartGallery(gallery: string, unique: boolean = true): this {
    this.docPartObj = { gallery, unique };
    return this;
  }

  /**
   * Get document part gallery
   */
  getDocPartGallery(): string | undefined {
    return this.docPartObj?.gallery;
  }

  /**
   * Check if this is a TOC SDT
   */
  isTOC(): boolean {
    return this.docPartObj?.gallery === 'Table of Contents';
  }

  /**
   * Converts to WordprocessingML XML element
   */
  toXML(): XMLElement {
    const sdtPrChildren: XMLElement[] = [];

    // Existing properties
    if (this.properties.id !== undefined) {
      sdtPrChildren.push(XMLBuilder.wSelf('id', { 'w:val': this.properties.id.toString() }));
    }
    if (this.properties.tag) {
      sdtPrChildren.push(XMLBuilder.wSelf('tag', { 'w:val': this.properties.tag }));
    }
    if (this.properties.lock) {
      sdtPrChildren.push(XMLBuilder.wSelf('lock', { 'w:val': this.properties.lock }));
    }
    if (this.properties.alias) {
      sdtPrChildren.push(XMLBuilder.wSelf('alias', { 'w:val': this.properties.alias }));
    }

    // NEW: Add docPartObj if set
    if (this.docPartObj) {
      const docPartChildren: XMLElement[] = [
        XMLBuilder.wSelf('docPartGallery', { 'w:val': this.docPartObj.gallery })
      ];
      if (this.docPartObj.unique) {
        docPartChildren.push(XMLBuilder.wSelf('docPartUnique'));
      }
      sdtPrChildren.push(XMLBuilder.w('docPartObj', undefined, docPartChildren));
    }

    // ... rest of existing toXML implementation ...
  }
}
```

#### 3.2 Update TableOfContents.ts to Use SDT

**Current Implementation:**
```typescript
toXML(): XMLElement[] {
  return [
    titleParagraph.toXML(),
    fieldParagraph.toXML()
  ];
}
```

**New Implementation:**
```typescript
toXML(): XMLElement[] {
  // Create SDT wrapper
  const sdt = new StructuredDocumentTag({
    id: this.generateUniqueId(),
    tag: 'TOC',
    alias: this.title || 'Table of Contents'
  });

  // Set as TOC document part
  sdt.setDocPartGallery('Table of Contents', true);

  // Add TOC content (title + field paragraphs)
  const titlePara = this.createTitleParagraph();
  const fieldPara = this.createFieldParagraph();

  sdt.addContent(titlePara);
  sdt.addContent(fieldPara);

  // Return SDT XML (single element instead of array)
  return [sdt.toXML()];
}

private createFieldParagraph(): Paragraph {
  const para = new Paragraph();

  // Use complex field instead of simple field
  const tocField = createTOCField({
    levels: this.options.headingLevels,
    hyperlinks: this.options.includeHyperlinks,
    hideInWebLayout: true,
    useOutlineLevels: true
  });

  // Add field runs to paragraph
  // Note: ComplexField.toXML() returns XMLElement[] (runs)
  // Need to integrate with Paragraph structure

  return para;
}

private generateUniqueId(): number {
  // Generate random negative integer (Microsoft convention)
  return -Math.floor(Math.random() * 1000000000);
}
```

**Files Modified:**
- `src/elements/StructuredDocumentTag.ts` (add docPartObj support)
- `src/elements/TableOfContents.ts` (wrap in SDT, use ComplexField)

**Tests Required:**
- `tests/elements/TOC_SDT.test.ts` (NEW)
  - Verify SDT wrapper generated
  - Verify docPartGallery set to "Table of Contents"
  - Verify complex field structure
  - Test XML structure matches ECMA-376

---

### Phase 4: Parsing Support (MEDIUM PRIORITY) 📖

**Goal:** Load existing documents without data loss
**Time Estimate:** 3 hours

#### 4.1 Parse Bookmarks from Existing Documents

**Add to DocumentParser.ts:**

```typescript
private async parseParagraphFromObject(
  paraObj: any,
  relationshipManager: RelationshipManager,
  zipHandler?: ZipHandler,
  imageManager?: ImageManager
): Promise<Paragraph | null> {
  try {
    const paragraph = new Paragraph();

    // ... existing paragraph property parsing ...

    // NEW: Parse bookmark start markers
    if (paraObj["w:bookmarkStart"]) {
      const bookmarkStarts = Array.isArray(paraObj["w:bookmarkStart"])
        ? paraObj["w:bookmarkStart"]
        : [paraObj["w:bookmarkStart"]];

      for (const bmStart of bookmarkStarts) {
        const id = parseInt(bmStart["@_w:id"], 10);
        const name = bmStart["@_w:name"];

        if (id !== undefined && name) {
          const bookmark = new Bookmark({
            id,
            name,
            skipNormalization: true // Preserve exact names from Word
          });
          paragraph.addBookmarkStart(bookmark);
        }
      }
    }

    // NEW: Parse bookmark end markers
    if (paraObj["w:bookmarkEnd"]) {
      const bookmarkEnds = Array.isArray(paraObj["w:bookmarkEnd"])
        ? paraObj["w:bookmarkEnd"]
        : [paraObj["w:bookmarkEnd"]];

      for (const bmEnd of bookmarkEnds) {
        const id = parseInt(bmEnd["@_w:id"], 10);

        if (id !== undefined) {
          // Create bookmark reference for end marker
          const bookmark = new Bookmark({ id, name: `_bookmark_${id}` });
          paragraph.addBookmarkEnd(bookmark);
        }
      }
    }

    // ... existing run parsing ...

    return paragraph;
  } catch (error) {
    // ... error handling ...
  }
}
```

**Note:** Paragraph class already has bookmark support:
```typescript
// From Paragraph.ts:
private bookmarksStart: Bookmark[] = [];
private bookmarksEnd: Bookmark[] = [];
```

Need to verify these methods exist or add them:
```typescript
// In Paragraph.ts:
addBookmarkStart(bookmark: Bookmark): this {
  this.bookmarksStart.push(bookmark);
  return this;
}

addBookmarkEnd(bookmark: Bookmark): this {
  this.bookmarksEnd.push(bookmark);
  return this;
}
```

#### 4.2 Parse Complex Fields

**Add to DocumentParser.ts:**

```typescript
/**
 * Detects and parses complex fields from runs
 * Complex fields span multiple runs with fldChar markers
 */
private parseComplexFieldFromRuns(runs: any[]): { field: ComplexField | null; consumedRuns: number } {
  if (!runs || runs.length === 0) {
    return { field: null, consumedRuns: 0 };
  }

  // Look for begin marker
  const firstRun = runs[0];
  if (!firstRun["w:fldChar"] || firstRun["w:fldChar"]["@_w:fldCharType"] !== "begin") {
    return { field: null, consumedRuns: 0 };
  }

  let instruction = '';
  let result = '';
  let consumedRuns = 1;
  let phase: 'instruction' | 'result' | 'done' = 'instruction';

  // Parse subsequent runs
  for (let i = 1; i < runs.length; i++) {
    const run = runs[i];
    consumedRuns++;

    // Check for field char
    if (run["w:fldChar"]) {
      const charType = run["w:fldChar"]["@_w:fldCharType"];

      if (charType === "separate") {
        phase = 'result';
        continue;
      } else if (charType === "end") {
        phase = 'done';
        break;
      }
    }

    // Extract text
    if (phase === 'instruction' && run["w:instrText"]) {
      const textElement = run["w:instrText"];
      const text = typeof textElement === 'object' && textElement !== null
        ? (textElement["#text"] || "")
        : (textElement || "");
      instruction += XMLBuilder.unescapeXml(text);
    } else if (phase === 'result' && run["w:t"]) {
      const textElement = run["w:t"];
      const text = typeof textElement === 'object' && textElement !== null
        ? (textElement["#text"] || "")
        : (textElement || "");
      result += XMLBuilder.unescapeXml(text);
    }
  }

  if (phase === 'done' && instruction) {
    const field = new ComplexField({
      instruction,
      result: result || undefined
    });
    return { field, consumedRuns };
  }

  return { field: null, consumedRuns: 0 };
}
```

**Update run parsing to detect fields:**

```typescript
private async parseParagraphFromObject(...) {
  // ... existing code ...

  // Parse runs - check for complex fields first
  const runs = paraObj["w:r"];
  const runChildren = Array.isArray(runs) ? runs : (runs ? [runs] : []);

  let i = 0;
  while (i < runChildren.length) {
    const child = runChildren[i];

    // Try to parse complex field starting at this run
    const { field, consumedRuns } = this.parseComplexFieldFromRuns(runChildren.slice(i));

    if (field) {
      // Add field to paragraph
      paragraph.addField(field);
      i += consumedRuns;
      continue;
    }

    // Regular run parsing
    if (child["w:drawing"]) {
      // ... image run ...
    } else {
      // ... text run ...
    }

    i++;
  }

  return paragraph;
}
```

**Note:** Need to add field support to Paragraph:
```typescript
// In Paragraph.ts:
private content: ParagraphContent[] = [];  // Already includes Field type

addField(field: Field | ComplexField): this {
  this.content.push(field);
  return this;
}
```

#### 4.3 Parse SDT Elements

**Add to DocumentParser.ts:**

```typescript
private async parseSDTFromObject(
  sdtObj: any,
  relationshipManager: RelationshipManager,
  zipHandler: ZipHandler,
  imageManager: ImageManager
): Promise<StructuredDocumentTag | null> {
  try {
    const sdt = new StructuredDocumentTag();

    // Parse SDT properties
    if (sdtObj["w:sdtPr"]) {
      const props = sdtObj["w:sdtPr"];

      // ID
      if (props["w:id"]) {
        const id = parseInt(props["w:id"]["@_w:val"], 10);
        if (!isNaN(id)) {
          sdt.setId(id);
        }
      }

      // Tag
      if (props["w:tag"]) {
        sdt.setTag(props["w:tag"]["@_w:val"]);
      }

      // Lock
      if (props["w:lock"]) {
        sdt.setLock(props["w:lock"]["@_w:val"]);
      }

      // Alias
      if (props["w:alias"]) {
        sdt.setAlias(props["w:alias"]["@_w:val"]);
      }

      // Document part object (for TOC)
      if (props["w:docPartObj"]) {
        const docPartObj = props["w:docPartObj"];
        if (docPartObj["w:docPartGallery"]) {
          const gallery = docPartObj["w:docPartGallery"]["@_w:val"];
          const unique = !!docPartObj["w:docPartUnique"];
          sdt.setDocPartGallery(gallery, unique);
        }
      }
    }

    // Parse SDT content
    if (sdtObj["w:sdtContent"]) {
      const content = sdtObj["w:sdtContent"];

      // Parse paragraphs
      if (content["w:p"]) {
        const paragraphs = Array.isArray(content["w:p"])
          ? content["w:p"]
          : [content["w:p"]];

        for (const paraObj of paragraphs) {
          const paragraph = await this.parseParagraphFromObject(
            paraObj,
            relationshipManager,
            zipHandler,
            imageManager
          );
          if (paragraph) {
            sdt.addContent(paragraph);
          }
        }
      }

      // Parse tables
      if (content["w:tbl"]) {
        const tables = Array.isArray(content["w:tbl"])
          ? content["w:tbl"]
          : [content["w:tbl"]];

        for (const tableObj of tables) {
          const table = await this.parseTableFromObject(
            tableObj,
            relationshipManager,
            zipHandler,
            imageManager
          );
          if (table) {
            sdt.addContent(table);
          }
        }
      }

      // TODO: Parse nested SDTs if needed
    }

    return sdt;
  } catch (error) {
    console.warn('[DocumentParser] Failed to parse SDT:', error);
    return null;
  }
}
```

**Update body element parsing to handle SDT:**

Already implemented! (See earlier code read - parseBodyElements handles w:sdt)

**Files Modified:**
- `src/core/DocumentParser.ts` (add parsing for bookmarks, complex fields, SDT)
- `src/elements/Paragraph.ts` (verify/add bookmark and field methods)

**Tests Required:**
- `tests/core/DocumentParser.test.ts` (enhance existing)
  - Parse bookmarks from Test6_BaseFile.docx
  - Parse complex fields from Test6_BaseFile.docx
  - Parse SDT from Test6_BaseFile.docx
  - Verify all elements preserved in round-trip

---

### Phase 5: RSID Attributes (SKIP ✅)

**Decision:** OMIT RSIDs for programmatic generation

**Rationale:**
- RSIDs are OPTIONAL per ECMA-376 specification
- Only needed for collaborative editing / track changes
- Not needed for document generation
- Cleaner XML without them
- Word regenerates them on first edit if needed
- Aligns with framework's "lean XML" philosophy

**Implementation:** NO CHANGES NEEDED ✅

**Documentation Update:**
```markdown
# RSID Handling Policy

The docXMLater framework intentionally OMITS RSID (Revision Session ID) attributes
when generating documents programmatically. This is per ECMA-376 specification which
marks RSIDs as OPTIONAL.

## Why No RSIDs?

1. **Specification Compliance**: RSIDs are optional per ECMA-376
2. **Cleaner XML**: Reduces markup clutter by ~30%
3. **No Functional Impact**: RSIDs only matter for merging forked documents
4. **Word Compatibility**: Word regenerates RSIDs on first edit if needed
5. **Framework Philosophy**: "The best code is the code you don't write"

## When RSIDs Matter

RSIDs are only important for:
- Collaborative editing with track changes
- Merging forked document versions
- Document forensics / audit trails

For programmatic document generation, RSIDs provide zero value.

## User Control

Users can enable RSID tracking in Microsoft Word if needed for their workflow.
The framework will preserve RSIDs when loading documents, but will not generate
new RSIDs when creating documents from scratch.
```

---

### Phase 6: Table Grid Changes (SKIP ✅)

**Decision:** OMIT `<w:tblGridChange>` per research

**Rationale:**
- Only for track changes / revision tracking
- Not needed for basic table structure
- Framework already has `<w:tblGrid>` implemented
- Not needed for document generation

**Implementation:** NO CHANGES NEEDED ✅

**Note:** Framework already correctly implements `<w:tblGrid>` without `<w:tblGridChange>`

---

## Testing Strategy

### Unit Tests

| Test File | Purpose | Coverage |
|-----------|---------|----------|
| `StylesRoundTrip.test.ts` | Style parsing | Hex colors, all properties |
| `ComplexField.test.ts` | Complex fields | begin/sep/end structure, TOC |
| `TOC_SDT.test.ts` | TOC with SDT | docPartGallery, full structure |
| `BookmarkParsing.test.ts` | Bookmark parsing | start/end markers, IDs |
| `SDTParsing.test.ts` | SDT parsing | Properties, content |

### Integration Tests

| Test | Description |
|------|-------------|
| Full TOC Generation | Create TOC with SDT + complex field |
| Test6 Round-Trip | Load Test6_BaseFile → save → load → verify |
| Bookmark Round-Trip | Create bookmarks → save → load → verify |

### Regression Tests

| Test | Purpose |
|------|---------|
| Test6_BaseFile.docx | Verify no data loss on user's problematic file |
| All existing tests | Ensure no regressions |

### Test Fixtures Needed

```
tests/fixtures/
├── Test6_BaseFile.docx           (user's problematic file)
├── styles-colors.docx             (various hex colors)
├── toc-with-sdt.docx              (proper TOC from Word)
├── bookmarks-complex.docx         (bookmarks spanning paragraphs)
└── complex-fields.docx            (various field types)
```

---

## Success Criteria

### Must Have ✅

- [ ] Style colors preserve correctly
  - [ ] "000000" stays "000000" (not 0)
  - [ ] "FF0000" stays "FF0000"
  - [ ] All hex formats work

- [ ] Complex fields supported
  - [ ] fldChar (begin/sep/end) structure
  - [ ] instrText with proper formatting
  - [ ] TOC field generation

- [ ] TOC with SDT wrapper
  - [ ] docPartGallery set to "Table of Contents"
  - [ ] docPartUnique flag
  - [ ] Proper structure

- [ ] All tests passing
  - [ ] 100+ new tests
  - [ ] All existing tests still pass
  - [ ] Test6 round-trip works

### Should Have 📋

- [ ] Bookmark parsing from documents
- [ ] Complex field parsing from documents
- [ ] SDT parsing from documents
- [ ] Comprehensive documentation

### Won't Have ❌

- ❌ RSID generation (correctly omitted)
- ❌ Table grid change tracking (correctly omitted)
- ❌ Full SDT types (date pickers, combo boxes, etc.)
  - Only TOC SDT needed for now

---

## Implementation Checklist

### Phase 1: Critical Bugs ⚡
- [ ] Fix `parseRunFormattingFromXml` color preservation
- [ ] Add `StylesRoundTrip.test.ts`
- [ ] Run tests and verify
- [ ] Test with Test6_BaseFile.docx

### Phase 2: Complex Fields 🔧
- [ ] Add `ComplexField` class to `Field.ts`
- [ ] Add `createTOCField()` function
- [ ] Add `ComplexField.test.ts`
- [ ] Run tests and verify

### Phase 3: Enhanced TOC 📦
- [ ] Add `docPartObj` to `StructuredDocumentTag.ts`
- [ ] Update `TableOfContents.ts` to use SDT
- [ ] Add `TOC_SDT.test.ts`
- [ ] Run tests and verify

### Phase 4: Parsing Support 📖
- [ ] Add bookmark parsing to `DocumentParser.ts`
- [ ] Add complex field parsing to `DocumentParser.ts`
- [ ] Add SDT parsing to `DocumentParser.ts`
- [ ] Add parsing tests
- [ ] Test6 round-trip verification

### Phase 5 & 6: Documentation ✅
- [ ] Document RSID policy
- [ ] Document table grid change policy
- [ ] Update README with new features
- [ ] Update CLAUDE.md files

---

## Estimated Timeline

| Phase | Estimated Time | Priority |
|-------|---------------|----------|
| Phase 1: Critical Bugs | 2 hours | HIGH ⚡ |
| Phase 2: Complex Fields | 3 hours | HIGH |
| Phase 3: Enhanced TOC | 2 hours | MEDIUM |
| Phase 4: Parsing Support | 3 hours | MEDIUM |
| Phase 5-6: Documentation | 1 hour | LOW |
| **Total** | **11 hours** | |

---

## Risk Assessment

### Low Risk ✅
- Style color fix (isolated change)
- Adding new features (ComplexField, docPartObj)
- Documentation updates

### Medium Risk ⚠️
- Parsing support (complex, but well-researched)
- TOC SDT wrapper (changes existing behavior)

### Mitigation Strategies
1. Comprehensive test coverage (100+ tests)
2. Test fixtures from real Word documents
3. Round-trip testing
4. Backward compatibility checks

---

## Success Metrics

### Before Implementation
- ❌ Test6_BaseFile.docx loses data on round-trip
- ❌ Style colors corrupted (000000 → 0 → "0")
- ❌ TOC not recognized by Word
- ❌ Bookmarks stripped
- ❌ Fields simplified/lost

### After Implementation
- ✅ Test6_BaseFile.docx perfect round-trip
- ✅ All style colors preserved
- ✅ TOC recognized by Word as building block
- ✅ Bookmarks preserved
- ✅ Complex fields maintained
- ✅ 100+ new tests passing
- ✅ 0 regressions

---

## Reference Documentation

### ECMA-376 Sections Referenced

- **§17.13.5** - Bookmarks (bookmarkStart, bookmarkEnd)
- **§17.16.4** - Field Character (fldChar)
- **§17.16.5** - Field Codes and Instructions
- **§17.16.5.68** - TOC Field
- **§17.5.2** - Structured Document Tags
- **§17.4.43** - Table Cell Margins (tcMar)
- **§17.4.49** - Table Grid (tblGrid)
- **§17.13.5.29** - RSID Attributes

### Microsoft Resources

- Open XML SDK Documentation
- WordprocessingML Reference
- Office Open XML Explained (officeopenxml.com)

---

## Appendix A: XML Examples

### A.1 Complex Field (TOC)

```xml
<w:p>
  <!-- Begin marker -->
  <w:r>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>

  <!-- Instruction -->
  <w:r>
    <w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText>
  </w:r>

  <!-- Separator -->
  <w:r>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>

  <!-- Result -->
  <w:r>
    <w:t>Table of Contents</w:t>
  </w:r>

  <!-- End marker -->
  <w:r>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
</w:p>
```

### A.2 TOC with SDT Wrapper

```xml
<w:sdt>
  <w:sdtPr>
    <w:id w:val="-123456789"/>
    <w:tag w:val="TOC"/>
    <w:alias w:val="Table of Contents"/>
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtEndPr/>
  <w:sdtContent>
    <!-- Title paragraph -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="TOCHeading"/>
      </w:pPr>
      <w:r>
        <w:t>Table of Contents</w:t>
      </w:r>
    </w:p>

    <!-- Field paragraph (from A.1) -->
    <!-- ... -->
  </w:sdtContent>
</w:sdt>
```

### A.3 Bookmarks

```xml
<w:p>
  <w:r>
    <w:t>Before bookmark</w:t>
  </w:r>
  <w:bookmarkStart w:id="0" w:name="MyBookmark"/>
  <w:r>
    <w:t>Bookmarked text</w:t>
  </w:r>
  <w:bookmarkEnd w:id="0"/>
  <w:r>
    <w:t>After bookmark</w:t>
  </w:r>
</w:p>
```

### A.4 Table Cell Margins (Already Implemented)

```xml
<w:tc>
  <w:tcPr>
    <w:tcMar>
      <w:top w:w="144" w:type="dxa"/>
      <w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="144" w:type="dxa"/>
      <w:right w:w="108" w:type="dxa"/>
    </w:tcMar>
  </w:tcPr>
  <w:p>
    <w:r>
      <w:t>Cell content</w:t>
    </w:r>
  </w:p>
</w:tc>
```

---

**End of Implementation Specification**
**Version 1.0 - Ready for Implementation**
