# Phase 4.6 Implementation Plan - Field Types

**Date:** October 23, 2025
**Estimated Time:** 2-3 hours
**Target:** Complete field implementation with parsing, integration, and testing

---

## Current Status Analysis

### Already Implemented ✅ (Field.ts - 708 lines)

**Field Class (Simple Fields):**
- 15 field types defined
- Static factory methods for common fields
- Run formatting support
- XML generation (fldSimple)
- Placeholder text generation

**Field Types Supported:**
1. PAGE - Current page number
2. NUMPAGES - Total pages
3. DATE - Current date
4. TIME - Current time
5. AUTHOR - Document author
6. TITLE - Document title
7. FILENAME - Document filename
8. FILENAMEWITHPATH - Filename with path
9. SUBJECT - Document subject
10. KEYWORDS - Document keywords
11. CREATEDATE - Creation date
12. SAVEDATE - Last save date
13. PRINTDATE - Last print date
14. SECTIONPAGES - Pages in section
15. SECTION - Section number

**ComplexField Class:**
- Begin/separate/end structure
- TOC field creation
- Custom instruction support
- Formatting for instruction and result

### Missing Components ❌

1. **Paragraph Integration** - No addField() method
2. **Field Parsing** - No parsing in DocumentParser
3. **Tests** - Zero tests for fields
4. **Round-trip Support** - Can generate but not load

---

## Implementation Tasks

### Task 1: Add Field Support to Paragraph
**File:** `src/elements/Paragraph.ts`
**Estimated Time:** 15 minutes

Add methods to support fields in paragraphs:

```typescript
/**
 * Adds a field to the paragraph
 * @param field Field or ComplexField to add
 */
addField(field: Field | ComplexField): this {
  // For simple fields (Field class)
  if (field instanceof Field) {
    const fieldXml = field.toXML();
    // Convert to Run-like structure
    const run = new Run(''); // Empty text
    run.setFieldXML(fieldXml);
    this.runs.push(run);
  }
  // For complex fields (ComplexField class)
  else if (field instanceof ComplexField) {
    // Complex fields generate multiple runs
    const runXmls = field.toXML();
    // Store as special marker
    this.complexFields = this.complexFields || [];
    this.complexFields.push(field);
  }
  return this;
}

/**
 * Adds a page number field
 * @param formatting Optional formatting
 */
addPageNumber(formatting?: RunFormatting): this {
  return this.addField(Field.createPageNumber(formatting));
}

/**
 * Adds a date field
 * @param format Date format
 * @param formatting Optional formatting
 */
addDate(format?: string, formatting?: RunFormatting): this {
  return this.addField(Field.createDate(format, formatting));
}
```

**Challenge:** Fields need special handling in toXML(). Need to decide:
- Option A: Store fields separately and inject during XML generation
- Option B: Convert fields to special Run objects
- **Decision: Option B** - Simpler, consistent with existing architecture

---

### Task 2: Add Field Parsing to DocumentParser
**File:** `src/core/DocumentParser.ts`
**Estimated Time:** 45 minutes

Add field parsing methods:

```typescript
/**
 * Parses simple field (fldSimple)
 * @private
 */
private parseSimpleField(fldSimple: any): Field | null {
  const instruction = fldSimple["@_w:instr"];
  if (!instruction) return null;

  // Extract field type from instruction
  const type = this.extractFieldType(instruction);

  // Parse formatting from run properties
  const rPr = fldSimple["w:rPr"];
  const formatting = rPr ? this.parseRunFormatting(rPr) : undefined;

  return Field.create({
    type,
    instruction,
    formatting,
  });
}

/**
 * Parses complex field (begin/separate/end structure)
 * @private
 */
private parseComplexField(runs: any[]): ComplexField | null {
  // Look for fldChar with type="begin"
  // Extract instrText
  // Look for fldChar with type="separate"
  // Extract result text
  // Look for fldChar with type="end"

  // Return ComplexField
}

/**
 * Extracts field type from instruction
 * @private
 */
private extractFieldType(instruction: string): FieldType {
  const match = instruction.trim().match(/^(\w+)/);
  return (match?.[1] || 'PAGE') as FieldType;
}
```

---

### Task 3: Integrate with Paragraph Parsing
**File:** `src/core/DocumentParser.ts`
**Estimated Time:** 30 minutes

Update `parseParagraphFromObject` to handle fields:

```typescript
// In paragraph parsing loop
for (const child of children) {
  if (child.name === 'w:fldSimple') {
    // Parse simple field
    const field = this.parseSimpleField(child);
    if (field) {
      paragraph.addField(field);
    }
  }
  else if (child.name === 'w:r') {
    // Check for complex field markers
    const fldChar = child.children?.find((c: any) => c.name === 'w:fldChar');
    if (fldChar) {
      // Handle complex field (requires state tracking)
      // ...
    } else {
      // Normal run
      // ...
    }
  }
}
```

---

### Task 4: Update Paragraph Class for Field XML Generation
**File:** `src/elements/Paragraph.ts`
**Estimated Time:** 20 minutes

Modify `toXML()` to generate field XML:

```typescript
toXML(): XMLElement {
  const children: XMLElement[] = [];

  // Add paragraph properties
  if (/* has properties */) {
    children.push(this.generatePPr());
  }

  // Add runs (including fields)
  for (const run of this.runs) {
    if (run.isField()) {
      // Field run - use field's XML
      children.push(run.getFieldXML());
    } else {
      // Normal run
      children.push(run.toXML());
    }
  }

  // Add complex fields
  if (this.complexFields) {
    for (const field of this.complexFields) {
      const fieldRuns = field.toXML();
      children.push(...fieldRuns);
    }
  }

  return XMLBuilder.w('p', undefined, children);
}
```

---

### Task 5: Create Comprehensive Tests
**File:** `tests/elements/Field.test.ts`
**Estimated Time:** 60 minutes

Test structure:
```typescript
describe('Field Tests', () => {
  describe('Simple Fields', () => {
    it('should create and serialize PAGE field', async () => {
      const doc = Document.create();
      const para = doc.addParagraph();
      para.addPageNumber();

      const buffer = await doc.toBuffer();
      const doc2 = await Document.loadFromBuffer(buffer);

      // Verify field exists and has correct instruction
    });

    it('should create DATE field with format', async () => { /* ... */ });
    it('should create NUMPAGES field', async () => { /* ... */ });
    it('should create AUTHOR field', async () => { /* ... */ });
    it('should create FILENAME field', async () => { /* ... */ });
  });

  describe('Complex Fields', () => {
    it('should create TOC field', async () => { /* ... */ });
    it('should handle TOC with custom options', async () => { /* ... */ });
  });

  describe('Field Formatting', () => {
    it('should apply formatting to field result', async () => { /* ... */ });
    it('should preserve MERGEFORMAT switch', async () => { /* ... */ });
  });

  describe('Round-Trip', () => {
    it('should preserve all field types through save/load', async () => { /* ... */ });
    it('should preserve field formatting', async () => { /* ... */ });
  });
});
```

**Expected Tests:** 20-25 tests

---

## Simplified Approach (Recommended)

Given complexity of full field integration, I recommend a **phased approach**:

### Phase A: Basic Field Support (1 hour)
1. Add Paragraph.addField() for simple fields only
2. Generate XML for fields in paragraph
3. Basic tests (10 tests) for field generation
4. **Skip parsing for now** - focus on generation

### Phase B: Field Parsing (1 hour - Future)
1. Add simple field parsing
2. Add complex field parsing
3. Round-trip tests

### Phase C: Advanced Features (1 hour - Future)
1. Field updates
2. Field switches
3. Complex field variations

**Recommendation:** Implement Phase A now (simpler, faster, still valuable)

---

## Success Criteria

**Minimum (Phase A):**
- ✅ Paragraph.addField() working
- ✅ Can generate PAGE, DATE, NUMPAGES fields
- ✅ 10+ tests passing
- ✅ Generated DOCX files open in Word
- ✅ Fields update correctly in Word

**Full (All Phases):**
- ✅ All 15 field types working
- ✅ Complex fields (TOC) working
- ✅ Field parsing implemented
- ✅ 20-25 tests passing
- ✅ Full round-trip support

---

## Decision Point

**Which approach?**

1. **Full Implementation** (2-3 hours)
   - Everything (generation + parsing + tests)
   - 20-25 tests
   - Complete round-trip

2. **Simplified (Phase A)** (1 hour) - **RECOMMENDED**
   - Field generation only
   - 10-15 tests
   - Parsing in future phase

---

**Proceeding with:** Simplified Phase A (generation + tests, skip parsing)
**Reason:** Faster, still valuable, can add parsing later
**Time:** ~1 hour
