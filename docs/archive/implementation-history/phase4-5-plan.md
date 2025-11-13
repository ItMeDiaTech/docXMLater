# Phase 4.5 Implementation Plan - Section Properties

**Date:** October 23, 2025
**Estimated Time:** 2-3 hours (reduced from 3-4, most properties already implemented)
**Target:** Complete section properties with enhanced support

---

## Current Status Analysis

### Already Implemented ✅ (11 properties)
1. Page size (width, height) - `setPageSize()`
2. Page orientation (portrait/landscape) - `setOrientation()`
3. Page margins (top, right, bottom, left, gutter, header, footer) - `setMargins()`
4. Column count - `setColumns()`
5. Column spacing - `setColumns(space param)`
6. Column equal widths - `setColumns()` sets equalWidth: true
7. Section type (nextPage, continuous, etc.) - `setSectionType()`
8. Page numbering format - `setPageNumbering(format param)`
9. Page numbering start - `setPageNumbering(start param)`
10. Title page (different first page) - `setTitlePage()`
11. Header/Footer references (default, first, even) - `setHeaderReference()`, `setFooterReference()`

### Missing Properties ❌ (4 properties)
1. **Vertical page alignment** - top, center, bottom, justify
2. **Paper source** - printer tray selection
3. **Column separator line** - visual separator between columns
4. **Text direction** - LTR/RTL support

---

## Implementation Tasks

### Task 1: Add Missing Properties to Interface
**File:** `src/elements/Section.ts`
**Estimated Time:** 15 minutes

Add to SectionProperties interface:
```typescript
export interface SectionProperties {
  // ... existing properties ...

  /** Vertical page alignment */
  verticalAlignment?: 'top' | 'center' | 'bottom' | 'both';

  /** Paper source (printer tray) */
  paperSource?: {
    first?: number;  // First page tray
    other?: number;  // Other pages tray
  };

  /** Text direction */
  textDirection?: 'ltr' | 'rtl' | 'tbRl' | 'btLr';
}

export interface Columns {
  // ... existing properties ...

  /** Column separator line */
  separator?: boolean;

  /** Individual column widths (for unequal columns) */
  columnWidths?: number[];
}
```

---

### Task 2: Add Setter Methods
**File:** `src/elements/Section.ts`
**Estimated Time:** 20 minutes

```typescript
/**
 * Sets vertical page alignment
 * @param alignment Vertical alignment
 */
setVerticalAlignment(alignment: 'top' | 'center' | 'bottom' | 'both'): this {
  this.properties.verticalAlignment = alignment;
  return this;
}

/**
 * Sets paper source (printer tray)
 * @param first First page tray number
 * @param other Other pages tray number
 */
setPaperSource(first?: number, other?: number): this {
  this.properties.paperSource = { first, other };
  return this;
}

/**
 * Sets column separator line
 * @param separator Whether to show column separator
 */
setColumnSeparator(separator: boolean = true): this {
  if (!this.properties.columns) {
    this.properties.columns = { count: 1 };
  }
  this.properties.columns.separator = separator;
  return this;
}

/**
 * Sets text direction
 * @param direction Text direction
 */
setTextDirection(direction: 'ltr' | 'rtl' | 'tbRl' | 'btLr'): this {
  this.properties.textDirection = direction;
  return this;
}

/**
 * Sets custom column widths
 * @param widths Array of column widths in twips
 */
setColumnWidths(widths: number[]): this {
  if (!this.properties.columns) {
    this.properties.columns = { count: widths.length };
  }
  this.properties.columns.columnWidths = widths;
  this.properties.columns.equalWidth = false;
  return this;
}
```

---

### Task 3: Update XML Generation
**File:** `src/elements/Section.ts` (toXML method)
**Estimated Time:** 25 minutes

Add XML generation for new properties:

```typescript
// Vertical alignment
if (this.properties.verticalAlignment) {
  children.push(
    XMLBuilder.wSelf('vAlign', { 'w:val': this.properties.verticalAlignment })
  );
}

// Paper source
if (this.properties.paperSource) {
  const attrs: Record<string, string> = {};
  if (this.properties.paperSource.first !== undefined) {
    attrs['w:first'] = this.properties.paperSource.first.toString();
  }
  if (this.properties.paperSource.other !== undefined) {
    attrs['w:other'] = this.properties.paperSource.other.toString();
  }
  if (Object.keys(attrs).length > 0) {
    children.push(XMLBuilder.wSelf('paperSrc', attrs));
  }
}

// Text direction
if (this.properties.textDirection) {
  children.push(
    XMLBuilder.wSelf('textDirection', { 'w:val': this.properties.textDirection })
  );
}

// Enhanced columns with separator and custom widths
if (this.properties.columns) {
  const attrs: Record<string, string> = {
    'w:num': this.properties.columns.count.toString(),
  };
  if (this.properties.columns.space !== undefined) {
    attrs['w:space'] = this.properties.columns.space.toString();
  }
  if (this.properties.columns.equalWidth !== undefined) {
    attrs['w:equalWidth'] = this.properties.columns.equalWidth ? '1' : '0';
  }
  if (this.properties.columns.separator !== undefined) {
    attrs['w:sep'] = this.properties.columns.separator ? '1' : '0';
  }

  const colChildren: XMLElement[] = [];

  // Add individual column widths if specified
  if (this.properties.columns.columnWidths) {
    for (const width of this.properties.columns.columnWidths) {
      colChildren.push(
        XMLBuilder.wSelf('col', { 'w:w': width.toString() })
      );
    }
  }

  children.push(
    colChildren.length > 0
      ? XMLBuilder.w('cols', attrs, colChildren)
      : XMLBuilder.wSelf('cols', attrs)
  );
}
```

---

### Task 4: Update Parsing
**File:** `src/core/DocumentParser.ts` (parseSectionProperties method)
**Estimated Time:** 30 minutes

Add parsing for new properties:

```typescript
// Parse vertical alignment
const vAlignElements = XMLParser.extractElements(sectPr, "w:vAlign");
if (vAlignElements.length > 0) {
  const vAlign = vAlignElements[0];
  if (vAlign) {
    const val = XMLParser.extractAttribute(vAlign, "w:val") as 'top' | 'center' | 'bottom' | 'both';
    if (val) {
      sectionProps.verticalAlignment = val;
    }
  }
}

// Parse paper source
const paperSrcElements = XMLParser.extractElements(sectPr, "w:paperSrc");
if (paperSrcElements.length > 0) {
  const paperSrc = paperSrcElements[0];
  if (paperSrc) {
    const first = XMLParser.extractAttribute(paperSrc, "w:first");
    const other = XMLParser.extractAttribute(paperSrc, "w:other");

    if (first || other) {
      sectionProps.paperSource = {
        first: first ? parseInt(first, 10) : undefined,
        other: other ? parseInt(other, 10) : undefined,
      };
    }
  }
}

// Parse text direction
const textDirElements = XMLParser.extractElements(sectPr, "w:textDirection");
if (textDirElements.length > 0) {
  const textDir = textDirElements[0];
  if (textDir) {
    const val = XMLParser.extractAttribute(textDir, "w:val");
    if (val) {
      sectionProps.textDirection = val as 'ltr' | 'rtl' | 'tbRl' | 'btLr';
    }
  }
}

// Enhanced column parsing (separator and custom widths)
const colsElements = XMLParser.extractElements(sectPr, "w:cols");
if (colsElements.length > 0) {
  const cols = colsElements[0];
  if (cols) {
    const num = XMLParser.extractAttribute(cols, "w:num");
    const space = XMLParser.extractAttribute(cols, "w:space");
    const equalWidth = XMLParser.extractAttribute(cols, "w:equalWidth");
    const sep = XMLParser.extractAttribute(cols, "w:sep");

    // Extract individual column widths
    const colElements = XMLParser.extractElements(cols, "w:col");
    const columnWidths: number[] = [];
    for (const col of colElements) {
      const width = XMLParser.extractAttribute(col, "w:w");
      if (width) {
        columnWidths.push(parseInt(width, 10));
      }
    }

    if (num) {
      sectionProps.columns = {
        count: parseInt(num, 10),
        space: space ? parseInt(space, 10) : undefined,
        equalWidth: equalWidth === "1" || equalWidth === "true",
        separator: sep === "1" || sep === "true",
        columnWidths: columnWidths.length > 0 ? columnWidths : undefined,
      };
    }
  }
}
```

---

### Task 5: Create Comprehensive Tests
**File:** `tests/elements/SectionPropertiesEnhanced.test.ts`
**Estimated Time:** 40 minutes

Test structure:
```typescript
describe('Section Properties - Phase 4.5 Enhancements', () => {
  describe('Vertical Alignment', () => {
    // Test top, center, bottom, both
    // Test round-trip preservation
  });

  describe('Paper Source', () => {
    // Test first page tray
    // Test other pages tray
    // Test both trays
    // Test round-trip
  });

  describe('Column Separator', () => {
    // Test separator enabled
    // Test separator disabled
    // Test round-trip
  });

  describe('Text Direction', () => {
    // Test ltr, rtl, tbRl, btLr
    // Test round-trip
  });

  describe('Custom Column Widths', () => {
    // Test unequal columns
    // Test round-trip
    // Test interaction with column count
  });

  describe('Combined Properties', () => {
    // Test all new properties together
    // Test with existing properties
    // Multi-cycle round-trip
  });
});
```

**Expected Tests:** 20-25 tests

---

### Task 6: Verification
**File:** All modified files
**Estimated Time:** 15 minutes

1. Run section property tests: `npm test -- SectionPropertiesEnhanced.test.ts`
2. Run full test suite: `npm test`
3. Verify zero regressions
4. Check test coverage (target: 100% for new properties)

---

## Success Criteria

- ✅ All 4 missing properties implemented
- ✅ All setter methods working with fluent API
- ✅ XML generation correct per ECMA-376
- ✅ Parsing handles all properties correctly
- ✅ 20-25 new tests passing (100% coverage)
- ✅ Full test suite passing (899+ tests)
- ✅ Zero regressions
- ✅ Round-trip verification for all properties

---

## Expected Outcome

**Before Phase 4.5:**
- Total features: 78/127 (61.4%)
- Total tests: 899

**After Phase 4.5:**
- Total features: 82/127 (64.6%) (+4 properties)
- Total tests: 920-925 (+20-25 tests)
- Section properties: 100% complete

---

## ECMA-376 Compliance References

### Vertical Alignment (vAlign)
- **Element:** `<w:vAlign>`
- **Attribute:** `w:val` - "top", "center", "bottom", "both" (justified)
- **Section:** Part 4, Section 2.6.14

### Paper Source (paperSrc)
- **Element:** `<w:paperSrc>`
- **Attributes:** `w:first` (first page tray), `w:other` (other pages tray)
- **Section:** Part 4, Section 2.6.11

### Column Separator (sep in cols)
- **Element:** `<w:cols>`
- **Attribute:** `w:sep` - "1" (show) or "0" (hide)
- **Section:** Part 4, Section 2.6.2

### Text Direction (textDirection)
- **Element:** `<w:textDirection>`
- **Attribute:** `w:val` - "ltr", "rtl", "tbRl", "btLr"
- **Section:** Part 4, Section 2.6.13

---

## Next Steps After Completion

After Phase 4.5 is complete, we'll be at 64.6% feature completion with options to:
1. Phase 4.6 - Field Types (11 types, ~2-3 hours)
2. Phase 5 - Advanced features
3. v1.0.0 release preparation

---

**Status:** Ready to implement
**Priority:** High
**Complexity:** Medium (most work already done)
