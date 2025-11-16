# Document Formatting Helpers - Implementation Plan

## Overview

## 🔍 Analysis: Bullet Symbol Handling

### Current Implementation

The framework supports **BOTH** configurable and default bullet symbols:

**Hardcoded Defaults** (AbstractNumbering.ts:282):

```typescript
static createBulletList(
  abstractNumId: number,
  levels: number = 9,
  bullets: string[] = ['•', '○', '▪']  // ✅ Hardcoded defaults
): AbstractNumbering
```

**User-Configurable** (Document.ts:2221):

```typescript
createBulletList(levels: number = 3, bullets?: string[]): number
// Example: doc.createBulletList(3, ['▪', '○', '▸'])
```

### Integration with Template_UI

The Template_UI project (or any external consumer) can:

1. **Use defaults**: `doc.createBulletList()` → Gets ['•', '○', '▪']
2. **Pass custom symbols**: `doc.createBulletList(3, customSymbols)` → Gets user-defined symbols

### ⚠️ Important Design Decision

The `standardizeBulletSymbols()` method will:

- ✅ **Preserve** the user's chosen bullet symbols (•, ○, ▪, ▸, etc.)
- ✅ **Only change** formatting (bold, size, color, font)
- ✅ **NOT** replace symbols with hardcoded values

This respects the user's content choices while ensuring consistent formatting.

**Rationale**: The bullet symbol itself is content/style choice (like "•" vs "→"). The standardization should only fix formatting properties (making them bold, correct size, etc.), not override the user's symbol selection.

---

This plan outlines the implementation of five formatting helper functions for the docXMLater framework to ensure consistent styling across documents.

---

## Task 1: Update `ensureBlankLinesAfter1x1Tables()` - Add Normal Style

**Location**: `src/core/Document.ts:3654-3735`

**Change Required**: Add `.setStyle('Normal')` to blank paragraphs created after 1x1 tables.

### Current Code (line ~3710)

```typescript
const blankPara = Paragraph.create();
blankPara.setSpaceAfter(spacingAfter);
if (markAsPreserved) {
  blankPara.setPreserved(true);
}
```

### Updated Code

```typescript
const blankPara = Paragraph.create();
blankPara.setStyle("Normal"); // ✅ NEW: Ensure Normal style
blankPara.setSpaceAfter(spacingAfter);
if (markAsPreserved) {
  blankPara.setPreserved(true);
}
```

**Lines to Update**: ~3710, ~3724

**Impact**: Ensures blank lines have explicit "Normal" style rather than inheriting potentially incorrect styles.

---

## Task 2: Create `ensureBlankLinesAfterOtherTables()` Helper

**Location**: `src/core/Document.ts` - Add after `ensureBlankLinesAfter1x1Tables()`

### Method Signature

```typescript
public ensureBlankLinesAfterOtherTables(options?: {
  spacingAfter?: number;
  markAsPreserved?: boolean;
  filter?: (table: Table, index: number) => boolean;
}): {
  tablesProcessed: number;
  blankLinesAdded: number;
  existingLinesMarked: number;
}
```

### Implementation Strategy

1. Filter for NON-1x1 tables: `if (rowCount === 1 && colCount === 1) continue;`
2. Rest is identical to `ensureBlankLinesAfter1x1Tables()`
3. Create blank paragraph with Normal style
4. Apply spacing and preserve flag

### Code Template

```typescript
public ensureBlankLinesAfterOtherTables(options?: {
  spacingAfter?: number;
  markAsPreserved?: boolean;
  filter?: (table: Table, index: number) => boolean;
}): {
  tablesProcessed: number;
  blankLinesAdded: number;
  existingLinesMarked: number;
} {
  const spacingAfter = options?.spacingAfter ?? 120;
  const markAsPreserved = options?.markAsPreserved ?? true;
  const filter = options?.filter;

  let tablesProcessed = 0;
  let blankLinesAdded = 0;
  let existingLinesMarked = 0;

  const tables = this.getAllTables();

  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (!table) continue;

    const rowCount = table.getRowCount();
    const colCount = table.getColumnCount();

    // ✅ KEY DIFFERENCE: Skip 1x1 tables
    if (rowCount === 1 && colCount === 1) {
      continue;
    }

    // Apply filter if provided
    if (filter && !filter(table, i)) {
      continue;
    }

    tablesProcessed++;

    // Find table index in body elements
    const tableIndex = this.bodyElements.indexOf(table);
    if (tableIndex === -1) continue;

    // Check next element
    const nextElement = this.bodyElements[tableIndex + 1];

    if (nextElement instanceof Paragraph) {
      if (this.isParagraphBlank(nextElement)) {
        if (markAsPreserved && !nextElement.isPreserved()) {
          nextElement.setPreserved(true);
          existingLinesMarked++;
        }
      } else {
        const blankPara = Paragraph.create();
        blankPara.setStyle('Normal');  // ✅ Ensure Normal style
        blankPara.setSpaceAfter(spacingAfter);
        if (markAsPreserved) {
          blankPara.setPreserved(true);
        }
        this.bodyElements.splice(tableIndex + 1, 0, blankPara);
        blankLinesAdded++;
      }
    } else {
      const blankPara = Paragraph.create();
      blankPara.setStyle('Normal');  // ✅ Ensure Normal style
      blankPara.setSpaceAfter(spacingAfter);
      if (markAsPreserved) {
        blankPara.setPreserved(true);
      }
      this.bodyElements.splice(tableIndex + 1, 0, blankPara);
      blankLinesAdded++;
    }
  }

  return {
    tablesProcessed,
    blankLinesAdded,
    existingLinesMarked,
  };
}
```

---

## Task 3: Bullet List Formatting Standardization

### Part A: Core Change - Extend `NumberingLevel` Class

**File**: `src/formatting/NumberingLevel.ts`

#### Step 1: Add Properties

Add to `NumberingLevelProperties` interface (line ~36):

```typescript
export interface NumberingLevelProperties {
  // ... existing properties ...

  /** Text color in hex (without #) */
  color?: string;

  /** Whether the numbering text is bold */
  bold?: boolean;
}
```

#### Step 2: Update Constructor

Update defaults in constructor (line ~83):

```typescript
this.properties = {
  // ... existing defaults ...
  color: properties.color || "000000",
  bold: properties.bold !== undefined ? properties.bold : false,
};
```

#### Step 3: Add Setters

```typescript
/**
 * Sets the text color
 * @param color Hex color without # (e.g., '000000')
 */
setColor(color: string): this {
  this.properties.color = color;
  return this;
}

/**
 * Sets whether the numbering text is bold
 * @param bold Whether to make bold
 */
setBold(bold: boolean): this {
  this.properties.bold = bold;
  return this;
}
```

#### Step 4: Update XML Generation

In `toXML()` method, add to `rPrChildren` array (after fontSize, line ~255):

```typescript
// Bold
if (this.properties.bold) {
  rPrChildren.push(XMLBuilder.wSelf("b"));
  rPrChildren.push(XMLBuilder.wSelf("bCs"));
}

// Color
if (this.properties.color) {
  rPrChildren.push(
    XMLBuilder.wSelf("color", { "w:val": this.properties.color })
  );
}
```

### Part B: Retroactive Fix - `standardizeBulletSymbols()`

**Location**: `src/core/Document.ts`

````typescript
/**
 * Standardizes all bullet list symbols to be bold, 12pt, and black (#000000)
 *
 * This helper ensures consistent bullet formatting across all bullet lists in the document.
 * It modifies the numbering definitions (not individual paragraphs).
 *
 * @param options Formatting options
 * @returns Statistics about lists updated
 *
 * @example
 * ```typescript
 * // Standardize all bullet symbols with defaults
 * const result = doc.standardizeBulletSymbols();
 * console.log(`Updated ${result.listsUpdated} bullet lists`);
 *
 * // Custom formatting
 * const result = doc.standardizeBulletSymbols({
 *   bold: true,
 *   fontSize: 28,  // 14pt
 *   color: 'FF0000',  // Red
 *   font: 'Calibri'
 * });
 * ```
 */
public standardizeBulletSymbols(options?: {
  bold?: boolean;
  fontSize?: number;
  color?: string;
  font?: string;
}): {
  listsUpdated: number;
  levelsModified: number;
} {
  const {
    bold = true,
    fontSize = 24,  // 12pt
    color = '000000',
    font = 'Calibri'
  } = options || {};

  let listsUpdated = 0;
  let levelsModified = 0;

  const instances = this.numberingManager.getAllInstances();

  for (const instance of instances) {
    const abstractNumId = instance.getAbstractNumId();
    const abstractNum = this.numberingManager.getAbstractNumbering(abstractNumId);

    if (!abstractNum) continue;

    // Only process bullet lists
    const level0 = abstractNum.getLevel(0);
    if (!level0 || level0.getFormat() !== 'bullet') continue;

    // Update all 9 levels (0-8)
    for (let levelIndex = 0; levelIndex < 9; levelIndex++) {
      const numLevel = abstractNum.getLevel(levelIndex);
      if (!numLevel) continue;

      numLevel.setFont(font);
      numLevel.setFontSize(fontSize);
      numLevel.setBold(bold);
      numLevel.setColor(color);

      levelsModified++;
    }

    listsUpdated++;
  }

  return { listsUpdated, levelsModified };
}
````

### Part C: New Defaults - Update `createBulletList()`

**Location**: `src/formatting/NumberingManager.ts`

Update default properties when creating bullet levels:

```typescript
// In createBulletList() method
font: 'Calibri',
fontSize: 24,      // 12pt
bold: true,        // ✅ NEW
color: '000000'    // ✅ NEW
```

---

## Task 4: Numbered List Prefix Standardization

### Part A: Retroactive Fix - `standardizeNumberedListPrefixes()`

**Location**: `src/core/Document.ts`

````typescript
/**
 * Standardizes numbered list prefixes (1., a., i., etc.) to Verdana 12pt bold black
 *
 * This only affects the prefix/number, not the text content after it.
 * It modifies the numbering definitions in the document.
 *
 * @param options Formatting options
 * @returns Statistics about lists updated
 *
 * @example
 * ```typescript
 * // Standardize all numbered list prefixes
 * const result = doc.standardizeNumberedListPrefixes();
 * console.log(`Updated ${result.listsUpdated} numbered lists`);
 *
 * // Custom formatting for prefixes
 * const result = doc.standardizeNumberedListPrefixes({
 *   bold: true,
 *   fontSize: 24,
 *   color: '000000',
 *   font: 'Verdana'
 * });
 * ```
 */
public standardizeNumberedListPrefixes(options?: {
  bold?: boolean;
  fontSize?: number;
  color?: string;
  font?: string;
}): {
  listsUpdated: number;
  levelsModified: number;
} {
  const {
    bold = true,
    fontSize = 24,  // 12pt
    color = '000000',
    font = 'Verdana'
  } = options || {};

  let listsUpdated = 0;
  let levelsModified = 0;

  const instances = this.numberingManager.getAllInstances();

  for (const instance of instances) {
    const abstractNumId = instance.getAbstractNumId();
    const abstractNum = this.numberingManager.getAbstractNumbering(abstractNumId);

    if (!abstractNum) continue;

    // Only process numbered lists (skip bullet lists)
    const level0 = abstractNum.getLevel(0);
    if (!level0 || level0.getFormat() === 'bullet') continue;

    // Update all 9 levels (0-8)
    for (let levelIndex = 0; levelIndex < 9; levelIndex++) {
      const numLevel = abstractNum.getLevel(levelIndex);
      if (!numLevel) continue;

      numLevel.setFont(font);
      numLevel.setFontSize(fontSize);
      numLevel.setBold(bold);
      numLevel.setColor(color);

      levelsModified++;
    }

    listsUpdated++;
  }

  return { listsUpdated, levelsModified };
}
````

### Part B: New Defaults - Update `createNumberedList()`

**Location**: `src/formatting/NumberingManager.ts`

Update default properties when creating numbered levels:

```typescript
// In createNumberedList() method
font: 'Verdana',
fontSize: 24,      // 12pt
bold: true,        // ✅ NEW
color: '000000'    // ✅ NEW
```

---

## Task 5: Hyperlink Standardization Helper

**Location**: `src/core/Document.ts`

````typescript
/**
 * Standardizes all hyperlinks in the document to Verdana 12pt blue (#0000FF) underline
 *
 * This applies consistent formatting to all hyperlinks throughout the document,
 * including those in tables and headers/footers.
 *
 * @param options Formatting options
 * @returns Number of hyperlinks updated
 *
 * @example
 * ```typescript
 * // Use default formatting (Verdana 12pt blue underline)
 * const count = doc.standardizeAllHyperlinks();
 * console.log(`Standardized ${count} hyperlinks`);
 *
 * // Custom hyperlink formatting
 * const count = doc.standardizeAllHyperlinks({
 *   font: 'Arial',
 *   size: 11,
 *   color: 'FF0000',  // Red
 *   underline: true
 * });
 * ```
 */
public standardizeAllHyperlinks(options?: {
  font?: string;
  size?: number;
  color?: string;
  underline?: boolean;
}): number {
  const {
    font = 'Verdana',
    size = 12,
    color = '0000FF',
    underline = true
  } = options || {};

  const hyperlinks = this.getHyperlinks();

  for (const { hyperlink } of hyperlinks) {
    hyperlink.setFormatting({
      font: font,
      size: size,
      color: color,
      underline: underline ? 'single' : false,
    });
  }

  return hyperlinks.length;
}
````

---

## API Summary

### New Public Methods in `Document.ts`

| Method                               | Purpose                                                  | Returns                                                     |
| ------------------------------------ | -------------------------------------------------------- | ----------------------------------------------------------- |
| `ensureBlankLinesAfterOtherTables()` | Add blank Normal-styled lines after non-1x1 tables       | `{ tablesProcessed, blankLinesAdded, existingLinesMarked }` |
| `standardizeBulletSymbols()`         | Fix existing bullet lists to be bold, 12pt, #000000      | `{ listsUpdated, levelsModified }`                          |
| `standardizeNumberedListPrefixes()`  | Fix existing numbered list prefixes to Verdana 12pt bold | `{ listsUpdated, levelsModified }`                          |
| `standardizeAllHyperlinks()`         | Set all hyperlinks to Verdana 12pt #0000FF underline     | `number`                                                    |

### Modified Methods

| File                  | Method                             | Change                                                   |
| --------------------- | ---------------------------------- | -------------------------------------------------------- |
| `Document.ts`         | `ensureBlankLinesAfter1x1Tables()` | Add `blankPara.setStyle('Normal')` at lines ~3710, ~3724 |
| `NumberingLevel.ts`   | Constructor                        | Add `color` and `bold` properties with defaults          |
| `NumberingLevel.ts`   | `toXML()`                          | Add bold and color XML generation                        |
| `NumberingManager.ts` | `createBulletList()`               | Update defaults: Calibri 12pt bold #000000               |
| `NumberingManager.ts` | `createNumberedList()`             | Update defaults: Verdana 12pt bold #000000               |

### New Properties in `NumberingLevel.ts`

```typescript
color?: string;   // Hex color without #
bold?: boolean;   // Whether numbering text is bold
```

### New Methods in `NumberingLevel.ts`

```typescript
setColor(color: string): this
setBold(bold: boolean): this
```

---

## Implementation Order

1. **First**: Update `NumberingLevel.ts` (core dependency)

   - Add `color` and `bold` properties
   - Add setters
   - Update XML generation

2. **Second**: Update `ensureBlankLinesAfter1x1Tables()` (simple change)

   - Add `.setStyle('Normal')` at 2 locations

3. **Third**: Add `ensureBlankLinesAfterOtherTables()` (copy/modify existing)

4. **Fourth**: Add `standardizeBulletSymbols()` (depends on NumberingLevel)

5. **Fifth**: Add `standardizeNumberedListPrefixes()` (similar to #4)

6. **Sixth**: Add `standardizeAllHyperlinks()` (independent)

7. **Seventh**: Update defaults in `NumberingManager.ts`

---

## Testing Checklist

- [ ] Test `ensureBlankLinesAfter1x1Tables()` adds Normal style
- [ ] Test `ensureBlankLinesAfterOtherTables()` skips 1x1 tables correctly
- [ ] Test `standardizeBulletSymbols()` updates all bullet levels
- [ ] Test `standardizeNumberedListPrefixes()` only affects numbered lists
- [ ] Test `standardizeAllHyperlinks()` updates all links
- [ ] Verify new list defaults apply to newly created lists
- [ ] Test backward compatibility with existing code

---

## Key Considerations

### Scalability

- All methods iterate document structures efficiently (O(n))
- Use existing managers (NumberingManager, StylesManager)
- No performance impact for small/medium documents

### Security

- All hex colors validated (existing validation in Run/Hyperlink)
- Style names validated through StylesManager
- No user input directly in XML

### Best Practices

- Follow existing patterns (e.g., `applyStandardListFormatting()`)
- Maintain backward compatibility
- Provide sensible defaults with override options
- Return statistics for debugging
- Comprehensive JSDoc comments with examples
