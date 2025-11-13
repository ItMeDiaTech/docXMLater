# Style Refactoring Reference Guide

**Version:** 2.2.0
**Date:** November 2025
**Status:** In Progress

---

## Overview

This document tracks the refactored style application API, providing a comprehensive reference for all new functions, helpers, and types added to support flexible style application with custom formatting.

---

## New Types and Interfaces

### Location: `src/types/formatting.ts`

#### `EmphasisType`

```typescript
type EmphasisType = "bold" | "italic" | "underline";
```

Text emphasis options for formatting.

#### `ListPrefix`

```typescript
interface ListPrefix {
  format: "bullet" | "number";
  style: string; // e.g., '•', '1.', 'a)'
}
```

Configuration for list prefix styling.

#### `FormatOptions`

```typescript
interface FormatOptions {
  // Text formatting
  font?: string; // Font family (e.g., 'Arial', 'Verdana')
  size?: number; // Font size in points
  color?: string; // Text color as 6-digit hex (e.g., 'FF0000')
  emphasis?: EmphasisType[]; // Array of emphasis types

  // Alignment
  alignment?: "left" | "right" | "center" | "justify";

  // Spacing (in points)
  spaceAbove?: number; // Space before paragraph
  spaceBelow?: number; // Space after paragraph
  lineSpacing?: number; // Line spacing

  // Indentation (in inches)
  indentLeft?: number;
  indentRight?: number;
  indentFirst?: number;
  indentHanging?: number;

  // Padding (in points) - for table cells
  paddingTop?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  paddingRight?: number;

  // List formatting
  prefixList?: string | ListPrefix;

  // Advanced options
  borderColor?: string; // 6-digit hex
  borderWidth?: number; // In points
  shading?: string; // Background color as 6-digit hex
  keepWithNext?: boolean; // Only set if true
  keepLines?: boolean; // Only set if true
}
```

Complete formatting configuration for style application.

**Unit Conversions:**

- Points to twips: `points * 20` (1 point = 20 twips)
- Inches to twips: `inches * 1440` (1 inch = 1440 twips)

#### `StyleApplyOptions`

```typescript
interface StyleApplyOptions {
  paragraphs?: Paragraph[]; // Specific paragraphs to apply style to
  keepProperties?: string[]; // Properties to preserve from existing formatting
  format?: FormatOptions; // Custom formatting to apply
}
```

Options for applying styles to paragraphs.

---

## Updated Public Methods

### Location: `src/core/Document.ts`

All style application methods now accept optional `StyleApplyOptions` parameter.

#### `applyH1(options?: StyleApplyOptions): number`

**Line:** 3794
Applies Heading 1 style to paragraphs with H1-like style names.

**Example:**

```typescript
// Simple usage
doc.applyH1();

// With custom formatting
doc.applyH1({
  format: { font: "Arial", size: 18, emphasis: ["bold"] },
});

// Preserve specific properties
doc.applyH1({
  keepProperties: ["bold", "color"],
  format: { font: "Verdana" },
});
```

#### `applyH2(options?: StyleApplyOptions): number`

**Line:** 3811
Applies Heading 2 style to paragraphs with H2-like style names.

**Example:**

```typescript
doc.applyH2({
  format: { font: "Verdana", size: 14, color: "000000" },
});
```

#### `applyH3(options?: StyleApplyOptions): number`

**Line:** 3828
Applies Heading 3 style to paragraphs with H3-like style names.

**Example:**

```typescript
doc.applyH3({
  format: { font: "Verdana", size: 12, emphasis: ["bold"] },
});
```

#### `applyNormal(options?: StyleApplyOptions): number`

**Line:** 3839
Applies Normal style to paragraphs without recognized styles.

**Example:**

```typescript
doc.applyNormal({
  format: {
    font: "Verdana",
    size: 12,
    alignment: "justify",
    spaceBelow: 3,
  },
});
```

#### `applyNumList(options?: StyleApplyOptions): number`

**Line:** 3870
Applies list style to numbered lists.

#### `applyBulletList(options?: StyleApplyOptions): number`

**Line:** 3881
Applies list style to bullet lists.

#### `applyTOC(options?: StyleApplyOptions): number`

**Line:** 3892
Applies Table of Contents style.

#### `applyTOD(options?: StyleApplyOptions): number`

**Line:** 3903
Applies Top of Document style.

#### `applyCaution(options?: StyleApplyOptions): number`

**Line:** 3914
Applies Caution/Warning style.

#### `applyCellHeader(options?: StyleApplyOptions): number`

**Line:** 3925
Applies header style to table cell paragraphs (typically first row).

**Example:**

```typescript
doc.applyCellHeader({
  format: {
    font: "Arial",
    size: 12,
    emphasis: ["bold"],
    alignment: "center",
  },
});
```

---

## New Helper Methods (Private)

### Heading Options: `src/core/Document.ts`

#### `applyFormatOptions(para: Paragraph, options: FormatOptions): void`

**Line:** 3928
**Purpose:** Applies formatting options to a paragraph.

**Functionality:**

- Text formatting (font, size, color, emphasis) applied to all runs
- Alignment applied to paragraph
- Spacing converted from points to twips (1pt = 20 twips)
- Indentation converted from inches to twips (1in = 1440 twips)
- Advanced options (keepWithNext, keepLines) only set if true

**Unit Conversions:**

```typescript
// Spacing: points → twips
spaceAbove * 20;
spaceBelow * 20;
lineSpacing * 20;

// Indentation: inches → twips
indentLeft * 1440;
indentRight * 1440;
indentFirst * 1440;
indentHanging * 1440;
```

#### `clearFormattingExcept(para: Paragraph, keepProperties: string[]): void`

**Line:** 3038
**Purpose:** Selectively clears formatting while preserving specific properties.

**Functionality:**

- Saves specified properties from paragraph formatting
- Clears all paragraph formatting
- Restores saved properties
- Handles run-level properties using appropriate setters

**Supported Properties to Keep:**

- `bold`, `italic`, `underline`
- `color`, `font`, `size`
- `highlight`, `strike`
- `subscript`, `superscript`

#### `applyStyleToMatching(targetStyle: string, options: StyleApplyOptions | undefined, matcher: (style: string) => boolean): number`

**Line:** 4032
**Purpose:** Helper to apply style to matching paragraphs.

**Functionality:**

- Filters paragraphs by style name using matcher function
- Skips preserved paragraphs
- Applies target style
- Handles selective property preservation if `keepProperties` specified
- Applies custom formatting if `format` option provided
- Returns count of paragraphs updated

---

## New Element Methods

### Location: `src/elements/Run.ts`

#### `clearFormatting(): this`

**Line:** 1280
**Purpose:** Clears all formatting from a run.

**Example:**

```typescript
run.clearFormatting();
```

### Location: `src/elements/Paragraph.ts`

#### `clearDirectFormatting(): this`

**Line:** 2343
**Purpose:** Clears all direct formatting from paragraph and its runs.

**Functionality:**

- Clears paragraph-level formatting
- Preserves style reference and numbering
- Clears formatting from all runs

**Example:**

```typescript
paragraph.clearDirectFormatting();
```

---

## Existing Methods (Updated Signatures)

### Multiple headings: `src/core/Document.ts`

#### `cleanFormatting(styleNames?: string[]): number`

**Line:** 3749
**Purpose:** Cleans direct formatting from paragraphs that have a style applied.

**Parameters:**

- `styleNames` (optional): Array of specific style names to clean

**Returns:** Number of paragraphs cleaned

**Example:**

```typescript
// Clean all styled paragraphs
doc.cleanFormatting();

// Clean specific styles only
doc.cleanFormatting(["Heading1", "Heading2", "Normal"]);
```

---

## Style Matchers

Regular expressions used to match style names:

| Method            | Regex Pattern                                                | Matches                                     |
| ----------------- | ------------------------------------------------------------ | ------------------------------------------- |
| `applyH1`         | `/^(heading\s*1\|header\s*1\|h1)$/i`                         | Heading1, Heading 1, Header1, H1            |
| `applyH2`         | `/^(heading\s*2\|header\s*2\|h2)$/i`                         | Heading2, Heading 2, Header2, H2            |
| `applyH3`         | `/^(heading\s*3\|header\s*3\|h3)$/i`                         | Heading3, Heading 3, Header3, H3            |
| `applyNumList`    | `/^(list\s*number\|numbered\s*list\|list\s*paragraph)$/i`    | List Number, Numbered List, List Paragraph  |
| `applyBulletList` | `/^(list\s*bullet\|bullet\s*list\|list\s*paragraph)$/i`      | List Bullet, Bullet List, List Paragraph    |
| `applyTOC`        | `/^(toc\|table\s*of\s*contents\|toc\s*heading)$/i`           | TOC, Table Of Contents, TOC Heading         |
| `applyTOD`        | `/^(tod\|top\s*of\s*document\|document\s*top)$/i`            | TOD, Top Of Document, Document Top          |
| `applyCaution`    | `/^(caution\|warning\|important\|alert)$/i`                  | Caution, Warning, Important, Alert          |
| `applyNormal`     | `/^(heading\|header\|h\d\|list\|toc\|tod\|caution\|table)/i` | Applies to styles NOT matching this pattern |

All patterns are case-insensitive (`i` flag).

---

## Usage Patterns

### Pattern 1: Simple Style Application

```typescript
// Apply default style, clear all formatting
doc.applyH1();
doc.applyH2();
doc.applyNormal();
```

### Pattern 2: Custom Formatting

```typescript
// Apply style with custom formatting
doc.applyH2({
  format: {
    font: "Verdana",
    size: 14,
    color: "000000",
    emphasis: ["bold"],
    alignment: "left",
    spaceBelow: 6,
    indentLeft: 0.25,
  },
});
```

### Pattern 3: Selective Preservation

```typescript
// Keep existing bold and color, apply new formatting
doc.applyH1({
  keepProperties: ["bold", "color"],
  format: {
    font: "Arial",
    size: 18,
    alignment: "center",
  },
});
```

### Pattern 4: Specific Paragraphs

```typescript
// Apply to specific paragraphs only
const someParagraphs = doc.getAllParagraphs().slice(0, 10);
doc.applyH1({
  paragraphs: someParagraphs,
  format: { font: "Verdana", size: 18 },
});
```

### Pattern 5: Complex Formatting

```typescript
// Full-featured example
doc.applyNormal({
  keepProperties: ["bold", "italic"],
  format: {
    font: "Verdana",
    size: 12,
    color: "000000",
    alignment: "justify",
    spaceAbove: 0,
    spaceBelow: 3,
    lineSpacing: 1.15,
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0.5,
    keepWithNext: true,
    shading: "F0F0F0",
  },
});
```

---

## Migration Guide

### Before (v2.1.0 and earlier)

```typescript
// Simple, no options
doc.applyH1(); // Returns number
```

### After (v2.2.0+)

```typescript
// Still backwards compatible
doc.applyH1(); // Works exactly the same

// New: With options
doc.applyH1({
  format: { font: "Arial", size: 18 },
});

// New: Preserve properties
doc.applyH1({
  keepProperties: ["bold", "color"],
});
```

**No breaking changes** - All existing code continues to work.

---

## Template_UI Integration

### Location: `src/services/document/WordDocumentProcessor.ts`

**Lines:** 653-673

Currently calls all style methods with default parameters (no options):

```typescript
const h1Count = doc.applyH1();
const h2Count = doc.applyH2();
const h3Count = doc.applyH3();
const numListCount = doc.applyNumList();
const bulletListCount = doc.applyBulletList();
const tocCount = doc.applyTOC();
const todCount = doc.applyTOD();
const cautionCount = doc.applyCaution();
const cellHeaderCount = doc.applyCellHeader();
const hyperlinkCount = doc.applyHyperlink();
const normalCount = doc.applyNormal();
const cleanedCount = doc.cleanFormatting();
```

**Future Enhancement:** Add UI to configure custom formatting per style.

---

## Testing Notes

### Test Cases to Add

1. **Format Options Application**

   - Font, size, color application
   - Emphasis (bold, italic, underline)
   - Alignment
   - Spacing conversions (points → twips)
   - Indentation conversions (inches → twips)

2. **Property Preservation**

   - Keep specified properties
   - Clear non-specified properties
   - Preserve paragraph-level properties
   - Preserve run-level properties

3. **Style Matching**

   - Case-insensitive matching
   - Multiple style name variants
   - Normal style fallback logic

4. **Edge Cases**
   - Empty paragraphs
   - Preserved paragraphs (should skip)
   - Paragraphs without style names
   - Tables with no first row

---

## Known Limitations

1. **Padding Properties:** `paddingTop/Bottom/Left/Right` are defined but not yet fully implemented for table cells (TODO).

2. **List Prefix:** `prefixList` option is defined but custom list styling not yet implemented (TODO).

3. **Border Properties:** `borderColor` and `borderWidth` are defined but not yet applied (TODO).

4. **Hanging Indent:** Set directly through `para.formatting.indentation.hanging` rather than dedicated setter method.

---

## Files Modified

### docXMLater

1. `src/types/formatting.ts` - New types and interfaces
2. `src/core/Document.ts` - Updated methods and new helpers
3. `src/elements/Paragraph.ts` - Added `clearDirectFormatting()`
4. `src/elements/Run.ts` - Added `clearFormatting()`
5. `src/index.ts` - Exported new types
6. `package.json` - Version bump to 2.2.0

### Template_UI

1. `src/services/document/WordDocumentProcessor.ts` - Integration (lines 653-673)
2. `package.json` - Updated docXMLater dependency

---

## Next Steps

1. Add unit tests for new functionality
2. Add UI in Template_UI to configure custom formatting per style
3. Implement remaining TODO items (padding, borders, list prefix)
4. Consider adding preset style configurations (e.g., "Corporate", "Academic", "Minimal")
5. Document performance implications of complex formatting operations

---

## Notes

- All style application methods are backwards compatible
- Options parameter is optional - default behavior unchanged
- Boolean properties simplified: only set if true (no undefined checks needed)
- Unit conversions handled automatically (points/inches → twips)
- Property preservation uses getters/setters (no direct formatting access)

---

**Last Updated:** November 13, 2025
**Next Review:** After user testing and feedback
