# Helper Functions for Future Implementation

This document lists potential helper functions that would simplify common document manipulation tasks, identified during the Test_Code.docx modification project.

## Date: October 29, 2025

## Context

While implementing a script to modify Test_Code.docx with the docxmlater framework, several patterns emerged that would benefit from dedicated helper functions. The current framework (v1.3.0) provides comprehensive functionality, but certain operations require verbose code that could be streamlined.

---

## 1. Document.wrapParagraphInTable()

**Current Approach (Verbose):**
```typescript
// Create 1x1 table
const table = new Table(1, 1);
const cell = table.getRow(0)!.getCell(0)!;
cell.setShading({ fill: 'BFBFBF' });
const currentMargins = cell.getFormatting().margins || {};
cell.setMargins({
  top: 0,
  bottom: 0,
  left: currentMargins.left || 100,
  right: currentMargins.right || 100
});
cell.addParagraph(paragraph);
table.setLayout('auto');
table.setWidth(5000).setWidthType('pct');
```

**Proposed Helper:**
```typescript
Document.wrapParagraphInTable(
  paragraph: Paragraph,
  options?: {
    cellShading?: string,
    cellMargins?: { top?: number, bottom?: number, left?: number, right?: number },
    tableWidth?: number,
    tableWidthType?: 'auto' | 'dxa' | 'pct',
    tableLayout?: 'auto' | 'fixed'
  }
): Table
```

**Benefits:**
- Reduces 10+ lines to 1-2 lines
- Common pattern for wrapping headings, notes, or callouts in tables
- Eliminates boilerplate for cell/table configuration

---

## 2. Hyperlink.setColor()

**Current Approach (Requires Recreation):**
```typescript
// Get hyperlink properties
const text = hyperlink.getText();
const url = hyperlink.getUrl();

// Create new hyperlink with different color
const newLink = Hyperlink.createExternal(url, text, {
  color: '0000FF',
  underline: 'single'
});

// Replace in paragraph content array
content[index] = newLink;
```

**Proposed Helper:**
```typescript
hyperlink.setColor(color: string): this
```

**Benefits:**
- Fluent API consistency with other elements
- Avoids recreating hyperlink objects
- Simplifies bulk color changes (e.g., theme updates)

**Alternative:** Add `updateFormatting(formatting: RunFormatting): this` for more flexibility

---

## 3. Document.getElementsByStyle()

**Current Approach (Manual Iteration):**
```typescript
const heading2Paragraphs: Paragraph[] = [];
for (const element of doc.getBodyElements()) {
  if (element instanceof Paragraph && element.getStyle() === 'Heading2') {
    heading2Paragraphs.push(element);
  }
}
```

**Proposed Helper:**
```typescript
Document.getElementsByStyle(styleId: string): Paragraph[]
// Or more generic:
Document.getElementsByStyle<T extends BodyElement>(styleId: string, type?: new () => T): T[]
```

**Benefits:**
- Query pattern common in document processing
- Reduces boilerplate for style-based operations
- Enables functional programming patterns (map, filter, reduce)

**Example Usage:**
```typescript
doc.getElementsByStyle('Heading1').forEach(p => {
  // Modify all Heading 1 paragraphs
});
```

---

## 4. Table.isMultiCell()

**Current Approach:**
```typescript
const rows = table.getRows();
const cols = rows[0]?.getCells().length || 0;

if (rows.length === 1 && cols === 1) {
  // Skip 1x1 table
} else {
  // Process multi-cell table
}
```

**Proposed Helper:**
```typescript
table.isMultiCell(): boolean
// Returns false if table is 1x1, true otherwise
```

**Benefits:**
- Semantic clarity (intent is obvious)
- Reduces duplication of row/column counting logic
- Useful for conditional processing (e.g., skip wrapper tables)

---

## 5. NumberingLevel.setBulletCharacter()

**Current Approach (Constructor-Based):**
```typescript
const newLevel = new NumberingLevel({
  level: 0,
  format: 'bullet',
  text: '●',
  alignment: 'left',
  leftIndent: 720,
  hangingIndent: 360,
  font: 'Symbol'
});
abstractNum.addLevel(newLevel);
```

**Proposed Helper (Factory Method):**
```typescript
NumberingLevel.createBullet(
  level: number,
  bulletChar: string,
  options?: { leftIndent?: number, hangingIndent?: number, font?: string }
): NumberingLevel
```

**Note:** This partially exists as `NumberingLevel.createBulletLevel()`, but the API could be clearer.

**Alternative:** Add fluent setters to NumberingLevel for in-place modifications:
```typescript
level.setBulletCharacter(char: string): this
level.setIndentation(left: number, hanging: number): this
```

**Benefits:**
- More intuitive than reconstructing entire level
- Fluent API consistency
- Easier to modify existing levels

---

## 6. Paragraph.replaceHyperlinks()

**Current Approach (Manual Loop with Type Checking):**
```typescript
const content = paragraph.getContent();
for (let i = 0; i < content.length; i++) {
  const item = content[i];
  if (item instanceof Hyperlink) {
    const newLink = Hyperlink.createExternal(url, newText, { color: '0000FF' });
    content[i] = newLink;
  }
}
```

**Proposed Helper:**
```typescript
paragraph.replaceHyperlinks(
  callback: (hyperlink: Hyperlink, index: number) => Hyperlink | null
): this
// Return null to remove hyperlink, return new hyperlink to replace
```

**Benefits:**
- Functional programming pattern
- Handles content array manipulation internally
- Reduces bugs from manual array indexing

**Example Usage:**
```typescript
paragraph.replaceHyperlinks((link, i) => {
  if (link.getText() === 'Old Text') {
    return Hyperlink.createExternal(link.getUrl()!, 'New Text', { color: '0000FF' });
  }
  return link; // Keep unchanged
});
```

---

## 7. Table.setFirstRowFormatting() (Enhancement)

**Current Approach:**
```typescript
const firstRow = table.getRow(0);
if (firstRow) {
  for (const cell of firstRow.getCells()) {
    cell.setShading({ fill: 'D9D9D9' });
    cell.setMargins({ top: 0, bottom: 0, left: 100, right: 100 });

    for (const para of cell.getParagraphs()) {
      para.setAlignment('center');
      para.setSpaceBefore(60);
      para.setSpaceAfter(60);

      for (const run of para.getRuns()) {
        run.setBold(true);
        run.setFont('Verdana', 12);
        run.setColor('000000');
      }
    }
  }
}
```

**Proposed Enhancement:**
```typescript
table.setFirstRowFormatting({
  cellShading?: string,
  cellMargins?: CellMargins,
  paragraphAlignment?: ParagraphAlignment,
  paragraphSpacing?: { before?: number, after?: number },
  runFormatting?: RunFormatting
}): this
```

**Note:** The current `table.setFirstRowShading()` exists but only handles shading. This enhancement would handle all first-row formatting in one call.

**Benefits:**
- Common pattern for table headers
- Reduces nested loops
- Single call for complete header styling

---

## Summary

These helper functions represent patterns encountered during document modification tasks. The current framework is fully functional and production-ready, but these additions would improve:

1. **Developer Experience**: Less boilerplate, more semantic code
2. **Maintainability**: Centralized common operations
3. **Consistency**: Fluent API across all elements
4. **Safety**: Framework-managed operations reduce bugs

## Priority Ranking

1. **High Priority** (Common patterns, significant boilerplate reduction):
   - `Document.getElementsByStyle()`
   - `Table.setFirstRowFormatting()` enhancement
   - `Hyperlink.setColor()`

2. **Medium Priority** (Convenience, clarity improvements):
   - `Document.wrapParagraphInTable()`
   - `Table.isMultiCell()`

3. **Low Priority** (Alternative approaches exist):
   - `NumberingLevel.setBulletCharacter()` (factory methods exist)
   - `Paragraph.replaceHyperlinks()` (manual approach works)

---

## Implementation Notes

- All helpers should follow the existing fluent API pattern (return `this` for chaining)
- TypeScript type safety should be maintained
- Helpers should be backward-compatible with existing code
- Unit tests should be added for all new helpers
- Documentation should include examples

---

**Generated during:** Test_Code.docx modification project (October 29, 2025)
**Framework Version:** docxmlater v1.3.0
**Total lines saved by proposed helpers:** ~50+ lines for this single task
