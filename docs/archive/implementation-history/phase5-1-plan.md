# Phase 5.1 Implementation Plan - Table Styles

**Phase:** 5.1 - Table Styles
**Features:** 4 major features
**Estimated Time:** 4-5 hours
**Priority:** High
**Complexity:** Medium-High

---

## Overview

Implement complete table style support in the Style class, including:
1. Table-level formatting properties (tblPr)
2. Table cell formatting properties (tcPr)
3. Table row formatting properties (trPr)
4. Conditional table formatting (tblStylePr) for 12 regions

This completes the style system by adding table-specific formatting that can be defined in styles.xml and applied to tables throughout a document.

---

## Feature Breakdown

### Feature 1: Table-Level Properties (tblPr)
**Complexity:** Medium
**Time:** 1 hour

Add table formatting properties that apply to the entire table:

**Properties to Add:**
- `tblInd` - Table indentation from left margin (twips)
- `tblCellSpacing` - Default cell spacing (twips)
- `tblBorders` - Table borders (top, bottom, left, right, insideH, insideV)
- `tblCellMar` - Default cell margins (top, bottom, left, right)
- `shd` - Table background shading (color, fill, pattern)
- `jc` - Table alignment (left, center, right)

**Implementation:**
```typescript
export interface TableStyleFormatting {
  indent?: number;              // Table indent in twips
  cellSpacing?: number;         // Cell spacing in twips
  borders?: TableBorders;       // Table borders
  cellMargins?: CellMargins;    // Default cell margins
  shading?: ShadingProperties;  // Table shading
  alignment?: TableAlignment;   // left, center, right
}
```

### Feature 2: Table Cell Properties (tcPr)
**Complexity:** Medium
**Time:** 1 hour

Add cell formatting properties that apply to table cells:

**Properties to Add:**
- `tcBorders` - Cell borders (top, bottom, left, right, tl2br, tr2bl, insideH, insideV)
- `shd` - Cell background shading
- `tcMar` - Cell margins (top, bottom, left, right)
- `vAlign` - Vertical alignment (top, center, bottom)

**Implementation:**
```typescript
export interface TableCellStyleFormatting {
  borders?: CellBorders;        // Cell borders (8 possible borders)
  shading?: ShadingProperties;  // Cell shading
  margins?: CellMargins;        // Cell-specific margins
  verticalAlignment?: 'top' | 'center' | 'bottom';
}
```

### Feature 3: Table Row Properties (trPr)
**Complexity:** Low
**Time:** 30 minutes

Add row formatting properties (limited per ECMA-376):

**Properties to Add:**
- `cantSplit` - Prevent row from splitting across pages
- `tblHeader` - Mark row as header row
- `trHeight` - Row height (exact, at least, auto)

**Implementation:**
```typescript
export interface TableRowStyleFormatting {
  cantSplit?: boolean;          // Prevent row split
  isHeader?: boolean;           // Header row
  height?: number;              // Height in twips
  heightRule?: 'auto' | 'exact' | 'atLeast';
}
```

### Feature 4: Conditional Formatting (tblStylePr)
**Complexity:** High
**Time:** 2 hours

Add conditional formatting for 12 table regions:

**Conditional Regions (12 types):**
1. `wholeTable` - Entire table
2. `firstRow` - First row
3. `lastRow` - Last row
4. `firstCol` - First column
5. `lastCol` - Last column
6. `band1Vert` - Odd column banding
7. `band2Vert` - Even column banding
8. `band1Horz` - Odd row banding
9. `band2Horz` - Even row banding
10. `nwCell` - Northwest (top-left) corner cell
11. `neCell` - Northeast (top-right) corner cell
12. `swCell` - Southwest (bottom-left) corner cell
13. `seCell` - Southeast (bottom-right) corner cell

**Implementation:**
```typescript
export type ConditionalFormattingType =
  | 'wholeTable' | 'firstRow' | 'lastRow' | 'firstCol' | 'lastCol'
  | 'band1Vert' | 'band2Vert' | 'band1Horz' | 'band2Horz'
  | 'nwCell' | 'neCell' | 'swCell' | 'seCell';

export interface ConditionalTableFormatting {
  type: ConditionalFormattingType;
  paragraphFormatting?: ParagraphFormatting;  // pPr
  runFormatting?: RunFormatting;               // rPr
  tableFormatting?: TableStyleFormatting;      // tblPr
  cellFormatting?: TableCellStyleFormatting;   // tcPr
  rowFormatting?: TableRowStyleFormatting;     // trPr
}

export interface TableStyleProperties {
  table?: TableStyleFormatting;
  cell?: TableCellStyleFormatting;
  row?: TableRowStyleFormatting;
  rowBandSize?: number;         // Rows per band (default 1)
  colBandSize?: number;         // Columns per band (default 1)
  conditionalFormatting?: ConditionalTableFormatting[];
}
```

---

## Supporting Types

### Border Properties
```typescript
export interface BorderProperties {
  style?: 'none' | 'single' | 'double' | 'dashed' | 'dotted' | 'thick';
  size?: number;    // Size in eighths of a point
  space?: number;   // Spacing in points
  color?: string;   // Hex color (6 chars)
}

export interface TableBorders {
  top?: BorderProperties;
  bottom?: BorderProperties;
  left?: BorderProperties;
  right?: BorderProperties;
  insideH?: BorderProperties;  // Inside horizontal
  insideV?: BorderProperties;  // Inside vertical
}

export interface CellBorders extends TableBorders {
  tl2br?: BorderProperties;    // Top-left to bottom-right diagonal
  tr2bl?: BorderProperties;    // Top-right to bottom-left diagonal
}
```

### Shading Properties
```typescript
export interface ShadingProperties {
  fill?: string;      // Background fill color (hex)
  color?: string;     // Foreground color for patterns (hex)
  val?: 'clear' | 'solid' | 'pct5' | 'pct10' | 'pct20' | 'pct25'
       | 'pct30' | 'pct40' | 'pct50' | 'pct60' | 'pct70' | 'pct75'
       | 'pct80' | 'pct90' | 'diagStripe' | 'horzStripe' | 'vertStripe'
       | 'reverseDiagStripe' | 'horzCross' | 'diagCross';
}
```

### Cell Margins
```typescript
export interface CellMargins {
  top?: number;       // Twips
  bottom?: number;    // Twips
  left?: number;      // Twips
  right?: number;     // Twips
}
```

---

## Files to Modify

### 1. `src/formatting/Style.ts` (~250 lines added)

**Changes:**
- Add `TableStyleProperties` to `StyleProperties` interface
- Add setter methods for table properties:
  - `setTableFormatting(formatting: TableStyleFormatting): this`
  - `setTableCellFormatting(formatting: TableCellStyleFormatting): this`
  - `setTableRowFormatting(formatting: TableRowStyleFormatting): this`
  - `setRowBandSize(size: number): this`
  - `setColBandSize(size: number): this`
  - `addConditionalFormatting(conditional: ConditionalTableFormatting): this`
- Update `toXML()` to generate table style XML
- Update `isValid()` to validate table properties

**XML Generation:**
```xml
<w:style w:type="table" w:styleId="TableGrid">
  <w:name w:val="Table Grid"/>
  <w:tblPr>
    <w:tblInd w:w="0" w:type="dxa"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/>
      <w:left w:w="108" w:type="dxa"/>
      <w:bottom w:w="0" w:type="dxa"/>
      <w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>
  <w:tcPr>
    <w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>
  </w:tcPr>
  <w:tblStylePr w:type="firstRow">
    <w:pPr>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
    </w:rPr>
    <w:tcPr>
      <w:shd w:val="clear" w:color="auto" w:fill="D0CECE"/>
    </w:tcPr>
  </w:tblStylePr>
  <w:tblStylePr w:type="band1Horz">
    <w:tcPr>
      <w:shd w:val="clear" w:color="auto" w:fill="F0F0F0"/>
    </w:tcPr>
  </w:tblStylePr>
</w:style>
```

### 2. `src/formatting/StylesManager.ts` (~40 lines added)

**Changes:**
- Add `createTableStyle()` helper method
- Add `getTableStyles()` filter method
- Add built-in table styles:
  - `TableNormal`
  - `TableGrid`

**Built-in Table Styles:**
```typescript
static createTableGridStyle(): Style {
  return Style.create({
    styleId: 'TableGrid',
    name: 'Table Grid',
    type: 'table',
    tableFormatting: {
      borders: {
        top: { style: 'single', size: 4, color: '000000' },
        bottom: { style: 'single', size: 4, color: '000000' },
        left: { style: 'single', size: 4, color: '000000' },
        right: { style: 'single', size: 4, color: '000000' },
        insideH: { style: 'single', size: 4, color: '000000' },
        insideV: { style: 'single', size: 4, color: '000000' }
      },
      cellMargins: {
        top: 0,
        left: 108,
        bottom: 0,
        right: 108
      }
    }
  });
}
```

### 3. `src/core/DocumentParser.ts` (~150 lines added)

**Changes:**
- Add `parseTableStyle()` method for table-specific parsing
- Parse `tblPr`, `tcPr`, `trPr` elements
- Parse `tblStylePr` conditional formatting
- Extract borders, shading, margins from style XML

**Parsing Functions:**
```typescript
private parseTableStyleProperties(styleXml: string): TableStyleProperties {
  // Parse tblPr
  // Parse tcPr
  // Parse trPr
  // Parse tblStylePr elements
  // Return complete table style properties
}

private parseConditionalFormatting(styleXml: string): ConditionalTableFormatting[] {
  // Find all tblStylePr elements
  // For each, extract type attribute
  // Parse pPr, rPr, tblPr, tcPr, trPr children
  // Return array of conditional formats
}
```

### 4. `tests/formatting/TableStyles.test.ts` (NEW - ~600 lines)

**Test Coverage:**

**Basic Table Properties (8 tests):**
- Table indentation
- Cell spacing
- Table borders
- Cell margins
- Table shading
- Table alignment
- Cell vertical alignment
- Row height

**Conditional Formatting (12 tests):**
- First row formatting
- Last row formatting
- First column formatting
- Last column formatting
- Row banding (band1Horz, band2Horz)
- Column banding (band1Vert, band2Vert)
- Corner cells (nwCell, neCell, swCell, seCell)

**XML Generation (5 tests):**
- Complete table style XML
- Borders XML generation
- Shading XML generation
- Margins XML generation
- Conditional formatting XML

**Round-Trip Testing (3 tests):**
- Save and load table style
- Preserve all table properties
- Preserve conditional formatting

**Total: 28 tests**

---

## Implementation Strategy

### Phase 1: Type Definitions (30 minutes)
1. Add all TypeScript interfaces
2. Add table properties to StyleProperties
3. Update StyleType validation

### Phase 2: Basic Properties (1.5 hours)
1. Implement table formatting (tblPr)
2. Implement cell formatting (tcPr)
3. Implement row formatting (trPr)
4. Add setter methods with fluent API

### Phase 3: Conditional Formatting (2 hours)
1. Implement ConditionalTableFormatting interface
2. Add conditional formatting storage
3. Implement addConditionalFormatting() method
4. Generate XML for all 12 region types

### Phase 4: XML Generation (1 hour)
1. Update toXML() for table styles
2. Generate tblPr element
3. Generate tcPr element
4. Generate trPr element
5. Generate tblStylePr elements

### Phase 5: Parsing Support (1.5 hours)
1. Parse tblPr from styles.xml
2. Parse tcPr from styles.xml
3. Parse trPr from styles.xml
4. Parse tblStylePr elements
5. Handle all border/shading variations

### Phase 6: Testing (1.5 hours)
1. Write 28 comprehensive tests
2. Test all property types
3. Test conditional formatting
4. Test XML generation
5. Test round-trip preservation

---

## Quality Checklist

- [ ] All TypeScript interfaces defined
- [ ] All setter methods implemented
- [ ] Fluent API working
- [ ] XML generation correct per ECMA-376
- [ ] Parsing support complete
- [ ] All 28 tests passing
- [ ] Zero regressions in existing tests
- [ ] Full JSDoc documentation
- [ ] Round-trip verification working

---

## Success Criteria

1. **Complete Implementation:**
   - All 4 features implemented
   - All table style properties supported
   - 12 conditional formatting regions working

2. **Test Coverage:**
   - 28 new tests passing (100% coverage)
   - Total tests: ~1004 passing
   - Zero regressions

3. **ECMA-376 Compliance:**
   - Valid XML generation
   - Correct element ordering
   - Proper namespace usage

4. **Round-Trip Support:**
   - Save table styles correctly
   - Load and preserve all properties
   - Conditional formatting preserved

---

## Estimated Timeline

| Task | Time | Running Total |
|------|------|---------------|
| Type definitions | 30 min | 30 min |
| Basic properties (tblPr, tcPr, trPr) | 1.5 hrs | 2 hrs |
| Conditional formatting | 2 hrs | 4 hrs |
| XML generation | 1 hr | 5 hrs |
| Parsing support | 1.5 hrs | 6.5 hrs |
| Testing | 1.5 hrs | 8 hrs |
| **Total** | **8 hours** | |

**Conservative Estimate:** 8 hours (includes buffer for debugging)
**Optimistic Estimate:** 4-5 hours (if everything goes smoothly)

---

## Dependencies

- Existing Style class
- Existing Table, TableRow, TableCell classes
- XMLBuilder for XML generation
- DocumentParser for parsing
- Jest for testing

---

## Risks & Mitigation

**Risk 1: Complex conditional formatting**
- Mitigation: Implement one region type at a time, test thoroughly

**Risk 2: XML generation complexity**
- Mitigation: Reference existing table XML from Phase 4.3

**Risk 3: Border/shading edge cases**
- Mitigation: Use existing border logic from Table/TableCell

---

## Next Steps After Completion

With Phase 5.1 complete, the framework will have:
- ✅ Complete style system (paragraph, character, table)
- ✅ Full table formatting support
- ✅ 103/127 features complete (81%)
- ✅ 1004+ tests passing

**Recommended Next Phase:**
- Phase 5.2: Content Controls (9 features, ~3-4 hours)
- Phase 4.6: Field Types (11 types, ~3-4 hours)

---

**Plan Created:** October 23, 2025
**Status:** Ready for implementation
**Priority:** High
