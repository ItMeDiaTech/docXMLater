# DocXMLater Phase 4-5 Implementation Plan
**Started:** October 23, 2025
**Current Status:** Phase 4.0 Complete (Critical Fixes) ✅

---

## Overview

Implementing 127 missing features across 11 phases to achieve v1.0.0 with full ECMA-376 compliance.

**Progress:** 5 of 127 features complete (3.9%)

**Phases Complete:** 4.0.1, 4.0.2, 4.0.3
**Test Count:** 652 tests (+26 new tests in Phase 4.0)
**Time Spent:** ~3 hours total for all Phase 4.0 critical fixes

---

## Phase 4.0: CRITICAL FIXES (Week 1)

### ✅ Phase 4.0.1: Paragraph Borders, Shading, Tabs (COMPLETED)
**Status:** Complete - All 14 tests passing
**Files Modified:**
- `src/elements/Paragraph.ts` - Added interfaces, fixed setters, implemented XML serialization
- `src/core/DocumentParser.ts` - Implemented parsing for borders/shading/tabs
- `tests/elements/ParagraphBordersShading.test.ts` - 14 comprehensive round-trip tests

**Accomplishments:**
- Fixed broken stub implementations (data loss issue resolved)
- Added proper TypeScript interfaces (BorderDefinition, ShadingPattern, TabStop)
- Implemented full XML serialization per ECMA-376
- Implemented complete parsing with round-trip support
- 100% test coverage for new features

---

### ✅ Phase 4.0.2: Image Parsing (COMPLETED)
**Status:** Complete - All inline images now preserved
**Files Modified:**
- `src/core/DocumentParser.ts:819-928` - Implemented `parseDrawingFromObject()` method (110 lines)
- `implement/phase4-0-2-complete.md` - Completion documentation

**Accomplishments:**
- Fixed complete data loss for embedded images
- Implemented DrawingML structure parsing (wp:inline → a:graphic → pic:pic → a:blip)
- Parse image dimensions (cx, cy in EMUs)
- Resolve relationship IDs to get image file paths
- Read image data from ZIP archive
- Create Image objects with proper metadata
- Register with ImageManager

**Limitations:**
- Only `wp:inline` implemented (floating images `wp:anchor` → Phase 4.4)

---

### ✅ Phase 4.0.3: Table Cell Properties Parsing (COMPLETED)
**Status:** Complete - All cell formatting preserved
**Files Modified:**
- `src/core/DocumentParser.ts:1007-1091` - Implemented cell properties parsing (85 lines)
- `src/elements/TableCell.ts:265-278` - Fixed XML element names (1 line)
- `tests/elements/TableCellProperties.test.ts` - 12 comprehensive round-trip tests (383 lines)
- `implement/phase4-0-3-complete.md` - Completion documentation

**Accomplishments:**
- Fixed complete data loss for table cell formatting
- Parse cell width (`w:tcW`)
- Parse cell borders (`w:tcBorders`)
- Parse cell shading (`w:shd`)
- Parse cell margins (`w:tcMar`) - PRIMARY GOAL ✅
- Parse vertical alignment (`w:vAlign`)
- Parse column span (`w:gridSpan`)
- 12 new tests (652 total, +1.9%)
- Full round-trip support for professional tables

**Note:** Cell margins API already existed (v0.24.0), only parsing was missing

---

## Phase 4.1: RUN PROPERTIES (Weeks 2-3) - 22 Features

### Critical Character Properties (Week 2)

**4.1.1 Character Style Reference**
- Add `characterStyle?: string` to RunFormatting
- Serialize as `<w:rStyle w:val="..."/>`
- Parse from w:rStyle
- Tests: 3 (set, save, load)

**4.1.2 Text Border**
- Add `border?: { style, width, color, space }` to RunFormatting
- Serialize as `<w:bdr .../>`
- Parse from w:bdr
- Tests: 5 (all border styles)

**4.1.3 Character Shading**
- Add `shading?: { fill, pattern, color }` to RunFormatting
- Serialize as `<w:shd .../>`
- Parse from w:shd
- Tests: 4 (solid, patterns)

**4.1.4 Emphasis Marks**
- Add `emphasis?: 'dot' | 'comma' | 'circle' | 'underDot'`
- Serialize as `<w:em w:val="..."/>`
- Tests: 4 (each type)

**4.1.5 Complex Script Variants**
- Add `complexScriptBold?: boolean`, `complexScriptItalic?: boolean`
- Serialize as `<w:bCs/>`, `<w:iCs/>`
- Tests: 3 (RTL text)

**4.1.6-4.1.10: Additional Properties**
- Character spacing, scaling, position, kerning, language
- Tests: 15 total

### Additional Character Effects (Week 3)

**4.1.11-4.1.22: Text Effects & Specialized**
- Outline, shadow, emboss, imprint, effects
- Fit text, East Asian layout, RTL
- Vanish/hidden, no proof, snap to grid
- Special vanish, math
- Tests: 25 total

**Week 2-3 Total:** 22 features, ~40 tests

---

## Phase 4.2: PARAGRAPH PROPERTIES (Weeks 4-5) - 28 Features

### Critical Paragraph Properties (Week 4)

**4.2.1 Widow/Orphan Control**
- Add `widowControl?: boolean`
- Serialize as `<w:widowControl w:val="..."/>`
- Tests: 2

**4.2.2 Outline Level**
- Add `outlineLevel?: number` (0-9 for TOC)
- Critical for table of contents hierarchy
- Tests: 3

**4.2.3 Frame Properties**
- Add `frame?: { width, height, x, y, hAnchor, vAnchor }`
- Text box positioning
- Tests: 5

**4.2.4-4.2.12: Additional Properties**
- Bidirectional, mirror indents, suppress line numbers
- East Asian typography (5 properties)
- Tests: 20

### Additional Paragraph Properties (Week 5)

**4.2.13-4.2.28: Specialized Properties**
- Text direction, text alignment, adjust right indent
- Snap to grid, text box tight wrap
- Div ID, conditional formatting, section properties
- Tests: 22

**Week 4-5 Total:** 28 features, ~42 tests

---

## Phase 4.3: TABLE PROPERTIES (Weeks 6-7) - 31 Features

### Table-Level Properties (Week 6)

**4.3.1 Table Position**
- Add `position?: { x, y, hAlign, vAlign, ... }`
- Floating table positioning
- Tests: 8

**4.3.2-4.3.7: Table Properties**
- Table overlap, bidirectional, table grid
- Caption, description, preferred width type
- Tests: 12

**4.3.8-4.3.15: Cell Properties**
- Text direction, fit text, no wrap
- Hidden, conditional formatting
- Tests: 10

### Row Properties (Week 7)

**4.3.16: Cell Margins** (Already in 4.0.3)

**4.3.17-4.3.31: Row Properties**
- Row borders, grid before/after
- Width before/after, hidden, conditional
- Tests: 20

**Week 6-7 Total:** 31 features, ~50 tests

---

## Phase 4.4: IMAGE PROPERTIES (Week 8) - 8 Features

**4.4.1 Text Wrapping**
- Add `wrapping?: { type, wrapText, distanceFromText }`
- Serialize as `<wp:wrapSquare/>`, `<wp:wrapTight/>`, etc.
- Tests: 8

**4.4.2 Positioning**
- Add `positioning?: { horizontal, vertical, offset }`
- Absolute/relative to page/margin
- Tests: 10

**4.4.3-4.4.8: Additional Properties**
- Alignment, distance from text, rotation
- Effects, cropping, alt text
- Tests: 18

**Week 8 Total:** 8 features, ~36 tests

---

## Phase 4.5: SECTION PROPERTIES (Week 9) - 15 Features

**4.5.1 Line Numbering**
- Add `lineNumbering?: { countBy, start, distance, restart }`
- Tests: 6

**4.5.2-4.5.15: Additional Section Properties**
- Paper source, text direction, vertical alignment
- Bidirectional, RTL gutter, document grid
- No endnote, form protection
- Tests: 18

**Week 9 Total:** 15 features, ~24 tests

---

## Phase 4.6: FIELD TYPES (Week 10) - 11 Features

**4.6.1-4.6.3: Page Fields**
- createPageField(), createNumPagesField(), createDateField()
- Tests: 6

**4.6.4-4.6.6: Reference Fields**
- createRefField(), createHyperlinkField(), createSeqField()
- Tests: 9

**4.6.7-4.6.11: Document Property Fields**
- AUTHOR, TITLE, FILENAME, TC, XE, INDEX
- Tests: 10

**Week 10 Total:** 11 features, ~25 tests

---

## Phase 5: ADVANCED FEATURES (Weeks 11-16)

### Phase 5.1: TABLE STYLES (Week 11) - 4 Features

**5.1.1-5.1.3: Table Formatting in Styles**
- Add tableFormatting, tableCellFormatting, tableRowFormatting to StyleProperties
- Update Style.toXML() and parser
- Tests: 16

**5.1.4: Conditional Table Formatting**
- Add conditionalFormatting to StyleProperties
- Support firstRow, lastRow, banding
- Tests: 12

**Week 11 Total:** 4 features, ~28 tests

---

### Phase 5.2: CONTENT CONTROLS (Week 12) - 9 Features

**5.2.1-5.2.9: All Control Types**
- Rich text, plain text, combo box, dropdown
- Date picker, checkbox, picture, building block, group
- Update StructuredDocumentTag class
- Tests: 18

**Week 12 Total:** 9 features, ~18 tests

---

### Phase 5.3: STYLE ENHANCEMENTS (Week 13) - 9 Features

**5.3.1-5.3.9: Style Gallery & Metadata**
- Quick style, UI priority, semi-hidden
- Unhide when used, locked, personal styles
- Linked style, auto-redefine
- Tests: 12

**Week 13 Total:** 9 features, ~12 tests

---

### Phase 5.4: DRAWING ELEMENTS (Weeks 14-15) - 5 Features

**5.4.1 Shapes**
- Create Shape.ts class
- Support rectangles, circles, arrows, lines
- Tests: 10

**5.4.2 Text Boxes**
- Create TextBox.ts class
- Floating text boxes with content
- Tests: 8

**5.4.3-5.4.5: Advanced Drawing**
- SmartArt, Charts, WordArt (basic support)
- Tests: 15

**Week 14-15 Total:** 5 features, ~33 tests

---

### Phase 5.5: DOCUMENT PROPERTIES (Week 16) - 8 Features

**5.5.1-5.5.8: Extended Properties**
- Custom properties, category, content status
- Language, version, application, company, manager
- Update DocumentProperties interface
- Tests: 12

**Week 16 Total:** 8 features, ~12 tests

---

## Success Metrics

### Phase 4 Complete (Week 10)
- ✅ 96 features implemented (76% of total)
- ✅ ~430 tests passing (66% increase from 226)
- ✅ Feature parity with docx library for common use cases
- ✅ No data loss on load/save for Phase 4 features

### Phase 5 Complete (Week 16)
- ✅ 127 features implemented (100%)
- ✅ ~650 tests passing (186% increase from 226)
- ✅ Full ECMA-376 compliance for implemented features
- ✅ Production-ready v1.0.0 release

---

## Version Release Schedule

| Phase | Version | Release Date | Features | Tests |
|-------|---------|--------------|----------|-------|
| 4.0 (Fixes) | v0.32.0 | Week 1 end | 3 critical fixes | +15 tests |
| 4.1 (Run) | v0.33.0 | Week 3 end | 22 run properties | +35 tests |
| 4.2 (Para) | v0.34.0 | Week 5 end | 28 paragraph properties | +42 tests |
| 4.3 (Table) | v0.35.0 | Week 7 end | 31 table properties | +50 tests |
| 4.4 (Image) | v0.36.0 | Week 8 end | 8 image properties | +36 tests |
| 4.5 (Section) | v0.37.0 | Week 9 end | 15 section properties | +24 tests |
| 4.6 (Fields) | v0.38.0 | Week 10 end | 11 field types | +25 tests |
| 5.1 (Table Styles) | v0.39.0 | Week 11 end | 4 table style features | +28 tests |
| 5.2 (Controls) | v0.40.0 | Week 12 end | 9 content controls | +18 tests |
| 5.3 (Style+) | v0.41.0 | Week 13 end | 9 style enhancements | +12 tests |
| 5.4 (Drawing) | v0.42.0 | Week 15 end | 5 drawing elements | +33 tests |
| 5.5 (Props) | v0.43.0 | Week 16 end | 8 doc properties | +12 tests |
| **1.0.0** | **v1.0.0** | Week 17 | **Full feature parity** | **~650 tests** |

---

## Next Implementation: Phase 4.0.2

**Immediate Focus:** Fix image parsing (parseDrawingFromObject)
**Priority:** CRITICAL - Complete data loss
**Estimated Time:** 2 days
**Files:** DocumentParser.ts:751-760

Starting implementation now...
