# Changelog

All notable changes to docxmlater will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [9.8.8] - 2026-02-19

### Added

- **ECMA-376 Image Attribute Coverage**
  - Group A: 14 typed properties (presetGeometry, compressionState, bwMode, inline distances, noChangeAspect, hidden, blipFillDpi, blipFillRotWithShape, picLocks, picNonVisualProps, isLinked, transparency)
  - Group B: 9 raw XML passthrough slots via `_rawPassthrough` Map for complex DrawingML subtrees; `collectUnmodeledChildren()` helper in parser
  - Group C: Enhanced `ImageBorder` interface replacing `{ width: number }` with full `a:ln` support (cap, compound, alignment, fill with srgbClr/schemeClr modifiers, dashPattern, join, headEnd, tailEnd)
  - Group D: SVG/EMF/WMF format support with buffer-based detection, dimension detection, MIME types, `xmlns:asvg` namespace, and linked images via `r:link`
  - 69 tests in ImageProperties.test.ts

- **ECMA-376 Gap A: CJK Paragraph Properties**
  - `setKinsoku()`, `setWordWrap()`, `setOverflowPunct()`, `setTopLinePunct()`, `setAutoSpaceDE()`, `setAutoSpaceDN()` on Paragraph
  - Parsed during document load and round-tripped on save

- **ECMA-376 Gap A: Document Background**
  - `getDocumentBackground()` / `setDocumentBackground()` on Document
  - Supports color, themeColor, themeTint, themeShade per ECMA-376 Part 1 Section 17.2.1

- **ECMA-376 Gap B: Form Field Data Preservation**
  - `w:fldChar` elements with `w:ffData` now parsed and round-tripped
  - Covers text input, checkbox, and dropdown form fields per ECMA-376 Section 17.16
  - `FormFieldData` type with name, enabled, calcOnExit, textInput, checkBox, ddList

- **ECMA-376 Gap C: Expanded Document Settings**
  - `getEvenAndOddHeaders()` / `setEvenAndOddHeaders()` for different odd/even headers
  - `getMirrorMargins()` / `setMirrorMargins()` for binding margins
  - `getAutoHyphenation()` / `setAutoHyphenation()` for automatic hyphenation
  - `getDecimalSymbol()` and `getListSeparator()` accessors

- **w14 Run Effects Passthrough**
  - `addRawW14Property(rawXml: string)` on Run for Word 2010+ effects (ligatures, numForm, textOutline, shadow, reflection, glow, etc.)
  - Stored in `formatting.rawW14Properties` array; output in `w:rPr` during generation

- **Tracked Changes: Full Property Round-Trip**
  - Run property changes (`w:rPrChange`): 30+ properties including dstrike, shadow, outline, emboss, imprint, vanish, RTL, complex scripts, spacing, kerning, border, shading, East Asian layout
  - Paragraph property changes (`w:pPrChange`): 25+ properties including widowControl, contextualSpacing, outlineLevel, bidi, borders, shading, tabs, textDirection
  - New types: `ParagraphBorderDef`, `ParagraphBorders`, `ParagraphShading`, `TabStopDef`, `ParagraphFormattingPartial`

- **Granular Character-Level Tracked Changes**
  - `diffText(oldText, newText): DiffSegment[]` prefix/suffix diff algorithm for minimal edits
  - `diffHasUnchangedParts(segments): boolean` for detecting partial text changes
  - Integration with `Run.setText()` when change tracking is enabled

- **Bookmark System: Body-Level Bookmarks**
  - Support for bookmarks appearing between block elements (paragraphs/tables) in document body
  - `extractBodyLevelBookmarkEnds()` and `extractBodyLevelBookmarkStarts()` in DocumentParser

- **Bookmark Validation and Auto-Repair**
  - `validateBookmarkPairs(): number` on Document returns count of repairs
  - Adds missing `w:bookmarkEnd` markers for orphaned starts
  - Removes orphaned `w:bookmarkEnd` without matching starts
  - Automatically called during `save()` / `toBuffer()`

- **Footnote/Endnote Round-Trip Pipeline**
  - Full round-trip: parse on load, regenerate on save, passthrough if unmodified
  - Dirty flag tracking (`_footnotesModified`, `_endnotesModified`)
  - Public API: `createFootnote()`, `createEndnote()`, `clearFootnotes()`, `clearEndnotes()`, `getFootnoteManager()`, `getEndnoteManager()`
  - Duplicate relationship prevention on round-trip

- **Unified Shading Model (ECMA-376 Closure)**
  - `ShadingConfig` type with 9 attributes: pattern, fill, color, themeFill, themeColor, themeFillTint, themeFillShade, themeTint, themeShade
  - `buildShadingAttributes()` shared XML generation helper
  - `parseShadingFromObj()` unified parser across 7 parse sites
  - Type aliases: `CharacterShading`, `TableShading`, `CellShading`, `ShadingProperties`, `ParagraphShading` all unified as `ShadingConfig`
  - Paragraph `setShading()` now integrates with change tracking
  - `resolveCellShading()` inheritance resolver (direct cell, tblPrEx, conditional style, default cell, table shading)
  - `decodeCnfStyle()` / `getActiveConditionalsInPriorityOrder()` utilities
  - `Document.getComputedCellShading(table, row, col)` convenience method

- **Run Enhancements**
  - `setUnderlineColor(color)` and `setUnderlineThemeColor(themeColor, themeTint?, themeShade?)` for underline styling
  - Theme font references: `setFontAsciiTheme()`, `setFontHAnsiTheme()`, `setFontEastAsiaTheme()`, `setFontCsTheme()`

- **Table Enhancements**
  - `getFirstParagraph()` / `getLastParagraph()` navigation helpers
  - Legacy horizontal merge (`hMerge`) via `setHorizontalMerge()` / `getHorizontalMerge()` on TableCell
  - `getLayout()` on Table returns `fixed` or `auto`

- **TOC Relative Indentation**
  - `indentPerLevel` option in `applyTocStyles()` for relative indentation calculation
  - First visible level gets 0 indent; each subsequent level adds `indentPerLevel` twips

- **Document Fidelity**
  - App.xml metadata preservation during round-trip (HeadingPairs, TotalTime, Pages, Words, etc.)
  - Namespace order preservation in generated XML
  - Style default attribute preservation (`w:default="1"`) with `setIsDefault()` / `getIsDefault()` on Style
  - Numbering definition preservation with Word 2013+ attributes (`w15:restartNumberingAfterBreak`, `w16cid:durableId`)
  - Style property preservation through load/modify/save cycles

- **Infrastructure**
  - webSettings.xml auto-generation with `optimizeForBrowser` and `allowPNG`
  - People.xml auto-registration for tracked change authors during save
  - `Paragraph.removeBookmarkEnd(id)` method

### Fixed

- **Border Parsing**: `space="0"` and `sz="0"` attributes on borders now correctly parsed as zero instead of being ignored as falsy
- **Duplicate Relationship Prevention**: `saveComments()` now checks before adding relationship, preventing Word freeze with duplicate comments.xml entries
- **Singleton Relationship Guards**: Footnotes, endnotes, and people.xml relationships now have check-before-add guards preventing duplicates on round-trip

### Changed

- **Breaking**: `val` removed from all shading interfaces; use `pattern` everywhere. All shading types unified as `ShadingConfig`

### Statistics

- 121 test suites, 2,683 tests passing
- 118 source files
- 24 new/modified test files added in this release

---

## [9.6.2] - 2026-02-05

### Added

- **Compatibility Mode Support (Phases 1-3)**
  - Settings.xml round-trip preservation with schema-aware merge on save
  - `getCompatibilityMode()` returns the document's Word version mode (11/12/14/15)
  - `isCompatibilityMode()` checks if document targets a legacy Word version (pre-2013)
  - `getCompatibilityInfo()` returns full parsed compat settings including legacy flags and modern entries
  - `upgradeToModernFormat()` upgrades document to mode 15, equivalent to Word's File > Info > Convert
  - `CompatibilityUpgrader` utility class for w:compat block manipulation
  - `CompatibilityMode` enum, `CompatibilityInfo` and `CompatSetting` types
  - Complete catalog of 65 legacy compat elements (`LEGACY_COMPAT_ELEMENTS`)
  - 5 modern compat settings (`MODERN_COMPAT_SETTINGS`)
  - 46 tests in SettingsRoundTrip.test.ts

- **Preserved Element Types for Round-Trip Fidelity**
  - `AlternateContent` preserves mc:AlternateContent blocks (Word 2010+ shapes/VML fallback)
  - `MathParagraph` / `MathExpression` preserves m:oMathPara and m:oMath elements
  - `CustomXmlBlock` preserves w:customXml block elements
  - `PreservedElement` generic wrapper for w:proofErr, w:permStart, w:permEnd, w:altChunk, w:ruby
  - All stored as raw XML for maximum fidelity; parsed at both body-level and inline contexts

- **Expanded Run Content Types**
  - 14 new run content types: lastRenderedPageBreak, separator, continuationSeparator, pageNumber, annotationRef, dayShort, dayLong, monthShort, monthLong, yearShort, yearLong, symbol, positionTab, embeddedObject
  - Symbol support with font and character code (w:sym)
  - Position tab support with alignment, relativeTo, and leader attributes (w:ptab)
  - Embedded OLE object preservation (w:object)

- **Run Formatting Properties**
  - `complexScript` (w:cs) per ECMA-376 Part 1 section 17.3.2.7
  - `webHidden` (w:webHidden) for web layout hidden text
  - `fontHAnsi`, `fontEastAsia`, `fontCs`, `fontHint` for multi-font family support per w:rFonts

- **Section Property Enhancements**
  - Expanded `PageNumberFormat` type with 40+ formats per ECMA-376 ST_NumberFormat
  - Footnote/endnote section-level properties (`NotePosition`, `NoteNumberRestart`)
  - Chapter separator type (`ChapterSeparator`)
  - 6 tests in SectionPropertiesParsing.test.ts

- **Table Property Enhancements**
  - `tblStyleRowBandSize` / `tblStyleColBandSize` for row/column band sizing in style alternation
  - `setStyleRowBandSize()` / `setStyleColBandSize()` methods

- **TableRow Property Enhancements**
  - `wBefore` / `wAfter` for width before/after row per ECMA-376
  - `cellSpacing` for row-level cell spacing override
  - `cnfStyle` for conditional formatting bitmask

- **Date Formatting Utility**
  - `formatDateForXml()` strips milliseconds from ISO dates for OOXML w:date compliance

### Fixed

- **mc:Ignorable Auto-Generation**
  - XMLBuilder.createDocument() now auto-generates mc:Ignorable attribute when extended namespaces (w14, w15, wp14, w16se, w16cid, etc.) are declared
  - Prevents Word "unreadable content" error caused by w14:paraId in raw XML passthrough zones
  - 15 regression tests in CorruptionFixes.test.ts

- **Orphaned numId References**
  - Numbered paragraphs pointing to removed numbering definitions are now detected and cleaned

- **People.xml Tracked Change Authors**
  - Missing tracked change authors are now added to people.xml during save

- **pPrChange Attribute Order**
  - Paragraph property change elements now use correct attribute ordering per ECMA-376

### Changed

- **index.ts Export Reorganization**
  - Exports reorganized with clear section headers for better discoverability
  - All new types, classes, and utilities properly exported

- **Dependency Cleanup**
  - Removed unused `all` and `tailwindcss` dependencies

### Removed

- **BaseManager Abstract Class** (`src/core/BaseManager.ts`)
  - Removed unused abstract base class; managers operate independently
  - No external API impact

- **Document.normalizeTableLists() Method**
  - List normalization moved to consumer applications
  - `ListNormalizationOptions` and `ListNormalizationReport` no longer exported from Document

---

## [9.5.33] - 2026-01-28

### Fixed

- Repository cleanup and bug fixes

---

## [9.5.31] - 2026-01-28

### Added

- **Revision Consolidation**
  - New `Document.consolidateAllRevisions()` method merges adjacent same-author revisions within a time window
  - New `Paragraph.consolidateRevisions()` for paragraph-level consolidation
  - Automatically called during `save()` when track changes is enabled
  - Prevents "random insertions and deletions" appearance in Word by matching Word's consolidation behavior

- **Section Property Getters**
  - Added individual getters: `getPageSize()`, `getOrientation()`, `getMargins()`, `getColumns()`, `getSectionType()`, `getPageNumbering()`, `getTitlePage()`, `getHeaderReferences()`, `getFooterReferences()`, `getVerticalAlignment()`, `getPaperSource()`, `getColumnSeparator()`, `getColumnWidths()`, `getTextDirection()`, `getBidi()`, `getRtlGutter()`, `getDocGrid()`

- **Table Property Getters**
  - Added: `getTblLook()`, `getPosition()`, `getOverlap()`, `getBidiVisual()`, `getTableGrid()`, `getCaption()`, `getDescription()`, `getCellSpacingType()`

- **TableCell Property Getters**
  - Added: `getFitText()`, `getNoWrap()`, `getHideMark()`, `getCnfStyle()`

- **Table Merge Conflict Detection**
  - New `Table.canMergeCells()` method checks for merge conflicts before attempting merge
  - `Table.mergeCells()` now throws descriptive errors when cells are already part of merge regions

- **Bookmark Raw Name Setter**
  - New `Bookmark.setRawName()` method sets bookmark name without normalization for round-trip fidelity

- **StylesManager Circular Reference Detection**
  - New `hasCircularReference()` method for explicit cycle detection
  - Improved `validateStyleReferences()` with better cycle path reporting

### Changed

- **Enhanced dispose() Method**
  - Now uses `clear()` methods on managers instead of creating new instances
  - Clears ZIP handler, tracking context, and preserved state
  - More thorough memory cleanup for API server use cases

- **TableCell.setVerticalMerge()**
  - Now accepts `undefined` to clear vertical merge
  - `Table.splitCell()` now clears both column span and vertical merge

### Fixed

- **Save State Rollback on Failure**
  - Document now creates state snapshot before save operations
  - On save failure, preserved paragraph flags and revision state are restored
  - Image data is only released after successful save, allowing retry without reload

- **Header/Footer Parse Error Handling**
  - Header/footer parsing errors no longer crash document loading
  - Document loads with warning when headers/footers cannot be parsed

- **List Normalizer Whitespace Trimming**
  - Fixed edge case where whitespace-only runs after prefix stripping didn't continue trimming

---

## [9.5.28] - 2026-01-27

### Fixed

- **Orphan List Level Normalization**
  - Fixed early exit path that skipped orphan level normalization when `recommendedAction === "none"`
  - Word lists at L1/L2 are now correctly shifted to L0/L1 when no L0 parent exists
  - Applies to all cells regardless of whether typed prefix conversion is needed

- **Format-Based Levels for Typed Prefixes**
  - Typed prefixes without explicit indentation now use semantic levels from `FORMAT_TO_LEVEL`
  - `decimal` (1., 2., 3.) maps to Level 0
  - `lowerLetter` (a., b., c.) maps to Level 1
  - `lowerRoman` (i., ii., iii.) maps to Level 2
  - Previously all typed prefixes defaulted to Level 0 when no extra indentation was present
  - Preserves semantic list hierarchy: "1." at L0, "a)" at L1, "i." at L2

---

## [9.5.20] - 2026-01-26

### Added

- **Relative Indentation Level Inference**
  - New `inferLevelFromRelativeIndentation()` function for table cell list normalization
  - Calculates list level based on paragraph indentation relative to cell baseline
  - Handles variable baseline indentation per cell (different from document-level lists)

### Changed

- **List Normalization Indentation-Based Levels**
  - Typed prefix lists now use indentation-based level detection instead of format-based
  - Calculates baseline (minimum) indentation within each cell
  - Sub-items are identified by relative indentation, not prefix format (a. vs 1.)
  - Numbered items following bullets are promoted to level 1 when indented beyond baseline

### Fixed

- **Mixed List Category Handling**
  - Improved detection of numbered items that should be sub-items under bullet lists
  - Added `numberedAsSubItemIndices` tracking for proper level assignment
  - Prevents flat list appearance when numbered items follow bullet parents

---

## [9.5.19] - 2026-01-26

### Fixed

- **Header/Footer Image Parsing**
  - Images in headers and footers now use their own relationship files (header1.xml.rels, footer1.xml.rels)
  - Previously, header/footer images incorrectly used document.xml.rels, causing image relationship mismatches
  - Added `currentPartName` tracking in DocumentParser to distinguish image sources
  - ImageManager now uses composite keys (e.g., `header1.xml:rId1`) to prevent relationship ID collisions

### Added

- **Public API Methods for Image Size Detection**
  - `isImageSmall(image)` - Check if image is small (< 100x100 pixels)
  - `isSmallImageParagraph(para)` - Check if paragraph contains a small image
  - These were previously private; now available for external use

### Changed

- **List Spacing Logic Improvement**
  - No longer adds blank line after list items when followed by indented content (e.g., "Example:" notes)
  - Improves layout for documents with continuation content after lists

---

## [9.5.13] - 2026-01-26

### Added

- **Centered Bold Text Blank Line Handling (Phases 11c/11d)**
  - Added `isCenteredBoldText()` helper to detect paragraphs that are centered with all text runs bold
  - Added `isTextOnlyParagraph()` helper to detect paragraphs containing only text (no images/shapes)
  - No blank line between centered bold text and following image (caption pattern)
  - Add blank line above centered bold text when preceded by text-only paragraph or list item
  - Applies to both body-level paragraphs and table cell paragraphs

---

## [9.5.12] - 2026-01-26

### Fixed

- **Style numPr Preservation**
  - Styles now preserve `w:numPr` (numbering properties) through round-trip
  - Fixes issue where paragraphs using `ListParagraph` style lost their numbered prefix after processing
  - Added `numPr` property to `Style` class for storing numId and ilvl
  - Added numPr parsing in `DocumentParser.parseStyle()` method
  - Added numPr generation in `Style.toXML()` method

---

## [9.5.11] - 2026-01-26

### Added

- **Blank Line After Long Text Paragraphs (Phase 9d)**
  - Adds blank line after paragraphs with >100 characters of text
  - Conditions: current and next elements must NOT be list items, NOT be blank, NOT be indented
  - Never adds blank if current element is last in table cell
  - Applies to both body-level paragraphs and table cell paragraphs
  - Helps visually separate dense content from following paragraphs

---

## [9.5.10] - 2025-01-26

### Fixed

- **No Blank Line Between Consecutive Small Images**
  - Don't add blank line after small image (<100x100) if next element is also a small image
  - Applies to both body-level and table cell images
  - Improves layout for documents with multiple inline icons or small graphics

---

## [9.5.9] - 2025-01-26

### Added

- **Indented Paragraph Handling in Table Cells (Phase 9c-table)**
  - Added blank line handling for indented-to-non-indented transitions inside table cells
  - Never adds blank line if element is last in table cell
  - All blank lines use "Normal" style (already the default)

---

## [9.5.8] - 2025-01-26

### Fixed

- **Blank Line Handling for Indented Paragraphs After Lists**
  - No longer adds unwanted blank line between list items and following indented paragraphs
  - Now adds blank line after indented paragraph blocks when transitioning to non-indented content
  - Affects `addStructureBlankLines()` Phase 9 and new Phase 9c
  - Example: List item followed by indented email text no longer has blank before, but does have blank after

---

## [9.5.7] - 2025-01-25

### Fixed

- **List Prefix Formatting Normalization**
  - List prefixes (bullets, numbers, letters) are now cleared of bold, italic, and underline when parsing documents
  - Previously, only bold was cleared; now italic and underline are also cleared
  - Ensures consistent, clean list prefix appearance across all processed documents
  - Affects `NumberingLevel.fromXML()` parsing behavior

---

## [9.5.6] - 2025-01-25

### Reverted

- **Reverted to 9.4.0 codebase**
  - All 9.5.x versions have been reverted due to stability issues
  - Reverted changes:
    - 9.5.5: Revert to 9.5.1 (partial revert)
    - 9.5.4: Bookmark-only paragraph preservation
    - 9.5.3: Bookmarks inside body-level deleted content handling
    - 9.5.2: Inter-row and body-level bookmark preservation
    - 9.5.1: Bold+colon blank line logic for indented paragraphs
    - 9.5.0: Granular numbering modification tracking for style preservation
  - The numbering tracking and bookmark handling approaches need redesign
  - This release restores the stable 9.4.0 feature set

---

## [9.4.0] - 2025-01-24

### Added

- **Table Cell Margins Support**
  - New `TableCellMargins` interface with top/bottom/left/right properties (in twips)
  - `Table.getCellMargins()` - retrieve current cell margins
  - `Table.setCellMargins(margins)` - set default cell margins for all cells
  - Full round-trip support: parsing and serialization of `w:tblCellMar`
  - Per ECMA-376 Part 1 Section 17.4.42

- **HLP Hyperlinks Styling** (`styleDocument()` enhancement)
  - New `aboveReturnToHLP` option to add blank lines above "Return to HLP" hyperlinks
  - Auto-creates `HLPHyperlinks` style with right alignment, Verdana 12pt, blue underline
  - Applies only to hyperlinks with exact text "return to hlp" (case-insensitive)

---

## [9.3.3] - 2025-01-23

### Fixed

- **Nested Bullets in Numbered Lists Flattening to Level 0**
  - Bullet points nested under numbered items are now converted to lettered sub-items (a, b, c) at Level 1
  - Previously, bullets were flattened to Level 0 and continued the numbering sequence (1, 2, 3, 4, 5, 6)
  - Now preserves hierarchy: `1. → • → • → 2.` becomes `1. → a. → b. → 2.`
  - Only affects cells where numbered items are the majority; bullet-only cells unchanged

---

## [9.3.2] - 2025-01-23

### Fixed

- **Hyperlinks in Tracked Changes Losing Styling**
  - Hyperlinks inside revisions (`w:ins`) no longer lose their blue color and underline after processing
  - Style application methods (`applyStandardTableFormatting()`, `applyH1()`, `applyH2()`, `applyH3()`, `applyListParagraph()`, `applyNormal()`) now skip runs with Hyperlink character style
  - Previously, explicit `w:color w:val="000000"` and `w:u w:val="none"` were being added, overriding the inherited Hyperlink style
  - Uses existing `Run.isHyperlinkStyled()` method to detect and preserve hyperlink formatting

---

## [9.3.1] - 2025-01-23

### Fixed

- **Empty/Self-Closing Hyperlinks**
  - Self-closing hyperlinks with no display text (e.g., `<w:hyperlink r:id="rId50"/>`) are now correctly skipped during parsing
  - Previously, the URL was incorrectly used as display text, causing URLs to appear as visible text in processed documents
  - This fix prevents invisible hyperlink markers from becoming visible text

### Added

- **Hyperlink Properties**
  - `isEmpty()` - Check if hyperlink is empty/invisible
  - `getTgtFrame()` - Get target frame attribute (e.g., "_blank")
  - `getHistory()` - Get history tracking attribute
  - `tgtFrame` and `history` attributes now preserved in XML output

---

## [9.3.0] - 2025-01-22

### Added

- **TableRow.clearHeight()**
  - New method to clear row height properties, allowing Word to auto-size rows based on content
  - Removes both `height` and `heightRule` properties
  - Supports method chaining

- **NumberingLevel Italic Support**
  - `setItalic(italic)` - Set italic formatting for numbering/bullet text
  - `getItalic()` - Get italic state (defaults to false)
  - New `italic?: boolean` property in NumberingLevelProperties interface
  - Generates `<w:i/>` and `<w:iCs/>` per ECMA-376

- **NumberingLevel Underline Support**
  - `setUnderline(style)` - Set underline style for numbering/bullet text
  - `getUnderline()` - Get underline style
  - `clearUnderline()` - Remove underline formatting
  - New `underline?: string` property supporting all Word underline styles (single, double, wave, dotted, dash, etc.)
  - Generates `<w:u w:val="..."/>` per ECMA-376

### Tests

- 3 new tests for TableRow.clearHeight()
- 18 new tests for NumberingLevel italic/underline (property handling, XML generation, XML parsing)

---

## [9.0.0] - 2025-01-22

### Added

- **Run Complex Script Font Size (w:szCs)**
  - `setSizeCs(size)` - Set font size for complex scripts (RTL text like Arabic, Hebrew)
  - `getSizeCs()` - Get complex script font size
  - New `sizeCs` property in RunFormatting interface
  - Per ECMA-376 Part 1 Section 17.3.2.40

- **Run Theme Color Support**
  - `setThemeColor(themeColor)` - Set color from document theme
  - `setThemeTint(tint)` - Apply tint (0-255, toward white)
  - `setThemeShade(shade)` - Apply shade (0-255, toward black)
  - New properties: `themeColor`, `themeTint`, `themeShade`
  - `ThemeColorValue` type exported with 16 standard theme colors

- **Revision Field Context Tracking**
  - `getFieldContext()` / `setFieldContext()` - Track revision position in fields
  - `isInsideField()` / `isInsideFieldResult()` / `isInsideFieldInstruction()`
  - New `FieldContext` interface with `position` and `instruction` properties

- **Table Cell Revision Support**
  - New revision types: `tableCellInsert`, `tableCellDelete`, `tableCellMerge`
  - `setCellRevision()` / `getCellRevision()` / `hasCellRevision()` on TableCell

- **NumberingLevel Restart Support (w:lvlRestart)**
  - `getLvlRestart()` / `setLvlRestart(level)` - Control which level restarts
  - Per ECMA-376 Part 1 Section 17.9.11
  - Useful for legal documents with continuous sub-numbering

### Fixed

- **Hyperlink Defragmentation with Track Changes**
  - `defragmentHyperlinks()` now guards against track changes conflicts
  - Prevents field corruption when track changes is enabled
  - Logs clear warning message when skipped

- **InstructionText Preservation in Run.setText()**
  - `setText()` now preserves `instructionText` content type for field instructions
  - Prevents field codes from displaying as visible text

### Tests

- 70+ new test cases across 7 new test files
- RunComplexScriptSize.test.ts (12 tests)
- RunThemeColor.test.ts (18 tests)
- RunInstructionText.test.ts (7 tests)
- RevisionFieldContext.test.ts (10+ tests)
- TableCellRevision.test.ts (10+ tests)
- NumberingLevelRestart.test.ts (20 tests)
- HyperlinkDefragmentWithRevisions.test.ts (regression tests)

---

## [8.1.0] - 2025-12-25

### Added

- **Document Search & Query Methods**
  - `findParagraphsByText(pattern)` - Search paragraphs by text or regex pattern
  - `getRunsByFont(fontName)` - Find all runs using a specific font
  - `getRunsByColor(color)` - Find all runs with a specific color
  - `getParagraphsByStyle(styleId)` - Get paragraphs with a specific style

- **Document Bulk Formatting Methods**
  - `setAllRunsFont(fontName)` - Apply font to all text in document
  - `setAllRunsSize(size)` - Apply font size to all text in document
  - `setAllRunsColor(color)` - Apply color to all text in document
  - `getFormattingReport()` - Get comprehensive formatting statistics

- **Document Convenience Methods**
  - `setAuthor(author)` - Alias for setCreator

- **Section Line Numbering** (ECMA-376 w:lnNumType)
  - `setLineNumbering(options)` - Enable line numbering with customizable options
  - `getLineNumbering()` - Get current line numbering settings
  - `clearLineNumbering()` - Remove line numbering
  - Supports countBy, start, distance, and restart options
  - Exported `LineNumbering` and `LineNumberingRestart` types

- **Comment Resolution** (ECMA-376 w:done attribute)
  - `Comment.resolve()` - Mark comment as resolved
  - `Comment.unresolve()` - Mark comment as unresolved
  - `Comment.isResolved()` - Check resolution status
  - `CommentManager.getResolvedComments()` - Get all resolved comments
  - `CommentManager.getUnresolvedComments()` - Get all unresolved comments
  - Updated `getStats()` to include resolved/unresolved counts

- **TableCell Convenience Methods**
  - `setTextAlignment(alignment)` - Set alignment for all paragraphs in cell
  - `setAllParagraphsStyle(styleId)` - Apply style to all paragraphs in cell
  - `setAllRunsFont(fontName)` - Apply font to all runs in cell
  - `setAllRunsSize(size)` - Apply font size to all runs in cell
  - `setAllRunsColor(color)` - Apply color to all runs in cell

- **Table Sorting**
  - `sortRows(columnIndex, options?)` - Sort table rows by column content
  - Options: ascending/descending, numeric/string comparison, skip header row

- **Style Methods**
  - `Style.reset()` - Reset style to minimal state (keeps id, name, type, basedOn)

### Changed

- Updated `API_METHODS_INVENTORY.md` with comprehensive documentation of all new methods

---

## [8.0.0] - 2025-12-24

### Added

- **Parsing Helpers (`src/utils/parsingHelpers.ts`)**: New utility functions for safe OOXML attribute parsing
  - `safeParseInt()` - Integer parsing with NaN handling and default values
  - `parseOoxmlBoolean()` - OOXML boolean parsing per ECMA-376 spec (handles self-closing tags, val="1", val="true")
  - `isExplicitlySet()` - Zero-value safe existence checking
  - `parseNumericAttribute()` - Numeric attribute parsing with zero-value handling
  - `parseOnOffAttribute()` - ST_OnOff type attribute parsing
  - All helpers exported from main index.ts

- **Run Property Change Parsing**: DocumentParser now parses `w:rPrChange` elements
  - Previous run formatting properties are preserved in Revision objects
  - Enables changelog generation for formatting changes

### Fixed

- **Zero-Value Handling Bug**: Fixed multiple locations in DocumentParser where values of 0 were incorrectly treated as falsy
  - Affected: spacing values, indentation, table grid widths, frame properties
  - Solution: Use `isExplicitlySet()` instead of truthy checks

- **Boolean Property Parsing**: Unified boolean parsing across DocumentParser
  - Now correctly handles all OOXML formats: self-closing tags, val="1", val="true", val="on"
  - Uses `parseOoxmlBoolean()` helper consistently

- **Revision Serialization**: Internal tracking types now return null from `toXML()` instead of throwing
  - Affected types: hyperlinkChange, imageChange, fieldChange, commentChange, bookmarkChange, contentControlChange
  - Prevents document save failures when internal tracking types are present

### Changed

- **Revision.toXML() Return Type**: Changed from `XMLElement` to `XMLElement | null`
  - Internal tracking types now gracefully return null
  - Paragraph serialization updated to skip null revisions

### Documentation

- **Nested Tables**: Added comprehensive documentation for nested table handling
  - Design philosophy (raw XML passthrough for round-trip fidelity)
  - Limitations and ECMA-376 compliance notes

- **Parsing Helpers**
  - Zero-value bug pattern explanation
  - OOXML boolean parsing rules
  - Usage examples from DocumentParser

### Tests

- **Nested Table Tests**: Added deep nesting and edge case test scenarios
  - 5-level deep nesting verification
  - Nested tables with revision markers (w:ins, w:del)
  - Multiple nested tables at different positions in same cell
  - SDT containing nested table
  - Performance validation for large structures (10 tables, 50 rows)
  - Edge cases: minimal structure, merged cells, complex borders, hyperlinks/bookmarks

---

## [5.1.0] - [7.6.8] - 2025-11-19 to 2025-12-10

### Note

For detailed changes between v5.1.0 and v7.6.8, see:

- Git commit history: `git log v5.0.0..v7.6.8`
- GitHub releases: <https://github.com/ItMeDiaTech/docXMLater/releases>

---

## [5.0.0] - 2025-11-19

### Added
- **CleanupHelper Class**: Comprehensive document cleanup utilities including unlocking SDTs, removing preserve flags, defragmenting hyperlinks, cleaning unused elements, removing customXML, unlocking fields/frames, and sanitizing tables.

### Fixed

- **TOC Field Instruction \o Switch Format Support**: Enhanced TOC outline level switch to support unquoted, single-quoted, and double-quoted formats
  - **Previous Behavior**: Only supported double-quoted format like `\o "1-3"`
  - **Issue**: Documents from some generators (e.g., Google Docs) use unquoted format `\o 1-3` which wasn't recognized
  - **Solution**: Updated regex in [`parseTOCFieldInstruction()`](src/core/DocumentParser.ts) and [`Document.parseTOCFieldInstruction()`](src/core/Document.ts) to support multiple formats
  - **New Regex**: `/\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/` handles:
    - Double-quoted: `\o "1-3"` (original format)
    - Single-quoted: `\o '1-3'` (alternative quoted format)
    - Unquoted: `\o 1-3` (unquoted format from Google Docs)
  - **Implementation**: Uses grouped captures with fallback logic: `parseInt(match[1] || match[3] || match[5]!, 10)`
  - **Backward Compatibility**: Existing documents continue to work unchanged
  - **Test Coverage**: Added 10 regression tests covering all three formats and edge cases

- **TOC Field Instruction Extraction**: Fixed critical bug in Test_Code.docx and similar documents where TOC field instructions couldn't be parsed
  - **Root Cause**: Single run can contain multiple `w:fldChar` elements in an array (e.g., both `begin` and `separate` in same run)
  - **Previous Behavior**: Code assumed `runObj["w:fldChar"]` was always a single object, resulting in `undefined` when accessing `fldChar["@_w:fldCharType"]` on an array
  - **Impact**: TOC elements were recognized but field instructions couldn't be extracted, returning `null` for TOC properties
  - **Solution**: Updated [`extractInstructionFromRawXML()`](src/core/DocumentParser.ts:3190) to handle `w:fldChar` as either object or array
  - **Implementation**: Added array detection and iteration - `const fldCharArray = Array.isArray(fldChar) ? fldChar : [fldChar]`
  - **Enhanced Logging**: Added diagnostic logging showing field marker counts and extraction steps
  - **Multi-Paragraph Support**: Field tracking now spans multiple paragraphs (TOC fields can have begin/separate in paragraph 1, end in paragraph 5)
  - **Tested With**: Test_Code.docx successfully extracts `TOC \h \u \z \t "Heading 2,2,"` instruction


### Removed

- **`Document.cleanFormatting()` Method**: Removed overly aggressive formatting cleanup method
  - **Reason**: Too aggressive - destroyed intentional direct formatting (bold, colors, fonts)
  - **Reason**: Redundant - all `applyX()` methods already clear formatting conflicts internally
  - **Reason**: Context-blind - didn't distinguish between body paragraphs and table cells
  - **Issue**: Was removing formatting from Header 2 paragraphs in table cells
  - **Replacement**: Use `Paragraph.clearDirectFormattingConflicts(style)` for smart conflict detection
  - **Impact**: None - single internal usage in WordDocumentProcessor was redundant
  - **Note**: The safe utility function `cleanFormatting()` in `src/utils/formatting.ts` (removes null/undefined from objects) is unchanged

### Changed

- **WordDocumentProcessor**: Removed redundant `doc.cleanFormatting()` call (line 797)
  - Direct formatting conflicts already handled by `applyH1()`, `applyH2()`, `applyH3()`, etc.
  - Each method internally calls `clearDirectFormattingConflicts()` which preserves non-conflicting formatting

---

## [3.2.0] - [4.9.0] - 2025-01-17 to 2025-11-18

### Note

For detailed changes between v3.2.0 and v3.5.0, see:

- Git commit history: `git log v3.1.0..v3.5.0`
- GitHub releases: <https://github.com/ItMeDiaTech/docXMLater/releases>

---

## [3.1.0] - 2025-01-17

### Added

- **TOC Range Format Support**: Enhanced `\t` switch to support numeric range format
  - New range format: `\t "2-3"` similar to `\o` switch behavior
  - Supports patterns like `\t "2-2"` → [2], `\t "2-3"` → [2, 3], `\t "1-5"` → [1, 2, 3, 4, 5]
  - Maintains backward compatibility with style name format: `\t "Heading 2,2,"`
  - Parser detects range format via regex `/^(\d+)-(\d+)$/` before processing style names

### Fixed

- **TOC Field Instruction Parsing**: Fixed critical bug where TOCs with ONLY `\t` switches incorrectly fell back to default levels [1,2,3]
  - Root cause: `parseTOCFieldInstruction()` returned default [1,2,3] whenever `levels.size === 0`, regardless of whether switches were present
  - Issue: Field instruction `TOC \h \u \z \t "Heading 2,2,"` should ONLY include Heading 2 paragraphs, but incorrectly included Heading 1 as well
  - Solution: Track whether `\t`, `\o`, or `\u` switches were found during parsing
  - Now only uses default [1,2,3] when NO switches are present
  - Returns empty array when switches exist but resulted in empty levels
  - Added support for `\u` switch (use outline levels from paragraph formatting)

### Examples

- `TOC \t "Heading 2,2,"` → [2] (not [1,2,3])
- `TOC \h \u \z \t "Heading 2,2,"` → [2] (not [1,2,3])
- `TOC \o "1-3"` → [1,2,3]
- `TOC` → [1,2,3] (default when no switches)
- `TOC \t "2-3"` → [2, 3] (new range format)

---

## [2.2.0] - 2025-11-13

### Added

- **Blank Line Preservation for 1x1 Tables**: New `ensureBlankLinesAfter1x1Tables()` method to preserve blank lines after single-cell tables
  - Automatically detects 1x1 tables with specific properties (10pt height, no borders)
  - Ensures exactly one blank paragraph follows each qualifying table
  - Prevents excessive blank line accumulation on repeated saves
  - Maintains document formatting consistency

### Documentation

- Added comprehensive Documentation Hub implementation comparison
- Updated project documentation and analysis files

### Technical Improvements

- Enhanced table processing logic for better formatting preservation
- Improved blank paragraph detection and management
- Added safety checks to prevent blank line duplication

---

## [1.18.0] - 2025-11-13

### Added

- **Comprehensive Tracked Changes Support**: Full implementation of all OpenXML revision types
  - `Revision` class enhancements: Support for all revision types (insert, delete, formatting, numbering, section properties, table properties, table row, table cell)
  - `RevisionManager` improvements: Enhanced tracking and management of revisions across document elements
  - New example: `examples/10-track-changes/advanced-track-changes.ts` (546 lines) demonstrating comprehensive tracked changes usage
  - Document-level tracked changes API: `Document.enableTrackChanges()`, `Document.disableTrackChanges()`, `Document.isTrackChangesEnabled()`
  - Paragraph-level tracked changes: `Paragraph.trackInsertion()`, `Paragraph.trackDeletion()`, `Paragraph.trackFormatting()`
  - Full round-trip support for reading and writing tracked changes

### Fixed

- **Automatic Indentation Conflict Resolution**: Fixed issues with numbered paragraph indentation
  - Automatically resolves conflicts between paragraph indentation and numbering indentation
  - Prevents double-indentation issues that occur when both paragraph and numbering define indentation
  - Implements smart merging: numbering indentation takes priority, paragraph indentation adjusts relatively
  - Comprehensive analysis document: `LIST-INDENTATION-ANALYSIS.md` (356 lines) documenting the implementation
  - New test suite: `tests/elements/ParagraphNumberingIndent.test.ts` (246 tests) ensuring correct behavior

### Technical Improvements

- Enhanced `Document.ts` with 180+ lines of tracked changes functionality
- Enhanced `Paragraph.ts` with 61 lines of indentation conflict resolution logic
- Expanded `Revision.ts` with 395+ lines supporting all revision types
- Improved `RevisionManager.ts` with 111+ lines of revision management features
- Added comprehensive formatting module documentation (52 lines)

### Tests

- All 1180 tests passing (53 test suites)
- New test coverage for tracked changes functionality
- New test coverage for paragraph numbering indentation
- Test output files cleaned up and removed from git tracking

### Documentation

- Added comprehensive list indentation analysis document
- Updated formatting module documentation with detailed specifications
- Added advanced tracked changes example with real-world scenarios

---

## [1.17.0] - 2025-11-13

Internal release with infrastructure improvements.

---

## [1.16.0] - 2025-11-13

### Documentation

- **Comprehensive Documentation Update**: Added complete documentation suite
  - New README.md with full feature matrix, API overview, and code examples
  - Updated project documentation to reflect all 5 phases complete (2073+ tests, 65 source files)
  - Added documentation consistency analysis (docs/analysis/)

### Added

- **Documentation Analysis Tools**:
  - `docs/analysis/DOCUMENTATION_CONSISTENCY_ANALYSIS.md` - 12-section analysis comparing implementation vs documentation
  - `docs/analysis/DOCUMENTATION_UPDATES_NEEDED.md` - Quick reference checklist for updates

### Changed

- **Phase Status Updates**: Marked Phase 4 (Rich Content) and Phase 5 (Polish) as Complete
  - Phase 4: Images, headers, footers, hyperlinks, bookmarks, shapes
  - Phase 5: Track changes, comments, TOC, fields, footnotes, content controls
- **Metrics Updates**:
  - Test count: 253 → 2073+ tests (59 test files)
  - Source files: 48 → 65 TypeScript files
  - Lines of code: ~10,000 → ~40,000+ lines

### Documentation Improvements

- Added comprehensive feature list covering all 31 element classes
- Added API overview with Document, Paragraph, Run, Table, TableCell classes
- Added code examples for common use cases
- Added TypeScript type examples
- Clarified RAG-CLI integration as development-only
- Added migration notes and performance considerations
- Added architecture overview and design principles

---

## [1.15.0] - 2025-11-14

### Added

- **Hyperlink Defragmentation API**: New methods to fix fragmented hyperlinks from Google Docs
  - `Document.defragmentHyperlinks(options)` - Merges fragmented hyperlinks with same URL across paragraphs
  - `Hyperlink.resetToStandardFormatting()` - Resets hyperlink to standard style (Calibri, blue, underline)
  - Enhanced `DocumentParser.mergeConsecutiveHyperlinks()` to handle non-consecutive fragments

### Improved

- **Hyperlink Merging Algorithm**: Now groups ALL hyperlinks by URL, not just consecutive ones
  - Handles hyperlinks separated by runs or other content
  - Optional formatting reset to fix corrupted fonts (e.g., Caveat from Google Docs)
  - Processes hyperlinks in both main content and tables

### Fixed

- **Hyperlink Fragmentation**: Fixed issue where hyperlinks with same URL were split into multiple fragments
- **Corrupted Hyperlink Fonts**: Added ability to reset hyperlinks to standard formatting
- **Non-Consecutive Hyperlink Merging**: Now properly merges hyperlinks even when separated by other content

### API Additions

```typescript
// Defragment hyperlinks in document
doc.defragmentHyperlinks({
  resetFormatting: boolean, // Reset to standard style
  cleanupRelationships: boolean, // Clean orphaned relationships
});

// Reset individual hyperlink formatting
hyperlink.resetToStandardFormatting();
```

### Technical Changes

- Enhanced `DocumentParser.mergeConsecutiveHyperlinks()` with URL grouping and optional formatting reset
- Added `getStandardHyperlinkFormatting()` helper in DocumentParser
- Added `resetToStandardFormatting()` method to Hyperlink class
- Added `defragmentHyperlinks()` public method to Document class

---

## [1.14.0] - 2025-11-13

### Added

- **New Helper Methods** for list formatting:
  - `NumberingLevel.getBulletSymbolWithFont(level, style)` - Get recommended bullet symbols with proper fonts for 5 different bullet styles (standard, circle, square, arrow, check)
  - `NumberingLevel.calculateStandardIndentation(level)` - Calculate standard Microsoft Word-compatible indentation values
  - `NumberingLevel.getStandardNumberFormat(level)` - Get recommended number format for any level (decimal, lowerLetter, lowerRoman, upperLetter, upperRoman)

### Changed

- **BREAKING (Minor)**: Default bullet font changed from 'Symbol' to 'Calibri' for better UI compatibility across platforms
- **List Indentation Formula**: Updated from `720 * (level + 1)` to `720 + (level * 360)` to match Microsoft Word standards
  - Level 0: 720 twips (0.5")
  - Level 1: 1080 twips (0.75") - was 1440 twips (1.0")
  - Level 2: 1440 twips (1.0") - was 2160 twips (1.5")
  - Level 3: 1800 twips (1.25") - was 2880 twips (2.0")
  - This fixes the issue where Level 3 appeared to have "double" indentation
- **Numbered List Formats**: Expanded from 3-level to 5-level format cycle
  - Level 0: decimal (1., 2., 3.)
  - Level 1: lowerLetter (a., b., c.)
  - Level 2: lowerRoman (i., ii., iii.)
  - Level 3: upperLetter (A., B., C.) - previously was decimal
  - Level 4: upperRoman (I., II., III.) - new
  - Level 5+: cycles back to decimal

### Fixed

- **Special Characters Serialization**: Tabs, newlines, and non-breaking hyphens now properly serialize as XML elements
  - `\t` (tab) → `<w:tab/>`
  - `\n` (newline) → `<w:br/>`
  - `\u2011` (non-breaking hyphen) → `<w:noBreakHyphen/>`
  - `\r` (carriage return) → `<w:cr/>`
  - `\u00AD` (soft hyphen) → `<w:softHyphen/>`
- **Run.getText()**: Now correctly reconstructs special characters from content elements
- **List Formatting**: Fixed Level 3+ numbered lists showing incorrect format (was "1., 2., 3." instead of "A., B., C.")
- **Bullet Display**: Improved bullet symbol display in UI contexts by using Calibri instead of Symbol font

### Technical Changes

- Added `parseTextWithSpecialCharacters()` private method to Run class for proper special character handling
- Updated Run constructor and setText() to use character parsing
- Enhanced Run.getText() to convert content elements back to their string representations
- Updated AbstractNumbering.createNumberedList() to support upperLetter and upperRoman formats
- Updated NumberingManager.getStandardIndentation() to use new formula

### Tests

- All 1188 tests passing (+21 from previous version)
- Added comprehensive test coverage for special character handling (19 tests)
- Updated test expectations for new list indentation and formatting behavior

## [1.13.0] - 2025-11-12

### Fixed

- **Hyperlink Duplication**: Fixed issue where hyperlinks from Google Docs would duplicate multiple times
  - Parse ALL runs within hyperlink elements, not just the first run
  - Added `mergeConsecutiveHyperlinks()` method to combine fragmented hyperlinks
  - Handles Google Docs-style hyperlinks that are split by formatting changes
- **Blank Paragraph Detection**: Enhanced logic to properly check for hyperlinks and other content before inserting blank paragraphs
  - Previously used `getText()` which only checked Run objects
  - Now uses `getContent()` to inspect all content types (Hyperlinks, Images, etc.)
- **List Indentation**: Fixed blank paragraph detection after Header 2 tables

### Added

- `Paragraph.clearContent()` method for removing all content from a paragraph

### Changed

- DocumentParser now correctly handles multi-run hyperlinks
- Enhanced blank paragraph insertion logic for better Word compatibility

## [1.12.0] - 2025-11-11

### Added

- Explicit spacing to Header 2 blank paragraphs (120 twips = 6pt) to ensure visibility in Word

### Fixed

- Blank paragraph spacing after Header 2 sections

## [1.11.0] - Previous Release

(See git history for earlier releases)

---

## Migration Guide

### Upgrading to 1.14.0

**List Indentation Changes:**
If you were relying on the specific indentation values, note that levels 1+ now have smaller indents:

- Old Level 1: 1440 twips → New Level 1: 1080 twips
- Old Level 2: 2160 twips → New Level 2: 1440 twips
- Old Level 3: 2880 twips → New Level 3: 1800 twips

To maintain old behavior, explicitly set indentation:

```typescript
const level = NumberingLevel.createBulletLevel(1);
level.setLeftIndent(1440); // Old value
```

**Bullet Font Changes:**
Bullets now use Calibri font by default instead of Symbol font. If you need Symbol font:

```typescript
const level = NumberingLevel.createBulletLevel(0, "•");
level.setFont("Symbol");
```

**Numbered List Formats:**
Level 3 now shows uppercase letters (A., B., C.) instead of numbers (1., 2., 3.). To maintain old behavior:

```typescript
const formats = ["decimal", "lowerLetter", "lowerRoman"]; // 3-level cycle
const abstractNum = AbstractNumbering.createNumberedList(1, 9, formats);
```

**Special Characters:**
Text containing tabs, newlines, etc. now automatically converts to proper XML elements. This is generally what you want, but if you need literal characters:

```typescript
// Tabs and newlines now auto-convert to XML elements
const run = new Run("Text\tWith\nSpecial");
// Generates: <w:t>Text</w:t><w:tab/><w:t>With</w:t><w:br/><w:t>Special</w:t>

// To preserve as literal text (not recommended):
const run = new Run("Text\\tWith\\nSpecial");
```
