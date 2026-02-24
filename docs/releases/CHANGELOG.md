# Changelog

All notable changes to docxmlater will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [10.1.6] - 2026-02-24

### Added

- **Numbering Restart Helper: `restartNumbering(numId, level?, startValue?)`** - Single-call method to restart list numbering. Creates a new `<w:num>` instance referencing the same abstract numbering with a `<w:lvlOverride>/<w:startOverride>`. Available on both `Document` and `NumberingManager`.

### Statistics

- 143 test suites, 3,084 tests passing
- 103 source files

---

## [10.1.1] - 2026-02-21

### Fixed

- **Table Width Parsing for Auto-Sized Tables**: Tables with `w:tblW w:w="0" w:type="auto"` now correctly parse as `width=0, widthType='auto'` instead of falling through to the constructor default of `9360/dxa`.
- **NaN-Safe Table Property Parsing**: Replaced raw `parseInt` calls with `safeParseInt` for table width and indentation parsing.

### Added

- **Table Indentation Parsing (`w:tblInd`)**: `DocumentParser` now parses `w:tblInd` from main `tblPr` properties per ECMA-376 section 17.4.43.

---

## [10.1.0] - 2026-02-21

### Added

- **Settings API Expansion**: New getter/setter pairs for `hideSpellingErrors`, `hideGrammaticalErrors`, `defaultTabStop`, `updateFields`, `embedTrueTypeFonts`, `saveSubsetFonts`, `doNotTrackMoves`.
- **Hyperlink Attributes**: `setDocLocation()`, `setTgtFrame()`, `setHistory()` with getters and round-trip support.
- **SDT Enhancements**: Placeholder, data binding, showingPlaceholder support; group and inline SDT types.
- **Numbering Enhancements**: Level pStyle association, full level override in NumberingInstance, AbstractNumbering template.
- **Styles Enhancements**: Latent styles support, `setPersonalCompose()` and `setPersonalReply()` on Style.
- **Dirty-tracking for settings**: `_modifiedBooleanSettings` Set for selective merging.

---

## [10.0.4] - 2026-02-21

### Fixed

- **Tracked Deletions Leaving Blank Lines in Simple Markup View**: Documents with pre-existing tracked deletions now preserve paragraph mark revision markers through round-trip processing.

### Added

- **Paragraph Mark Insertion Tracking**: `markParagraphMarkAsInserted()`, `clearParagraphMarkInsertion()`, `isParagraphMarkInserted()`.
- **Paragraph Mark Revision Parsing**: `DocumentParser` extracts `w:del` and `w:ins` from `w:pPr/w:rPr`.
- **Paragraph Mark Revision Acceptance**: `acceptRevisionsInMemory()` and `SelectiveRevisionAcceptor` support paragraph mark markers.

---

## [10.0.3] - 2026-02-21

### Fixed

- **OOXML Compliance**: `tblStyleRowBandSize`/`tblStyleColBandSize` no longer serialize into direct table `w:tblPr`.
- **OOXML Compliance**: Hyperlinks in headers/footers now use part-level `.rels` files.
- **OOXML Compliance**: Property change revisions silently skipped during paragraph content serialization.

---

## [10.0.0] - 2026-02-20

### Added

- **Document Sanitization: `clearDirectSpacingForStyles()`** - Removes direct `w:spacing` overrides from styled paragraphs so style spacing takes effect
- **Image Optimization: `optimizeImages()`** - Lossless PNG re-compression and BMP-to-PNG conversion (zero dependencies)
- **Run Property Change Tracking API** - `getPropertyChangeRevision()`, `setPropertyChangeRevision()`, `clearPropertyChangeRevision()`, `hasPropertyChangeRevision()`
- **Run: `clearMatchingFormatting()`** - Removes formatting matching a style definition for style inheritance
- **Paragraph: `clearSpacing()`** - Clears direct spacing for style inheritance
- **Normal/NormalWeb Style Linking** in `applyStyles()` with preservation flags
- **ImageManager: `updateEntryFilename()`** - Update image filenames after format conversion
- **NumberingManager: `markAbstractNumberingModified()`** - Manual dirty flag for direct level modifications
- **TrackingContext: `createInsertion()`/`createDeletion()`** - Factory methods eliminating circular dependencies

### Changed

- `clearDirectFormattingConflicts()` now clears direct indentation consistently (was conditional)
- `centerLargeImages()` clears direct indentation before centering
- `applyStyles()` caps ListParagraph hanging indent to left indent

### Fixed

- Save state rollback now restores post-processing flags on failure
- Circular dependency in Run tracked changes resolved via TrackingContext factory methods

### Statistics

- 128 test suites, 2,819 tests passing

---

## [9.9.3] - 2026-02-19

### Added

- **Numbering Cleanup Pipeline: Header/Footer/Footnote/Endnote Scanning**
  - `cleanupUnusedNumbering()` now scans headers, footers, footnotes, and endnotes for numId references
  - Previously only scanned body paragraphs and document.xml, which could incorrectly delete numbering definitions used exclusively in headers/footers/notes
  - Raw XML safety net extended to scan all header/footer/footnote/endnote XML files

- **AbstractNumbering: numStyleLink and styleLink Support (ECMA-376 §17.9.21, §17.9.27)**
  - `getNumStyleLink()` / `setNumStyleLink()` and `getStyleLink()` / `setStyleLink()` on AbstractNumbering
  - Parsed from XML and serialized in output

- **Numbering Consolidation: Style Link Fingerprint**
  - Fingerprint now includes `numStyleLink` and `styleLink` to prevent incorrect merging

- **Document Sanitization Pipeline**
  - `flattenFieldCodes()`, `stripOrphanRSIDs()`, `_postProcessDocumentXml()`

### Fixed

- Numbering cleanup false positives for definitions referenced only in headers/footers/notes
- Consolidation style association bug for abstractNums with different style links

### Statistics

- 124 test suites, 2,752 tests passing

---

## [9.2.0] - 2025-01-22

### Added

- **60+ new getter/helper methods** for convenient property access across element classes:
  - **Run**: `getBold()`, `getItalic()`, `getUnderline()`, `getStrike()`, `getFont()`, `getSize()`, `getColor()`, `getHighlight()`, `getSubscript()`, `getSuperscript()`, `isRTL()`, `getSmallCaps()`, `getAllCaps()`
  - **Paragraph**: `getAlignment()`, `getLeftIndent()`, `getRightIndent()`, `getFirstLineIndent()`, `getHangingIndent()`, `getSpaceBefore()`, `getSpaceAfter()`, `getLineSpacing()`, `getKeepNext()`, `getKeepLines()`, `getPageBreakBefore()`, `getOutlineLevel()`, `getTextDirection()`, `getWidowControl()`, `getContextualSpacing()`, `hasNumbering()`, `hasFields()`, `hasBookmarks()`, `hasComments()`, `hasRevisions()`, `isEmpty()`
  - **Table**: `getWidth()`, `getWidthType()`, `getAlignment()`, `getLayout()`, `getIndent()`, `getBorders()`, `getColumnWidths()`, `getCellSpacing()`, `getStyle()`, `hasRows()`, `isFloating()`, `hasStyle()`
  - **TableRow**: `getHeight()`, `getHeightRule()`, `getIsHeader()`, `getCantSplit()`, `getJustification()`, `isHidden()`
  - **TableCell**: `getWidth()`, `getWidthType()`, `getVerticalAlignment()`, `getVerticalMerge()`, `getMargins()`, `getBorders()`, `getShading()`, `getTextDirection()`
  - **Document**: `getParagraphAt()`, `getTableAt()`, `getBodyElementAt()`, `getParagraphIndex()`, `getTableIndex()`, `getNextParagraph()`, `getPreviousParagraph()`
  - **Hyperlink**: `getColor()`, `getUnderline()`, `getBold()`, `getItalic()`, `getFont()`, `getSize()`

### Fixed

- **List bullet/number formatting**: Bullets and numbered list prefixes no longer apply bold formatting by default
  - Previously, list numbering factory methods (createDecimal, createBullet, etc.) set `bold: true`
  - Now defaults to `bold: false` for all list number/bullet symbols
  - Source document bold formatting on bullets is now ignored during parsing

### Changed

- Internal refactoring: Replaced `.getFormatting().property` patterns with new getter methods throughout codebase

## [0.28.0] - 2025-01-20

### Fixed

- **CRITICAL**: Fixed Microsoft Word corruption error for documents with justified text alignment
  - Mapped `alignment: 'justify'` to correct ECMA-376 value `w:val="both"` (was incorrectly using `w:val="justify"`)
  - Fixed in both Style.ts and Paragraph.ts XML generation
  - Resolves "Word found unreadable content" error when opening documents with justified paragraphs

### Added

- **REQUIRED**: Added three mandatory DOCX files per ECMA-376 standard to prevent corruption warnings
  - `word/fontTable.xml` - Font metadata definitions for document fonts
  - `word/settings.xml` - Document settings and Word compatibility configuration
  - `word/theme/theme1.xml` - Office theme with color scheme and font definitions
  - Added automatic registration in [Content_Types].xml and document relationships
  - These files are now included in all documents created with `Document.create()`

### Changed

- Custom paragraph styles now correctly use `customStyle: true` flag to prevent invalid `qFormat` element
- Updated showcase.ts example to demonstrate proper custom style creation
- Improved hanging indent implementation to use `w:hanging` attribute instead of negative `w:firstLine`

### Technical Details

- **Alignment Fix**: Per ECMA-376 Part 1 §17.3.1.13, justified alignment uses enumeration value "both", not "justify"
- **Required Files**: ECMA-376 §11-15 specifies fontTable, settings, and theme as required parts for Office 2007+ compatibility
- **Custom Styles**: Per ECMA-376 §17.7.4.17, `qFormat` element should only appear on built-in Quick Styles, not user-defined custom styles

## [0.5.0] - 2025-01-17

### Security

- **CRITICAL**: Fixed XML injection vulnerability in `Relationship.toXML()` by adding proper XML escaping for all relationship attributes (ID, Type, Target)
- **CRITICAL**: Enhanced path traversal validation to prevent bypass via URL-encoded paths (`%2E`, `%2F`, `%5C`)
- **HIGH**: Added URL protocol validation in Hyperlink class to reject dangerous protocols (javascript:, file:, data:, vbscript:, about:)
- **HIGH**: Added XML special character validation in hyperlink URLs to prevent relationship XML corruption

### Fixed

- **HIGH**: Fixed ECMA-376 property ordering violations in `Paragraph.toXML()` - properties now generated in correct order per §17.3.1.26
- **HIGH**: Fixed ECMA-376 property ordering violations in `Run.toXML()` - properties now generated in correct order per §17.3.2.28
- **MEDIUM**: Fixed circular URL swap bug in `Document.updateHyperlinkUrls()` by implementing two-phase update (collect then apply atomically)
- **MEDIUM**: Added validation in `processHyperlinks()` to throw early error for external hyperlinks without URLs
- **MEDIUM**: Added validation in `Hyperlink.setUrl()` to prevent creating invalid hyperlinks when clearing URL without anchor

### Changed

- Relationship Target URLs are now automatically XML-escaped to prevent injection attacks
- Path traversal validation now checks for URL-encoded bypass attempts
- Hyperlink URL validation now enforces protocol whitelist (HTTP/HTTPS/MAILTO/FTP only)
- `Paragraph.toXML()` now generates properties in ECMA-376 compliant order
- `Run.toXML()` now generates properties in ECMA-376 compliant order
- `Document.updateHyperlinkUrls()` now handles circular URL swaps correctly
- `Hyperlink.setUrl(undefined)` now throws error if result would have no URL and no anchor

### Breaking Changes

- `Hyperlink` constructor and `setUrl()` now validate URLs and throw errors for dangerous protocols or XML special characters
- `normalizePath()` now throws errors for URL-encoded path traversal attempts
- `processHyperlinks()` now throws error early for external hyperlinks without URLs (previously silent failure)
- `Hyperlink.setUrl(undefined)` now throws error if clearing URL would create invalid hyperlink

### Technical Details

All changes maintain backward compatibility for valid use cases. Only invalid/dangerous operations now throw errors:

- **XML Injection Fix**: Uses `XMLBuilder.escapeXmlAttribute()` for proper escaping per ECMA-376 Part 2 §9
- **Path Traversal Fix**: Enhanced `normalizePath()` with URL-encoding detection and better validation
- **URL Validation**: Whitelist approach per ECMA-376 §17.16.22 security guidelines
- **Property Ordering**: Strict compliance with ECMA-376 Part 1 for document integrity with strict parsers
- **Circular Swaps**: Two-phase algorithm prevents URL update order dependencies

## [0.4.0] - 2025-01-16

### Added

- Hyperlink URL update functionality via `Document.updateHyperlinkUrls()`
- `Hyperlink.setUrl()` method with automatic relationship re-registration
- Improved text fallback chain in Hyperlink (text → url → anchor → 'Link')
- Comprehensive hyperlink parsing and validation tests

### Fixed

- Text preservation logic when updating hyperlink URLs
- Relationship re-registration workflow for URL changes

## [0.3.0] - 2025-01-15

### Added

- Security hardening and production improvements
- Enhanced error handling and validation

## [0.2.0] - 2025-01-14

### Added

- Core document functionality
- Paragraph and Run elements
- ZIP handling and XML generation

## [0.1.0] - 2025-01-13

### Added

- Initial release
- Basic DOCX creation and manipulation

[0.5.0]: https://github.com/ItMeDiaTech/docXMLater/compare/v0.4.0...v0.5.0
[0.4.0]: https://github.com/ItMeDiaTech/docXMLater/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/ItMeDiaTech/docXMLater/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/ItMeDiaTech/docXMLater/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/ItMeDiaTech/docXMLater/releases/tag/v0.1.0
