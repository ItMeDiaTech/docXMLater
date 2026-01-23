# Changelog

All notable changes to DocXML will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
