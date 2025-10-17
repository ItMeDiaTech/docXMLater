# Changelog

All notable changes to DocXML will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
