# Changelog

All notable changes to docxmlater will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
- Added comprehensive formatting module documentation: `src/formatting/CLAUDE.md` (52 lines)

### Tests
- All 1180 tests passing (53 test suites)
- New test coverage for tracked changes functionality
- New test coverage for paragraph numbering indentation
- Test output files cleaned up and removed from git tracking

### Documentation
- Added comprehensive list indentation analysis document
- Updated formatting module CLAUDE.md with detailed specifications
- Added advanced tracked changes example with real-world scenarios

---

## [1.17.0] - 2025-11-13

Internal release with infrastructure improvements.

---

## [1.16.0] - 2025-11-13

### Documentation
- **Comprehensive Documentation Update**: Added complete documentation suite
  - New README.md with full feature matrix, API overview, and code examples
  - Updated CLAUDE.md to reflect all 5 phases complete (2073+ tests, 65 source files)
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
  resetFormatting?: boolean,      // Reset to standard style
  cleanupRelationships?: boolean  // Clean orphaned relationships
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
const level = NumberingLevel.createBulletLevel(0, '•');
level.setFont('Symbol');
```

**Numbered List Formats:**
Level 3 now shows uppercase letters (A., B., C.) instead of numbers (1., 2., 3.). To maintain old behavior:
```typescript
const formats = ['decimal', 'lowerLetter', 'lowerRoman']; // 3-level cycle
const abstractNum = AbstractNumbering.createNumberedList(1, 9, formats);
```

**Special Characters:**
Text containing tabs, newlines, etc. now automatically converts to proper XML elements. This is generally what you want, but if you need literal characters:
```typescript
// Tabs and newlines now auto-convert to XML elements
const run = new Run('Text\tWith\nSpecial');
// Generates: <w:t>Text</w:t><w:tab/><w:t>With</w:t><w:br/><w:t>Special</w:t>

// To preserve as literal text (not recommended):
const run = new Run('Text\\tWith\\nSpecial');
```
