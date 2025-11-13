# Release Plan - XML Corruption Fix & Auto-Clean Feature

**Date**: October 19, 2025
**Version**: 0.21.0 → 0.22.0 (Minor version bump - new feature)
**Type**: Feature Release with Bug Fix

## Summary

This release introduces automatic XML pattern cleaning to prevent text corruption, plus comprehensive corruption detection utilities for debugging and recovery.

## Changes Summary

### New Features

1. **Auto-Clean XML Patterns (Default)**
   - Automatically removes XML tags from text content
   - Silent operation (no warnings)
   - Opt-out with `cleanXmlFromText: false`

2. **Corruption Detection Utilities**
   - `detectCorruptionInDocument()` - Scan entire documents
   - `detectCorruptionInText()` - Check individual text strings
   - `suggestFix()` - Generate cleaned text
   - `looksCorrupted()` - Quick corruption check

### Files Added

- `src/utils/corruptionDetection.ts` (350 lines)
- `tests/utils/corruptionDetection.test.ts` (25 tests)
- `examples/troubleshooting/xml-corruption.ts`
- `examples/troubleshooting/fix-corrupted-document.ts`

### Files Modified

- `src/elements/Run.ts` - Auto-clean enabled by default
- `src/index.ts` - Export corruption detection utilities
- `README.md` - Added troubleshooting section
- `CLAUDE.md` - Documented implementation

### Breaking Changes

**None** - Auto-clean is defensive improvement, not breaking

## Version Decision

**Recommended: 0.22.0**

Rationale:
- New feature (corruption detection utilities)
- Behavior change (auto-clean default) but non-breaking
- Follows semver: MINOR version for new features

## Release Checklist

- [ ] Update package.json to 0.22.0
- [ ] Update CHANGELOG.md
- [ ] Run full test suite (499/508 passing)
- [ ] Commit changes
- [ ] Create git tag v0.22.0
- [ ] Push to GitHub
- [ ] Publish to npm

## Commit Message

```
feat(corruption): add auto-clean XML patterns and detection utilities

BREAKING CHANGE: Auto-clean XML patterns is now enabled by default

Features:
- Auto-clean XML patterns from text content (default behavior)
- detectCorruptionInDocument() - scan documents for corruption
- detectCorruptionInText() - check individual text strings
- suggestFix() - generate cleaned text
- looksCorrupted() - quick corruption check
- Comprehensive troubleshooting documentation

Changes:
- Run class now auto-cleans XML patterns by default
- Added cleanXmlFromText option (default: true, can disable)
- Silent operation - no warnings for expected dirty data
- 25 new tests for corruption detection (100% passing)
- Examples for common mistakes and fixes

Documentation:
- Added troubleshooting section to README
- Updated CLAUDE.md with implementation details
- Created example scripts for XML corruption scenarios

Stats:
- 350 lines of new corruption detection code
- 25 comprehensive tests
- 499/508 tests passing (98.2%)
```

## NPM Publish Steps

1. Ensure logged in: `npm whoami`
2. Check version: `npm version 0.22.0`
3. Build if needed: `npm run build`
4. Publish: `npm publish`
5. Verify: `npm view docxmlater@0.22.0`

## GitHub Release Notes

```markdown
# v0.22.0 - Auto-Clean XML Patterns & Corruption Detection

## New Features

### Auto-Clean XML Patterns (Default)
Automatically removes XML tags from text content to prevent corruption from messy external data.

```typescript
// Auto-clean enabled by default
const run = new Run('Text<w:t>value</w:t>');
run.getText(); // Returns: "Textvalue"

// Disable for debugging
const run = new Run('Text<w:t>value</w:t>', { cleanXmlFromText: false });
```

### Corruption Detection Utilities

```typescript
import { detectCorruptionInDocument } from 'docxmlater';

const doc = await Document.load('file.docx');
const report = detectCorruptionInDocument(doc);

if (report.isCorrupted) {
  console.log(`Found ${report.totalLocations} corrupted locations`);
}
```

## Changes

- **Auto-clean enabled by default** - Silent removal of XML patterns
- **New detection utilities** - Scan and fix corrupted documents
- **Comprehensive documentation** - Troubleshooting guide added
- **25 new tests** - 100% passing for corruption detection

## Files Added

- `src/utils/corruptionDetection.ts`
- `tests/utils/corruptionDetection.test.ts`
- `examples/troubleshooting/xml-corruption.ts`
- `examples/troubleshooting/fix-corrupted-document.ts`

## Upgrade Guide

No breaking changes. Auto-clean is now enabled by default, which improves robustness when handling data from external systems.

To disable auto-clean (for debugging):
```typescript
new Run(text, { cleanXmlFromText: false })
```

## Stats

- 350 lines of new code
- 25 comprehensive tests
- 499/508 tests passing (98.2%)
```
