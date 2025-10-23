# Bug Report: Document Corruption Issues

## Date: 2025-10-23
## Reporter: User
## Status: PARTIALLY FIXED

## Issues Identified

### 1. ❌ Global Underline Bug (NOT REPRODUCED)
**Claim:** Everything in document gets underlined
**Finding:** CANNOT REPRODUCE - Framework works correctly
- Test6_BaseFile.docx loads and saves without adding underline
- Test6_Processed_Corrupted.docx was NOT created by current framework
- User may have different version or workflow

### 2. ✅ FIXED: XML Parser Color Bug
**Problem:** Hex color "000000" converted to number 0
**Root Cause:** `XMLParser.parseValue()` in `src/xml/XMLParser.ts`
**Fix Applied:** Lines 757-761 - Added hex color detection

```typescript
// Preserve hex color codes (3 or 6 hex characters)
if (/^[0-9A-Fa-f]{3}$/.test(value) || /^[0-9A-Fa-f]{6}$/.test(value)) {
  return value.toUpperCase(); // Normalize to uppercase per Microsoft convention
}
```

**Impact:** parseToObject() now correctly preserves hex colors

### 3. ⚠️ REMAINING: Style Color Parsing
**Problem:** Heading styles still show wrong colors after fix
**Why:** Styles use `parseRunFormattingFromXml` (old method) not `parseToObject`
**Location:** `DocumentParser.parseStyle()` line 1306

**Next Steps:**
1. Investigate `parseRunFormattingFromXml` color handling
2. OR migrate style parsing to use `parseToObject`
3. Add comprehensive tests for style round-trip

## Missing Framework Features

These elements are stripped during load/save:
- ✗ Bookmarks (`<w:bookmarkStart>`, `<w:bookmarkEnd>`)
- ✗ SDT - Structured Document Tags (`<w:sdt>`)
- ✗ Field codes (`<w:fldChar>`, `<w:instrText>`)
- ✗ Table of Contents (TOC)
- ✗ Table grid changes (`<w:tblGridChange>`)
- ✗ Cell margins (`<w:tcMar>`)
- ✗ Paragraph borders (`<w:pBdr>`)
- ✗ RSID attributes (revision session IDs)

## Test Files

- `Test6_BaseFile.docx` - Original file (styles have correct colors)
- `Test6_Processed_Corrupted.docx` - User's corrupted file (NOT created by current framework)
- `Test6_Resaved.docx` - Framework resave (NO global underline, but style colors wrong)

## Recommendations

1. **Complete style color fix** - Fix old XML parsing path
2. **Add style round-trip tests** - Ensure colors preserved
3. **Document unsupported features** - Warn users about stripped elements
4. **Version tag check** - Verify Test6_Processed_Corrupted creation method

## Code Changes

### Modified Files:
1. `src/xml/XMLParser.ts` - Added hex color preservation (lines 757-761)

### Test Scripts Created:
1. `test-underline-bug.ts` - Cannot reproduce global underline
2. `test-style-color-bug.ts` - Confirms style generation works
3. `test-style-parsing-bug.ts` - Shows parsing still has issues
4. `test-xmlparser-bug.ts` - Confirms parseValue fix works
