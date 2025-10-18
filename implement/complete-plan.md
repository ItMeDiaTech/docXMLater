# Implementation Plan: getRawXml & Unimplemented Features

**Session Date**: October 18, 2025
**Task**: Implement getRawXml method and all other "not yet implemented" markers
**Status**: Active Implementation

## Executive Summary

OOXML validation is failing due to missing `getRawXml` method. This plan consolidates:
1. **getRawXml implementation** (critical for OOXML validation)
2. **Table parsing** (DocumentParser - not yet implemented)
3. **NumberingInstance level overrides** (formatting - placeholders only)

## Issues Identified

### Issue 1: getRawXml Not Implemented (CRITICAL)
**Error**:
```
[warn] getRawXml not yet implemented in DocXMLater for part: word/document.xml
[warn] getRawXml not yet implemented in DocXMLater for part: word/_rels/document.xml.rels
```

**Location**: Should be method on Document or Part classes
**Impact**: OOXML validation fails, cannot validate document structure
**Priority**: CRITICAL

### Issue 2: Table Parsing Not Yet Implemented
**Location**: `src/core/DocumentParser.ts` - line with comment
**Content**: "Check for tables (not yet implemented)"
**Impact**: Tables in loaded documents not properly parsed
**Priority**: HIGH

### Issue 3: NumberingInstance Level Overrides Placeholder
**Location**: `src/formatting/NumberingInstance.ts` (2 instances)
**Content**: "Placeholder for level overrides (not yet implemented)"
**Impact**: Complex numbering with overrides not fully supported
**Priority**: MEDIUM

## Phase 1: Implement getRawXml Method (CRITICAL)

### Analysis
`getRawXml` is needed for OOXML validation - it returns the raw XML content of a document part as a string without parsing.

### Implementation Location
Add to `src/core/Document.ts` as async method

### Methods to Add
1. `getRawXml(partName: string): Promise<string | null>` - Get raw XML for any part
2. Optional: `getAllRawXml(): Promise<Map<string, string>>` - Get all XML parts

### Implementation Details
- Get part from zipHandler
- If content is string, return as-is
- If content is Buffer, decode as UTF-8
- Return null if part not found

## Phase 2: Implement Table Parsing (HIGH PRIORITY)

### Current State
DocumentParser exists but table parsing is not implemented (has TODO comment)

### Location
`src/core/DocumentParser.ts` - parseDocument() or related method

### Implementation
Parse `<w:tbl>` elements from word/document.xml

### Components
- Parse table rows (w:tr)
- Parse table cells (w:tc)
- Handle cell content
- Support merged cells
- Support nested tables

## Phase 3: Implement NumberingInstance Level Overrides (MEDIUM PRIORITY)

### Current State
`src/formatting/NumberingInstance.ts` has placeholder comments for level overrides

### Implementation
- Add levelOverrides Map to store overrides
- Implement addLevelOverride method
- Implement getLevelOverride method
- Generate proper XML with overrides

## Implementation Checklist

### Phase 1: getRawXml (CRITICAL)
- [ ] Add getRawXml method to Document.ts
- [ ] Add optional getAllRawXml method
- [ ] Add JSDoc documentation
- [ ] Test with various part types
- [ ] Verify OOXML validation passes
- [ ] Run full test suite

### Phase 2: Table Parsing (HIGH)
- [ ] Implement parseTable in DocumentParser
- [ ] Add table row parsing logic
- [ ] Add table cell parsing logic
- [ ] Handle merged cells
- [ ] Handle nested tables
- [ ] Write 5+ table parsing tests
- [ ] Verify round-trip preservation

### Phase 3: NumberingInstance Overrides (MEDIUM)
- [ ] Add levelOverrides Map to NumberingInstance
- [ ] Implement addLevelOverride method
- [ ] Implement getLevelOverride method
- [ ] Generate proper XML with overrides
- [ ] Write 3+ override tests
- [ ] Update CLAUDE.md

## Affected Files

### Modified Files
1. `src/core/Document.ts` - Add getRawXml, getAllRawXml methods
2. `src/core/DocumentParser.ts` - Implement table parsing
3. `src/formatting/NumberingInstance.ts` - Implement level overrides

### Test Files to Add/Update
1. `tests/core/Document.test.ts` - Add getRawXml tests
2. `tests/core/DocumentParser.test.ts` - Add table parsing tests
3. `tests/formatting/Numbering.test.ts` - Add level override tests

## Success Criteria

✅ getRawXml method works for all part types
✅ OOXML validation no longer shows "getRawXml not yet implemented"
✅ Tables parsed correctly from loaded documents
✅ Table round-trip preservation (load → modify → save)
✅ NumberingInstance level overrides functional
✅ All 12+ new tests passing
✅ No breaking changes
✅ Full backward compatibility

## Implementation Order

1. **getRawXml** (30 min) - Critical, fixes validation
2. **Table Parsing** (40 min) - High priority, completes parsing
3. **NumberingInstance Overrides** (20 min) - Lower priority, enhances formatting
4. **Testing** (30 min) - Comprehensive validation
5. **Documentation** (15 min) - Update CLAUDE.md

**Total Estimated Time**: ~135 minutes

## Version Impact

- **Current**: 0.9.0 (Font Embedding Planning)
- **Next**: 0.10.0 (Implementation Complete)
- **Breaking Changes**: None
- **New Public APIs**:
  - `Document.getRawXml(partName): Promise<string | null>`
  - `Document.getAllRawXml(): Promise<Map<string, string>>`
- **Deprecations**: None

## Risk Mitigation

**Risk**: getRawXml performance with large files
**Mitigation**: Cache results, document performance implications

**Risk**: Table parsing edge cases
**Mitigation**: Start with simple tables, add complexity incrementally

**Risk**: Breaking existing numbering behavior
**Mitigation**: Thorough testing before release, semantic versioning
