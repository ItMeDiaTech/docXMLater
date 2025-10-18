# Implementation Plan - RSID & Polish Fixes

## Source Analysis
- **Source Type**: Best Practice Guidelines (ECMA-376 + Microsoft Recommendations)
- **Core Features**:
  1. Remove invalid all-zero RSID attributes (not needed for programmatic generation)
  2. Normalize color values to uppercase hex
  3. Optimize XML by omitting default property values
- **Dependencies**: None (internal refactoring)
- **Complexity**: LOW (120 LOC total, straightforward refactoring)

## Target Integration
- **Integration Points**:
  - DocumentGenerator: Remove RSID generation
  - Run.ts, Hyperlink.ts: Color normalization
  - Paragraph.ts, Run.ts, TableCell.ts: Strip redundant defaults
- **Affected Files**: 6-8 files
- **Pattern Matching**: Existing defensive coding patterns → lean, optimized patterns

## Implementation Tasks

### Phase 1: Remove RSID Generation (RECOMMENDED)
- [ ] Search for all RSID attribute generation in codebase
- [ ] Remove rsidR, rsidRDefault, rsidP, rsidRPr, rsidDel attributes
- [ ] Document decision in CLAUDE.md
- [ ] Update tests to not expect RSID values
- [ ] Verify Word still opens documents without errors

### Phase 2: Normalize Color Values (POLISH)
- [ ] Add `normalizeColor()` utility function
- [ ] Update all color setters (Run.ts, Hyperlink.ts)
- [ ] Add validation for hex format
- [ ] Normalize to uppercase per Microsoft convention
- [ ] Write tests for color normalization

### Phase 3: Strip Redundant Default Properties (OPTIMIZATION)
- [ ] Create XMLOptimizer helper
- [ ] Update property generation logic to skip defaults
- [ ] Replace explicit false/0 values with element omission
- [ ] Test file size reduction
- [ ] Update tests to expect optimal XML

## Validation Checklist
- [ ] All RSIDs removed (or valid if kept)
- [ ] Colors normalized to uppercase hex
- [ ] No redundant default properties in XML
- [ ] Tests passing (226+)
- [ ] File size reduction verified (target: 20-30%)
- [ ] No regressions in existing functionality
- [ ] Documentation updated
- [ ] Build succeeds with 0 errors

## Success Criteria
1. All RSIDs removed
2. Colors consistently uppercase
3. No default values in XML output
4. File size reduced 20-30%
5. All tests passing
6. No Word compatibility issues
