# Data Loss Prevention - Implementation Plan

**Created:** October 24, 2025
**Priority:** CRITICAL - Fixing data loss issues before v1.0.0

## Context

During thorough testing, multiple data loss patterns were discovered:
- SDT parsing was completely broken (FIXED)
- Run special elements were lost (FIXED)
- Comments, bookmarks, and complex fields still being lost

## Completed in Current Session

### ✅ Phase 1: SDT Parsing (CRITICAL - COMPLETE)
- **Issue:** parseSDTFromObject() returned null, deleting all SDT content
- **Impact:** Lost tables, TOCs, content controls
- **Solution:** Implemented full 165-line parser
- **Result:** 100% SDT preservation verified

### ✅ Phase 2: Run Special Elements (HIGH - COMPLETE)
- **Issue:** w:br, w:tab, w:sym ignored in runs
- **Impact:** Lost formatting and layout
- **Solution:** Enhanced parseRunFromObject() and Run.toXML()
- **Result:** Special characters preserved in round-trip

## Remaining Implementation Tasks

### Phase 3: Comment/Annotation Parsing (HIGH PRIORITY)
**Risk:** Loss of review feedback and collaboration notes

**Implementation:**
1. [ ] Create parseCommentFromObject() method in DocumentParser
2. [ ] Handle w:commentRangeStart/End in paragraph parsing
3. [ ] Parse w:commentReference in runs
4. [ ] Integrate with existing CommentManager
5. [ ] Test with documents containing comments

**Files to modify:**
- src/core/DocumentParser.ts (add parsing methods)
- src/core/Document.ts (integrate comment loading)

### Phase 4: Bookmark Parsing (MEDIUM PRIORITY)
**Risk:** Bookmarks lost on document round-trip

**Implementation:**
1. [ ] Create parseBookmarkFromObject() method
2. [ ] Handle w:bookmarkStart/End pairs
3. [ ] Track bookmark IDs and names
4. [ ] Add to paragraph child order reconstruction

**Files to modify:**
- src/core/DocumentParser.ts (add parsing methods)
- src/elements/Bookmark.ts (ensure compatibility)

### Phase 5: Complex Field Parsing (MEDIUM PRIORITY)
**Risk:** Mail merge fields and conditional content lost

**Implementation:**
1. [ ] Create parseComplexFieldFromObject() method
2. [ ] Handle w:fldChar (begin/separate/end)
3. [ ] Parse w:instrText for field codes
4. [ ] Support w:fldData for field values

**Files to modify:**
- src/core/DocumentParser.ts (add parsing methods)
- src/elements/Field.ts (extend for complex fields)

### Phase 6: Paragraph Child Order (MEDIUM PRIORITY)
**Risk:** Incorrect ordering of mixed content

**Implementation:**
1. [ ] Extend extractParagraphChildOrder regex
2. [ ] Include: w:bookmarkStart/End, w:commentRangeStart/End
3. [ ] Maintain proper sequencing of all inline elements

**Files to modify:**
- src/core/DocumentParser.ts (line 2234, extend regex)

### Phase 7: Additional Body Elements (LOW PRIORITY)
**Risk:** Rare content types ignored

**Implementation:**
1. [ ] Add w:altChunk parsing (alternate content)
2. [ ] Add w:customXml parsing (custom markup)
3. [ ] Handle track changes elements if needed

**Files to modify:**
- src/core/DocumentParser.ts (parseBodyElements method)

## Testing Strategy

### Test Documents Needed
1. **comments-test.docx** - Document with multiple comments and replies
2. **bookmarks-test.docx** - Document with bookmarks and cross-references
3. **fields-test.docx** - Document with complex fields (mail merge, IF)
4. **mixed-content.docx** - Document with complex inline element ordering

### Validation Process
1. Load test document
2. Save to new file
3. Compare XML structure
4. Verify no content loss
5. Check element ordering

## Success Criteria

- [ ] Zero content loss on round-trip
- [ ] All tests passing (current + new)
- [ ] Comments preserved with author/date
- [ ] Bookmarks maintain IDs and names
- [ ] Complex fields retain instructions
- [ ] Element ordering preserved
- [ ] No regression in existing features

## Time Estimate

- Phase 3 (Comments): 2 hours
- Phase 4 (Bookmarks): 1 hour
- Phase 5 (Complex Fields): 2 hours
- Phase 6 (Child Order): 30 minutes
- Phase 7 (Body Elements): 1 hour
- Testing & Validation: 1.5 hours

**Total: ~8 hours**

## Next Action

Start with Phase 3 (Comment Parsing) as it represents the highest risk of losing important user content.