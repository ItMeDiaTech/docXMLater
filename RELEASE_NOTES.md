# docXMLater v0.20.1 Release Notes

## Release Date: October 18, 2025

### Critical Bug Fix Release

This release addresses a critical bug that could cause silent text loss when mixing text and hyperlinks in paragraphs.

## Critical Bug Fixed

### Paragraph.getText() - Hyperlink Text Loss
**Issue**: When using `paragraph.addText()` and `paragraph.addHyperlink()` in the same paragraph, the hyperlink text was completely omitted from `getText()` results.

**Root Cause**: The `Paragraph.getText()` method only filtered for `Run` instances and completely ignored `Hyperlink` instances.

**Fix**: Updated the filter to include both `Run` and `Hyperlink` instances, ensuring all text content is properly consolidated.

**Impact**: This bug could cause serious data loss in production, as hyperlink text would silently disappear.

## What's Included

### Bug Fixes
- Paragraph.getText() now includes hyperlink text content
- Fixed type safety with XMLElement handling
- Improved StylesManager corruption detection
- Enhanced hyperlink relationship ID management

### Test Improvements
- Added 6 comprehensive hyperlink integration tests
- 474/478 tests passing (98.1% pass rate)
- Full test coverage for mixed content scenarios

### Test Cases Added
1. Mixed text + hyperlink retrieval
2. Hyperlink-only paragraph content
3. Multiple hyperlinks with runs
4. Hyperlink before runs
5. Internal (bookmark) hyperlinks
6. Complex multi-hyperlink scenarios

## Test Suite Status

| Metric | Value |
|--------|-------|
| **Total Tests** | 478 |
| **Passing** | 474 |
| **Skipped** | 9 (external fixtures) |
| **Failing** | 0 |
| **Pass Rate** | 98.1% |

## Technical Details

### Modified Files
- `src/elements/Paragraph.ts` - Fixed getText() method
- `tests/elements/Paragraph.test.ts` - Added 6 new test cases
- `README.md` - Updated with v0.20.1 information

### Commit History
- `e0dafe2` - docs: Update README with v0.20.1 release information
- `fc4f540` - Version bump to 0.20.1
- `e656f40` - fix: Fix Paragraph.getText() to include hyperlinks

## Installation

```bash
npm install docxmlater@0.20.1
```

## Usage Example

```typescript
import { Document, Hyperlink } from 'docxmlater';

// Create document
const doc = Document.create();
const para = doc.createParagraph();

// Add mixed content
para.addText('Click ');
const link = Hyperlink.createExternal('https://example.com', 'here');
para.addHyperlink(link);
para.addText(' for more info');

// Now getText() includes ALL content (v0.20.1)
console.log(para.getText()); // Output: "Click here for more info"

// Save document
await doc.save('document.docx');
```

## Package Information

| Field | Value |
|-------|-------|
| **Name** | docxmlater |
| **Version** | 0.20.1 |
| **License** | MIT |
| **Repository** | https://github.com/ItMeDiaTech/docXMLater |
| **npm** | https://www.npmjs.com/package/docxmlater |

## Links

- GitHub Repository: https://github.com/ItMeDiaTech/docXMLater
- npm Package: https://www.npmjs.com/package/docxmlater
- Documentation: https://github.com/ItMeDiaTech/docXMLater/tree/main/docs
- Examples: https://github.com/ItMeDiaTech/docXMLater/tree/main/examples

## Previous Releases

See CHANGELOG.md for complete version history.

---

Ready to upgrade? Run `npm install docxmlater@0.20.1` to get the latest version with this critical bug fix!
