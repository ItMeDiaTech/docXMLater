# docXMLater v10.2.9 Release Notes

## Release Date: March 4, 2026

### Bug Fix Release

This release fixes incorrect text ordering when accepting tracked changes inside ComplexField result sections, removes dead code from the parser, and strengthens test assertions.

## Bug Fixes

### ComplexField Revision Acceptance Text Ordering

- `acceptAllRevisions()` now produces correct interleaved text when accepting revisions inside ComplexField HYPERLINK results
- Previously, insertion text was appended to the end of the result string instead of being merged in document order
- Example: interleaved content like "Co" + ins("mmercial...") + "verage..." now correctly produces "Commercial PA Appeals - Coverage Determination Denial Reasons (CMS-PRD1-086897)" instead of "Coverage Determination...(CMS-PRD1-086897)mmercial PA Appeals - Co"
- Uses `getAcceptedResultText()` which processes `resultContent` XML elements in document order, including plain runs and insertions while skipping deletions

## Code Quality

### Dead Code Removal in DocumentParser

- Removed ~30 lines of unreachable code in `tryPromoteRevisionFieldCode()` after the early return guard for content revision types (insert, delete, moveFrom, moveTo)
- Property change revision types never contain field character sequences, so the downstream field detection and promotion logic was never executed

## Test Suite Status

| Metric           | Value |
| ---------------- | ----- |
| **Test Suites**  | 146   |
| **Total Tests**  | 3,120 |
| **Passing**      | 100%  |
| **Source Files** | 120   |

## Installation

```bash
npm install docxmlater@10.2.9
```

## Package Information

| Field          | Value                                     |
| -------------- | ----------------------------------------- |
| **Name**       | docxmlater                                |
| **Version**    | 10.2.9                                    |
| **License**    | MIT                                       |
| **Repository** | https://github.com/ItMeDiaTech/docXMLater |
| **npm**        | https://www.npmjs.com/package/docxmlater  |

## Links

- GitHub Repository: https://github.com/ItMeDiaTech/docXMLater
- npm Package: https://www.npmjs.com/package/docxmlater
- Documentation: https://github.com/ItMeDiaTech/docXMLater/tree/main/docs

## Previous Releases

See CHANGELOG.md for complete version history.
