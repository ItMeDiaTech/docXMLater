# docXMLater v10.1.6 Release Notes

## Release Date: February 24, 2026

### Feature Release

This release adds a numbering restart helper and continues the ECMA-376 comprehensive gap analysis work with OOXML compliance fixes, expanded APIs, and extensive test coverage.

## Key Features

### Numbering Restart Helper

- `restartNumbering(numId, level?, startValue?)` on both `Document` and `NumberingManager`
- Creates a new numbering instance referencing the same abstract numbering with a level override
- Replaces the previous multi-step manual process

```typescript
const listId = doc.createNumberedList();
doc.createParagraph('Item 1').setNumbering(listId, 0);
doc.createParagraph('Item 2').setNumbering(listId, 0);

// Restart numbering from 1
const restartId = doc.restartNumbering(listId);
doc.createParagraph('New item 1').setNumbering(restartId, 0);
```

### Settings API Expansion (10.1.0)

- New getter/setter pairs for `hideSpellingErrors`, `hideGrammaticalErrors`, `defaultTabStop`, `updateFields`, `embedTrueTypeFonts`, `saveSubsetFonts`, `doNotTrackMoves`
- Dirty-tracking for selective merging with original settings XML

### SDT and Numbering Enhancements (10.1.0)

- Structured Document Tag: placeholder, data binding, showingPlaceholder support
- Numbering: level pStyle association, full level overrides, AbstractNumbering template
- Styles: latent styles support, `setPersonalCompose()`, `setPersonalReply()`

### Hyperlink Attributes (10.1.0)

- `setDocLocation()`, `setTgtFrame()`, `setHistory()` with getters and round-trip support

## Bug Fixes

### Table Width Parsing (10.1.1)

- Auto-sized tables (`w:tblW w:w="0" w:type="auto"`) now parse correctly
- NaN-safe table property parsing prevents propagation from malformed input
- Table indentation (`w:tblInd`) parsing from main `tblPr` properties

### Paragraph Mark Revision Tracking (10.0.4)

- Tracked deletions no longer leave blank lines in Simple Markup View
- Full paragraph mark insertion/deletion tracking API

### OOXML Compliance (10.0.3)

- `tblStyleRowBandSize`/`tblStyleColBandSize` correctly limited to style definitions
- Header/footer hyperlinks use part-scoped relationship files
- Property change revisions properly skipped in paragraph content serialization

## Test Suite Status

| Metric           | Value |
| ---------------- | ----- |
| **Test Suites**  | 143   |
| **Total Tests**  | 3,084 |
| **Passing**      | 100%  |
| **Source Files** | 103   |

## Installation

```bash
npm install docxmlater@10.1.6
```

## Package Information

| Field          | Value                                     |
| -------------- | ----------------------------------------- |
| **Name**       | docxmlater                                |
| **Version**    | 10.1.6                                    |
| **License**    | MIT                                       |
| **Repository** | https://github.com/ItMeDiaTech/docXMLater |
| **npm**        | https://www.npmjs.com/package/docxmlater  |

## Links

- GitHub Repository: https://github.com/ItMeDiaTech/docXMLater
- npm Package: https://www.npmjs.com/package/docxmlater
- Documentation: https://github.com/ItMeDiaTech/docXMLater/tree/main/docs

## Previous Releases

See CHANGELOG.md for complete version history.
