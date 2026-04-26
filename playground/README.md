# docxmlater Playground

A live, runnable sandbox for trying docxmlater in your browser. No install required.

## Open in StackBlitz

[Click here to open the playground](https://stackblitz.com/github/ItMeDiaTech/docXMLater/tree/main/playground)

StackBlitz boots a full Node.js environment in the browser, installs the dependencies, and runs the default example. After it finishes you'll see generated `.docx` files in the Files panel - right-click any file and choose **Download** to open it in Word.

## Default example

```bash
npm start
```

Round-trips an existing document while recording every edit as a Word revision. This is docxmlater's signature feature: editing existing files without corrupting them, with full tracked-changes attribution.

## Topical examples

Each example is small and focused. Run any of them with `npm run <name>`.

| #   | Script                   | What it shows                                                                             |
| --- | ------------------------ | ----------------------------------------------------------------------------------------- |
| 01  | `01-basic`               | Smallest possible program: create a document and save it.                                 |
| 02  | `02-text`                | Character formatting: bold, italic, color, highlight, sub/superscript, strike, caps.      |
| 03  | `03-lists`               | Bulleted, numbered, and multi-level nested lists with mixed numbering formats.            |
| 04  | `04-styles`              | Apply built-in styles (Title, Heading1-9, Subtitle, Normal) and define a custom style.    |
| 05  | `05-images`              | Embed a PNG image with sizing and alt text.                                               |
| 06  | `06-headers-footers`     | Document title in the header, "Page X of Y" page numbering in the footer.                 |
| 07  | `07-hyperlinks`          | External web links, email links, custom-styled links, and inline links inside paragraphs. |
| 08  | `08-table-of-contents`   | Insert a TOC field that Word populates from the document's headings.                      |
| 09  | `09-bookmarks`           | Mark a section, then jump to it from a hyperlink elsewhere in the document.               |
| 10  | `10-track-changes`       | Edit an existing document with every change recorded as a Word revision.                  |
| 11  | `11-comments`            | Attach review comments to paragraphs, with both open and resolved state.                  |
| 12  | `12-tables`              | A styled table with header row shading, cell borders, and content per cell.               |
| 13  | `13-logging`             | Configure docxmlater's internal log level via `DOCXMLATER_LOG_LEVEL`.                     |
| 14  | `14-fonts`               | Apply different font families and sizes per run; set a document-wide default.             |
| 15  | `15-footnotes-endnotes`  | Register footnotes and endnotes; inspect the managers that own them.                      |
| 16  | `16-content-controls`    | Build a Word form with plain text, date picker, dropdown, and checkbox SDTs.              |
| 17  | `17-complex-fields`      | Insert dynamic fields: author, title, date, page number, total pages, filename.           |
| 18  | `18-math-equations`      | OMML round-trip workflow notes (math is preserved verbatim from Word-authored documents). |
| 19  | `19-document-protection` | Lock a document into tracked-changes-only editing mode.                                   |
| 20  | `20-compatibility-mode`  | Detect a document's Word compatibility mode and the upgrade-to-modern API.                |

You can also run **all** examples in one shot:

```bash
npm run all
```

## Running locally

```bash
git clone https://github.com/ItMeDiaTech/docXMLater.git
cd docXMLater/playground
npm install
npm start
```

Generated `.docx` files appear in the playground directory after each run.

## Modifying the examples

Edit any file in `examples/` and re-run that script. Some ideas:

- In `examples/10-track-changes.ts`, change the author name and watch how Word attributes the edits.
- In `examples/11-comments.ts`, swap which comments call `.resolve()` and observe the open/resolved counts.
- In `examples/16-content-controls.ts`, add a new dropdown option or a fifth field type.
- In `examples/12-tables.ts`, change the shading color or add a totals row.

## How it works

Each example imports `docxmlater` from npm, creates or loads a `Document`, performs operations, and writes the result to disk via Node's `fs` module. StackBlitz exposes those files in its Files panel for download.

For full API documentation, see the [main README](../README.md).
