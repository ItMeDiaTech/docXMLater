# docxmlater Playground

A live, runnable sandbox for trying docxmlater in your browser. No install required.

## Open in StackBlitz

[Click here to open the playground](https://stackblitz.com/github/ItMeDiaTech/docXMLater/tree/main/playground)

StackBlitz boots a full Node.js environment in the browser, installs the dependencies, and runs the default example. After it finishes you'll see generated `.docx` files in the Files panel - right-click any file and choose **Download** to open it in Word.

## Examples

| Command                 | What it shows                                                                                            |
| ----------------------- | -------------------------------------------------------------------------------------------------------- |
| `npm start`             | The default example: round-trip a document while recording every edit as a Word revision.                |
| `npm run create`        | Build a polished proposal document from scratch with headings, tables, and styled text.                  |
| `npm run edit`          | Load an existing `.docx`, replace placeholders, highlight text, and save without losing formatting.      |
| `npm run track-changes` | Apply edits to an existing document with full tracked-change attribution. Open in Word to accept/reject. |
| `npm run comments`      | Attach review comments to specific paragraphs, resolve some of them, and verify resolution state.        |

## Running locally instead

If you'd rather try it on your own machine:

```bash
git clone https://github.com/ItMeDiaTech/docXMLater.git
cd docXMLater/playground
npm install
npm start
```

You'll get the same generated `.docx` files in this directory.

## What to try

The fastest way to see what the library does is to edit one of the files in `examples/`. Some ideas:

- In `examples/track-changes.ts`, change the author name and watch how Word attributes the edits.
- In `examples/comments.ts`, swap which comments call `.resolve()` and observe the open/resolved counts.
- In `examples/edit.ts`, run it once to generate `input.docx`, edit it manually in Word, then re-run to see your manual edits preserved alongside the programmatic ones.
- In `examples/create.ts`, add a second table or experiment with shading and borders.

## How it works

Each example imports `docxmlater` from npm, creates or loads a `Document`, performs operations, and writes the result to disk via Node's `fs` module. StackBlitz exposes those files in its Files panel for download.

For full API documentation, see the [main README](../README.md).
