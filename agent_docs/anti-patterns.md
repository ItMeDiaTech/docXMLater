# Anti-Patterns

Things that look correct but break documents or tests.

## XML in Text Content
**Wrong:** `paragraph.addText('text<w:t>value</w:t>')`
**Why:** The framework correctly escapes `<` to `&lt;`. Passing XML as text is user error.
**Fix:** Use the proper API methods instead of embedding raw XML in text.

## Missing Content Types
**Wrong:** Adding a new part to the ZIP without updating `[Content_Types].xml`.
**Why:** Word will report the document as corrupt and refuse to open it.
**Fix:** Always register the content type in Generator when adding new part types.

## cnfStyle Preservation
**Wrong:** Removing paragraphs from table cells without checking cnfStyle.
**Why:** `cnfStyle` bitmasks control table style conditionals (firstRow, lastRow, etc.). Removing a paragraph with cnfStyle causes Word to apply unexpected default shading.
**Fix:** Use `isParaBlank()` which preserves paragraphs with cnfStyle.

## RSIDs in Generated Content
**Wrong:** Generating rsid attributes for programmatically created content.
**Why:** RSIDs are OPTIONAL per ECMA-376 and Word regenerates them on first edit. Adding them increases file size and creates non-deterministic output that breaks golden tests.
**Fix:** Omit rsid attributes entirely for new content.

## Setting Color to 'auto'
**Wrong:** `run.setColor('auto')` — this fails.
**Why:** The setter only accepts hex color strings.
**Fix:** Use `'000000'` for black.

## Dropping Run Properties
**Wrong:** Calling `setFont()`, `setSize()`, or `setBold()` and assuming other properties persist.
**Why:** Some setters can drop existing run properties like color or underline.
**Fix:** Always re-set color/underline after modifying formatting on Hyperlink-styled runs.

## Cleanup Deleting User Content
**Wrong:** `CleanupHelper.run({ cleanupNumbering: true })` after creating numbering definitions.
**Why:** The cleanup method removes numbering definitions it considers orphaned, including ones just created.
**Fix:** Use `cleanupNumbering: false` when running cleanup after programmatic numbering creation.

## Treating PreservedElement as Hyperlink
**Wrong:** Using `isHyperlink()` type guard on content from real documents.
**Why:** `para.getContent()` returns `PreservedElement` for hyperlinks in loaded documents, not `Hyperlink` instances.
**Fix:** Check for PreservedElement when processing loaded document content.

## Modifying After Raw XML Accept
**Wrong:** Modifying content after calling `acceptAllRevisionsRawXml()`.
**Why:** This legacy method sets `skipDocumentXmlRegeneration = true`, so subsequent modifications are NOT saved.
**Fix:** Use `acceptAllRevisions()` (in-memory) instead, which fully supports subsequent modifications.
