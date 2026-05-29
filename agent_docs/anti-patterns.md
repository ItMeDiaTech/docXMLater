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

## Setting Color to 'auto' Programmatically

**Wrong:** `run.setColor('auto')` — this throws because `normalizeColor()` validates hex format.
**Why:** The setter runs hex validation. The `auto` value is only preserved during round-trip (parser bypasses the setter for `auto`).
**Fix:** For programmatic use, use `'000000'` for black. For round-trip, `auto` is handled automatically.

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

## addHyperlink Return Type

**Wrong:** Chaining `para.addHyperlink('https://example.com').addText('more text')` and expecting paragraph methods.
**Why:** `addHyperlink(url: string)` returns the `Hyperlink` object (for configuring the link). Only `addHyperlink(hyperlink: Hyperlink)` returns `this` (the paragraph). The overloaded signatures have different return types.
**Fix:** Use two statements: `const link = para.addHyperlink('url'); link.setText('text');` — or pass a pre-built Hyperlink: `para.addHyperlink(new Hyperlink({ url, text }))` which returns `this` for chaining.

## Direct Property Mutation

**Wrong:** Directly assigning to `doc.bodyElements`, `paragraph.content`, or `paragraph.formatting` without understanding the implications.
**Why:** Some properties are mutable for internal use but modifying them directly bypasses validation, dirty flag updates, and event tracking. The object model may become inconsistent.
**Fix:** Use API methods (`addParagraph()`, `removeParagraph()`, `addText()`, `setStyle()`) instead of direct property assignment. Direct mutation is only safe when explicitly documented.

## Concurrent Save Calls

**Wrong:** Calling `save()` or `toBuffer()` concurrently on the same Document instance.

```typescript
// This races:
await Promise.all([doc.save(path1), doc.save(path2)]);
```

**Why:** `prepareSave()` modifies internal managers (StylesManager, NumberingManager) without locking. Concurrent calls see partially-updated state, potentially corrupting output.
**Fix:** Serialize save operations — wait for one save to complete before starting another. Each Document instance should be used by one async caller at a time.

## Table Cell Content Assumption

**Wrong:** Calling `cell.getParagraphs()[0].addText('text')` on a newly created table.
**Why:** Newly created table cells may have empty paragraph arrays. The `[0]` index returns `undefined`.
**Fix:** Use `cell.addParagraph(new Paragraph().addText('text'))` to populate cells explicitly.

## Relationship ID Ordering

**Wrong:** Assuming relationship IDs (`rId1`, `rId2`, etc.) are assigned in a specific order or are stable across save/load cycles.
**Why:** RelationshipManager assigns sequential IDs, but the order depends on when relationships are registered. Loading a document may assign different IDs than the original.
**Fix:** Never hardcode relationship IDs. Use the API to look up relationships by type or target.

## CT_OnOff Truthiness Check

**Wrong:** Using `if (formatting.bold)` to decide whether to emit a boolean XML element.
**Why:** This drops explicit `false` values. Per ECMA-376, `<w:b w:val="0"/>` means "explicitly not bold" and overrides style inheritance. Omitting the element means "inherit from style" — different semantics.
**Fix:** Use `if (formatting.bold !== undefined)` and emit `w:val="0"` for false. This applies to all CT_OnOff properties: bold, italic, strike, caps, keepNext, keepLines, pageBreakBefore, contextualSpacing, etc.

## Removing Last Table Row

**Wrong:** `table.removeRow(0)` on a single-row table.
**Why:** Per ECMA-376 §17.4.38, `w:tbl` must contain at least one `w:tr`. An empty table produces invalid OOXML that Word rejects.
**Fix:** `removeRow()` returns `false` when attempting to remove the last row (unless change tracking is enabled, where the row is marked as deleted instead of physically removed).

## insertRow Column Count with Merged Cells

**Wrong:** Assuming `getColumnCount()` returns the table grid width.
**Why:** `getColumnCount()` returns the max number of `<w:tc>` elements, not accounting for `gridSpan`. A row with 2 cells where one spans 3 columns has 4 grid columns but `getColumnCount()` returns 2.
**Fix:** Use `getTotalGridSpan()` (or the explicit `tableGrid`) when creating rows to match the actual grid width. `insertRow()`/`insertRows()` handle this correctly.

## Change Tracking Property Omissions

**Wrong:** Serializing fewer elements in `*Change` previous properties than in the main property block.
**Why:** When Word shows "Original" markup, it uses the previous properties from the change tracking element. Missing properties cause incorrect display of the document's pre-change state.
**Fix:** Every property serialized in the main block (tblPr, tcPr, trPr, pPr, sectPr) must also be serialized in the corresponding change tracking block (tblPrChange, tcPrChange, trPrChange, pPrChange, sectPrChange).
