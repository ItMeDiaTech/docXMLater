# TOC Pre-Population: Final Implementation Plan

## Executive Summary

**Goal**: Generate Table of Contents with actual heading entries visible when document is first opened in Word, while maintaining field structure for "Update Field" functionality.

**Good News**: The complex logic already exists in [`Document.replaceTableOfContents()`](src/core/Document.ts:5108-5184). We just need to integrate it into the save flow.

## Architecture Documents Created

1. [`TOC_FIELD_ARCHITECTURE.md`](TOC_FIELD_ARCHITECTURE.md) - Overall architecture and 17-hour timeline
2. [`TOC_IMPLEMENTATION_GUIDE.md`](TOC_IMPLEMENTATION_GUIDE.md) - Field structure deep dive with diagrams
3. [`TOC_FIELD_IMPLEMENTATION_SUMMARY.md`](TOC_FIELD_IMPLEMENTATION_SUMMARY.md) - Current status summary
4. [`TOC_PREPOPULATION_PLAN.md`](TOC_PREPOPULATION_PLAN.md) - Detailed pre-population design
5. [`TOC_PREPOPULATION_ACTION_PLAN.md`](TOC_PREPOPULATION_ACTION_PLAN.md) - Step-by-step action plan

## Implementation Summary

### Changes Required

| File                                           | Change                                    | Impact                |
| ---------------------------------------------- | ----------------------------------------- | --------------------- |
| [`src/core/Document.ts`](src/core/Document.ts) | Add `autoPopulateTOCs` flag               | Low - additive        |
| [`src/core/Document.ts`](src/core/Document.ts) | Add `setAutoPopulateTOCs()` method        | Low - additive        |
| [`src/core/Document.ts`](src/core/Document.ts) | Modify `save()` to populate TOCs          | Medium - modification |
| [`src/core/Document.ts`](src/core/Document.ts) | Modify `toBuffer()` to populate TOCs      | Medium - modification |
| [`src/core/Document.ts`](src/core/Document.ts) | Extract `populateAllTOCsInXML()`          | Low - refactor        |
| [`src/core/Document.ts`](src/core/Document.ts) | Add `createPrePopulatedTableOfContents()` | Low - additive        |

### Code Reuse

**Excellent news**: ~95% of the code already exists!

Reusable methods:

- ✅ [`findHeadingsForTOCFromXML()`](src/core/Document.ts:4734-4878) - Finds headings in document
- ✅ [`generateTOCXML()`](src/core/Document.ts:4940-5035) - Generates populated TOC XML
- ✅ [`buildTOCEntryXML()`](src/core/Document.ts:5043-5066) - Formats single TOC entry
- ✅ [`parseTOCFieldInstruction()`](src/core/Document.ts:4666-4722) - Parses field code
- ✅ [`escapeXml()`](src/core/Document.ts:5074-5081) - XML escaping

## Simple Implementation Steps

### Step 1: Add Flag (5 lines)

```typescript
private autoPopulateTOCs: boolean = false;

public setAutoPopulateTOCs(enabled: boolean): this {
  this.autoPopulateTOCs = enabled;
  return this;
}
```

### Step 2: Extract Method (30 lines)

Extract TOC population loop from `replaceTableOfContents()` into `populateAllTOCsInXML()`.

### Step 3: Integrate into Save (10 lines)

```typescript
// In save() method, after zipHandler.save(tempPath):
if (this.autoPopulateTOCs) {
  const handler = new ZipHandler();
  await handler.load(tempPath);
  const docXml = handler.getFileAsString("word/document.xml");
  if (docXml) {
    const populatedXml = this.populateAllTOCsInXML(docXml);
    handler.updateFile("word/document.xml", populatedXml);
    await handler.save(tempPath);
  }
}
```

### Step 4: Add Convenience Method (10 lines)

```typescript
public createPrePopulatedTableOfContents(
  title?: string,
  options?: Partial<TOCProperties>
): this {
  this.createTableOfContents(title);
  this.setAutoPopulateTOCs(true);
  return this;
}
```

**Total New Code**: ~60 lines

## Usage Examples

### Example 1: Simple Pre-Populated TOC

```typescript
const doc = Document.create();

doc.createParagraph("Chapter 1: Introduction").setStyle("Heading1");
doc.createParagraph("Section 1.1: Overview").setStyle("Heading2");
doc.createParagraph("Chapter 2: Methods").setStyle("Heading1");

// Single method call
doc.createPrePopulatedTableOfContents();

await doc.save("output.docx");
// Opens with TOC entries already visible!
```

### Example 2: Manual Control

```typescript
const doc = Document.create();

// Add headings first
doc.createParagraph("Introduction").setStyle("Heading1");
doc.createParagraph("Methods").setStyle("Heading1");

// Create TOC
doc.createTableOfContents("Contents");

// Enable auto-population
doc.setAutoPopulateTOCs(true);

await doc.save("output.docx");
```

### Example 3: Custom Options

```typescript
doc.createPrePopulatedTableOfContents("Table of Contents", {
  levels: 4,
  useHyperlinks: true,
  hideInWebLayout: true,
  tabLeader: "dot",
});

await doc.save("output.docx");
```

## Field Structure After Population

```xml
<w:sdt>
  <w:sdtPr><!-- Standard SDT properties --></w:sdtPr>
  <w:sdtContent>

    <!-- First paragraph: Field markers + first entry -->
    <w:p>
      <w:pPr>
        <w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/>
      </w:pPr>

      <!-- FIELD BEGIN -->
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>

      <!-- FIELD INSTRUCTION -->
      <w:r>
        <w:instrText xml:space="preserve">TOC \o "1-3" \h \* MERGEFORMAT</w:instrText>
      </w:r>

      <!-- FIELD SEPARATOR -->
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>

      <!-- FIRST ENTRY (actual heading, not placeholder!) -->
      <w:hyperlink w:anchor="_Toc123">
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/>
            <w:color w:val="0000FF"/>
            <w:sz w:val="24"/>
            <w:u w:val="single"/>
          </w:rPr>
          <w:t>Introduction</w:t>
        </w:r>
      </w:hyperlink>
    </w:p>

    <!-- Subsequent entries (one paragraph per entry) -->
    <w:p>
      <w:pPr>
        <w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/>
        <w:ind w:left="360"/>  <!-- Indented for Heading2 -->
      </w:pPr>
      <w:hyperlink w:anchor="_Toc124">
        <w:r>
          <w:rPr><!-- Same formatting --></w:rPr>
          <w:t>Background</w:t>
        </w:r>
      </w:hyperlink>
    </w:p>

    <!-- More entries... -->

    <!-- Final paragraph: Field end marker -->
    <w:p>
      <w:pPr>
        <w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/>
      </w:pPr>
      <!-- FIELD END -->
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>

  </w:sdtContent>
</w:sdt>
```

## Testing Strategy

### Unit Tests

```typescript
describe("TOC Pre-Population", () => {
  it("should populate TOC with heading entries", async () => {
    const doc = Document.create();
    doc.createParagraph("Test").setStyle("Heading1");
    doc.setAutoPopulateTOCs(true);
    doc.createTableOfContents();

    const buffer = await doc.toBuffer();
    const xml = buffer.toString("utf-8");

    expect(xml).toContain("Test");
    expect(xml).not.toContain("Right-click to update");
  });

  it("should preserve field structure", async () => {
    // Verify all 5 field components present
  });

  it("should create bookmarks for headings", async () => {
    // Verify bookmark creation
  });
});
```

### Integration Tests

- Save document and verify TOC entries visible
- Open in Word (manual test) and verify "Update Field" works
- Test with multiple heading levels
- Test with custom styles

## Rollout Plan

### Phase 1: Core Implementation (Week 1)

- [ ] Add auto-population flag
- [ ] Extract population methods
- [ ] Integrate into save/toBuffer
- [ ] Unit tests

### Phase 2: API Enhancement (Week 2)

- [ ] Add convenience method
- [ ] Create examples
- [ ] Integration tests

### Phase 3: Documentation (Week 3)

- [ ] Update README
- [ ] API documentation
- [ ] User guide

**Total Timeline**: 3 weeks (including testing and documentation)

## Risk Assessment

| Risk                         | Impact | Mitigation                                     |
| ---------------------------- | ------ | ---------------------------------------------- |
| Field structure corruption   | High   | Comprehensive validation in tests              |
| Backward compatibility break | Medium | Auto-populate is opt-in only                   |
| Performance degradation      | Low    | Population only on save, reuses efficient code |

## Conclusion

This is a straightforward enhancement that:

- ✅ Reuses 95% of existing, tested code
- ✅ Maintains backward compatibility
- ✅ Provides clear opt-in API
- ✅ Preserves field updateability in Word
- ✅ Estimated ~90 minutes of implementation time

**Ready to proceed with implementation in Code mode.**
