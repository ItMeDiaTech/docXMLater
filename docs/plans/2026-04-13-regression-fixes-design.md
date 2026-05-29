# Regression Fixes: Theme Fonts, Hyperlink Parsing, Image Cropping

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Fix three regressions: style theme font contamination in docxmlater, hyperlink-with-revisions not editable in docxmlater, and image border crop not removing whitespace+border in dochub-app.

**Architecture:** Fix A modifies the `applyStyles()` merge in `Document.ts` to strip theme font attributes when an explicit font is provided. Fix B modifies `DocumentParser.ts` to flatten revision children inside hyperlinks before parsing (instead of wrapping as PreservedElement). Fix C modifies `ImageBorderCropper.ts` scanLine to resume scanning past the border into the whitespace gap before returning the crop position, and removes the safety margin pullback.

**Tech Stack:** TypeScript, Jest, docxmlater (OOXML framework), dochub-app (Electron app)

---

## Task 1: Strip theme font attributes in applyStyles() merge

**Files:**

- Modify: `src/core/Document.ts:6860-6916` (the five config merges in `applyStyles()`)
- Test: `tests/core/ApplyStylesThemeFont.test.ts`

### Step 1: Write the failing test

Create `tests/core/ApplyStylesThemeFont.test.ts`:

```typescript
/**
 * Regression test: applyStyles() must strip theme font attributes from style
 * definitions when user provides an explicit font. Theme fonts (fontAsciiTheme,
 * fontHAnsiTheme) override explicit font names in Word, causing the user's
 * chosen font to be ignored.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function createDocxWithThemeFontStyle(): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`
  );

  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );

  // ListParagraph style with theme font (minorHAnsi = Calibri in most themes)
  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:rPr>
      <w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>
      <w:sz w:val="22"/>
    </w:rPr>
    <w:pPr>
      <w:spacing w:line="256" w:lineRule="auto"/>
      <w:ind w:left="720"/>
      <w:contextualSpacing/>
    </w:pPr>
  </w:style>
</w:styles>`
  );

  // A ListParagraph paragraph with direct run formatting (Verdana 12pt)
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:sz w:val="24"/></w:rPr>
        <w:t>List item text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
  );

  return await zipHandler.toBuffer();
}

describe('applyStyles theme font stripping', () => {
  it('should strip theme font attributes when explicit font is provided', async () => {
    const buffer = await createDocxWithThemeFontStyle();
    const doc = await Document.loadFromBuffer(buffer);

    // Apply styles with explicit Verdana font for ListParagraph
    doc.applyStyles({
      listParagraph: {
        run: { font: 'Verdana', size: 12, color: '000000' },
        paragraph: {
          spacing: { before: 0, after: 120, line: 240, lineRule: 'auto' as const },
        },
      },
    });

    // Verify the style definition no longer has theme font attributes
    const sm = doc.getStylesManager();
    const lpStyle = sm.getStyle('ListParagraph');
    const runFmt = lpStyle!.getRunFormatting()!;

    expect(runFmt.font).toBe('Verdana');
    expect(runFmt.fontAsciiTheme).toBeUndefined();
    expect(runFmt.fontHAnsiTheme).toBeUndefined();

    doc.dispose();
  });

  it('should strip theme fonts from all style configs (h1, h2, h3, normal, listParagraph)', async () => {
    const buffer = await createDocxWithThemeFontStyle();
    const doc = await Document.loadFromBuffer(buffer);

    doc.applyStyles({
      normal: {
        run: { font: 'Arial', size: 11 },
        paragraph: { spacing: { before: 0, after: 0 } },
      },
    });

    const sm = doc.getStylesManager();
    const normalStyle = sm.getStyle('Normal');
    const runFmt = normalStyle!.getRunFormatting()!;

    expect(runFmt.font).toBe('Arial');
    // Even if original Normal had no theme fonts, verify no contamination
    expect(runFmt.fontAsciiTheme).toBeUndefined();
    expect(runFmt.fontHAnsiTheme).toBeUndefined();

    doc.dispose();
  });

  it('should preserve theme fonts when no explicit font is provided', async () => {
    const buffer = await createDocxWithThemeFontStyle();
    const doc = await Document.loadFromBuffer(buffer);

    // Apply styles WITHOUT specifying a font — theme fonts should survive
    doc.applyStyles({
      listParagraph: {
        run: { size: 14, color: 'FF0000' },
        paragraph: { spacing: { before: 0, after: 120 } },
      },
    });

    const sm = doc.getStylesManager();
    const lpStyle = sm.getStyle('ListParagraph');
    const runFmt = lpStyle!.getRunFormatting()!;

    expect(runFmt.fontAsciiTheme).toBe('minorHAnsi');
    expect(runFmt.fontHAnsiTheme).toBe('minorHAnsi');

    doc.dispose();
  });
});
```

### Step 2: Run test to verify it fails

Run: `npx jest tests/core/ApplyStylesThemeFont.test.ts --no-coverage`
Expected: FAIL — `fontAsciiTheme` is still `"minorHAnsi"` instead of `undefined`

### Step 3: Implement the fix

In `src/core/Document.ts`, add a helper function before `applyStyles()` (around line 6838):

```typescript
/**
 * Strip theme font attributes from run config when an explicit font is set.
 * In OOXML, theme fonts (w:asciiTheme, w:hAnsiTheme) override explicit font
 * names (w:ascii, w:hAnsi) — so setting font: 'Verdana' is meaningless if
 * fontAsciiTheme: 'minorHAnsi' is also present.
 */
private static stripThemeFontsIfExplicitFont(runConfig: any): void {
  if (runConfig?.font) {
    delete runConfig.fontAsciiTheme;
    delete runConfig.fontHAnsiTheme;
    delete runConfig.fontEastAsiaTheme;
    delete runConfig.fontCsTheme;
  }
}
```

Then call it on each merged config, right after the five config blocks are built (after line 6916, before line 6918):

```typescript
// Strip theme font attributes that would override explicit font names
Document.stripThemeFontsIfExplicitFont(h1Config.run);
Document.stripThemeFontsIfExplicitFont(h2Config.run);
Document.stripThemeFontsIfExplicitFont(h3Config.run);
Document.stripThemeFontsIfExplicitFont(normalConfig.run);
Document.stripThemeFontsIfExplicitFont(listParaConfig.run);
```

### Step 4: Run test to verify it passes

Run: `npx jest tests/core/ApplyStylesThemeFont.test.ts --no-coverage`
Expected: PASS

### Step 5: Run full test suite

Run: `npm test`
Expected: All existing tests still pass

### Step 6: Commit

```bash
git add src/core/Document.ts tests/core/ApplyStylesThemeFont.test.ts
git commit -m "Fix applyStyles() theme font contamination: strip fontAsciiTheme/fontHAnsiTheme when explicit font is set"
```

---

## Task 2: Parse hyperlinks with revisions as Hyperlink (not PreservedElement)

**Files:**

- Modify: `src/core/DocumentParser.ts:1339-1363`
- Modify: `tests/core/HyperlinkRevisionPreservation.test.ts` (update expectations)
- Test: `tests/core/HyperlinkRevisionFlattening.test.ts` (new)

### Step 1: Write the failing test

Create `tests/core/HyperlinkRevisionFlattening.test.ts`:

```typescript
/**
 * Tests that hyperlinks containing tracked changes (w:ins/w:del) are parsed
 * as editable Hyperlink objects by flattening the revisions.
 *
 * Replaces the old PreservedElement behavior. The dochub-app Power Automate
 * pipeline needs setUrl()/setText() on these hyperlinks.
 */
import { Document } from '../../src/core/Document';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { PreservedElement } from '../../src/elements/PreservedElement';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function createDocxWithHyperlinks(
  documentXml: string,
  rels: { id: string; target: string }[] = []
): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );

  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  const relEntries = rels
    .map(
      (r) =>
        `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${r.target}" TargetMode="External"/>`
    )
    .join('\n  ');

  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${relEntries}
</Relationships>`
  );

  zipHandler.addFile('word/document.xml', documentXml);
  return await zipHandler.toBuffer();
}

describe('Hyperlink Revision Flattening', () => {
  it('should parse hyperlink with w:ins/w:del as Hyperlink (keeping inserted text, dropping deleted)', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:del w:id="10" w:author="User" w:date="2024-01-01T00:00:00Z">
          <w:r><w:delText>old title</w:delText></w:r>
        </w:del>
        <w:ins w:id="11" w:author="User" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>new title</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxWithHyperlinks(documentXml, [
      { id: 'rId5', target: 'https://thesource.example.com/doc/123456' },
    ]);

    const doc = await Document.loadFromBuffer(buffer, {
      revisionHandling: 'preserve',
    });

    const content = doc.getParagraphs()[0]!.getContent();

    // Should be a Hyperlink, NOT a PreservedElement
    const hyperlinks = content.filter((item) => item instanceof Hyperlink);
    const preserved = content.filter((item) => item instanceof PreservedElement);
    expect(hyperlinks.length).toBe(1);
    expect(preserved.length).toBe(0);

    // The hyperlink text should be the inserted text (not the deleted text)
    const hyperlink = hyperlinks[0] as Hyperlink;
    expect(hyperlink.getText()).toBe('new title');
    expect(hyperlink.getUrl()).toBe('https://thesource.example.com/doc/123456');

    // Should be editable
    hyperlink.setText('updated title (123456)');
    expect(hyperlink.getText()).toBe('updated title (123456)');

    hyperlink.setUrl('https://thesource.example.com/doc/789');
    expect(hyperlink.getUrl()).toBe('https://thesource.example.com/doc/789');

    doc.dispose();
  });

  it('should parse hyperlink with only w:ins (no w:del) as Hyperlink', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:ins w:id="11" w:author="User" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>inserted link</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxWithHyperlinks(documentXml, [
      { id: 'rId5', target: 'https://example.com' },
    ]);

    const doc = await Document.loadFromBuffer(buffer, {
      revisionHandling: 'preserve',
    });

    const content = doc.getParagraphs()[0]!.getContent();
    const hyperlinks = content.filter((item) => item instanceof Hyperlink);
    expect(hyperlinks.length).toBe(1);
    expect((hyperlinks[0] as Hyperlink).getText()).toBe('inserted link');

    doc.dispose();
  });

  it('should handle hyperlink with w:ins and direct w:r children', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:r><w:t>kept text </w:t></w:r>
        <w:ins w:id="11" w:author="User" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>added text</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxWithHyperlinks(documentXml, [
      { id: 'rId5', target: 'https://example.com' },
    ]);

    const doc = await Document.loadFromBuffer(buffer, {
      revisionHandling: 'preserve',
    });

    const content = doc.getParagraphs()[0]!.getContent();
    const hyperlinks = content.filter((item) => item instanceof Hyperlink);
    expect(hyperlinks.length).toBe(1);
    // Should contain both the direct run text and the inserted text
    expect((hyperlinks[0] as Hyperlink).getText()).toContain('kept text');

    doc.dispose();
  });
});
```

### Step 2: Run test to verify it fails

Run: `npx jest tests/core/HyperlinkRevisionFlattening.test.ts --no-coverage`
Expected: FAIL — hyperlinks with revisions are still PreservedElement

### Step 3: Implement the fix

In `src/core/DocumentParser.ts`, replace lines 1339-1350 (the `hasRevisionChildren` block):

```typescript
          // Hyperlinks containing tracked changes (w:del/w:ins inside w:hyperlink):
          // Flatten revisions to make the hyperlink editable (setUrl/setText).
          // - w:ins runs are unwrapped (kept as direct w:r children)
          // - w:del runs are dropped (deleted content)
          // - w:moveFrom runs are dropped, w:moveTo runs are unwrapped
          // This trades revision fidelity inside the hyperlink for editability,
          // which is required by the dochub-app Power Automate pipeline.
          const hasRevisionChildren =
            hyperlinkObj['w:del'] ||
            hyperlinkObj['w:ins'] ||
            hyperlinkObj['w:moveFrom'] ||
            hyperlinkObj['w:moveTo'];
          if (hasRevisionChildren) {
            // Build a flattened copy with revisions resolved
            const flattenedObj = { ...hyperlinkObj };

            // Collect all runs: direct w:r + unwrapped from w:ins/w:moveTo
            const allRuns: any[] = [];

            // Keep existing direct runs
            if (flattenedObj['w:r']) {
              const directRuns = Array.isArray(flattenedObj['w:r'])
                ? flattenedObj['w:r']
                : [flattenedObj['w:r']];
              allRuns.push(...directRuns);
            }

            // Unwrap w:ins runs (inserted content — keep)
            if (flattenedObj['w:ins']) {
              const insArr = Array.isArray(flattenedObj['w:ins'])
                ? flattenedObj['w:ins']
                : [flattenedObj['w:ins']];
              for (const ins of insArr) {
                if (ins['w:r']) {
                  const insRuns = Array.isArray(ins['w:r']) ? ins['w:r'] : [ins['w:r']];
                  allRuns.push(...insRuns);
                }
              }
            }

            // Unwrap w:moveTo runs (move destination — keep)
            if (flattenedObj['w:moveTo']) {
              const moveToArr = Array.isArray(flattenedObj['w:moveTo'])
                ? flattenedObj['w:moveTo']
                : [flattenedObj['w:moveTo']];
              for (const mt of moveToArr) {
                if (mt['w:r']) {
                  const mtRuns = Array.isArray(mt['w:r']) ? mt['w:r'] : [mt['w:r']];
                  allRuns.push(...mtRuns);
                }
              }
            }

            // Drop w:del and w:moveFrom (deleted/moved-away content)
            // by simply not including their runs

            // Replace runs and remove revision wrappers
            flattenedObj['w:r'] = allRuns.length > 0 ? allRuns : undefined;
            delete flattenedObj['w:del'];
            delete flattenedObj['w:ins'];
            delete flattenedObj['w:moveFrom'];
            delete flattenedObj['w:moveTo'];

            const result = this.parseHyperlinkFromObject(flattenedObj, relationshipManager);
            if (result.hyperlink) {
              paragraph.addHyperlink(result.hyperlink);
            }
            for (const bookmark of result.bookmarkStarts) {
              paragraph.addBookmarkStart(bookmark);
            }
            for (const bookmark of result.bookmarkEnds) {
              paragraph.addBookmarkEnd(bookmark);
            }
          } else {
```

### Step 4: Run test to verify it passes

Run: `npx jest tests/core/HyperlinkRevisionFlattening.test.ts --no-coverage`
Expected: PASS

### Step 5: Update existing HyperlinkRevisionPreservation tests

The existing `tests/core/HyperlinkRevisionPreservation.test.ts` expects PreservedElement — update to expect Hyperlink instead. The test at line 93-98 should now expect a Hyperlink, not PreservedElement.

Update the first test (`should preserve tracked changes inside hyperlinks during round-trip`) to expect a Hyperlink with the inserted text. The round-trip test should verify the output contains the hyperlink with the text "new link text" (the accepted/flattened version).

Update the second test (`should preserve w:moveFrom and w:moveTo inside hyperlinks`) similarly.

### Step 6: Run full test suite

Run: `npm test`
Expected: All tests pass (existing hyperlink tests may need adjustment)

### Step 7: Commit

```bash
git add src/core/DocumentParser.ts tests/core/HyperlinkRevisionFlattening.test.ts tests/core/HyperlinkRevisionPreservation.test.ts
git commit -m "Flatten revisions inside hyperlinks during parsing for editability (setUrl/setText)"
```

---

## Task 3: Fix ImageBorderCropper scanLine to crop whitespace + border

**Files:**

- Modify: `C:\Users\DiaTech\Projects\DocHub\development\dochub-app\src\services\document\helpers\ImageBorderCropper.ts:418-469` (scanLine function)
- Modify: `C:\Users\DiaTech\Projects\DocHub\development\dochub-app\src\services\document\helpers\ImageBorderCropper.ts:300-315` (detectEmbeddedBorder — remove CONTENT_SAFETY_MARGIN)
- Test: `C:\Users\DiaTech\Projects\DocHub\development\dochub-app\src\services\document\helpers\__tests__\ImageBorderCropper.test.ts`

### Step 1: Write the failing test

Add to the existing `ImageBorderCropper.test.ts`:

```typescript
describe('scanLine whitespace+border cropping', () => {
  it('should crop past whitespace AND border to content edge', () => {
    // Layout: 10px whitespace, 2px dark border, then content
    const W = 200;
    const H = 200;
    const pixels = makePixels(W, H, WHITE);

    // Top edge: rows 0-9 white (whitespace), rows 10-11 dark (border), rows 12+ content
    fillRow(pixels, W, 10, DARK);
    fillRow(pixels, W, 11, DARK);

    const result = scanLine(pixels, W, H, 'top', 50);

    // Should crop everything from edge through border: crop at 12 (past border)
    // NOT at 11-CONTENT_SAFETY_MARGIN = 8 (inside whitespace)
    expect(result).toBe(12);
  });

  it('should crop border at edge with no whitespace', () => {
    // Layout: 2px dark border at edge, then content
    const W = 200;
    const H = 200;
    const pixels = makePixels(W, H, WHITE);

    fillRow(pixels, W, 0, DARK);
    fillRow(pixels, W, 1, DARK);

    const result = scanLine(pixels, W, H, 'top', 50);

    // Should crop right past the border
    expect(result).toBe(2);
  });

  it('should crop whitespace + border + gap to content', () => {
    // Layout: 5px whitespace, 1px dark border, 5px whitespace gap, then content
    const W = 200;
    const H = 200;
    const pixels = makePixels(W, H, WHITE);

    fillRow(pixels, W, 5, DARK); // border at depth 5

    const result = scanLine(pixels, W, H, 'top', 50);

    // Should crop past the border AND the whitespace gap
    // The gap between border and content should also be removed
    expect(result).toBeGreaterThan(5);
  });
});
```

### Step 2: Run test to verify it fails

Run: `cd C:\Users\DiaTech\Projects\DocHub\development\dochub-app && npx jest src/services/document/helpers/__tests__/ImageBorderCropper.test.ts --no-coverage`
Expected: FAIL — scanLine returns `lastDarkDepth + 1` minus safety margin, not the full crop

### Step 3: Implement the fix

Replace the `scanLine` function in `ImageBorderCropper.ts` (lines 418-469):

```typescript
export function scanLine(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  edge: Edge,
  lineIndex: number
): number | null {
  const depthDimension = edge === 'top' || edge === 'bottom' ? height : width;
  const maxDepth = Math.floor(depthDimension * MAX_CROP_FRACTION);

  // ── Phase 1: Find first dark pixel (skip initial white gap) ──
  let firstDarkDepth = -1;
  const skipLimit = Math.min(MAX_INITIAL_SKIP + 1, maxDepth);
  for (let depth = 0; depth < skipLimit; depth++) {
    if (getPixelLuminance(pixels, width, height, edge, lineIndex, depth) <= DARK_THRESHOLD) {
      firstDarkDepth = depth;
      break;
    }
  }
  if (firstDarkDepth === -1) return null; // no border pixels found within skip zone

  // ── Phase 2: Scan border zone anchored to firstDarkDepth ──
  let lastDarkDepth = firstDarkDepth;
  let darkCount = 1;
  const borderZoneEnd = Math.min(firstDarkDepth + MAX_BORDER_ZONE, maxDepth);
  for (let depth = firstDarkDepth + 1; depth < borderZoneEnd; depth++) {
    if (getPixelLuminance(pixels, width, height, edge, lineIndex, depth) <= DARK_THRESHOLD) {
      lastDarkDepth = depth;
      darkCount++;
    }
  }
  if (darkCount > MAX_BORDER_THICKNESS) return null; // too many dark pixels for a border

  // ── Phase 3: Verify border ended ─────────────────────────────────
  // Need MIN_POST_BORDER_NONDARK consecutive non-dark pixels after last dark
  let nondarkRun = 0;
  let borderEndConfirmed = false;
  for (let depth = lastDarkDepth + 1; depth < maxDepth; depth++) {
    const lum = getPixelLuminance(pixels, width, height, edge, lineIndex, depth);
    if (lum > DARK_THRESHOLD) {
      nondarkRun++;
      if (nondarkRun >= MIN_POST_BORDER_NONDARK) {
        borderEndConfirmed = true;
        break;
      }
    } else {
      return null; // more dark pixels beyond border zone — not a clean border
    }
  }
  if (!borderEndConfirmed) return null;

  // ── Phase 4: Skip whitespace gap past border to reach content ──
  // Scan from just past the border until we either hit non-white content
  // or reach maxDepth. Crop everything up to this point.
  // This removes: initial whitespace + border + post-border gap.
  let cropDepth = lastDarkDepth + 1;
  for (let depth = lastDarkDepth + 1; depth < maxDepth; depth++) {
    const lum = getPixelLuminance(pixels, width, height, edge, lineIndex, depth);
    if (lum <= DARK_THRESHOLD) {
      // Hit dark content — stop before it
      cropDepth = depth;
      break;
    }
    cropDepth = depth + 1;
  }

  return cropDepth;
}
```

Also remove `CONTENT_SAFETY_MARGIN` from `detectEmbeddedBorder` (around line 303):

```typescript
// Before (remove safeCrop):
// const safeCrop = (pos: number): number => Math.max(0, pos - CONTENT_SAFETY_MARGIN);

// After: use crop positions directly
return {
  cropRect: {
    top: results.top.detected ? results.top.cropPosition : 0,
    bottom: results.bottom.detected ? results.bottom.cropPosition : 0,
    left: results.left.detected ? results.left.cropPosition : 0,
    right: results.right.detected ? results.right.cropPosition : 0,
  },
  detectedEdges: detectedCount,
};
```

Remove the `CONTENT_SAFETY_MARGIN` constant export (line 31) and its import from the test file.

### Step 4: Run test to verify it passes

Run: `cd C:\Users\DiaTech\Projects\DocHub\development\dochub-app && npx jest src/services/document/helpers/__tests__/ImageBorderCropper.test.ts --no-coverage`
Expected: PASS

### Step 5: Run full dochub-app test suite

Run: `cd C:\Users\DiaTech\Projects\DocHub\development\dochub-app && npm test`
Expected: All tests pass (some existing scanLine tests may need updated expected values)

### Step 6: Commit

```bash
cd C:\Users\DiaTech\Projects\DocHub\development\dochub-app
git add src/services/document/helpers/ImageBorderCropper.ts src/services/document/helpers/__tests__/ImageBorderCropper.test.ts
git commit -m "Fix image border crop: scan past whitespace+border to content, remove safety margin pullback"
```

---

## Task 4: Integration verification with actual documents

### Step 1: Verify theme font fix with Original.docx

Write a quick Node.js script in docxmlater root:

```bash
node -e "
const { Document } = require('./dist/index.js');
const fs = require('fs');
async function main() {
  const buf = fs.readFileSync('Original.docx');
  const doc = await Document.loadFromBuffer(buf);
  doc.applyStyles({
    listParagraph: {
      run: { font: 'Verdana', size: 12, color: '000000' },
      paragraph: { spacing: { before: 0, after: 120, line: 240, lineRule: 'auto' } },
    },
    normal: {
      run: { font: 'Verdana', size: 12, color: '000000' },
      paragraph: { spacing: { before: 60, after: 60, line: 240, lineRule: 'auto' } },
    },
  });
  const sm = doc.getStylesManager();
  const lp = sm.getStyle('ListParagraph');
  const fmt = lp.getRunFormatting();
  console.log('LP font:', fmt.font);
  console.log('LP fontAsciiTheme:', fmt.fontAsciiTheme);
  console.log('LP fontHAnsiTheme:', fmt.fontHAnsiTheme);
  // Check a LP paragraph's runs
  for (const p of doc.getAllParagraphs()) {
    if (p.getStyle() === 'ListParagraph') {
      const runs = p.getRuns();
      if (runs.length > 0) {
        console.log('First LP run font:', runs[0].getFormatting().font);
        console.log('First LP run size:', runs[0].getFormatting().size);
      }
      break;
    }
  }
  doc.dispose();
}
main();
"
```

Expected output:

```
LP font: Verdana
LP fontAsciiTheme: undefined
LP fontHAnsiTheme: undefined
First LP run font: undefined (inherits Verdana from style)
First LP run size: undefined (inherits 12 from style)
```

### Step 2: Verify hyperlink fix

```bash
node -e "
const { Document } = require('./dist/index.js');
const fs = require('fs');
async function main() {
  const buf = fs.readFileSync('Original.docx');
  const doc = await Document.loadFromBuffer(buf, { revisionHandling: 'preserve' });
  let hyperlinkCount = 0;
  let preservedCount = 0;
  for (const p of doc.getAllParagraphs()) {
    for (const item of p.getContent()) {
      if (item.constructor.name === 'Hyperlink') hyperlinkCount++;
      if (item.constructor.name === 'PreservedElement' &&
          item.getElementType() === 'w:hyperlink') preservedCount++;
    }
  }
  console.log('Hyperlinks:', hyperlinkCount);
  console.log('PreservedElement hyperlinks:', preservedCount);
  doc.dispose();
}
main();
"
```

Expected: `PreservedElement hyperlinks: 0` (all parsed as Hyperlink)

---

## Execution Order

1. **Task 1** (docxmlater: theme font fix) — independent, can start immediately
2. **Task 2** (docxmlater: hyperlink parsing) — independent, can start immediately
3. **Task 3** (dochub-app: image cropper) — independent, can start immediately
4. **Task 4** (integration verification) — depends on Tasks 1-3

Tasks 1, 2, and 3 are fully independent and can be done in parallel.
