/**
 * Paragraph `w:pStyle` and run `w:rStyle` — style-ID references with
 * purely-numeric values must round-trip as strings (type-contract
 * safety).
 *
 * Per ECMA-376 Part 1 §17.3.1.27 (pStyle) and §17.3.2.36 (rStyle),
 * the `w:val` attribute is `ST_String` — a reference to a style ID.
 * Style IDs are usually descriptive strings like `"Heading1"` or
 * `"Hyperlink"`, but custom styles can legitimately use any string
 * including purely-numeric ones like `"1"`, `"42"`, or `"2025"`.
 *
 * XMLParser's `parseAttributeValue: true` coerces purely-numeric
 * attribute values to JS numbers. Previously the parsers stored the
 * coerced value:
 *
 *     paragraph.setStyle(pPrObj['w:pStyle']['@_w:val']);  // could be 1 (number)
 *     run.setCharacterStyle(styleId);                      // could be 42 (number)
 *
 * Both `setStyle` and `setCharacterStyle` type the argument as
 * `string`, so storing a number violated the type contract.
 * Downstream code calling `.startsWith(...)`, `.toLowerCase()`, or
 * template-string concatenation on the styleId would either throw
 * (string methods on number) or produce wrong output (template
 * literals coerce but some comparisons would fail).
 *
 * Iteration 126 casts both style references via `String(...)`.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadFirstParagraph(pPrInner: string) {
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
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>${pPrInner}</w:pPr>
      <w:r>
        <w:rPr><w:rStyle w:val="42"/></w:rPr>
        <w:t>x</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const paragraph = doc.getParagraphs()[0]!;
  const styleId = paragraph.getStyle();
  const runStyleId = paragraph.getRuns()[0]?.getFormatting().characterStyle;
  doc.dispose();
  return { styleId, runStyleId };
}

describe('<w:pStyle> / <w:rStyle> numeric style-ID type contract', () => {
  it('stores <w:pStyle w:val="1"/> as the STRING "1" (not number 1)', async () => {
    const { styleId } = await loadFirstParagraph('<w:pStyle w:val="1"/>');
    expect(typeof styleId).toBe('string');
    expect(styleId).toBe('1');
  });

  it('stores <w:pStyle w:val="2025"/> as the STRING "2025"', async () => {
    const { styleId } = await loadFirstParagraph('<w:pStyle w:val="2025"/>');
    expect(typeof styleId).toBe('string');
    expect(styleId).toBe('2025');
  });

  it('stores <w:rStyle w:val="42"/> as the STRING "42"', async () => {
    const { runStyleId } = await loadFirstParagraph('<w:pStyle w:val="Normal"/>');
    expect(typeof runStyleId).toBe('string');
    expect(runStyleId).toBe('42');
  });

  it('preserves non-numeric style IDs (regression guard)', async () => {
    const { styleId } = await loadFirstParagraph('<w:pStyle w:val="Heading1"/>');
    expect(styleId).toBe('Heading1');
  });

  it('string methods are callable on parsed numeric styleId', async () => {
    const { styleId } = await loadFirstParagraph('<w:pStyle w:val="2025"/>');
    // Pre-fix: styleId was number 2025, .startsWith would throw.
    expect(() => styleId?.startsWith('20')).not.toThrow();
    expect(styleId?.startsWith('20')).toBe(true);
  });
});
