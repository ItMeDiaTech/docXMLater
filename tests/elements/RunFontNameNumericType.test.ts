/**
 * `<w:rFonts>` literal-font attributes — numeric-looking values must
 * round-trip as strings (type-contract safety).
 *
 * Per ECMA-376 Part 1 §17.3.2.26 CT_Fonts, the four literal-font
 * attributes (`w:ascii`, `w:hAnsi`, `w:eastAsia`, `w:cs`) are all
 * `ST_String`. ECMA-376 Part 4 schema confirms ST_String = xsd:string
 * so any string is valid including purely-numeric names (a numbered
 * custom font like `"2010"`, East-Asian fonts named `"1600"`, etc.).
 *
 * Bug: XMLParser's `parseAttributeValue: true` coerces purely-numeric
 * strings to JS numbers. The main-path parser (DocumentParser.ts
 * §5405) and the `rPrChange` previous-font parser (§5545) both store
 * the raw coerced value:
 *
 *     if (rFonts['@_w:ascii']) run.setFont(rFonts['@_w:ascii']);
 *     // ↑ receives number 2010 when the XML was w:ascii="2010"
 *
 * `RunFormatting.font`/`fontHAnsi`/`fontEastAsia`/`fontCs` are all
 * declared `string`. Storing a JS number violates the contract and
 * any downstream `.toLowerCase()` / `.startsWith(...)` on a resolved
 * font name would throw.
 *
 * Iteration 131 adds `String(...)` casts at every literal-font
 * attribute read site in both the main-path and rPrChange parsers.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Run } from '../../src/elements/Run';

async function loadRunWith(rPrInnerXml: string) {
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
      <w:r>
        <w:rPr>${rPrInnerXml}</w:rPr>
        <w:t>x</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  return Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
}

describe('<w:rFonts> literal-font attribute type-contract preservation', () => {
  it('stores <w:rFonts w:ascii="2010"/> as the STRING "2010"', async () => {
    const doc = await loadRunWith('<w:rFonts w:ascii="2010"/>');
    const run = doc.getParagraphs()[0]!.getRuns()[0] as Run;
    const font = run.getFormatting().font;
    doc.dispose();
    expect(typeof font).toBe('string');
    expect(font).toBe('2010');
  });

  it('stores <w:rFonts w:hAnsi="42"/> as the STRING "42"', async () => {
    const doc = await loadRunWith('<w:rFonts w:hAnsi="42"/>');
    const run = doc.getParagraphs()[0]!.getRuns()[0] as Run;
    const font = run.getFormatting().fontHAnsi;
    doc.dispose();
    expect(typeof font).toBe('string');
    expect(font).toBe('42');
  });

  it('stores <w:rFonts w:eastAsia="1600"/> as the STRING "1600"', async () => {
    const doc = await loadRunWith('<w:rFonts w:eastAsia="1600"/>');
    const run = doc.getParagraphs()[0]!.getRuns()[0] as Run;
    const font = run.getFormatting().fontEastAsia;
    doc.dispose();
    expect(typeof font).toBe('string');
    expect(font).toBe('1600');
  });

  it('stores <w:rFonts w:cs="99"/> as the STRING "99"', async () => {
    const doc = await loadRunWith('<w:rFonts w:cs="99"/>');
    const run = doc.getParagraphs()[0]!.getRuns()[0] as Run;
    const font = run.getFormatting().fontCs;
    doc.dispose();
    expect(typeof font).toBe('string');
    expect(font).toBe('99');
  });

  it('string methods are callable on parsed numeric font name', async () => {
    const doc = await loadRunWith('<w:rFonts w:ascii="2010"/>');
    const run = doc.getParagraphs()[0]!.getRuns()[0] as Run;
    const font = run.getFormatting().font as string;
    doc.dispose();
    // Pre-fix: font was number 2010, .startsWith would throw.
    expect(() => font.startsWith('20')).not.toThrow();
    expect(font.startsWith('20')).toBe(true);
  });

  it('preserves non-numeric font "Calibri" (regression guard)', async () => {
    const doc = await loadRunWith('<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>');
    const run = doc.getParagraphs()[0]!.getRuns()[0] as Run;
    const fmt = run.getFormatting();
    doc.dispose();
    expect(fmt.font).toBe('Calibri');
    expect(fmt.fontHAnsi).toBe('Calibri');
  });

  it('preserves numeric font in rPrChange previousProperties', async () => {
    const doc = await loadRunWith(
      `<w:rFonts w:ascii="Calibri"/>
       <w:rPrChange w:id="1" w:author="A" w:date="2026-04-24T10:00:00Z">
         <w:rPr><w:rFonts w:ascii="2010"/></w:rPr>
       </w:rPrChange>`
    );
    const run = doc.getParagraphs()[0]!.getRuns()[0] as Run;
    const rPrChange = run.getPropertyChangeRevision();
    doc.dispose();
    expect(rPrChange).toBeDefined();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const prevFont = (rPrChange!.previousProperties as any).font;
    expect(typeof prevFont).toBe('string');
    expect(prevFont).toBe('2010');
  });
});
