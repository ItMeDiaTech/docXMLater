/**
 * `<w:pPrChange w:id="0" …>` — tracked paragraph-property change with
 * revision ID 0 must round-trip with the ID intact.
 *
 * Per ECMA-376 Part 1 §17.13.5.27 CT_TrackChange, `w:id` is
 * ST_DecimalNumber (xsd:integer) and REQUIRED. Any integer including
 * `0` is a legal revision identifier.
 *
 * Two compounding bugs silently dropped `w:id="0"` on load:
 *   1. XMLParser coerces `"0"` → number `0`.
 *   2. The pPrChange parser at `DocumentParser.ts:2630` used
 *      `if (changeObj['@_w:id']) change.id = String(…)`, where the
 *      coerced numeric zero fails the truthy gate.
 *
 * The emitter at `Paragraph.ts:3589` re-emits `if (change.id) …` —
 * also truthy-gated, so the lost id never reappears on save. Result:
 * every `<w:pPrChange w:id="0" …/>` emitted as
 * `<w:pPrChange w:author="…" w:date="…">` — missing the REQUIRED
 * `w:id` attribute, failing strict OOXML validation, and losing the
 * revision linkage Word uses to correlate pPrChange instances with
 * ins / del counterparts in the same document.
 *
 * Sibling parsers (`trPrChange` / `tblPrChange` / `tcPrChange` /
 * `sectPrChange`) already handle id=0 via `|| '0'` or `!== undefined`
 * — only pPrChange was out of step.
 *
 * Iteration 113 gates on `!== undefined` so id=0 survives.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadAndResaveDocXml(xml: string): Promise<string> {
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
  zipHandler.addFile('word/document.xml', xml);
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const content = zip.getFile('word/document.xml')?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

describe('<w:pPrChange w:id="…"> zero-id round-trip', () => {
  it('preserves w:id="0" on pPrChange', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:pPrChange w:id="0" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    // Previously the output had `<w:pPrChange w:author="Tester"
    // w:date="…">` with no w:id. Now w:id="0" round-trips.
    expect(out).toMatch(/<w:pPrChange[^>]*w:id="0"/);
  });

  it('preserves w:id="42" on pPrChange (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:pPrChange w:id="42" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    expect(out).toMatch(/<w:pPrChange[^>]*w:id="42"/);
  });

  it('still preserves author and date alongside id=0', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:pPrChange w:id="0" w:author="Zed" w:date="2026-02-02T12:00:00Z">
          <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const block = out.match(/<w:pPrChange[^>]*>/)?.[0] ?? '';
    expect(block).toMatch(/w:id="0"/);
    expect(block).toMatch(/w:author="Zed"/);
    expect(block).toMatch(/w:date="2026-02-02T12:00:00Z"/);
  });
});
