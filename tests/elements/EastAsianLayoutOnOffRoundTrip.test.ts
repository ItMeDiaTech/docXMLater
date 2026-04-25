/**
 * `<w:rPr><w:eastAsianLayout w:vert="…" w:vertCompress="…" w:combine="…">`
 * — ST_OnOff attribute round-trip.
 *
 * Per ECMA-376 Part 1 §17.3.2.10 CT_EastAsianLayout, the three
 * boolean attributes `w:vert`, `w:vertCompress`, and `w:combine` are
 * ST_OnOff — accept every literal ("1"/"0"/"true"/"false"/"on"/"off").
 *
 * Two compounding bugs silently dropped explicit-false overrides and
 * mishandled the "off" literal:
 *   1. The parser (main rPr at `DocumentParser.ts:5335`, rPrChange at
 *      `:5741`) used a truthy gate. XMLParser coerced `"0"` to the
 *      number `0`, which failed the gate — so `w:vert="0"` was
 *      dropped entirely. Worse, `w:vert="off"` stayed as the truthy
 *      string `"off"` and passed the gate, yielding `layout.vert = true`
 *      — the opposite of the author's intent.
 *   2. The emitter (Run.ts `generateRunPropertiesXML` §3106) re-emitted
 *      only when the value was truthy — explicit-false values stored
 *      in the model (`layout.vert = false`) were never written back.
 *
 * Iteration 119 routes all three through `parseOnOffAttribute` on
 * parse and changes the emitter guard to `!== undefined`, so
 * explicit-false vertical-typography overrides survive round-trip.
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
  const doc = await Document.loadFromBuffer(buffer);
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const content = zip.getFile('word/document.xml')?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

function buildRunDoc(eaAttrs: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:eastAsianLayout${eaAttrs}/>
        </w:rPr>
        <w:t>text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
}

function extractEastAsianLayout(xml: string): string {
  return xml.match(/<w:eastAsianLayout[^/]*\/>/)?.[0] ?? '';
}

describe('<w:eastAsianLayout> ST_OnOff attribute round-trip', () => {
  it('preserves w:vert="1" (baseline)', async () => {
    const out = await loadAndResaveDocXml(buildRunDoc(' w:vert="1"'));
    expect(extractEastAsianLayout(out)).toMatch(/w:vert="1"/);
  });

  it('preserves w:vert="0" — explicit-false override (previously dropped)', async () => {
    const out = await loadAndResaveDocXml(buildRunDoc(' w:vert="0"'));
    expect(extractEastAsianLayout(out)).toMatch(/w:vert="0"/);
  });

  it('honours w:vert="off" — correctly stores false (previously wrong)', async () => {
    // Previously "off" (truthy string) passed the truthy gate and was
    // incorrectly stored as true. Now parseOnOffAttribute resolves it
    // to false and we emit the canonical "0" form.
    const out = await loadAndResaveDocXml(buildRunDoc(' w:vert="off"'));
    expect(extractEastAsianLayout(out)).toMatch(/w:vert="0"/);
  });

  it('preserves w:combine="0" — explicit-false override', async () => {
    const out = await loadAndResaveDocXml(buildRunDoc(' w:combine="0"'));
    expect(extractEastAsianLayout(out)).toMatch(/w:combine="0"/);
  });

  it('preserves w:vertCompress="false" — explicit-false literal form', async () => {
    const out = await loadAndResaveDocXml(buildRunDoc(' w:vertCompress="false"'));
    expect(extractEastAsianLayout(out)).toMatch(/w:vertCompress="0"/);
  });

  it('omits all three when absent on input (regression guard)', async () => {
    const out = await loadAndResaveDocXml(buildRunDoc(' w:id="1"'));
    const tag = extractEastAsianLayout(out);
    expect(tag).toMatch(/w:id="1"/);
    expect(tag).not.toMatch(/w:vert=/);
    expect(tag).not.toMatch(/w:vertCompress=/);
    expect(tag).not.toMatch(/w:combine=/);
  });
});
