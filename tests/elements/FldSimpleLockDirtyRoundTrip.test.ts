/**
 * `<w:fldSimple>` — `w:fldLock` and `w:dirty` attribute round-trip.
 *
 * Per ECMA-376 Part 1 §17.16.16 CT_SimpleField, `<w:fldSimple>`
 * carries three attributes:
 *   - w:instr (required) — the field instruction
 *   - w:fldLock (CT_OnOff, optional) — field is locked against
 *     automatic updates
 *   - w:dirty (CT_OnOff, optional) — field's cached result is stale
 *     and Word should refresh on open
 *
 * Both ST_OnOff attributes were unsupported: the parser read only
 * `w:instr`, and the emitter wrote only `w:instr`. Any Word document
 * whose `<w:fldSimple>` carried `w:fldLock="1"` or `w:dirty="1"`
 * silently lost the flag on every load → save cycle — a user's
 * "update field" reminder or "lock field" override reverted to the
 * spec default ("absent ≡ false") on every round-trip.
 *
 * Iteration 110 adds parse + emit support for both attributes via
 * `parseOnOffAttribute`, honouring every ST_OnOff literal ("1" / "0"
 * / "true" / "false" / "on" / "off") on the parse side and writing
 * the canonical `"1"` / `"0"` form on the emit side.
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

function buildFldSimpleDoc(attrFragment: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:fldSimple w:instr="PAGE"${attrFragment}>
        <w:r><w:t>1</w:t></w:r>
      </w:fldSimple>
    </w:p>
  </w:body>
</w:document>`;
}

function extractFldSimpleOpenTag(xml: string): string {
  return xml.match(/<w:fldSimple[^>]*>/)?.[0] ?? '';
}

describe('<w:fldSimple> w:fldLock / w:dirty round-trip', () => {
  it('preserves w:fldLock="1"', async () => {
    const out = await loadAndResaveDocXml(buildFldSimpleDoc(' w:fldLock="1"'));
    expect(extractFldSimpleOpenTag(out)).toMatch(/w:fldLock="1"/);
  });

  it('preserves w:dirty="1"', async () => {
    const out = await loadAndResaveDocXml(buildFldSimpleDoc(' w:dirty="1"'));
    expect(extractFldSimpleOpenTag(out)).toMatch(/w:dirty="1"/);
  });

  it('preserves both attributes together', async () => {
    const out = await loadAndResaveDocXml(buildFldSimpleDoc(' w:fldLock="1" w:dirty="1"'));
    const tag = extractFldSimpleOpenTag(out);
    expect(tag).toMatch(/w:fldLock="1"/);
    expect(tag).toMatch(/w:dirty="1"/);
  });

  it('normalises w:fldLock="true" to "1"', async () => {
    const out = await loadAndResaveDocXml(buildFldSimpleDoc(' w:fldLock="true"'));
    expect(extractFldSimpleOpenTag(out)).toMatch(/w:fldLock="1"/);
  });

  it('normalises w:dirty="on" to "1"', async () => {
    const out = await loadAndResaveDocXml(buildFldSimpleDoc(' w:dirty="on"'));
    expect(extractFldSimpleOpenTag(out)).toMatch(/w:dirty="1"/);
  });

  it('preserves w:fldLock="0" (explicit false override)', async () => {
    const out = await loadAndResaveDocXml(buildFldSimpleDoc(' w:fldLock="0"'));
    expect(extractFldSimpleOpenTag(out)).toMatch(/w:fldLock="0"/);
  });

  it('omits both attributes when absent on input (regression guard)', async () => {
    const out = await loadAndResaveDocXml(buildFldSimpleDoc(''));
    const tag = extractFldSimpleOpenTag(out);
    expect(tag).not.toMatch(/w:fldLock/);
    expect(tag).not.toMatch(/w:dirty/);
  });
});
