/**
 * `<w:p w14:paraId="…" w14:textId="…">` — Word 2010+ paragraph IDs
 * must survive load → save round-trip.
 *
 * Per MC-DOCX §2.6.19 (Word 2010+ extension), every `<w:p>` can carry
 * two `ST_LongHexNumber` attributes — `w14:paraId` and `w14:textId` —
 * that Word uses as stable IDs for change-tracking and comments.
 * Values are 8 hexadecimal digits (e.g., `"3F0E1A22"`).
 *
 * Bug: both paragraph parse paths (`parseParagraphFromElement` at
 * `DocumentParser.ts:921` and `parseParagraphFromObject` at
 * `DocumentParser.ts:2030`) read the IDs via the WRONG key shape:
 *
 *     const paraId = pElement['w14:paraId'];   // always undefined
 *
 * XMLParser stores attributes under the `@_` prefix, so the correct
 * key is `pElement['@_w14:paraId']`. The prior lookup silently
 * returned `undefined`, so every `w14:paraId` / `w14:textId` on every
 * loaded Word document was discarded. On subsequent save the
 * generator auto-assigned fresh IDs, invalidating any external
 * references (comments.xml threads, tracked-change linkage, Word's
 * own change log).
 *
 * Also: XMLParser's `parseAttributeValue: true` coerces purely-numeric
 * hex strings like `"00000001"` to the number `1`, so we normalise the
 * parsed value back to 8-char uppercase hex.
 *
 * Iteration 108 fixes both parse sites.
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

describe('<w:p w14:paraId / w14:textId> round-trip', () => {
  it('preserves w14:paraId and w14:textId with mixed hex/digit values', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="3F0E1A22" w14:textId="1B2C3D4E">
      <w:r><w:t>hello</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    expect(out).toMatch(/w14:paraId="3F0E1A22"/);
    expect(out).toMatch(/w14:textId="1B2C3D4E"/);
  });

  it('preserves purely-digit hex IDs (normalises XMLParser numeric coercion)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="00000001" w14:textId="12345678">
      <w:r><w:t>hello</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    // Previously XMLParser coerced "00000001" to number 1; after the
    // fix the value is stored as the canonical 8-char uppercase form.
    expect(out).toMatch(/w14:paraId="00000001"/);
    expect(out).toMatch(/w14:textId="12345678"/);
  });

  it('still works for paragraphs without w14 IDs (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>plain</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    // Paragraphs without w14 IDs should load cleanly. The generator
    // auto-assigns IDs on save (existing behaviour), so out DOES contain
    // fresh w14:paraId/w14:textId attributes — just not the ones from
    // our input (there were none). Sanity-check that we didn't crash
    // and that the paragraph survived.
    expect(out).toContain('plain');
  });
});
