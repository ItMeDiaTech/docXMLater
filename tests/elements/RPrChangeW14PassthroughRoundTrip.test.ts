/**
 * `<w:rPrChange><w:rPr><w14:textOutline …/>` — tracked-change history
 * of Word 2010+ w14: text effects must round-trip.
 *
 * The `w14:` namespace (MC-DOCX §2.6.x) declares a suite of text
 * effects beyond the ECMA-376 CT_RPr schema: w14:textOutline,
 * w14:shadow, w14:reflection, w14:glow, w14:ligatures, w14:numForm,
 * w14:numSpacing, w14:cntxtAlts, w14:stylisticSets. The main rPr
 * parser collects these as raw-XML passthrough strings on
 * `RunFormatting.rawW14Properties`, and `generateRunPropertiesXML`
 * re-emits them in their schema-mandated position (after all
 * ECMA-376 CT_RPr children).
 *
 * The rPrChange parser, however, skipped w14: children entirely,
 * so when `<w:rPrChange>` recorded a previous rPr containing any
 * w14 text effect the previous state silently vanished on every
 * round-trip — breaking Word's "Original" view of that revision
 * and Ribbon preview of accepting / rejecting the change.
 *
 * Iteration 116 mirrors the main-rPr w14 collection loop onto the
 * previous-rPr parse, populating `prevProps.rawW14Properties` so the
 * emitter (which already handles the field) can write the effects
 * back.
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

function extractRPrChangeBlock(xml: string): string {
  return xml.match(/<w:rPrChange[\s\S]*?<\/w:rPrChange>/)?.[0] ?? '';
}

describe('<w:rPrChange> previous w14: text-effect passthrough', () => {
  it('preserves w14:shadow on the previous rPr', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:rPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
            <w:rPr>
              <w14:shadow w14:blurRad="50800" w14:dist="38100" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">
                <w14:schemeClr w14:val="tx1"/>
              </w14:shadow>
            </w:rPr>
          </w:rPrChange>
        </w:rPr>
        <w:t>text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = extractRPrChangeBlock(out);
    // Previously this block emitted only <w:rPr><w:b …/></w:rPr> with
    // no trace of w14:shadow — the tracked "previous" text-shadow
    // was silently lost.
    expect(changeBlock).toMatch(/w14:shadow/);
    expect(changeBlock).toMatch(/w14:blurRad="50800"/);
    expect(changeBlock).toMatch(/w14:algn="tl"/);
  });

  it('preserves w14:textOutline on the previous rPr', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:rPrChange w:id="2" w:author="Tester" w:date="2026-01-01T00:00:00Z">
            <w:rPr>
              <w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr">
                <w14:solidFill>
                  <w14:schemeClr w14:val="accent1"/>
                </w14:solidFill>
              </w14:textOutline>
            </w:rPr>
          </w:rPrChange>
        </w:rPr>
        <w:t>outlined</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = extractRPrChangeBlock(out);
    expect(changeBlock).toMatch(/w14:textOutline/);
    expect(changeBlock).toMatch(/w14:w="9525"/);
    expect(changeBlock).toMatch(/w14:cap="flat"/);
  });

  it('does not emit w14 on previous rPr when absent (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:rPrChange w:id="3" w:author="Tester" w:date="2026-01-01T00:00:00Z">
            <w:rPr>
              <w:i/>
            </w:rPr>
          </w:rPrChange>
        </w:rPr>
        <w:t>plain</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = extractRPrChangeBlock(out);
    expect(changeBlock).not.toMatch(/w14:/);
    // Plain italic should still round-trip.
    expect(changeBlock).toMatch(/<w:i/);
  });
});
