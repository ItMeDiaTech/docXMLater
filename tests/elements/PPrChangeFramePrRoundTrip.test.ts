/**
 * `<w:pPrChange><w:pPr><w:framePr …/>` — tracked-change history of
 * paragraph frame properties (text frames, drop caps, floating
 * positioning) must round-trip.
 *
 * Per ECMA-376 Part 1 §17.3.1.11 CT_FramePr, `<w:framePr>` on a
 * paragraph's pPr carries 17+ attributes covering text-frame size
 * (w/h/hRule), positioning (x/y/xAlign/yAlign/hAnchor/vAnchor),
 * spacing (hSpace/vSpace), wrap mode, drop-cap config (dropCap/lines),
 * and anchor locking (anchorLock, ST_OnOff).
 *
 * The pPrChange emitter at `Paragraph.ts:3634` already rebuilds every
 * framePr attribute onto the previous-properties block. The parser at
 * `DocumentParser.ts:2774` — which extracts the same previous-properties
 * for the round-trip — never read `w:framePr` at all. Round-trip
 * impact: a tracked change to any frame property (flipping drop-cap
 * on/off, repositioning a text-box paragraph, changing wrap mode) lost
 * the entire "previous" frame state on load → save, breaking Word's
 * "Original" markup view for that revision.
 *
 * Iteration 114 mirrors the main-pPr framePr parser onto the
 * pPrChange path: every numeric attr via isExplicitlySet + safeParseInt
 * (zero/negative coordinates are valid), enum/anchor strings via
 * String(), and w:anchorLock via parseOnOffAttribute so every ST_OnOff
 * literal ("1"/"0"/"true"/"false"/"on"/"off") resolves.
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

function extractPrevFramePr(xml: string): string {
  const changeBlock = xml.match(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/)?.[0] ?? '';
  return changeBlock.match(/<w:framePr[^/]*\/>/)?.[0] ?? '';
}

describe('<w:pPrChange> previous <w:framePr> round-trip', () => {
  it('preserves sizing and wrap attributes on previous framePr', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:framePr w:w="4000" w:h="2000" w:hRule="exact" w:wrap="around" w:hAnchor="margin" w:vAnchor="text"/>
        <w:pPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:framePr w:w="3000" w:h="1500" w:hRule="atLeast" w:wrap="notBeside" w:hAnchor="page" w:vAnchor="page"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const prev = extractPrevFramePr(out);
    expect(prev).toMatch(/w:w="3000"/);
    expect(prev).toMatch(/w:h="1500"/);
    expect(prev).toMatch(/w:hRule="atLeast"/);
    expect(prev).toMatch(/w:wrap="notBeside"/);
    expect(prev).toMatch(/w:hAnchor="page"/);
    expect(prev).toMatch(/w:vAnchor="page"/);
  });

  it('preserves drop-cap and lines on previous framePr', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:framePr w:dropCap="drop" w:lines="3" w:hAnchor="margin" w:vAnchor="text"/>
        <w:pPrChange w:id="2" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:framePr w:dropCap="margin" w:lines="4" w:hAnchor="margin" w:vAnchor="text"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const prev = extractPrevFramePr(out);
    expect(prev).toMatch(/w:dropCap="margin"/);
    expect(prev).toMatch(/w:lines="4"/);
  });

  it('preserves zero-valued coordinates (x="0" / y="0")', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:framePr w:x="100" w:y="100" w:hAnchor="page" w:vAnchor="page"/>
        <w:pPrChange w:id="3" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:framePr w:x="0" w:y="0" w:hAnchor="page" w:vAnchor="page"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const prev = extractPrevFramePr(out);
    expect(prev).toMatch(/w:x="0"/);
    expect(prev).toMatch(/w:y="0"/);
  });

  it('honours ST_OnOff "on"/"off" literals on previous framePr anchorLock', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:framePr w:anchorLock="0" w:hAnchor="margin" w:vAnchor="text"/>
        <w:pPrChange w:id="4" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:framePr w:anchorLock="on" w:hAnchor="margin" w:vAnchor="text"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const prev = extractPrevFramePr(out);
    // anchorLock="on" → stored as true → emitted as "1"
    expect(prev).toMatch(/w:anchorLock="1"/);
  });
});
