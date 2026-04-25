/**
 * `<w:pPr><w:pBdr>` — main paragraph border parser must capture all
 * CT_Border attributes per ECMA-376 Part 1 §17.18.2.
 *
 * CT_Border has nine attributes:
 *   w:val, w:sz, w:color, w:space, w:themeColor, w:themeTint,
 *   w:themeShade, w:shadow (CT_OnOff), w:frame (CT_OnOff).
 *
 * The paragraph emitter (Paragraph.ts `createBorder`) emits all nine —
 * but the paragraph parser (`DocumentParser.parseParagraphFormattingFromXml`
 * → `parseBorder`) historically read only four: val / sz / color / space.
 *
 * Impact: a Word-authored paragraph with
 *   <w:pBdr>
 *     <w:top w:val="single" w:sz="4" w:color="auto"
 *            w:themeColor="accent1" w:themeTint="80"
 *            w:shadow="1" w:frame="0"/>
 *   </w:pBdr>
 * loaded with `border.themeColor`/`themeShade`/`shadow`/`frame` as
 * `undefined`, so on save those five attributes were silently stripped.
 * The iter-79 fix covered the `pPrChange` previous-border path but
 * missed the active-pPr path.
 *
 * Iteration 96 extends the main parser so all nine CT_Border attributes
 * survive round-trip.
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

describe('<w:pBdr> CT_Border full attribute round-trip', () => {
  it('preserves w:themeColor / w:themeTint / w:themeShade on paragraph borders', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="4" w:space="1" w:color="auto" w:themeColor="accent1" w:themeTint="80" w:themeShade="40"/>
          <w:bottom w:val="double" w:sz="6" w:space="2" w:color="FF0000"/>
        </w:pBdr>
      </w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const topBorder = out.match(/<w:top[^/]*\/>/)?.[0] ?? '';
    expect(topBorder).toMatch(/w:themeColor="accent1"/);
    expect(topBorder).toMatch(/w:themeTint="80"/);
    expect(topBorder).toMatch(/w:themeShade="40"/);
  });

  it('preserves w:shadow and w:frame on paragraph borders', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="8" w:color="000000" w:shadow="1" w:frame="0"/>
        </w:pBdr>
      </w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const topBorder = out.match(/<w:top[^/]*\/>/)?.[0] ?? '';
    // shadow=1 → emitted as w:shadow="1"; frame=0 → emitted as w:frame="0".
    expect(topBorder).toMatch(/w:shadow="1"/);
    expect(topBorder).toMatch(/w:frame="0"/);
  });

  it('honours "off"/"on" ST_OnOff literals on w:shadow/w:frame', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="4" w:color="000000" w:shadow="on" w:frame="off"/>
        </w:pBdr>
      </w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const topBorder = out.match(/<w:top[^/]*\/>/)?.[0] ?? '';
    // shadow="on" → stored as true → emitted as "1"
    expect(topBorder).toMatch(/w:shadow="1"/);
    // frame="off" → stored as false → emitted as "0"
    expect(topBorder).toMatch(/w:frame="0"/);
  });

  it('keeps size 0 borders intact (regression guard on safeParseInt zero-handling)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="nil" w:sz="0" w:space="0" w:color="auto"/>
        </w:pBdr>
      </w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const topBorder = out.match(/<w:top[^/]*\/>/)?.[0] ?? '';
    expect(topBorder).toMatch(/w:val="nil"/);
    // sz="0" must survive — it is a legal (and common) CT_Border value.
    expect(topBorder).toMatch(/w:sz="0"/);
  });
});
