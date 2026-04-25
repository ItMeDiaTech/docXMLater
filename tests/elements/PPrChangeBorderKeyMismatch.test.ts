/**
 * `<w:pPrChange><w:pPr><w:pBdr>` — previous-border parser stored the
 * attributes under the wrong field names (`val`, `sz`, `space` …)
 * while the emitter reads `style`, `size`, `space` …. The mismatch
 * silently flattened every tracked-change previous border to the
 * `<w:top w:val="nil"/>` default, losing style and width entirely.
 *
 * In addition, the parser only ever read five of CT_Border's nine
 * attributes (§17.18.2) — no themeTint / themeShade / shadow / frame.
 *
 * Iteration 97 rewrites the pPrChange `parseBorder` helper to:
 *   - use the same `{ style, size, space, color, themeColor,
 *     themeTint, themeShade, shadow, frame }` keys as the main
 *     paragraph parser and the emitter;
 *   - route `shadow`/`frame` through `parseOnOffAttribute` so every
 *     ST_OnOff literal ("true"/"false"/"1"/"0"/"on"/"off") resolves.
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

function extractPPrChangeTopBorder(xml: string): string {
  const changeBlock = xml.match(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/)?.[0] ?? '';
  return changeBlock.match(/<w:top[^/]*\/>/)?.[0] ?? '';
}

describe('<w:pPrChange> previous <w:pBdr> key-name round-trip', () => {
  it('preserves w:val (style) and w:sz (size) on previous top border', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="double" w:sz="12" w:space="4" w:color="00FF00"/>
        </w:pBdr>
        <w:pPrChange w:id="1" w:author="T" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:pBdr>
              <w:top w:val="single" w:sz="4" w:space="1" w:color="FF0000"/>
            </w:pBdr>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const topBorder = extractPPrChangeTopBorder(out);
    // Previously this collapsed to `<w:top w:val="nil" w:color="FF0000"/>`
    // because the parser stored the hex-style under the wrong key.
    expect(topBorder).toMatch(/w:val="single"/);
    expect(topBorder).toMatch(/w:sz="4"/);
    expect(topBorder).toMatch(/w:space="1"/);
    expect(topBorder).toMatch(/w:color="FF0000"/);
  });

  it('preserves themed + shadow/frame on previous top border', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="4" w:color="auto"/>
        </w:pBdr>
        <w:pPrChange w:id="2" w:author="T" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:pBdr>
              <w:top w:val="wave" w:sz="6" w:color="auto" w:themeColor="accent2" w:themeTint="80" w:themeShade="40" w:shadow="1" w:frame="0"/>
            </w:pBdr>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const topBorder = extractPPrChangeTopBorder(out);
    expect(topBorder).toMatch(/w:val="wave"/);
    expect(topBorder).toMatch(/w:themeColor="accent2"/);
    expect(topBorder).toMatch(/w:themeTint="80"/);
    expect(topBorder).toMatch(/w:themeShade="40"/);
    expect(topBorder).toMatch(/w:shadow="1"/);
    expect(topBorder).toMatch(/w:frame="0"/);
  });

  it('honours ST_OnOff "on"/"off" literals on previous shadow/frame', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="4" w:color="auto"/>
        </w:pBdr>
        <w:pPrChange w:id="3" w:author="T" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:pBdr>
              <w:top w:val="single" w:sz="4" w:color="auto" w:shadow="on" w:frame="off"/>
            </w:pBdr>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const topBorder = extractPPrChangeTopBorder(out);
    expect(topBorder).toMatch(/w:shadow="1"/);
    expect(topBorder).toMatch(/w:frame="0"/);
  });
});
