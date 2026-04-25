/**
 * `<w:cols>` — num attribute is OPTIONAL per ECMA-376 Part 1 §17.6.4
 * CT_Columns. Defaults: num=1, equalWidth=true, sep=false, space=720.
 *
 * The section parser previously had `if (num) { ... }` which skipped
 * every cols element lacking an explicit `w:num`. A section that toggled
 * only the separator (Word's common single-column-with-separator form:
 * `<w:cols w:sep="1" w:space="720"/>`) silently lost both attributes on
 * round-trip because `sectionProps.columns` was never assigned.
 *
 * Iteration 99 defaults num per spec (1, or the count of `<w:col>`
 * children if present) so attributes on a num-less cols element survive.
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

describe('<w:cols> default num parsing', () => {
  it('preserves w:sep and w:space when w:num is absent (spec-default num=1)', async () => {
    // Word's "single column with separator line" form omits w:num.
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>x</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:cols w:sep="1" w:space="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    // Previously the <w:cols> was dropped entirely because the `if (num)`
    // gate skipped it. The parser now defaults num to 1 per §17.6.4, so
    // the sep/space round-trip.
    expect(out).toMatch(/<w:cols[^/]*w:sep="1"/);
    expect(out).toMatch(/<w:cols[^/]*w:space="720"/);
  });

  it('defaults num to child <w:col> count when num attribute is absent', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>x</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:cols w:equalWidth="0" w:space="720">
        <w:col w:w="3000" w:space="720"/>
        <w:col w:w="5000"/>
      </w:cols>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    expect(out).toMatch(/<w:cols[^>]*w:num="2"/);
    expect(out).toMatch(/<w:col[^/]*w:w="3000"/);
    expect(out).toMatch(/<w:col[^/]*w:w="5000"/);
  });

  it('still parses w:num when present (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>x</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:cols w:num="3" w:space="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    expect(out).toMatch(/<w:cols[^>]*w:num="3"/);
  });
});
