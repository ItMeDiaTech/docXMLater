/**
 * `<w:rPrChange>` previous-`<w:u>` — underline color / theme attribute
 * round-trip.
 *
 * Per ECMA-376 §17.3.2.40 CT_Underline declares five attributes:
 *   val / color / themeColor / themeTint / themeShade
 *
 * The rPrChange previous-rPr parser previously only read `val`, silently
 * dropping underline-color metadata from tracked-change history. A
 * paragraph whose "previous" underline was a themed accent color (e.g.
 * `themeColor="accent1"`) lost that theme linkage on every load→save
 * round-trip.
 *
 * This iteration extends the rPrChange parser to read all five attrs,
 * matching the main-rPr parser's CT_Underline coverage. The emitter
 * (`Run.generateRunPropertiesXML`) already handles them.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('<w:rPrChange> previous <w:u> underline color/theme round-trip', () => {
  it('preserves color / themeColor / themeTint / themeShade on previous underline', async () => {
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
        <w:rPr>
          <w:u w:val="single" w:color="000000"/>
          <w:rPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
            <w:rPr>
              <w:u w:val="double" w:color="FF0000" w:themeColor="accent1" w:themeTint="66" w:themeShade="80"/>
            </w:rPr>
          </w:rPrChange>
        </w:rPr>
        <w:t>underlined</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
    );
    const buffer = await zipHandler.toBuffer();

    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    const rPrChangeBlock = xml.match(/<w:rPrChange[\s\S]*?<\/w:rPrChange>/)?.[0] ?? '';

    // Previous-underline must carry all five CT_Underline attributes.
    expect(rPrChangeBlock).toMatch(/<w:u[^/>]*w:val="double"/);
    expect(rPrChangeBlock).toMatch(/<w:u[^/>]*w:color="FF0000"/);
    expect(rPrChangeBlock).toMatch(/<w:u[^/>]*w:themeColor="accent1"/);
    expect(rPrChangeBlock).toMatch(/<w:u[^/>]*w:themeTint="66"/);
    expect(rPrChangeBlock).toMatch(/<w:u[^/>]*w:themeShade="80"/);
  });
});
