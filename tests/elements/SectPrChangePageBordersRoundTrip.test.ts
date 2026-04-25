/**
 * `<w:sectPrChange>` previous-`<w:pgBorders>` round-trip.
 *
 * Per ECMA-376 §17.13.5.32 CT_SectPrChange contains a child `<w:sectPr>`
 * whose content model is CT_SectPrBase — including `<w:pgBorders>`
 * (§17.6.10). The main sectPr parser already reads pgBorders, and the
 * sectPrChange EMITTER already supports `prev.pageBorders` — but the
 * sectPrChange PARSER was missing pgBorders handling entirely.
 *
 * Consequence: any tracked-change history of page-border edits (e.g. a
 * user changed the page border color from red to blue, with track-changes
 * enabled) lost the entire "previous" border configuration on every
 * load→save round-trip. Users reviewing the Original-markup view saw no
 * border change recorded, even though the change WAS tracked by Word.
 *
 * This iteration closes the final sectPrChange coverage gap for page
 * borders — the parser now mirrors the main sectPr parser for the same
 * element, preserving all CT_Border attributes including themed colors
 * and shadow/frame flags.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('<w:sectPrChange> previous <w:pgBorders> round-trip', () => {
  it('preserves page border with themed colors in sectPrChange history', async () => {
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
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:pgBorders w:offsetFrom="page">
        <w:top w:val="single" w:sz="12" w:color="0000FF" w:shadow="0" w:frame="0"/>
      </w:pgBorders>
      <w:sectPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:pgBorders w:offsetFrom="page">
            <w:top w:val="double" w:sz="24" w:color="auto" w:themeColor="accent3" w:themeTint="66" w:themeShade="80" w:shadow="1"/>
          </w:pgBorders>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
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

    // Isolate the sectPrChange block for precise assertions.
    const changeBlock = xml.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    expect(changeBlock).toBeTruthy();

    // The previous pgBorders must survive the round-trip with all CT_Border attrs.
    expect(changeBlock).toMatch(/<w:pgBorders[^>]*w:offsetFrom="page"/);
    expect(changeBlock).toMatch(/<w:top[^/>]*w:val="double"/);
    expect(changeBlock).toMatch(/<w:top[^/>]*w:sz="24"/);
    expect(changeBlock).toMatch(/<w:top[^/>]*w:themeColor="accent3"/);
    expect(changeBlock).toMatch(/<w:top[^/>]*w:themeTint="66"/);
    expect(changeBlock).toMatch(/<w:top[^/>]*w:themeShade="80"/);
    expect(changeBlock).toMatch(/<w:top[^/>]*w:shadow="1"/);
  });
});
