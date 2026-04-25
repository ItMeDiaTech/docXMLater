/**
 * `<w:rPrChange>` previous-`<w:rFonts>` theme attribute round-trip.
 *
 * Per ECMA-376 §17.3.2.26 CT_Fonts, `<w:rFonts>` carries nine attributes:
 *   hint / ascii / hAnsi / eastAsia / cs / asciiTheme / hAnsiTheme /
 *   eastAsiaTheme / cstheme
 *
 * The rPrChange parser previously only read the five "literal" attributes
 * (hint + ascii/hAnsi/eastAsia/cs), silently dropping all four theme-font
 * references from tracked-change history. A paragraph whose previous font
 * was a theme reference (e.g. `asciiTheme="minorHAnsi"`) lost the theme
 * linkage on round-trip — the rPrChange history re-emitted as plain fonts
 * instead of theme-font references.
 *
 * This iteration extends the rPrChange parser to read all four theme
 * attributes, matching the main-rPr parser and the style-level rPr parser
 * which already handled them.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('<w:rPrChange> previous <w:rFonts> theme font round-trip', () => {
  it('preserves asciiTheme / hAnsiTheme / eastAsiaTheme / cstheme on previous-rPr', async () => {
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
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
          <w:rPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
            <w:rPr>
              <w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>
            </w:rPr>
          </w:rPrChange>
        </w:rPr>
        <w:t>themed</w:t>
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

    // Extract just the rPrChange block so assertions target the tracked history.
    const rPrChangeBlock = xml.match(/<w:rPrChange[\s\S]*?<\/w:rPrChange>/)?.[0] ?? '';

    expect(rPrChangeBlock).toMatch(/<w:rFonts[^/>]*w:asciiTheme="minorHAnsi"/);
    expect(rPrChangeBlock).toMatch(/<w:rFonts[^/>]*w:hAnsiTheme="minorHAnsi"/);
    expect(rPrChangeBlock).toMatch(/<w:rFonts[^/>]*w:eastAsiaTheme="minorEastAsia"/);
    expect(rPrChangeBlock).toMatch(/<w:rFonts[^/>]*w:cstheme="minorBidi"/);
  });
});
