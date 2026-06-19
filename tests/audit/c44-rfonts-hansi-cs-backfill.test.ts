/**
 * `<w:rFonts>` w:hAnsi / w:cs must not be synthesized from the ascii font.
 *
 * Per ECMA-376 §17.3.2.26, each rFonts slot (ascii, hAnsi, eastAsia, cs)
 * inherits independently from the style/theme hierarchy when absent. The
 * serializer previously back-filled w:hAnsi and w:cs from formatting.font,
 * so a parsed run carrying only `<w:rFonts w:ascii="Times New Roman"/>`
 * re-serialized as `<w:rFonts w:ascii=".." w:hAnsi=".." w:cs=".."/>` —
 * overriding style/theme-resolved hAnsi/cs fonts (e.g. a minorBidi
 * complex-script font) and polluting rPrChange tracked-change history.
 *
 * w:hAnsi / w:cs are now emitted only when explicitly set; setFont()
 * mirrors the chosen font into both slots for programmatic use.
 */

import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function buildDocx(runXml: string): Promise<Buffer> {
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
    <w:p>${runXml}</w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

async function roundTrip(runXml: string): Promise<string> {
  const buffer = await buildDocx(runXml);
  const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
  try {
    const saved = await doc.toBuffer();
    const zip = new ZipHandler();
    await zip.loadFromBuffer(saved);
    return zip.getFileAsString('word/document.xml') ?? '';
  } finally {
    doc.dispose();
  }
}

describe('rFonts w:hAnsi / w:cs are not back-filled from the ascii font', () => {
  it('round-trips <w:rFonts w:ascii="Times New Roman"/> without inventing w:hAnsi or w:cs', async () => {
    const xml = await roundTrip(
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/></w:rPr><w:t>x</w:t></w:r>'
    );
    const rFonts = xml.match(/<w:rFonts[^>]*>/)?.[0] ?? '';

    expect(rFonts).toContain('w:ascii="Times New Roman"');
    expect(rFonts).not.toContain('w:hAnsi=');
    expect(rFonts).not.toContain('w:cs=');
  });

  it('round-trips explicit w:hAnsi and w:cs values unchanged', async () => {
    const xml = await roundTrip(
      '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Helvetica" w:cs="Arial Unicode MS"/></w:rPr><w:t>x</w:t></w:r>'
    );
    const rFonts = xml.match(/<w:rFonts[^>]*>/)?.[0] ?? '';

    expect(rFonts).toContain('w:ascii="Times New Roman"');
    expect(rFonts).toContain('w:hAnsi="Helvetica"');
    expect(rFonts).toContain('w:cs="Arial Unicode MS"');
  });

  it('does not fabricate w:cs in rPrChange previous-rPr history', async () => {
    const xml = await roundTrip(
      `<w:r>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
          <w:rPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
            <w:rPr>
              <w:rFonts w:ascii="Calibri"/>
            </w:rPr>
          </w:rPrChange>
        </w:rPr>
        <w:t>x</w:t>
      </w:r>`
    );
    const rPrChangeBlock = xml.match(/<w:rPrChange[\s\S]*?<\/w:rPrChange>/)?.[0] ?? '';
    const prevRFonts = rPrChangeBlock.match(/<w:rFonts[^>]*>/)?.[0] ?? '';

    expect(prevRFonts).toContain('w:ascii="Calibri"');
    expect(prevRFonts).not.toContain('w:hAnsi=');
    expect(prevRFonts).not.toContain('w:cs=');
  });

  it('serializes only w:ascii when the formatting object sets font alone', () => {
    const run = new Run('x', { font: 'Courier New' });
    const xml = XMLBuilder.elementToString(run.toXML());
    const rFonts = xml.match(/<w:rFonts[^>]*>/)?.[0] ?? '';

    expect(rFonts).toContain('w:ascii="Courier New"');
    expect(rFonts).not.toContain('w:hAnsi=');
    expect(rFonts).not.toContain('w:cs=');
  });

  it('setFont() mirrors the font into w:hAnsi and w:cs for programmatic use', () => {
    const run = new Run('x');
    run.setFont('Arial');
    const xml = XMLBuilder.elementToString(run.toXML());
    const rFonts = xml.match(/<w:rFonts[^>]*>/)?.[0] ?? '';

    expect(rFonts).toContain('w:ascii="Arial"');
    expect(rFonts).toContain('w:hAnsi="Arial"');
    expect(rFonts).toContain('w:cs="Arial"');
  });
});
