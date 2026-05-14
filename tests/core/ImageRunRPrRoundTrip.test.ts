/**
 * Tests that w:rPr on a run wrapping an inline drawing survives round-trip
 * outside of revision (w:del/w:ins) contexts.
 *
 * Bug: DocumentParser created ImageRun from <w:r><w:rPr/><w:drawing/></w:r>
 * without applying the parent run's properties, and ImageRun.toXML() emitted
 * <w:r><w:drawing/></w:r> with no w:rPr. The dropped w:rFonts shifted line
 * metrics in Word — images with effectExtent (drop shadow) overflowed their
 * containing cell and clipped into the previous row.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function buildDocx(documentXml: string): Promise<Buffer> {
  const zip = new ZipHandler();
  zip.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`
  );
  const pngBuffer = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAB' +
      'Nl7BcQAAAABJRU5ErkJggg==',
    'base64'
  );
  zip.addFile('word/media/image1.png', pngBuffer);
  zip.addFile('word/document.xml', documentXml);
  return zip.toBuffer();
}

const DRAWING_XML = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:effectExtent l="38100" t="38100" r="38100" b="38100"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

describe('ImageRun w:rPr round-trip (no revision wrapper)', () => {
  it('preserves w:rFonts and w:noProof on the run wrapping an inline drawing', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/>
          <w:noProof/>
          <w:color w:val="FF0000"/>
        </w:rPr>
        ${DRAWING_XML}
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await buildDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);
    const outBuffer = await doc.toBuffer();
    doc.dispose();

    const outZip = new ZipHandler();
    await outZip.loadFromBuffer(outBuffer);
    const outXml = outZip.getFileAsString('word/document.xml') ?? '';

    const drawingRunMatch = outXml.match(/<w:r[^>]*>(<w:rPr>[\s\S]*?<\/w:rPr>)?<w:drawing/);
    expect(drawingRunMatch).not.toBeNull();
    const rPr = drawingRunMatch?.[1] ?? '';

    expect(rPr).toContain('<w:rFonts');
    expect(rPr).toMatch(/w:ascii="Verdana"/);
    expect(rPr).toContain('<w:noProof');
    expect(rPr).toMatch(/<w:color[^/]*w:val="FF0000"/);
  });

  it('preserves w:rPr on drawing-runs inside table cells', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:r>
              <w:rPr>
                <w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/>
                <w:b/>
                <w:noProof/>
              </w:rPr>
              ${DRAWING_XML}
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

    const buffer = await buildDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);
    const outBuffer = await doc.toBuffer();
    doc.dispose();

    const outZip = new ZipHandler();
    await outZip.loadFromBuffer(outBuffer);
    const outXml = outZip.getFileAsString('word/document.xml') ?? '';

    expect(outXml).toMatch(/<w:r[^>]*><w:rPr>[\s\S]*?<w:rFonts[^/]*w:ascii="Verdana"/);
    expect(outXml).toMatch(/<w:rPr>[\s\S]*?<w:noProof[\s\S]*?<\/w:rPr><w:drawing/);
  });

  it('round-trips a paragraph with no rPr on the drawing-run without inventing one', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:p>
      <w:r>${DRAWING_XML}</w:r>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await buildDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);
    const outBuffer = await doc.toBuffer();
    doc.dispose();

    const outZip = new ZipHandler();
    await outZip.loadFromBuffer(outBuffer);
    const outXml = outZip.getFileAsString('word/document.xml') ?? '';

    // When the source had no rPr on the drawing run, we must not synthesize one.
    expect(outXml).toMatch(/<w:r><w:drawing/);
  });
});
