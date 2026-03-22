/**
 * Tests for tracked image deletion/insertion in table cells being preserved during round-trip.
 *
 * Issue: When Word creates a tracked deletion of an image in a table cell,
 * ImageRun.toXML() strips w:rPr (containing w:noProof) and regenerates w:drawing
 * from parsed properties instead of preserving verbatim. This causes cell corruption
 * when the tracked change is accepted in Word.
 *
 * Fix: Store the original raw <w:r> XML on ImageRun during revision parsing.
 * When toXML() is called and raw XML is available, return __rawXml passthrough.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function createDocxWithImage(documentXml: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
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
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`
  );

  // Minimal 1x1 PNG
  const pngBuffer = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAB' +
      'Nl7BcQAAAABJRU5ErkJggg==',
    'base64'
  );
  zipHandler.addFile('word/media/image1.png', pngBuffer);

  zipHandler.addFile('word/document.xml', documentXml);

  return await zipHandler.toBuffer();
}

const DRAWING_XML = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

describe('ImageRun Revision Round-Trip', () => {
  describe('w:del with image run', () => {
    it('should preserve w:rPr with w:noProof in deleted image run during round-trip', async () => {
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
            <w:del w:id="1" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
              <w:r><w:rPr><w:noProof/></w:rPr>${DRAWING_XML}</w:r>
            </w:del>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

      const buffer = await createDocxWithImage(documentXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const outBuffer = await doc.toBuffer();
      doc.dispose();

      // Read back the saved document XML
      const outZip = new ZipHandler();
      await outZip.loadFromBuffer(outBuffer);
      const outXml = outZip.getFileAsString('word/document.xml');

      // The output must contain w:rPr with w:noProof inside the deleted image run
      expect(outXml).toContain('<w:rPr>');
      expect(outXml).toContain('<w:noProof');
      // The deletion revision must be preserved
      expect(outXml).toContain('w:del');
      expect(outXml).toContain('w:author="TestUser"');
    });
  });

  describe('w:ins with image run', () => {
    it('should preserve raw XML for inserted image run during round-trip', async () => {
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
            <w:ins w:id="2" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
              <w:r><w:rPr><w:noProof/></w:rPr>${DRAWING_XML}</w:r>
            </w:ins>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

      const buffer = await createDocxWithImage(documentXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const outBuffer = await doc.toBuffer();
      doc.dispose();

      const outZip = new ZipHandler();
      await outZip.loadFromBuffer(outBuffer);
      const outXml = outZip.getFileAsString('word/document.xml');

      expect(outXml).toContain('<w:rPr>');
      expect(outXml).toContain('<w:noProof');
      expect(outXml).toContain('w:ins');
    });
  });

  describe('image + text in same cell', () => {
    it('should preserve deleted image without affecting text runs in same cell', async () => {
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
            <w:r><w:t>Some text</w:t></w:r>
            <w:del w:id="3" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
              <w:r><w:rPr><w:noProof/></w:rPr>${DRAWING_XML}</w:r>
            </w:del>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

      const buffer = await createDocxWithImage(documentXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const outBuffer = await doc.toBuffer();
      doc.dispose();

      const outZip = new ZipHandler();
      await outZip.loadFromBuffer(outBuffer);
      const outXml = outZip.getFileAsString('word/document.xml');

      // Text run preserved
      expect(outXml).toContain('Some text');
      // Image deletion preserved with run properties
      expect(outXml).toContain('<w:noProof');
      expect(outXml).toContain('w:del');
    });
  });
});
