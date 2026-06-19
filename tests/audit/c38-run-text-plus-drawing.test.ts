/**
 * A single w:r may legally mix a drawing with text/tab/break siblings
 * (ECMA-376 §17.3.3, EG_RunInnerContent is a repeated choice). The run
 * dispatch used to be exclusive: when w:drawing was present only the
 * ImageRun was emitted and any sibling w:t in the same run was dropped, so
 * the text vanished on load/save round-trip. The parser now emits the
 * ImageRun plus a separate Run carrying the sibling text.
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

function documentXmlWithMixedRun(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">RUNTEXT-KEEP </w:t>${DRAWING_XML}</w:r>
    </w:p>
  </w:body>
</w:document>`;
}

async function saveAndReadDocumentXml(doc: Document): Promise<string> {
  const outBuffer = await doc.toBuffer();
  const outZip = new ZipHandler();
  await outZip.loadFromBuffer(outBuffer);
  const outXml = outZip.getFileAsString('word/document.xml');
  expect(outXml).toBeDefined();
  return outXml as string;
}

describe('Run mixing text and a drawing preserves the text (C38)', () => {
  it('keeps the in-model text of a run that also contains a drawing', async () => {
    const buffer = await createDocxWithImage(documentXmlWithMixedRun());
    const doc = await Document.loadFromBuffer(buffer);
    try {
      const text = doc
        .getAllParagraphs()
        .map((p) => p.getText())
        .join('');
      expect(text).toContain('RUNTEXT-KEEP');
      // The image must also be present, not displaced by the text run.
      expect(doc.getImages().length).toBe(1);
    } finally {
      doc.dispose();
    }
  });

  it('round-trips both the text and the drawing in saved document.xml', async () => {
    const buffer = await createDocxWithImage(documentXmlWithMixedRun());
    const doc = await Document.loadFromBuffer(buffer);
    const outXml = await saveAndReadDocumentXml(doc);
    doc.dispose();

    expect(outXml).toContain('RUNTEXT-KEEP');
    expect(outXml).toContain('<w:drawing');
  });
});
