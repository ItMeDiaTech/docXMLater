/**
 * Image mutators must mark the image as mutated so changes survive saving.
 *
 * For images parsed inside tracked revisions (w:ins/w:del), ImageRun captures
 * the original raw run XML and emits it verbatim on save unless the image
 * reports isMutated(). Previously only setSize() and setBorder() set the flag,
 * so setWidth(), setAltText(), rotate(), removeBorder(), etc. on a
 * revision-nested image were silently dropped: the saved document was
 * byte-identical to the input despite the mutation.
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

function drawingXml(spPrExtra = ''): string {
  return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPrExtra}</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

function documentXmlWithTrackedImage(spPrExtra = ''): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:p>
      <w:ins w:id="1" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
        <w:r><w:rPr><w:noProof/></w:rPr>${drawingXml(spPrExtra)}</w:r>
      </w:ins>
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

describe('Image mutators flag mutation for revision-nested images (raw run XML refresh)', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('setWidth() on a tracked-insertion image is reflected in saved wp:extent', async () => {
    const buffer = await createDocxWithImage(documentXmlWithTrackedImage());
    doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    const images = doc.getImages();
    expect(images.length).toBe(1);
    images[0]!.image.setWidth(457200);

    const outXml = await saveAndReadDocumentXml(doc);
    expect(outXml).toContain('cx="457200"');
    expect(outXml).not.toContain('cx="914400"');
    // Captured run properties from the original run must survive the refresh
    expect(outXml).toContain('<w:noProof');
  });

  it('setAltText() on a tracked-insertion image is reflected in saved wp:docPr', async () => {
    const buffer = await createDocxWithImage(documentXmlWithTrackedImage());
    doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    doc.getImages()[0]!.image.setAltText('Company logo');

    const outXml = await saveAndReadDocumentXml(doc);
    expect(outXml).toContain('descr="Company logo"');
  });

  it('rotate() on a tracked-insertion image emits a:xfrm rot in saved XML', async () => {
    const buffer = await createDocxWithImage(documentXmlWithTrackedImage());
    doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    doc.getImages()[0]!.image.rotate(90);

    const outXml = await saveAndReadDocumentXml(doc);
    // 90 degrees = 5400000 60000ths of a degree (ECMA-376 ST_Angle)
    expect(outXml).toContain('rot="5400000"');
  });

  it('removeBorder() on a tracked-insertion image strips a:ln from saved XML', async () => {
    const borderLn = '<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>';
    const buffer = await createDocxWithImage(documentXmlWithTrackedImage(borderLn));
    doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    const image = doc.getImages()[0]!.image;
    expect(image.getBorder()).toBeDefined();
    image.removeBorder();

    const outXml = await saveAndReadDocumentXml(doc);
    expect(outXml).not.toContain('<a:ln w="25400"');
  });
});
