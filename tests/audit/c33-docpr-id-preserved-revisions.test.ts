/**
 * assignUniqueIds must not assign a wp:docPr id that collides with a drawing
 * whose original XML is preserved verbatim on save.
 *
 * Images inside preserved revisions (w:ins/w:del with revisionHandling:
 * 'preserve') keep their captured raw run XML — including the original
 * wp:docPr id — while top-level drawings are renumbered sequentially from 1.
 * Previously a preserved revision image with id="1" plus any normal image
 * produced two <wp:docPr id="1"> elements in document.xml, violating the
 * uniqueness of wp:docPr/@id (ECMA-376 Part 1 §20.4.2.5).
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { ImageRun } from '../../src/elements/ImageRun';

async function createDocx(documentXml: string): Promise<Buffer> {
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

function drawingXml(docPrId: number): string {
  return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="${docPrId}" name="Picture ${docPrId}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

// Tracked-insertion image with docPr id=1 followed by a normal image with docPr id=2
const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:p>
      <w:ins w:id="1" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
        <w:r><w:rPr><w:noProof/></w:rPr>${drawingXml(1)}</w:r>
      </w:ins>
    </w:p>
    <w:p>
      <w:r>${drawingXml(2)}</w:r>
    </w:p>
  </w:body>
</w:document>`;

function extractDocPrIds(xml: string): string[] {
  const ids: string[] = [];
  const docPrPattern = /<wp:docPr\b[^>]*\bid="(\d+)"/g;
  let match: RegExpExecArray | null;
  while ((match = docPrPattern.exec(xml)) !== null) {
    ids.push(match[1]!);
  }
  return ids;
}

async function saveAndReadDocumentXml(doc: Document): Promise<string> {
  const outBuffer = await doc.toBuffer();
  const outZip = new ZipHandler();
  await outZip.loadFromBuffer(outBuffer);
  const outXml = outZip.getFileAsString('word/document.xml');
  expect(outXml).toBeDefined();
  return outXml as string;
}

describe('wp:docPr id uniqueness with preserved revision images', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('does not renumber a normal image into the id of a preserved tracked-insertion image', async () => {
    const buffer = await createDocx(DOCUMENT_XML);
    doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    const outXml = await saveAndReadDocumentXml(doc);
    const ids = extractDocPrIds(outXml);

    expect(ids.length).toBe(2);
    expect(new Set(ids).size).toBe(ids.length);
    // The preserved revision run keeps its original id verbatim
    expect(outXml).toContain('<wp:docPr id="1" name="Picture 1"/>');
  });

  it('gives a mutated tracked-insertion image a fresh id distinct from renumbered drawings', async () => {
    const buffer = await createDocx(DOCUMENT_XML);
    doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    // Mutate the revision-nested image so its drawing is regenerated from the
    // model (spliced into the captured run XML) instead of emitted verbatim
    const revisions = doc.getParagraphs()[0]!.getRevisions();
    expect(revisions.length).toBe(1);
    const trackedImageRun = revisions[0]!
      .getContent()
      .find((item): item is ImageRun => item instanceof ImageRun);
    expect(trackedImageRun).toBeDefined();
    trackedImageRun!.getImageElement().setWidth(457200, false);

    const outXml = await saveAndReadDocumentXml(doc);
    const ids = extractDocPrIds(outXml);

    expect(ids.length).toBe(2);
    expect(new Set(ids).size).toBe(ids.length);
    // The mutation itself must survive
    expect(outXml).toContain('cx="457200"');
  });
});
