/**
 * CleanupHelper.unlockFields / unlockFrames must clear lock state on the
 * in-memory model so the unlock survives save.
 *
 * prepareSave() regenerates word/document.xml from the model, so editing
 * the raw ZIP copy of document.xml is overwritten: the serializers re-emit
 * w:fldLock (Run fldChar content, Field fldSimple) and w:anchorLock
 * (paragraph framePr) from the parsed model state. The cleanup therefore
 * has to mutate the model itself for the unlock to persist in the saved
 * document.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { CleanupHelper } from '../../src/helpers/CleanupHelper';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';

async function buildDocx(documentXml: string): Promise<Buffer> {
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
  zipHandler.addFile('word/document.xml', documentXml);
  return zipHandler.toBuffer();
}

async function extractDocumentXml(buffer: Buffer): Promise<string> {
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  const content = zip.getFile('word/document.xml')?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

describe('CleanupHelper unlockFields/unlockFrames persist through save', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('unlockFields removes w:fldLock from a loaded fldSimple and the saved XML', async () => {
    const buffer = await buildDocx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:fldSimple w:instr="PAGE" w:fldLock="1">
        <w:r><w:t>42</w:t></w:r>
      </w:fldSimple>
    </w:p>
  </w:body>
</w:document>`);
    doc = await Document.loadFromBuffer(buffer);

    const cleanup = new CleanupHelper(doc);
    const report = cleanup.run({ unlockFields: true });
    expect(report.fieldsUnlocked).toBe(1);

    const out = await extractDocumentXml(await doc.toBuffer());
    expect(out).toContain('w:fldSimple');
    expect(out).not.toMatch(/w:fldLock/);
  });

  it('unlockFields leaves explicit w:fldLock="0" untouched (already unlocked)', async () => {
    const buffer = await buildDocx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:fldSimple w:instr="PAGE" w:fldLock="0">
        <w:r><w:t>42</w:t></w:r>
      </w:fldSimple>
    </w:p>
  </w:body>
</w:document>`);
    doc = await Document.loadFromBuffer(buffer);

    const cleanup = new CleanupHelper(doc);
    const report = cleanup.run({ unlockFields: true });
    expect(report.fieldsUnlocked).toBe(0);

    const out = await extractDocumentXml(await doc.toBuffer());
    expect(out).toMatch(/w:fldLock="0"/);
  });

  it('unlockFields clears fieldCharLocked on run-level w:fldChar content', async () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addRun(
      Run.createFromContent([{ type: 'fieldChar', fieldCharType: 'begin', fieldCharLocked: true }])
    );
    para.addRun(Run.createFromContent([{ type: 'instructionText', value: ' PAGE ' }]));
    para.addRun(Run.createFromContent([{ type: 'fieldChar', fieldCharType: 'end' }]));
    doc.addParagraph(para);

    const cleanup = new CleanupHelper(doc);
    const report = cleanup.run({ unlockFields: true });
    expect(report.fieldsUnlocked).toBe(1);

    const out = await extractDocumentXml(await doc.toBuffer());
    expect(out).toContain('w:fldChar');
    expect(out).not.toMatch(/w:fldLock/);
  });

  it('unlockFrames removes w:anchorLock from a loaded framePr and the saved XML', async () => {
    const buffer = await buildDocx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:framePr w:w="2000" w:h="1000" w:hAnchor="margin" w:vAnchor="text" w:anchorLock="1"/>
      </w:pPr>
      <w:r><w:t>framed</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`);
    doc = await Document.loadFromBuffer(buffer);

    const cleanup = new CleanupHelper(doc);
    const report = cleanup.run({ unlockFrames: true });
    expect(report.framesUnlocked).toBe(1);

    const out = await extractDocumentXml(await doc.toBuffer());
    const framePr = out.match(/<w:framePr[^>]*\/>/)?.[0] ?? '';
    expect(framePr).not.toBe('');
    expect(framePr).not.toMatch(/w:anchorLock/);
    // The rest of the frame definition must survive the unlock
    expect(framePr).toMatch(/w:w="2000"/);
    expect(framePr).toMatch(/w:hAnchor="margin"/);
  });

  it('unlockFrames counts paragraphs inside table cells', async () => {
    const buffer = await buildDocx(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:p>
            <w:pPr>
              <w:framePr w:hAnchor="margin" w:vAnchor="text" w:anchorLock="true"/>
            </w:pPr>
            <w:r><w:t>cell frame</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
  </w:body>
</w:document>`);
    doc = await Document.loadFromBuffer(buffer);

    const cleanup = new CleanupHelper(doc);
    const report = cleanup.run({ unlockFrames: true });
    expect(report.framesUnlocked).toBe(1);

    const out = await extractDocumentXml(await doc.toBuffer());
    expect(out).not.toMatch(/w:anchorLock/);
  });
});
