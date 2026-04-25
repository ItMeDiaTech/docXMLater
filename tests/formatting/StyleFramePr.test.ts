/**
 * Style-level paragraph — `<w:framePr>` (text frame properties) round-trip.
 *
 * Per ECMA-376 Part 1 §17.3.1.11 CT_FramePr, `<w:framePr>` is a child of
 * CT_PPrBase carrying drop-cap / floating-frame positioning for styles
 * like "drop cap" or custom magazine-layout frames. It has ~16 attributes
 * (w, h, hRule, x, y, xAlign, yAlign, hAnchor, vAnchor, hSpace, vSpace,
 * wrap, dropCap, lines, anchorLock).
 *
 * The Paragraph-level serializer emits it correctly, but the style-level
 * `generateParagraphProperties` in `formatting/Style.ts` silently dropped
 * it (commented "not yet serialized on styles"), and the style-level
 * `parseParagraphFormattingFromXml` never extracted it. A paragraph style
 * carrying drop-cap configuration lost it entirely on any save that went
 * through the managed style emission path.
 *
 * Schema position within CT_PPrBase: #5, between `pageBreakBefore` (#4)
 * and `widowControl` (#6).
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleFramePr(framePrXml: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="FrameTest">
    <w:name w:val="FrameTest"/>
    <w:pPr>${framePrXml}</w:pPr>
  </w:style>
</w:styles>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

function getFrame(doc: Document) {
  return doc.getStylesManager().getStyle('FrameTest')?.getProperties().paragraphFormatting?.framePr;
}

describe('Style-level paragraph framePr (§17.3.1.11 CT_FramePr)', () => {
  it('parses dropCap=drop / lines=3 / wrap=around framePr configuration', async () => {
    const buffer = await makeDocxWithStyleFramePr(
      '<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:hSpace="144"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const fp = getFrame(doc);
    expect(fp?.dropCap).toBe('drop');
    expect(fp?.lines).toBe(3);
    expect(fp?.wrap).toBe('around');
    expect(fp?.vAnchor).toBe('text');
    expect(fp?.hAnchor).toBe('margin');
    expect(fp?.hSpace).toBe(144);
    doc.dispose();
  });

  it('parses floating-frame w/h/x/y/anchor attributes', async () => {
    const buffer = await makeDocxWithStyleFramePr(
      '<w:framePr w:w="2880" w:h="1440" w:hRule="atLeast" w:x="720" w:y="720" w:xAlign="center" w:yAlign="top" w:vSpace="180" w:anchorLock="1"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const fp = getFrame(doc);
    expect(fp?.w).toBe(2880);
    expect(fp?.h).toBe(1440);
    expect(fp?.hRule).toBe('atLeast');
    expect(fp?.x).toBe(720);
    expect(fp?.y).toBe(720);
    expect(fp?.xAlign).toBe('center');
    expect(fp?.yAlign).toBe('top');
    expect(fp?.vSpace).toBe(180);
    expect(fp?.anchorLock).toBe(true);
    doc.dispose();
  });

  it('round-trips a dropCap framePr through save', async () => {
    const buffer = await makeDocxWithStyleFramePr(
      '<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const stylesFile = zip.getFile('word/styles.xml');
    const content = stylesFile?.content;
    const styles = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(styles).toMatch(/<w:framePr\b/);
    expect(styles).toMatch(/w:dropCap="drop"/);
    expect(styles).toMatch(/w:lines="3"/);
    expect(styles).toMatch(/w:wrap="around"/);
    expect(styles).toMatch(/w:vAnchor="text"/);
  });

  it('emits framePr between pageBreakBefore and widowControl per CT_PPrBase schema', async () => {
    // Author a style exercising several CT_PPrBase children simultaneously.
    const zipHandler = new ZipHandler();
    zipHandler.addFile(
      '[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    );
    zipHandler.addFile(
      'word/styles.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="OrderTest">
    <w:name w:val="OrderTest"/>
    <w:pPr>
      <w:pageBreakBefore/>
      <w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>
      <w:widowControl/>
    </w:pPr>
  </w:style>
</w:styles>`
    );
    zipHandler.addFile(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
    );
    const buffer = await zipHandler.toBuffer();

    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const stylesFile = zip.getFile('word/styles.xml');
    const content = stylesFile?.content;
    const styles = content instanceof Buffer ? content.toString('utf8') : String(content);

    const pbbIdx = styles.indexOf('<w:pageBreakBefore');
    const frameIdx = styles.indexOf('<w:framePr');
    const widowIdx = styles.indexOf('<w:widowControl');

    expect(pbbIdx).toBeGreaterThan(-1);
    expect(frameIdx).toBeGreaterThan(-1);
    expect(widowIdx).toBeGreaterThan(-1);
    expect(pbbIdx).toBeLessThan(frameIdx);
    expect(frameIdx).toBeLessThan(widowIdx);
  });
});
