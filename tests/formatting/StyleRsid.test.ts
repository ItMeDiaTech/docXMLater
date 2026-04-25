/**
 * CT_Style — `<w:rsid>` (style-level revision-save ID) round-trip.
 *
 * Per ECMA-376 Part 1 §17.7.4 CT_Style, the `<w:rsid>` child element of
 * `<w:style>` carries a CT_LongHexNumber (§17.18.50) identifying the
 * revision session in which this style definition was last modified.
 *
 * Schema position: between `<w:personalReply>` (#15) and `<w:pPr>` (#17)
 * in CT_Style's child-element sequence.
 *
 * Previously the parser never extracted the element and the generator
 * never emitted it, so any Word-authored document carrying rsid on its
 * styles silently lost it across a load/save round-trip.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleRsid(rsidValue: string): Promise<Buffer> {
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
  <w:style w:type="paragraph" w:styleId="RsidTest">
    <w:name w:val="RsidTest"/>
    <w:rsid w:val="${rsidValue}"/>
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

describe('CT_Style rsid (§17.7.4)', () => {
  it('parses <w:rsid w:val="HEX"/> on a style', async () => {
    const buffer = await makeDocxWithStyleRsid('00A12B3C');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('RsidTest');
    expect(style?.getProperties().rsid).toBe('00A12B3C');
    doc.dispose();
  });

  it('leaves rsid undefined when element absent', async () => {
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
  <w:style w:type="paragraph" w:styleId="NoRsid">
    <w:name w:val="NoRsid"/>
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
    const style = doc.getStylesManager().getStyle('NoRsid');
    expect(style?.getProperties().rsid).toBeUndefined();
    doc.dispose();
  });

  it('round-trips rsid through save', async () => {
    const buffer = await makeDocxWithStyleRsid('00DEADBE');
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const stylesFile = zip.getFile('word/styles.xml');
    const content = stylesFile?.content;
    const styles = content instanceof Buffer ? content.toString('utf8') : String(content);
    expect(styles).toMatch(/<w:rsid\s+w:val="00DEADBE"\s*\/>/);
  });

  it('emits rsid after personalReply and before pPr per CT_Style schema order', async () => {
    // Author a style that exercises several children simultaneously so that
    // schema ordering is observable. Note: personalReply + rsid + pPr.
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
    <w:personalReply/>
    <w:rsid w:val="00ABCDEF"/>
    <w:pPr><w:keepNext/></w:pPr>
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

    const prIdx = styles.indexOf('<w:personalReply');
    const rsidIdx = styles.indexOf('<w:rsid');
    const pPrIdx = styles.indexOf('<w:pPr');

    expect(prIdx).toBeGreaterThan(-1);
    expect(rsidIdx).toBeGreaterThan(-1);
    expect(pPrIdx).toBeGreaterThan(-1);
    expect(prIdx).toBeLessThan(rsidIdx);
    expect(rsidIdx).toBeLessThan(pPrIdx);
  });
});
