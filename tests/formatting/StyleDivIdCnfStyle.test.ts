/**
 * Style-level pPr — divId and cnfStyle round-trip.
 *
 * CT_PPrBase positions #32 (divId §17.3.1.10 — HTML div association) and
 * #33 (cnfStyle §17.3.1.8 — conditional formatting bitmask). Both are
 * valid on a paragraph style's pPr per ECMA-376.
 *
 * The main paragraph parser handles both. The style-level parser
 * (`parseParagraphFormattingFromXml`) and the style serializer
 * (`Style.generateParagraphProperties`) silently dropped them — the
 * generator never emitted either, and the parser never read them.
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStylePPr(pPrInner: string): Promise<Buffer> {
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
  <w:style w:type="paragraph" w:styleId="DivCnf">
    <w:name w:val="DivCnf"/>
    <w:pPr>${pPrInner}</w:pPr>
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

function getFmt(doc: Document) {
  return doc.getStylesManager().getStyle('DivCnf')?.getParagraphFormatting();
}

describe('Style pPr — divId (§17.3.1.10)', () => {
  it('parses <w:divId w:val="42"/> as divId: 42', async () => {
    const buffer = await makeDocxWithStylePPr(`<w:divId w:val="42"/>`);
    const doc = await Document.loadFromBuffer(buffer);
    expect(getFmt(doc)?.divId).toBe(42);
    doc.dispose();
  });

  it('emits <w:divId w:val="42"/> via toXML()', () => {
    const style = new Style({
      styleId: 'DivCnf',
      type: 'paragraph',
      name: 'DivCnf',
      paragraphFormatting: { divId: 42 },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toContain('<w:divId w:val="42"/>');
  });

  it('omits <w:divId/> when undefined', () => {
    const style = new Style({
      styleId: 'DivCnf',
      type: 'paragraph',
      name: 'DivCnf',
      paragraphFormatting: {},
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).not.toContain('<w:divId');
  });
});

describe('Style pPr — cnfStyle (§17.3.1.8)', () => {
  it('parses <w:cnfStyle w:val="100000000000"/> as cnfStyle bitmask', async () => {
    const buffer = await makeDocxWithStylePPr(`<w:cnfStyle w:val="100000000000"/>`);
    const doc = await Document.loadFromBuffer(buffer);
    expect(getFmt(doc)?.cnfStyle).toBe('100000000000');
    doc.dispose();
  });

  it('emits <w:cnfStyle w:val="010000000010"/> via toXML()', () => {
    const style = new Style({
      styleId: 'DivCnf',
      type: 'paragraph',
      name: 'DivCnf',
      paragraphFormatting: { cnfStyle: '010000000010' },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toContain('<w:cnfStyle w:val="010000000010"/>');
  });

  it('omits <w:cnfStyle/> when undefined', () => {
    const style = new Style({
      styleId: 'DivCnf',
      type: 'paragraph',
      name: 'DivCnf',
      paragraphFormatting: {},
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).not.toContain('<w:cnfStyle');
  });
});

describe('Style pPr — divId and cnfStyle child order (§17.3.1.26)', () => {
  it('emits outlineLvl → divId → cnfStyle in CT_PPrBase order', () => {
    const style = new Style({
      styleId: 'DivCnf',
      type: 'paragraph',
      name: 'DivCnf',
      paragraphFormatting: {
        outlineLevel: 2,
        divId: 42,
        cnfStyle: '100000000000',
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    const olIdx = xml.indexOf('<w:outlineLvl');
    const divIdx = xml.indexOf('<w:divId');
    const cnfIdx = xml.indexOf('<w:cnfStyle');
    expect(olIdx).toBeGreaterThan(-1);
    expect(divIdx).toBeGreaterThan(-1);
    expect(cnfIdx).toBeGreaterThan(-1);
    expect(divIdx).toBeGreaterThan(olIdx);
    expect(cnfIdx).toBeGreaterThan(divIdx);
  });
});
