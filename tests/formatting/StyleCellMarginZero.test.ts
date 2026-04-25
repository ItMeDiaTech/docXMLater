/**
 * Style-level tcPr — cell margin zero-value parsing.
 *
 * `parseCellMarginsFromXml` parses `<w:tblCellMar>` / `<w:tcMar>` into
 * `{ top, left, bottom, right }` millimetre-twip values. The current
 * implementation uses `if (w)` to guard each side, where `w` comes from
 * `XMLParser.extractAttribute(...)` which coerces numeric strings to
 * numbers when `parseAttributeValue: true`. For explicit zero (`w:w="0"`
 * — a valid "no margin on this side" configuration), the numeric `0`
 * is falsy and the side is silently dropped.
 *
 * This is the same class of bug the pPr CT_OnOff parser fixed in
 * iteration 25 (`parseOoxmlBoolean` / `!== undefined` semantics) —
 * generalised to a numeric-zero context.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithCellMargins(tblCellMarInner: string): Promise<Buffer> {
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
  <w:style w:type="table" w:styleId="MarginTest">
    <w:name w:val="MarginTest"/>
    <w:tblPr>
      <w:tblCellMar>${tblCellMarInner}</w:tblCellMar>
    </w:tblPr>
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

function getMargins(doc: Document) {
  return doc.getStylesManager().getStyle('MarginTest')?.getProperties().tableStyle?.table
    ?.cellMargins;
}

describe('Style tblCellMar — explicit zero margin round-trip (§17.4.43)', () => {
  it('parses top="0" as margins.top = 0 (not dropped)', async () => {
    const buffer = await makeDocxWithCellMargins(`<w:top w:w="0" w:type="dxa"/>`);
    const doc = await Document.loadFromBuffer(buffer);
    expect(getMargins(doc)?.top).toBe(0);
    doc.dispose();
  });

  it('parses bottom="0" as margins.bottom = 0', async () => {
    const buffer = await makeDocxWithCellMargins(`<w:bottom w:w="0" w:type="dxa"/>`);
    const doc = await Document.loadFromBuffer(buffer);
    expect(getMargins(doc)?.bottom).toBe(0);
    doc.dispose();
  });

  it('parses left="0" as margins.left = 0', async () => {
    const buffer = await makeDocxWithCellMargins(`<w:left w:w="0" w:type="dxa"/>`);
    const doc = await Document.loadFromBuffer(buffer);
    expect(getMargins(doc)?.left).toBe(0);
    doc.dispose();
  });

  it('parses right="0" as margins.right = 0', async () => {
    const buffer = await makeDocxWithCellMargins(`<w:right w:w="0" w:type="dxa"/>`);
    const doc = await Document.loadFromBuffer(buffer);
    expect(getMargins(doc)?.right).toBe(0);
    doc.dispose();
  });

  it('parses all four zero margins together', async () => {
    const buffer = await makeDocxWithCellMargins(
      `<w:top w:w="0" w:type="dxa"/>
       <w:left w:w="0" w:type="dxa"/>
       <w:bottom w:w="0" w:type="dxa"/>
       <w:right w:w="0" w:type="dxa"/>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const m = getMargins(doc);
    expect(m?.top).toBe(0);
    expect(m?.left).toBe(0);
    expect(m?.bottom).toBe(0);
    expect(m?.right).toBe(0);
    doc.dispose();
  });

  it('still parses non-zero margin values correctly (no regression)', async () => {
    const buffer = await makeDocxWithCellMargins(`<w:top w:w="108" w:type="dxa"/>`);
    const doc = await Document.loadFromBuffer(buffer);
    expect(getMargins(doc)?.top).toBe(108);
    doc.dispose();
  });
});
