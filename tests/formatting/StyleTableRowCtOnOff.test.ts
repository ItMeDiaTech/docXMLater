/**
 * Style-level table row CT_OnOff — cantSplit / tblHeader.
 *
 * `DocumentParser.parseTableRowFormattingFromXml` (the style-level trPr
 * parser) detects `<w:cantSplit>` and `<w:tblHeader>` via substring scan:
 *
 *   if (trPrXml.includes('<w:cantSplit/>') || trPrXml.includes('<w:cantSplit ')) {
 *     formatting.cantSplit = true;
 *   }
 *
 * Same bug that iteration 25 fixed on pPr and iteration 40 on rPr:
 *   - `<w:cantSplit w:val="off"/>` matches `<w:cantSplit ` and is
 *     hard-coded to `true`, silently flipping a tblStylePr's explicit-off
 *     override (e.g., inside a header-row conditional that un-splits the
 *     parent style's cantSplit) into an enabled flag.
 *   - Parser also never writes `false` for these fields, so explicit-off
 *     has nowhere to land.
 *
 * Both `w:cantSplit` (§17.4.6) and `w:tblHeader` (§17.4.49/50) are
 * OnOffOnlyType per the memory `project_onoffonly_elements.md`, so
 * explicit-false appears as `w:val="off"` (not `"0"` or `"false"`).
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithTableStyleTrPr(trPrInner: string): Promise<Buffer> {
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
  <w:style w:type="table" w:styleId="TblStyleTrPr">
    <w:name w:val="TblStyleTrPr"/>
    <w:trPr>${trPrInner}</w:trPr>
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

function getRowFmt(doc: Document) {
  return doc.getStylesManager().getStyle('TblStyleTrPr')?.getProperties().tableStyle?.row;
}

describe('Style-level trPr — cantSplit (§17.4.6, OnOffOnlyType)', () => {
  it('<w:cantSplit/> (bare) parses as true', async () => {
    const buffer = await makeDocxWithTableStyleTrPr('<w:cantSplit/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRowFmt(doc)?.cantSplit).toBe(true);
    doc.dispose();
  });

  it('<w:cantSplit w:val="on"/> parses as true', async () => {
    const buffer = await makeDocxWithTableStyleTrPr('<w:cantSplit w:val="on"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRowFmt(doc)?.cantSplit).toBe(true);
    doc.dispose();
  });

  it('<w:cantSplit w:val="off"/> parses as false (not true)', async () => {
    const buffer = await makeDocxWithTableStyleTrPr('<w:cantSplit w:val="off"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRowFmt(doc)?.cantSplit).toBe(false);
    doc.dispose();
  });

  it('absent <w:cantSplit> leaves cantSplit undefined', async () => {
    const buffer = await makeDocxWithTableStyleTrPr('');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRowFmt(doc)?.cantSplit).toBeUndefined();
    doc.dispose();
  });
});

describe('Style-level trPr — tblHeader (§17.4.49/50, OnOffOnlyType)', () => {
  it('<w:tblHeader/> (bare) parses as isHeader: true', async () => {
    const buffer = await makeDocxWithTableStyleTrPr('<w:tblHeader/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRowFmt(doc)?.isHeader).toBe(true);
    doc.dispose();
  });

  it('<w:tblHeader w:val="on"/> parses as isHeader: true', async () => {
    const buffer = await makeDocxWithTableStyleTrPr('<w:tblHeader w:val="on"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRowFmt(doc)?.isHeader).toBe(true);
    doc.dispose();
  });

  it('<w:tblHeader w:val="off"/> parses as isHeader: false (not true)', async () => {
    const buffer = await makeDocxWithTableStyleTrPr('<w:tblHeader w:val="off"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRowFmt(doc)?.isHeader).toBe(false);
    doc.dispose();
  });

  it('absent <w:tblHeader> leaves isHeader undefined', async () => {
    const buffer = await makeDocxWithTableStyleTrPr('');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRowFmt(doc)?.isHeader).toBeUndefined();
    doc.dispose();
  });
});
