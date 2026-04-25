/**
 * Style-level rPr — fitText and eastAsianLayout.
 *
 * Two more CT_RPr children the style-level rPr parser
 * (`parseRunFormattingFromXml`) silently dropped:
 *
 *   - `w:fitText` (§17.3.2.15)            → fitText (manual run width, twips)
 *   - `w:eastAsianLayout` (§17.3.2.10)    → eastAsianLayout (object)
 *
 * Main run parser handles both; the style-level parser never did, so a
 * character style setting either was silently stripped on load.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleRPr(rPrInner: string): Promise<Buffer> {
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
  <w:style w:type="character" w:styleId="FitEal">
    <w:name w:val="FitEal"/>
    <w:rPr>${rPrInner}</w:rPr>
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

function getRPr(doc: Document) {
  return doc.getStylesManager().getStyle('FitEal')?.getRunFormatting();
}

describe('Style rPr — fitText (w:fitText §17.3.2.15)', () => {
  it('parses <w:fitText w:val="720"/> as fitText: 720 twips', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:fitText w:val="720"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.fitText).toBe(720);
    doc.dispose();
  });

  it('parses <w:fitText w:val="0"/> as fitText: 0 (explicit zero width)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:fitText w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.fitText).toBe(0);
    doc.dispose();
  });

  it('leaves fitText undefined when absent', async () => {
    const buffer = await makeDocxWithStyleRPr('');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.fitText).toBeUndefined();
    doc.dispose();
  });
});

describe('Style rPr — eastAsianLayout (w:eastAsianLayout §17.3.2.10)', () => {
  it('parses a single-attribute layout (vert=true)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:eastAsianLayout w:vert="1"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const layout = getRPr(doc)?.eastAsianLayout;
    expect(layout).toBeDefined();
    expect(layout?.vert).toBe(true);
    doc.dispose();
  });

  it('parses a multi-attribute layout with id, combine, and combineBrackets', async () => {
    const buffer = await makeDocxWithStyleRPr(
      '<w:eastAsianLayout w:id="5" w:combine="1" w:combineBrackets="round"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const layout = getRPr(doc)?.eastAsianLayout;
    expect(layout).toBeDefined();
    expect(layout?.id).toBe(5);
    expect(layout?.combine).toBe(true);
    expect(layout?.combineBrackets).toBe('round');
    doc.dispose();
  });

  it('leaves eastAsianLayout undefined when absent', async () => {
    const buffer = await makeDocxWithStyleRPr('');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.eastAsianLayout).toBeUndefined();
    doc.dispose();
  });
});
