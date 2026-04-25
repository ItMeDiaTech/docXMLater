/**
 * Style-level rPr — text border (w:bdr) round-trip.
 *
 * CT_RPr child `<w:bdr>` per ECMA-376 §17.3.2.5 (run/text border) is
 * handled by the main run parser and the rPrChange parser but the
 * style-level `parseRunFormattingFromXml` never read it. A character
 * style setting a text border was silently dropped on parse.
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
  <w:style w:type="character" w:styleId="BdrTest">
    <w:name w:val="BdrTest"/>
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
  return doc.getStylesManager().getStyle('BdrTest')?.getRunFormatting();
}

describe('Style rPr — text border (w:bdr §17.3.2.5)', () => {
  it('parses <w:bdr> with style/size/color/space attributes', async () => {
    const buffer = await makeDocxWithStyleRPr(
      '<w:bdr w:val="single" w:sz="4" w:space="1" w:color="FF0000"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const border = getRPr(doc)?.border;
    expect(border).toBeDefined();
    expect(border?.style).toBe('single');
    expect(border?.size).toBe(4);
    expect(border?.space).toBe(1);
    expect(border?.color).toBe('FF0000');
    doc.dispose();
  });

  it('parses a minimal <w:bdr> with only style', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:bdr w:val="double"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const border = getRPr(doc)?.border;
    expect(border?.style).toBe('double');
    doc.dispose();
  });

  it('leaves border undefined when <w:bdr> is absent', async () => {
    const buffer = await makeDocxWithStyleRPr('');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.border).toBeUndefined();
    doc.dispose();
  });
});
