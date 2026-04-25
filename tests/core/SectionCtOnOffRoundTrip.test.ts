/**
 * Section CT_OnOff Round-Trip Tests
 *
 * Section pPr contains five CT_OnOff flags that drive real rendering:
 *   w:formProt, w:noEndnote, w:titlePg, w:bidi, w:rtlGutter
 *
 * The main section parser used XMLParser.hasSelfClosingTag(...) to detect
 * these, which returns true for ANY presence of `<w:x/>` or `<w:x ...` — it
 * does NOT read w:val. A source document containing `<w:bidi w:val="0"/>`
 * (explicit override of an inherited true) was therefore silently flipped
 * to bidi=true. The sectPrChange (tracked section property change) parser
 * had the same bug plus a naive `sectPrXml.includes('<w:bidi')` substring
 * check, which would also fire on any element whose name starts with "bidi"
 * (the only OOXML example is `w:bidiVisual`, a table property — a false
 * positive that would be theoretical in a well-formed w:sectPr but still
 * fragile).
 *
 * Per ECMA-376 Part 1 §17.17.4 the five elements are full CT_OnOff and
 * w:val honours every ST_OnOff literal ("1"/"0"/"true"/"false"/"on"/"off").
 * This suite locks that behaviour on both the main parser and the
 * w:sectPrChange tracked-change parser, and verifies the generator
 * preserves an explicit false across a round-trip.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithSectPr(sectPrInner: string): Promise<Buffer> {
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

  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
  );

  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>test</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      ${sectPrInner}
    </w:sectPr>
  </w:body>
</w:document>`
  );

  return await zipHandler.toBuffer();
}

describe('Main sectPr CT_OnOff — parser honours w:val', () => {
  it('parses <w:titlePg w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithSectPr('<w:titlePg w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getTitlePage()).toBe(false);
    doc.dispose();
  });

  it('parses <w:titlePg/> as true', async () => {
    const buffer = await makeDocxWithSectPr('<w:titlePg/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getTitlePage()).toBe(true);
    doc.dispose();
  });

  it('parses <w:titlePg w:val="false"/> as false', async () => {
    const buffer = await makeDocxWithSectPr('<w:titlePg w:val="false"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getTitlePage()).toBe(false);
    doc.dispose();
  });

  it('parses <w:bidi w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithSectPr('<w:bidi w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getBidi()).toBe(false);
    doc.dispose();
  });

  it('parses <w:bidi/> as true', async () => {
    const buffer = await makeDocxWithSectPr('<w:bidi/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getBidi()).toBe(true);
    doc.dispose();
  });

  it('parses <w:rtlGutter w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithSectPr('<w:rtlGutter w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getRtlGutter()).toBe(false);
    doc.dispose();
  });

  it('parses <w:noEndnote w:val="off"/> as false', async () => {
    const buffer = await makeDocxWithSectPr('<w:noEndnote w:val="off"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getProperties().noEndnote).toBe(false);
    doc.dispose();
  });

  it('parses <w:formProt w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithSectPr('<w:formProt w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getProperties().formProt).toBe(false);
    doc.dispose();
  });

  it('parses <w:formProt/> as true', async () => {
    const buffer = await makeDocxWithSectPr('<w:formProt/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getSection().getProperties().formProt).toBe(true);
    doc.dispose();
  });
});

describe('sectPrChange (tracked section property change) — parser honours w:val', () => {
  const wrapInChange = (prevInner: string) =>
    `<w:sectPrChange w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z">
      <w:sectPr>${prevInner}</w:sectPr>
    </w:sectPrChange>`;

  it('parses <w:bidi w:val="0"/> inside sectPrChange as false', async () => {
    const buffer = await makeDocxWithSectPr(wrapInChange('<w:bidi w:val="0"/>'));
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getSection().getSectPrChange();
    expect(change?.previousProperties.bidi).toBe(false);
    doc.dispose();
  });

  it('parses <w:bidi/> inside sectPrChange as true', async () => {
    const buffer = await makeDocxWithSectPr(wrapInChange('<w:bidi/>'));
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getSection().getSectPrChange();
    expect(change?.previousProperties.bidi).toBe(true);
    doc.dispose();
  });

  it('parses <w:bidi w:val="false"/> inside sectPrChange as false', async () => {
    const buffer = await makeDocxWithSectPr(wrapInChange('<w:bidi w:val="false"/>'));
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getSection().getSectPrChange();
    expect(change?.previousProperties.bidi).toBe(false);
    doc.dispose();
  });

  it('parses <w:titlePg w:val="0"/> inside sectPrChange as false', async () => {
    const buffer = await makeDocxWithSectPr(wrapInChange('<w:titlePg w:val="0"/>'));
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getSection().getSectPrChange();
    expect(change?.previousProperties.titlePage).toBe(false);
    doc.dispose();
  });

  it('parses <w:formProt w:val="off"/> inside sectPrChange as false', async () => {
    const buffer = await makeDocxWithSectPr(wrapInChange('<w:formProt w:val="off"/>'));
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getSection().getSectPrChange();
    expect(change?.previousProperties.formProt).toBe(false);
    doc.dispose();
  });

  it('parses <w:noEndnote w:val="0"/> inside sectPrChange as false', async () => {
    const buffer = await makeDocxWithSectPr(wrapInChange('<w:noEndnote w:val="0"/>'));
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getSection().getSectPrChange();
    expect(change?.previousProperties.noEndnote).toBe(false);
    doc.dispose();
  });

  it('parses <w:rtlGutter w:val="0"/> inside sectPrChange as false', async () => {
    const buffer = await makeDocxWithSectPr(wrapInChange('<w:rtlGutter w:val="0"/>'));
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getSection().getSectPrChange();
    expect(change?.previousProperties.rtlGutter).toBe(false);
    doc.dispose();
  });

  it('parses absent w:bidi inside sectPrChange as undefined', async () => {
    const buffer = await makeDocxWithSectPr(wrapInChange('<w:pgSz w:w="11000" w:h="14000"/>'));
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getSection().getSectPrChange();
    expect(change?.previousProperties.bidi).toBeUndefined();
    doc.dispose();
  });
});
