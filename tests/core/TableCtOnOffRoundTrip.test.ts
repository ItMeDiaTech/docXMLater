/**
 * Table / Row / Cell CT_OnOff Round-Trip Tests
 *
 * Seven CT_OnOff properties cover table/row/cell direct formatting:
 *   Table  : w:bidiVisual
 *   Row    : w:tblHeader, w:cantSplit, w:hidden
 *   Cell   : w:noWrap, w:hideMark, w:tcFitText
 *
 * Per ECMA-376 Part 1 §17.17.4 (ST_OnOff), the presence of a CT_OnOff
 * element does NOT imply true — it means "read w:val". When w:val is
 * absent, the value is true (presence = on); when w:val is "0", "false",
 * or "off", the value is explicitly false.
 *
 * Bugs this suite guards against:
 *   - Main-path parsers hardcoded `= true` regardless of w:val, so a
 *     source document containing `<w:cantSplit w:val="0"/>` was silently
 *     flipped to cantSplit=true, misrepresenting row pagination rules.
 *   - `parseGenericPreviousProperties` (shared by tblPrChange, trPrChange,
 *     tcPrChange) had the same bug, flipping the tracked "previous" value
 *     and so flipping the recorded change history.
 *   - Generators emitted only a bare `<w:cantSplit/>` for true, with no
 *     way to represent explicit false, so a parsed explicit-false could
 *     not round-trip.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithTable(tblPrXml: string, trPrXml = '', tcPrXml = ''): Promise<Buffer> {
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
    <w:tbl>
      <w:tblPr>${tblPrXml}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr>${trPrXml}</w:trPr>
        <w:tc>
          <w:tcPr>${tcPrXml}</w:tcPr>
          <w:p><w:r><w:t>test</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`
  );

  return await zipHandler.toBuffer();
}

describe('Table CT_OnOff — main-path parser honours w:val', () => {
  it('parses <w:bidiVisual w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithTable('<w:bidiVisual w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getBidiVisual()).toBe(false);
    doc.dispose();
  });

  it('parses <w:bidiVisual/> as true', async () => {
    const buffer = await makeDocxWithTable('<w:bidiVisual/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getBidiVisual()).toBe(true);
    doc.dispose();
  });

  it('parses <w:bidiVisual w:val="false"/> as false', async () => {
    const buffer = await makeDocxWithTable('<w:bidiVisual w:val="false"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getBidiVisual()).toBe(false);
    doc.dispose();
  });

  it('parses <w:bidiVisual w:val="off"/> as false', async () => {
    const buffer = await makeDocxWithTable('<w:bidiVisual w:val="off"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getBidiVisual()).toBe(false);
    doc.dispose();
  });
});

describe('Row CT_OnOff — main-path parser honours w:val', () => {
  it('parses <w:cantSplit w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithTable('', '<w:cantSplit w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const row = doc.getTables()[0]!.getRows()[0]!;
    expect(row.getCantSplit()).toBe(false);
    doc.dispose();
  });

  it('parses <w:cantSplit/> as true', async () => {
    const buffer = await makeDocxWithTable('', '<w:cantSplit/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getRows()[0]!.getCantSplit()).toBe(true);
    doc.dispose();
  });

  it('parses <w:tblHeader w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithTable('', '<w:tblHeader w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getRows()[0]!.getIsHeader()).toBe(false);
    doc.dispose();
  });

  it('parses <w:tblHeader w:val="false"/> as false', async () => {
    const buffer = await makeDocxWithTable('', '<w:tblHeader w:val="false"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getRows()[0]!.getIsHeader()).toBe(false);
    doc.dispose();
  });

  it('parses <w:hidden w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithTable('', '<w:hidden w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getRows()[0]!.isHidden()).toBe(false);
    doc.dispose();
  });

  it('parses <w:hidden w:val="off"/> as false', async () => {
    const buffer = await makeDocxWithTable('', '<w:hidden w:val="off"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getRows()[0]!.isHidden()).toBe(false);
    doc.dispose();
  });
});

describe('Cell CT_OnOff — main-path parser honours w:val', () => {
  it('parses <w:noWrap w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithTable('', '', '<w:noWrap w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const cell = doc.getTables()[0]!.getRows()[0]!.getCells()[0]!;
    expect(cell.getNoWrap()).toBe(false);
    doc.dispose();
  });

  it('parses <w:hideMark w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithTable('', '', '<w:hideMark w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const cell = doc.getTables()[0]!.getRows()[0]!.getCells()[0]!;
    expect(cell.getHideMark()).toBe(false);
    doc.dispose();
  });

  it('parses <w:tcFitText w:val="0"/> as false', async () => {
    const buffer = await makeDocxWithTable('', '', '<w:tcFitText w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const cell = doc.getTables()[0]!.getRows()[0]!.getCells()[0]!;
    expect(cell.getFitText()).toBe(false);
    doc.dispose();
  });

  it('parses <w:noWrap w:val="true"/> as true', async () => {
    const buffer = await makeDocxWithTable('', '', '<w:noWrap w:val="true"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getTables()[0]!.getRows()[0]!.getCells()[0]!.getNoWrap()).toBe(true);
    doc.dispose();
  });
});

describe('tblPrChange / trPrChange / tcPrChange — generic parser honours w:val', () => {
  it('parses <w:bidiVisual w:val="0"/> inside tblPrChange as false', async () => {
    const buffer = await makeDocxWithTable(
      `<w:tblPrChange w:id="1" w:author="T" w:date="2024-01-01T00:00:00Z">
        <w:tblPr><w:bidiVisual w:val="0"/></w:tblPr>
      </w:tblPrChange>`
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getTables()[0]!.getTblPrChange();
    expect(change?.previousProperties.bidiVisual).toBe(false);
    doc.dispose();
  });

  it('parses <w:cantSplit w:val="0"/> inside trPrChange as false', async () => {
    const buffer = await makeDocxWithTable(
      '',
      `<w:trPrChange w:id="1" w:author="T" w:date="2024-01-01T00:00:00Z">
        <w:trPr><w:cantSplit w:val="0"/></w:trPr>
      </w:trPrChange>`
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getTables()[0]!.getRows()[0]!.getTrPrChange();
    expect(change?.previousProperties.cantSplit).toBe(false);
    doc.dispose();
  });

  it('parses <w:tblHeader w:val="false"/> inside trPrChange as false', async () => {
    const buffer = await makeDocxWithTable(
      '',
      `<w:trPrChange w:id="1" w:author="T" w:date="2024-01-01T00:00:00Z">
        <w:trPr><w:tblHeader w:val="false"/></w:trPr>
      </w:trPrChange>`
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getTables()[0]!.getRows()[0]!.getTrPrChange();
    expect(change?.previousProperties.isHeader).toBe(false);
    doc.dispose();
  });

  it('parses <w:hidden w:val="0"/> inside trPrChange as false', async () => {
    const buffer = await makeDocxWithTable(
      '',
      `<w:trPrChange w:id="1" w:author="T" w:date="2024-01-01T00:00:00Z">
        <w:trPr><w:hidden w:val="0"/></w:trPr>
      </w:trPrChange>`
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getTables()[0]!.getRows()[0]!.getTrPrChange();
    expect(change?.previousProperties.hidden).toBe(false);
    doc.dispose();
  });

  it('parses <w:noWrap w:val="0"/> inside tcPrChange as false', async () => {
    const buffer = await makeDocxWithTable(
      '',
      '',
      `<w:tcPrChange w:id="1" w:author="T" w:date="2024-01-01T00:00:00Z">
        <w:tcPr><w:noWrap w:val="0"/></w:tcPr>
      </w:tcPrChange>`
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getTables()[0]!.getRows()[0]!.getCells()[0]!.getTcPrChange();
    expect(change?.previousProperties.noWrap).toBe(false);
    doc.dispose();
  });

  it('parses <w:hideMark w:val="off"/> inside tcPrChange as false', async () => {
    const buffer = await makeDocxWithTable(
      '',
      '',
      `<w:tcPrChange w:id="1" w:author="T" w:date="2024-01-01T00:00:00Z">
        <w:tcPr><w:hideMark w:val="off"/></w:tcPr>
      </w:tcPrChange>`
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getTables()[0]!.getRows()[0]!.getCells()[0]!.getTcPrChange();
    expect(change?.previousProperties.hideMark).toBe(false);
    doc.dispose();
  });

  it('parses <w:tcFitText w:val="0"/> inside tcPrChange as false', async () => {
    const buffer = await makeDocxWithTable(
      '',
      '',
      `<w:tcPrChange w:id="1" w:author="T" w:date="2024-01-01T00:00:00Z">
        <w:tcPr><w:tcFitText w:val="0"/></w:tcPr>
      </w:tcPrChange>`
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = doc.getTables()[0]!.getRows()[0]!.getCells()[0]!.getTcPrChange();
    expect(change?.previousProperties.fitText).toBe(false);
    doc.dispose();
  });
});

describe('Table CT_OnOff — explicit false round-trips through generator', () => {
  it('round-trips bidiVisual=false through load → save → load', async () => {
    const buffer1 = await makeDocxWithTable('<w:bidiVisual w:val="0"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(doc1.getTables()[0]!.getBidiVisual()).toBe(false);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(doc2.getTables()[0]!.getBidiVisual()).toBe(false);
    doc2.dispose();
  });

  it('round-trips cantSplit=false through load → save → load', async () => {
    const buffer1 = await makeDocxWithTable('', '<w:cantSplit w:val="0"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(doc1.getTables()[0]!.getRows()[0]!.getCantSplit()).toBe(false);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(doc2.getTables()[0]!.getRows()[0]!.getCantSplit()).toBe(false);
    doc2.dispose();
  });

  it('round-trips noWrap=false through load → save → load', async () => {
    const buffer1 = await makeDocxWithTable('', '', '<w:noWrap w:val="0"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(doc1.getTables()[0]!.getRows()[0]!.getCells()[0]!.getNoWrap()).toBe(false);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(doc2.getTables()[0]!.getRows()[0]!.getCells()[0]!.getNoWrap()).toBe(false);
    doc2.dispose();
  });
});
