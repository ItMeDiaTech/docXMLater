/**
 * Table row tracked insertion / deletion — ECMA-376 Part 1 §17.13.5.19 / §17.13.5.14
 *
 * Per CT_TrPr (§17.4.79), a table row's `<w:trPr>` MAY contain `<w:ins>` and
 * `<w:del>` child elements that mark the entire row as a tracked insertion
 * or deletion. These children appear AFTER the CT_TrPrBase content
 * (cnfStyle / divId / gridBefore / … / hidden) and BEFORE `<w:trPrChange>`:
 *
 *   <xsd:complexType name="CT_TrPr">
 *     <xsd:complexContent>
 *       <xsd:extension base="CT_TrPrBase">
 *         <xsd:sequence>
 *           <xsd:element name="ins" type="CT_TrackChange" minOccurs="0"/>
 *           <xsd:element name="del" type="CT_TrackChange" minOccurs="0"/>
 *           <xsd:element name="trPrChange" type="CT_TrPrChange" minOccurs="0"/>
 *         </xsd:sequence>
 *       </xsd:extension>
 *     </xsd:complexContent>
 *   </xsd:complexType>
 *
 * Bug guarded against: the parser and generator previously ignored
 * `<w:ins>` and `<w:del>` on rows entirely. Loading a DOCX with tracked
 * row insertions or deletions and saving it back silently dropped the
 * revision markers, losing the tracked-change history on rows.
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithRowIns(insOrDel: 'ins' | 'del'): Promise<Buffer> {
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
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr>
          <w:${insOrDel} w:id="7" w:author="Jane Doe" w:date="2026-01-15T10:00:00Z"/>
        </w:trPr>
        <w:tc><w:tcPr/><w:p><w:r><w:t>tracked row</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('TableRow tracked insertion (w:ins in w:trPr, §17.13.5.19)', () => {
  it('parses <w:ins> in <w:trPr> as rowInsertion metadata', async () => {
    const buffer = await makeDocxWithRowIns('ins');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const table = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    expect(table).toBeDefined();
    const row = table.getRows()[0]!;
    expect(row.getFormatting().rowInsertion).toBeDefined();
    expect(row.getFormatting().rowInsertion?.author).toBe('Jane Doe');
    expect(row.getFormatting().rowInsertion?.id).toBe('7');
    expect(row.getFormatting().rowInsertion?.date).toBe('2026-01-15T10:00:00Z');
    doc.dispose();
  });

  it('emits <w:ins> inside <w:trPr> for rowInsertion', () => {
    const row = new TableRow(1, {
      rowInsertion: {
        id: '3',
        author: 'Alice',
        date: '2026-01-15T10:00:00Z',
      },
    });
    const xml = XMLBuilder.elementToString(row.toXML());
    // w:ins is treated by XMLBuilder as a container element, so it may emit
    // open/close pair instead of self-closing. Both are valid per schema.
    expect(xml).toMatch(
      /<w:ins w:id="3" w:author="Alice" w:date="2026-01-15T10:00:00Z"(\/>|><\/w:ins>)/
    );
  });
});

describe('TableRow tracked deletion (w:del in w:trPr, §17.13.5.14)', () => {
  it('parses <w:del> in <w:trPr> as rowDeletion metadata', async () => {
    const buffer = await makeDocxWithRowIns('del');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const table = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const row = table.getRows()[0]!;
    expect(row.getFormatting().rowDeletion).toBeDefined();
    expect(row.getFormatting().rowDeletion?.author).toBe('Jane Doe');
    expect(row.getFormatting().rowDeletion?.id).toBe('7');
    doc.dispose();
  });

  it('emits <w:del> inside <w:trPr> for rowDeletion', () => {
    const row = new TableRow(1, {
      rowDeletion: {
        id: '5',
        author: 'Bob',
        date: '2026-01-15T10:00:00Z',
      },
    });
    const xml = XMLBuilder.elementToString(row.toXML());
    // w:del is treated by XMLBuilder as a container element; both forms valid.
    expect(xml).toMatch(
      /<w:del w:id="5" w:author="Bob" w:date="2026-01-15T10:00:00Z"(\/>|><\/w:del>)/
    );
  });
});

describe('TableRow CT_TrPr child order (§17.4.79)', () => {
  it('emits ins/del AFTER CT_TrPrBase children and BEFORE trPrChange', () => {
    const row = new TableRow(1, {
      cantSplit: true,
      isHeader: true,
      rowInsertion: { id: '1', author: 'X', date: '2026-01-15T10:00:00Z' },
      rowDeletion: { id: '2', author: 'X', date: '2026-01-15T10:00:00Z' },
    });
    row.setTrPrChange({
      id: '99',
      author: 'X',
      date: '2026-01-15T10:00:00Z',
      previousProperties: {},
    });

    const xml = XMLBuilder.elementToString(row.toXML());
    const cantSplitIdx = xml.indexOf('<w:cantSplit');
    const tblHeaderIdx = xml.indexOf('<w:tblHeader');
    const insIdx = xml.indexOf('<w:ins ');
    const delIdx = xml.indexOf('<w:del ');
    const trPrChangeIdx = xml.indexOf('<w:trPrChange');

    expect(cantSplitIdx).toBeGreaterThan(-1);
    expect(tblHeaderIdx).toBeGreaterThan(-1);
    expect(insIdx).toBeGreaterThan(-1);
    expect(delIdx).toBeGreaterThan(-1);
    expect(trPrChangeIdx).toBeGreaterThan(-1);

    // CT_TrPrBase children FIRST
    expect(cantSplitIdx).toBeLessThan(insIdx);
    expect(tblHeaderIdx).toBeLessThan(insIdx);
    // ins BEFORE del (EG sequence)
    expect(insIdx).toBeLessThan(delIdx);
    // trPrChange LAST
    expect(delIdx).toBeLessThan(trPrChangeIdx);
  });
});

describe('TableRow tracked ins/del round-trip (full DOCX)', () => {
  it('round-trips rowInsertion through Document save/load', async () => {
    const buffer = await makeDocxWithRowIns('ins');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const rebuffered = await doc.toBuffer();
    doc.dispose();

    const reloaded = await Document.loadFromBuffer(rebuffered, { revisionHandling: 'preserve' });
    const table = reloaded.getBodyElements().find((el) => el instanceof Table) as Table;
    const row = table.getRows()[0]!;
    expect(row.getFormatting().rowInsertion).toBeDefined();
    expect(row.getFormatting().rowInsertion?.author).toBe('Jane Doe');
    reloaded.dispose();
  });

  it('round-trips rowDeletion through Document save/load', async () => {
    const buffer = await makeDocxWithRowIns('del');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const rebuffered = await doc.toBuffer();
    doc.dispose();

    const reloaded = await Document.loadFromBuffer(rebuffered, { revisionHandling: 'preserve' });
    const table = reloaded.getBodyElements().find((el) => el instanceof Table) as Table;
    const row = table.getRows()[0]!;
    expect(row.getFormatting().rowDeletion).toBeDefined();
    expect(row.getFormatting().rowDeletion?.author).toBe('Jane Doe');
    reloaded.dispose();
  });
});
