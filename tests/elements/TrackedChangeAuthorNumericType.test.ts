/**
 * Tracked-change author/date attributes — numeric-looking values must
 * round-trip as strings (type-contract safety).
 *
 * Per ECMA-376 Part 1:
 *   - `w:author` on CT_TrackChange is ST_String (§17.18.84)
 *   - `w:date` on CT_TrackChange is ST_DateTime (§17.18.10)
 *
 * Affected elements: CT_TrackChange-derived elements including
 *   - `w:cellIns` / `w:cellDel` / `w:cellMerge` (§17.13.5.4-.6)
 *   - `w:tblGridChange` (§17.13.5.35)
 *   - `w:tblPrChange` / `w:trPrChange` / `w:tcPrChange` (§17.13.5.36-.38)
 *
 * Bug: XMLParser's `parseAttributeValue: true` coerces purely-digit
 * `w:author` strings (e.g. a user ID like `"42"`) to JS numbers. Seven
 * parser sites in `DocumentParser.ts` stored the raw coerced value:
 *   - tblGridChange author + date (line 6666-7)
 *   - tblPrChange author + date (line 6999-7000)
 *   - trPrChange author + date (line 7217-8)
 *   - cellIns author + dateAttr (line 7524-6)
 *   - cellDel author + dateAttr (line 7542-4)
 *   - cellMerge author + dateAttr (line 7560-2)
 *   - tcPrChange author + date (line 7587-8)
 *
 * Both `Revision.author` and the various *Change interfaces declare
 * `author` as `string`. Storing a number violates the contract and any
 * downstream `.toLowerCase()` / `.startsWith(...)` on the author would
 * throw.
 *
 * Iteration 127 adds `String(...)` casts at every site and passes
 * `String(dateAttr)` into `new Date(...)` so a coerced numeric date
 * can't be interpreted as an epoch-ms timestamp.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadDocWith(tableBodyXml: string) {
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
    ${tableBodyXml}
    <w:p><w:r><w:t>trailer</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  return Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
}

const tableWithCellIns = (author: string, date: string) => `
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="pct"/>
            <w:cellIns w:id="1" w:author="${author}" w:date="${date}"/>
          </w:tcPr>
          <w:p><w:r><w:t>ins</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

const tableWithCellDel = (author: string) => `
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="pct"/>
            <w:cellDel w:id="1" w:author="${author}" w:date="2026-04-24T10:00:00Z"/>
          </w:tcPr>
          <w:p><w:r><w:t>del</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

const tableWithCellMerge = (author: string) => `
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="pct"/>
            <w:cellMerge w:id="1" w:author="${author}" w:date="2026-04-24T10:00:00Z" w:vMerge="rest"/>
          </w:tcPr>
          <w:p><w:r><w:t>merge</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

const tableWithTcPrChange = (author: string) => `
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="pct"/>
            <w:tcPrChange w:id="2" w:author="${author}" w:date="2026-04-24T10:00:00Z">
              <w:tcPr><w:tcW w:w="3000" w:type="pct"/></w:tcPr>
            </w:tcPrChange>
          </w:tcPr>
          <w:p><w:r><w:t>chg</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

const tableWithTrPrChange = (author: string) => `
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr>
          <w:trPrChange w:id="3" w:author="${author}" w:date="2026-04-24T10:00:00Z">
            <w:trPr><w:cantSplit/></w:trPr>
          </w:trPrChange>
        </w:trPr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>row</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

const tableWithTblPrChange = (author: string) => `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblPrChange w:id="4" w:author="${author}" w:date="2026-04-24T10:00:00Z">
          <w:tblPr><w:tblW w:w="3000" w:type="pct"/></w:tblPr>
        </w:tblPrChange>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>tbl</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

const tableWithTblGridChange = (author: string) => `
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="5000"/>
        <w:tblGridChange w:id="5" w:author="${author}" w:date="2026-04-24T10:00:00Z">
          <w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid>
        </w:tblGridChange>
      </w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>grid</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

describe('Tracked-change author attribute type-contract preservation', () => {
  it('stores <w:cellIns w:author="42"/> as the STRING "42"', async () => {
    const doc = await loadDocWith(tableWithCellIns('42', '2026-04-24T10:00:00Z'));
    const table = doc.getTables()[0]!;
    const cell = table.getRows()[0]!.getCells()[0]!;
    const revision = cell.getCellRevision();
    doc.dispose();
    expect(revision).toBeDefined();
    expect(typeof revision!.getAuthor()).toBe('string');
    expect(revision!.getAuthor()).toBe('42');
  });

  it('stores <w:cellDel w:author="2025"/> as the STRING "2025"', async () => {
    const doc = await loadDocWith(tableWithCellDel('2025'));
    const table = doc.getTables()[0]!;
    const cell = table.getRows()[0]!.getCells()[0]!;
    const revision = cell.getCellRevision();
    doc.dispose();
    expect(typeof revision!.getAuthor()).toBe('string');
    expect(revision!.getAuthor()).toBe('2025');
  });

  it('stores <w:cellMerge w:author="7"/> as the STRING "7"', async () => {
    const doc = await loadDocWith(tableWithCellMerge('7'));
    const table = doc.getTables()[0]!;
    const cell = table.getRows()[0]!.getCells()[0]!;
    const revision = cell.getCellRevision();
    doc.dispose();
    expect(typeof revision!.getAuthor()).toBe('string');
    expect(revision!.getAuthor()).toBe('7');
  });

  it('stores <w:tcPrChange w:author="99"/> as the STRING "99"', async () => {
    const doc = await loadDocWith(tableWithTcPrChange('99'));
    const table = doc.getTables()[0]!;
    const cell = table.getRows()[0]!.getCells()[0]!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const tcPrChange = (cell as any).tcPrChange ?? (cell as any).getTcPrChange?.();
    doc.dispose();
    expect(tcPrChange).toBeDefined();
    expect(typeof tcPrChange.author).toBe('string');
    expect(tcPrChange.author).toBe('99');
  });

  it('stores <w:trPrChange w:author="11"/> as the STRING "11"', async () => {
    const doc = await loadDocWith(tableWithTrPrChange('11'));
    const table = doc.getTables()[0]!;
    const row = table.getRows()[0]!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const trPrChange = (row as any).trPrChange ?? (row as any).getTrPrChange?.();
    doc.dispose();
    expect(trPrChange).toBeDefined();
    expect(typeof trPrChange.author).toBe('string');
    expect(trPrChange.author).toBe('11');
  });

  it('stores <w:tblPrChange w:author="22"/> as the STRING "22"', async () => {
    const doc = await loadDocWith(tableWithTblPrChange('22'));
    const table = doc.getTables()[0]!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const tblPrChange = (table as any).tblPrChange ?? (table as any).getTblPrChange?.();
    doc.dispose();
    expect(tblPrChange).toBeDefined();
    expect(typeof tblPrChange.author).toBe('string');
    expect(tblPrChange.author).toBe('22');
  });

  it('stores <w:tblGridChange w:author="33"/> as the STRING "33"', async () => {
    const doc = await loadDocWith(tableWithTblGridChange('33'));
    const table = doc.getTables()[0]!;
    const gridChange = table.getTblGridChange();
    doc.dispose();
    expect(gridChange).toBeDefined();
    expect(typeof gridChange!.getAuthor()).toBe('string');
    expect(gridChange!.getAuthor()).toBe('33');
  });

  it('string methods are callable on parsed numeric author (cellIns)', async () => {
    const doc = await loadDocWith(tableWithCellIns('42', '2026-04-24T10:00:00Z'));
    const table = doc.getTables()[0]!;
    const cell = table.getRows()[0]!.getCells()[0]!;
    const author = cell.getCellRevision()!.getAuthor();
    doc.dispose();
    // Pre-fix: author was number 42, .startsWith would throw.
    expect(() => author.startsWith('4')).not.toThrow();
    expect(author.startsWith('4')).toBe(true);
  });

  it('preserves non-numeric author "John Doe" (regression guard)', async () => {
    const doc = await loadDocWith(tableWithCellIns('John Doe', '2026-04-24T10:00:00Z'));
    const table = doc.getTables()[0]!;
    const cell = table.getRows()[0]!.getCells()[0]!;
    const author = cell.getCellRevision()!.getAuthor();
    doc.dispose();
    expect(author).toBe('John Doe');
  });
});
