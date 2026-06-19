/**
 * Per ECMA-376 §17.4.14/§17.4.15, w:gridBefore/w:gridAfter consume grid
 * columns before the first and after the last cell of a row, so a row's grid
 * footprint is gridBefore + sum(gridSpan) + gridAfter. The auto-generated
 * w:tblGrid (Table.toXML when no explicit grid is set) must declare that many
 * w:gridCol entries — otherwise offset rows extend past the declared grid and
 * Word has to infer the missing columns, shifting the layout.
 */
import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';
import { Document } from '../../src/core/Document';

const JSZip = require('jszip');

function countGridCols(table: Table): number {
  const xml = table.toXML();
  const tblGrid = (xml.children ?? []).find((c: any) => c.name === 'w:tblGrid') as any;
  return (tblGrid?.children ?? []).filter((c: any) => c.name === 'w:gridCol').length;
}

describe('C46: auto-generated tblGrid accounts for gridBefore/gridAfter', () => {
  it('includes gridBefore and gridAfter in getTotalGridSpan()', () => {
    const row = new TableRow(2);
    row.setGridBefore(2);
    row.setGridAfter(1);

    // 2 before + 2 cells + 1 after
    expect(row.getTotalGridSpan()).toBe(5);
  });

  it('combines gridBefore/gridAfter with cell column spans', () => {
    const row = new TableRow();
    row.createCell('A');
    row.createCell('B').setColumnSpan(2);
    row.setGridBefore(1);
    row.setGridAfter(2);

    // 1 before + (1 + 2) cell spans + 2 after
    expect(row.getTotalGridSpan()).toBe(6);
  });

  it('auto-generates enough gridCol entries for a row offset by gridBefore', () => {
    const table = new Table(2, 3); // no explicit tableGrid -> auto-generated
    table.getRow(0)!.setGridBefore(2);

    // Row 0 occupies 2 + 3 = 5 grid columns
    expect(countGridCols(table)).toBe(5);
  });

  it('sizes new rows from insertRows to the full grid footprint', () => {
    const table = new Table(1, 3);
    table.getRow(0)!.setGridBefore(1).setGridAfter(1);

    const [inserted] = table.insertRows(1, 1);

    // Full-width row must cover gridBefore + cells + gridAfter of the widest row
    expect(inserted!.getCellCount()).toBe(5);
  });

  it('saves a tblGrid wide enough for gridBefore/gridAfter rows', async () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(2, 3);
      table.getRow(0)!.setGridBefore(1);
      table.getRow(0)!.setGridAfter(1);

      const buffer = await doc.toBuffer();
      const zip = await JSZip.loadAsync(buffer);
      const docXml = await zip.file('word/document.xml')!.async('string');

      const gridCols = docXml.match(/<w:gridCol[^>]*\/>/g) ?? [];
      expect(gridCols.length).toBe(5);
      expect(docXml).toContain('<w:gridBefore w:val="1"/>');
      expect(docXml).toContain('<w:gridAfter w:val="1"/>');
    } finally {
      doc.dispose();
    }
  });
});
