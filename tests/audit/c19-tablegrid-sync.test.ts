/**
 * addColumn()/removeColumn() must keep formatting.tableGrid in sync. toXML()
 * trusts the persisted grid for the gridCol count whenever it is set — and
 * the parser sets it for every loaded table with a w:tblGrid — so a stale
 * grid serializes the wrong number of w:gridCol entries and Word applies
 * surviving widths to the wrong columns in fixed-layout tables
 * (ECMA-376 §17.4.49).
 */
import { Table } from '../../src/elements/Table';
import { Document } from '../../src/core/Document';

const JSZip = require('jszip');

function countGridCols(table: Table): number {
  const xml = table.toXML();
  const tblGrid = (xml.children ?? []).find((c: any) => c.name === 'w:tblGrid') as any;
  return (tblGrid?.children ?? []).filter((c: any) => c.name === 'w:gridCol').length;
}

describe('C19: addColumn/removeColumn keep w:tblGrid in sync', () => {
  it('appends a grid width when adding a column at the end', () => {
    const table = new Table(2, 3);
    table.setTableGrid([2000, 3000, 4000]);

    table.addColumn();

    expect(table.getTableGrid()).toEqual([2000, 3000, 4000, 4000]);
    expect(countGridCols(table)).toBe(4);
  });

  it('inserts the average of neighboring widths when adding a column in the middle', () => {
    const table = new Table(2, 3);
    table.setTableGrid([2000, 3000, 4000]);

    table.addColumn(1);

    expect(table.getTableGrid()).toEqual([2000, 2500, 3000, 4000]);
    expect(countGridCols(table)).toBe(4);
  });

  it('removes the corresponding grid width when removing a column', () => {
    const table = new Table(2, 3);
    table.setTableGrid([2000, 3000, 4000]);

    expect(table.removeColumn(1)).toBe(true);

    expect(table.getTableGrid()).toEqual([2000, 4000]);
    expect(countGridCols(table)).toBe(2);
  });

  it('shrinks the grid when removeEmptyColumns drops columns', () => {
    const table = Table.fromArray([
      ['Name', '', 'Age'],
      ['Alice', '', '30'],
    ]);
    table.setTableGrid([2000, 3000, 4000]);

    expect(table.removeEmptyColumns()).toBe(1);

    expect(table.getTableGrid()).toEqual([2000, 4000]);
  });

  it('leaves the grid alone when a tracked removeColumn keeps the cells', () => {
    const doc = Document.create();
    try {
      doc.enableTrackChanges({ author: 'Author' });
      const table = new Table(2, 3);
      doc.addTable(table);
      table.setTableGrid([2000, 3000, 4000]);

      expect(table.removeColumn(1)).toBe(true);

      // Cells stay in place with cellDel markers, so the grid must not shrink
      expect(table.getRow(0)!.getCellCount()).toBe(3);
      expect(table.getTableGrid()).toEqual([2000, 3000, 4000]);
    } finally {
      doc.dispose();
    }
  });

  it('keeps gridCol count equal to row cell count after mutating a loaded table', async () => {
    const doc = Document.create();
    let loaded: Document | undefined;
    try {
      const table = doc.createTable(2, 3);
      table.setTableGrid([2000, 3000, 4000]);
      const buffer = await doc.toBuffer();
      loaded = await Document.loadFromBuffer(buffer);

      const loadedTable = loaded.getTables()[0]!;
      loadedTable.addColumn();
      loadedTable.removeColumn(0);

      const savedBuffer = await loaded.toBuffer();
      const zip = await JSZip.loadAsync(savedBuffer);
      const docXml = await zip.file('word/document.xml')!.async('string');

      const gridCols = docXml.match(/<w:gridCol[^>]*\/>/g) ?? [];
      expect(gridCols.length).toBe(3);
      const firstRow = /<w:tr[ >][\s\S]*?<\/w:tr>/.exec(docXml)![0];
      expect(firstRow.match(/<w:tc[ >]/g)!.length).toBe(3);
    } finally {
      loaded?.dispose();
      doc.dispose();
    }
  });
});
