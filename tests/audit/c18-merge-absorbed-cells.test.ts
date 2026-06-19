/**
 * Horizontal mergeCells() must remove the absorbed w:tc elements from the
 * row. Per ECMA-376 §17.4.17 a gridSpan cell consumes the columns it spans;
 * leaving the absorbed cells behind widens the table grid (span 3 + 1 + 1
 * = 5 grid columns for a 3-column table) and misaligns every other row
 * beneath the merged cell.
 */
import { Table } from '../../src/elements/Table';
import { Document } from '../../src/core/Document';

const JSZip = require('jszip');

function countChildren(element: any, name: string): number {
  return (element?.children ?? []).filter((c: any) => c.name === name).length;
}

describe('C18: horizontal mergeCells removes absorbed cells', () => {
  it('removes absorbed cells from a single-row merge and migrates their content', () => {
    const table = new Table(3, 3);
    table.getCell(0, 0)!.createParagraph('A');
    table.getCell(0, 1)!.createParagraph('B');
    table.getCell(0, 2)!.createParagraph('C');

    table.mergeCells(0, 0, 0, 2);

    const row0 = table.getRow(0)!;
    expect(row0.getCellCount()).toBe(1);
    const merged = row0.getCell(0)!;
    expect(merged.getColumnSpan()).toBe(3);
    expect(merged.getText()).toBe('A\nB\nC');
    // Untouched rows keep their cells
    expect(table.getRow(1)!.getCellCount()).toBe(3);
    expect(table.getRow(2)!.getCellCount()).toBe(3);
  });

  it('serializes a 3-column tblGrid instead of widening to 5', () => {
    const table = new Table(3, 3);
    table.mergeCells(0, 0, 0, 2);

    const xml = table.toXML();
    const tblGrid = (xml.children ?? []).find((c: any) => c.name === 'w:tblGrid');
    expect(countChildren(tblGrid, 'w:gridCol')).toBe(3);

    const rows = (xml.children ?? []).filter((c: any) => c.name === 'w:tr');
    expect(countChildren(rows[0], 'w:tc')).toBe(1);
    expect(countChildren(rows[1], 'w:tc')).toBe(3);
  });

  it('removes absorbed cells in every row of a block merge, keeping vMerge continue cells', () => {
    const table = new Table(3, 3);
    table.mergeCells(0, 0, 1, 1);

    expect(table.getRow(0)!.getCellCount()).toBe(2);
    expect(table.getRow(1)!.getCellCount()).toBe(2);
    expect(table.getRow(2)!.getCellCount()).toBe(3);

    // Continuation cell stays with vMerge continue and matching gridSpan
    const continuation = table.getCell(1, 0)!;
    expect(continuation.getVerticalMerge()).toBe('continue');
    expect(continuation.getColumnSpan()).toBe(2);

    // Every row spans the same 3 grid columns
    for (const row of table.getRows()) {
      expect(row.getTotalGridSpan()).toBe(3);
    }
  });

  it('round-trips a merged table without widening the saved grid', async () => {
    const doc = Document.create();
    let loaded: Document | undefined;
    try {
      const table = doc.createTable(3, 3);
      table.getCell(0, 0)!.createParagraph('Merged');
      table.mergeCells(0, 0, 0, 2);

      const buffer = await doc.toBuffer();
      const zip = await JSZip.loadAsync(buffer);
      const docXml = await zip.file('word/document.xml')!.async('string');
      const gridCols = docXml.match(/<w:gridCol[^>]*\/>/g) ?? [];
      expect(gridCols.length).toBe(3);

      loaded = await Document.loadFromBuffer(buffer);
      const loadedTable = loaded.getTables()[0]!;
      expect(loadedTable.getRow(0)!.getCellCount()).toBe(1);
      expect(loadedTable.getRow(0)!.getCell(0)!.getColumnSpan()).toBe(3);
    } finally {
      loaded?.dispose();
      doc.dispose();
    }
  });
});
