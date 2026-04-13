/**
 * Tests for TableCell.clone(), TableRow.clone(), and Table.duplicateRow()
 */

import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';
import { TableCell } from '../../src/elements/TableCell';
import { Paragraph } from '../../src/elements/Paragraph';

describe('TableCell.clone()', () => {
  it('clones an empty cell', () => {
    const cell = new TableCell();
    const clone = cell.clone();

    expect(clone).not.toBe(cell);
    expect(clone.getParagraphs()).toHaveLength(0);
  });

  it('deep-clones formatting', () => {
    const cell = new TableCell({
      width: 2400,
      widthType: 'dxa',
      columnSpan: 2,
      shading: { fill: 'FF0000', pattern: 'clear' },
    });
    const clone = cell.clone();

    expect(clone.getFormatting().width).toBe(2400);
    expect(clone.getFormatting().widthType).toBe('dxa');
    expect(clone.getFormatting().columnSpan).toBe(2);
    expect(clone.getFormatting().shading?.fill).toBe('FF0000');

    // Mutating clone formatting does not affect original
    clone.setWidth(4800);
    expect(cell.getFormatting().width).toBe(2400);
  });

  it('deep-clones paragraphs', () => {
    const cell = new TableCell();
    cell.createParagraph('Hello');
    cell.createParagraph('World');

    const clone = cell.clone();
    expect(clone.getParagraphs()).toHaveLength(2);
    expect(clone.getText()).toBe('Hello\nWorld');

    // Mutating cloned paragraphs does not affect original
    clone.getParagraphs()[0]!.setText('Changed');
    expect(cell.getParagraphs()[0]!.getText()).toBe('Hello');
  });

  it('clones raw nested content', () => {
    const cell = new TableCell();
    cell.createParagraph('Before');
    cell.addRawNestedContent(0, '<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>', 'table');

    const clone = cell.clone();
    expect(clone.hasNestedTables()).toBe(true);

    const rawContent = clone.getRawNestedContent();
    expect(rawContent).toHaveLength(1);
    expect(rawContent[0]!.type).toBe('table');
  });

  it('clones tcPrChange tracking', () => {
    const cell = new TableCell({ width: 3000 });
    cell.setTcPrChange({
      author: 'Test',
      date: '2024-01-01T00:00:00Z',
      id: '1',
      previousProperties: { width: 2000 },
    });

    const clone = cell.clone();
    const change = clone.getTcPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('Test');
    expect(change!.previousProperties.width).toBe(2000);

    // Deep clone: mutating clone's change doesn't affect original
    change!.previousProperties.width = 9999;
    expect(cell.getTcPrChange()!.previousProperties.width).toBe(2000);
  });

  it('does not copy parent row reference', () => {
    const row = new TableRow();
    const cell = row.createCell('Test');

    const clone = cell.clone();
    // Clone should work independently (no parent set)
    expect(clone.getText()).toBe('Test');
  });
});

describe('TableRow.clone()', () => {
  it('clones an empty row', () => {
    const row = new TableRow();
    const clone = row.clone();

    expect(clone).not.toBe(row);
    expect(clone.getCellCount()).toBe(0);
  });

  it('deep-clones cells and formatting', () => {
    const row = new TableRow(3, { isHeader: true, height: 720, heightRule: 'exact' });
    row.getCell(0)!.createParagraph('A');
    row.getCell(1)!.createParagraph('B');
    row.getCell(2)!.createParagraph('C');

    const clone = row.clone();
    expect(clone.getCellCount()).toBe(3);
    expect(clone.getIsHeader()).toBe(true);
    expect(clone.getHeight()).toBe(720);
    expect(clone.getHeightRule()).toBe('exact');

    // Cells are deep clones
    expect(clone.getCell(0)!.getText()).toBe('A');
    clone.getCell(0)!.createParagraph(' modified');
    expect(row.getCell(0)!.getParagraphs()).toHaveLength(1);
    expect(clone.getCell(0)!.getParagraphs()).toHaveLength(2);
  });

  it('deep-clones formatting independently', () => {
    const row = new TableRow(1, { cantSplit: true, justification: 'center' });
    const clone = row.clone();

    clone.setCantSplit(false);
    clone.setJustification('left');

    expect(row.getCantSplit()).toBe(true);
    expect(row.getJustification()).toBe('center');
  });

  it('clones trPrChange tracking', () => {
    const row = new TableRow(2);
    row.setTrPrChange({
      author: 'User',
      date: '2024-06-15T12:00:00Z',
      id: '42',
      previousProperties: { height: 500 },
    });

    const clone = row.clone();
    const change = clone.getTrPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('User');
    expect(change!.id).toBe('42');

    // Deep clone
    change!.previousProperties.height = 9999;
    expect(row.getTrPrChange()!.previousProperties.height).toBe(500);
  });

  it('preserves cell column spans', () => {
    const row = new TableRow();
    const cell = new TableCell({ columnSpan: 3, width: 7200, widthType: 'dxa' });
    cell.createParagraph('Merged');
    row.addCell(cell);

    const clone = row.clone();
    expect(clone.getCell(0)!.getFormatting().columnSpan).toBe(3);
    expect(clone.getTotalGridSpan()).toBe(3);
  });
});

describe('Table.duplicateRow()', () => {
  it('duplicates a single row', () => {
    const table = new Table(2, 3);
    table.getCell(0, 0)!.createParagraph('Header');
    table.getCell(1, 0)!.createParagraph('Data');

    const [copy] = table.duplicateRow(1);
    expect(table.getRowCount()).toBe(3);
    expect(copy!.getCell(0)!.getText()).toBe('Data');

    // Modify copy, original untouched
    copy!.getCell(0)!.getParagraphs()[0]!.setText('Modified');
    expect(table.getRow(1)!.getCell(0)!.getText()).toBe('Data');
    expect(table.getRow(2)!.getCell(0)!.getText()).toBe('Modified');
  });

  it('duplicates multiple copies', () => {
    const table = new Table(1, 2);
    table.getCell(0, 0)!.createParagraph('Template');
    table.getCell(0, 1)!.createParagraph('Row');

    const copies = table.duplicateRow(0, 3);
    expect(copies).toHaveLength(3);
    expect(table.getRowCount()).toBe(4);

    // All copies have same content
    for (const copy of copies) {
      expect(copy.getCell(0)!.getText()).toBe('Template');
      expect(copy.getCell(1)!.getText()).toBe('Row');
    }

    // Each copy is independent
    copies[0]!.getCell(0)!.getParagraphs()[0]!.setText('A');
    copies[1]!.getCell(0)!.getParagraphs()[0]!.setText('B');
    copies[2]!.getCell(0)!.getParagraphs()[0]!.setText('C');

    expect(table.getRow(0)!.getCell(0)!.getText()).toBe('Template');
    expect(table.getRow(1)!.getCell(0)!.getText()).toBe('A');
    expect(table.getRow(2)!.getCell(0)!.getText()).toBe('B');
    expect(table.getRow(3)!.getCell(0)!.getText()).toBe('C');
  });

  it('throws RangeError for out-of-bounds index', () => {
    const table = new Table(2, 2);

    expect(() => table.duplicateRow(-1)).toThrow(RangeError);
    expect(() => table.duplicateRow(2)).toThrow(RangeError);
    expect(() => table.duplicateRow(99)).toThrow(RangeError);
  });

  it('returns empty array for count < 1', () => {
    const table = new Table(2, 2);
    const result = table.duplicateRow(0, 0);
    expect(result).toEqual([]);
    expect(table.getRowCount()).toBe(2);
  });

  it('preserves row formatting in duplicates', () => {
    const table = new Table(2, 2);
    table.getRow(0)!.setHeader(true).setHeight(720, 'exact').setCantSplit(true);

    const [copy] = table.duplicateRow(0);
    expect(copy!.getIsHeader()).toBe(true);
    expect(copy!.getHeight()).toBe(720);
    expect(copy!.getHeightRule()).toBe('exact');
    expect(copy!.getCantSplit()).toBe(true);
  });

  it('preserves cell formatting in duplicates', () => {
    const table = new Table(1, 2);
    const cell = table.getCell(0, 0)!;
    cell.setShading({ fill: 'FFFF00', pattern: 'clear' });
    cell.setVerticalAlignment('center');

    const [copy] = table.duplicateRow(0);
    const clonedCell = copy!.getCell(0)!;
    expect(clonedCell.getFormatting().shading?.fill).toBe('FFFF00');
    expect(clonedCell.getFormatting().verticalAlignment).toBe('center');
  });

  it('inserts duplicates after the source row', () => {
    const table = new Table(3, 1);
    table.getCell(0, 0)!.createParagraph('Row 0');
    table.getCell(1, 0)!.createParagraph('Row 1');
    table.getCell(2, 0)!.createParagraph('Row 2');

    table.duplicateRow(1, 2);
    expect(table.getRowCount()).toBe(5);

    // Order: Row 0, Row 1, Row 1 (copy), Row 1 (copy), Row 2
    expect(table.getRow(0)!.getCell(0)!.getText()).toBe('Row 0');
    expect(table.getRow(1)!.getCell(0)!.getText()).toBe('Row 1');
    expect(table.getRow(2)!.getCell(0)!.getText()).toBe('Row 1');
    expect(table.getRow(3)!.getCell(0)!.getText()).toBe('Row 1');
    expect(table.getRow(4)!.getCell(0)!.getText()).toBe('Row 2');
  });
});

describe('Table.clone() uses TableRow.clone()', () => {
  it('deep-clones rows with formatting', () => {
    const table = new Table(2, 2);
    table.getRow(0)!.setHeader(true);
    table.getCell(0, 0)!.createParagraph('Header');
    table.getCell(1, 0)!.createParagraph('Data');

    const clone = table.clone();
    expect(clone.getRowCount()).toBe(2);
    expect(clone.getRow(0)!.getIsHeader()).toBe(true);
    expect(clone.getCell(0, 0)!.getText()).toBe('Header');

    // Independence
    clone.getCell(0, 0)!.getParagraphs()[0]!.setText('Changed');
    expect(table.getCell(0, 0)!.getText()).toBe('Header');
  });

  it('preserves raw nested content through clone chain', () => {
    const table = new Table(1, 1);
    const cell = table.getCell(0, 0)!;
    cell.createParagraph('Main');
    cell.addRawNestedContent(0, '<w:tbl/>', 'table');

    const clone = table.clone();
    const clonedCell = clone.getCell(0, 0)!;
    expect(clonedCell.hasNestedTables()).toBe(true);
    expect(clonedCell.getRawNestedContent()).toHaveLength(1);
  });
});

describe('Table.getColumnCells()', () => {
  it('returns all cells in a column', () => {
    const table = new Table(3, 4);
    table.getCell(0, 1)!.createParagraph('R0C1');
    table.getCell(1, 1)!.createParagraph('R1C1');
    table.getCell(2, 1)!.createParagraph('R2C1');

    const cells = table.getColumnCells(1);
    expect(cells).toHaveLength(3);
    expect(cells[0]!.getText()).toBe('R0C1');
    expect(cells[1]!.getText()).toBe('R1C1');
    expect(cells[2]!.getText()).toBe('R2C1');
  });

  it('returns empty array for empty table', () => {
    const table = new Table(0, 0);
    expect(table.getColumnCells(0)).toEqual([]);
  });

  it('skips rows with fewer cells', () => {
    const table = new Table(3, 2);
    // Remove last cell from middle row to create jagged table
    table.getRow(1)!.removeCellAt(1);

    const cells = table.getColumnCells(1);
    expect(cells).toHaveLength(2); // Only rows 0 and 2
  });

  it('returns first column correctly', () => {
    const table = new Table(2, 3);
    table.getCell(0, 0)!.createParagraph('A');
    table.getCell(1, 0)!.createParagraph('B');

    const cells = table.getColumnCells(0);
    expect(cells).toHaveLength(2);
    expect(cells[0]!.getText()).toBe('A');
    expect(cells[1]!.getText()).toBe('B');
  });

  it('returns empty array for out-of-range column', () => {
    const table = new Table(2, 2);
    expect(table.getColumnCells(99)).toEqual([]);
  });
});

describe('Table.forEachCell()', () => {
  it('iterates all cells with correct indices', () => {
    const table = new Table(2, 3);
    const visited: [number, number][] = [];

    table.forEachCell((row, col) => {
      visited.push([row, col]);
    });

    expect(visited).toEqual([
      [0, 0],
      [0, 1],
      [0, 2],
      [1, 0],
      [1, 1],
      [1, 2],
    ]);
  });

  it('provides actual cell references', () => {
    const table = new Table(2, 2);
    table.getCell(1, 0)!.createParagraph('Target');

    let foundText = '';
    table.forEachCell((row, col, cell) => {
      if (row === 1 && col === 0) {
        foundText = cell.getText();
      }
    });

    expect(foundText).toBe('Target');
  });

  it('supports early termination by returning false', () => {
    const table = new Table(3, 3);
    const visited: [number, number][] = [];

    table.forEachCell((row, col) => {
      visited.push([row, col]);
      if (row === 1 && col === 0) return false;
      return undefined;
    });

    // Should stop after (1,0) — visited row 0 fully plus (1,0)
    expect(visited).toEqual([
      [0, 0],
      [0, 1],
      [0, 2],
      [1, 0],
    ]);
  });

  it('handles empty table', () => {
    const table = new Table(0, 0);
    const visited: [number, number][] = [];

    table.forEachCell((row, col) => {
      visited.push([row, col]);
    });

    expect(visited).toEqual([]);
  });

  it('can be used for bulk formatting', () => {
    const table = new Table(3, 2);

    // Apply alternating row shading
    table.forEachCell((row, _col, cell) => {
      if (row % 2 === 1) {
        cell.setShading({ fill: 'F0F0F0', pattern: 'clear' });
      }
    });

    expect(table.getCell(0, 0)!.getFormatting().shading).toBeUndefined();
    expect(table.getCell(1, 0)!.getFormatting().shading?.fill).toBe('F0F0F0');
    expect(table.getCell(1, 1)!.getFormatting().shading?.fill).toBe('F0F0F0');
    expect(table.getCell(2, 0)!.getFormatting().shading).toBeUndefined();
  });

  it('can find a cell by content', () => {
    const table = new Table(3, 3);
    table.getCell(2, 1)!.createParagraph('Total: 42');

    let found: { row: number; col: number } | undefined;
    table.forEachCell((row, col, cell) => {
      if (cell.getText().includes('Total')) {
        found = { row, col };
        return false;
      }
      return undefined;
    });

    expect(found).toEqual({ row: 2, col: 1 });
  });
});
