/**
 * Tests for Table.findCell(), Table.filterRows(), Table.removeEmptyRows(), Table.removeEmptyColumns()
 */

import { Table } from '../../src/elements/Table';

describe('Table.findCell()', () => {
  it('finds a cell by text content', () => {
    const table = Table.fromArray([
      ['Name', 'Age'],
      ['Alice', '30'],
      ['Bob', '25'],
    ]);

    const result = table.findCell((cell) => cell.getText() === 'Bob');

    expect(result).toBeDefined();
    expect(result!.row).toBe(2);
    expect(result!.col).toBe(0);
    expect(result!.cell.getText()).toBe('Bob');
  });

  it('finds first match in row-major order', () => {
    const table = Table.fromArray([
      ['A', 'X'],
      ['X', 'B'],
    ]);

    const result = table.findCell((cell) => cell.getText() === 'X');

    expect(result!.row).toBe(0);
    expect(result!.col).toBe(1);
  });

  it('returns undefined when no match', () => {
    const table = Table.fromArray([['A', 'B']]);
    expect(table.findCell((cell) => cell.getText() === 'Z')).toBeUndefined();
  });

  it('returns undefined for empty table', () => {
    const table = new Table(0, 0);
    expect(table.findCell(() => true)).toBeUndefined();
  });

  it('provides row and col indices to predicate', () => {
    const table = Table.fromArray([
      ['A', 'B'],
      ['C', 'D'],
    ]);

    const result = table.findCell((_cell, row, col) => row === 1 && col === 1);

    expect(result!.cell.getText()).toBe('D');
  });

  it('finds cell with partial text match', () => {
    const table = Table.fromArray([['Revenue: $100'], ['Costs: $50'], ['Total: $50']]);

    const result = table.findCell((cell) => cell.getText().startsWith('Total'));

    expect(result!.row).toBe(2);
    expect(result!.cell.getText()).toBe('Total: $50');
  });
});

describe('Table.filterRows()', () => {
  it('returns indices of matching rows', () => {
    const table = Table.fromArray([
      ['Name', 'Status'],
      ['Alice', 'active'],
      ['Bob', 'inactive'],
      ['Charlie', 'active'],
    ]);

    const active = table.filterRows((cells) => cells[1]?.getText() === 'active');

    expect(active).toEqual([1, 3]);
  });

  it('returns empty array when no match', () => {
    const table = Table.fromArray([['A'], ['B']]);
    const result = table.filterRows((cells) => cells[0]?.getText() === 'Z');

    expect(result).toEqual([]);
  });

  it('returns all indices when all match', () => {
    const table = Table.fromArray([['X'], ['X'], ['X']]);
    const result = table.filterRows((cells) => cells[0]?.getText() === 'X');

    expect(result).toEqual([0, 1, 2]);
  });

  it('provides row index to predicate', () => {
    const table = Table.fromArray([['A'], ['B'], ['C'], ['D']]);

    // Select even rows
    const even = table.filterRows((_cells, rowIndex) => rowIndex % 2 === 0);

    expect(even).toEqual([0, 2]);
  });

  it('finds empty rows', () => {
    const table = Table.fromArray([
      ['Data', 'Here'],
      ['', ''],
      ['More', 'Data'],
      ['', ''],
    ]);

    const empty = table.filterRows((cells) => cells.every((c) => c.getText().trim() === ''));

    expect(empty).toEqual([1, 3]);
  });
});

describe('Table.removeEmptyRows()', () => {
  it('removes rows where all cells are empty', () => {
    const table = Table.fromArray([
      ['Name', 'Age'],
      ['', ''],
      ['Alice', '30'],
      ['', ''],
    ]);

    const removed = table.removeEmptyRows();

    expect(removed).toBe(2);
    expect(table.getRowCount()).toBe(2);
    expect(table.toArray()).toEqual([
      ['Name', 'Age'],
      ['Alice', '30'],
    ]);
  });

  it('keeps rows with any non-empty cell', () => {
    const table = Table.fromArray([
      ['', 'Data'],
      ['', ''],
      ['Data', ''],
    ]);

    const removed = table.removeEmptyRows();

    expect(removed).toBe(1);
    expect(table.getRowCount()).toBe(2);
  });

  it('returns 0 when no empty rows', () => {
    const table = Table.fromArray([
      ['A', 'B'],
      ['C', 'D'],
    ]);

    expect(table.removeEmptyRows()).toBe(0);
    expect(table.getRowCount()).toBe(2);
  });

  it('keeps at least one row when all are empty', () => {
    const table = Table.fromArray([
      ['', ''],
      ['', ''],
      ['', ''],
    ]);

    const removed = table.removeEmptyRows();

    expect(removed).toBe(2);
    expect(table.getRowCount()).toBe(1);
  });

  it('treats whitespace-only cells as empty', () => {
    const table = Table.fromArray([['Data'], ['   '], [' \t ']]);

    const removed = table.removeEmptyRows();

    expect(removed).toBe(2);
    expect(table.getRowCount()).toBe(1);
  });

  it('handles single-row table (no removal)', () => {
    const table = Table.fromArray([['']]);
    const removed = table.removeEmptyRows();

    expect(removed).toBe(0);
    expect(table.getRowCount()).toBe(1);
  });
});

describe('Table.removeEmptyColumns()', () => {
  it('removes columns where all cells are empty', () => {
    const table = Table.fromArray([
      ['Name', '', 'Age'],
      ['Alice', '', '30'],
      ['Bob', '', '25'],
    ]);

    const removed = table.removeEmptyColumns();

    expect(removed).toBe(1);
    expect(table.toArray()).toEqual([
      ['Name', 'Age'],
      ['Alice', '30'],
      ['Bob', '25'],
    ]);
  });

  it('removes multiple empty columns', () => {
    const table = Table.fromArray([
      ['', 'A', '', 'B', ''],
      ['', 'C', '', 'D', ''],
    ]);

    const removed = table.removeEmptyColumns();

    expect(removed).toBe(3);
    expect(table.toArray()).toEqual([
      ['A', 'B'],
      ['C', 'D'],
    ]);
  });

  it('returns 0 when no empty columns', () => {
    const table = Table.fromArray([
      ['A', 'B'],
      ['C', 'D'],
    ]);

    expect(table.removeEmptyColumns()).toBe(0);
  });

  it('keeps columns with any non-empty cell', () => {
    const table = Table.fromArray([
      ['', 'B'],
      ['A', ''],
    ]);

    expect(table.removeEmptyColumns()).toBe(0);
    expect(table.getColumnCount()).toBe(2);
  });

  it('handles single-column table', () => {
    const table = Table.fromArray([[''], [''], ['']]);
    const removed = table.removeEmptyColumns();

    // removeColumn doesn't have the min-1 constraint like removeRow
    expect(removed).toBe(1);
  });

  it('treats whitespace-only cells as empty', () => {
    const table = Table.fromArray([
      ['A', '   '],
      ['B', ' \t '],
    ]);

    const removed = table.removeEmptyColumns();

    expect(removed).toBe(1);
  });
});

describe('combined usage', () => {
  it('cleanup pipeline: remove empty rows then empty columns', () => {
    const table = Table.fromArray([
      ['Name', '', 'Age', ''],
      ['', '', '', ''],
      ['Alice', '', '30', ''],
      ['', '', '', ''],
      ['Bob', '', '25', ''],
    ]);

    const rowsRemoved = table.removeEmptyRows();
    const colsRemoved = table.removeEmptyColumns();

    expect(rowsRemoved).toBe(2);
    expect(colsRemoved).toBe(2);
    expect(table.toArray()).toEqual([
      ['Name', 'Age'],
      ['Alice', '30'],
      ['Bob', '25'],
    ]);
  });

  it('findCell + filterRows for data queries', () => {
    const table = Table.fromArray([
      ['Product', 'Price', 'Stock'],
      ['Widget', '10', '100'],
      ['Gadget', '25', '0'],
      ['Doohickey', '5', '50'],
    ]);

    // Find the "Price" column
    const priceHeader = table.findCell((c) => c.getText() === 'Price');
    expect(priceHeader).toBeDefined();
    const priceCol = priceHeader!.col;

    // Find rows with zero stock
    const outOfStock = table.filterRows((cells) => cells[2]?.getText() === '0');
    expect(outOfStock).toEqual([2]);

    // Get the product name for out-of-stock items
    const productName = table.getCell(outOfStock[0]!, 0)!.getText();
    expect(productName).toBe('Gadget');
  });
});
