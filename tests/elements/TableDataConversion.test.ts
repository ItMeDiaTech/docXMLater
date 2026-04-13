/**
 * Tests for Table.toArray(), Table.toPlainText(), and Table.fromArray()
 */

import { Table } from '../../src/elements/Table';
import { Paragraph } from '../../src/elements/Paragraph';

describe('Table.toArray()', () => {
  it('extracts simple table data', () => {
    const table = new Table(2, 3);
    table.getCell(0, 0)!.createParagraph('Name');
    table.getCell(0, 1)!.createParagraph('Age');
    table.getCell(0, 2)!.createParagraph('City');
    table.getCell(1, 0)!.createParagraph('Alice');
    table.getCell(1, 1)!.createParagraph('30');
    table.getCell(1, 2)!.createParagraph('NYC');

    const data = table.toArray();
    expect(data).toEqual([
      ['Name', 'Age', 'City'],
      ['Alice', '30', 'NYC'],
    ]);
  });

  it('returns empty arrays for empty cells', () => {
    const table = new Table(2, 2);
    table.getCell(0, 0)!.createParagraph('Only this');

    const data = table.toArray();
    expect(data[0]![0]).toBe('Only this');
    expect(data[0]![1]).toBe('');
    expect(data[1]![0]).toBe('');
    expect(data[1]![1]).toBe('');
  });

  it('joins multi-paragraph cells with newlines', () => {
    const table = new Table(1, 1);
    const cell = table.getCell(0, 0)!;
    cell.createParagraph('Line 1');
    cell.createParagraph('Line 2');

    const data = table.toArray();
    expect(data[0]![0]).toBe('Line 1\nLine 2');
  });

  it('returns empty array for empty table', () => {
    const table = new Table(0, 0);
    expect(table.toArray()).toEqual([]);
  });

  it('handles single-cell table', () => {
    const table = new Table(1, 1);
    table.getCell(0, 0)!.createParagraph('Solo');

    expect(table.toArray()).toEqual([['Solo']]);
  });

  it('handles large tables', () => {
    const table = new Table(10, 5);
    for (let r = 0; r < 10; r++) {
      for (let c = 0; c < 5; c++) {
        table.getCell(r, c)!.createParagraph(`R${r}C${c}`);
      }
    }

    const data = table.toArray();
    expect(data).toHaveLength(10);
    expect(data[0]).toHaveLength(5);
    expect(data[3]![2]).toBe('R3C2');
    expect(data[9]![4]).toBe('R9C4');
  });
});

describe('Table.toPlainText()', () => {
  it('renders tab-separated values by default', () => {
    const table = new Table(2, 2);
    table.getCell(0, 0)!.createParagraph('A');
    table.getCell(0, 1)!.createParagraph('B');
    table.getCell(1, 0)!.createParagraph('C');
    table.getCell(1, 1)!.createParagraph('D');

    expect(table.toPlainText()).toBe('A\tB\nC\tD');
  });

  it('supports custom column separator (CSV)', () => {
    const table = new Table(2, 2);
    table.getCell(0, 0)!.createParagraph('Name');
    table.getCell(0, 1)!.createParagraph('Age');
    table.getCell(1, 0)!.createParagraph('Alice');
    table.getCell(1, 1)!.createParagraph('30');

    expect(table.toPlainText(',')).toBe('Name,Age\nAlice,30');
  });

  it('supports custom row separator', () => {
    const table = new Table(2, 2);
    table.getCell(0, 0)!.createParagraph('A');
    table.getCell(0, 1)!.createParagraph('B');
    table.getCell(1, 0)!.createParagraph('C');
    table.getCell(1, 1)!.createParagraph('D');

    expect(table.toPlainText(' | ', ' // ')).toBe('A | B // C | D');
  });

  it('returns empty string for empty table', () => {
    const table = new Table(0, 0);
    expect(table.toPlainText()).toBe('');
  });

  it('handles single cell', () => {
    const table = new Table(1, 1);
    table.getCell(0, 0)!.createParagraph('Solo');

    expect(table.toPlainText()).toBe('Solo');
  });
});

describe('Table.fromArray()', () => {
  it('creates a table from a 2D array', () => {
    const table = Table.fromArray([
      ['Name', 'Age', 'City'],
      ['Alice', '30', 'New York'],
      ['Bob', '25', 'London'],
    ]);

    expect(table.getRowCount()).toBe(3);
    expect(table.getColumnCount()).toBe(3);
    expect(table.getCell(0, 0)!.getText()).toBe('Name');
    expect(table.getCell(1, 2)!.getText()).toBe('New York');
    expect(table.getCell(2, 1)!.getText()).toBe('25');
  });

  it('round-trips through toArray()', () => {
    const original = [
      ['A', 'B', 'C'],
      ['D', 'E', 'F'],
      ['G', 'H', 'I'],
    ];

    const table = Table.fromArray(original);
    const result = table.toArray();

    expect(result).toEqual(original);
  });

  it('handles empty data', () => {
    const table = Table.fromArray([]);
    expect(table.getRowCount()).toBe(0);
  });

  it('handles single cell', () => {
    const table = Table.fromArray([['Solo']]);
    expect(table.getRowCount()).toBe(1);
    expect(table.getCell(0, 0)!.getText()).toBe('Solo');
  });

  it('pads jagged rows to rectangular grid', () => {
    const table = Table.fromArray([['A', 'B', 'C'], ['D'], ['E', 'F']]);

    expect(table.getRowCount()).toBe(3);
    // All rows should have 3 cells
    expect(table.getRow(0)!.getCellCount()).toBe(3);
    expect(table.getRow(1)!.getCellCount()).toBe(3);
    expect(table.getRow(2)!.getCellCount()).toBe(3);

    expect(table.getCell(1, 1)!.getText()).toBe('');
    expect(table.getCell(1, 2)!.getText()).toBe('');
    expect(table.getCell(2, 2)!.getText()).toBe('');
  });

  it('handles empty strings in data', () => {
    const table = Table.fromArray([
      ['', 'B'],
      ['C', ''],
    ]);

    expect(table.getCell(0, 0)!.getText()).toBe('');
    expect(table.getCell(0, 1)!.getText()).toBe('B');
    expect(table.getCell(1, 0)!.getText()).toBe('C');
    expect(table.getCell(1, 1)!.getText()).toBe('');
  });

  it('applies optional formatting', () => {
    const table = Table.fromArray([['A', 'B']], { alignment: 'center', layout: 'fixed' });

    expect(table.getAlignment()).toBe('center');
    expect(table.getLayout()).toBe('fixed');
  });

  it('generates valid XML', () => {
    const table = Table.fromArray([
      ['Header 1', 'Header 2'],
      ['Data 1', 'Data 2'],
    ]);

    const xml = table.toXML();
    expect(xml.name).toBe('w:tbl');
    expect(xml.children).toBeDefined();
    // Should have tblPr + 2 rows
    const rows = xml.children!.filter((c) => typeof c !== 'string' && c.name === 'w:tr');
    expect(rows).toHaveLength(2);
  });

  describe('fromArray + toArray round-trip', () => {
    it('preserves data through create → extract cycle', () => {
      const datasets = [
        [['Single']],
        [
          ['A', 'B'],
          ['C', 'D'],
        ],
        [
          ['Row1Col1', 'Row1Col2', 'Row1Col3'],
          ['Row2Col1', 'Row2Col2', 'Row2Col3'],
        ],
        [['Unicode: \u00E9\u00E8\u00EA', '\u4E16\u754C']],
      ];

      for (const data of datasets) {
        const table = Table.fromArray(data);
        expect(table.toArray()).toEqual(data);
      }
    });

    it('toPlainText matches fromArray input when using CSV format', () => {
      const table = Table.fromArray([
        ['Name', 'Score'],
        ['Alice', '95'],
        ['Bob', '87'],
      ]);

      expect(table.toPlainText(',')).toBe('Name,Score\nAlice,95\nBob,87');
    });
  });
});
