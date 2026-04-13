/**
 * Tests for Table.toCSV() and Table.fromCSV()
 */

import { Table } from '../../src/elements/Table';

describe('Table.toCSV()', () => {
  it('exports simple table as CSV', () => {
    const table = Table.fromArray([
      ['Name', 'Age'],
      ['Alice', '30'],
      ['Bob', '25'],
    ]);

    expect(table.toCSV()).toBe('Name,Age\nAlice,30\nBob,25');
  });

  it('quotes fields containing commas', () => {
    const table = Table.fromArray([
      ['City', 'Pop'],
      ['New York, NY', '8336817'],
    ]);

    const csv = table.toCSV();
    expect(csv).toBe('City,Pop\n"New York, NY",8336817');
  });

  it('escapes embedded quotes by doubling', () => {
    const table = Table.fromArray([['Quote'], ['She said "hello"']]);

    const csv = table.toCSV();
    expect(csv).toBe('Quote\n"She said ""hello"""');
  });

  it('quotes fields containing newlines', () => {
    const table = new Table(1, 1);
    const cell = table.getCell(0, 0)!;
    cell.createParagraph('Line 1');
    cell.createParagraph('Line 2');

    const csv = table.toCSV();
    expect(csv).toContain('"Line 1\nLine 2"');
  });

  it('does not quote clean fields', () => {
    const table = Table.fromArray([['Simple', 'Clean', 'Data']]);
    expect(table.toCSV()).toBe('Simple,Clean,Data');
  });

  it('handles empty cells', () => {
    const table = Table.fromArray([
      ['A', '', 'C'],
      ['', 'B', ''],
    ]);

    expect(table.toCSV()).toBe('A,,C\n,B,');
  });

  it('handles empty table', () => {
    const table = new Table(0, 0);
    expect(table.toCSV()).toBe('');
  });

  it('supports custom delimiter (TSV)', () => {
    const table = Table.fromArray([
      ['Name', 'Age'],
      ['Alice', '30'],
    ]);

    expect(table.toCSV('\t')).toBe('Name\tAge\nAlice\t30');
  });

  it('quotes fields containing custom delimiter', () => {
    const table = Table.fromArray([['A\tB']]);

    expect(table.toCSV('\t')).toBe('"A\tB"');
  });

  it('single cell table', () => {
    const table = Table.fromArray([['Solo']]);
    expect(table.toCSV()).toBe('Solo');
  });
});

describe('Table.fromCSV()', () => {
  it('parses simple CSV', () => {
    const table = Table.fromCSV('Name,Age\nAlice,30\nBob,25');

    expect(table.getRowCount()).toBe(3);
    expect(table.toArray()).toEqual([
      ['Name', 'Age'],
      ['Alice', '30'],
      ['Bob', '25'],
    ]);
  });

  it('handles quoted fields with commas', () => {
    const table = Table.fromCSV('"City, State",Pop\n"New York, NY",8336817');

    expect(table.getCell(0, 0)!.getText()).toBe('City, State');
    expect(table.getCell(1, 0)!.getText()).toBe('New York, NY');
    expect(table.getCell(1, 1)!.getText()).toBe('8336817');
  });

  it('handles escaped quotes (doubled)', () => {
    const table = Table.fromCSV('Quote\n"She said ""hello"""');

    expect(table.getCell(1, 0)!.getText()).toBe('She said "hello"');
  });

  it('handles newlines within quoted fields', () => {
    const table = Table.fromCSV('Data\n"Line 1\nLine 2"');

    expect(table.getCell(1, 0)!.getText()).toBe('Line 1\nLine 2');
  });

  it('handles CRLF line endings', () => {
    const table = Table.fromCSV('A,B\r\nC,D\r\n');

    expect(table.getRowCount()).toBe(2);
    expect(table.toArray()).toEqual([
      ['A', 'B'],
      ['C', 'D'],
    ]);
  });

  it('handles empty fields', () => {
    const table = Table.fromCSV('A,,C\n,B,');

    expect(table.toArray()).toEqual([
      ['A', '', 'C'],
      ['', 'B', ''],
    ]);
  });

  it('handles empty string', () => {
    const table = Table.fromCSV('');
    expect(table.getRowCount()).toBe(0);
  });

  it('handles single cell', () => {
    const table = Table.fromCSV('Solo');
    expect(table.toArray()).toEqual([['Solo']]);
  });

  it('handles single row', () => {
    const table = Table.fromCSV('A,B,C');
    expect(table.toArray()).toEqual([['A', 'B', 'C']]);
  });

  it('supports custom delimiter (TSV)', () => {
    const table = Table.fromCSV('Name\tAge\nAlice\t30', '\t');

    expect(table.toArray()).toEqual([
      ['Name', 'Age'],
      ['Alice', '30'],
    ]);
  });

  it('applies optional formatting', () => {
    const table = Table.fromCSV('A,B', ',', { alignment: 'center' });
    expect(table.getAlignment()).toBe('center');
  });
});

describe('CSV round-trip', () => {
  it('toCSV → fromCSV preserves simple data', () => {
    const original = Table.fromArray([
      ['Product', 'Price', 'Stock'],
      ['Widget', '10.00', '100'],
      ['Gadget', '25.00', '50'],
    ]);

    const csv = original.toCSV();
    const restored = Table.fromCSV(csv);

    expect(restored.toArray()).toEqual(original.toArray());
  });

  it('round-trips data with commas and quotes', () => {
    const original = Table.fromArray([
      ['Name', 'Address'],
      ['Alice', '123 Main St, Apt 4'],
      ['Bob', 'Said "hi" today'],
    ]);

    const csv = original.toCSV();
    const restored = Table.fromCSV(csv);

    expect(restored.toArray()).toEqual(original.toArray());
  });

  it('fromCSV → toCSV → fromCSV is stable', () => {
    const csv1 = '"Name","City, State"\n"Alice","New York, NY"\n"Bob","Los Angeles, CA"';
    const table1 = Table.fromCSV(csv1);
    const csv2 = table1.toCSV();
    const table2 = Table.fromCSV(csv2);

    expect(table2.toArray()).toEqual(table1.toArray());
  });

  it('practical: import spreadsheet data into DOCX table', () => {
    // Simulating data from a CSV export
    const spreadsheetCSV = [
      'Employee,Department,Salary',
      'Alice Smith,Engineering,"$120,000"',
      'Bob Jones,Marketing,"$95,000"',
      'Charlie Brown,Sales,"$88,500"',
    ].join('\n');

    const table = Table.fromCSV(spreadsheetCSV);

    expect(table.getRowCount()).toBe(4);
    expect(table.getCell(0, 0)!.getText()).toBe('Employee');
    expect(table.getCell(1, 2)!.getText()).toBe('$120,000');
    expect(table.getCell(3, 1)!.getText()).toBe('Sales');

    // Can export back
    const exported = table.toCSV();
    expect(exported).toContain('"$120,000"');
  });
});
