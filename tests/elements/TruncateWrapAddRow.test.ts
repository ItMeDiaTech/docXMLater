/**
 * Tests for Paragraph.truncate(), Paragraph.wrap(), Table.addRowFromArray()
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';

// ============================================================================
// Paragraph.truncate()
// ============================================================================

describe('Paragraph.truncate()', () => {
  it('truncates long text with default ellipsis', () => {
    const para = new Paragraph().addText('The quick brown fox jumps over the lazy dog');
    para.truncate(20);

    expect(para.getText()).toBe('The quick brown f...');
    expect(para.getText().length).toBe(20);
  });

  it('does nothing when text fits within limit', () => {
    const para = new Paragraph().addText('Short');
    para.truncate(100);

    expect(para.getText()).toBe('Short');
  });

  it('does nothing when text equals limit', () => {
    const para = new Paragraph().addText('Exact');
    para.truncate(5);

    expect(para.getText()).toBe('Exact');
  });

  it('uses custom suffix', () => {
    const para = new Paragraph().addText('Hello World');
    para.truncate(10, ' [more]');

    expect(para.getText()).toBe('Hel [more]');
    expect(para.getText().length).toBe(10);
  });

  it('uses empty suffix', () => {
    const para = new Paragraph().addText('Hello World');
    para.truncate(5, '');

    expect(para.getText()).toBe('Hello');
  });

  it('handles maxLength shorter than suffix', () => {
    const para = new Paragraph().addText('Hello World');
    para.truncate(2, '...');

    // cutoff = max(0, 2-3) = 0, so all text deleted, then "..." appended
    expect(para.getText()).toBe('...');
  });

  it('works with multi-run paragraphs', () => {
    const para = new Paragraph();
    para.addRun(new Run('Hello '));
    para.addRun(new Run('Beautiful '));
    para.addRun(new Run('World'));

    para.truncate(12);

    expect(para.getText()).toBe('Hello Bea...');
    expect(para.getText().length).toBe(12);
  });

  it('returns this for chaining', () => {
    const para = new Paragraph().addText('Long text here');
    expect(para.truncate(10)).toBe(para);
  });

  it('handles empty paragraph', () => {
    const para = new Paragraph();
    para.truncate(10);

    expect(para.getText()).toBe('');
  });

  it('practical: create preview text', () => {
    const para = new Paragraph().addText(
      'This document describes the architecture and implementation details of the system.'
    );
    para.truncate(40);

    expect(para.getText().length).toBe(40);
    expect(para.getText().endsWith('...')).toBe(true);
  });
});

// ============================================================================
// Paragraph.wrap()
// ============================================================================

describe('Paragraph.wrap()', () => {
  it('wraps content with prefix and suffix', () => {
    const para = new Paragraph().addText('content');
    para.wrap('[', ']');

    expect(para.getText()).toBe('[content]');
  });

  it('adds only prefix when suffix is empty', () => {
    const para = new Paragraph().addText('World');
    para.wrap('Hello ', '');

    expect(para.getText()).toBe('Hello World');
  });

  it('adds only suffix when prefix is empty', () => {
    const para = new Paragraph().addText('Hello');
    para.wrap('', '!');

    expect(para.getText()).toBe('Hello!');
  });

  it('preserves existing run formatting', () => {
    const para = new Paragraph();
    para.addRun(new Run('bold', { bold: true }));
    para.wrap('(', ')');

    const runs = para.getRuns();
    expect(runs).toHaveLength(3);
    expect(runs[0]!.getText()).toBe('(');
    expect(runs[1]!.getText()).toBe('bold');
    expect(runs[1]!.getFormatting().bold).toBe(true);
    expect(runs[2]!.getText()).toBe(')');
  });

  it('applies formatting to wrapper runs', () => {
    const para = new Paragraph().addText('text');
    para.wrap('>>>', '<<<', { bold: true, color: 'FF0000' });

    const runs = para.getRuns();
    expect(runs[0]!.getFormatting().bold).toBe(true);
    expect(runs[0]!.getFormatting().color).toBe('FF0000');
    expect(runs[2]!.getFormatting().bold).toBe(true);
  });

  it('returns this for chaining', () => {
    const para = new Paragraph().addText('x');
    expect(para.wrap('a', 'b')).toBe(para);
  });

  it('works on empty paragraph', () => {
    const para = new Paragraph();
    para.wrap('[', ']');

    expect(para.getText()).toBe('[]');
  });

  it('can wrap multiple times', () => {
    const para = new Paragraph().addText('core');
    para.wrap('{', '}');
    para.wrap('[', ']');

    expect(para.getText()).toBe('[{core}]');
  });

  it('practical: add quotation marks around text', () => {
    const para = new Paragraph().addText('To be or not to be');
    para.wrap('\u201C', '\u201D'); // Smart quotes

    expect(para.getText()).toBe('\u201CTo be or not to be\u201D');
  });
});

// ============================================================================
// Table.addRowFromArray()
// ============================================================================

describe('Table.addRowFromArray()', () => {
  it('adds a row from string array', () => {
    const table = Table.fromArray([['Name', 'Age']]);
    table.addRowFromArray(['Alice', '30']);

    expect(table.getRowCount()).toBe(2);
    expect(table.getCell(1, 0)!.getText()).toBe('Alice');
    expect(table.getCell(1, 1)!.getText()).toBe('30');
  });

  it('returns the created row', () => {
    const table = new Table(0, 0);
    const row = table.addRowFromArray(['A', 'B']);

    expect(row.getCellCount()).toBe(2);
    expect(row.getCell(0)!.getText()).toBe('A');
  });

  it('appends multiple rows incrementally', () => {
    const table = Table.fromArray([['Header']]);
    table.addRowFromArray(['Row 1']);
    table.addRowFromArray(['Row 2']);
    table.addRowFromArray(['Row 3']);

    expect(table.getRowCount()).toBe(4);
    expect(table.toArray()).toEqual([['Header'], ['Row 1'], ['Row 2'], ['Row 3']]);
  });

  it('handles empty strings', () => {
    const table = new Table(0, 0);
    table.addRowFromArray(['', 'B', '']);

    expect(table.getCell(0, 0)!.getText()).toBe('');
    expect(table.getCell(0, 1)!.getText()).toBe('B');
    expect(table.getCell(0, 2)!.getText()).toBe('');
  });

  it('handles empty array', () => {
    const table = new Table(0, 0);
    const row = table.addRowFromArray([]);

    expect(row.getCellCount()).toBe(0);
    expect(table.getRowCount()).toBe(1);
  });

  it('builds a table from header + data rows', () => {
    const table = new Table(0, 0);
    table.addRowFromArray(['Product', 'Price', 'Stock']);
    table.addRowFromArray(['Widget', '10.00', '100']);
    table.addRowFromArray(['Gadget', '25.00', '50']);
    table.addRowFromArray(['Doohickey', '5.00', '200']);

    expect(table.getRowCount()).toBe(4);
    expect(table.toArray()).toEqual([
      ['Product', 'Price', 'Stock'],
      ['Widget', '10.00', '100'],
      ['Gadget', '25.00', '50'],
      ['Doohickey', '5.00', '200'],
    ]);
  });

  it('row is part of the table (parent set)', () => {
    const table = new Table(0, 0);
    const row = table.addRowFromArray(['A']);

    // The row should be retrievable via getRow
    expect(table.getRow(0)).toBe(row);
  });

  it('works with mapColumn after adding rows', () => {
    const table = new Table(0, 0);
    table.addRowFromArray(['Name', 'Score']);
    table.addRowFromArray(['alice', '95']);
    table.addRowFromArray(['bob', '87']);

    table.mapColumn(0, (t, r) => (r === 0 ? t : t.toUpperCase()));

    expect(table.getCell(1, 0)!.getText()).toBe('ALICE');
    expect(table.getCell(2, 0)!.getText()).toBe('BOB');
  });
});

// ============================================================================
// Integration
// ============================================================================

describe('combined usage', () => {
  it('build table incrementally then truncate long cells', () => {
    const table = new Table(0, 0);
    table.addRowFromArray(['Title', 'Description']);
    table.addRowFromArray([
      'Report',
      'This is a very long description that should be truncated for the summary view',
    ]);
    table.addRowFromArray([
      'Analysis',
      'Another lengthy description of the analysis methodology and results',
    ]);

    // Truncate descriptions
    table.mapColumn(1, (text, row) => {
      if (row === 0) return text;
      if (text.length > 30) return text.slice(0, 27) + '...';
      return text;
    });

    expect(table.getCell(1, 1)!.getText().length).toBeLessThanOrEqual(30);
    expect(table.getCell(2, 1)!.getText().endsWith('...')).toBe(true);
  });

  it('wrap paragraph content then truncate', () => {
    const para = new Paragraph().addText('Long content here');
    para.wrap('"', '"');
    para.truncate(15);

    expect(para.getText().length).toBe(15);
    expect(para.getText().startsWith('"')).toBe(true);
  });
});
