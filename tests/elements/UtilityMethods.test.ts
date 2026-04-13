/**
 * Tests for Paragraph.contains(), TableRow.getText()/toArray(),
 * Table.getColumnTexts(), Document.clear()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';

// ============================================================================
// Paragraph.contains()
// ============================================================================

describe('Paragraph.contains()', () => {
  it('returns true when text is found', () => {
    const para = new Paragraph().addText('Hello World');
    expect(para.contains('World')).toBe(true);
  });

  it('returns false when text is not found', () => {
    const para = new Paragraph().addText('Hello World');
    expect(para.contains('xyz')).toBe(false);
  });

  it('is case-insensitive by default', () => {
    const para = new Paragraph().addText('Hello World');
    expect(para.contains('hello')).toBe(true);
    expect(para.contains('WORLD')).toBe(true);
  });

  it('supports case-sensitive mode', () => {
    const para = new Paragraph().addText('Hello World');
    expect(para.contains('Hello', true)).toBe(true);
    expect(para.contains('hello', true)).toBe(false);
  });

  it('handles empty paragraph', () => {
    const para = new Paragraph();
    expect(para.contains('test')).toBe(false);
  });

  it('handles empty search string', () => {
    const para = new Paragraph().addText('Hello');
    expect(para.contains('')).toBe(true);
  });

  it('finds text across multiple runs', () => {
    const para = new Paragraph();
    para.addText('Hello ');
    para.addText('World');
    // getText() returns "Hello World"
    expect(para.contains('lo Wo')).toBe(true);
  });

  it('practical: filter paragraphs by content', () => {
    const doc = Document.create();
    doc.createParagraph('TODO: fix this bug');
    doc.createParagraph('Normal paragraph');
    doc.createParagraph('TODO: add tests');
    doc.createParagraph('Another normal one');

    const todos = doc.getParagraphs().filter((p) => p.contains('TODO'));
    expect(todos).toHaveLength(2);
    doc.dispose();
  });
});

// ============================================================================
// TableRow.toArray() and TableRow.getText()
// ============================================================================

describe('TableRow.toArray()', () => {
  it('returns cell texts as string array', () => {
    const row = new TableRow(3);
    row.getCell(0)!.createParagraph('A');
    row.getCell(1)!.createParagraph('B');
    row.getCell(2)!.createParagraph('C');

    expect(row.toArray()).toEqual(['A', 'B', 'C']);
  });

  it('returns empty strings for empty cells', () => {
    const row = new TableRow(2);
    expect(row.toArray()).toEqual(['', '']);
  });

  it('returns empty array for row with no cells', () => {
    const row = new TableRow();
    expect(row.toArray()).toEqual([]);
  });

  it('works with Table.getRow()', () => {
    const table = Table.fromArray([
      ['Name', 'Age'],
      ['Alice', '30'],
    ]);

    expect(table.getRow(0)!.toArray()).toEqual(['Name', 'Age']);
    expect(table.getRow(1)!.toArray()).toEqual(['Alice', '30']);
  });
});

describe('TableRow.getText()', () => {
  it('returns tab-separated text by default', () => {
    const row = new TableRow(3);
    row.getCell(0)!.createParagraph('A');
    row.getCell(1)!.createParagraph('B');
    row.getCell(2)!.createParagraph('C');

    expect(row.getText()).toBe('A\tB\tC');
  });

  it('supports custom separator', () => {
    const row = new TableRow(2);
    row.getCell(0)!.createParagraph('Hello');
    row.getCell(1)!.createParagraph('World');

    expect(row.getText(', ')).toBe('Hello, World');
    expect(row.getText(' | ')).toBe('Hello | World');
  });

  it('handles single cell', () => {
    const row = new TableRow(1);
    row.getCell(0)!.createParagraph('Solo');

    expect(row.getText()).toBe('Solo');
  });

  it('handles empty row', () => {
    const row = new TableRow();
    expect(row.getText()).toBe('');
  });
});

// ============================================================================
// Table.getColumnTexts()
// ============================================================================

describe('Table.getColumnTexts()', () => {
  it('returns text values for a column', () => {
    const table = Table.fromArray([
      ['Name', 'Age'],
      ['Alice', '30'],
      ['Bob', '25'],
    ]);

    expect(table.getColumnTexts(0)).toEqual(['Name', 'Alice', 'Bob']);
    expect(table.getColumnTexts(1)).toEqual(['Age', '30', '25']);
  });

  it('returns empty array for out-of-range column', () => {
    const table = Table.fromArray([['A']]);
    expect(table.getColumnTexts(99)).toEqual([]);
  });

  it('returns empty array for empty table', () => {
    const table = new Table(0, 0);
    expect(table.getColumnTexts(0)).toEqual([]);
  });

  it('practical: sum a numeric column', () => {
    const table = Table.fromArray([
      ['Item', 'Price'],
      ['Widget', '10'],
      ['Gadget', '25'],
      ['Doohickey', '5'],
    ]);

    const prices = table.getColumnTexts(1).slice(1); // Skip header
    const total = prices.reduce((sum, v) => sum + Number(v), 0);

    expect(total).toBe(40);
  });

  it('practical: get unique values in a column', () => {
    const table = Table.fromArray([
      ['Name', 'Dept'],
      ['Alice', 'Eng'],
      ['Bob', 'Sales'],
      ['Charlie', 'Eng'],
      ['Diana', 'Sales'],
    ]);

    const depts = [...new Set(table.getColumnTexts(1).slice(1))];
    expect(depts).toEqual(['Eng', 'Sales']);
  });
});

// ============================================================================
// Document.clear()
// ============================================================================

describe('Document.clear()', () => {
  it('removes all body elements', () => {
    const doc = Document.create();
    doc.createParagraph('Para 1');
    doc.createParagraph('Para 2');
    doc.addTable(Table.fromArray([['A']]));

    doc.clear();

    expect(doc.getParagraphs()).toHaveLength(0);
    expect(doc.getTables()).toHaveLength(0);
    expect(doc.getBodyElements()).toHaveLength(0);
    doc.dispose();
  });

  it('returns this for chaining', () => {
    const doc = Document.create();
    doc.createParagraph('Content');

    const result = doc.clear();
    expect(result).toBe(doc);
    doc.dispose();
  });

  it('allows adding new content after clearing', () => {
    const doc = Document.create();
    doc.createParagraph('Old content');

    doc.clear();
    doc.addHeading('New Title', 1);
    doc.createParagraph('New content');

    expect(doc.getParagraphs()).toHaveLength(2);
    expect(doc.toPlainText()).toContain('New Title');
    expect(doc.toPlainText()).not.toContain('Old content');
    doc.dispose();
  });

  it('preserves styles after clearing', () => {
    const doc = Document.create();
    doc.setDefaultFont('Georgia', 12);
    doc.createParagraph('Content');

    doc.clear();

    // Styles should still be accessible
    const normal = doc.getStylesManager().getStyle('Normal');
    expect(normal?.getRunFormatting()?.font).toBe('Georgia');
    doc.dispose();
  });

  it('handles already-empty document', () => {
    const doc = Document.create();
    doc.clear();

    expect(doc.getBodyElements()).toHaveLength(0);
    doc.dispose();
  });

  it('produces valid DOCX after clear + rebuild', async () => {
    const doc = Document.create();
    doc.createParagraph('Original');

    doc.clear();
    doc.addHeading('Rebuilt', 1);
    doc.createParagraph('Fresh content.');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.toPlainText()).toContain('Rebuilt');
    expect(loaded.toPlainText()).not.toContain('Original');
    loaded.dispose();
    doc.dispose();
  });

  it('practical: template reset pattern', () => {
    const doc = Document.create();
    doc.setDefaultFont('Arial', 11);
    doc.getStylesManager().cloneStyle('Heading1', 'CustomH1');

    // First fill
    doc.addHeading('Report v1', 1);
    doc.createParagraph('Version 1 data.');
    expect(doc.toPlainText()).toContain('Report v1');

    // Reset and refill
    doc.clear();
    doc.addHeading('Report v2', 1);
    doc.createParagraph('Version 2 data.');
    expect(doc.toPlainText()).toContain('Report v2');
    expect(doc.toPlainText()).not.toContain('Report v1');

    // Custom style still available
    expect(doc.getStylesManager().hasStyle('CustomH1')).toBe(true);
    doc.dispose();
  });
});

// ============================================================================
// Combined usage
// ============================================================================

describe('combined utility methods', () => {
  it('filter rows using contains + row.toArray', () => {
    const doc = Document.create();
    const table = doc.createTableFromCSV('Name,Status\nAlice,active\nBob,inactive\nCharlie,active');

    const activeRows = table
      .getRows()
      .filter((row) => row.toArray()[1] === 'active')
      .map((row) => row.toArray());

    expect(activeRows).toEqual([
      ['Alice', 'active'],
      ['Charlie', 'active'],
    ]);
    doc.dispose();
  });

  it('getColumnTexts + contains for data analysis', () => {
    const table = Table.fromArray([
      ['Product', 'Category'],
      ['Widget', 'Hardware'],
      ['Software Pro', 'Software'],
      ['Cable', 'Hardware'],
    ]);

    const hardware = table
      .getRows()
      .slice(1)
      .filter((_, i) => table.getColumnTexts(1)[i + 1] === 'Hardware')
      .map((row) => row.toArray()[0]);

    expect(hardware).toEqual(['Widget', 'Cable']);
  });
});
