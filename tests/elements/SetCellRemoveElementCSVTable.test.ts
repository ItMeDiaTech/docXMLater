/**
 * Tests for Table.setCell(), Document.removeElement(), Document.createTableFromCSV()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';

// ============================================================================
// Table.setCell()
// ============================================================================

describe('Table.setCell()', () => {
  it('sets text in an empty cell', () => {
    const table = new Table(2, 2);
    table.setCell(0, 0, 'Hello');

    expect(table.getCell(0, 0)!.getText()).toBe('Hello');
  });

  it('replaces existing cell text', () => {
    const table = Table.fromArray([['Old']]);
    table.setCell(0, 0, 'New');

    expect(table.getCell(0, 0)!.getText()).toBe('New');
  });

  it('returns this for chaining', () => {
    const table = new Table(2, 2);
    const result = table
      .setCell(0, 0, 'A')
      .setCell(0, 1, 'B')
      .setCell(1, 0, 'C')
      .setCell(1, 1, 'D');

    expect(result).toBe(table);
    expect(table.toArray()).toEqual([
      ['A', 'B'],
      ['C', 'D'],
    ]);
  });

  it('removes extra paragraphs (multi-paragraph cell becomes single)', () => {
    const table = new Table(1, 1);
    const cell = table.getCell(0, 0)!;
    cell.createParagraph('Line 1');
    cell.createParagraph('Line 2');
    cell.createParagraph('Line 3');

    table.setCell(0, 0, 'Single line');

    expect(cell.getParagraphs()).toHaveLength(1);
    expect(cell.getText()).toBe('Single line');
  });

  it('is a no-op for out-of-range indices', () => {
    const table = new Table(1, 1);
    table.setCell(0, 0, 'Valid');
    table.setCell(99, 99, 'Invalid');

    expect(table.getCell(0, 0)!.getText()).toBe('Valid');
  });

  it('can set empty string', () => {
    const table = Table.fromArray([['Content']]);
    table.setCell(0, 0, '');

    expect(table.getCell(0, 0)!.getText()).toBe('');
  });

  it('practical: populate a table by coordinates', () => {
    const table = new Table(3, 3);
    const data = [
      ['Name', 'Department', 'Salary'],
      ['Alice', 'Engineering', '$120K'],
      ['Bob', 'Marketing', '$95K'],
    ];

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r]!.length; c++) {
        table.setCell(r, c, data[r]![c]!);
      }
    }

    expect(table.toArray()).toEqual(data);
  });

  it('works with mapColumn after setCell', () => {
    const table = new Table(2, 2);
    table.setCell(0, 0, 'Name').setCell(0, 1, 'Score');
    table.setCell(1, 0, 'alice').setCell(1, 1, '95');

    table.mapColumn(0, (t, r) => (r === 0 ? t : t.toUpperCase()));

    expect(table.getCell(1, 0)!.getText()).toBe('ALICE');
  });
});

// ============================================================================
// Document.removeElement()
// ============================================================================

describe('Document.removeElement()', () => {
  it('removes a paragraph by reference', () => {
    const doc = Document.create();
    const p1 = doc.createParagraph('Keep');
    const p2 = doc.createParagraph('Remove');
    const p3 = doc.createParagraph('Keep too');

    const result = doc.removeElement(p2);

    expect(result).toBe(true);
    const texts = doc.getParagraphs().map((p) => p.getText());
    expect(texts).toEqual(['Keep', 'Keep too']);
    doc.dispose();
  });

  it('removes a table by reference', () => {
    const doc = Document.create();
    doc.createParagraph('Before');
    const table = doc.createTable(1, 1);
    doc.createParagraph('After');

    doc.removeElement(table);

    expect(doc.getTables()).toHaveLength(0);
    expect(doc.getParagraphs().map((p) => p.getText())).toEqual(['Before', 'After']);
    doc.dispose();
  });

  it('returns false when element not found', () => {
    const doc = Document.create();
    doc.createParagraph('Existing');

    const orphan = new Paragraph().addText('Not in doc');
    expect(doc.removeElement(orphan)).toBe(false);
    doc.dispose();
  });

  it('can remove multiple elements in a loop', () => {
    const doc = Document.create();
    doc.createParagraph('Text');
    doc.addTable(Table.fromArray([['T1']]));
    doc.createParagraph('More text');
    doc.addTable(Table.fromArray([['T2']]));

    // Remove all tables
    const tables = [...doc.getTables()];
    for (const table of tables) {
      doc.removeElement(table);
    }

    expect(doc.getTables()).toHaveLength(0);
    expect(doc.getParagraphs()).toHaveLength(2);
    doc.dispose();
  });

  it('preserves element order after removal', () => {
    const doc = Document.create();
    const p1 = doc.createParagraph('A');
    const p2 = doc.createParagraph('B');
    const p3 = doc.createParagraph('C');
    const p4 = doc.createParagraph('D');

    doc.removeElement(p2);
    doc.removeElement(p4);

    const elements = doc.getBodyElements();
    expect(elements).toEqual([p1, p3]);
    doc.dispose();
  });

  it('practical: remove paragraphs matching a condition', () => {
    const doc = Document.create();
    doc.createParagraph('Keep this');
    doc.createParagraph('DRAFT: Remove this');
    doc.createParagraph('Also keep');
    doc.createParagraph('DRAFT: And this too');

    const drafts = doc.getParagraphs().filter((p) => p.getText().startsWith('DRAFT:'));
    for (const draft of drafts) {
      doc.removeElement(draft);
    }

    expect(doc.getParagraphs()).toHaveLength(2);
    expect(doc.getParagraphs().every((p) => !p.getText().startsWith('DRAFT:'))).toBe(true);
    doc.dispose();
  });
});

// ============================================================================
// Document.createTableFromCSV()
// ============================================================================

describe('Document.createTableFromCSV()', () => {
  it('creates and adds a table from CSV', () => {
    const doc = Document.create();
    const table = doc.createTableFromCSV('Name,Age\nAlice,30\nBob,25');

    expect(doc.getTables()).toHaveLength(1);
    expect(table.toArray()).toEqual([
      ['Name', 'Age'],
      ['Alice', '30'],
      ['Bob', '25'],
    ]);
    doc.dispose();
  });

  it('returns the table for further customization', () => {
    const doc = Document.create();
    const table = doc.createTableFromCSV('A,B\n1,2');

    table.setAlignment('center');
    expect(table.getAlignment()).toBe('center');
    doc.dispose();
  });

  it('handles quoted CSV fields', () => {
    const doc = Document.create();
    const table = doc.createTableFromCSV('"City, State",Pop\n"New York, NY",8336817');

    expect(table.getCell(1, 0)!.getText()).toBe('New York, NY');
    doc.dispose();
  });

  it('supports TSV delimiter', () => {
    const doc = Document.create();
    const table = doc.createTableFromCSV('Name\tAge\nAlice\t30', '\t');

    expect(table.toArray()).toEqual([
      ['Name', 'Age'],
      ['Alice', '30'],
    ]);
    doc.dispose();
  });

  it('places table in document body at current position', () => {
    const doc = Document.create();
    doc.addHeading('Data', 1);
    doc.createTableFromCSV('A,B\n1,2');
    doc.createParagraph('After table');

    const elements = doc.getBodyElements();
    expect(elements[0] instanceof Paragraph).toBe(true);
    expect(elements[1] instanceof Table).toBe(true);
    expect(elements[2] instanceof Paragraph).toBe(true);
    doc.dispose();
  });

  it('generates valid DOCX', async () => {
    const doc = Document.create();
    doc.addHeading('Sales Report', 1);
    doc.createTableFromCSV('Product,Revenue\nWidget,"$1,200"\nGadget,"$3,500"');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.getTables()).toHaveLength(1);
    loaded.dispose();
    doc.dispose();
  });
});

// ============================================================================
// Combined usage
// ============================================================================

describe('combined patterns', () => {
  it('import CSV, set cells, remove element workflow', () => {
    const doc = Document.create();
    doc.addHeading('Report', 1);

    // Import CSV data
    const table = doc.createTableFromCSV('Name,Status\nAlice,Active\nBob,Inactive');

    // Update a cell
    table.setCell(2, 1, 'Active');

    // Verify
    expect(table.getCell(2, 1)!.getText()).toBe('Active');

    // Remove a paragraph by reference
    const heading = doc.getParagraphs().find((p) => p.getStyle() === 'Heading1');
    expect(heading).toBeDefined();

    // All elements present
    expect(doc.getBodyElements()).toHaveLength(2);
    doc.dispose();
  });

  it('setCell + toCSV round-trip', () => {
    const table = new Table(2, 3);
    table
      .setCell(0, 0, 'A')
      .setCell(0, 1, 'B')
      .setCell(0, 2, 'C')
      .setCell(1, 0, '1')
      .setCell(1, 1, '2')
      .setCell(1, 2, '3');

    const csv = table.toCSV();
    const restored = Table.fromCSV(csv);

    expect(restored.toArray()).toEqual([
      ['A', 'B', 'C'],
      ['1', '2', '3'],
    ]);
  });
});
