/**
 * Tests for Document.getStatistics(), Document.forEachParagraph(), Document.forEachTable()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';

describe('Document.getStatistics()', () => {
  it('returns zero counts for empty document', () => {
    const doc = Document.create();
    const stats = doc.getStatistics();

    expect(stats.words).toBe(0);
    expect(stats.characters).toBe(0);
    expect(stats.charactersNoSpaces).toBe(0);
    expect(stats.paragraphs).toBe(0);
    expect(stats.tables).toBe(0);
    expect(stats.headings).toBe(0);
    expect(stats.lists).toBe(0);
    doc.dispose();
  });

  it('counts words correctly', () => {
    const doc = Document.create();
    doc.createParagraph('The quick brown fox');
    doc.createParagraph('jumps over');

    const stats = doc.getStatistics();
    expect(stats.words).toBe(6);
    doc.dispose();
  });

  it('counts characters with and without spaces', () => {
    const doc = Document.create();
    doc.createParagraph('Hello World');

    const stats = doc.getStatistics();
    expect(stats.characters).toBe(11);
    expect(stats.charactersNoSpaces).toBe(10);
    doc.dispose();
  });

  it('counts paragraphs including those in tables', () => {
    const doc = Document.create();
    doc.createParagraph('Top level');
    const table = doc.createTable(2, 2);
    table.getCell(0, 0)!.createParagraph('Cell text');
    doc.createParagraph('Bottom level');

    const stats = doc.getStatistics();
    // 2 top-level + cell paragraphs (table creates default empty paragraphs + our added one)
    expect(stats.paragraphs).toBeGreaterThanOrEqual(3);
    doc.dispose();
  });

  it('counts tables', () => {
    const doc = Document.create();
    doc.addTable(Table.fromArray([['A']]));
    doc.addTable(Table.fromArray([['B']]));
    doc.addTable(Table.fromArray([['C']]));

    expect(doc.getStatistics().tables).toBe(3);
    doc.dispose();
  });

  it('counts headings', () => {
    const doc = Document.create();
    doc.addHeading('Title', 1);
    doc.addHeading('Section', 2);
    doc.createParagraph('Normal text');
    doc.addHeading('Subsection', 3);

    expect(doc.getStatistics().headings).toBe(3);
    doc.dispose();
  });

  it('counts words in table cells', () => {
    const doc = Document.create();
    const table = Table.fromArray([
      ['Hello World', 'Foo Bar'],
      ['One', 'Two Three'],
    ]);
    doc.addTable(table);

    const stats = doc.getStatistics();
    expect(stats.words).toBe(7); // Hello World Foo Bar One Two Three
    doc.dispose();
  });

  it('returns all metric fields', () => {
    const doc = Document.create();
    doc.createParagraph('Test');

    const stats = doc.getStatistics();

    // Verify all fields exist
    expect(typeof stats.words).toBe('number');
    expect(typeof stats.characters).toBe('number');
    expect(typeof stats.charactersNoSpaces).toBe('number');
    expect(typeof stats.paragraphs).toBe('number');
    expect(typeof stats.tables).toBe('number');
    expect(typeof stats.images).toBe('number');
    expect(typeof stats.headings).toBe('number');
    expect(typeof stats.lists).toBe('number');
    expect(typeof stats.hyperlinks).toBe('number');
    expect(typeof stats.bookmarks).toBe('number');
    expect(typeof stats.footnotes).toBe('number');
    expect(typeof stats.endnotes).toBe('number');
    expect(typeof stats.comments).toBe('number');
    expect(typeof stats.sections).toBe('number');
    doc.dispose();
  });

  it('provides accurate mixed document statistics', () => {
    const doc = Document.create();
    doc.addHeading('Report Title', 1);
    doc.createParagraph('This is the introduction paragraph.');
    doc.addHeading('Data Section', 2);
    doc.addTable(
      Table.fromArray([
        ['Name', 'Value'],
        ['Alpha', '100'],
      ])
    );
    doc.createParagraph('Conclusion text here.');

    const stats = doc.getStatistics();

    expect(stats.headings).toBe(2);
    expect(stats.tables).toBe(1);
    expect(stats.paragraphs).toBeGreaterThanOrEqual(4); // headings + paras + table cells
    expect(stats.words).toBeGreaterThanOrEqual(12);
    doc.dispose();
  });
});

describe('Document.forEachParagraph()', () => {
  it('iterates top-level paragraphs only', () => {
    const doc = Document.create();
    doc.createParagraph('Para 1');
    doc.addTable(Table.fromArray([['Table cell']]));
    doc.createParagraph('Para 2');

    const texts: string[] = [];
    doc.forEachParagraph((para) => {
      texts.push(para.getText());
    });

    expect(texts).toEqual(['Para 1', 'Para 2']);
    doc.dispose();
  });

  it('provides sequential paragraph index', () => {
    const doc = Document.create();
    doc.createParagraph('A');
    doc.addTable(Table.fromArray([['skip']]));
    doc.createParagraph('B');
    doc.createParagraph('C');

    const indices: number[] = [];
    doc.forEachParagraph((_para, index) => {
      indices.push(index);
    });

    expect(indices).toEqual([0, 1, 2]);
    doc.dispose();
  });

  it('returns count of paragraphs visited', () => {
    const doc = Document.create();
    doc.createParagraph('A');
    doc.createParagraph('B');
    doc.createParagraph('C');

    const count = doc.forEachParagraph(() => {});

    expect(count).toBe(3);
    doc.dispose();
  });

  it('supports early termination by returning false', () => {
    const doc = Document.create();
    doc.createParagraph('First');
    doc.createParagraph('Second');
    doc.createParagraph('Third');

    const visited: string[] = [];
    doc.forEachParagraph((para) => {
      visited.push(para.getText());
      if (para.getText() === 'Second') return false;
      return undefined;
    });

    expect(visited).toEqual(['First', 'Second']);
    doc.dispose();
  });

  it('returns 0 for empty document', () => {
    const doc = Document.create();
    const count = doc.forEachParagraph(() => {});

    expect(count).toBe(0);
    doc.dispose();
  });

  it('can modify paragraphs during iteration', () => {
    const doc = Document.create();
    doc.createParagraph('lower case');
    doc.createParagraph('another one');

    doc.forEachParagraph((para) => {
      const upper = para.getText().toUpperCase();
      para.getRuns().forEach((r) => r.setText(upper));
    });

    const texts = doc.getParagraphs().map((p) => p.getText());
    expect(texts).toEqual(['LOWER CASE', 'ANOTHER ONE']);
    doc.dispose();
  });
});

describe('Document.forEachTable()', () => {
  it('iterates top-level tables only', () => {
    const doc = Document.create();
    doc.createParagraph('Intro');
    doc.addTable(Table.fromArray([['Table 1']]));
    doc.createParagraph('Middle');
    doc.addTable(Table.fromArray([['Table 2']]));

    const tableTexts: string[] = [];
    doc.forEachTable((table) => {
      tableTexts.push(table.getCell(0, 0)!.getText());
    });

    expect(tableTexts).toEqual(['Table 1', 'Table 2']);
    doc.dispose();
  });

  it('provides sequential table index', () => {
    const doc = Document.create();
    doc.addTable(Table.fromArray([['A']]));
    doc.createParagraph('gap');
    doc.addTable(Table.fromArray([['B']]));

    const indices: number[] = [];
    doc.forEachTable((_table, index) => {
      indices.push(index);
    });

    expect(indices).toEqual([0, 1]);
    doc.dispose();
  });

  it('returns count of tables visited', () => {
    const doc = Document.create();
    doc.addTable(Table.fromArray([['A']]));
    doc.addTable(Table.fromArray([['B']]));

    expect(doc.forEachTable(() => {})).toBe(2);
    doc.dispose();
  });

  it('supports early termination', () => {
    const doc = Document.create();
    doc.addTable(Table.fromArray([['First']]));
    doc.addTable(Table.fromArray([['Second']]));
    doc.addTable(Table.fromArray([['Third']]));

    let found: Table | undefined;
    doc.forEachTable((table) => {
      if (table.getCell(0, 0)!.getText() === 'Second') {
        found = table;
        return false;
      }
      return undefined;
    });

    expect(found).toBeDefined();
    expect(found!.getCell(0, 0)!.getText()).toBe('Second');
    doc.dispose();
  });

  it('returns 0 for document without tables', () => {
    const doc = Document.create();
    doc.createParagraph('No tables');

    expect(doc.forEachTable(() => {})).toBe(0);
    doc.dispose();
  });

  it('can apply operations to all tables', () => {
    const doc = Document.create();
    doc.addTable(
      Table.fromArray([
        ['A', ''],
        ['', ''],
      ])
    );
    doc.addTable(
      Table.fromArray([
        ['B', ''],
        ['', ''],
      ])
    );

    doc.forEachTable((table) => {
      table.removeEmptyRows();
      table.removeEmptyColumns();
    });

    const tables = doc.getTables();
    expect(tables[0]!.toArray()).toEqual([['A']]);
    expect(tables[1]!.toArray()).toEqual([['B']]);
    doc.dispose();
  });
});
