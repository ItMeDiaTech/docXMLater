/**
 * Tests for Document.insertAfter(), Document.insertBefore(), Document.replaceElement()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';

describe('Document.insertAfter()', () => {
  it('inserts a paragraph after another paragraph', () => {
    const doc = Document.create();
    const first = doc.createParagraph('First');
    doc.createParagraph('Third');

    const second = new Paragraph().addText('Second');
    const result = doc.insertAfter(first, second);

    expect(result).toBe(true);
    const texts = doc.getParagraphs().map((p) => p.getText());
    expect(texts).toEqual(['First', 'Second', 'Third']);
    doc.dispose();
  });

  it('inserts a table after a paragraph', () => {
    const doc = Document.create();
    const para = doc.createParagraph('Before table');
    doc.createParagraph('After table');

    const table = Table.fromArray([['A', 'B']]);
    doc.insertAfter(para, table);

    const elements = doc.getBodyElements();
    expect(elements[0]).toBe(para);
    expect(elements[1]).toBe(table);
    doc.dispose();
  });

  it('inserts after the last element', () => {
    const doc = Document.create();
    const last = doc.createParagraph('Last');

    const appended = new Paragraph().addText('Appended');
    doc.insertAfter(last, appended);

    const paras = doc.getParagraphs();
    expect(paras[paras.length - 1]!.getText()).toBe('Appended');
    doc.dispose();
  });

  it('returns false when reference not found', () => {
    const doc = Document.create();
    doc.createParagraph('Existing');

    const orphan = new Paragraph().addText('Not in doc');
    const newPara = new Paragraph().addText('New');

    expect(doc.insertAfter(orphan, newPara)).toBe(false);
    expect(doc.getParagraphs()).toHaveLength(1);
    doc.dispose();
  });

  it('works with splitAt pattern', () => {
    const doc = Document.create();
    const para = doc.createParagraph('Hello World');

    const tail = para.splitAt(5);
    const table = Table.fromArray([['Data']]);

    doc.insertAfter(para, table);
    doc.insertAfter(table, tail);

    const elements = doc.getBodyElements();
    expect(elements[0]).toBe(para);
    expect(elements[1]).toBe(table);
    expect(elements[2]).toBe(tail);

    expect(para.getText()).toBe('Hello');
    expect(tail.getText()).toBe(' World');
    doc.dispose();
  });
});

describe('Document.insertBefore()', () => {
  it('inserts a paragraph before another paragraph', () => {
    const doc = Document.create();
    doc.createParagraph('First');
    const third = doc.createParagraph('Third');

    const second = new Paragraph().addText('Second');
    const result = doc.insertBefore(third, second);

    expect(result).toBe(true);
    const texts = doc.getParagraphs().map((p) => p.getText());
    expect(texts).toEqual(['First', 'Second', 'Third']);
    doc.dispose();
  });

  it('inserts before the first element', () => {
    const doc = Document.create();
    const first = doc.createParagraph('Was first');

    const newFirst = new Paragraph().addText('Now first');
    doc.insertBefore(first, newFirst);

    expect(doc.getParagraphs()[0]!.getText()).toBe('Now first');
    doc.dispose();
  });

  it('inserts a heading before a table', () => {
    const doc = Document.create();
    const table = doc.createTable(1, 2);
    table.getCell(0, 0)!.createParagraph('Data');

    const heading = new Paragraph().addText('Table 1');
    heading.setStyle('Heading2');
    doc.insertBefore(table, heading);

    const elements = doc.getBodyElements();
    expect(elements[0]).toBe(heading);
    expect(elements[1]).toBe(table);
    doc.dispose();
  });

  it('returns false when reference not found', () => {
    const doc = Document.create();
    doc.createParagraph('Existing');

    const orphan = new Paragraph().addText('Not in doc');
    const newPara = new Paragraph().addText('New');

    expect(doc.insertBefore(orphan, newPara)).toBe(false);
    doc.dispose();
  });
});

describe('Document.replaceElement()', () => {
  it('replaces a paragraph with another paragraph', () => {
    const doc = Document.create();
    doc.createParagraph('First');
    const old = doc.createParagraph('Replace me');
    doc.createParagraph('Third');

    const replacement = new Paragraph().addText('Replaced!');
    const result = doc.replaceElement(old, replacement);

    expect(result).toBe(true);
    const texts = doc.getParagraphs().map((p) => p.getText());
    expect(texts).toEqual(['First', 'Replaced!', 'Third']);
    doc.dispose();
  });

  it('replaces a paragraph with a table', () => {
    const doc = Document.create();
    const placeholder = doc.createParagraph('{{TABLE}}');
    doc.createParagraph('After');

    const table = Table.fromArray([
      ['Name', 'Value'],
      ['A', '1'],
    ]);
    doc.replaceElement(placeholder, table);

    const elements = doc.getBodyElements();
    expect(elements[0]).toBe(table);
    expect(elements[0] instanceof Table).toBe(true);
    doc.dispose();
  });

  it('replaces a table with a paragraph', () => {
    const doc = Document.create();
    doc.createParagraph('Before');
    const table = doc.createTable(1, 1);
    doc.createParagraph('After');

    const summary = new Paragraph().addText('Table removed');
    doc.replaceElement(table, summary);

    const texts = doc.getParagraphs().map((p) => p.getText());
    expect(texts).toContain('Table removed');
    expect(doc.getTables()).toHaveLength(0);
    doc.dispose();
  });

  it('returns false when old element not found', () => {
    const doc = Document.create();
    doc.createParagraph('Existing');

    const orphan = new Paragraph().addText('Not in doc');
    const replacement = new Paragraph().addText('New');

    expect(doc.replaceElement(orphan, replacement)).toBe(false);
    doc.dispose();
  });

  it('preserves position in document', () => {
    const doc = Document.create();
    const p1 = doc.createParagraph('One');
    const p2 = doc.createParagraph('Two');
    const p3 = doc.createParagraph('Three');

    const replacement = new Paragraph().addText('TWO');
    doc.replaceElement(p2, replacement);

    const elements = doc.getBodyElements();
    expect(elements[0]).toBe(p1);
    expect(elements[1]).toBe(replacement);
    expect(elements[2]).toBe(p3);
    doc.dispose();
  });
});

describe('combined patterns', () => {
  it('find-and-insert: add heading before each table', () => {
    const doc = Document.create();
    doc.createParagraph('Intro');
    doc.addTable(Table.fromArray([['Table 1 data']]));
    doc.createParagraph('Between');
    doc.addTable(Table.fromArray([['Table 2 data']]));

    const tables = doc.getTables();
    tables.forEach((table, i) => {
      const heading = new Paragraph().addText(`Table ${i + 1}`);
      heading.setStyle('Heading3');
      doc.insertBefore(table, heading);
    });

    const headings = doc.getParagraphs().filter((p) => p.getStyle() === 'Heading3');
    expect(headings).toHaveLength(2);
    expect(headings[0]!.getText()).toBe('Table 1');
    expect(headings[1]!.getText()).toBe('Table 2');
    doc.dispose();
  });

  it('template replacement: replace placeholder paragraphs with tables', () => {
    const doc = Document.create();
    doc.addHeading('Report', 1);
    doc.createParagraph('{{SALES_TABLE}}');
    doc.createParagraph('Some text between.');
    doc.createParagraph('{{COSTS_TABLE}}');

    // Replace placeholders
    const allParas = doc.getParagraphs();
    for (const para of allParas) {
      if (para.getText() === '{{SALES_TABLE}}') {
        doc.replaceElement(
          para,
          Table.fromArray([
            ['Product', 'Sales'],
            ['Widget', '100'],
          ])
        );
      } else if (para.getText() === '{{COSTS_TABLE}}') {
        doc.replaceElement(
          para,
          Table.fromArray([
            ['Category', 'Cost'],
            ['Labor', '50'],
          ])
        );
      }
    }

    expect(doc.getTables()).toHaveLength(2);
    expect(doc.getTables()[0]!.getCell(0, 0)!.getText()).toBe('Product');
    expect(doc.getTables()[1]!.getCell(0, 0)!.getText()).toBe('Category');
    doc.dispose();
  });

  it('generates valid DOCX after operations', async () => {
    const doc = Document.create();
    const para = doc.createParagraph('Hello World');

    const tail = para.splitAt(5);
    const table = Table.fromArray([['Inserted']]);

    doc.insertAfter(para, table);
    doc.insertAfter(table, tail);

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.getTables()).toHaveLength(1);
    loaded.dispose();
    doc.dispose();
  });
});
