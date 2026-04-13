/**
 * Tests for Document.toMarkdown()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';
import { Hyperlink } from '../../src/elements/Hyperlink';

describe('Document.toMarkdown()', () => {
  describe('headings', () => {
    it('converts heading levels to # syntax', () => {
      const doc = Document.create();
      doc.addHeading('Title', 1);
      doc.addHeading('Section', 2);
      doc.addHeading('Subsection', 3);

      const md = doc.toMarkdown();
      expect(md).toContain('# Title');
      expect(md).toContain('## Section');
      expect(md).toContain('### Subsection');
      doc.dispose();
    });

    it('supports heading levels 1-6', () => {
      const doc = Document.create();
      for (let i = 1; i <= 6; i++) {
        doc.addHeading(`Level ${i}`, i as 1 | 2 | 3 | 4 | 5 | 6);
      }

      const md = doc.toMarkdown();
      expect(md).toContain('# Level 1');
      expect(md).toContain('## Level 2');
      expect(md).toContain('###### Level 6');
      doc.dispose();
    });
  });

  describe('paragraphs', () => {
    it('outputs plain paragraphs as text', () => {
      const doc = Document.create();
      doc.createParagraph('First paragraph.');
      doc.createParagraph('Second paragraph.');

      const md = doc.toMarkdown();
      expect(md).toContain('First paragraph.');
      expect(md).toContain('Second paragraph.');
      doc.dispose();
    });

    it('separates paragraphs with blank lines', () => {
      const doc = Document.create();
      doc.createParagraph('Para 1');
      doc.createParagraph('Para 2');

      const md = doc.toMarkdown();
      expect(md).toBe('Para 1\n\nPara 2');
      doc.dispose();
    });

    it('skips empty paragraphs', () => {
      const doc = Document.create();
      doc.createParagraph('Before');
      doc.createParagraph('');
      doc.createParagraph('After');

      const md = doc.toMarkdown();
      expect(md).toBe('Before\n\nAfter');
      doc.dispose();
    });
  });

  describe('inline formatting', () => {
    it('wraps bold text in **', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('bold text', { bold: true }));

      const md = doc.toMarkdown();
      expect(md).toBe('**bold text**');
      doc.dispose();
    });

    it('wraps italic text in *', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('italic text', { italic: true }));

      const md = doc.toMarkdown();
      expect(md).toBe('*italic text*');
      doc.dispose();
    });

    it('wraps bold+italic in ***', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('bold italic', { bold: true, italic: true }));

      const md = doc.toMarkdown();
      expect(md).toBe('***bold italic***');
      doc.dispose();
    });

    it('wraps strikethrough in ~~', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('deleted', { strike: true }));

      const md = doc.toMarkdown();
      expect(md).toBe('~~deleted~~');
      doc.dispose();
    });

    it('detects monospace fonts as inline code', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('some code', { font: 'Courier New' }));

      const md = doc.toMarkdown();
      expect(md).toBe('`some code`');
      doc.dispose();
    });

    it('handles mixed formatting in a paragraph', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('Normal '));
      para.addRun(new Run('bold', { bold: true }));
      para.addRun(new Run(' and '));
      para.addRun(new Run('italic', { italic: true }));

      const md = doc.toMarkdown();
      expect(md).toBe('Normal **bold** and *italic*');
      doc.dispose();
    });
  });

  describe('hyperlinks', () => {
    it('converts hyperlinks to Markdown link syntax', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('Visit '));
      para.addHyperlink(new Hyperlink({ url: 'https://example.com', text: 'Example' }));

      const md = doc.toMarkdown();
      expect(md).toBe('Visit [Example](https://example.com)');
      doc.dispose();
    });
  });

  describe('tables', () => {
    it('converts a simple table to Markdown', () => {
      const doc = Document.create();
      const table = Table.fromArray([
        ['Name', 'Age'],
        ['Alice', '30'],
        ['Bob', '25'],
      ]);
      doc.addTable(table);

      const md = doc.toMarkdown();
      const lines = md.split('\n');

      expect(lines[0]).toBe('| Name | Age |');
      expect(lines[1]).toBe('| --- | --- |');
      expect(lines[2]).toBe('| Alice | 30 |');
      expect(lines[3]).toBe('| Bob | 25 |');
      doc.dispose();
    });

    it('handles single-row table (header only)', () => {
      const doc = Document.create();
      const table = Table.fromArray([['A', 'B', 'C']]);
      doc.addTable(table);

      const md = doc.toMarkdown();
      expect(md).toContain('| A | B | C |');
      expect(md).toContain('| --- | --- | --- |');
      doc.dispose();
    });

    it('escapes pipe characters in cell text', () => {
      const doc = Document.create();
      const table = Table.fromArray([['Header'], ['A | B']]);
      doc.addTable(table);

      const md = doc.toMarkdown();
      expect(md).toContain('A \\| B');
      doc.dispose();
    });

    it('replaces newlines in cells with spaces', () => {
      const doc = Document.create();
      const table = new Table(1, 1);
      const cell = table.getCell(0, 0)!;
      cell.createParagraph('Line 1');
      cell.createParagraph('Line 2');
      doc.addTable(table);

      const md = doc.toMarkdown();
      expect(md).toContain('Line 1 Line 2');
      doc.dispose();
    });
  });

  describe('mixed content', () => {
    it('converts a full document with headings, paragraphs, and tables', () => {
      const doc = Document.create();

      doc.addHeading('Report Title', 1);
      doc.createParagraph('This is the introduction.');

      doc.addHeading('Data', 2);
      const table = Table.fromArray([
        ['Item', 'Value'],
        ['Revenue', '$1M'],
        ['Costs', '$500K'],
      ]);
      doc.addTable(table);

      doc.addHeading('Conclusion', 2);
      doc.createParagraph('The results are positive.');

      const md = doc.toMarkdown();

      expect(md).toContain('# Report Title');
      expect(md).toContain('This is the introduction.');
      expect(md).toContain('## Data');
      expect(md).toContain('| Item | Value |');
      expect(md).toContain('| Revenue | $1M |');
      expect(md).toContain('## Conclusion');
      expect(md).toContain('The results are positive.');
      doc.dispose();
    });

    it('handles document with page breaks (ignored in markdown)', () => {
      const doc = Document.create();
      doc.createParagraph('Before');
      doc.addPageBreak();
      doc.createParagraph('After');

      const md = doc.toMarkdown();
      expect(md).toContain('Before');
      expect(md).toContain('After');
      doc.dispose();
    });
  });

  describe('edge cases', () => {
    it('returns empty string for empty document', () => {
      const doc = Document.create();
      expect(doc.toMarkdown()).toBe('');
      doc.dispose();
    });

    it('handles document with only empty paragraphs', () => {
      const doc = Document.create();
      doc.createParagraph('');
      doc.createParagraph('');

      expect(doc.toMarkdown()).toBe('');
      doc.dispose();
    });

    it('does not end with trailing newlines', () => {
      const doc = Document.create();
      doc.createParagraph('Content');

      const md = doc.toMarkdown();
      expect(md).not.toMatch(/\n$/);
      doc.dispose();
    });
  });
});
