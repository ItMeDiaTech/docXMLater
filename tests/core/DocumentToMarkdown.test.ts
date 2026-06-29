/**
 * Tests for Document.toMarkdown()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Image } from '../../src/elements/Image';
import { ImageRun } from '../../src/elements/ImageRun';

/** Minimal 8-byte PNG signature, sufficient for format detection. */
const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

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

  describe('rich inline formatting (HTML fallback)', () => {
    it('renders underline as <u>', () => {
      const doc = Document.create();
      doc.createParagraph().addRun(new Run('underlined', { underline: true }));
      expect(doc.toMarkdown()).toBe('<u>underlined</u>');
      doc.dispose();
    });

    it('renders superscript as <sup> and subscript as <sub>', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('E=mc'));
      para.addRun(new Run('2', { superscript: true }));
      para.addRun(new Run(' and H'));
      para.addRun(new Run('2', { subscript: true }));
      para.addRun(new Run('O'));
      expect(doc.toMarkdown()).toBe('E=mc<sup>2</sup> and H<sub>2</sub>O');
      doc.dispose();
    });

    it('renders highlight as <mark>', () => {
      const doc = Document.create();
      doc.createParagraph().addRun(new Run('marked', { highlight: 'yellow' }));
      expect(doc.toMarkdown()).toBe('<mark>marked</mark>');
      doc.dispose();
    });

    it('renders text color as a span', () => {
      const doc = Document.create();
      doc.createParagraph().addRun(new Run('red', { color: 'FF0000' }));
      expect(doc.toMarkdown()).toBe('<span style="color:#FF0000">red</span>');
      doc.dispose();
    });

    it('combines emphasis with HTML fallback', () => {
      const doc = Document.create();
      doc.createParagraph().addRun(new Run('x', { bold: true, superscript: true }));
      expect(doc.toMarkdown()).toBe('<sup>**x**</sup>');
      doc.dispose();
    });

    it('drops HTML fallback when htmlFallback: false', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('a', { underline: true }));
      para.addRun(new Run('b', { highlight: 'yellow' }));
      para.addRun(new Run('c', { color: 'FF0000' }));
      expect(doc.toMarkdown({ htmlFallback: false })).toBe('abc');
      doc.dispose();
    });
  });

  describe('breaks, tabs, and escaping', () => {
    it('renders line breaks as <br>', () => {
      const doc = Document.create();
      const run = new Run('Line1');
      run.addBreak();
      run.appendText('Line2');
      doc.createParagraph().addRun(run);
      expect(doc.toMarkdown()).toBe('Line1<br>Line2');
      doc.dispose();
    });

    it('renders line breaks as newline when htmlFallback: false', () => {
      const doc = Document.create();
      const run = new Run('Line1');
      run.addBreak();
      run.appendText('Line2');
      doc.createParagraph().addRun(run);
      expect(doc.toMarkdown({ htmlFallback: false })).toBe('Line1\nLine2');
      doc.dispose();
    });

    it('ignores page breaks within text (layout only)', () => {
      const doc = Document.create();
      const run = new Run('Before');
      run.addBreak('page');
      run.appendText('After');
      doc.createParagraph().addRun(run);
      expect(doc.toMarkdown()).toBe('BeforeAfter');
      doc.dispose();
    });

    it('preserves tabs', () => {
      const doc = Document.create();
      const run = new Run('A');
      run.addTab();
      run.appendText('B');
      doc.createParagraph().addRun(run);
      expect(doc.toMarkdown()).toBe('A\tB');
      doc.dispose();
    });

    it('escapes Markdown-significant characters in literal text', () => {
      const doc = Document.create();
      doc.createParagraph('Use *stars* and _under_ and [brackets]');
      const md = doc.toMarkdown();
      expect(md).toBe('Use \\*stars\\* and \\_under\\_ and \\[brackets\\]');
      doc.dispose();
    });
  });

  describe('lists', () => {
    it('renders bullet lists with - markers', () => {
      const doc = Document.create();
      doc.addBulletListFromArray(['Apple', 'Banana']);
      const md = doc.toMarkdown();
      expect(md).toContain('- Apple');
      expect(md).toContain('- Banana');
      doc.dispose();
    });

    it('renders numbered lists with ordered markers', () => {
      const doc = Document.create();
      doc.addNumberedListFromArray(['First', 'Second']);
      const md = doc.toMarkdown();
      expect(md).toContain('1. First');
      expect(md).toContain('1. Second');
      doc.dispose();
    });

    it('indents nested list levels by two spaces', () => {
      const doc = Document.create();
      doc.addNumberedListFromArray(['One', { text: 'Sub', level: 1 }, 'Two']);
      const md = doc.toMarkdown();
      expect(md).toContain('\n  1. Sub');
      doc.dispose();
    });
  });

  describe('block quotes', () => {
    it('prefixes Quote-styled paragraphs with >', () => {
      const doc = Document.create();
      const para = doc.createParagraph('A wise quote.');
      para.setStyle('Quote');
      expect(doc.toMarkdown()).toBe('> A wise quote.');
      doc.dispose();
    });
  });

  describe('images', () => {
    it('renders inline images as Markdown image syntax', async () => {
      const doc = Document.create();
      const image = await Image.fromBuffer(PNG_SIGNATURE, { width: 914400, height: 914400 });
      image.setAltText('Company logo');
      doc.createParagraph().addRun(new ImageRun(image));
      const md = doc.toMarkdown();
      expect(md).toMatch(/^!\[Company logo\]\(.+\)$/);
      doc.dispose();
    });

    it('omits images when images: false', async () => {
      const doc = Document.create();
      const image = await Image.fromBuffer(PNG_SIGNATURE, { width: 914400, height: 914400 });
      image.setAltText('logo');
      const para = doc.createParagraph();
      para.addRun(new Run('Text'));
      para.addRun(new ImageRun(image));
      expect(doc.toMarkdown({ images: false })).toBe('Text');
      doc.dispose();
    });
  });

  describe('footnotes', () => {
    it('emits footnote markers and appends definitions', () => {
      const doc = Document.create();
      const footnote = doc.createFootnote('The footnote body.');
      const para = doc.createParagraph();
      para.addRun(new Run('Anchor'));
      para.addRun(
        Run.createFromContent([{ type: 'footnoteReference', footnoteId: footnote.getId() }])
      );

      const md = doc.toMarkdown();
      const id = footnote.getId();
      expect(md).toContain(`Anchor[^fn${id}]`);
      expect(md).toContain(`[^fn${id}]: The footnote body.`);
      doc.dispose();
    });

    it('omits footnote markers when footnotes: false', () => {
      const doc = Document.create();
      const footnote = doc.createFootnote('Body.');
      const para = doc.createParagraph();
      para.addRun(new Run('Anchor'));
      para.addRun(
        Run.createFromContent([{ type: 'footnoteReference', footnoteId: footnote.getId() }])
      );

      const md = doc.toMarkdown({ footnotes: false });
      expect(md).toBe('Anchor');
      doc.dispose();
    });
  });

  describe('complex tables', () => {
    it('falls back to HTML for tables with horizontally merged cells', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getCell(0, 0)!.createParagraph('Spanning header');
      table.getCell(0, 0)!.setColumnSpan(2);
      table.getCell(1, 0)!.createParagraph('A');
      table.getCell(1, 1)!.createParagraph('B');
      doc.addTable(table);

      const md = doc.toMarkdown();
      expect(md).toContain('<table>');
      expect(md).toContain('colspan="2"');
      expect(md).toContain('Spanning header');
      doc.dispose();
    });

    it('falls back to HTML for tables with vertically merged cells', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getCell(0, 0)!.createParagraph('Tall');
      table.getCell(0, 0)!.setVerticalMerge('restart');
      table.getCell(1, 0)!.setVerticalMerge('continue');
      doc.addTable(table);

      const md = doc.toMarkdown();
      expect(md).toContain('<table>');
      expect(md).toContain('rowspan="2"');
      doc.dispose();
    });

    it('renders cell inline formatting in simple pipe tables', () => {
      const doc = Document.create();
      const table = new Table(1, 1);
      const cell = table.getCell(0, 0)!;
      cell.createParagraph().addRun(new Run('bold', { bold: true }));
      doc.addTable(table);

      const md = doc.toMarkdown();
      expect(md).toContain('| **bold** |');
      doc.dispose();
    });
  });

  describe('round-trip from a saved document', () => {
    it('preserves headings, formatting, lists, and tables through save/load', async () => {
      const doc = Document.create();
      doc.addHeading('Report', 1);
      doc.createParagraph().addRun(new Run('Intro with bold', { bold: true }));
      // whole-run bold renders as **Intro with bold**
      doc.addBulletListFromArray(['Point A', 'Point B']);
      const table = Table.fromArray([
        ['Key', 'Value'],
        ['x', '1'],
      ]);
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const md = loaded.toMarkdown();

      expect(md).toContain('# Report');
      expect(md).toContain('**Intro with bold**');
      expect(md).toContain('- Point A');
      expect(md).toContain('| Key | Value |');
      expect(md).toContain('| x | 1 |');
      loaded.dispose();
      doc.dispose();
    });
  });
});
