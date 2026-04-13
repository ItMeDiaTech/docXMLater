/**
 * Tests for Document.toHTML()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';
import { Hyperlink } from '../../src/elements/Hyperlink';

describe('Document.toHTML()', () => {
  describe('headings', () => {
    it('renders headings as h1-h6 tags', () => {
      const doc = Document.create();
      doc.addHeading('Title', 1);
      doc.addHeading('Section', 2);
      doc.addHeading('Sub', 3);

      const html = doc.toHTML();

      expect(html).toContain('<h1>Title</h1>');
      expect(html).toContain('<h2>Section</h2>');
      expect(html).toContain('<h3>Sub</h3>');
      doc.dispose();
    });

    it('supports all 6 heading levels', () => {
      const doc = Document.create();
      for (let i = 1; i <= 6; i++) {
        doc.addHeading(`H${i}`, i as 1 | 2 | 3 | 4 | 5 | 6);
      }

      const html = doc.toHTML();

      for (let i = 1; i <= 6; i++) {
        expect(html).toContain(`<h${i}>H${i}</h${i}>`);
      }
      doc.dispose();
    });
  });

  describe('paragraphs', () => {
    it('renders paragraphs as p tags', () => {
      const doc = Document.create();
      doc.createParagraph('First paragraph.');
      doc.createParagraph('Second paragraph.');

      const html = doc.toHTML();

      expect(html).toContain('<p>First paragraph.</p>');
      expect(html).toContain('<p>Second paragraph.</p>');
      doc.dispose();
    });

    it('skips empty paragraphs', () => {
      const doc = Document.create();
      doc.createParagraph('Content');
      doc.createParagraph('');

      const html = doc.toHTML();

      expect(html).not.toContain('<p></p>');
      expect(html).toContain('<p>Content</p>');
      doc.dispose();
    });
  });

  describe('inline formatting', () => {
    it('renders bold as <strong>', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('bold text', { bold: true }));

      expect(doc.toHTML()).toContain('<strong>bold text</strong>');
      doc.dispose();
    });

    it('renders italic as <em>', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('italic text', { italic: true }));

      expect(doc.toHTML()).toContain('<em>italic text</em>');
      doc.dispose();
    });

    it('renders strikethrough as <s>', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('deleted', { strike: true }));

      expect(doc.toHTML()).toContain('<s>deleted</s>');
      doc.dispose();
    });

    it('renders underline as <u>', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('underlined', { underline: 'single' }));

      expect(doc.toHTML()).toContain('<u>underlined</u>');
      doc.dispose();
    });

    it('nests bold + italic tags', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('both', { bold: true, italic: true }));

      const html = doc.toHTML();

      expect(html).toContain('<strong>');
      expect(html).toContain('<em>');
      expect(html).toContain('both');
      doc.dispose();
    });

    it('renders monospace fonts as <code>', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('let x = 1', { font: 'Consolas' }));

      expect(doc.toHTML()).toContain('<code>let x = 1</code>');
      doc.dispose();
    });

    it('handles mixed formatting in a paragraph', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('Normal '));
      para.addRun(new Run('bold', { bold: true }));
      para.addRun(new Run(' and '));
      para.addRun(new Run('italic', { italic: true }));

      const html = doc.toHTML();

      expect(html).toContain('Normal <strong>bold</strong> and <em>italic</em>');
      doc.dispose();
    });
  });

  describe('hyperlinks', () => {
    it('renders hyperlinks as <a> tags', () => {
      const doc = Document.create();
      const para = doc.createParagraph();
      para.addRun(new Run('Visit '));
      para.addHyperlink(new Hyperlink({ url: 'https://example.com', text: 'Example' }));

      const html = doc.toHTML();

      expect(html).toContain('<a href="https://example.com">Example</a>');
      doc.dispose();
    });
  });

  describe('tables', () => {
    it('renders a table with thead and tbody', () => {
      const doc = Document.create();
      doc.addTable(
        Table.fromArray([
          ['Name', 'Age'],
          ['Alice', '30'],
          ['Bob', '25'],
        ])
      );

      const html = doc.toHTML();

      expect(html).toContain('<table>');
      expect(html).toContain('<thead>');
      expect(html).toContain('<th>Name</th>');
      expect(html).toContain('<th>Age</th>');
      expect(html).toContain('<tbody>');
      expect(html).toContain('<td>Alice</td>');
      expect(html).toContain('<td>25</td>');
      expect(html).toContain('</table>');
      doc.dispose();
    });

    it('handles single-row table (header only)', () => {
      const doc = Document.create();
      doc.addTable(Table.fromArray([['A', 'B', 'C']]));

      const html = doc.toHTML();

      expect(html).toContain('<thead>');
      expect(html).toContain('<th>A</th>');
      expect(html).not.toContain('<tbody>');
      doc.dispose();
    });
  });

  describe('lists', () => {
    it('renders bullet list items as <ul><li>', () => {
      const doc = Document.fromMarkdown('- First\n- Second\n- Third');

      const html = doc.toHTML();

      expect(html).toContain('<ul>');
      expect(html).toContain('<li>First</li>');
      expect(html).toContain('<li>Second</li>');
      expect(html).toContain('<li>Third</li>');
      expect(html).toContain('</ul>');
      doc.dispose();
    });

    it('renders numbered list items as <ol><li>', () => {
      const doc = Document.fromMarkdown('1. First\n2. Second');

      const html = doc.toHTML();

      expect(html).toContain('<ol>');
      expect(html).toContain('<li>First</li>');
      expect(html).toContain('<li>Second</li>');
      expect(html).toContain('</ol>');
      doc.dispose();
    });

    it('closes list when non-list content follows', () => {
      const doc = Document.fromMarkdown('- Item\n\nParagraph after list.');

      const html = doc.toHTML();

      expect(html).toContain('</ul>');
      expect(html).toContain('<p>Paragraph after list.</p>');
      // </ul> should appear before <p>
      const ulClose = html.indexOf('</ul>');
      const pOpen = html.indexOf('<p>Paragraph');
      expect(ulClose).toBeLessThan(pOpen);
      doc.dispose();
    });
  });

  describe('HTML escaping', () => {
    it('escapes special characters in text', () => {
      const doc = Document.create();
      doc.createParagraph('a < b & c > d "quoted"');

      const html = doc.toHTML();

      expect(html).toContain('a &lt; b &amp; c &gt; d &quot;quoted&quot;');
      doc.dispose();
    });

    it('escapes special characters in table cells', () => {
      const doc = Document.create();
      doc.addTable(Table.fromArray([['<script>alert("xss")</script>']]));

      const html = doc.toHTML();

      expect(html).not.toContain('<script>');
      expect(html).toContain('&lt;script&gt;');
      doc.dispose();
    });
  });

  describe('wrapInDocument option', () => {
    it('returns fragment by default', () => {
      const doc = Document.create();
      doc.createParagraph('Hello');

      const html = doc.toHTML();

      expect(html).not.toContain('<!DOCTYPE');
      expect(html).not.toContain('<html>');
      expect(html).toBe('<p>Hello</p>');
      doc.dispose();
    });

    it('wraps in full HTML document when requested', () => {
      const doc = Document.create();
      doc.createParagraph('Hello');

      const html = doc.toHTML({ wrapInDocument: true, title: 'Test Doc' });

      expect(html).toContain('<!DOCTYPE html>');
      expect(html).toContain('<html>');
      expect(html).toContain('<title>Test Doc</title>');
      expect(html).toContain('<body>');
      expect(html).toContain('<p>Hello</p>');
      expect(html).toContain('</body>');
      expect(html).toContain('</html>');
      doc.dispose();
    });

    it('uses default title when not provided', () => {
      const doc = Document.create();
      doc.createParagraph('Content');

      const html = doc.toHTML({ wrapInDocument: true });

      expect(html).toContain('<title>Document</title>');
      doc.dispose();
    });

    it('escapes title for HTML safety', () => {
      const doc = Document.create();
      const html = doc.toHTML({ wrapInDocument: true, title: 'Title <with> & "special"' });

      expect(html).toContain('Title &lt;with&gt; &amp; &quot;special&quot;');
      doc.dispose();
    });
  });

  describe('mixed content', () => {
    it('renders a complete document', () => {
      const doc = Document.create();
      doc.addHeading('Report', 1);
      doc.createParagraph('Introduction with **bold** context.');
      doc.addHeading('Data', 2);
      doc.addTable(
        Table.fromArray([
          ['Metric', 'Value'],
          ['Revenue', '$1M'],
        ])
      );
      doc.createParagraph('Conclusion.');

      const html = doc.toHTML();

      expect(html).toContain('<h1>Report</h1>');
      expect(html).toContain('<h2>Data</h2>');
      expect(html).toContain('<table>');
      expect(html).toContain('<p>Conclusion.</p>');
      doc.dispose();
    });
  });

  describe('edge cases', () => {
    it('returns empty string for empty document', () => {
      const doc = Document.create();
      expect(doc.toHTML()).toBe('');
      doc.dispose();
    });

    it('handles document with only tables', () => {
      const doc = Document.create();
      doc.addTable(Table.fromArray([['Only table']]));

      const html = doc.toHTML();

      expect(html).toContain('<table>');
      expect(html).not.toContain('<p>');
      doc.dispose();
    });
  });
});
