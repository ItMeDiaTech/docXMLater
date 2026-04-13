/**
 * Tests for Document.fromMarkdown()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';

describe('Document.fromMarkdown()', () => {
  describe('headings', () => {
    it('parses heading levels 1-6', () => {
      const md = [
        '# Heading 1',
        '## Heading 2',
        '### Heading 3',
        '#### Heading 4',
        '##### Heading 5',
        '###### Heading 6',
      ].join('\n\n');

      const doc = Document.fromMarkdown(md);
      const paras = doc.getParagraphs();

      expect(paras.filter((p) => p.getStyle() === 'Heading1')).toHaveLength(1);
      expect(paras.filter((p) => p.getStyle() === 'Heading2')).toHaveLength(1);
      expect(paras.filter((p) => p.getStyle() === 'Heading6')).toHaveLength(1);
      doc.dispose();
    });

    it('preserves heading text', () => {
      const doc = Document.fromMarkdown('# My Title');
      const heading = doc.getParagraphs().find((p) => p.getStyle() === 'Heading1');

      expect(heading).toBeDefined();
      expect(heading!.getText()).toBe('My Title');
      doc.dispose();
    });

    it('applies inline formatting within headings', () => {
      const doc = Document.fromMarkdown('# Title with **bold** word');
      const heading = doc.getParagraphs()[0]!;
      const runs = heading.getRuns();

      expect(runs.some((r) => r.getFormatting().bold && r.getText() === 'bold')).toBe(true);
      doc.dispose();
    });
  });

  describe('paragraphs', () => {
    it('creates paragraphs from plain text', () => {
      const doc = Document.fromMarkdown('First paragraph.\n\nSecond paragraph.');
      const paras = doc.getParagraphs();

      expect(paras.some((p) => p.getText() === 'First paragraph.')).toBe(true);
      expect(paras.some((p) => p.getText() === 'Second paragraph.')).toBe(true);
      doc.dispose();
    });

    it('joins continuation lines into one paragraph', () => {
      const doc = Document.fromMarkdown('Line one\nline two\nline three.');
      const paras = doc.getParagraphs();

      expect(paras.some((p) => p.getText() === 'Line one line two line three.')).toBe(true);
      doc.dispose();
    });

    it('skips blank lines', () => {
      const doc = Document.fromMarkdown('\n\n\nSolo\n\n\n');
      const paras = doc.getParagraphs().filter((p) => p.getText().trim() !== '');

      expect(paras).toHaveLength(1);
      expect(paras[0]!.getText()).toBe('Solo');
      doc.dispose();
    });
  });

  describe('inline formatting', () => {
    it('parses **bold**', () => {
      const doc = Document.fromMarkdown('This is **bold** text.');
      const runs = doc.getParagraphs()[0]!.getRuns();
      const boldRun = runs.find((r) => r.getText() === 'bold');

      expect(boldRun).toBeDefined();
      expect(boldRun!.getFormatting().bold).toBe(true);
      doc.dispose();
    });

    it('parses *italic*', () => {
      const doc = Document.fromMarkdown('This is *italic* text.');
      const runs = doc.getParagraphs()[0]!.getRuns();
      const italicRun = runs.find((r) => r.getText() === 'italic');

      expect(italicRun).toBeDefined();
      expect(italicRun!.getFormatting().italic).toBe(true);
      doc.dispose();
    });

    it('parses ***bold+italic***', () => {
      const doc = Document.fromMarkdown('This is ***both*** styled.');
      const runs = doc.getParagraphs()[0]!.getRuns();
      const bothRun = runs.find((r) => r.getText() === 'both');

      expect(bothRun).toBeDefined();
      expect(bothRun!.getFormatting().bold).toBe(true);
      expect(bothRun!.getFormatting().italic).toBe(true);
      doc.dispose();
    });

    it('parses ~~strikethrough~~', () => {
      const doc = Document.fromMarkdown('This is ~~deleted~~ text.');
      const runs = doc.getParagraphs()[0]!.getRuns();
      const strikeRun = runs.find((r) => r.getText() === 'deleted');

      expect(strikeRun).toBeDefined();
      expect(strikeRun!.getFormatting().strike).toBe(true);
      doc.dispose();
    });

    it('parses `inline code`', () => {
      const doc = Document.fromMarkdown('Use `console.log()` for output.');
      const runs = doc.getParagraphs()[0]!.getRuns();
      const codeRun = runs.find((r) => r.getText() === 'console.log()');

      expect(codeRun).toBeDefined();
      expect(codeRun!.getFormatting().font).toBe('Courier New');
      doc.dispose();
    });

    it('parses [link](url)', () => {
      const doc = Document.fromMarkdown('Visit [Example](https://example.com) now.');
      const para = doc.getParagraphs()[0]!;
      const content = para.getContent();

      // Should contain a hyperlink
      const hasHyperlink = content.some((item) => item.constructor.name === 'Hyperlink');
      expect(hasHyperlink).toBe(true);
      expect(para.getText()).toContain('Example');
      doc.dispose();
    });

    it('handles multiple inline formats in one paragraph', () => {
      const doc = Document.fromMarkdown('**Bold** and *italic* and `code`.');
      const runs = doc.getParagraphs()[0]!.getRuns();

      expect(runs.find((r) => r.getText() === 'Bold')!.getFormatting().bold).toBe(true);
      expect(runs.find((r) => r.getText() === 'italic')!.getFormatting().italic).toBe(true);
      expect(runs.find((r) => r.getText() === 'code')!.getFormatting().font).toBe('Courier New');
      doc.dispose();
    });

    it('preserves plain text between formatted spans', () => {
      const doc = Document.fromMarkdown('A **B** C');
      const text = doc.getParagraphs()[0]!.getText();

      expect(text).toBe('A B C');
      doc.dispose();
    });
  });

  describe('lists', () => {
    it('parses bullet lists with -', () => {
      const doc = Document.fromMarkdown('- First\n- Second\n- Third');
      const paras = doc.getParagraphs();
      const bullets = paras.filter((p) => p.getStyle() === 'ListBullet');

      expect(bullets).toHaveLength(3);
      expect(bullets[0]!.getText()).toBe('First');
      expect(bullets[2]!.getText()).toBe('Third');
      doc.dispose();
    });

    it('parses bullet lists with *', () => {
      const doc = Document.fromMarkdown('* Item A\n* Item B');
      const bullets = doc.getParagraphs().filter((p) => p.getStyle() === 'ListBullet');

      expect(bullets).toHaveLength(2);
      doc.dispose();
    });

    it('parses numbered lists', () => {
      const doc = Document.fromMarkdown('1. First\n2. Second\n3. Third');
      const numbered = doc.getParagraphs().filter((p) => p.getStyle() === 'ListNumber');

      expect(numbered).toHaveLength(3);
      expect(numbered[0]!.getText()).toBe('First');
      doc.dispose();
    });

    it('applies inline formatting within list items', () => {
      const doc = Document.fromMarkdown('- A **bold** item');
      const para = doc.getParagraphs().find((p) => p.getStyle() === 'ListBullet')!;
      const boldRun = para.getRuns().find((r) => r.getFormatting().bold);

      expect(boldRun).toBeDefined();
      expect(boldRun!.getText()).toBe('bold');
      doc.dispose();
    });
  });

  describe('tables', () => {
    it('parses a simple Markdown table', () => {
      const md = ['| Name | Age |', '| --- | --- |', '| Alice | 30 |', '| Bob | 25 |'].join('\n');

      const doc = Document.fromMarkdown(md);
      const tables = doc.getTables();

      expect(tables).toHaveLength(1);
      expect(tables[0]!.toArray()).toEqual([
        ['Name', 'Age'],
        ['Alice', '30'],
        ['Bob', '25'],
      ]);
      doc.dispose();
    });

    it('skips the separator row', () => {
      const md = '| H1 | H2 |\n| --- | --- |\n| D1 | D2 |';
      const doc = Document.fromMarkdown(md);

      expect(doc.getTables()[0]!.getRowCount()).toBe(2); // header + 1 data row
      doc.dispose();
    });

    it('handles table with surrounding content', () => {
      const md = 'Before\n\n| A | B |\n| - | - |\n| 1 | 2 |\n\nAfter';
      const doc = Document.fromMarkdown(md);

      expect(doc.getTables()).toHaveLength(1);
      const paras = doc.getParagraphs().filter((p) => p.getText().trim() !== '');
      const texts = paras.map((p) => p.getText());
      expect(texts).toContain('Before');
      expect(texts).toContain('After');
      doc.dispose();
    });
  });

  describe('horizontal rules', () => {
    it('parses --- as horizontal rule', () => {
      const doc = Document.fromMarkdown('Above\n\n---\n\nBelow');
      const paras = doc.getParagraphs();

      // HR creates a paragraph with bottom border
      const hrPara = paras.find((p) => p.getFormatting().borders?.bottom);
      expect(hrPara).toBeDefined();
      doc.dispose();
    });

    it('parses *** as horizontal rule', () => {
      const doc = Document.fromMarkdown('***');
      const paras = doc.getParagraphs();
      const hrPara = paras.find((p) => p.getFormatting().borders?.bottom);

      expect(hrPara).toBeDefined();
      doc.dispose();
    });

    it('parses ___ as horizontal rule', () => {
      const doc = Document.fromMarkdown('___');
      const paras = doc.getParagraphs();
      const hrPara = paras.find((p) => p.getFormatting().borders?.bottom);

      expect(hrPara).toBeDefined();
      doc.dispose();
    });
  });

  describe('round-trip', () => {
    it('fromMarkdown produces valid DOCX', async () => {
      const md = [
        '# Title',
        '',
        'A paragraph with **bold** and *italic*.',
        '',
        '| Col 1 | Col 2 |',
        '| --- | --- |',
        '| A | B |',
        '',
        '- Item 1',
        '- Item 2',
        '',
        '---',
        '',
        'Final paragraph.',
      ].join('\n');

      const doc = Document.fromMarkdown(md);
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);

      // Verify round-trip load
      const loaded = await Document.loadFromBuffer(buffer);
      const text = loaded.toPlainText();
      expect(text).toContain('Title');
      expect(text).toContain('Final paragraph');
      loaded.dispose();
      doc.dispose();
    });

    it('toMarkdown and fromMarkdown preserve structure', () => {
      const original = Document.create();
      original.addHeading('Report', 1);
      original.createParagraph('Introduction text.');
      original.addHeading('Section', 2);
      original.addTable(
        Table.fromArray([
          ['A', 'B'],
          ['1', '2'],
        ])
      );

      const md = original.toMarkdown();
      const rebuilt = Document.fromMarkdown(md);

      // Verify heading count and text
      const headings = rebuilt.getParagraphs().filter((p) => p.getStyle()?.startsWith('Heading'));
      expect(headings).toHaveLength(2);
      expect(rebuilt.getTables()).toHaveLength(1);

      original.dispose();
      rebuilt.dispose();
    });
  });

  describe('edge cases', () => {
    it('handles empty string', () => {
      const doc = Document.fromMarkdown('');
      expect(doc.getParagraphs()).toHaveLength(0);
      doc.dispose();
    });

    it('handles whitespace-only string', () => {
      const doc = Document.fromMarkdown('   \n\n   \n');
      expect(doc.getParagraphs()).toHaveLength(0);
      doc.dispose();
    });
  });
});
