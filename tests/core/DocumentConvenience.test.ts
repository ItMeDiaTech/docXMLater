/**
 * Tests for Document convenience methods: addHeading, addPageBreak, addHorizontalRule
 */

import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';

describe('Document convenience methods', () => {
  describe('addHeading()', () => {
    it('creates a Heading1 paragraph by default', () => {
      const doc = Document.create();
      const para = doc.addHeading('Introduction');

      expect(para.getText()).toBe('Introduction');
      expect(para.getStyle()).toBe('Heading1');
      doc.dispose();
    });

    it('creates headings at specified levels', () => {
      const doc = Document.create();
      const h1 = doc.addHeading('Chapter', 1);
      const h2 = doc.addHeading('Section', 2);
      const h3 = doc.addHeading('Subsection', 3);
      const h4 = doc.addHeading('Detail', 4);

      expect(h1.getStyle()).toBe('Heading1');
      expect(h2.getStyle()).toBe('Heading2');
      expect(h3.getStyle()).toBe('Heading3');
      expect(h4.getStyle()).toBe('Heading4');
      doc.dispose();
    });

    it('appends heading to document body', () => {
      const doc = Document.create();
      doc.addHeading('First');
      doc.addHeading('Second', 2);

      const paragraphs = doc.getParagraphs();
      expect(paragraphs.length).toBeGreaterThanOrEqual(2);
      doc.dispose();
    });

    it('returns paragraph for further customization', () => {
      const doc = Document.create();
      const para = doc.addHeading('Centered Heading', 1);
      para.setAlignment('center');

      expect(para.getAlignment()).toBe('center');
      doc.dispose();
    });

    it('supports all heading levels 1-9', () => {
      const doc = Document.create();
      for (let level = 1; level <= 9; level++) {
        const para = doc.addHeading(`Level ${level}`, level as 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9);
        expect(para.getStyle()).toBe(`Heading${level}`);
      }
      doc.dispose();
    });
  });

  describe('addPageBreak()', () => {
    it('creates a paragraph with a page break', () => {
      const doc = Document.create();
      const para = doc.addPageBreak();

      // Should have a run containing a break
      const runs = para.getRuns();
      expect(runs.length).toBeGreaterThanOrEqual(1);

      // The run should contain a page break in its content
      const breakRun = runs.find((r) => {
        const content = r.getContent();
        return content.some((c) => c.type === 'break' && c.breakType === 'page');
      });
      expect(breakRun).toBeDefined();
      doc.dispose();
    });

    it('appends page break to document body', () => {
      const doc = Document.create();
      doc.createParagraph('Before');
      doc.addPageBreak();
      doc.createParagraph('After');

      const paragraphs = doc.getParagraphs();
      expect(paragraphs.length).toBeGreaterThanOrEqual(3);
      doc.dispose();
    });

    it('returns paragraph for chaining', () => {
      const doc = Document.create();
      const para = doc.addPageBreak();

      // Can add more content after the break in the same paragraph
      expect(para).toBeDefined();
      expect(typeof para.addText).toBe('function');
      doc.dispose();
    });

    it('generates valid XML with page break', async () => {
      const doc = Document.create();
      doc.createParagraph('Page 1');
      doc.addPageBreak();
      doc.createParagraph('Page 2');

      // Should save without errors
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  describe('addHorizontalRule()', () => {
    it('creates a paragraph with a bottom border', () => {
      const doc = Document.create();
      const para = doc.addHorizontalRule();

      // The paragraph should have border formatting
      const formatting = para.getFormatting();
      expect(formatting.borders?.bottom).toBeDefined();
      expect(formatting.borders!.bottom!.style).toBe('single');
      doc.dispose();
    });

    it('uses default color and size', () => {
      const doc = Document.create();
      const para = doc.addHorizontalRule();

      const border = para.getFormatting().borders?.bottom;
      expect(border?.color).toBe('auto');
      expect(border?.size).toBe(4);
      doc.dispose();
    });

    it('accepts custom color', () => {
      const doc = Document.create();
      const para = doc.addHorizontalRule('FF0000');

      const border = para.getFormatting().borders?.bottom;
      expect(border?.color).toBe('FF0000');
      doc.dispose();
    });

    it('accepts custom size', () => {
      const doc = Document.create();
      const para = doc.addHorizontalRule('auto', 12);

      const border = para.getFormatting().borders?.bottom;
      expect(border?.size).toBe(12);
      doc.dispose();
    });

    it('appends to document body', () => {
      const doc = Document.create();
      doc.createParagraph('Above');
      doc.addHorizontalRule();
      doc.createParagraph('Below');

      const paragraphs = doc.getParagraphs();
      expect(paragraphs.length).toBeGreaterThanOrEqual(3);
      doc.dispose();
    });

    it('generates valid XML', async () => {
      const doc = Document.create();
      doc.createParagraph('Content');
      doc.addHorizontalRule('000000', 8);
      doc.createParagraph('More content');

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  describe('combined usage', () => {
    it('builds a structured document with all convenience methods', async () => {
      const doc = Document.create();

      doc.addHeading('Document Title', 1);
      doc.createParagraph('Opening paragraph with introduction text.');
      doc.addHorizontalRule();

      doc.addHeading('Chapter 1', 2);
      doc.createParagraph('Chapter 1 content goes here.');

      doc.addPageBreak();

      doc.addHeading('Chapter 2', 2);
      doc.createParagraph('Chapter 2 content goes here.');

      doc.addHorizontalRule('0000FF', 6);
      doc.createParagraph('Footer text.');

      // Verify structure
      const paragraphs = doc.getParagraphs();
      expect(paragraphs.length).toBeGreaterThanOrEqual(8);

      // Verify it saves correctly
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);

      doc.dispose();
    });
  });
});
