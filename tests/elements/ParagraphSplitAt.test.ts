/**
 * Tests for Paragraph.splitAt()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Hyperlink } from '../../src/elements/Hyperlink';

describe('Paragraph.splitAt()', () => {
  describe('single run splitting', () => {
    it('splits a paragraph in the middle of a run', () => {
      const para = new Paragraph().addText('Hello World');
      const tail = para.splitAt(5);

      expect(para.getText()).toBe('Hello');
      expect(tail.getText()).toBe(' World');
    });

    it('splits at position 0 (moves everything to tail)', () => {
      const para = new Paragraph().addText('Hello');
      const tail = para.splitAt(0);

      expect(para.getText()).toBe('');
      expect(tail.getText()).toBe('Hello');
    });

    it('splits at end (returns empty tail)', () => {
      const para = new Paragraph().addText('Hello');
      const tail = para.splitAt(5);

      expect(para.getText()).toBe('Hello');
      expect(tail.getText()).toBe('');
    });

    it('splits past end (returns empty tail)', () => {
      const para = new Paragraph().addText('Hello');
      const tail = para.splitAt(100);

      expect(para.getText()).toBe('Hello');
      expect(tail.getText()).toBe('');
    });

    it('splits at first character', () => {
      const para = new Paragraph().addText('ABCDEF');
      const tail = para.splitAt(1);

      expect(para.getText()).toBe('A');
      expect(tail.getText()).toBe('BCDEF');
    });

    it('splits at last character', () => {
      const para = new Paragraph().addText('ABCDEF');
      const tail = para.splitAt(5);

      expect(para.getText()).toBe('ABCDE');
      expect(tail.getText()).toBe('F');
    });
  });

  describe('multi-run splitting', () => {
    it('splits at a run boundary', () => {
      const para = new Paragraph();
      para.addRun(new Run('AAA'));
      para.addRun(new Run('BBB'));
      para.addRun(new Run('CCC'));

      const tail = para.splitAt(3); // Exactly at AAA/BBB boundary

      expect(para.getText()).toBe('AAA');
      expect(tail.getText()).toBe('BBBCCC');
      expect(para.getRuns()).toHaveLength(1);
      expect(tail.getRuns()).toHaveLength(2);
    });

    it('splits within the second run', () => {
      const para = new Paragraph();
      para.addRun(new Run('AAA'));
      para.addRun(new Run('BBB'));
      para.addRun(new Run('CCC'));

      const tail = para.splitAt(5); // Mid "BBB"

      expect(para.getText()).toBe('AAABB');
      expect(tail.getText()).toBe('BCCC');
    });

    it('splits within the first run of many', () => {
      const para = new Paragraph();
      para.addRun(new Run('AAA'));
      para.addRun(new Run('BBB'));

      const tail = para.splitAt(1);

      expect(para.getText()).toBe('A');
      expect(tail.getText()).toBe('AABBB');
    });
  });

  describe('formatting preservation', () => {
    it('inherits paragraph style', () => {
      const para = new Paragraph().addText('Heading Text');
      para.setStyle('Heading1');
      para.setAlignment('center');

      const tail = para.splitAt(7);

      expect(tail.getStyle()).toBe('Heading1');
      expect(tail.getAlignment()).toBe('center');
    });

    it('preserves run formatting in both halves', () => {
      const para = new Paragraph();
      para.addRun(new Run('Bold', { bold: true }));
      para.addRun(new Run('Italic', { italic: true }));

      const tail = para.splitAt(4); // Split at Bold/Italic boundary

      const headRuns = para.getRuns();
      expect(headRuns[0]!.getFormatting().bold).toBe(true);

      const tailRuns = tail.getRuns();
      expect(tailRuns[0]!.getFormatting().italic).toBe(true);
    });

    it('preserves formatting when splitting within a formatted run', () => {
      const para = new Paragraph();
      para.addRun(new Run('Hello World', { bold: true, color: 'FF0000' }));

      const tail = para.splitAt(5);

      expect(para.getRuns()[0]!.getFormatting().bold).toBe(true);
      expect(para.getRuns()[0]!.getFormatting().color).toBe('FF0000');
      expect(tail.getRuns()[0]!.getFormatting().bold).toBe(true);
      expect(tail.getRuns()[0]!.getFormatting().color).toBe('FF0000');
    });

    it('deep-clones paragraph formatting (mutations are independent)', () => {
      const para = new Paragraph().addText('Hello World');
      para.setAlignment('center');

      const tail = para.splitAt(5);
      tail.setAlignment('left');

      expect(para.getAlignment()).toBe('center');
      expect(tail.getAlignment()).toBe('left');
    });
  });

  describe('non-Run content', () => {
    it('moves hyperlinks that appear after the split point', () => {
      const para = new Paragraph();
      para.addRun(new Run('Visit '));
      para.addHyperlink(new Hyperlink({ url: 'https://example.com', text: 'here' }));

      // "Visit " is 6 chars, split at 3
      const tail = para.splitAt(3);

      expect(para.getText()).toBe('Vis');
      // tail should have "it " + hyperlink
      expect(tail.getText()).toContain('it ');
      expect(tail.getText()).toContain('here');
    });
  });

  describe('edge cases', () => {
    it('handles empty paragraph', () => {
      const para = new Paragraph();
      const tail = para.splitAt(0);

      expect(para.getText()).toBe('');
      expect(tail.getText()).toBe('');
    });

    it('handles negative offset (moves everything)', () => {
      const para = new Paragraph().addText('Hello');
      const tail = para.splitAt(-5);

      expect(para.getText()).toBe('');
      expect(tail.getText()).toBe('Hello');
    });

    it('preserves total text content', () => {
      const original = 'The quick brown fox jumps over the lazy dog';
      const para = new Paragraph().addText(original);

      const tail = para.splitAt(20);

      expect(para.getText() + tail.getText()).toBe(original);
    });
  });

  describe('integration with Document', () => {
    it('can insert a table between split paragraph halves', () => {
      const doc = Document.create();
      const para = doc.createParagraph('Before table. After table.');
      const table = doc.createTable(1, 2);
      table.getCell(0, 0)!.createParagraph('A');
      table.getCell(0, 1)!.createParagraph('B');

      // Split the paragraph at position 14 ("Before table. " | "After table.")
      const tail = para.splitAt(14);

      // Find the paragraph index and insert tail after the table
      const paraIndex = doc.getBodyElementIndex(para);
      // Table is already the last element, insert tail after it
      doc.insertParagraphAt(doc.getBodyElementCount(), tail);

      // Verify document structure
      const elements = doc.getBodyElements();
      const paragraphs = elements.filter((e) => e instanceof Paragraph);
      expect(paragraphs.length).toBeGreaterThanOrEqual(2);

      // First paragraph should be "Before table. "
      expect(para.getText()).toBe('Before table. ');
      // Tail paragraph should be "After table."
      expect(tail.getText()).toBe('After table.');

      doc.dispose();
    });

    it('generates valid XML after split', () => {
      const para = new Paragraph().addText('Hello World');
      const tail = para.splitAt(5);

      const headXml = para.toXML();
      const tailXml = tail.toXML();

      expect(headXml.name).toBe('w:p');
      expect(tailXml.name).toBe('w:p');

      // Both should have runs
      const headRuns = headXml.children?.filter((c) => typeof c !== 'string' && c.name === 'w:r');
      const tailRuns = tailXml.children?.filter((c) => typeof c !== 'string' && c.name === 'w:r');
      expect(headRuns!.length).toBeGreaterThanOrEqual(1);
      expect(tailRuns!.length).toBeGreaterThanOrEqual(1);
    });
  });

  describe('chaining with other operations', () => {
    it('splitAt + consolidateRuns works correctly', () => {
      const para = new Paragraph();
      para.addRun(new Run('AAABBB', { bold: true }));

      const tail = para.splitAt(3);

      // Both paragraphs should have bold runs
      expect(para.getRuns()[0]!.getFormatting().bold).toBe(true);
      expect(tail.getRuns()[0]!.getFormatting().bold).toBe(true);

      // Can consolidate after further operations
      para.addRun(new Run('CCC', { bold: true }));
      para.consolidateRuns();
      expect(para.getRuns()).toHaveLength(1);
      expect(para.getText()).toBe('AAACCC');
    });

    it('splitAt + applyFormattingToRange works correctly', () => {
      const para = new Paragraph().addText('Hello Beautiful World');

      // Split, then format the tail
      const tail = para.splitAt(6);
      tail.applyFormattingToRange(0, 9, { bold: true });

      expect(para.getText()).toBe('Hello ');
      expect(tail.getText()).toBe('Beautiful World');

      const boldRun = tail.getRuns().find((r) => r.getFormatting().bold);
      expect(boldRun).toBeDefined();
      expect(boldRun!.getText()).toBe('Beautiful');
    });
  });
});
