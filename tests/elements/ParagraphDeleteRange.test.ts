/**
 * Tests for Paragraph.deleteRange()
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';

describe('Paragraph.deleteRange()', () => {
  describe('single run', () => {
    it('deletes from the middle of a run', () => {
      const para = new Paragraph().addText('Hello Beautiful World');
      para.deleteRange(5, 15);

      expect(para.getText()).toBe('Hello World');
    });

    it('deletes from the beginning', () => {
      const para = new Paragraph().addText('Hello World');
      para.deleteRange(0, 6);

      expect(para.getText()).toBe('World');
    });

    it('deletes from the end', () => {
      const para = new Paragraph().addText('Hello World');
      para.deleteRange(5, 11);

      expect(para.getText()).toBe('Hello');
    });

    it('deletes entire content', () => {
      const para = new Paragraph().addText('Hello');
      para.deleteRange(0, 5);

      expect(para.getText()).toBe('');
      expect(para.getRuns()).toHaveLength(0);
    });

    it('does nothing when start >= end', () => {
      const para = new Paragraph().addText('Hello');
      para.deleteRange(3, 3);

      expect(para.getText()).toBe('Hello');
    });

    it('does nothing when start > end', () => {
      const para = new Paragraph().addText('Hello');
      para.deleteRange(5, 2);

      expect(para.getText()).toBe('Hello');
    });
  });

  describe('multi-run', () => {
    it('deletes spanning two runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Hello '));
      para.addRun(new Run('World'));

      para.deleteRange(3, 8);

      expect(para.getText()).toBe('Helrld');
    });

    it('deletes an entire middle run', () => {
      const para = new Paragraph();
      para.addRun(new Run('AAA'));
      para.addRun(new Run('BBB'));
      para.addRun(new Run('CCC'));

      para.deleteRange(3, 6);

      expect(para.getText()).toBe('AAACCC');
      expect(para.getRuns()).toHaveLength(2);
    });

    it('deletes spanning all runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('AA'));
      para.addRun(new Run('BB'));
      para.addRun(new Run('CC'));

      para.deleteRange(0, 6);

      expect(para.getText()).toBe('');
    });

    it('preserves formatting on trimmed runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Bold', { bold: true }));
      para.addRun(new Run('Normal'));

      para.deleteRange(2, 8);

      expect(para.getText()).toBe('Boal');
      const runs = para.getRuns();
      expect(runs[0]!.getFormatting().bold).toBe(true);
    });
  });

  describe('edge cases', () => {
    it('handles empty paragraph', () => {
      const para = new Paragraph();
      para.deleteRange(0, 5);

      expect(para.getText()).toBe('');
    });

    it('handles range beyond text length', () => {
      const para = new Paragraph().addText('Hi');
      para.deleteRange(0, 100);

      expect(para.getText()).toBe('');
    });

    it('returns this for chaining', () => {
      const para = new Paragraph().addText('Hello World');
      const result = para.deleteRange(0, 5);

      expect(result).toBe(para);
    });

    it('delete single character', () => {
      const para = new Paragraph().addText('Hello');
      para.deleteRange(2, 3);

      expect(para.getText()).toBe('Helo');
    });
  });

  describe('text integrity', () => {
    it('preserves content outside the deleted range', () => {
      const original = 'The quick brown fox jumps over the lazy dog';
      const para = new Paragraph().addText(original);

      para.deleteRange(10, 19); // delete "brown fox"

      expect(para.getText()).toBe('The quick  jumps over the lazy dog');
    });
  });
});
