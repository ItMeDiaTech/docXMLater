/**
 * Tests for Paragraph.applyFormattingToRange()
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';

describe('Paragraph.applyFormattingToRange()', () => {
  describe('single run scenarios', () => {
    it('applies formatting to entire run', () => {
      const para = new Paragraph().addText('Hello World');
      para.applyFormattingToRange(0, 11, { bold: true });

      const runs = para.getRuns();
      expect(runs).toHaveLength(1);
      expect(runs[0]!.getText()).toBe('Hello World');
      expect(runs[0]!.getFormatting().bold).toBe(true);
    });

    it('applies formatting to beginning of run', () => {
      const para = new Paragraph().addText('Hello World');
      para.applyFormattingToRange(0, 5, { bold: true });

      const runs = para.getRuns();
      expect(runs).toHaveLength(2);
      expect(runs[0]!.getText()).toBe('Hello');
      expect(runs[0]!.getFormatting().bold).toBe(true);
      expect(runs[1]!.getText()).toBe(' World');
      expect(runs[1]!.getFormatting().bold).toBeUndefined();
    });

    it('applies formatting to end of run', () => {
      const para = new Paragraph().addText('Hello World');
      para.applyFormattingToRange(6, 11, { italic: true });

      const runs = para.getRuns();
      expect(runs).toHaveLength(2);
      expect(runs[0]!.getText()).toBe('Hello ');
      expect(runs[0]!.getFormatting().italic).toBeUndefined();
      expect(runs[1]!.getText()).toBe('World');
      expect(runs[1]!.getFormatting().italic).toBe(true);
    });

    it('applies formatting to middle of run (three-way split)', () => {
      const para = new Paragraph().addText('Hello Beautiful World');
      para.applyFormattingToRange(6, 15, { bold: true, color: 'FF0000' });

      const runs = para.getRuns();
      expect(runs).toHaveLength(3);
      expect(runs[0]!.getText()).toBe('Hello ');
      expect(runs[0]!.getFormatting().bold).toBeUndefined();
      expect(runs[1]!.getText()).toBe('Beautiful');
      expect(runs[1]!.getFormatting().bold).toBe(true);
      expect(runs[1]!.getFormatting().color).toBe('FF0000');
      expect(runs[2]!.getText()).toBe(' World');
      expect(runs[2]!.getFormatting().bold).toBeUndefined();
    });
  });

  describe('multi-run scenarios', () => {
    it('applies formatting spanning two runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Hello '));
      para.addRun(new Run('World'));

      // Bold from position 3 to 8: "lo Wo"
      para.applyFormattingToRange(3, 8, { bold: true });

      const runs = para.getRuns();
      // "Hel" | "lo " (bold) | "Wo" (bold) | "rld"
      expect(runs).toHaveLength(4);
      expect(runs[0]!.getText()).toBe('Hel');
      expect(runs[0]!.getFormatting().bold).toBeUndefined();
      expect(runs[1]!.getText()).toBe('lo ');
      expect(runs[1]!.getFormatting().bold).toBe(true);
      expect(runs[2]!.getText()).toBe('Wo');
      expect(runs[2]!.getFormatting().bold).toBe(true);
      expect(runs[3]!.getText()).toBe('rld');
      expect(runs[3]!.getFormatting().bold).toBeUndefined();
    });

    it('applies formatting to exactly one full run among many', () => {
      const para = new Paragraph();
      para.addRun(new Run('AAA'));
      para.addRun(new Run('BBB'));
      para.addRun(new Run('CCC'));

      para.applyFormattingToRange(3, 6, { highlight: 'yellow' });

      const runs = para.getRuns();
      expect(runs).toHaveLength(3);
      expect(runs[0]!.getText()).toBe('AAA');
      expect(runs[0]!.getFormatting().highlight).toBeUndefined();
      expect(runs[1]!.getText()).toBe('BBB');
      expect(runs[1]!.getFormatting().highlight).toBe('yellow');
      expect(runs[2]!.getText()).toBe('CCC');
      expect(runs[2]!.getFormatting().highlight).toBeUndefined();
    });

    it('applies formatting spanning all runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('A'));
      para.addRun(new Run('B'));
      para.addRun(new Run('C'));

      para.applyFormattingToRange(0, 3, { underline: 'double' });

      const runs = para.getRuns();
      expect(runs).toHaveLength(3);
      for (const run of runs) {
        expect(run.getFormatting().underline).toBe('double');
      }
    });
  });

  describe('edge cases', () => {
    it('does nothing when start >= end', () => {
      const para = new Paragraph().addText('Hello');
      para.applyFormattingToRange(5, 3, { bold: true });

      const runs = para.getRuns();
      expect(runs).toHaveLength(1);
      expect(runs[0]!.getFormatting().bold).toBeUndefined();
    });

    it('does nothing when start equals end', () => {
      const para = new Paragraph().addText('Hello');
      para.applyFormattingToRange(3, 3, { bold: true });

      const runs = para.getRuns();
      expect(runs).toHaveLength(1);
      expect(runs[0]!.getFormatting().bold).toBeUndefined();
    });

    it('handles range beyond text length', () => {
      const para = new Paragraph().addText('Hi');
      para.applyFormattingToRange(0, 100, { bold: true });

      const runs = para.getRuns();
      expect(runs).toHaveLength(1);
      expect(runs[0]!.getText()).toBe('Hi');
      expect(runs[0]!.getFormatting().bold).toBe(true);
    });

    it('handles empty paragraph', () => {
      const para = new Paragraph();
      para.applyFormattingToRange(0, 5, { bold: true });
      expect(para.getRuns()).toHaveLength(0);
    });

    it('preserves existing formatting on non-targeted runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Bold', { bold: true }));
      para.addRun(new Run(' Normal'));

      para.applyFormattingToRange(5, 11, { italic: true });

      const runs = para.getRuns();
      expect(runs[0]!.getText()).toBe('Bold');
      expect(runs[0]!.getFormatting().bold).toBe(true);
      expect(runs[0]!.getFormatting().italic).toBeUndefined();
    });
  });

  describe('multiple formatting properties', () => {
    it('applies multiple properties at once', () => {
      const para = new Paragraph().addText('Format me');
      para.applyFormattingToRange(0, 6, {
        bold: true,
        italic: true,
        color: '0000FF',
        size: 14,
        font: 'Courier New',
      });

      const runs = para.getRuns();
      const formatted = runs[0]!;
      expect(formatted.getText()).toBe('Format');
      expect(formatted.getFormatting().bold).toBe(true);
      expect(formatted.getFormatting().italic).toBe(true);
      expect(formatted.getFormatting().color).toBe('0000FF');
      expect(formatted.getFormatting().size).toBe(14);
      expect(formatted.getFormatting().font).toBe('Courier New');
    });
  });

  describe('chaining', () => {
    it('returns paragraph for chaining', () => {
      const para = new Paragraph().addText('Hello World');
      const result = para.applyFormattingToRange(0, 5, { bold: true });
      expect(result).toBe(para);
    });

    it('supports multiple non-overlapping ranges', () => {
      const para = new Paragraph().addText('AABBCC');

      para.applyFormattingToRange(0, 2, { bold: true });
      para.applyFormattingToRange(4, 6, { italic: true });

      const runs = para.getRuns();
      // "AA" (bold) | "BB" | "CC" (italic)
      expect(runs).toHaveLength(3);
      expect(runs[0]!.getText()).toBe('AA');
      expect(runs[0]!.getFormatting().bold).toBe(true);
      expect(runs[1]!.getText()).toBe('BB');
      expect(runs[1]!.getFormatting().bold).toBeUndefined();
      expect(runs[1]!.getFormatting().italic).toBeUndefined();
      expect(runs[2]!.getText()).toBe('CC');
      expect(runs[2]!.getFormatting().italic).toBe(true);
    });
  });

  describe('text integrity', () => {
    it('preserves total text content after formatting', () => {
      const original = 'The quick brown fox jumps over the lazy dog';
      const para = new Paragraph().addText(original);

      para.applyFormattingToRange(4, 9, { bold: true });
      para.applyFormattingToRange(20, 25, { italic: true });

      expect(para.getText()).toBe(original);
    });

    it('generates valid XML after splitting', () => {
      const para = new Paragraph().addText('Hello World');
      para.applyFormattingToRange(6, 11, { bold: true });

      const xml = para.toXML();
      expect(xml.name).toBe('w:p');
      // Should have runs as children
      const runs = xml.children?.filter((c) => typeof c !== 'string' && c.name === 'w:r');
      expect(runs!.length).toBe(2);
    });
  });
});
