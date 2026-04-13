/**
 * Tests for Paragraph.getRunAtOffset() and Paragraph.getFormattingAtOffset()
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';

describe('Paragraph.getRunAtOffset()', () => {
  describe('single run', () => {
    it('returns the run at offset 0', () => {
      const para = new Paragraph().addText('Hello');
      const result = para.getRunAtOffset(0);

      expect(result).toBeDefined();
      expect(result!.run.getText()).toBe('Hello');
      expect(result!.localOffset).toBe(0);
    });

    it('returns correct local offset within run', () => {
      const para = new Paragraph().addText('Hello');
      const result = para.getRunAtOffset(3);

      expect(result).toBeDefined();
      expect(result!.run.getText()).toBe('Hello');
      expect(result!.localOffset).toBe(3);
    });

    it('returns last character position', () => {
      const para = new Paragraph().addText('Hello');
      const result = para.getRunAtOffset(4);

      expect(result).toBeDefined();
      expect(result!.localOffset).toBe(4);
    });

    it('returns undefined for offset at text length (past end)', () => {
      const para = new Paragraph().addText('Hello');
      expect(para.getRunAtOffset(5)).toBeUndefined();
    });

    it('returns undefined for offset past text length', () => {
      const para = new Paragraph().addText('Hi');
      expect(para.getRunAtOffset(100)).toBeUndefined();
    });
  });

  describe('multiple runs', () => {
    it('returns first run for offsets in first run', () => {
      const run1 = new Run('Hello ', { bold: true });
      const run2 = new Run('World');
      const para = new Paragraph();
      para.addRun(run1);
      para.addRun(run2);

      const result = para.getRunAtOffset(3);
      expect(result).toBeDefined();
      expect(result!.run).toBe(run1);
      expect(result!.localOffset).toBe(3);
    });

    it('returns second run for offsets in second run', () => {
      const run1 = new Run('Hello ', { bold: true });
      const run2 = new Run('World');
      const para = new Paragraph();
      para.addRun(run1);
      para.addRun(run2);

      const result = para.getRunAtOffset(8);
      expect(result).toBeDefined();
      expect(result!.run).toBe(run2);
      expect(result!.localOffset).toBe(2); // 'r' in World
    });

    it('returns second run at the exact boundary', () => {
      const para = new Paragraph();
      para.addRun(new Run('AAA'));
      para.addRun(new Run('BBB'));

      const result = para.getRunAtOffset(3);
      expect(result).toBeDefined();
      expect(result!.run.getText()).toBe('BBB');
      expect(result!.localOffset).toBe(0);
    });

    it('handles three runs correctly', () => {
      const para = new Paragraph();
      para.addRun(new Run('AB'));
      para.addRun(new Run('CD'));
      para.addRun(new Run('EF'));

      expect(para.getRunAtOffset(0)!.run.getText()).toBe('AB');
      expect(para.getRunAtOffset(1)!.run.getText()).toBe('AB');
      expect(para.getRunAtOffset(2)!.run.getText()).toBe('CD');
      expect(para.getRunAtOffset(3)!.run.getText()).toBe('CD');
      expect(para.getRunAtOffset(4)!.run.getText()).toBe('EF');
      expect(para.getRunAtOffset(5)!.run.getText()).toBe('EF');
      expect(para.getRunAtOffset(6)).toBeUndefined();
    });
  });

  describe('edge cases', () => {
    it('returns undefined for empty paragraph', () => {
      const para = new Paragraph();
      expect(para.getRunAtOffset(0)).toBeUndefined();
    });

    it('returns undefined for negative offset', () => {
      const para = new Paragraph().addText('Hello');
      expect(para.getRunAtOffset(-1)).toBeUndefined();
    });

    it('skips empty runs', () => {
      const para = new Paragraph();
      para.addRun(new Run(''));
      para.addRun(new Run('Hello'));

      const result = para.getRunAtOffset(0);
      expect(result).toBeDefined();
      expect(result!.run.getText()).toBe('Hello');
      expect(result!.localOffset).toBe(0);
    });
  });
});

describe('Paragraph.getFormattingAtOffset()', () => {
  describe('basic formatting queries', () => {
    it('returns formatting of a bold run', () => {
      const para = new Paragraph();
      para.addRun(new Run('Bold text', { bold: true, color: 'FF0000' }));

      const fmt = para.getFormattingAtOffset(3);
      expect(fmt).toBeDefined();
      expect(fmt!.bold).toBe(true);
      expect(fmt!.color).toBe('FF0000');
    });

    it('returns different formatting for different runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Bold', { bold: true }));
      para.addRun(new Run('Italic', { italic: true }));

      const fmt1 = para.getFormattingAtOffset(2); // in "Bold"
      expect(fmt1!.bold).toBe(true);
      expect(fmt1!.italic).toBeUndefined();

      const fmt2 = para.getFormattingAtOffset(6); // in "Italic"
      expect(fmt2!.bold).toBeUndefined();
      expect(fmt2!.italic).toBe(true);
    });

    it('returns empty formatting for unformatted run', () => {
      const para = new Paragraph().addText('Plain text');
      const fmt = para.getFormattingAtOffset(0);

      expect(fmt).toBeDefined();
      expect(fmt!.bold).toBeUndefined();
      expect(fmt!.italic).toBeUndefined();
      expect(fmt!.color).toBeUndefined();
    });
  });

  describe('complex formatting', () => {
    it('returns full formatting object', () => {
      const para = new Paragraph();
      para.addRun(
        new Run('Styled', {
          bold: true,
          italic: true,
          underline: 'double',
          font: 'Courier New',
          size: 14,
          color: '0000FF',
          highlight: 'yellow',
        })
      );

      const fmt = para.getFormattingAtOffset(0)!;
      expect(fmt.bold).toBe(true);
      expect(fmt.italic).toBe(true);
      expect(fmt.underline).toBe('double');
      expect(fmt.font).toBe('Courier New');
      expect(fmt.size).toBe(14);
      expect(fmt.color).toBe('0000FF');
      expect(fmt.highlight).toBe('yellow');
    });
  });

  describe('edge cases', () => {
    it('returns undefined for out-of-range offset', () => {
      const para = new Paragraph().addText('Hi');
      expect(para.getFormattingAtOffset(100)).toBeUndefined();
    });

    it('returns undefined for empty paragraph', () => {
      const para = new Paragraph();
      expect(para.getFormattingAtOffset(0)).toBeUndefined();
    });

    it('returns undefined for negative offset', () => {
      const para = new Paragraph().addText('Hi');
      expect(para.getFormattingAtOffset(-1)).toBeUndefined();
    });
  });

  describe('integration with applyFormattingToRange', () => {
    it('reflects formatting applied to a range', () => {
      const para = new Paragraph().addText('Hello World');
      para.applyFormattingToRange(6, 11, { bold: true, color: 'FF0000' });

      // Before the range
      const before = para.getFormattingAtOffset(2)!;
      expect(before.bold).toBeUndefined();

      // Inside the range
      const inside = para.getFormattingAtOffset(8)!;
      expect(inside.bold).toBe(true);
      expect(inside.color).toBe('FF0000');
    });

    it('can verify formatting at split boundaries', () => {
      const para = new Paragraph().addText('AAABBBCCC');
      para.applyFormattingToRange(3, 6, { italic: true });

      // At boundary: last char before range
      expect(para.getFormattingAtOffset(2)!.italic).toBeUndefined();
      // At boundary: first char of range
      expect(para.getFormattingAtOffset(3)!.italic).toBe(true);
      // At boundary: last char of range
      expect(para.getFormattingAtOffset(5)!.italic).toBe(true);
      // At boundary: first char after range
      expect(para.getFormattingAtOffset(6)!.italic).toBeUndefined();
    });
  });

  describe('format painting use case', () => {
    it('can copy formatting from one offset to a range', () => {
      const para = new Paragraph();
      para.addRun(new Run('Source ', { bold: true, font: 'Georgia', size: 16 }));
      para.addRun(new Run('Target text'));

      // "Paint" formatting from offset 0 to the target range
      const sourceFmt = para.getFormattingAtOffset(0)!;
      para.applyFormattingToRange(7, 18, {
        bold: sourceFmt.bold,
        font: sourceFmt.font,
        size: sourceFmt.size,
      });

      // Verify target now has the same formatting
      const targetFmt = para.getFormattingAtOffset(10)!;
      expect(targetFmt.bold).toBe(true);
      expect(targetFmt.font).toBe('Georgia');
      expect(targetFmt.size).toBe(16);
    });
  });
});
