/**
 * Tests for Paragraph.consolidateRuns()
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Hyperlink } from '../../src/elements/Hyperlink';

describe('Paragraph.consolidateRuns()', () => {
  describe('basic merging', () => {
    it('merges two adjacent runs with identical formatting', () => {
      const para = new Paragraph();
      para.addRun(new Run('Hello ', { bold: true }));
      para.addRun(new Run('World', { bold: true }));

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(1);
      const runs = para.getRuns();
      expect(runs).toHaveLength(1);
      expect(runs[0]!.getText()).toBe('Hello World');
      expect(runs[0]!.getFormatting().bold).toBe(true);
    });

    it('merges three adjacent runs with identical formatting', () => {
      const para = new Paragraph();
      para.addRun(new Run('A'));
      para.addRun(new Run('B'));
      para.addRun(new Run('C'));

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(2);
      const runs = para.getRuns();
      expect(runs).toHaveLength(1);
      expect(runs[0]!.getText()).toBe('ABC');
    });

    it('does not merge runs with different formatting', () => {
      const para = new Paragraph();
      para.addRun(new Run('Bold', { bold: true }));
      para.addRun(new Run('Normal'));
      para.addRun(new Run('Italic', { italic: true }));

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(0);
      expect(para.getRuns()).toHaveLength(3);
    });

    it('merges selectively within mixed formatting', () => {
      const para = new Paragraph();
      para.addRun(new Run('A', { bold: true }));
      para.addRun(new Run('B', { bold: true }));
      para.addRun(new Run('C'));
      para.addRun(new Run('D'));
      para.addRun(new Run('E', { italic: true }));

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(2);
      const runs = para.getRuns();
      expect(runs).toHaveLength(3);
      expect(runs[0]!.getText()).toBe('AB');
      expect(runs[0]!.getFormatting().bold).toBe(true);
      expect(runs[1]!.getText()).toBe('CD');
      expect(runs[2]!.getText()).toBe('E');
      expect(runs[2]!.getFormatting().italic).toBe(true);
    });
  });

  describe('complex formatting comparison', () => {
    it('merges runs with matching multi-property formatting', () => {
      const fmt = { bold: true, italic: true, color: 'FF0000', font: 'Arial', size: 12 };
      const para = new Paragraph();
      para.addRun(new Run('Part 1', fmt));
      para.addRun(new Run('Part 2', fmt));

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(1);
      const run = para.getRuns()[0]!;
      expect(run.getText()).toBe('Part 1Part 2');
      expect(run.getFormatting().bold).toBe(true);
      expect(run.getFormatting().color).toBe('FF0000');
      expect(run.getFormatting().font).toBe('Arial');
    });

    it('does not merge when only one property differs', () => {
      const para = new Paragraph();
      para.addRun(new Run('A', { bold: true, color: 'FF0000' }));
      para.addRun(new Run('B', { bold: true, color: '0000FF' }));

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(0);
      expect(para.getRuns()).toHaveLength(2);
    });
  });

  describe('non-Run content boundaries', () => {
    it('does not merge runs across hyperlinks', () => {
      const para = new Paragraph();
      para.addRun(new Run('Before'));
      para.addHyperlink(new Hyperlink({ url: 'https://example.com', text: 'Link' }));
      para.addRun(new Run('After'));

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(0);
      const runs = para.getRuns();
      // getRuns() also extracts runs from hyperlinks
      expect(runs.length).toBeGreaterThanOrEqual(2);
    });
  });

  describe('edge cases', () => {
    it('handles empty paragraph', () => {
      const para = new Paragraph();
      expect(para.consolidateRuns()).toBe(0);
    });

    it('handles single run', () => {
      const para = new Paragraph().addText('Solo');
      expect(para.consolidateRuns()).toBe(0);
      expect(para.getRuns()).toHaveLength(1);
    });

    it('handles runs with empty text', () => {
      const para = new Paragraph();
      para.addRun(new Run(''));
      para.addRun(new Run(''));

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(1);
      expect(para.getRuns()).toHaveLength(1);
    });

    it('preserves text content after consolidation', () => {
      const para = new Paragraph();
      para.addRun(new Run('The '));
      para.addRun(new Run('quick '));
      para.addRun(new Run('brown '));
      para.addRun(new Run('fox'));

      para.consolidateRuns();

      expect(para.getText()).toBe('The quick brown fox');
    });
  });

  describe('integration with applyFormattingToRange', () => {
    it('consolidates after formatting adjacent ranges identically', () => {
      const para = new Paragraph().addText('Hello World');

      // Apply bold to "Hello" and " World" separately (creates fragments)
      para.applyFormattingToRange(0, 5, { bold: true });
      para.applyFormattingToRange(5, 11, { bold: true });

      // Before consolidation: 2 bold runs
      expect(para.getRuns().length).toBeGreaterThanOrEqual(2);

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBeGreaterThanOrEqual(1);
      expect(para.getText()).toBe('Hello World');
      const runs = para.getRuns();
      expect(runs).toHaveLength(1);
      expect(runs[0]!.getFormatting().bold).toBe(true);
    });

    it('does not consolidate differently-formatted ranges', () => {
      const para = new Paragraph().addText('Hello World');

      para.applyFormattingToRange(0, 5, { bold: true });
      para.applyFormattingToRange(6, 11, { italic: true });

      para.consolidateRuns();

      const runs = para.getRuns();
      expect(runs.length).toBeGreaterThanOrEqual(3);
      // "Hello" (bold), " " (plain), "World" (italic)
    });

    it('handles three-way split followed by consolidation', () => {
      const para = new Paragraph().addText('AAABBBCCC');

      // Bold the middle "BBB"
      para.applyFormattingToRange(3, 6, { bold: true });

      // Now un-bold it (making all runs have same formatting again)
      para.applyFormattingToRange(3, 6, { bold: false });

      // All three fragments should now have bold=false or undefined
      // After consolidation, should merge back
      const beforeRuns = para.getRuns().length;
      para.consolidateRuns();
      const afterRuns = para.getRuns().length;

      expect(afterRuns).toBeLessThanOrEqual(beforeRuns);
      expect(para.getText()).toBe('AAABBBCCC');
    });
  });

  describe('special run content', () => {
    it('merges runs with tabs preserving content structure', () => {
      const run1 = new Run('');
      run1.appendText('A');
      run1.addTab();

      const run2 = new Run('');
      run2.appendText('B');

      const para = new Paragraph();
      para.addRun(run1);
      para.addRun(run2);

      const eliminated = para.consolidateRuns();

      expect(eliminated).toBe(1);
      const merged = para.getRuns()[0]!;
      expect(merged.getText()).toBe('A\tB');
      // Content includes tab element between text segments
      const content = merged.getContent();
      expect(content.some((c) => c.type === 'tab')).toBe(true);
    });
  });

  describe('XML validity', () => {
    it('produces valid XML after consolidation', () => {
      const para = new Paragraph();
      para.addRun(new Run('A'));
      para.addRun(new Run('B'));
      para.addRun(new Run('C'));

      para.consolidateRuns();

      const xml = para.toXML();
      expect(xml.name).toBe('w:p');
      const runs = xml.children?.filter((c) => typeof c !== 'string' && c.name === 'w:r');
      expect(runs).toHaveLength(1);
    });
  });
});
