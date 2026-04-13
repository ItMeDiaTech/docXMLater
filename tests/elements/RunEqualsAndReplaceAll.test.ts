/**
 * Tests for Run.equals(), Run.hasSameFormatting(), Paragraph.replaceAll()
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';

// ============================================================================
// Run.equals()
// ============================================================================

describe('Run.equals()', () => {
  it('returns true for identical runs', () => {
    const a = new Run('Hello', { bold: true });
    const b = new Run('Hello', { bold: true });

    expect(a.equals(b)).toBe(true);
  });

  it('returns false for different text', () => {
    const a = new Run('Hello', { bold: true });
    const b = new Run('World', { bold: true });

    expect(a.equals(b)).toBe(false);
  });

  it('returns false for different formatting', () => {
    const a = new Run('Hello', { bold: true });
    const b = new Run('Hello', { italic: true });

    expect(a.equals(b)).toBe(false);
  });

  it('returns true for both unformatted', () => {
    const a = new Run('text');
    const b = new Run('text');

    expect(a.equals(b)).toBe(true);
  });

  it('returns false when one has formatting and other does not', () => {
    const a = new Run('text', { bold: true });
    const b = new Run('text');

    expect(a.equals(b)).toBe(false);
  });

  it('compares complex formatting deeply', () => {
    const fmt = { bold: true, italic: true, color: 'FF0000', font: 'Arial', size: 12 };
    const a = new Run('text', { ...fmt });
    const b = new Run('text', { ...fmt });

    expect(a.equals(b)).toBe(true);
  });

  it('detects single property difference in complex formatting', () => {
    const a = new Run('text', { bold: true, color: 'FF0000' });
    const b = new Run('text', { bold: true, color: '0000FF' });

    expect(a.equals(b)).toBe(false);
  });

  it('returns true for empty runs', () => {
    const a = new Run('');
    const b = new Run('');

    expect(a.equals(b)).toBe(true);
  });
});

// ============================================================================
// Run.hasSameFormatting()
// ============================================================================

describe('Run.hasSameFormatting()', () => {
  it('returns true when formatting matches (different text)', () => {
    const a = new Run('Hello', { bold: true });
    const b = new Run('World', { bold: true });

    expect(a.hasSameFormatting(b)).toBe(true);
  });

  it('returns false when formatting differs', () => {
    const a = new Run('text', { bold: true });
    const b = new Run('text', { italic: true });

    expect(a.hasSameFormatting(b)).toBe(false);
  });

  it('returns true for both unformatted', () => {
    const a = new Run('different');
    const b = new Run('text');

    expect(a.hasSameFormatting(b)).toBe(true);
  });

  it('practical: find runs that can be merged', () => {
    const para = new Paragraph();
    para.addRun(new Run('A', { bold: true }));
    para.addRun(new Run('B', { bold: true }));
    para.addRun(new Run('C', { italic: true }));

    const runs = para.getRuns();
    const canMerge = runs[0]!.hasSameFormatting(runs[1]!);
    const cantMerge = runs[0]!.hasSameFormatting(runs[2]!);

    expect(canMerge).toBe(true);
    expect(cantMerge).toBe(false);
  });
});

// ============================================================================
// Paragraph.replaceAll()
// ============================================================================

describe('Paragraph.replaceAll()', () => {
  it('replaces all occurrences of text', () => {
    const para = new Paragraph().addText('foo bar foo baz foo');
    const count = para.replaceAll('foo', 'qux');

    expect(count).toBe(3);
    expect(para.getText()).toBe('qux bar qux baz qux');
  });

  it('handles cross-run text', () => {
    const para = new Paragraph();
    para.addRun(new Run('{{'));
    para.addRun(new Run('name'));
    para.addRun(new Run('}}'));

    const count = para.replaceAll('{{name}}', 'Alice');

    expect(count).toBe(1);
    expect(para.getText()).toBe('Alice');
  });

  it('is case-insensitive', () => {
    const para = new Paragraph().addText('Hello HELLO hello');
    const count = para.replaceAll('hello', 'hi');

    expect(count).toBe(3);
  });

  it('returns 0 when no match', () => {
    const para = new Paragraph().addText('Nothing here');
    expect(para.replaceAll('xyz', 'abc')).toBe(0);
  });

  it('preserves surrounding text', () => {
    const para = new Paragraph().addText('The cat sat on the mat');
    para.replaceAll('cat', 'dog');

    expect(para.getText()).toBe('The dog sat on the mat');
  });

  it('handles empty replacement', () => {
    const para = new Paragraph().addText('Remove {{this}} text');
    para.replaceAll('{{this}} ', '');

    expect(para.getText()).toBe('Remove text');
  });

  it('practical: quick template fill', () => {
    const para = new Paragraph().addText('Dear {{name}}, welcome to {{company}}!');
    para.replaceAll('{{name}}', 'Alice');
    para.replaceAll('{{company}}', 'Acme Corp');

    expect(para.getText()).toBe('Dear Alice, welcome to Acme Corp!');
  });
});

// ============================================================================
// Integration
// ============================================================================

describe('Run.equals + consolidateRuns integration', () => {
  it('can identify mergeable runs then consolidate', () => {
    const para = new Paragraph();
    const r1 = new Run('Hello ', { bold: true });
    const r2 = new Run('World', { bold: true });
    const r3 = new Run('!');
    para.addRun(r1);
    para.addRun(r2);
    para.addRun(r3);

    // r1 and r2 have same formatting
    expect(r1.hasSameFormatting(r2)).toBe(true);
    expect(r1.hasSameFormatting(r3)).toBe(false);

    // Consolidate merges them
    para.consolidateRuns();
    expect(para.getRuns()).toHaveLength(2);
    expect(para.getText()).toBe('Hello World!');
  });
});
