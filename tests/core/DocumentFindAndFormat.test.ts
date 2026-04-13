/**
 * Tests for Document.findAndHighlight() and Document.findAndFormat()
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { Run } from '../../src/elements/Run';

describe('Document.findAndHighlight()', () => {
  it('highlights all occurrences in yellow by default', () => {
    const doc = Document.create();
    doc.createParagraph('This is important. Very important indeed.');

    const count = doc.findAndHighlight('important');

    expect(count).toBe(2);

    // Check that matching runs have yellow highlight
    const runs = doc.getParagraphs()[0]!.getRuns();
    const highlighted = runs.filter((r) => r.getFormatting().highlight === 'yellow');
    expect(highlighted).toHaveLength(2);
    expect(highlighted[0]!.getText()).toBe('important');

    doc.dispose();
  });

  it('uses custom highlight color', () => {
    const doc = Document.create();
    doc.createParagraph('Error found in line 5.');

    doc.findAndHighlight('Error', 'red');

    const runs = doc.getParagraphs()[0]!.getRuns();
    const redRun = runs.find((r) => r.getFormatting().highlight === 'red');
    expect(redRun).toBeDefined();
    expect(redRun!.getText()).toBe('Error');

    doc.dispose();
  });

  it('supports case-sensitive search', () => {
    const doc = Document.create();
    doc.createParagraph('Error error ERROR');

    const count = doc.findAndHighlight('Error', 'yellow', { caseSensitive: true });

    expect(count).toBe(1);
    doc.dispose();
  });

  it('is case-insensitive by default', () => {
    const doc = Document.create();
    doc.createParagraph('Hello HELLO hello');

    const count = doc.findAndHighlight('hello');

    expect(count).toBe(3);
    doc.dispose();
  });

  it('highlights across multiple paragraphs', () => {
    const doc = Document.create();
    doc.createParagraph('First important paragraph.');
    doc.createParagraph('Second paragraph.');
    doc.createParagraph('Third important paragraph.');

    const count = doc.findAndHighlight('important');

    expect(count).toBe(2);
    doc.dispose();
  });

  it('highlights text inside table cells', () => {
    const doc = Document.create();
    doc.addTable(
      Table.fromArray([
        ['Name', 'Status'],
        ['Alpha', 'active'],
        ['Beta', 'active'],
        ['Gamma', 'inactive'],
      ])
    );

    const count = doc.findAndHighlight('active', 'green');

    // "active" appears in rows 1, 2, 3 (also within "inactive")
    expect(count).toBeGreaterThanOrEqual(3);
    doc.dispose();
  });

  it('returns 0 when text not found', () => {
    const doc = Document.create();
    doc.createParagraph('Hello World');

    expect(doc.findAndHighlight('xyz')).toBe(0);
    doc.dispose();
  });

  it('handles cross-run fragmented text', () => {
    const doc = Document.create();
    const para = doc.createParagraph();
    para.addRun(new Run('imp'));
    para.addRun(new Run('ortant'));

    const count = doc.findAndHighlight('important');

    expect(count).toBe(1);
    doc.dispose();
  });

  it('preserves surrounding text', () => {
    const doc = Document.create();
    doc.createParagraph('The word important is key.');

    doc.findAndHighlight('important');

    expect(doc.getParagraphs()[0]!.getText()).toBe('The word important is key.');
    doc.dispose();
  });

  it('produces valid DOCX', async () => {
    const doc = Document.create();
    doc.createParagraph('Highlight this word and this word too.');
    doc.findAndHighlight('word', 'cyan');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.toPlainText()).toContain('Highlight this word');
    loaded.dispose();
    doc.dispose();
  });
});

describe('Document.findAndFormat()', () => {
  it('applies bold to all matches', () => {
    const doc = Document.create();
    doc.createParagraph('Bold this and bold that.');

    const count = doc.findAndFormat('bold', { bold: true });

    expect(count).toBe(2);

    const runs = doc.getParagraphs()[0]!.getRuns();
    const boldRuns = runs.filter((r) => r.getFormatting().bold);
    expect(boldRuns).toHaveLength(2);
    doc.dispose();
  });

  it('applies multiple formatting properties', () => {
    const doc = Document.create();
    doc.createParagraph('Mark this critical term clearly.');

    doc.findAndFormat('critical', {
      bold: true,
      color: 'FF0000',
      underline: 'single',
    });

    const runs = doc.getParagraphs()[0]!.getRuns();
    const formatted = runs.find((r) => r.getText() === 'critical');
    expect(formatted).toBeDefined();
    expect(formatted!.getFormatting().bold).toBe(true);
    expect(formatted!.getFormatting().color).toBe('FF0000');
    expect(formatted!.getFormatting().underline).toBe('single');
    doc.dispose();
  });

  it('applies strikethrough for deprecated terms', () => {
    const doc = Document.create();
    doc.createParagraph('The deprecated function is still used.');

    doc.findAndFormat('deprecated', { strike: true, color: '888888' });

    const runs = doc.getParagraphs()[0]!.getRuns();
    const struck = runs.find((r) => r.getFormatting().strike);
    expect(struck).toBeDefined();
    expect(struck!.getText()).toBe('deprecated');
    doc.dispose();
  });

  it('applies italic to specific terms', () => {
    const doc = Document.create();
    doc.createParagraph('The genus Homo sapiens is remarkable.');

    doc.findAndFormat('Homo sapiens', { italic: true });

    const runs = doc.getParagraphs()[0]!.getRuns();
    const italicRun = runs.find((r) => r.getFormatting().italic && r.getText().includes('Homo'));
    expect(italicRun).toBeDefined();
    doc.dispose();
  });

  it('formats across multiple paragraphs and tables', () => {
    const doc = Document.create();
    doc.createParagraph('Status: active');
    doc.addTable(
      Table.fromArray([
        ['Item', 'Status'],
        ['A', 'active'],
      ])
    );
    doc.createParagraph('All items are active.');

    const count = doc.findAndFormat('active', { bold: true, color: '00AA00' });

    expect(count).toBeGreaterThanOrEqual(3);
    doc.dispose();
  });

  it('supports case-sensitive formatting', () => {
    const doc = Document.create();
    doc.createParagraph('TODO: Fix this. todo: not this.');

    const count = doc.findAndFormat(
      'TODO',
      { bold: true, highlight: 'yellow' },
      { caseSensitive: true }
    );

    expect(count).toBe(1);
    doc.dispose();
  });

  it('returns 0 for no matches', () => {
    const doc = Document.create();
    doc.createParagraph('Nothing here.');

    expect(doc.findAndFormat('xyz', { bold: true })).toBe(0);
    doc.dispose();
  });

  it('preserves total text content', () => {
    const doc = Document.create();
    const original = 'The quick brown fox jumps over the lazy dog.';
    doc.createParagraph(original);

    doc.findAndFormat('fox', { bold: true });
    doc.findAndFormat('dog', { italic: true });

    expect(doc.getParagraphs()[0]!.getText()).toBe(original);
    doc.dispose();
  });
});

describe('combined findAndHighlight + findAndFormat', () => {
  it('can apply different colors to different terms', () => {
    const doc = Document.create();
    doc.createParagraph('Warning: potential error detected. Info: system normal.');

    doc.findAndHighlight('Warning', 'yellow');
    doc.findAndHighlight('error', 'red');
    doc.findAndFormat('Info', { bold: true, color: '0000FF' });

    // Text preserved
    const text = doc.getParagraphs()[0]!.getText();
    expect(text).toContain('Warning');
    expect(text).toContain('error');
    expect(text).toContain('Info');
    doc.dispose();
  });

  it('practical: document review annotation', () => {
    const doc = Document.create();
    doc.addHeading('Contract Review', 1);
    doc.createParagraph('The party shall indemnify all claims.');
    doc.createParagraph('Payment terms: net 30 days.');
    doc.createParagraph('The party assumes all liability for damages.');

    // Highlight legal keywords
    doc.findAndHighlight('shall', 'cyan');
    doc.findAndHighlight('liability', 'yellow');
    doc.findAndFormat('indemnify', { bold: true, underline: 'single' });

    const stats = doc.getStatistics();
    expect(stats.paragraphs).toBeGreaterThanOrEqual(3);
    doc.dispose();
  });
});
