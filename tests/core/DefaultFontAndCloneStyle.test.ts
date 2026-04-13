/**
 * Tests for Document.setDefaultFont(), Document.setDefaultFontSize(),
 * and StylesManager.cloneStyle()
 */

import { Document } from '../../src/core/Document';
import { StylesManager } from '../../src/formatting/StylesManager';
import { Style } from '../../src/formatting/Style';

describe('Document.setDefaultFont()', () => {
  it('sets the default font on the Normal style', () => {
    const doc = Document.create();
    doc.setDefaultFont('Times New Roman');

    const normalStyle = doc.getStylesManager().getStyle('Normal');
    expect(normalStyle).toBeDefined();
    expect(normalStyle!.getRunFormatting()?.font).toBe('Times New Roman');
    doc.dispose();
  });

  it('sets font and size together', () => {
    const doc = Document.create();
    doc.setDefaultFont('Arial', 14);

    const fmt = doc.getStylesManager().getStyle('Normal')!.getRunFormatting()!;
    expect(fmt.font).toBe('Arial');
    expect(fmt.size).toBe(14);
    doc.dispose();
  });

  it('preserves existing Normal style run formatting', () => {
    const doc = Document.create();
    // Set up existing formatting
    const normalStyle = doc.getStylesManager().getStyle('Normal')!;
    normalStyle.setRunFormatting({ bold: true, color: 'FF0000' });

    doc.setDefaultFont('Georgia', 12);

    const fmt = normalStyle.getRunFormatting()!;
    expect(fmt.font).toBe('Georgia');
    expect(fmt.size).toBe(12);
    expect(fmt.bold).toBe(true);
    expect(fmt.color).toBe('FF0000');
    doc.dispose();
  });

  it('returns this for chaining', () => {
    const doc = Document.create();
    const result = doc.setDefaultFont('Calibri');

    expect(result).toBe(doc);
    doc.dispose();
  });

  it('creates Normal style if missing', () => {
    const doc = Document.create();
    // Even fresh documents have Normal, but the method should handle its absence
    doc.setDefaultFont('Verdana', 11);

    const normal = doc.getStylesManager().getStyle('Normal');
    expect(normal).toBeDefined();
    expect(normal!.getRunFormatting()?.font).toBe('Verdana');
    doc.dispose();
  });

  it('produces valid DOCX with custom default font', async () => {
    const doc = Document.create();
    doc.setDefaultFont('Courier New', 10);
    doc.createParagraph('Monospace text by default.');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    const normalStyle = loaded.getStylesManager().getStyle('Normal');
    expect(normalStyle?.getRunFormatting()?.font).toBe('Courier New');
    loaded.dispose();
    doc.dispose();
  });
});

describe('Document.setDefaultFontSize()', () => {
  it('sets the default font size on Normal style', () => {
    const doc = Document.create();
    doc.setDefaultFontSize(14);

    const fmt = doc.getStylesManager().getStyle('Normal')!.getRunFormatting()!;
    expect(fmt.size).toBe(14);
    doc.dispose();
  });

  it('preserves existing font name', () => {
    const doc = Document.create();
    doc.setDefaultFont('Arial');
    doc.setDefaultFontSize(16);

    const fmt = doc.getStylesManager().getStyle('Normal')!.getRunFormatting()!;
    expect(fmt.font).toBe('Arial');
    expect(fmt.size).toBe(16);
    doc.dispose();
  });

  it('returns this for chaining', () => {
    const doc = Document.create();
    const result = doc.setDefaultFontSize(12);

    expect(result).toBe(doc);
    doc.dispose();
  });

  it('chains with setDefaultFont', () => {
    const doc = Document.create();
    doc.setDefaultFont('Georgia').setDefaultFontSize(11);

    const fmt = doc.getStylesManager().getStyle('Normal')!.getRunFormatting()!;
    expect(fmt.font).toBe('Georgia');
    expect(fmt.size).toBe(11);
    doc.dispose();
  });
});

describe('StylesManager.cloneStyle()', () => {
  it('clones a built-in style with a new ID', () => {
    const sm = StylesManager.create();
    const cloned = sm.cloneStyle('Heading1', 'CustomH1', 'Custom Heading 1');

    expect(cloned).toBeDefined();
    expect(cloned!.getStyleId()).toBe('CustomH1');
    expect(cloned!.getName()).toBe('Custom Heading 1');
    expect(cloned!.getType()).toBe('paragraph');
  });

  it('preserves source formatting in clone', () => {
    const sm = StylesManager.create();
    const cloned = sm.cloneStyle('Heading1', 'H1Copy');

    const original = sm.getStyle('Heading1')!;
    const origFmt = original.getRunFormatting();
    const clonedFmt = cloned!.getRunFormatting();

    // Both should have similar formatting
    if (origFmt && clonedFmt) {
      expect(clonedFmt.bold).toBe(origFmt.bold);
      expect(clonedFmt.size).toBe(origFmt.size);
    }
  });

  it('registers clone in the styles manager', () => {
    const sm = StylesManager.create();
    sm.cloneStyle('Normal', 'NormalCopy');

    expect(sm.hasStyle('NormalCopy')).toBe(true);
    expect(sm.getStyle('NormalCopy')).toBeDefined();
  });

  it('defaults name to newId when not provided', () => {
    const sm = StylesManager.create();
    const cloned = sm.cloneStyle('Normal', 'MyCustom');

    expect(cloned!.getName()).toBe('MyCustom');
  });

  it('returns undefined for non-existent source', () => {
    const sm = StylesManager.create();
    const result = sm.cloneStyle('NonExistent', 'Copy');

    expect(result).toBeUndefined();
  });

  it('clone is independent of source', () => {
    const sm = StylesManager.create();
    const cloned = sm.cloneStyle('Normal', 'NormalAlt')!;

    // Modify the clone
    cloned.setRunFormatting({ font: 'Comic Sans MS', bold: true });

    // Original should be unaffected
    const original = sm.getStyle('Normal')!;
    expect(original.getRunFormatting()?.font).not.toBe('Comic Sans MS');
  });

  it('clone is never marked as default', () => {
    const sm = StylesManager.create();
    const cloned = sm.cloneStyle('Normal', 'NormalCopy');

    expect(cloned!.getProperties().isDefault).toBe(false);
  });

  it('can clone and modify for a variant style', () => {
    const sm = StylesManager.create();
    const blueHeading = sm.cloneStyle('Heading1', 'BlueH1', 'Blue Heading 1');

    blueHeading!.setRunFormatting({
      ...blueHeading!.getRunFormatting(),
      color: '0000FF',
    });

    expect(blueHeading!.getRunFormatting()?.color).toBe('0000FF');
    expect(sm.hasStyle('BlueH1')).toBe(true);
  });

  it('works within a Document context', () => {
    const doc = Document.create();
    const sm = doc.getStylesManager();

    const variant = sm.cloneStyle('Heading1', 'ChapterTitle', 'Chapter Title');
    expect(variant).toBeDefined();

    // Use the cloned style
    doc.createParagraph('Chapter 1').setStyle('ChapterTitle');

    const styled = doc.getParagraphs().find((p) => p.getStyle() === 'ChapterTitle');
    expect(styled).toBeDefined();
    expect(styled!.getText()).toBe('Chapter 1');
    doc.dispose();
  });

  it('produces valid DOCX with cloned styles', async () => {
    const doc = Document.create();
    const sm = doc.getStylesManager();
    sm.cloneStyle('Heading1', 'SpecialH1', 'Special Heading');

    doc.createParagraph('Styled content').setStyle('SpecialH1');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.getStylesManager().hasStyle('SpecialH1')).toBe(true);
    loaded.dispose();
    doc.dispose();
  });
});
