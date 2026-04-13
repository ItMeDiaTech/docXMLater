/**
 * Tests for Document.addBulletListFromArray() and Document.addNumberedListFromArray()
 */

import { Document } from '../../src/core/Document';

describe('Document.addBulletListFromArray()', () => {
  it('creates a bullet list from string array', () => {
    const doc = Document.create();
    const paras = doc.addBulletListFromArray(['First', 'Second', 'Third']);

    expect(paras).toHaveLength(3);
    expect(paras[0]!.getText()).toBe('First');
    expect(paras[1]!.getText()).toBe('Second');
    expect(paras[2]!.getText()).toBe('Third');

    // All should have numbering
    for (const para of paras) {
      expect(para.hasNumbering()).toBe(true);
    }
    doc.dispose();
  });

  it('supports nested items with level objects', () => {
    const doc = Document.create();
    const paras = doc.addBulletListFromArray([
      'Top level',
      { text: 'Nested', level: 1 },
      { text: 'Deep nested', level: 2 },
      'Back to top',
    ]);

    expect(paras).toHaveLength(4);
    expect(paras[0]!.getText()).toBe('Top level');
    expect(paras[1]!.getText()).toBe('Nested');
    expect(paras[2]!.getText()).toBe('Deep nested');
    expect(paras[3]!.getText()).toBe('Back to top');

    // All have numbering
    for (const para of paras) {
      expect(para.hasNumbering()).toBe(true);
    }
    doc.dispose();
  });

  it('applies optional formatting to all items', () => {
    const doc = Document.create();
    const paras = doc.addBulletListFromArray(['Bold item', 'Another bold'], { bold: true });

    for (const para of paras) {
      const runs = para.getRuns();
      expect(runs[0]!.getFormatting().bold).toBe(true);
    }
    doc.dispose();
  });

  it('returns empty array for empty input', () => {
    const doc = Document.create();
    const paras = doc.addBulletListFromArray([]);

    expect(paras).toHaveLength(0);
    doc.dispose();
  });

  it('appends paragraphs to the document body', () => {
    const doc = Document.create();
    doc.createParagraph('Before list');
    doc.addBulletListFromArray(['Item A', 'Item B']);
    doc.createParagraph('After list');

    const allParas = doc.getParagraphs();
    expect(allParas.length).toBeGreaterThanOrEqual(4);
    doc.dispose();
  });

  it('all items share the same numId', () => {
    const doc = Document.create();
    const paras = doc.addBulletListFromArray(['A', 'B', 'C']);

    const numIds = paras.map((p) => p.getFormatting().numbering?.numId);
    expect(numIds[0]).toBeDefined();
    expect(numIds.every((id) => id === numIds[0])).toBe(true);
    doc.dispose();
  });

  it('defaults level to 0 for string items', () => {
    const doc = Document.create();
    const paras = doc.addBulletListFromArray(['Item']);

    expect(paras[0]!.getFormatting().numbering?.level).toBe(0);
    doc.dispose();
  });

  it('defaults level to 0 for objects without level', () => {
    const doc = Document.create();
    const paras = doc.addBulletListFromArray([{ text: 'Item' }]);

    expect(paras[0]!.getFormatting().numbering?.level).toBe(0);
    doc.dispose();
  });

  it('generates valid DOCX', async () => {
    const doc = Document.create();
    doc.addHeading('Shopping List', 1);
    doc.addBulletListFromArray(['Apples', 'Bananas', 'Oranges']);

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    const text = loaded.toPlainText();
    expect(text).toContain('Apples');
    expect(text).toContain('Bananas');
    loaded.dispose();
    doc.dispose();
  });
});

describe('Document.addNumberedListFromArray()', () => {
  it('creates a numbered list from string array', () => {
    const doc = Document.create();
    const paras = doc.addNumberedListFromArray(['Step 1', 'Step 2', 'Step 3']);

    expect(paras).toHaveLength(3);
    expect(paras[0]!.getText()).toBe('Step 1');
    expect(paras[2]!.getText()).toBe('Step 3');

    for (const para of paras) {
      expect(para.hasNumbering()).toBe(true);
    }
    doc.dispose();
  });

  it('supports nested items', () => {
    const doc = Document.create();
    const paras = doc.addNumberedListFromArray([
      'Chapter 1',
      { text: 'Section 1.1', level: 1 },
      { text: 'Section 1.2', level: 1 },
      'Chapter 2',
      { text: 'Section 2.1', level: 1 },
    ]);

    expect(paras).toHaveLength(5);
    doc.dispose();
  });

  it('applies optional formatting', () => {
    const doc = Document.create();
    const paras = doc.addNumberedListFromArray(['Italic step'], { italic: true });

    expect(paras[0]!.getRuns()[0]!.getFormatting().italic).toBe(true);
    doc.dispose();
  });

  it('returns empty array for empty input', () => {
    const doc = Document.create();
    expect(doc.addNumberedListFromArray([])).toHaveLength(0);
    doc.dispose();
  });

  it('all items share the same numId', () => {
    const doc = Document.create();
    const paras = doc.addNumberedListFromArray(['A', 'B', 'C']);

    const numIds = paras.map((p) => p.getFormatting().numbering?.numId);
    expect(numIds[0]).toBeDefined();
    expect(numIds.every((id) => id === numIds[0])).toBe(true);
    doc.dispose();
  });

  it('generates valid DOCX', async () => {
    const doc = Document.create();
    doc.addHeading('Instructions', 1);
    doc.addNumberedListFromArray(['Open the application', 'Click Settings', 'Configure options']);

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    const text = loaded.toPlainText();
    expect(text).toContain('Open the application');
    expect(text).toContain('Configure options');
    loaded.dispose();
    doc.dispose();
  });
});

describe('bullet and numbered lists together', () => {
  it('creates independent lists with separate numIds', () => {
    const doc = Document.create();
    const bullets = doc.addBulletListFromArray(['Bullet A', 'Bullet B']);
    const numbers = doc.addNumberedListFromArray(['Number 1', 'Number 2']);

    const bulletNumId = bullets[0]!.getFormatting().numbering?.numId;
    const numberNumId = numbers[0]!.getFormatting().numbering?.numId;

    expect(bulletNumId).toBeDefined();
    expect(numberNumId).toBeDefined();
    expect(bulletNumId).not.toBe(numberNumId);
    doc.dispose();
  });

  it('renders correctly in toMarkdown output', () => {
    const doc = Document.create();
    doc.addHeading('Lists', 1);
    doc.addBulletListFromArray(['Apple', 'Banana']);
    doc.addNumberedListFromArray(['Step 1', 'Step 2']);

    // Lists should appear in the document content
    const text = doc.toPlainText();
    expect(text).toContain('Apple');
    expect(text).toContain('Step 1');
    doc.dispose();
  });

  it('full document with mixed content', async () => {
    const doc = Document.create();
    doc.setDefaultFont('Calibri', 11);
    doc.addHeading('Project Plan', 1);
    doc.createParagraph('This document outlines the project plan.');

    doc.addHeading('Goals', 2);
    doc.addBulletListFromArray([
      'Increase revenue by 20%',
      'Reduce costs by 10%',
      'Improve customer satisfaction',
    ]);

    doc.addHeading('Steps', 2);
    doc.addNumberedListFromArray([
      'Research market trends',
      { text: 'Analyze competitors', level: 1 },
      { text: 'Review pricing', level: 1 },
      'Develop strategy',
      'Execute plan',
    ]);

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.getStatistics().paragraphs).toBeGreaterThanOrEqual(10);
    loaded.dispose();
    doc.dispose();
  });
});
