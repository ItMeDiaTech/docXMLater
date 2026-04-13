/**
 * Tests for Document.clone() and Paragraph.findTextCrossRun()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';

// ============================================================================
// Document.clone()
// ============================================================================

describe('Document.clone()', () => {
  it('creates an independent copy of the document', async () => {
    const original = Document.create();
    original.createParagraph('Original content');

    const clone = await original.clone();

    expect(clone.toPlainText()).toContain('Original content');

    // Modify clone — original should be unaffected
    clone.createParagraph('Added to clone');
    expect(original.getParagraphs().length).toBeLessThan(clone.getParagraphs().length);

    clone.dispose();
    original.dispose();
  });

  it('preserves headings and styles', async () => {
    const original = Document.create();
    original.addHeading('Title', 1);
    original.addHeading('Section', 2);
    original.createParagraph('Body text.');

    const clone = await original.clone();

    expect(clone.getStatistics().headings).toBe(2);
    expect(clone.toPlainText()).toContain('Body text.');

    clone.dispose();
    original.dispose();
  });

  it('preserves tables', async () => {
    const original = Document.create();
    original.addTable(
      Table.fromArray([
        ['Name', 'Value'],
        ['Alpha', '100'],
      ])
    );

    const clone = await original.clone();

    expect(clone.getTables()).toHaveLength(1);
    expect(clone.getTables()[0]!.getCell(1, 0)!.getText()).toBe('Alpha');

    clone.dispose();
    original.dispose();
  });

  it('preserves default font settings', async () => {
    const original = Document.create();
    original.setDefaultFont('Georgia', 14);
    original.createParagraph('Styled text');

    const clone = await original.clone();

    const normalStyle = clone.getStylesManager().getStyle('Normal');
    expect(normalStyle?.getRunFormatting()?.font).toBe('Georgia');
    expect(normalStyle?.getRunFormatting()?.size).toBe(14);

    clone.dispose();
    original.dispose();
  });

  it('clone is fully independent — mutations do not cross', async () => {
    const original = Document.create();
    original.createParagraph('Shared');

    const clone = await original.clone();

    // Modify original after cloning
    original.createParagraph('Only in original');

    // Clone should not see the new paragraph
    expect(clone.toPlainText()).not.toContain('Only in original');

    clone.dispose();
    original.dispose();
  });

  it('enables template-based batch generation', async () => {
    const template = Document.create();
    template.addHeading('{{title}}', 1);
    template.createParagraph('Dear {{name}}, your order {{orderId}} is ready.');

    const records = [
      { title: 'Order Confirmation', name: 'Alice', orderId: 'A001' },
      { title: 'Order Confirmation', name: 'Bob', orderId: 'B002' },
      { title: 'Order Confirmation', name: 'Charlie', orderId: 'C003' },
    ];

    const docs: Document[] = [];
    for (const record of records) {
      const doc = await template.clone();
      doc.fillTemplate(record);
      docs.push(doc);
    }

    // Each document should have its own data
    expect(docs[0]!.toPlainText()).toContain('Alice');
    expect(docs[0]!.toPlainText()).toContain('A001');
    expect(docs[1]!.toPlainText()).toContain('Bob');
    expect(docs[2]!.toPlainText()).toContain('Charlie');

    // Template unchanged
    expect(template.toPlainText()).toContain('{{name}}');

    for (const doc of docs) doc.dispose();
    template.dispose();
  });

  it('cloned document can be saved', async () => {
    const original = Document.create();
    original.createParagraph('Saveable clone');

    const clone = await original.clone();
    const buffer = await clone.toBuffer();

    expect(buffer.length).toBeGreaterThan(0);

    // Verify the saved clone loads correctly
    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.toPlainText()).toContain('Saveable clone');

    loaded.dispose();
    clone.dispose();
    original.dispose();
  });

  it('handles empty document', async () => {
    const original = Document.create();
    const clone = await original.clone();

    expect(clone.getParagraphs()).toHaveLength(0);

    clone.dispose();
    original.dispose();
  });
});

// ============================================================================
// Paragraph.findTextCrossRun()
// ============================================================================

describe('Paragraph.findTextCrossRun()', () => {
  describe('single run (non-fragmented)', () => {
    it('finds text within a single run', () => {
      const para = new Paragraph().addText('Hello World');
      const matches = para.findTextCrossRun('World');

      expect(matches).toHaveLength(1);
      expect(matches[0]!.offset).toBe(6);
      expect(matches[0]!.text).toBe('World');
    });

    it('finds multiple occurrences', () => {
      const para = new Paragraph().addText('foo bar foo baz foo');
      const matches = para.findTextCrossRun('foo');

      expect(matches).toHaveLength(3);
      expect(matches[0]!.offset).toBe(0);
      expect(matches[1]!.offset).toBe(8);
      expect(matches[2]!.offset).toBe(16);
    });

    it('is case-insensitive by default', () => {
      const para = new Paragraph().addText('Hello HELLO hello');
      const matches = para.findTextCrossRun('hello');

      expect(matches).toHaveLength(3);
    });

    it('supports case-sensitive search', () => {
      const para = new Paragraph().addText('Hello HELLO hello');
      const matches = para.findTextCrossRun('Hello', { caseSensitive: true });

      expect(matches).toHaveLength(1);
      expect(matches[0]!.offset).toBe(0);
    });

    it('returns empty array when no match', () => {
      const para = new Paragraph().addText('Hello World');
      expect(para.findTextCrossRun('xyz')).toHaveLength(0);
    });
  });

  describe('cross-run matching', () => {
    it('finds text spanning two runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Hel'));
      para.addRun(new Run('lo'));

      const matches = para.findTextCrossRun('Hello');

      expect(matches).toHaveLength(1);
      expect(matches[0]!.offset).toBe(0);
    });

    it('finds placeholder split across three runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('{{'));
      para.addRun(new Run('name'));
      para.addRun(new Run('}}'));

      const matches = para.findTextCrossRun('{{name}}');

      expect(matches).toHaveLength(1);
      expect(matches[0]!.offset).toBe(0);
      expect(matches[0]!.text).toBe('{{name}}');
    });

    it('finds text with surrounding content in boundary runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Dear {{'));
      para.addRun(new Run('name'));
      para.addRun(new Run('}}, welcome!'));

      const matches = para.findTextCrossRun('{{name}}');

      expect(matches).toHaveLength(1);
      expect(matches[0]!.offset).toBe(5);
    });

    it('finds multiple fragmented placeholders', () => {
      const para = new Paragraph();
      para.addRun(new Run('{{'));
      para.addRun(new Run('first'));
      para.addRun(new Run('}} and {{'));
      para.addRun(new Run('second'));
      para.addRun(new Run('}}'));

      const firstMatches = para.findTextCrossRun('{{first}}');
      const secondMatches = para.findTextCrossRun('{{second}}');

      expect(firstMatches).toHaveLength(1);
      expect(secondMatches).toHaveLength(1);
    });
  });

  describe('edge cases', () => {
    it('handles empty paragraph', () => {
      const para = new Paragraph();
      expect(para.findTextCrossRun('test')).toHaveLength(0);
    });

    it('handles special regex characters in search text', () => {
      const para = new Paragraph().addText('Price: $10.00 (USD)');
      const matches = para.findTextCrossRun('$10.00 (USD)');

      expect(matches).toHaveLength(1);
    });

    it('returns correct text with case-insensitive match', () => {
      const para = new Paragraph().addText('Hello WORLD');
      const matches = para.findTextCrossRun('world');

      expect(matches[0]!.text).toBe('WORLD');
    });
  });

  describe('practical usage', () => {
    it('can list all placeholders in a paragraph', () => {
      const para = new Paragraph();
      para.addRun(new Run('Hello {{'));
      para.addRun(new Run('name'));
      para.addRun(new Run('}}, your order {{orderId}} is ready.'));

      // Find all {{...}} patterns using findTextCrossRun for specific keys
      const nameMatches = para.findTextCrossRun('{{name}}');
      const orderMatches = para.findTextCrossRun('{{orderId}}');

      expect(nameMatches).toHaveLength(1);
      expect(orderMatches).toHaveLength(1);
    });

    it('pairs with replaceTextCrossRun for preview-then-replace', () => {
      const para = new Paragraph().addText('Replace {{token}} here');

      // Preview: check if token exists
      const preview = para.findTextCrossRun('{{token}}');
      expect(preview).toHaveLength(1);

      // Replace
      para.replaceTextCrossRun('{{token}}', 'VALUE');
      expect(para.getText()).toBe('Replace VALUE here');

      // Verify it's gone
      expect(para.findTextCrossRun('{{token}}')).toHaveLength(0);
    });
  });
});
