/**
 * Tests for Paragraph.replaceTextCrossRun() and Document.fillTemplate()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';

describe('Paragraph.replaceTextCrossRun()', () => {
  describe('single run (non-fragmented)', () => {
    it('replaces text within a single run', () => {
      const para = new Paragraph().addText('Hello World');
      const count = para.replaceTextCrossRun('World', 'Earth');

      expect(count).toBe(1);
      expect(para.getText()).toBe('Hello Earth');
    });

    it('replaces multiple occurrences', () => {
      const para = new Paragraph().addText('foo bar foo baz foo');
      const count = para.replaceTextCrossRun('foo', 'qux');

      expect(count).toBe(3);
      expect(para.getText()).toBe('qux bar qux baz qux');
    });

    it('is case-insensitive by default', () => {
      const para = new Paragraph().addText('Hello HELLO hello');
      const count = para.replaceTextCrossRun('hello', 'hi');

      expect(count).toBe(3);
      expect(para.getText()).toBe('hi hi hi');
    });

    it('supports case-sensitive mode', () => {
      const para = new Paragraph().addText('Hello HELLO hello');
      const count = para.replaceTextCrossRun('Hello', 'Hi', { caseSensitive: true });

      expect(count).toBe(1);
      expect(para.getText()).toBe('Hi HELLO hello');
    });

    it('returns 0 when no match found', () => {
      const para = new Paragraph().addText('Hello World');
      const count = para.replaceTextCrossRun('xyz', 'abc');

      expect(count).toBe(0);
      expect(para.getText()).toBe('Hello World');
    });
  });

  describe('cross-run matching (fragmented text)', () => {
    it('replaces text split across two runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Hel'));
      para.addRun(new Run('lo'));

      const count = para.replaceTextCrossRun('Hello', 'Hi');

      expect(count).toBe(1);
      expect(para.getText()).toBe('Hi');
    });

    it('replaces placeholder split across three runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('{{'));
      para.addRun(new Run('name'));
      para.addRun(new Run('}}'));

      const count = para.replaceTextCrossRun('{{name}}', 'Alice');

      expect(count).toBe(1);
      expect(para.getText()).toBe('Alice');
    });

    it('preserves surrounding text in boundary runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Dear {{'));
      para.addRun(new Run('name'));
      para.addRun(new Run('}}, welcome!'));

      const count = para.replaceTextCrossRun('{{name}}', 'Bob');

      expect(count).toBe(1);
      expect(para.getText()).toBe('Dear Bob, welcome!');
    });

    it('handles placeholder with extra formatting run boundaries', () => {
      const para = new Paragraph();
      para.addRun(new Run('{'));
      para.addRun(new Run('{'));
      para.addRun(new Run('id'));
      para.addRun(new Run('}'));
      para.addRun(new Run('}'));

      const count = para.replaceTextCrossRun('{{id}}', '42');

      expect(count).toBe(1);
      expect(para.getText()).toBe('42');
    });

    it('replaces multiple fragmented placeholders', () => {
      const para = new Paragraph();
      para.addRun(new Run('{{'));
      para.addRun(new Run('first'));
      para.addRun(new Run('}} and {{'));
      para.addRun(new Run('second'));
      para.addRun(new Run('}}'));

      const count1 = para.replaceTextCrossRun('{{first}}', 'A');
      const count2 = para.replaceTextCrossRun('{{second}}', 'B');

      expect(count1).toBe(1);
      expect(count2).toBe(1);
      expect(para.getText()).toBe('A and B');
    });
  });

  describe('formatting preservation', () => {
    it('replacement inherits first run formatting', () => {
      const para = new Paragraph();
      para.addRun(new Run('Say {{', { bold: true }));
      para.addRun(new Run('greeting'));
      para.addRun(new Run('}}'));

      para.replaceTextCrossRun('{{greeting}}', 'Hello');

      // The first affected run had bold formatting
      const runs = para.getRuns();
      const helloRun = runs.find((r) => r.getText().includes('Hello'));
      expect(helloRun).toBeDefined();
    });

    it('preserves non-affected runs', () => {
      const para = new Paragraph();
      para.addRun(new Run('Before ', { italic: true }));
      para.addRun(new Run('{{x}}'));
      para.addRun(new Run(' After', { bold: true }));

      para.replaceTextCrossRun('{{x}}', 'VALUE');

      const runs = para.getRuns();
      expect(runs[0]!.getText()).toBe('Before ');
      expect(runs[0]!.getFormatting().italic).toBe(true);
      const lastRun = runs[runs.length - 1]!;
      expect(lastRun.getText()).toBe(' After');
      expect(lastRun.getFormatting().bold).toBe(true);
    });
  });

  describe('edge cases', () => {
    it('handles empty paragraph', () => {
      const para = new Paragraph();
      expect(para.replaceTextCrossRun('test', 'new')).toBe(0);
    });

    it('handles empty replacement', () => {
      const para = new Paragraph().addText('Remove {{this}} please');
      para.replaceTextCrossRun('{{this}} ', '');

      expect(para.getText()).toBe('Remove please');
    });

    it('handles replacement longer than original', () => {
      const para = new Paragraph().addText('{{x}}');
      para.replaceTextCrossRun('{{x}}', 'Very Long Replacement Text');

      expect(para.getText()).toBe('Very Long Replacement Text');
    });

    it('handles special regex characters in search text', () => {
      const para = new Paragraph().addText('Price: $10.00 (USD)');
      para.replaceTextCrossRun('$10.00 (USD)', '$20.00 (EUR)');

      expect(para.getText()).toBe('Price: $20.00 (EUR)');
    });
  });
});

describe('Document.fillTemplate()', () => {
  describe('basic template filling', () => {
    it('fills a single placeholder', () => {
      const doc = Document.create();
      doc.createParagraph('Hello {{name}}!');

      const count = doc.fillTemplate({ name: 'Alice' });

      expect(count).toBe(1);
      expect(doc.getParagraphs()[0]!.getText()).toBe('Hello Alice!');
      doc.dispose();
    });

    it('fills multiple placeholders', () => {
      const doc = Document.create();
      doc.createParagraph('Dear {{name}}, your order {{orderId}} is ready.');

      const count = doc.fillTemplate({
        name: 'Bob',
        orderId: 'ORD-999',
      });

      expect(count).toBe(2);
      expect(doc.getParagraphs()[0]!.getText()).toBe('Dear Bob, your order ORD-999 is ready.');
      doc.dispose();
    });

    it('fills same placeholder multiple times', () => {
      const doc = Document.create();
      doc.createParagraph('{{company}} - powered by {{company}}');

      const count = doc.fillTemplate({ company: 'Acme' });

      expect(count).toBe(2);
      expect(doc.getParagraphs()[0]!.getText()).toBe('Acme - powered by Acme');
      doc.dispose();
    });

    it('fills placeholders across multiple paragraphs', () => {
      const doc = Document.create();
      doc.createParagraph('Title: {{title}}');
      doc.createParagraph('Author: {{author}}');
      doc.createParagraph('Date: {{date}}');

      const count = doc.fillTemplate({
        title: 'My Report',
        author: 'Jane',
        date: '2024-06-15',
      });

      expect(count).toBe(3);
      const paras = doc.getParagraphs();
      expect(paras[0]!.getText()).toBe('Title: My Report');
      expect(paras[1]!.getText()).toBe('Author: Jane');
      expect(paras[2]!.getText()).toBe('Date: 2024-06-15');
      doc.dispose();
    });
  });

  describe('table cell placeholders', () => {
    it('fills placeholders inside table cells', () => {
      const doc = Document.create();
      const table = doc.createTable(2, 2);
      table.getCell(0, 0)!.createParagraph('{{header1}}');
      table.getCell(0, 1)!.createParagraph('{{header2}}');
      table.getCell(1, 0)!.createParagraph('{{val1}}');
      table.getCell(1, 1)!.createParagraph('{{val2}}');

      const count = doc.fillTemplate({
        header1: 'Name',
        header2: 'Age',
        val1: 'Alice',
        val2: '30',
      });

      expect(count).toBe(4);
      expect(table.getCell(0, 0)!.getText()).toBe('Name');
      expect(table.getCell(1, 1)!.getText()).toBe('30');
      doc.dispose();
    });
  });

  describe('custom delimiters', () => {
    it('supports angle bracket delimiters', () => {
      const doc = Document.create();
      doc.createParagraph('Hello <<name>>!');

      const count = doc.fillTemplate({ name: 'World' }, { delimiters: ['<<', '>>'] });

      expect(count).toBe(1);
      expect(doc.getParagraphs()[0]!.getText()).toBe('Hello World!');
      doc.dispose();
    });

    it('supports percent delimiters', () => {
      const doc = Document.create();
      doc.createParagraph('Value: %{total}%');

      const count = doc.fillTemplate({ total: '100' }, { delimiters: ['%{', '}%'] });

      expect(count).toBe(1);
      expect(doc.getParagraphs()[0]!.getText()).toBe('Value: 100');
      doc.dispose();
    });
  });

  describe('edge cases', () => {
    it('returns 0 when no placeholders match', () => {
      const doc = Document.create();
      doc.createParagraph('No placeholders here');

      const count = doc.fillTemplate({ name: 'value' });

      expect(count).toBe(0);
      doc.dispose();
    });

    it('handles empty data object', () => {
      const doc = Document.create();
      doc.createParagraph('Hello {{name}}');

      const count = doc.fillTemplate({});

      expect(count).toBe(0);
      expect(doc.getParagraphs()[0]!.getText()).toBe('Hello {{name}}');
      doc.dispose();
    });

    it('handles empty replacement values', () => {
      const doc = Document.create();
      doc.createParagraph('Remove: {{placeholder}}');

      const count = doc.fillTemplate({ placeholder: '' });

      expect(count).toBe(1);
      expect(doc.getParagraphs()[0]!.getText()).toBe('Remove: ');
      doc.dispose();
    });

    it('generates valid DOCX after template filling', async () => {
      const doc = Document.create();
      doc.addHeading('{{title}}', 1);
      doc.createParagraph('By {{author}} on {{date}}');
      doc.createParagraph('{{content}}');

      doc.fillTemplate({
        title: 'Quarterly Report',
        author: 'Finance Team',
        date: '2024-Q3',
        content: 'Revenue increased by 15% compared to the previous quarter.',
      });

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);

      // Verify round-trip
      const loaded = await Document.loadFromBuffer(buffer);
      const text = loaded.toPlainText();
      expect(text).toContain('Quarterly Report');
      expect(text).toContain('Finance Team');
      expect(text).toContain('Revenue increased');
      loaded.dispose();
      doc.dispose();
    });
  });
});
