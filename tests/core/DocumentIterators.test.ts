/**
 * Lazy iterator API: iterateParagraphs / iterateBodyElements /
 * iterateSections. Designed for streaming over very large documents
 * without materialising the full paragraph array.
 */
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';
import { StructuredDocumentTag } from '../../src/elements/StructuredDocumentTag';

describe('Document iterators', () => {
  describe('iterateParagraphs', () => {
    it('yields body paragraphs in document order', () => {
      const doc = Document.create();
      doc.createParagraph('A');
      doc.createParagraph('B');
      doc.createParagraph('C');

      const seen: string[] = [];
      for (const p of doc.iterateParagraphs()) seen.push(p.getText());
      expect(seen).toEqual(['A', 'B', 'C']);
      doc.dispose();
    });

    it('yields paragraphs nested inside tables', () => {
      const doc = Document.create();
      doc.createParagraph('before');
      const t = new Table(2, 2);
      t.getCell(0, 0)!.addParagraph(new Paragraph().addText('cell00'));
      t.getCell(1, 1)!.addParagraph(new Paragraph().addText('cell11'));
      doc.addTable(t);
      doc.createParagraph('after');

      const seen: string[] = [];
      for (const p of doc.iterateParagraphs()) seen.push(p.getText());
      // Body paragraphs interleaved with cell paragraphs.
      expect(seen).toContain('before');
      expect(seen).toContain('after');
      expect(seen).toContain('cell00');
      expect(seen).toContain('cell11');
      doc.dispose();
    });

    it('yields paragraphs nested inside SDTs (matches getAllParagraphs)', () => {
      const doc = Document.create();
      doc.createParagraph('outer');
      const inner = new Paragraph().addText('inside-sdt');
      doc.addBodyElement(new StructuredDocumentTag({}, [inner]));
      doc.createParagraph('after');

      const iterated = [...doc.iterateParagraphs()].map((p) => p.getText());
      const arr = doc.getAllParagraphs().map((p) => p.getText());
      expect(iterated).toEqual(arr);
      expect(iterated).toContain('inside-sdt');
      doc.dispose();
    });

    it('returns an iterator that is lazy (yields on demand)', () => {
      const doc = Document.create();
      for (let i = 0; i < 100; i++) doc.createParagraph(`p${i}`);

      const it = doc.iterateParagraphs();
      const first = it.next();
      expect(first.value!.getText()).toBe('p0');
      const second = it.next();
      expect(second.value!.getText()).toBe('p1');
      // Stop after two yields — no requirement to drain.
      doc.dispose();
    });
  });

  describe('iterateBodyElements', () => {
    it('preserves table boundaries (does not flatten cells)', () => {
      const doc = Document.create();
      doc.createParagraph('A');
      doc.addTable(new Table(2, 2));
      doc.createParagraph('B');

      const types: string[] = [];
      for (const el of doc.iterateBodyElements()) {
        if (el instanceof Paragraph) types.push('paragraph');
        else if (el instanceof Table) types.push('table');
        else types.push('other');
      }
      expect(types).toEqual(['paragraph', 'table', 'paragraph']);
      doc.dispose();
    });
  });

  describe('iterateSections', () => {
    it('yields the trailing group with the document root section when no inline sectPr exists', () => {
      const doc = Document.create();
      doc.createParagraph('p1');
      doc.createParagraph('p2');

      const sections = [...doc.iterateSections()];
      expect(sections).toHaveLength(1);
      expect(sections[0]!.elements).toHaveLength(2);
      expect(sections[0]!.sectPr).toBeDefined();
      doc.dispose();
    });

    it('splits at paragraphs carrying inline sectPr', () => {
      const doc = Document.create();
      doc.createParagraph('section1-end').formatting.sectPr =
        '<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';
      doc.createParagraph('section2-start');
      doc.createParagraph('section2-end');

      const sections = [...doc.iterateSections()];
      expect(sections).toHaveLength(2);
      // First group ended at the section-break paragraph.
      expect(sections[0]!.elements).toHaveLength(1);
      expect((sections[0]!.elements[0] as Paragraph).getText()).toBe('section1-end');
      // Second group is the trailing two paragraphs.
      expect(sections[1]!.elements).toHaveLength(2);
      doc.dispose();
    });
  });
});
