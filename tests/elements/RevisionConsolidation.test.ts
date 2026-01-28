/**
 * Tests for revision consolidation functionality
 * Addresses the "random insertions and deletions" problem where Word displays
 * many small revisions instead of consolidated ones.
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Revision } from '../../src/elements/Revision';
import { Run } from '../../src/elements/Run';
import { Document } from '../../src/core/Document';

describe('Revision Consolidation', () => {
  describe('Paragraph.consolidateRevisions()', () => {
    it('should consolidate adjacent insertions from same author', () => {
      const para = new Paragraph();
      const now = new Date();

      // Add two insertions from same author within time window
      const rev1 = Revision.createInsertion('Author', new Run('Hello '), now);
      const rev2 = Revision.createInsertion('Author', new Run('World'), new Date(now.getTime() + 100));

      para.addRevision(rev1);
      para.addRevision(rev2);

      expect(para.getRevisions()).toHaveLength(2);

      // Consolidate
      const consolidated = para.consolidateRevisions(1000);

      // Should have merged into one
      expect(para.getRevisions()).toHaveLength(1);
      expect(consolidated).toBe(1);

      // Combined content should have both runs
      const mergedRevision = para.getRevisions()[0];
      expect(mergedRevision?.getContent()).toHaveLength(2);
      expect(mergedRevision?.getText()).toBe('Hello World');
    });

    it('should consolidate adjacent deletions from same author', () => {
      const para = new Paragraph();
      const now = new Date();

      const rev1 = Revision.createDeletion('Author', new Run('foo '), now);
      const rev2 = Revision.createDeletion('Author', new Run('bar'), new Date(now.getTime() + 100));

      para.addRevision(rev1);
      para.addRevision(rev2);

      const consolidated = para.consolidateRevisions(1000);

      expect(para.getRevisions()).toHaveLength(1);
      expect(consolidated).toBe(1);
      expect(para.getRevisions()[0]?.getType()).toBe('delete');
    });

    it('should NOT consolidate revisions from different authors', () => {
      const para = new Paragraph();
      const now = new Date();

      const rev1 = Revision.createInsertion('Author1', new Run('Hello '), now);
      const rev2 = Revision.createInsertion('Author2', new Run('World'), new Date(now.getTime() + 100));

      para.addRevision(rev1);
      para.addRevision(rev2);

      const consolidated = para.consolidateRevisions(1000);

      expect(para.getRevisions()).toHaveLength(2);
      expect(consolidated).toBe(0);
    });

    it('should NOT consolidate revisions of different types', () => {
      const para = new Paragraph();
      const now = new Date();

      const rev1 = Revision.createInsertion('Author', new Run('Hello '), now);
      const rev2 = Revision.createDeletion('Author', new Run('World'), new Date(now.getTime() + 100));

      para.addRevision(rev1);
      para.addRevision(rev2);

      const consolidated = para.consolidateRevisions(1000);

      expect(para.getRevisions()).toHaveLength(2);
      expect(consolidated).toBe(0);
    });

    it('should NOT consolidate revisions outside time window', () => {
      const para = new Paragraph();
      const now = new Date();

      // Second revision is 2 seconds later (outside 1 second window)
      const rev1 = Revision.createInsertion('Author', new Run('Hello '), now);
      const rev2 = Revision.createInsertion('Author', new Run('World'), new Date(now.getTime() + 2000));

      para.addRevision(rev1);
      para.addRevision(rev2);

      const consolidated = para.consolidateRevisions(1000);

      expect(para.getRevisions()).toHaveLength(2);
      expect(consolidated).toBe(0);
    });

    it('should stop consolidation at non-revision content boundaries', () => {
      const para = new Paragraph();
      const now = new Date();

      // Revision -> Run -> Revision pattern
      const rev1 = Revision.createInsertion('Author', new Run('A'), now);
      para.addRevision(rev1);

      // Add a non-revision Run (simulating normal content) using addRun
      para.addRun(new Run('Normal'));

      const rev2 = Revision.createInsertion('Author', new Run('B'), new Date(now.getTime() + 100));
      para.addRevision(rev2);

      const consolidated = para.consolidateRevisions(1000);

      // Should not consolidate because of the Run in between
      expect(para.getRevisions()).toHaveLength(2);
      expect(consolidated).toBe(0);
    });

    it('should handle multiple consecutive revisions', () => {
      const para = new Paragraph();
      const now = new Date();

      // Add 5 insertions from same author within time window
      for (let i = 0; i < 5; i++) {
        const rev = Revision.createInsertion('Author', new Run(`Part${i}`), new Date(now.getTime() + i * 100));
        para.addRevision(rev);
      }

      expect(para.getRevisions()).toHaveLength(5);

      const consolidated = para.consolidateRevisions(1000);

      expect(para.getRevisions()).toHaveLength(1);
      expect(consolidated).toBe(4);
      expect(para.getRevisions()[0]?.getText()).toBe('Part0Part1Part2Part3Part4');
    });

    it('should return 0 when no revisions exist', () => {
      const para = new Paragraph();
      para.addText('Normal text');

      const consolidated = para.consolidateRevisions();

      expect(consolidated).toBe(0);
    });

    it('should return 0 when single revision exists', () => {
      const para = new Paragraph();
      para.addRevision(Revision.createInsertion('Author', new Run('Hello')));

      const consolidated = para.consolidateRevisions();

      expect(para.getRevisions()).toHaveLength(1);
      expect(consolidated).toBe(0);
    });

    it('should NOT consolidate property change revisions', () => {
      const para = new Paragraph();
      const now = new Date();

      const rev1 = Revision.createRunPropertiesChange('Author', new Run('text'), { bold: true }, now);
      const rev2 = Revision.createRunPropertiesChange('Author', new Run('text'), { italic: true }, new Date(now.getTime() + 100));

      para.addRevision(rev1);
      para.addRevision(rev2);

      const consolidated = para.consolidateRevisions(1000);

      // Property changes should NOT be consolidated
      expect(para.getRevisions()).toHaveLength(2);
      expect(consolidated).toBe(0);
    });

    it('should use custom time window', () => {
      const para = new Paragraph();
      const now = new Date();

      // Two revisions 500ms apart
      para.addRevision(Revision.createInsertion('Author', new Run('A'), now));
      para.addRevision(Revision.createInsertion('Author', new Run('B'), new Date(now.getTime() + 500)));

      // With 250ms window, should NOT consolidate
      let consolidated = para.consolidateRevisions(250);
      expect(para.getRevisions()).toHaveLength(2);
      expect(consolidated).toBe(0);

      // Create fresh paragraph for second test
      const para2 = new Paragraph();
      para2.addRevision(Revision.createInsertion('Author', new Run('A'), now));
      para2.addRevision(Revision.createInsertion('Author', new Run('B'), new Date(now.getTime() + 500)));

      // With 1000ms window, should consolidate
      consolidated = para2.consolidateRevisions(1000);
      expect(para2.getRevisions()).toHaveLength(1);
      expect(consolidated).toBe(1);
    });
  });

  describe('Document.consolidateAllRevisions()', () => {
    it('should consolidate revisions across all paragraphs', () => {
      const doc = Document.create();
      const now = new Date();

      // Create paragraphs and add revisions to them
      const para1 = Paragraph.create();
      para1.addRevision(Revision.createInsertion('Author', new Run('A'), now));
      para1.addRevision(Revision.createInsertion('Author', new Run('B'), new Date(now.getTime() + 100)));
      doc.addParagraph(para1);

      const para2 = Paragraph.create();
      para2.addRevision(Revision.createInsertion('Author', new Run('C'), now));
      para2.addRevision(Revision.createInsertion('Author', new Run('D'), new Date(now.getTime() + 100)));
      doc.addParagraph(para2);

      const result = doc.consolidateAllRevisions(1000);

      // Should have processed both paragraphs and consolidated 2 revisions (1 per paragraph)
      expect(result.paragraphsProcessed).toBeGreaterThanOrEqual(2);
      expect(result.revisionsConsolidated).toBe(2);

      // Each paragraph should have 1 consolidated revision
      expect(para1.getRevisions()).toHaveLength(1);
      expect(para2.getRevisions()).toHaveLength(1);
    });

    it('should handle document with no revisions', () => {
      const doc = Document.create();
      const para1 = Paragraph.create('Normal text');
      const para2 = Paragraph.create('More text');
      doc.addParagraph(para1);
      doc.addParagraph(para2);

      const result = doc.consolidateAllRevisions();

      expect(result.revisionsConsolidated).toBe(0);
    });
  });
});
