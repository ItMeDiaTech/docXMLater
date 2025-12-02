/**
 * Tests for InMemoryRevisionAcceptor
 *
 * Verifies that the in-memory DOM transformation approach correctly:
 * 1. Accepts insertions (unwraps content)
 * 2. Accepts deletions (removes content)
 * 3. Accepts move operations
 * 4. Accepts property changes
 * 5. Allows subsequent modifications to work correctly
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import {
  acceptRevisionsInMemory,
  paragraphHasRevisions,
  getRevisionsFromParagraph,
  countRevisionsByType,
} from '../../src/utils/InMemoryRevisionAcceptor';

describe('InMemoryRevisionAcceptor', () => {
  describe('acceptRevisionsInMemory', () => {
    it('should accept insertion revisions by unwrapping content', () => {
      // Create document with insertion revision
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add a regular run
      para.addRun(new Run('Regular text. '));

      // Add an insertion revision with content
      const insertedRun = new Run('Inserted text.');
      const insertionRevision = Revision.createInsertion('Test Author', insertedRun);
      para.addRevision(insertionRevision);

      // Verify revision exists before acceptance
      expect(paragraphHasRevisions(para)).toBe(true);
      expect(para.getRevisions().length).toBe(1);

      // Accept revisions
      const result = acceptRevisionsInMemory(doc);

      // Verify revision was accepted
      expect(result.insertionsAccepted).toBe(1);
      expect(result.totalAccepted).toBe(1);

      // Verify revision was removed and content was kept
      expect(paragraphHasRevisions(para)).toBe(false);
      expect(para.getRuns().length).toBe(2);
      expect(para.getText()).toBe('Regular text. Inserted text.');
    });

    it('should accept deletion revisions by removing content', () => {
      // Create document with deletion revision
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add a regular run
      para.addRun(new Run('Keep this. '));

      // Add a deletion revision (content should be removed)
      const deletedRun = new Run('Delete this.');
      const deletionRevision = Revision.createDeletion('Test Author', deletedRun);
      para.addRevision(deletionRevision);

      // Add another regular run
      para.addRun(new Run(' And keep this.'));

      // Verify revision exists before acceptance
      expect(paragraphHasRevisions(para)).toBe(true);

      // Accept revisions
      const result = acceptRevisionsInMemory(doc);

      // Verify revision was accepted
      expect(result.deletionsAccepted).toBe(1);

      // Verify revision was removed and content was deleted
      expect(paragraphHasRevisions(para)).toBe(false);
      expect(para.getText()).toBe('Keep this.  And keep this.');
    });

    it('should accept moveFrom revisions by removing content', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add moveFrom revision (source location - should be removed)
      const movedRun = new Run('Moved text');
      const moveFromRevision = Revision.createMoveFrom('Test Author', movedRun, 'move-1');
      para.addRevision(moveFromRevision);

      // Accept revisions
      const result = acceptRevisionsInMemory(doc);

      // Verify moveFrom was accepted (content removed)
      expect(result.movesAccepted).toBe(1);
      expect(paragraphHasRevisions(para)).toBe(false);
      expect(para.getText()).toBe('');
    });

    it('should accept moveTo revisions by keeping content', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add moveTo revision (destination location - keep content)
      const movedRun = new Run('Moved text');
      const moveToRevision = Revision.createMoveTo('Test Author', movedRun, 'move-1');
      para.addRevision(moveToRevision);

      // Accept revisions
      const result = acceptRevisionsInMemory(doc);

      // Verify moveTo was accepted (content kept)
      expect(result.movesAccepted).toBe(1);
      expect(paragraphHasRevisions(para)).toBe(false);
      expect(para.getText()).toBe('Moved text');
    });

    it('should accept property change revisions', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add run with formatting change tracked
      const run = new Run('Formatted text');
      const propertyChangeRevision = Revision.createRunPropertiesChange(
        'Test Author',
        run,
        { b: true } // Previous: was bold
      );
      para.addRevision(propertyChangeRevision);

      // Accept revisions
      const result = acceptRevisionsInMemory(doc);

      // Verify property change was accepted
      expect(result.propertyChangesAccepted).toBe(1);
      expect(paragraphHasRevisions(para)).toBe(false);
      // Content should be preserved
      expect(para.getText()).toBe('Formatted text');
    });

    it('should handle multiple revisions in a single paragraph', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add insertion
      para.addRevision(Revision.createInsertion('Author1', new Run('Inserted. ')));

      // Add regular text
      para.addRun(new Run('Regular. '));

      // Add deletion
      para.addRevision(Revision.createDeletion('Author2', new Run('Deleted. ')));

      // Add another insertion
      para.addRevision(Revision.createInsertion('Author1', new Run('Also inserted.')));

      // Accept all
      const result = acceptRevisionsInMemory(doc);

      expect(result.insertionsAccepted).toBe(2);
      expect(result.deletionsAccepted).toBe(1);
      expect(result.totalAccepted).toBe(3);

      // Final text should be: "Inserted. Regular. Also inserted."
      // (Deleted text is removed)
      expect(para.getText()).toBe('Inserted. Regular. Also inserted.');
    });

    it('should handle revisions across multiple paragraphs', () => {
      const doc = Document.create();

      // First paragraph with insertion
      const para1 = doc.createParagraph();
      para1.addRevision(Revision.createInsertion('Author', new Run('Para 1 inserted')));

      // Second paragraph with deletion
      const para2 = doc.createParagraph();
      para2.addRevision(Revision.createDeletion('Author', new Run('Para 2 deleted')));

      // Accept all
      const result = acceptRevisionsInMemory(doc);

      expect(result.insertionsAccepted).toBe(1);
      expect(result.deletionsAccepted).toBe(1);
      expect(result.totalAccepted).toBe(2);

      expect(para1.getText()).toBe('Para 1 inserted');
      expect(para2.getText()).toBe('');
    });
  });

  describe('subsequent modifications after acceptance', () => {
    it('should allow text modifications after accepting revisions', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add revision
      para.addRevision(Revision.createInsertion('Author', new Run('Initial text')));

      // Accept revisions
      acceptRevisionsInMemory(doc);

      // Now modify the text
      const runs = para.getRuns();
      expect(runs.length).toBe(1);
      expect(runs[0]).toBeDefined();
      runs[0]!.setText('Modified text');

      // Verify modification worked
      expect(para.getText()).toBe('Modified text');
    });

    it('should allow formatting changes after accepting revisions', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add revision
      para.addRevision(Revision.createInsertion('Author', new Run('Text to format')));

      // Accept revisions
      acceptRevisionsInMemory(doc);

      // Now apply formatting
      const runs = para.getRuns();
      expect(runs[0]).toBeDefined();
      runs[0]!.setBold(true);
      runs[0]!.setColor('FF0000');

      // Verify formatting was applied
      expect(runs[0]!.getFormatting().bold).toBe(true);
      expect(runs[0]!.getFormatting().color).toBe('FF0000');
    });

    it('should allow adding new content after accepting revisions', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      // Add revision
      para.addRevision(Revision.createInsertion('Author', new Run('Original')));

      // Accept revisions
      acceptRevisionsInMemory(doc);

      // Add new content
      para.addRun(new Run(' plus new content'));

      // Verify new content was added
      expect(para.getText()).toBe('Original plus new content');
      expect(para.getRuns().length).toBe(2);
    });
  });

  describe('selective acceptance options', () => {
    it('should only accept insertions when other options are disabled', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      para.addRevision(Revision.createInsertion('Author', new Run('Insert')));
      para.addRevision(Revision.createDeletion('Author', new Run('Delete')));

      const result = acceptRevisionsInMemory(doc, {
        acceptInsertions: true,
        acceptDeletions: false,
      });

      expect(result.insertionsAccepted).toBe(1);
      expect(result.deletionsAccepted).toBe(0);

      // Insertion content should be present, deletion revision should remain
      const revisions = para.getRevisions();
      expect(revisions.length).toBe(1);
      expect(revisions[0]).toBeDefined();
      expect(revisions[0]!.getType()).toBe('delete');
    });

    it('should only accept deletions when other options are disabled', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      para.addRevision(Revision.createInsertion('Author', new Run('Insert')));
      para.addRevision(Revision.createDeletion('Author', new Run('Delete')));

      const result = acceptRevisionsInMemory(doc, {
        acceptInsertions: false,
        acceptDeletions: true,
      });

      expect(result.insertionsAccepted).toBe(0);
      expect(result.deletionsAccepted).toBe(1);

      // Deletion should be accepted (removed), insertion revision should remain
      const revisions = para.getRevisions();
      expect(revisions.length).toBe(1);
      expect(revisions[0]).toBeDefined();
      expect(revisions[0]!.getType()).toBe('insert');
    });
  });

  describe('helper functions', () => {
    describe('paragraphHasRevisions', () => {
      it('should return true when paragraph has revisions', () => {
        const para = new Paragraph();
        para.addRevision(Revision.createInsertion('Author', new Run('text')));
        expect(paragraphHasRevisions(para)).toBe(true);
      });

      it('should return false when paragraph has no revisions', () => {
        const para = new Paragraph();
        para.addRun(new Run('text'));
        expect(paragraphHasRevisions(para)).toBe(false);
      });
    });

    describe('getRevisionsFromParagraph', () => {
      it('should return all revisions from paragraph', () => {
        const para = new Paragraph();
        para.addRevision(Revision.createInsertion('A', new Run('1')));
        para.addRun(new Run('regular'));
        para.addRevision(Revision.createDeletion('B', new Run('2')));

        const revisions = getRevisionsFromParagraph(para);
        expect(revisions.length).toBe(2);
        expect(revisions[0]).toBeDefined();
        expect(revisions[1]).toBeDefined();
        expect(revisions[0]!.getType()).toBe('insert');
        expect(revisions[1]!.getType()).toBe('delete');
      });
    });

    describe('countRevisionsByType', () => {
      it('should count revisions by type across document', () => {
        const doc = Document.create();

        const para1 = doc.createParagraph();
        para1.addRevision(Revision.createInsertion('A', new Run('1')));
        para1.addRevision(Revision.createInsertion('A', new Run('2')));

        const para2 = doc.createParagraph();
        para2.addRevision(Revision.createDeletion('B', new Run('3')));

        const counts = countRevisionsByType(doc);
        expect(counts.get('insert')).toBe(2);
        expect(counts.get('delete')).toBe(1);
      });
    });
  });
});
