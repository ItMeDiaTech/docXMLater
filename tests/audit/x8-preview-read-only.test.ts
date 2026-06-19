/**
 * SelectiveRevisionAcceptor.preview() must be read-only.
 *
 * preview() is documented as "Preview what would happen without making
 * changes" — it must classify revisions against the criteria without
 * transforming paragraph content, clearing paragraph-mark markers, or
 * removing revisions from the RevisionManager.
 */

import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { SelectiveRevisionAcceptor } from '../../src/processors/SelectiveRevisionAcceptor';

describe('SelectiveRevisionAcceptor.preview() read-only contract', () => {
  let doc: Document;

  afterEach(() => {
    doc.dispose();
  });

  it('preview-reject does not remove inserted content from the paragraph', () => {
    doc = Document.create();
    const para = doc.createParagraph();
    para.addRun(new Run('Existing text. '));
    const insertion = Revision.createInsertion('Alice', new Run('Inserted text.'));
    para.addRevision(insertion);

    const result = SelectiveRevisionAcceptor.preview(doc, { authors: ['Alice'] }, 'reject');

    expect(result.rejected).toEqual([insertion.getId().toString()]);
    expect(result.summary.rejectedCount).toBe(1);

    // The insertion (and its text) must survive the preview
    expect(para.getRevisions()).toHaveLength(1);
    expect(para.getRevisions()[0]).toBe(insertion);
    expect(
      insertion
        .getRuns()
        .map((r) => r.getText())
        .join('')
    ).toBe('Inserted text.');
    expect(para.getText()).toBe('Existing text. ');
  });

  it('preview-accept does not unwrap revision markup or strip deletions', () => {
    doc = Document.create();
    const para = doc.createParagraph();
    const insertion = Revision.createInsertion('Alice', new Run('Added.'));
    const deletion = Revision.createDeletion('Alice', new Run('Removed.'));
    para.addRevision(insertion);
    para.addRevision(deletion);

    const result = SelectiveRevisionAcceptor.preview(doc, { authors: ['Alice'] }, 'accept');

    expect(result.accepted).toHaveLength(2);
    expect(result.rejected).toHaveLength(0);
    expect(result.summary.acceptedCount).toBe(2);

    // Both revisions must still be wrapped in the paragraph
    expect(para.getRevisions()).toHaveLength(2);
    expect(para.getRevisions()[0]).toBe(insertion);
    expect(para.getRevisions()[1]).toBe(deletion);
  });

  it('preview does not remove matched revisions from the RevisionManager', () => {
    doc = Document.create();
    const para = doc.createParagraph();
    const insertion = Revision.createInsertion('Alice', new Run('Tracked.'));
    para.addRevision(insertion);
    doc.getRevisionManager().register(insertion);

    SelectiveRevisionAcceptor.preview(doc, { authors: ['Alice'] }, 'accept');

    const allRevisions = doc.getRevisionManager().getAllRevisions();
    expect(allRevisions).toHaveLength(1);
    expect(allRevisions[0]).toBe(insertion);
  });

  it('preview-accept does not clear paragraph-mark deletion markers', () => {
    doc = Document.create();
    const para = doc.createParagraph();
    para.addRun(new Run('First'));
    const second = doc.createParagraph();
    second.addRun(new Run('Second'));
    para.markParagraphMarkAsDeleted(42, 'Alice');

    const result = SelectiveRevisionAcceptor.preview(doc, { authors: ['Alice'] }, 'accept');

    expect(result.accepted).toContain('42');
    // The marker must survive so a later accept() can still merge paragraphs
    expect(para.isParagraphMarkDeleted()).toBe(true);
    expect(doc.getAllParagraphs()).toHaveLength(2);
  });

  it('preview is repeatable: a second call reports the same counts', () => {
    doc = Document.create();
    const para = doc.createParagraph();
    para.addRevision(Revision.createInsertion('Alice', new Run('Once.')));

    const first = SelectiveRevisionAcceptor.preview(doc, { authors: ['Alice'] }, 'reject');
    const second = SelectiveRevisionAcceptor.preview(doc, { authors: ['Alice'] }, 'reject');

    expect(first.summary).toEqual(second.summary);
    expect(second.rejected).toHaveLength(1);
  });
});
