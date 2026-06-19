/**
 * Rejecting a run/paragraph property-change revision (w:rPrChange /
 * w:pPrChange, ECMA-376 §17.13.5.31 / §17.13.5.29) must restore the
 * previous properties stored in the change element — not keep the new
 * formatting, which would make reject behave identically to accept.
 * When no previous-property snapshot exists, the revision must stay in
 * place instead of being silently discarded.
 */

import { Document } from '../../src/core/Document';
import { Revision } from '../../src/elements/Revision';
import { Run } from '../../src/elements/Run';
import { SelectiveRevisionAcceptor } from '../../src/processors/SelectiveRevisionAcceptor';

describe('C36: rejecting property-change revisions restores previous properties', () => {
  let doc: Document;

  afterEach(() => {
    doc.dispose();
  });

  it('restores the previous run formatting when rejecting a runPropertiesChange', () => {
    doc = Document.create();
    const para = doc.createParagraph();
    const run = new Run('Changed text', { bold: true, color: 'FF0000' });
    const revision = Revision.createRunPropertiesChange(
      'Bob',
      run,
      { italic: true },
      new Date('2024-03-03')
    );
    doc.getRevisionManager().register(revision);
    para.addRevision(revision);

    const result = SelectiveRevisionAcceptor.rejectByType(doc, ['runPropertiesChange']);

    expect(result.rejected).toContain(revision.getId().toString());
    // The snapshot is the complete pre-change rPr: restored wholesale
    expect(run.getFormatting()).toEqual({ italic: true });
    expect(run.getText()).toBe('Changed text');
    // The wrapper is gone — the run sits directly in paragraph content
    const content = para.getContent();
    expect(content).toHaveLength(1);
    expect(content[0]).toBe(run);
    expect(doc.getRevisionManager().getAllRevisions()).toHaveLength(0);
  });

  it('restores the previous paragraph properties when rejecting via rejectFormattingChanges', () => {
    doc = Document.create();
    const para = doc.createParagraph();
    para.setAlignment('center');
    const run = new Run('Paragraph text');
    const revision = Revision.createParagraphPropertiesChange(
      'Carol',
      run,
      { alignment: 'left' },
      new Date('2024-03-03')
    );
    doc.getRevisionManager().register(revision);
    para.addRevision(revision);

    const result = SelectiveRevisionAcceptor.rejectFormattingChanges(doc);

    expect(result.rejected).toContain(revision.getId().toString());
    expect(para.getFormatting().alignment).toBe('left');
    expect(para.getText()).toBe('Paragraph text');
    const content = para.getContent();
    expect(content).toHaveLength(1);
    expect(content[0]).toBe(run);
  });

  it('keeps the revision in place when no previous-property snapshot exists', () => {
    doc = Document.create();
    const para = doc.createParagraph();
    const run = new Run('No snapshot', { bold: true });
    const revision = new Revision({
      author: 'Bob',
      type: 'runPropertiesChange',
      content: run,
      date: new Date('2024-03-03'),
    });
    doc.getRevisionManager().register(revision);
    para.addRevision(revision);

    const result = SelectiveRevisionAcceptor.rejectByType(doc, ['runPropertiesChange']);

    // Not rejected: reported as remaining, kept in content and manager
    expect(result.rejected).toHaveLength(0);
    expect(result.remaining).toContain(revision.getId().toString());
    const content = para.getContent();
    expect(content).toHaveLength(1);
    expect(content[0]).toBe(revision);
    expect(run.getFormatting()).toEqual({ bold: true });
    expect(doc.getRevisionManager().getAllRevisions()).toHaveLength(1);
  });

  it('round-trips the restored formatting through save and reload', async () => {
    doc = Document.create();
    const para = doc.createParagraph();
    const run = new Run('Round trip', { color: 'FF0000' });
    const revision = Revision.createRunPropertiesChange(
      'Bob',
      run,
      { bold: true },
      new Date('2024-03-03')
    );
    doc.getRevisionManager().register(revision);
    para.addRevision(revision);

    SelectiveRevisionAcceptor.rejectByType(doc, ['runPropertiesChange']);

    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    try {
      const paragraphs = reloaded.getAllParagraphs();
      const reloadedRun = paragraphs[0]?.getRuns()[0];
      expect(reloadedRun?.getText()).toBe('Round trip');
      expect(reloadedRun?.getBold()).toBe(true);
      expect(reloadedRun?.getColor()).toBeUndefined();
    } finally {
      reloaded.dispose();
    }
  });
});
