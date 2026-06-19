/**
 * Paragraph.setText() replaced the content array directly, so with track
 * changes enabled the old content vanished with no w:del and the new text
 * appeared with no w:ins -- a silent untracked edit. The new Run also never
 * received parent-paragraph/tracking wiring, so subsequent run.setText()
 * tracking and conditional-formatting resolution silently no-oped.
 *
 * These tests pin the tracking-aware behavior mirroring clearContent() +
 * addText().
 */
import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('C45: Paragraph.setText() under change tracking', () => {
  it('wraps old content in delete revisions and the new run in an insert revision', () => {
    const doc = new Document();
    try {
      const para = doc.createParagraph();
      para.addText('Old text');

      doc.enableTrackChanges({ author: 'TestUser' });
      para.setText('New text');

      const content = para.getContent();
      expect(content).toHaveLength(2);
      expect(content[0]).toBeInstanceOf(Revision);
      expect((content[0] as Revision).getType()).toBe('delete');
      expect(content[1]).toBeInstanceOf(Revision);
      expect((content[1] as Revision).getType()).toBe('insert');
      expect((content[1] as Revision).getAuthor()).toBe('TestUser');

      const revManager = doc.getRevisionManager();
      expect(revManager.getDeletionCount()).toBe(1);
      expect(revManager.getInsertionCount()).toBe(1);
    } finally {
      doc.dispose();
    }
  });

  it('wires parent paragraph and tracking context on the inserted run', () => {
    const doc = new Document();
    try {
      const para = doc.createParagraph();
      para.addText('Old');

      doc.enableTrackChanges({ author: 'TestUser' });
      para.setText('New');

      const content = para.getContent();
      const insertion = content[content.length - 1] as Revision;
      const insertedRun = insertion.getContent()[0] as Run;
      expect(insertedRun).toBeInstanceOf(Run);
      expect(insertedRun._getParentParagraph()).toBe(para);
    } finally {
      doc.dispose();
    }
  });

  it('emits w:del for the old text and w:ins for the new text on save', async () => {
    const doc = Document.create();
    let saved: Buffer;
    try {
      const para = doc.createParagraph('Old text');
      doc.enableTrackChanges({ author: 'TestUser' });
      para.setText('New text');
      saved = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = new ZipHandler();
    await zip.loadFromBuffer(saved);
    const xml = zip.getFileAsString('word/document.xml')!;

    const del = /<w:del [\s\S]*?<\/w:del>/.exec(xml)?.[0];
    expect(del).toBeDefined();
    expect(del).toContain('Old text');

    const ins = /<w:ins [\s\S]*?<\/w:ins>/.exec(xml)?.[0];
    expect(ins).toBeDefined();
    expect(ins).toContain('New text');
  });

  it('replaces content directly and wires parent when tracking is disabled', () => {
    const doc = new Document();
    try {
      const para = doc.createParagraph();
      para.addText('First');
      para.addText('Second');

      para.setText('Replaced');

      const content = para.getContent();
      expect(content).toHaveLength(1);
      expect(content[0]).toBeInstanceOf(Run);
      expect(para.getText()).toBe('Replaced');
      expect((content[0] as Run)._getParentParagraph()).toBe(para);
    } finally {
      doc.dispose();
    }
  });
});
