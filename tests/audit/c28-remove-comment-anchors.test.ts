/**
 * Tests that Document.removeComment() removes the comment's body anchors
 * (w:commentRangeStart, w:commentRangeEnd, w:commentReference) along with
 * the comment definition, for both programmatic and loaded documents.
 * Dangling anchors reference a comment that no longer exists in
 * comments.xml, which Word reports as unreadable content.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('removeComment anchor cleanup', () => {
  it('removes anchors from programmatic paragraphs', async () => {
    const doc = Document.create();
    try {
      const para = doc.createParagraph('Commented text');
      const comment = doc.addCommentToParagraph(para, 'Author', 'Review this');

      expect(doc.removeComment(comment.getId())).toBe(true);

      // In-memory anchors are gone
      expect(para.getCommentsStart()).toHaveLength(0);
      expect(para.getCommentsEnd()).toHaveLength(0);

      const buffer = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);

      // The comments part is removed (last comment) — no anchors may remain
      expect(zip.hasFile('word/comments.xml')).toBe(false);
      const docXml = zip.getFileAsString('word/document.xml')!;
      expect(docXml).not.toContain('w:commentRangeStart');
      expect(docXml).not.toContain('w:commentRangeEnd');
      expect(docXml).not.toContain('w:commentReference');
    } finally {
      doc.dispose();
    }
  });

  it('removes preserved raw-XML anchors after load round-trip', async () => {
    const doc = Document.create();
    const para = doc.createParagraph('Commented text');
    doc.addCommentToParagraph(para, 'Author', 'Review this');
    const buffer1 = await doc.toBuffer();
    doc.dispose();

    // Loaded anchors are PreservedElement raw XML, not Comment objects
    const doc2 = await Document.loadFromBuffer(buffer1);
    try {
      const comments = doc2.getAllComments();
      expect(comments).toHaveLength(1);
      expect(doc2.removeComment(comments[0]!.getId())).toBe(true);

      const buffer2 = await doc2.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer2);

      expect(zip.hasFile('word/comments.xml')).toBe(false);
      const docXml = zip.getFileAsString('word/document.xml')!;
      expect(docXml).not.toContain('w:commentRangeStart');
      expect(docXml).not.toContain('w:commentRangeEnd');
      expect(docXml).not.toContain('w:commentReference');
    } finally {
      doc2.dispose();
    }
  });

  it('removes anchors of replies deleted along with the parent comment', async () => {
    const doc = Document.create();
    try {
      const para = doc.createParagraph('Discussion');
      const parent = doc.createComment('User1', 'Initial comment');
      const reply = doc.createReply(parent.getId(), 'User2', 'Reply');
      para.addComment(parent);
      para.addComment(reply);

      expect(doc.removeComment(parent.getId())).toBe(true);

      const buffer = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);

      expect(zip.hasFile('word/comments.xml')).toBe(false);
      const docXml = zip.getFileAsString('word/document.xml')!;
      expect(docXml).not.toContain('w:commentRangeStart');
      expect(docXml).not.toContain('w:commentRangeEnd');
      expect(docXml).not.toContain('w:commentReference');
    } finally {
      doc.dispose();
    }
  });

  it('keeps anchors of other comments intact', async () => {
    const doc = Document.create();
    const para1 = doc.createParagraph('First commented text');
    const para2 = doc.createParagraph('Second commented text');
    const comment1 = doc.addCommentToParagraph(para1, 'Author', 'First');
    const comment2 = doc.addCommentToParagraph(para2, 'Author', 'Second');
    const buffer1 = await doc.toBuffer();
    doc.dispose();

    // Remove only the first comment on a loaded document
    const doc2 = await Document.loadFromBuffer(buffer1);
    try {
      expect(doc2.removeComment(comment1.getId())).toBe(true);

      const buffer2 = await doc2.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer2);

      // The surviving comment keeps its definition and anchors
      expect(zip.hasFile('word/comments.xml')).toBe(true);
      const docXml = zip.getFileAsString('word/document.xml')!;
      expect(docXml).not.toContain(`<w:commentRangeStart w:id="${comment1.getId()}"`);
      expect(docXml).not.toContain(`<w:commentRangeEnd w:id="${comment1.getId()}"`);
      expect(docXml).not.toContain(`<w:commentReference w:id="${comment1.getId()}"`);
      expect(docXml).toContain(`<w:commentRangeStart w:id="${comment2.getId()}"`);
      expect(docXml).toContain(`<w:commentRangeEnd w:id="${comment2.getId()}"`);
      expect(docXml).toContain(`<w:commentReference w:id="${comment2.getId()}"`);
    } finally {
      doc2.dispose();
    }
  });
});
