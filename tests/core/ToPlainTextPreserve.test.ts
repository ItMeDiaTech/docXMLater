/**
 * Regression test for issue #32 — toPlainText() truncated tracked text in
 * preserve mode.
 *
 * Before the fix, toPlainText() delegated to Paragraph.getText(), which filters
 * content to Run | Hyperlink and drops Revision objects. For a preserve-mode
 * document, inserted (w:t) and deleted (w:delText) text vanished, so the edit
 * "Hello World" -> "Hi World" rendered as "H World". toPlainText() now walks a
 * revision-aware path while the default getText() filter is unchanged.
 */

import { Document, Paragraph } from '../../src';

/**
 * Builds a buffer for a single-paragraph document whose only paragraph contains
 * a tracked edit changing "Hello World" to "Hi World" (deletes "ello", inserts
 * "i"). The saved buffer retains w:ins/w:del markup so it can be reloaded under
 * each revisionHandling mode.
 */
async function buildTrackedEditBuffer(): Promise<Buffer> {
  const doc = Document.create();
  const para = new Paragraph();
  para.addText('Hello World');
  doc.addParagraph(para);
  doc.enableTrackChanges({ author: 'Reviewer' });
  para.replaceAll('Hello', 'Hi');
  const buffer = await doc.toBuffer();
  doc.dispose();
  return buffer;
}

describe('Document.toPlainText with preserved revisions (issue #32)', () => {
  it('does not truncate tracked text in preserve mode', async () => {
    const buffer = await buildTrackedEditBuffer();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    const text = doc.toPlainText();

    // Document-order concatenation: "H" + deleted "ello" + inserted "i" + " World".
    expect(text).toBe('Helloi World');
    // Both the deleted and inserted fragments survive (the bug dropped them).
    expect(text).toContain('ello');
    expect(text).toContain('i');
    // The truncated output produced by the bug must not appear.
    expect(text).not.toBe('H World');

    doc.dispose();
  });

  it('keeps accept-mode toPlainText unchanged (insertions kept, deletions dropped)', async () => {
    const buffer = await buildTrackedEditBuffer();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'accept' });

    expect(doc.toPlainText()).toBe('Hi World');

    doc.dispose();
  });

  it('keeps strip-mode toPlainText unchanged (same resulting text as accept)', async () => {
    const buffer = await buildTrackedEditBuffer();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'strip' });

    expect(doc.toPlainText()).toBe('Hi World');

    doc.dispose();
  });
});
