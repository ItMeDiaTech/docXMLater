/**
 * Regression test for issue #33 — bindTrackingToAllElements skipped Hyperlink.
 *
 * Before the fix, editing a hyperlink's display text while track changes were
 * enabled silently replaced the text with no w:ins/w:del markup, because the
 * Hyperlink never received the document's tracking context. The fix binds any
 * paragraph-content item exposing _setTrackingContext (not just ComplexField),
 * so Hyperlink.setText() now produces tracked deletion/insertion revisions.
 */

import { Document, Paragraph, Hyperlink } from '../../src';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('Hyperlink tracked changes (issue #33)', () => {
  it('emits w:ins and w:del inside w:hyperlink when setText runs under track changes', async () => {
    const doc = Document.create();
    const para = new Paragraph();
    const link = new Hyperlink({ url: 'https://example.com', text: 'OLD' });
    para.addHyperlink(link);
    doc.addParagraph(para);

    // Enabling track changes binds the tracking context to the hyperlink.
    doc.enableTrackChanges({ author: 'Reviewer' });
    link.setText('NEW');

    const buffer = await doc.toBuffer();
    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    const xml = zip.getFileAsString('word/document.xml') || '';

    // Tracked markup must be generated for the hyperlink edit.
    expect(xml).toContain('<w:ins ');
    expect(xml).toContain('<w:del ');
    expect(xml).toContain('<w:hyperlink');

    // The deletion carries the old text (w:delText), the insertion the new text.
    expect(xml).toContain('<w:delText xml:space="preserve">OLD</w:delText>');
    expect(xml).toContain('<w:t xml:space="preserve">NEW</w:t>');

    // The revisions are nested inside w:hyperlink elements.
    expect(xml).toMatch(/<w:hyperlink[^>]*>\s*<w:del[ >]/);
    expect(xml).toMatch(/<w:hyperlink[^>]*>\s*<w:ins[ >]/);

    doc.dispose();
  });
});
