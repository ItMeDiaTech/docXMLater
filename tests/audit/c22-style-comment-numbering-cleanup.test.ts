/**
 * cleanupUnusedNumbering() must treat style-bound numbering (w:numPr inside a
 * style's pPr) and numbering referenced from comments.xml as "in use".
 *
 * Word commonly attaches multilevel list numbering to paragraph styles
 * (e.g., Headings, list styles); paragraphs using such a style carry no
 * direct w:numPr, so without scanning styles the cleanup deletes numbering
 * definitions that are still referenced, silently destroying list numbering.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Style } from '../../src/formatting/Style';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

describe('cleanupUnusedNumbering() style and comment scanning', () => {
  it('preserves numbering referenced only by a paragraph style numPr', () => {
    const doc = Document.create();
    try {
      const manager = doc.getNumberingManager();

      const bodyNumId = manager.createBulletList();
      const styleNumId = manager.createNumberedList();

      // Body paragraph with a direct numId
      const bodyPara = new Paragraph().addText('Body item');
      bodyPara.setNumbering(bodyNumId, 0);
      doc.addParagraph(bodyPara);

      // Style carries the only reference to styleNumId
      doc.getStylesManager().addStyle(
        Style.create({
          styleId: 'StyledList',
          name: 'Styled List',
          type: 'paragraph',
          numPr: { numId: styleNumId, ilvl: 0 },
        })
      );

      // Paragraph numbered purely via the style — no direct w:numPr
      const styledPara = new Paragraph().addText('Styled item');
      styledPara.setStyle('StyledList');
      doc.addParagraph(styledPara);

      doc.cleanupUnusedNumbering();

      expect(manager.getInstance(styleNumId)).toBeDefined();
      expect(manager.getInstanceCount()).toBe(2);
      expect(manager.getAbstractNumberingCount()).toBe(2);
    } finally {
      doc.dispose();
    }
  });

  it('retains style-bound numbering in saved numbering.xml after round-trip cleanup', async () => {
    const doc = Document.create();
    const manager = doc.getNumberingManager();

    const bodyNumId = manager.createBulletList();
    const styleNumId = manager.createNumberedList();

    const bodyPara = new Paragraph().addText('Body item');
    bodyPara.setNumbering(bodyNumId, 0);
    doc.addParagraph(bodyPara);

    doc.getStylesManager().addStyle(
      Style.create({
        styleId: 'StyledList',
        name: 'Styled List',
        type: 'paragraph',
        numPr: { numId: styleNumId, ilvl: 0 },
      })
    );

    const styledPara = new Paragraph().addText('Styled item');
    styledPara.setStyle('StyledList');
    doc.addParagraph(styledPara);

    const buffer = await doc.toBuffer();
    doc.dispose();

    const reloaded = await Document.loadFromBuffer(buffer);
    try {
      reloaded.cleanupUnusedNumbering();

      const output = await reloaded.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING)!;

      expect(numberingXml).toContain(`<w:num w:numId="${styleNumId}"`);
    } finally {
      reloaded.dispose();
    }
  });

  it('preserves numbering referenced only inside comments.xml', async () => {
    const doc = Document.create();
    const manager = doc.getNumberingManager();

    const bodyNumId = manager.createBulletList();
    const commentNumId = manager.createNumberedList();

    const bodyPara = new Paragraph().addText('Body item');
    bodyPara.setNumbering(bodyNumId, 0);
    doc.addParagraph(bodyPara);

    const buffer = await doc.toBuffer();
    doc.dispose();

    // Inject a comments part whose paragraph is the only commentNumId reference
    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer);
    zipHandler.addFile(
      'word/comments.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z" w:initials="R">
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${commentNumId}"/></w:numPr></w:pPr>
      <w:r><w:t>Comment list item</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>`
    );
    const modifiedBuffer = await zipHandler.toBuffer();

    const reloaded = await Document.loadFromBuffer(modifiedBuffer);
    try {
      reloaded.cleanupUnusedNumbering();

      const reloadedManager = reloaded.getNumberingManager();
      expect(reloadedManager.getInstance(commentNumId)).toBeDefined();
      expect(reloadedManager.getInstance(bodyNumId)).toBeDefined();
    } finally {
      reloaded.dispose();
    }
  });
});
