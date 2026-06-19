/**
 * In-memory revision acceptance must drop the boundary range markers that
 * delimit a move / ins / del span, not just the w:moveFrom/w:moveTo (or
 * ins/del) revisions themselves.
 *
 * Per ECMA-376 §17.13.5.21-28, w:moveFromRangeStart/End and
 * w:moveToRangeStart/End carry w:author/w:date and reference a move
 * operation. If acceptAllRevisions() unwraps the move but leaves these
 * markers behind, Word still reports the document as containing tracked
 * changes, defeating acceptance. The raw-XML acceptor
 * (acceptRevisions.ts) and stripRevisionsFromXml already remove the same
 * markers; the in-memory paragraph path (used by doc.acceptAllRevisions())
 * must match.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { RangeMarker } from '../../src/elements/RangeMarker';
import { MoveOperationHelper } from '../../src/processors/MoveOperationHelper';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function savedDocumentXml(doc: Document): Promise<string> {
  const buffer = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  return zip.getFileAsString('word/document.xml') ?? '';
}

describe('In-memory acceptAllRevisions removes move/ins/del range markers', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('strips all four move range markers after acceptAllRevisions()', async () => {
    doc = Document.create();
    const sourcePara = new Paragraph();
    sourcePara.addText('source ');
    const destPara = new Paragraph();
    destPara.addText('destination ');

    MoveOperationHelper.addMoveOperation(sourcePara, destPara, {
      author: 'MoveAuthor',
      content: new Run('moved content'),
      moveId: 'move-x10',
    });

    doc.addParagraph(sourcePara);
    doc.addParagraph(destPara);

    // Sanity: markers are present before acceptance.
    const before = await savedDocumentXml(doc);
    expect(before).toMatch(/<w:moveFromRangeStart\b/);
    expect(before).toMatch(/<w:moveToRangeStart\b/);

    await doc.acceptAllRevisions();

    const after = await savedDocumentXml(doc);

    // The move revisions themselves are gone...
    expect(after).not.toMatch(/<w:moveFrom\b/);
    expect(after).not.toMatch(/<w:moveTo\b/);

    // ...and so are the orphaned boundary range markers (the bug).
    expect(after).not.toContain('moveFromRangeStart');
    expect(after).not.toContain('moveFromRangeEnd');
    expect(after).not.toContain('moveToRangeStart');
    expect(after).not.toContain('moveToRangeEnd');

    // The moved content survives at the destination.
    expect(after).toContain('moved content');
  });

  it('strips customXml ins/del range markers when accepting insertions/deletions', async () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('text ');
    para.addRangeMarker(RangeMarker.createCustomXmlInsStart(2001, 'InsAuthor'));
    para.addRangeMarker(RangeMarker.createCustomXmlInsEnd(2001));
    para.addRangeMarker(RangeMarker.createCustomXmlDelStart(2002, 'DelAuthor'));
    para.addRangeMarker(RangeMarker.createCustomXmlDelEnd(2002));
    doc.addParagraph(para);

    const before = await savedDocumentXml(doc);
    expect(before).toContain('customXmlInsRangeStart');
    expect(before).toContain('customXmlDelRangeStart');

    await doc.acceptAllRevisions();

    const after = await savedDocumentXml(doc);
    expect(after).not.toContain('customXmlInsRangeStart');
    expect(after).not.toContain('customXmlInsRangeEnd');
    expect(after).not.toContain('customXmlDelRangeStart');
    expect(after).not.toContain('customXmlDelRangeEnd');
  });
});
