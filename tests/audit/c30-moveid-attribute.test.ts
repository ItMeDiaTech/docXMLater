/**
 * Per ECMA-376 §17.13.5.22/.23, w:moveFrom and w:moveTo are CT_RunTrackChange,
 * whose only attributes are w:id, w:author, and w:date. The schema declares no
 * w:moveId attribute; an undeclared attribute in the main w: namespace fails
 * Open XML schema validation and can trigger Word's unreadable-content repair.
 *
 * Move source/destination pairing must stay an in-memory key (Revision.moveId,
 * used by getMovePair/validateMovePairs) and be expressed on disk solely via
 * the w:name attribute on w:moveFromRangeStart/w:moveToRangeStart.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { MoveOperationHelper } from '../../src/processors/MoveOperationHelper';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function savedDocumentXml(doc: Document): Promise<string> {
  const buffer = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  return zip.getFileAsString('word/document.xml') ?? '';
}

describe('w:moveFrom/w:moveTo carry only CT_RunTrackChange attributes (no w:moveId)', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('omits w:moveId from moveFrom/moveTo toXML() while keeping the internal pairing key', () => {
    const moveFrom = Revision.createMoveFrom('Author', new Run('moved'), 'move-abc');
    moveFrom.setId(1);
    const moveTo = Revision.createMoveTo('Author', new Run('moved'), 'move-abc');
    moveTo.setId(2);

    const fromXml = moveFrom.toXML();
    expect(fromXml).not.toBeNull();
    expect(fromXml!.name).toBe('w:moveFrom');
    expect(Object.keys(fromXml!.attributes ?? {})).toEqual(['w:id', 'w:author', 'w:date']);

    const toXml = moveTo.toXML();
    expect(toXml).not.toBeNull();
    expect(toXml!.name).toBe('w:moveTo');
    expect(Object.keys(toXml!.attributes ?? {})).toEqual(['w:id', 'w:author', 'w:date']);

    // The pairing key remains available for getMovePair/validateMovePairs
    expect(moveFrom.getMoveId()).toBe('move-abc');
    expect(moveTo.getMoveId()).toBe('move-abc');
  });

  it('saves move operations without w:moveId, pairing via w:name on range start markers', async () => {
    doc = Document.create();
    const sourcePara = new Paragraph();
    sourcePara.addText('source ');
    const destPara = new Paragraph();
    destPara.addText('destination ');

    MoveOperationHelper.addMoveOperation(sourcePara, destPara, {
      author: 'MoveAuthor',
      content: new Run('moved content'),
      moveId: 'move-c30',
    });

    doc.addParagraph(sourcePara);
    doc.addParagraph(destPara);

    const xml = await savedDocumentXml(doc);

    // The undeclared attribute must never reach document.xml
    expect(xml).not.toContain('w:moveId');

    // Move markup itself is intact
    expect(xml).toMatch(/<w:moveFrom\b/);
    expect(xml).toMatch(/<w:moveTo\b/);

    // On-disk pairing is the w:name attribute on the CT_MoveBookmark range starts
    expect(xml).toMatch(/<w:moveFromRangeStart\b[^>]*w:name="move-c30"/);
    expect(xml).toMatch(/<w:moveToRangeStart\b[^>]*w:name="move-c30"/);
  });
});
