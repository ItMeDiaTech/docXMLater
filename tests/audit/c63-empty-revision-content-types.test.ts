/**
 * REV103 (empty revision) must recognize every content shape ECMA-376
 * allows inside w:ins/w:del: hyperlinks (w:hyperlink) and image runs
 * (w:r with w:drawing) in addition to text runs, plus the structurally
 * contentless table cell markers w:cellIns/w:cellDel/w:cellMerge.
 *
 * A runs-only text check flagged hyperlink-only and image-only tracked
 * changes as "empty" and the auto-fixer then deregistered these valid
 * revisions, corrupting manager-driven outputs (changelogs, statistics).
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import type { RevisionType } from '../../src/elements/Revision';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Image } from '../../src/elements/Image';
import { ImageRun } from '../../src/elements/ImageRun';
import { RevisionValidator } from '../../src/validation/RevisionValidator';
import { RevisionAutoFixer } from '../../src/validation/RevisionAutoFixer';

/** 1x1 transparent PNG */
function createTestImageBuffer(): Buffer {
  return Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);
}

async function createImageRun(): Promise<ImageRun> {
  const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
  return new ImageRun(image);
}

/** Revision IDs flagged REV103 by a full validation pass */
function rev103Ids(doc: Document): number[] {
  const result = RevisionValidator.validate(doc);
  return result.warnings
    .filter((i) => i.code === 'REV103')
    .map((i) => i.location?.revisionId ?? -1);
}

describe('REV103 empty-revision check — hyperlink, image, and cell-marker content (validator)', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('does not flag an insertion whose only content is a hyperlink with display text', () => {
    doc = Document.create();
    const para = new Paragraph();
    const rev = new Revision({
      id: 1,
      type: 'insert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [Hyperlink.createExternal('https://example.com', 'Example')],
    });
    para.addRevision(rev);
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(rev);

    expect(rev103Ids(doc)).not.toContain(1);
  });

  it('does not flag an insertion whose only hyperlink has no display text', () => {
    doc = Document.create();
    const para = new Paragraph();
    const emptyLink = new Hyperlink({ url: 'https://example.com', isEmpty: true });
    const rev = new Revision({
      id: 2,
      type: 'insert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [emptyLink],
    });
    para.addRevision(rev);
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(rev);

    expect(emptyLink.getText()).toBe('');
    expect(rev103Ids(doc)).not.toContain(2);
  });

  it('does not flag an image-only insertion', async () => {
    doc = Document.create();
    const para = new Paragraph();
    const rev = new Revision({
      id: 3,
      type: 'insert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [await createImageRun()],
    });
    para.addRevision(rev);
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(rev);

    expect(rev103Ids(doc)).not.toContain(3);
  });

  it('does not flag table cell markers (cellIns/cellDel/cellMerge)', () => {
    doc = Document.create();
    const types: RevisionType[] = ['tableCellInsert', 'tableCellDelete', 'tableCellMerge'];
    types.forEach((type, index) => {
      const rev = new Revision({
        id: 10 + index,
        type,
        author: 'Author',
        date: new Date('2025-01-15T10:00:00Z'),
        content: [],
      });
      doc!.getRevisionManager().registerExisting(rev);
    });

    const flagged = rev103Ids(doc);
    expect(flagged).not.toContain(10);
    expect(flagged).not.toContain(11);
    expect(flagged).not.toContain(12);
  });

  it('still flags an insertion with only empty text runs', () => {
    doc = Document.create();
    const para = new Paragraph();
    const rev = new Revision({
      id: 4,
      type: 'insert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [new Run('')],
    });
    para.addRevision(rev);
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(rev);

    expect(rev103Ids(doc)).toContain(4);
  });
});

describe('REV103 auto-fix — keeps hyperlink-only and image-only revisions (fixer)', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('removes only the truly empty revision, not hyperlink/image content', async () => {
    doc = Document.create();
    const para = new Paragraph();

    const hyperlinkRev = new Revision({
      id: 1,
      type: 'insert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [Hyperlink.createExternal('https://example.com', 'Example')],
    });
    const imageRev = new Revision({
      id: 2,
      type: 'insert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [await createImageRun()],
    });
    const emptyRev = new Revision({
      id: 3,
      type: 'insert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [new Run('')],
    });

    para.addRevision(hyperlinkRev);
    para.addRevision(imageRev);
    para.addRevision(emptyRev);
    doc.addParagraph(para);
    const manager = doc.getRevisionManager();
    manager.registerExisting(hyperlinkRev);
    manager.registerExisting(imageRev);
    manager.registerExisting(emptyRev);

    const result = RevisionAutoFixer.fix(doc, { onlyRules: ['REV103'] });

    const fixedIds = result.actions
      .filter((a) => a.issue.code === 'REV103')
      .map((a) => a.issue.location?.revisionId);
    expect(fixedIds).toEqual([3]);

    expect(manager.getAllRevisions()).toContain(hyperlinkRev);
    expect(manager.getAllRevisions()).toContain(imageRev);
    expect(manager.getAllRevisions()).not.toContain(emptyRev);
    expect(para.getRevisions()).toContain(hyperlinkRev);
    expect(para.getRevisions()).toContain(imageRev);
    expect(para.getRevisions()).not.toContain(emptyRev);
  });

  it('does not deregister table cell markers', () => {
    doc = Document.create();
    const cellMarker = new Revision({
      id: 7,
      type: 'tableCellInsert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [],
    });
    const manager = doc.getRevisionManager();
    manager.registerExisting(cellMarker);

    const result = RevisionAutoFixer.fix(doc, { onlyRules: ['REV103'] });

    expect(result.actions.filter((a) => a.issue.code === 'REV103')).toHaveLength(0);
    expect(manager.getAllRevisions()).toContain(cellMarker);
  });
});
