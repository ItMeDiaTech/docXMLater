/**
 * RevisionAutoFixer removal fixes (REV003/REV004 orphaned move markers,
 * REV103 empty revisions) must remove the Revision from the owning
 * paragraph's content, not just from the RevisionManager registry.
 *
 * Paragraph.toXML() serializes revisions from paragraph content and the
 * generator never consults the manager registry, so registry-only pruning
 * left the error-severity w:moveFrom/w:moveTo markup in saved
 * document.xml while fix() reported the issues as fixed.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { Table } from '../../src/elements/Table';
import { RevisionAutoFixer } from '../../src/validation/RevisionAutoFixer';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function savedDocumentXml(doc: Document): Promise<string> {
  const buffer = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  return zip.getFileAsString('word/document.xml') ?? '';
}

describe('RevisionAutoFixer — removal fixes update paragraph content (not just the registry)', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('removes orphaned moveFrom markup from the paragraph and saved XML (REV003)', async () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('Before');
    const orphan = new Revision({
      id: 1,
      type: 'moveFrom',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [new Run('moved away')],
      moveId: 'move-1',
    });
    para.addRevision(orphan);
    para.addText('After');
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(orphan);

    const result = RevisionAutoFixer.fix(doc, { onlyRules: ['REV003', 'REV004'] });

    expect(result.actions.some((a) => a.issue.code === 'REV003')).toBe(true);
    expect(para.getRevisions()).toHaveLength(0);

    const xml = await savedDocumentXml(doc);
    expect(xml).not.toMatch(/<w:moveFrom\b/);
    expect(xml).not.toContain('moved away');
    expect(xml).toContain('Before');
    expect(xml).toContain('After');
  });

  it('removes orphaned moveTo markup from a table cell paragraph (REV004)', async () => {
    doc = Document.create();
    const table = new Table(1, 1);
    const cellPara = new Paragraph();
    cellPara.addText('Cell text');
    const orphan = new Revision({
      id: 2,
      type: 'moveTo',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [new Run('moved here')],
      moveId: 'move-2',
    });
    cellPara.addRevision(orphan);
    table.getCell(0, 0)!.addParagraph(cellPara);
    doc.addBodyElement(table);
    doc.getRevisionManager().registerExisting(orphan);

    const result = RevisionAutoFixer.fix(doc, { onlyRules: ['REV003', 'REV004'] });

    expect(result.actions.some((a) => a.issue.code === 'REV004')).toBe(true);
    expect(cellPara.getRevisions()).toHaveLength(0);

    const xml = await savedDocumentXml(doc);
    expect(xml).not.toMatch(/<w:moveTo\b/);
    expect(xml).not.toContain('moved here');
    expect(xml).toContain('Cell text');
  });

  it('removes empty revisions from the paragraph and saved XML (REV103)', async () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('Kept');
    const empty = new Revision({
      id: 3,
      type: 'insert',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [new Run('')],
    });
    para.addRevision(empty);
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(empty);

    const result = RevisionAutoFixer.fix(doc, { onlyRules: ['REV103'] });

    expect(result.actions.some((a) => a.issue.code === 'REV103')).toBe(true);
    expect(para.getRevisions()).toHaveLength(0);

    const xml = await savedDocumentXml(doc);
    expect(xml).not.toMatch(/<w:ins\b/);
    expect(xml).toContain('Kept');
  });

  it('does not touch paragraph content in dry-run mode', async () => {
    doc = Document.create();
    const para = new Paragraph();
    const orphan = new Revision({
      id: 4,
      type: 'moveFrom',
      author: 'Author',
      date: new Date('2025-01-15T10:00:00Z'),
      content: [new Run('still here')],
      moveId: 'move-4',
    });
    para.addRevision(orphan);
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(orphan);

    const preview = RevisionAutoFixer.preview(doc, { onlyRules: ['REV003', 'REV004'] });

    expect(preview.actions.some((a) => a.issue.code === 'REV003')).toBe(true);
    expect(para.getRevisions()).toHaveLength(1);
    expect(para.getRevisions()[0]).toBe(orphan);
    expect(doc.getRevisionManager().getAllRevisions()).toContain(orphan);
  });
});
