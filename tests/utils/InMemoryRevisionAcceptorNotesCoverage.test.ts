/**
 * `acceptRevisionsInMemory` must walk footnote and endnote paragraphs.
 *
 * Per ECMA-376 Part 1 §17.11.4 (w:endnote) and §17.11.15 (w:footnote),
 * notes can contain block-level content including tracked changes.
 * The in-memory acceptor previously walked body paragraphs, tables,
 * headers, footers, and the section but never reached
 * `FootnoteManager.getAllFootnotes` or `EndnoteManager.getAllEndnotes`.
 *
 * On the load path, iteration 135's raw-XML acceptor extension now
 * strips revisions from `word/footnotes.xml` etc. on default-accept
 * load, so existing-document footnote revisions are accepted. But
 * **programmatically added** revisions — e.g. someone calls
 * `footnote.getParagraphs()[0].addContent(new Revision(...))` after
 * loading — were not visited by the in-memory acceptor and stayed
 * in the model after `doc.acceptAllRevisions()` returned.
 *
 * Iteration 137 adds footnote and endnote paragraph traversal to
 * `acceptRevisionsInMemory` so programmatic note revisions are
 * processed alongside body / header / footer revisions.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { Footnote } from '../../src/elements/Footnote';
import { Endnote } from '../../src/elements/Endnote';

async function loadEmptyDoc() {
  const zip = new ZipHandler();
  zip.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return Document.loadFromBuffer(await zip.toBuffer());
}

function paragraphHasRevisionType(p: Paragraph, type: string): boolean {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const content = (p as any).getContent?.() ?? [];
  for (const item of content) {
    if (
      item &&
      typeof item === 'object' &&
      'getType' in item &&
      typeof item.getType === 'function' &&
      item.getType() === type
    ) {
      return true;
    }
  }
  return false;
}

describe('acceptRevisionsInMemory — footnote / endnote paragraph traversal', () => {
  it('accepts a programmatic insertion revision inside a footnote paragraph', async () => {
    const doc = await loadEmptyDoc();
    const fn = new Footnote({ id: 1 });
    const para = new Paragraph();
    para.addRun(new Run('kept '));
    const insertedRun = new Run('inserted');
    const ins = new Revision({
      id: 10,
      author: 'A',
      date: new Date('2026-04-24T10:00:00Z'),
      type: 'insert',
      content: [insertedRun],
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (para as any).addContent(ins);
    fn.addParagraph(para);
    doc.getFootnoteManager().register(fn);

    expect(paragraphHasRevisionType(para, 'insert')).toBe(true);

    await doc.acceptAllRevisions();

    expect(paragraphHasRevisionType(para, 'insert')).toBe(false);
    expect(para.getText()).toContain('inserted');
    expect(para.getText()).toContain('kept');
    doc.dispose();
  });

  it('accepts a programmatic deletion revision inside a footnote paragraph', async () => {
    const doc = await loadEmptyDoc();
    const fn = new Footnote({ id: 2 });
    const para = new Paragraph();
    para.addRun(new Run('kept '));
    const del = new Revision({
      id: 11,
      author: 'A',
      date: new Date('2026-04-24T10:00:00Z'),
      type: 'delete',
      content: [new Run('removed')],
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (para as any).addContent(del);
    fn.addParagraph(para);
    doc.getFootnoteManager().register(fn);

    expect(paragraphHasRevisionType(para, 'delete')).toBe(true);

    await doc.acceptAllRevisions();

    expect(paragraphHasRevisionType(para, 'delete')).toBe(false);
    expect(para.getText()).toContain('kept');
    expect(para.getText()).not.toContain('removed');
    doc.dispose();
  });

  it('accepts a programmatic insertion revision inside an endnote paragraph', async () => {
    const doc = await loadEmptyDoc();
    const en = new Endnote({ id: 1 });
    const para = new Paragraph();
    const ins = new Revision({
      id: 12,
      author: 'A',
      date: new Date('2026-04-24T10:00:00Z'),
      type: 'insert',
      content: [new Run('endnote-ins')],
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (para as any).addContent(ins);
    en.addParagraph(para);
    doc.getEndnoteManager().register(en);

    expect(paragraphHasRevisionType(para, 'insert')).toBe(true);

    await doc.acceptAllRevisions();

    expect(paragraphHasRevisionType(para, 'insert')).toBe(false);
    expect(para.getText()).toContain('endnote-ins');
    doc.dispose();
  });

  it('regression guard: footnote without revisions is untouched', async () => {
    const doc = await loadEmptyDoc();
    const fn = new Footnote({ id: 3 });
    const para = new Paragraph();
    para.addRun(new Run('plain note text'));
    fn.addParagraph(para);
    doc.getFootnoteManager().register(fn);

    await doc.acceptAllRevisions();

    expect(para.getText()).toBe('plain note text');
    doc.dispose();
  });
});
