/**
 * `SelectiveRevisionAcceptor.accept` / `reject` must walk headers,
 * footers, footnotes, and endnotes alongside body paragraphs and
 * tables.
 *
 * Per ECMA-376 Part 1, tracked-change revisions can appear inside:
 *   - Body paragraphs (§17.2)
 *   - Tables (§17.4)
 *   - Headers and footers (§17.10.3-4)
 *   - Footnotes (§17.11.15) and endnotes (§17.11.4)
 *
 * Bug: `SelectiveRevisionAcceptor` (`utils/SelectiveRevisionAcceptor.ts`)
 * only walks `doc.getAllParagraphs()` and `doc.getTables()`. Selective
 * acceptance / rejection of revisions living inside headers, footers,
 * footnotes, or endnotes silently no-ops on the in-memory model — the
 * revision survives in the paragraph's content array even when the
 * caller's criteria match it.
 *
 * Iteration 138 extends the walker to mirror the in-memory acceptor's
 * traversal (iter 137): tables, headers, footers, footnote paragraphs,
 * endnote paragraphs.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { Footnote } from '../../src/elements/Footnote';
import { Endnote } from '../../src/elements/Endnote';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';
import { SelectiveRevisionAcceptor } from '../../src/utils/SelectiveRevisionAcceptor';

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

function buildParagraphWithIns(text: string, insText: string, author: string): Paragraph {
  const p = new Paragraph();
  p.addRun(new Run(text));
  const ins = new Revision({
    id: 0,
    author,
    date: new Date('2026-04-24T10:00:00Z'),
    type: 'insert',
    content: [new Run(insText)],
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (p as any).addContent(ins);
  return p;
}

describe('SelectiveRevisionAcceptor — non-body coverage', () => {
  it('accepts a revision matching the criteria inside a footnote', async () => {
    const doc = await loadEmptyDoc();
    const fn = new Footnote({ id: 1 });
    const para = buildParagraphWithIns('kept ', 'fn-ins', 'Alice');
    fn.addParagraph(para);
    doc.getFootnoteManager().register(fn);

    expect(paragraphHasRevisionType(para, 'insert')).toBe(true);

    SelectiveRevisionAcceptor.accept(doc, { authors: ['Alice'] });

    expect(paragraphHasRevisionType(para, 'insert')).toBe(false);
    expect(para.getText()).toContain('fn-ins');
    expect(para.getText()).toContain('kept');
    doc.dispose();
  });

  it('rejects a revision matching the criteria inside an endnote (removes inserted content)', async () => {
    const doc = await loadEmptyDoc();
    const en = new Endnote({ id: 1 });
    const para = buildParagraphWithIns('kept ', 'en-ins', 'Alice');
    en.addParagraph(para);
    doc.getEndnoteManager().register(en);

    SelectiveRevisionAcceptor.reject(doc, { authors: ['Alice'] });

    expect(paragraphHasRevisionType(para, 'insert')).toBe(false);
    // Reject of an insertion removes the inserted content
    expect(para.getText()).toContain('kept');
    expect(para.getText()).not.toContain('en-ins');
    doc.dispose();
  });

  it('accepts a revision matching the criteria inside a registered header', async () => {
    const doc = await loadEmptyDoc();
    const para = buildParagraphWithIns('hdr ', 'header-ins', 'Bob');
    const header = new Header();
    header.addParagraph(para);
    doc.getHeaderFooterManager().registerHeader(header, 'rIdHdr1');

    expect(paragraphHasRevisionType(para, 'insert')).toBe(true);

    SelectiveRevisionAcceptor.accept(doc, { authors: ['Bob'] });

    expect(paragraphHasRevisionType(para, 'insert')).toBe(false);
    expect(para.getText()).toContain('header-ins');
    doc.dispose();
  });

  it('accepts a revision matching the criteria inside a registered footer', async () => {
    const doc = await loadEmptyDoc();
    const para = buildParagraphWithIns('ftr ', 'footer-ins', 'Bob');
    const footer = new Footer();
    footer.addParagraph(para);
    doc.getHeaderFooterManager().registerFooter(footer, 'rIdFtr1');

    expect(paragraphHasRevisionType(para, 'insert')).toBe(true);

    SelectiveRevisionAcceptor.accept(doc, { authors: ['Bob'] });

    expect(paragraphHasRevisionType(para, 'insert')).toBe(false);
    expect(para.getText()).toContain('footer-ins');
    doc.dispose();
  });

  it('preserves revisions whose author does NOT match (regression guard)', async () => {
    const doc = await loadEmptyDoc();
    const fn = new Footnote({ id: 1 });
    const para = buildParagraphWithIns('kept ', 'preserved-ins', 'Alice');
    fn.addParagraph(para);
    doc.getFootnoteManager().register(fn);

    // Different author — should NOT match
    SelectiveRevisionAcceptor.accept(doc, { authors: ['Bob'] });

    expect(paragraphHasRevisionType(para, 'insert')).toBe(true);
    doc.dispose();
  });
});
