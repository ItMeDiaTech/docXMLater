/**
 * clearDirectRunFormatting() must mutate run formatting in place instead of
 * rebuilding runs from getText().
 *
 * Reconstruction via `new Run(run.getText(), ...)` flattens non-text run
 * content: `<w:br w:type="page"/>` re-parses as a plain `<w:br/>` line
 * break, and content for which getText() returns '' (footnoteReference,
 * endnoteReference, fieldChar, symbol) is deleted outright — orphaning
 * notes in footnotes.xml. Per ECMA-376 §17.3.3.1 the break type is part
 * of document content, not formatting, so clearing formatting must leave
 * it untouched.
 */

import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('clearDirectRunFormatting() preserves non-text run content', () => {
  let doc: Document;

  beforeEach(() => {
    doc = Document.create();
  });

  afterEach(() => {
    doc.dispose();
  });

  it('keeps a page break with its break type when clearing ALL formatting', async () => {
    const para = doc.createParagraph();
    const run = new Run('Before the break', { bold: true, color: 'FF0000' });
    run.addBreak('page');
    para.addRun(run);

    para.clearDirectRunFormatting();

    const cleared = para.getRuns()[0]!;
    expect(cleared.getFormatting()).toEqual({});
    const breakContent = cleared.getContent().find((c) => c.type === 'break');
    expect(breakContent).toBeDefined();
    expect(breakContent!.breakType).toBe('page');

    const saved = await doc.toBuffer();
    const outZip = new ZipHandler();
    await outZip.loadFromBuffer(saved);
    const savedXml = outZip.getFileAsString('word/document.xml') ?? '';
    expect(savedXml).toContain('<w:br w:type="page"/>');
  });

  it('keeps a footnote reference when clearing ALL formatting', () => {
    const para = doc.createParagraph('Anchor text');
    const refRun = Run.createFromContent([{ type: 'footnoteReference', footnoteId: 2 }], {
      superscript: true,
    });
    para.addRun(refRun);

    para.clearDirectRunFormatting();

    const runs = para.getRuns();
    const refContent = runs
      .flatMap((r) => r.getContent())
      .find((c) => c.type === 'footnoteReference');
    expect(refContent).toBeDefined();
    expect(refContent!.footnoteId).toBe(2);
  });

  it('keeps content and untargeted formatting when clearing selected properties', () => {
    const para = doc.createParagraph();
    const run = new Run('Chapter end', { bold: true, color: 'FF0000' });
    run.addBreak('page');
    para.addRun(run);

    para.clearDirectRunFormatting(['color']);

    const cleared = para.getRuns()[0]!;
    expect(cleared.getFormatting().color).toBeUndefined();
    expect(cleared.getFormatting().bold).toBe(true);
    const breakContent = cleared.getContent().find((c) => c.type === 'break');
    expect(breakContent).toBeDefined();
    expect(breakContent!.breakType).toBe('page');
  });

  it('keeps field characters when clearing ALL formatting', () => {
    const para = doc.createParagraph();
    const fieldRun = Run.createFromContent([{ type: 'fieldChar', fieldCharType: 'begin' }], {
      bold: true,
    });
    para.addRun(fieldRun);

    para.clearDirectRunFormatting();

    const runs = para.getRuns();
    const fieldContent = runs.flatMap((r) => r.getContent()).find((c) => c.type === 'fieldChar');
    expect(fieldContent).toBeDefined();
    expect(fieldContent!.fieldCharType).toBe('begin');
    expect(runs[0]!.getFormatting()).toEqual({});
  });
});
