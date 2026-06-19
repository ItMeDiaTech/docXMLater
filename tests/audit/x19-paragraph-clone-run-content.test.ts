/**
 * Paragraph.clone() must deep-clone runs via Run.clone() instead of
 * rebuilding them from getText().
 *
 * Reconstruction via `new Run(run.getText(), ...)` flattens non-text run
 * content: getText() maps every break to '\n' (losing breakType and
 * breakClear, so `<w:br w:type="page"/>` re-parses as a plain line break)
 * and returns '' for fieldChar, footnoteReference, endnoteReference,
 * symbol, and VML — deleting them from the clone. TableCell.clone(),
 * TableRow.clone(), Table.clone(), and Table.duplicateRow() all funnel
 * through Paragraph.clone(), so the loss propagates to table duplication.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('Paragraph.clone() preserves non-text run content', () => {
  it('preserves page break type and clear attribute', () => {
    const para = new Paragraph();
    const run = new Run('Before break', { bold: true });
    run.addBreak('page');
    para.addRun(run);

    const clone = para.clone();
    const content = clone
      .getRuns()
      .flatMap((r) => r.getContent())
      .filter((c) => c.type === 'break');

    expect(content).toHaveLength(1);
    expect(content[0]!.breakType).toBe('page');
    expect(clone.getRuns()[0]!.getFormatting().bold).toBe(true);
  });

  it('preserves textWrapping break clear attribute', () => {
    const para = new Paragraph();
    const run = new Run('Wrap');
    run.addBreak('textWrapping', 'all');
    para.addRun(run);

    const breakContent = para
      .clone()
      .getRuns()
      .flatMap((r) => r.getContent())
      .find((c) => c.type === 'break');

    expect(breakContent).toBeDefined();
    expect(breakContent!.breakType).toBe('textWrapping');
    expect(breakContent!.breakClear).toBe('all');
  });

  it('preserves footnote references, field chars, and symbols', () => {
    const para = new Paragraph();
    para.addRun(
      Run.createFromContent(
        [
          { type: 'fieldChar', fieldCharType: 'begin' },
          { type: 'instructionText', value: ' PAGE ' },
          { type: 'fieldChar', fieldCharType: 'end' },
        ],
        {}
      )
    );
    para.addRun(
      Run.createFromContent([{ type: 'footnoteReference', footnoteId: 2 }], {
        superscript: true,
      })
    );
    para.addRun(
      Run.createFromContent([{ type: 'symbol', symbolFont: 'Wingdings', symbolChar: 'F0FC' }], {})
    );

    const content = para
      .clone()
      .getRuns()
      .flatMap((r) => r.getContent());

    const fieldChars = content.filter((c) => c.type === 'fieldChar');
    expect(fieldChars.map((c) => c.fieldCharType)).toEqual(['begin', 'end']);
    expect(content.find((c) => c.type === 'instructionText')?.value).toBe(' PAGE ');

    const ref = content.find((c) => c.type === 'footnoteReference');
    expect(ref).toBeDefined();
    expect(ref!.footnoteId).toBe(2);

    const sym = content.find((c) => c.type === 'symbol');
    expect(sym).toBeDefined();
    expect(sym!.symbolFont).toBe('Wingdings');
    expect(sym!.symbolChar).toBe('F0FC');
  });

  it('deep-copies run content so the clone is independent', () => {
    const para = new Paragraph();
    const run = new Run('Text');
    run.addBreak('page');
    para.addRun(run);

    const clone = para.clone();
    const clonedBreak = clone
      .getRuns()[0]!
      .getContent()
      .find((c) => c.type === 'break')!;
    clonedBreak.breakType = 'column';

    const originalBreak = run.getContent().find((c) => c.type === 'break')!;
    expect(originalBreak.breakType).toBe('page');
  });

  it('saves a cloned page break as <w:br w:type="page"/>', async () => {
    const doc = Document.create();
    try {
      const para = doc.createParagraph();
      const run = new Run('Chapter end');
      run.addBreak('page');
      para.addRun(run);

      doc.addParagraph(para.clone());

      const saved = await doc.toBuffer();
      const outZip = new ZipHandler();
      await outZip.loadFromBuffer(saved);
      const savedXml = outZip.getFileAsString('word/document.xml') ?? '';

      const pageBreaks = savedXml.match(/<w:br w:type="page"\/>/g) ?? [];
      expect(pageBreaks).toHaveLength(2);
    } finally {
      doc.dispose();
    }
  });

  it('preserves page breaks through Table.duplicateRow()', () => {
    const table = new Table(1, 1);
    const cellPara = table.getCell(0, 0)!.createParagraph();
    const run = new Run('Cell content');
    run.addBreak('page');
    cellPara.addRun(run);

    table.duplicateRow(0);

    const duplicated = table
      .getRow(1)!
      .getCell(0)!
      .getParagraphs()[0]!
      .getRuns()
      .flatMap((r) => r.getContent())
      .find((c) => c.type === 'break');

    expect(duplicated).toBeDefined();
    expect(duplicated!.breakType).toBe('page');
  });
});
