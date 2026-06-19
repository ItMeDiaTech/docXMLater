/**
 * Regression tests: conditional table-style formatting (w:tblStylePr pPr/rPr)
 * must not leak into the style's unconditional base formatting. Per CT_Style
 * (ECMA-376 §17.7.4.17) the root pPr/rPr precede w:tblStylePr blocks, so a
 * first-match search over the full w:style element misattributes a conditional
 * block's pPr/rPr as base formatting whenever the root pPr/rPr are absent —
 * the common shape of Word built-in banded table styles (firstRow-only
 * bold/centered with no root-level formatting).
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="table" w:styleId="BandedConditionalOnly">
    <w:name w:val="Banded Conditional Only"/>
    <w:tblStylePr w:type="firstRow">
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:rPr><w:b/></w:rPr>
    </w:tblStylePr>
  </w:style>
  <w:style w:type="table" w:styleId="RootPlusConditional">
    <w:name w:val="Root Plus Conditional"/>
    <w:pPr><w:jc w:val="right"/></w:pPr>
    <w:rPr><w:i/></w:rPr>
    <w:tblStylePr w:type="firstRow">
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:rPr><w:b/></w:rPr>
    </w:tblStylePr>
  </w:style>
</w:styles>`;

async function buildDocxWithCustomStyles(): Promise<Buffer> {
  const doc = Document.create();
  try {
    doc.addParagraph(new Paragraph().addText('Test'));
    const buffer = await doc.toBuffer();
    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer);
    zipHandler.updateFile(DOCX_PATHS.STYLES, STYLES_XML);
    return await zipHandler.toBuffer();
  } finally {
    doc.dispose();
  }
}

describe('table style conditional pPr/rPr does not leak into base formatting (C42)', () => {
  it('leaves base formatting empty when the style has only tblStylePr formatting', async () => {
    const buffer = await buildDocxWithCustomStyles();
    const doc = await Document.loadFromBuffer(buffer);
    try {
      const style = doc.getStyle('BandedConditionalOnly');
      expect(style).toBeDefined();

      // firstRow-only formatting must not become unconditional base formatting
      expect(style!.getParagraphFormatting()).toBeUndefined();
      expect(style!.getRunFormatting()).toBeUndefined();

      // The conditional block itself must still be parsed
      const conditionals = style!.getProperties().tableStyle?.conditionalFormatting;
      expect(conditionals).toBeDefined();
      expect(conditionals).toHaveLength(1);
      expect(conditionals![0]!.type).toBe('firstRow');
      expect(conditionals![0]!.paragraphFormatting?.alignment).toBe('center');
      expect(conditionals![0]!.runFormatting?.bold).toBe(true);
    } finally {
      doc.dispose();
    }
  });

  it('keeps root-level pPr/rPr as base formatting alongside conditional blocks', async () => {
    const buffer = await buildDocxWithCustomStyles();
    const doc = await Document.loadFromBuffer(buffer);
    try {
      const style = doc.getStyle('RootPlusConditional');
      expect(style).toBeDefined();

      // Root pPr/rPr must win, not the conditional block's pPr/rPr
      expect(style!.getParagraphFormatting()?.alignment).toBe('right');
      expect(style!.getRunFormatting()?.italic).toBe(true);
      expect(style!.getRunFormatting()?.bold).toBeUndefined();

      const conditionals = style!.getProperties().tableStyle?.conditionalFormatting;
      expect(conditionals).toHaveLength(1);
      expect(conditionals![0]!.paragraphFormatting?.alignment).toBe('center');
      expect(conditionals![0]!.runFormatting?.bold).toBe(true);
    } finally {
      doc.dispose();
    }
  });
});
