/**
 * Tests that NumberingLevel does not fabricate explicit run properties
 * (w:sz/w:szCs/w:color) for levels that never specified them. Per ECMA-376
 * an absent rPr child means "inherit", so a parsed <w:lvl> without w:sz or
 * w:color must re-serialize without them when its abstractNum is regenerated.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { NumberingLevel } from '../../src/formatting/NumberingLevel';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/** Level 0 lvl XML with rPr carrying only rFonts — no w:sz, no w:color */
const LVL_NO_SZ_NO_COLOR =
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
  '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>' +
  '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond" w:hint="default"/></w:rPr></w:lvl>';

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    ${LVL_NO_SZ_NO_COLOR}
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

async function createDocxWithNumbering(numberingXml: string): Promise<Buffer> {
  const doc = Document.create();
  const para = new Paragraph().addText('Item');
  doc.addParagraph(para);
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.addFile(DOCX_PATHS.NUMBERING, numberingXml);
  return await zipHandler.toBuffer();
}

describe('NumberingLevel run-property defaults (inherit when unset)', () => {
  it('keeps fontSize and color undefined when not provided to the constructor', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'decimal',
      text: '%1.',
    });

    const props = level.getProperties();
    expect(props.fontSize).toBeUndefined();
    expect(props.color).toBeUndefined();

    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).not.toContain('<w:sz ');
    expect(xml).not.toContain('<w:szCs ');
    expect(xml).not.toContain('<w:color ');
  });

  it('parses a lvl lacking w:sz/w:color without coercing defaults', () => {
    const level = NumberingLevel.fromXML(LVL_NO_SZ_NO_COLOR);

    const props = level.getProperties();
    expect(props.fontSize).toBeUndefined();
    expect(props.color).toBeUndefined();
    expect(props.font).toBe('Garamond');

    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).not.toContain('<w:sz ');
    expect(xml).not.toContain('<w:color ');
  });

  it('still emits explicitly set fontSize and color', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'decimal',
      text: '%1.',
      fontSize: 28,
      color: 'FF0000',
    });

    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).toContain('<w:sz w:val="28"/>');
    expect(xml).toContain('<w:szCs w:val="28"/>');
    expect(xml).toContain('<w:color w:val="FF0000"/>');
  });

  it('does not inject w:sz/w:color when a modified abstractNum is re-serialized on save', async () => {
    const buffer = await createDocxWithNumbering(NUMBERING_XML);

    const doc = await Document.loadFromBuffer(buffer);
    let output: Buffer;
    try {
      // Force regeneration of abstractNum 0 through the merge-on-save path
      doc.getNumberingManager().markAbstractNumberingModified(0);
      output = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = new ZipHandler();
    await zip.loadFromBuffer(output);
    const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING);
    expect(numberingXml).toBeTruthy();

    expect(numberingXml!).not.toContain('<w:sz ');
    expect(numberingXml!).not.toContain('<w:color ');
    // The source font must survive regeneration
    expect(numberingXml!).toContain('w:ascii="Garamond"');
  });
});
