/**
 * Paragraph mark `<w:rPr>` child element order — ECMA-376 Part 1 §17.3.1.29
 * CT_ParaRPr schema compliance for tracked-change markers.
 *
 * CT_ParaRPr schema (ECMA-376 Part 1, Annex A / wml.xsd):
 *
 *   <xsd:complexType name="CT_ParaRPr">
 *     <xsd:sequence>
 *       <xsd:group ref="EG_ParaRPrTrackChanges" minOccurs="0"/>
 *         <!-- ins, del, moveFrom, moveTo -->
 *       <xsd:group ref="EG_RPrBase" minOccurs="0" maxOccurs="unbounded"/>
 *         <!-- rStyle, rFonts, b, bCs, i, iCs, caps, ... -->
 *       <xsd:element name="rPrChange" type="CT_ParaRPrChange" minOccurs="0"/>
 *     </xsd:sequence>
 *   </xsd:complexType>
 *
 * Bug guarded against: the generator previously pushed the EG_RPrBase run
 * property children FIRST and then `<w:ins>` / `<w:del>` AFTER. Word tolerates
 * both orders, but strict OOXML validators (e.g. Open XML SDK Validator)
 * flag the inverted order as a schema violation and tracked-change-aware
 * tools may misinterpret the revision state of the paragraph mark.
 *
 * Correct order per CT_ParaRPr: `[ins?, del?, moveFrom?, moveTo?,
 * ...EG_RPrBase..., rPrChange?]`.
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

function buildParaXml(paragraph: Paragraph): string {
  return XMLBuilder.elementToString(paragraph.toXML());
}

describe('Paragraph mark w:rPr child order (CT_ParaRPr §17.3.1.29)', () => {
  it('emits <w:ins> before EG_RPrBase run properties in the paragraph mark rPr', () => {
    const p = new Paragraph();
    p.addText('x');
    p.setParagraphMarkFormatting({ bold: true, color: 'FF0000' });
    p.markParagraphMarkAsInserted(7, 'Jane', new Date('2026-01-15T10:00:00Z'));

    const xml = buildParaXml(p);
    const pPrIdx = xml.indexOf('<w:pPr>');
    const pPrEnd = xml.indexOf('</w:pPr>');
    expect(pPrIdx).toBeGreaterThan(-1);
    const pPrBlock = xml.substring(pPrIdx, pPrEnd);

    const insIdx = pPrBlock.indexOf('<w:ins ');
    const boldIdx = pPrBlock.indexOf('<w:b ');
    const colorIdx = pPrBlock.indexOf('<w:color ');
    expect(insIdx).toBeGreaterThan(-1);
    expect(boldIdx).toBeGreaterThan(-1);
    expect(colorIdx).toBeGreaterThan(-1);
    // ins must precede every EG_RPrBase child per schema
    expect(insIdx).toBeLessThan(boldIdx);
    expect(insIdx).toBeLessThan(colorIdx);
  });

  it('emits <w:del> before EG_RPrBase run properties in the paragraph mark rPr', () => {
    const p = new Paragraph();
    p.addText('x');
    p.setParagraphMarkFormatting({ italic: true, size: 24 });
    p.markParagraphMarkAsDeleted(9, 'Bob', new Date('2026-01-15T10:00:00Z'));

    const xml = buildParaXml(p);
    const pPrIdx = xml.indexOf('<w:pPr>');
    const pPrEnd = xml.indexOf('</w:pPr>');
    const pPrBlock = xml.substring(pPrIdx, pPrEnd);

    const delIdx = pPrBlock.indexOf('<w:del ');
    const iIdx = pPrBlock.indexOf('<w:i ');
    const szIdx = pPrBlock.indexOf('<w:sz ');
    expect(delIdx).toBeGreaterThan(-1);
    expect(iIdx).toBeGreaterThan(-1);
    expect(szIdx).toBeGreaterThan(-1);
    expect(delIdx).toBeLessThan(iIdx);
    expect(delIdx).toBeLessThan(szIdx);
  });

  it('emits <w:ins> before <w:del> when both are present (EG_ParaRPrTrackChanges order)', () => {
    const p = new Paragraph();
    p.addText('x');
    p.markParagraphMarkAsInserted(1, 'Alice', new Date('2026-01-15T10:00:00Z'));
    p.markParagraphMarkAsDeleted(2, 'Alice', new Date('2026-01-15T10:00:00Z'));

    const xml = buildParaXml(p);
    const pPrIdx = xml.indexOf('<w:pPr>');
    const pPrEnd = xml.indexOf('</w:pPr>');
    const pPrBlock = xml.substring(pPrIdx, pPrEnd);

    const insIdx = pPrBlock.indexOf('<w:ins ');
    const delIdx = pPrBlock.indexOf('<w:del ');
    expect(insIdx).toBeGreaterThan(-1);
    expect(delIdx).toBeGreaterThan(-1);
    expect(insIdx).toBeLessThan(delIdx);
  });

  it('emits ins + del + run props in strict CT_ParaRPr order', () => {
    const p = new Paragraph();
    p.addText('x');
    p.setParagraphMarkFormatting({ bold: true });
    p.markParagraphMarkAsInserted(1, 'A', new Date('2026-01-15T10:00:00Z'));
    p.markParagraphMarkAsDeleted(2, 'A', new Date('2026-01-15T10:00:00Z'));

    const xml = buildParaXml(p);
    const pPrIdx = xml.indexOf('<w:pPr>');
    const pPrEnd = xml.indexOf('</w:pPr>');
    const pPrBlock = xml.substring(pPrIdx, pPrEnd);

    const insIdx = pPrBlock.indexOf('<w:ins ');
    const delIdx = pPrBlock.indexOf('<w:del ');
    const boldIdx = pPrBlock.indexOf('<w:b ');

    // Order: ins → del → bold (EG_ParaRPrTrackChanges → EG_RPrBase)
    expect(insIdx).toBeLessThan(delIdx);
    expect(delIdx).toBeLessThan(boldIdx);
  });

  it('preserves paragraph mark rPr when only run properties are set (no track changes)', () => {
    const p = new Paragraph();
    p.addText('x');
    p.setParagraphMarkFormatting({ bold: true, color: '00FF00' });
    const xml = buildParaXml(p);
    expect(xml).toContain('<w:rPr>');
    expect(xml).toMatch(/<w:b(\/>|\s+w:val="1"\s*\/>)/);
    expect(xml).toContain('<w:color w:val="00FF00"/>');
    // No track-change markers should be emitted
    const pPrBlock = xml.substring(xml.indexOf('<w:pPr>'), xml.indexOf('</w:pPr>'));
    expect(pPrBlock).not.toContain('<w:ins ');
    expect(pPrBlock).not.toContain('<w:del ');
  });
});
