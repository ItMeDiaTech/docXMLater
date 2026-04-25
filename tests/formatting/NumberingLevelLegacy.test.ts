/**
 * NumberingLevel — `<w:legacy>` child (ECMA-376 §17.9.10) round-trip.
 *
 * Per ECMA-376 Part 1 §17.9.10, `<w:lvl>` can carry a `<w:legacy>`
 * child that captures Word 97–era legacy numbering state:
 *
 *   <xsd:complexType name="CT_LvlLegacy">
 *     <xsd:attribute name="legacy" type="ST_OnOff"/>
 *     <xsd:attribute name="legacySpace" type="ST_TwipsMeasure"/>
 *     <xsd:attribute name="legacyIndent" type="ST_SignedTwipsMeasure"/>
 *   </xsd:complexType>
 *
 * Bug this suite guards against:
 *   - `NumberingLevel.fromXML` didn't read `<w:legacy>` at all; `toXML`
 *     didn't emit it. Any level imported from an older Word document
 *     that preserved its legacy-numbering state (hybrid lists, mixed
 *     imports, upgraded templates) lost that state on any round-trip
 *     through docxmlater. The visible symptom: the level's numbering
 *     switched to Word's modern rendering rules, which can change
 *     indentation and bullet/number spacing.
 *
 * This suite covers parse, generate, and full round-trip for all
 * three CT_LvlLegacy attributes.
 */

import { NumberingLevel } from '../../src/formatting/NumberingLevel';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

const buildLvlXml = (legacyInner: string): string =>
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>${legacyInner}<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:hint="default"/></w:rPr></w:lvl>`;

describe('NumberingLevel — w:legacy attribute parsing per ECMA-376 §17.9.10', () => {
  it('parses <w:legacy w:legacy="1"/> as legacy.legacy=true', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:legacy w:legacy="1"/>'));
    expect(level.getProperties().legacy?.legacy).toBe(true);
  });

  it('parses <w:legacy w:legacy="0"/> as legacy.legacy=false', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:legacy w:legacy="0"/>'));
    expect(level.getProperties().legacy?.legacy).toBe(false);
  });

  it('parses <w:legacy w:legacy="on"/> as true (honours ST_OnOff)', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:legacy w:legacy="on"/>'));
    expect(level.getProperties().legacy?.legacy).toBe(true);
  });

  it('parses <w:legacy w:legacy="off"/> as false', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:legacy w:legacy="off"/>'));
    expect(level.getProperties().legacy?.legacy).toBe(false);
  });

  it('parses w:legacySpace and w:legacyIndent as twips integers', () => {
    const level = NumberingLevel.fromXML(
      buildLvlXml('<w:legacy w:legacy="1" w:legacySpace="120" w:legacyIndent="-360"/>')
    );
    const legacy = level.getProperties().legacy;
    expect(legacy?.legacy).toBe(true);
    expect(legacy?.legacySpace).toBe(120);
    expect(legacy?.legacyIndent).toBe(-360);
  });

  it('parses <w:legacy/> bare (no attributes) as an empty object', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:legacy/>'));
    // Element present without attributes → legacy object exists but all fields undefined
    expect(level.getProperties().legacy).toBeDefined();
    expect(level.getProperties().legacy?.legacy).toBeUndefined();
    expect(level.getProperties().legacy?.legacySpace).toBeUndefined();
    expect(level.getProperties().legacy?.legacyIndent).toBeUndefined();
  });

  it('leaves legacy undefined when element absent', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(''));
    expect(level.getProperties().legacy).toBeUndefined();
  });
});

describe('NumberingLevel — w:legacy generation', () => {
  it('emits <w:legacy> with all three attributes when all set', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'decimal',
      text: '%1.',
      alignment: 'left',
      start: 1,
      font: 'Calibri',
      fontSize: 22,
      suffix: 'tab',
      legacy: { legacy: true, legacySpace: 120, legacyIndent: -360 },
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).toMatch(/<w:legacy[^>]*w:legacy="1"/);
    expect(xml).toMatch(/<w:legacy[^>]*w:legacySpace="120"/);
    expect(xml).toMatch(/<w:legacy[^>]*w:legacyIndent="-360"/);
  });

  it('emits only the attributes that are set', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'decimal',
      text: '%1.',
      alignment: 'left',
      start: 1,
      font: 'Calibri',
      fontSize: 22,
      suffix: 'tab',
      legacy: { legacy: true },
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).toMatch(/<w:legacy[^>]*w:legacy="1"/);
    expect(xml).not.toMatch(/w:legacySpace=/);
    expect(xml).not.toMatch(/w:legacyIndent=/);
  });

  it('emits w:legacy="0" for explicit false', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'decimal',
      text: '%1.',
      alignment: 'left',
      start: 1,
      font: 'Calibri',
      fontSize: 22,
      suffix: 'tab',
      legacy: { legacy: false, legacySpace: 0 },
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).toMatch(/<w:legacy[^>]*w:legacy="0"/);
  });

  it('omits <w:legacy> entirely when legacy property is undefined', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'decimal',
      text: '%1.',
      alignment: 'left',
      start: 1,
      font: 'Calibri',
      fontSize: 22,
      suffix: 'tab',
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).not.toContain('<w:legacy');
  });
});

describe('NumberingLevel — w:legacy full round-trip', () => {
  it('preserves legacy=true, legacySpace=120, legacyIndent=-360', () => {
    const src = buildLvlXml('<w:legacy w:legacy="1" w:legacySpace="120" w:legacyIndent="-360"/>');
    const level1 = NumberingLevel.fromXML(src);
    expect(level1.getProperties().legacy).toEqual({
      legacy: true,
      legacySpace: 120,
      legacyIndent: -360,
    });
    const xml = XMLBuilder.elementToString(level1.toXML());
    const level2 = NumberingLevel.fromXML(xml);
    expect(level2.getProperties().legacy).toEqual({
      legacy: true,
      legacySpace: 120,
      legacyIndent: -360,
    });
  });
});
