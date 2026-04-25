/**
 * w:isLgl (legal numbering) round-trip tests.
 *
 * Per ECMA-376 Part 1 §17.9.4, `<w:isLgl>` is a CT_OnOff child of
 * `<w:lvl>`. When present (or `w:val="1"/"true"/"on"`), the numbering
 * level renders using legal numbering style: every level above the
 * current one is formatted using its own numFmt EXCEPT the current
 * level (which renders as Arabic decimal). This is the "1., 2., 3.1.,
 * 3.2." behaviour familiar from legal documents.
 *
 * Bug this suite guards against:
 *   - `NumberingLevel.fromXML` had no w:isLgl parsing at all. A source
 *     document with `<w:isLgl/>` on a level was silently dropped on
 *     load — the level reverted to non-legal format on any round-trip
 *     through docxmlater.
 *
 * The generator already handles the model→XML direction (line ~470:
 * `if (properties.isLegalNumberingStyle) emit <w:isLgl/>`), so the fix
 * is purely on the parse side.
 */

import { NumberingLevel } from '../../src/formatting/NumberingLevel';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

const buildLvlXml = (isLglInner: string): string =>
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>${isLglInner}<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:hint="default"/></w:rPr></w:lvl>`;

describe('NumberingLevel.fromXML — w:isLgl parsing', () => {
  it('parses <w:isLgl/> as isLegalNumberingStyle=true', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:isLgl/>'));
    expect(level.getProperties().isLegalNumberingStyle).toBe(true);
  });

  it('parses <w:isLgl w:val="1"/> as true', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:isLgl w:val="1"/>'));
    expect(level.getProperties().isLegalNumberingStyle).toBe(true);
  });

  it('parses <w:isLgl w:val="true"/> as true', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:isLgl w:val="true"/>'));
    expect(level.getProperties().isLegalNumberingStyle).toBe(true);
  });

  it('parses <w:isLgl w:val="on"/> as true', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:isLgl w:val="on"/>'));
    expect(level.getProperties().isLegalNumberingStyle).toBe(true);
  });

  it('parses <w:isLgl w:val="0"/> as false', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:isLgl w:val="0"/>'));
    expect(level.getProperties().isLegalNumberingStyle).toBe(false);
  });

  it('parses <w:isLgl w:val="false"/> as false', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:isLgl w:val="false"/>'));
    expect(level.getProperties().isLegalNumberingStyle).toBe(false);
  });

  it('parses <w:isLgl w:val="off"/> as false', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('<w:isLgl w:val="off"/>'));
    expect(level.getProperties().isLegalNumberingStyle).toBe(false);
  });

  it('defaults to false when w:isLgl absent', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(''));
    expect(level.getProperties().isLegalNumberingStyle).toBe(false);
  });
});

describe('NumberingLevel w:isLgl round-trip (fromXML → toXML → fromXML)', () => {
  it('preserves w:isLgl=true through a full XML round-trip', () => {
    const level1 = NumberingLevel.fromXML(buildLvlXml('<w:isLgl/>'));
    expect(level1.getProperties().isLegalNumberingStyle).toBe(true);

    // Serialize and re-parse
    const xml2 = XMLBuilder.elementToString(level1.toXML());
    expect(xml2).toContain('<w:isLgl');

    const level2 = NumberingLevel.fromXML(xml2);
    expect(level2.getProperties().isLegalNumberingStyle).toBe(true);
  });

  it('omits w:isLgl when isLegalNumberingStyle=false (default)', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(''));
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).not.toContain('<w:isLgl');
  });
});
