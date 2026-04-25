/**
 * w:lvl @w:tentative attribute round-trip tests.
 *
 * Per ECMA-376 Part 1 §17.9.29, `<w:lvl>` carries an optional
 * `w:tentative` attribute (ST_OnOff). When true, the level is
 * "tentative" — present in the definition but not displayed by default
 * until the user chooses it (this is the mechanism Word uses for the
 * unused levels 3–8 in hybrid-multilevel lists so the list picker can
 * offer them without polluting the visible outline).
 *
 * Bug this suite guards against:
 *   - Neither `NumberingLevel.fromXML` nor `toXML` had ANY reference
 *     to `w:tentative`. On load, the attribute was silently discarded,
 *     collapsing tentative and non-tentative levels into the same
 *     on-disk shape after round-trip — so a hybrid-multilevel list
 *     edited and saved through docxmlater reverted every tentative
 *     level to "normal", expanding the list picker.
 */

import { NumberingLevel } from '../../src/formatting/NumberingLevel';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

const buildLvlXml = (extraAttr: string): string =>
  `<w:lvl w:ilvl="0"${extraAttr}><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:hint="default"/></w:rPr></w:lvl>`;

describe('NumberingLevel — w:tentative attribute parsing per ECMA-376 §17.9.29', () => {
  it('parses w:tentative="1" as tentative=true', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(' w:tentative="1"'));
    expect(level.getProperties().tentative).toBe(true);
  });

  it('parses w:tentative="true" as tentative=true', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(' w:tentative="true"'));
    expect(level.getProperties().tentative).toBe(true);
  });

  it('parses w:tentative="on" as tentative=true', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(' w:tentative="on"'));
    expect(level.getProperties().tentative).toBe(true);
  });

  it('parses w:tentative="0" as tentative=false', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(' w:tentative="0"'));
    expect(level.getProperties().tentative).toBe(false);
  });

  it('parses w:tentative="false" as tentative=false', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(' w:tentative="false"'));
    expect(level.getProperties().tentative).toBe(false);
  });

  it('parses w:tentative="off" as tentative=false', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(' w:tentative="off"'));
    expect(level.getProperties().tentative).toBe(false);
  });

  it('leaves tentative undefined when attribute absent (preserves inherit semantics)', () => {
    // Undefined vs false are semantically equivalent at render time (both
    // non-tentative), but the distinction matters for serialization: absent
    // sources should round-trip as absent, not as an explicit "off".
    const level = NumberingLevel.fromXML(buildLvlXml(''));
    expect(level.getProperties().tentative).toBeUndefined();
  });
});

describe('NumberingLevel — w:tentative attribute generation', () => {
  it('emits w:tentative="1" when tentative=true', () => {
    const level = new NumberingLevel({
      level: 3,
      format: 'decimal',
      text: '%4.',
      alignment: 'left',
      start: 1,
      font: 'Calibri',
      fontSize: 22,
      suffix: 'tab',
      tentative: true,
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    // Tentative should appear on the <w:lvl ...> opening tag
    expect(xml).toMatch(/<w:lvl[^>]*w:tentative="1"/);
  });

  it('omits w:tentative when tentative=false (default)', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'decimal',
      text: '%1.',
      alignment: 'left',
      start: 1,
      font: 'Calibri',
      fontSize: 22,
      suffix: 'tab',
      tentative: false,
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).not.toMatch(/w:tentative=/);
  });

  it('omits w:tentative when undefined', () => {
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
    expect(xml).not.toMatch(/w:tentative=/);
  });
});

describe('NumberingLevel — w:tentative full round-trip', () => {
  it('preserves tentative=true through parse → serialize → parse', () => {
    const level1 = NumberingLevel.fromXML(buildLvlXml(' w:tentative="1"'));
    expect(level1.getProperties().tentative).toBe(true);

    const xml = XMLBuilder.elementToString(level1.toXML());
    expect(xml).toMatch(/w:tentative="1"/);

    const level2 = NumberingLevel.fromXML(xml);
    expect(level2.getProperties().tentative).toBe(true);
  });
});
