/**
 * NumberingLevel — w:tplc attribute and w:lvlPicBulletId child round-trip.
 *
 * Two more ECMA-376 CT_Lvl members that were silently dropped on parse
 * and could never be regenerated on serialize:
 *
 *   - `w:tplc` (§17.9.30) — ST_ShortHexNumber attribute on `<w:lvl>`.
 *     Word uses this 8-char hex to identify the level template; losing
 *     it means a round-tripped definition will no longer match a sibling
 *     definition in the user's Word list gallery.
 *
 *   - `<w:lvlPicBulletId w:val="…"/>` (§17.9.15) — a child element of
 *     `<w:lvl>` that references a picture-bullet resource defined by
 *     `<w:numPicBullet>` elsewhere in numbering.xml. Parsing silently
 *     dropped the reference, so on save no `<w:lvlPicBulletId>` was
 *     emitted. Worse, `Document.removeOrphanedNumPicBullets` then
 *     treated every `<w:numPicBullet>` as unreferenced and DELETED the
 *     picture-bullet definitions entirely. Net effect: a document with
 *     picture bullets had them silently stripped on any load → save.
 *
 * This suite covers parse + generate for both, plus a full
 * parse → serialize → parse round-trip for each.
 */

import { NumberingLevel } from '../../src/formatting/NumberingLevel';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

const buildLvlXml = (tplcAttr: string, picBulletChild: string): string =>
  `<w:lvl w:ilvl="0"${tplcAttr}><w:start w:val="1"/><w:numFmt w:val="decimal"/>${picBulletChild}<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:hint="default"/></w:rPr></w:lvl>`;

describe('NumberingLevel — w:tplc attribute parsing per ECMA-376 §17.9.30', () => {
  it('parses w:tplc="04090019" as an 8-char hex template code', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(' w:tplc="04090019"', ''));
    expect(level.getProperties().tplc).toBe('04090019');
  });

  it('parses w:tplc="0C0A0017" (with letters)', () => {
    const level = NumberingLevel.fromXML(buildLvlXml(' w:tplc="0C0A0017"', ''));
    expect(level.getProperties().tplc).toBe('0C0A0017');
  });

  it('leaves tplc undefined when attribute absent', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('', ''));
    expect(level.getProperties().tplc).toBeUndefined();
  });
});

describe('NumberingLevel — w:tplc attribute generation', () => {
  it('emits w:tplc attribute when set', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'decimal',
      text: '%1.',
      alignment: 'left',
      start: 1,
      font: 'Calibri',
      fontSize: 22,
      suffix: 'tab',
      tplc: '04090019',
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).toMatch(/<w:lvl[^>]*w:tplc="04090019"/);
  });

  it('omits w:tplc when undefined', () => {
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
    expect(xml).not.toMatch(/w:tplc=/);
  });

  it('round-trips w:tplc through parse → serialize → parse', () => {
    const level1 = NumberingLevel.fromXML(buildLvlXml(' w:tplc="04090019"', ''));
    expect(level1.getProperties().tplc).toBe('04090019');
    const xml = XMLBuilder.elementToString(level1.toXML());
    const level2 = NumberingLevel.fromXML(xml);
    expect(level2.getProperties().tplc).toBe('04090019');
  });
});

describe('NumberingLevel — w:lvlPicBulletId child parsing per ECMA-376 §17.9.15', () => {
  it('parses <w:lvlPicBulletId w:val="0"/> as lvlPicBulletId=0', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('', '<w:lvlPicBulletId w:val="0"/>'));
    expect(level.getProperties().lvlPicBulletId).toBe(0);
  });

  it('parses <w:lvlPicBulletId w:val="5"/> as lvlPicBulletId=5', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('', '<w:lvlPicBulletId w:val="5"/>'));
    expect(level.getProperties().lvlPicBulletId).toBe(5);
  });

  it('leaves lvlPicBulletId undefined when element absent', () => {
    const level = NumberingLevel.fromXML(buildLvlXml('', ''));
    expect(level.getProperties().lvlPicBulletId).toBeUndefined();
  });
});

describe('NumberingLevel — w:lvlPicBulletId child generation', () => {
  it('emits <w:lvlPicBulletId w:val="X"/> when set', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'bullet',
      text: '•',
      alignment: 'left',
      start: 1,
      font: 'Symbol',
      fontSize: 22,
      suffix: 'tab',
      lvlPicBulletId: 3,
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).toContain('<w:lvlPicBulletId w:val="3"/>');
  });

  it('emits w:lvlPicBulletId=0 (valid zero ID)', () => {
    const level = new NumberingLevel({
      level: 0,
      format: 'bullet',
      text: '•',
      alignment: 'left',
      start: 1,
      font: 'Symbol',
      fontSize: 22,
      suffix: 'tab',
      lvlPicBulletId: 0,
    });
    const xml = XMLBuilder.elementToString(level.toXML());
    expect(xml).toContain('<w:lvlPicBulletId w:val="0"/>');
  });

  it('omits w:lvlPicBulletId when undefined', () => {
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
    expect(xml).not.toContain('<w:lvlPicBulletId');
  });

  it('round-trips w:lvlPicBulletId through parse → serialize → parse', () => {
    const level1 = NumberingLevel.fromXML(buildLvlXml('', '<w:lvlPicBulletId w:val="7"/>'));
    expect(level1.getProperties().lvlPicBulletId).toBe(7);
    const xml = XMLBuilder.elementToString(level1.toXML());
    expect(xml).toContain('<w:lvlPicBulletId w:val="7"/>');
    const level2 = NumberingLevel.fromXML(xml);
    expect(level2.getProperties().lvlPicBulletId).toBe(7);
  });
});

describe('NumberingLevel — combined w:tplc + w:lvlPicBulletId round-trip', () => {
  it('preserves both through a full load → save', () => {
    const level1 = NumberingLevel.fromXML(
      buildLvlXml(' w:tplc="04090019"', '<w:lvlPicBulletId w:val="2"/>')
    );
    expect(level1.getProperties().tplc).toBe('04090019');
    expect(level1.getProperties().lvlPicBulletId).toBe(2);
    const xml = XMLBuilder.elementToString(level1.toXML());
    const level2 = NumberingLevel.fromXML(xml);
    expect(level2.getProperties().tplc).toBe('04090019');
    expect(level2.getProperties().lvlPicBulletId).toBe(2);
  });
});
