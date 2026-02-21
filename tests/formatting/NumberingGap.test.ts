/**
 * Numbering gap tests: pStyle, tmpl, full level override
 * Phase 4 of ECMA-376 gap analysis
 */

import { NumberingLevel } from '../../src/formatting/NumberingLevel';
import { AbstractNumbering } from '../../src/formatting/AbstractNumbering';
import { NumberingInstance } from '../../src/formatting/NumberingInstance';
import { Document } from '../../src/core/Document';

describe('NumberingLevel pStyle', () => {
  test('should set and get paragraph style', () => {
    const level = NumberingLevel.createDecimalLevel(0);
    level.setParagraphStyle('ListParagraph');
    expect(level.getParagraphStyle()).toBe('ListParagraph');
  });

  test('should generate w:pStyle in XML', () => {
    const level = NumberingLevel.createDecimalLevel(0);
    level.setParagraphStyle('Heading1');

    const xml = level.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('w:pStyle');
    expect(xmlStr).toContain('Heading1');
  });

  test('should parse pStyle from XML', () => {
    const xmlStr = `<w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:pStyle w:val="ListBullet"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:hint="default"/><w:sz w:val="22"/></w:rPr>
    </w:lvl>`;

    const level = NumberingLevel.fromXML(xmlStr);
    expect(level.getParagraphStyle()).toBe('ListBullet');
  });

  test('should preserve pStyle through AbstractNumbering', () => {
    const level = NumberingLevel.createDecimalLevel(0);
    level.setParagraphStyle('TOCHeading');

    const abstractNum = AbstractNumbering.create({
      abstractNumId: 1,
      levels: [level],
    });

    expect(abstractNum.getLevel(0)?.getParagraphStyle()).toBe('TOCHeading');
  });
});

describe('AbstractNumbering Template', () => {
  test('should set and get template', () => {
    const abstractNum = new AbstractNumbering(1);
    abstractNum.setTemplate('04090001');
    expect(abstractNum.getTemplate()).toBe('04090001');
  });

  test('should generate w:tmpl in XML', () => {
    const abstractNum = AbstractNumbering.create({
      abstractNumId: 1,
      tmpl: 'ABCD1234',
      levels: [NumberingLevel.createDecimalLevel(0)],
    });

    const xml = abstractNum.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('w:tmpl');
    expect(xmlStr).toContain('ABCD1234');
  });

  test('should parse tmpl from XML', () => {
    const xmlStr = `<w:abstractNum w:abstractNumId="0">
      <w:multiLevelType w:val="multilevel"/>
      <w:tmpl w:val="DEADBEEF"/>
      <w:lvl w:ilvl="0">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/>
        <w:lvlJc w:val="left"/>
        <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
        <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:hint="default"/><w:sz w:val="22"/></w:rPr>
      </w:lvl>
    </w:abstractNum>`;

    const abstractNum = AbstractNumbering.fromXML(xmlStr);
    expect(abstractNum.getTemplate()).toBe('DEADBEEF');
  });

  test('should round-trip template through document', async () => {
    const doc = Document.create();
    const nm = doc.getNumberingManager();

    const level = NumberingLevel.createBulletLevel(0);
    const abstractNum = AbstractNumbering.create({
      abstractNumId: 100,
      tmpl: '04090001',
      levels: [level],
    });
    nm.addAbstractNumbering(abstractNum);

    const instance = NumberingInstance.create(100, 100);
    nm.addInstance(instance);

    const para = doc.createParagraph('Bullet item');
    para.setNumbering(100, 0);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedNm = loaded.getNumberingManager();
    const loadedAbstract = loadedNm.getAllAbstractNumberings()
      .find(a => a.getTemplate() === '04090001');
    expect(loadedAbstract).toBeDefined();

    doc.dispose();
    loaded.dispose();
  });
});

describe('NumberingInstance Level Overrides', () => {
  test('should set and get startOverride', () => {
    const instance = new NumberingInstance(1, 0);
    instance.setLevelOverride(0, 5);
    expect(instance.getLevelOverride(0)).toBe(5);
  });

  test('should set and get full level override', () => {
    const instance = new NumberingInstance(1, 0);
    const overrideLevel = NumberingLevel.createDecimalLevel(0);
    overrideLevel.setText('(%1)');

    instance.setFullLevelOverride(0, overrideLevel);
    expect(instance.getFullLevelOverride(0)).toBeDefined();
    expect(instance.getFullLevelOverride(0)?.getProperties().text).toBe('(%1)');
  });

  test('should generate lvlOverride with startOverride in XML', () => {
    const instance = new NumberingInstance(1, 0);
    instance.setLevelOverride(0, 10);

    const xml = instance.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('w:lvlOverride');
    expect(xmlStr).toContain('w:startOverride');
  });

  test('should generate lvlOverride with full w:lvl in XML', () => {
    const instance = new NumberingInstance(1, 0);
    const overrideLevel = NumberingLevel.createBulletLevel(0);

    instance.setFullLevelOverride(0, overrideLevel);

    const xml = instance.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('w:lvlOverride');
    expect(xmlStr).toContain('w:lvl');
    expect(xmlStr).toContain('bullet');
  });

  test('should parse lvlOverride with startOverride from XML', () => {
    const xmlStr = `<w:num w:numId="1">
      <w:abstractNumId w:val="0"/>
      <w:lvlOverride w:ilvl="0">
        <w:startOverride w:val="5"/>
      </w:lvlOverride>
    </w:num>`;

    const instance = NumberingInstance.fromXML(xmlStr);
    expect(instance.getLevelOverride(0)).toBe(5);
  });

  test('should parse lvlOverride with full w:lvl from XML', () => {
    const xmlStr = `<w:num w:numId="2">
      <w:abstractNumId w:val="0"/>
      <w:lvlOverride w:ilvl="1">
        <w:lvl w:ilvl="1">
          <w:start w:val="1"/>
          <w:numFmt w:val="lowerLetter"/>
          <w:lvlText w:val="%2)"/>
          <w:lvlJc w:val="left"/>
          <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
          <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:hint="default"/><w:sz w:val="22"/></w:rPr>
        </w:lvl>
      </w:lvlOverride>
    </w:num>`;

    const instance = NumberingInstance.fromXML(xmlStr);
    const fullOverride = instance.getFullLevelOverride(1);
    expect(fullOverride).toBeDefined();
    expect(fullOverride?.getProperties().format).toBe('lowerLetter');
    expect(fullOverride?.getProperties().text).toBe('%2)');
  });

  test('should clear full level override', () => {
    const instance = new NumberingInstance(1, 0);
    const level = NumberingLevel.createDecimalLevel(0);
    instance.setFullLevelOverride(0, level);
    expect(instance.getFullLevelOverride(0)).toBeDefined();

    instance.clearFullLevelOverride(0);
    expect(instance.getFullLevelOverride(0)).toBeUndefined();
  });

  test('should include both startOverride and full lvl when both set', () => {
    const instance = new NumberingInstance(1, 0);
    instance.setLevelOverride(0, 5);
    const level = NumberingLevel.createDecimalLevel(0);
    instance.setFullLevelOverride(0, level);

    const xml = instance.toXML();
    const xmlStr = JSON.stringify(xml);
    // Full level override should contain both startOverride and lvl
    expect(xmlStr).toContain('w:startOverride');
    expect(xmlStr).toContain('w:lvl');
  });
});

describe('NumberingLevel pStyle round-trip', () => {
  test('should round-trip pStyle through document', async () => {
    const doc = Document.create();
    const nm = doc.getNumberingManager();

    const level = NumberingLevel.createDecimalLevel(0);
    level.setParagraphStyle('ListParagraph');

    const abstractNum = AbstractNumbering.create({
      abstractNumId: 200,
      levels: [level],
    });
    nm.addAbstractNumbering(abstractNum);

    const instance = NumberingInstance.create(200, 200);
    nm.addInstance(instance);

    const para = doc.createParagraph('Styled list item');
    para.setNumbering(200, 0);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedNm = loaded.getNumberingManager();
    // Find abstract numbering that has pStyle
    let foundPStyle = false;
    for (const an of loadedNm.getAllAbstractNumberings()) {
      const lvl0 = an.getLevel(0);
      if (lvl0?.getParagraphStyle() === 'ListParagraph') {
        foundPStyle = true;
        break;
      }
    }
    expect(foundPStyle).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});
