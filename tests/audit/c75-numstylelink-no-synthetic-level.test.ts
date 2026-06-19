/**
 * Tests that AbstractNumbering.toXML() does not inject a synthetic default
 * level into numStyleLink reference definitions. Per ECMA-376 §17.9.21 a
 * numStyleLink abstractNum is a pure reference to a numbering style and
 * carries no w:lvl children — the linked style supplies the level formats.
 */

import { AbstractNumbering } from '../../src/formatting/AbstractNumbering';
import { NumberingManager } from '../../src/formatting/NumberingManager';
import { NumberingInstance } from '../../src/formatting/NumberingInstance';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('AbstractNumbering.toXML for level-less definitions', () => {
  it('emits no w:lvl children for a numStyleLink reference definition', () => {
    const abstractNum = new AbstractNumbering({
      abstractNumId: 0,
      numStyleLink: 'ListBullet',
    });

    const xml = XMLBuilder.elementToString(abstractNum.toXML());

    expect(xml).toContain('<w:numStyleLink w:val="ListBullet"/>');
    expect(xml).not.toContain('<w:lvl ');
  });

  it('still injects a default level for level-less definitions without numStyleLink', () => {
    const abstractNum = new AbstractNumbering({ abstractNumId: 0 });

    const xml = XMLBuilder.elementToString(abstractNum.toXML());

    expect(xml).toContain('<w:lvl w:ilvl="0"');
    expect(xml).toContain('<w:numFmt w:val="decimal"/>');
  });

  it('keeps numStyleLink definitions level-free through generateNumberingXml', () => {
    const manager = new NumberingManager();
    manager.addAbstractNumbering(
      new AbstractNumbering({ abstractNumId: 0, numStyleLink: 'ListNumber' })
    );
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));

    const xml = manager.generateNumberingXml();

    expect(xml).toContain('<w:numStyleLink w:val="ListNumber"/>');
    expect(xml).not.toContain('<w:lvl ');
  });
});
