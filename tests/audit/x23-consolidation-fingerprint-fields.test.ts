/**
 * Tests that consolidateNumbering's fingerprint accounts for pStyle,
 * lvlPicBulletId, and legacy. All three are render-affecting per ECMA-376
 * (style-to-level linkage §17.9.23, picture bullets §17.9.15, legacy
 * indentation §17.9.10), so abstractNums differing on any of them must not
 * be merged — merging would remap lists to the wrong bullet picture or style
 * linkage and orphan the duplicate's numPicBullet.
 */

import { NumberingManager } from '../../src/formatting/NumberingManager';
import { NumberingInstance } from '../../src/formatting/NumberingInstance';
import { AbstractNumbering } from '../../src/formatting/AbstractNumbering';
import { NumberingLevel, NumberingLevelProperties } from '../../src/formatting/NumberingLevel';

function makeAbstractNum(
  id: number,
  levelOverrides: Partial<NumberingLevelProperties> = {}
): AbstractNumbering {
  const abstractNum = new AbstractNumbering({ abstractNumId: id, multiLevelType: 1 });
  abstractNum.addLevel(
    NumberingLevel.create({
      level: 0,
      format: 'bullet',
      text: '',
      font: 'Symbol',
      ...levelOverrides,
    })
  );
  return abstractNum;
}

function setupManager(abs0: AbstractNumbering, abs1: AbstractNumbering): NumberingManager {
  const manager = new NumberingManager();
  manager.addAbstractNumbering(abs0);
  manager.addAbstractNumbering(abs1);
  manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
  manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
  manager.resetModified();
  return manager;
}

describe('consolidateNumbering fingerprint includes pStyle/lvlPicBulletId/legacy', () => {
  it('does not merge abstractNums differing only in lvlPicBulletId', () => {
    const manager = setupManager(
      makeAbstractNum(0, { text: '', lvlPicBulletId: 0 }),
      makeAbstractNum(1, { text: '', lvlPicBulletId: 1 })
    );

    const result = manager.consolidateNumbering();

    expect(result.abstractNumsRemoved).toBe(0);
    expect(result.instancesRemapped).toBe(0);
    expect(manager.getAbstractNumberingCount()).toBe(2);
    expect(manager.getInstance(2)!.getAbstractNumId()).toBe(1);
  });

  it('does not merge abstractNums differing only in pStyle', () => {
    const manager = setupManager(
      makeAbstractNum(0, { pStyle: 'ListBullet' }),
      makeAbstractNum(1, { pStyle: 'ListBullet2' })
    );

    const result = manager.consolidateNumbering();

    expect(result.abstractNumsRemoved).toBe(0);
    expect(manager.getAbstractNumberingCount()).toBe(2);
  });

  it('does not merge abstractNums differing only in legacy', () => {
    const manager = setupManager(
      makeAbstractNum(0, { legacy: { legacy: true, legacySpace: 144, legacyIndent: 288 } }),
      makeAbstractNum(1)
    );

    const result = manager.consolidateNumbering();

    expect(result.abstractNumsRemoved).toBe(0);
    expect(manager.getAbstractNumberingCount()).toBe(2);
  });

  it('still merges abstractNums whose pStyle/lvlPicBulletId/legacy are identical', () => {
    const manager = setupManager(
      makeAbstractNum(0, { pStyle: 'ListBullet', lvlPicBulletId: 3 }),
      makeAbstractNum(1, { pStyle: 'ListBullet', lvlPicBulletId: 3 })
    );

    const result = manager.consolidateNumbering();

    expect(result.abstractNumsRemoved).toBe(1);
    expect(result.groupsConsolidated).toBe(1);
    expect(manager.getInstance(2)!.getAbstractNumId()).toBe(0);
  });
});
