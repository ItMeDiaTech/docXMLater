/**
 * Tests for NumberingManager.restartNumbering() and Document.restartNumbering()
 */

import { NumberingManager } from '../../src/formatting/NumberingManager';
import { Document } from '../../src/core/Document';
import { XMLElement } from '../../src/xml/XMLBuilder';

/**
 * Helper to filter and safely access XMLElement children
 */
function filterXMLElements(children?: (XMLElement | string)[]): XMLElement[] {
  return (children || []).filter((c): c is XMLElement => typeof c !== 'string');
}

describe('NumberingManager.restartNumbering()', () => {
  let manager: NumberingManager;
  let originalNumId: number;

  beforeEach(() => {
    manager = NumberingManager.create();
    originalNumId = manager.createNumberedList();
  });

  it('should create a new instance with the same abstractNumId', () => {
    const newNumId = manager.restartNumbering(originalNumId);

    expect(newNumId).not.toBe(originalNumId);

    const originalInstance = manager.getInstance(originalNumId);
    const newInstance = manager.getInstance(newNumId);
    expect(newInstance).toBeDefined();
    expect(newInstance!.getAbstractNumId()).toBe(originalInstance!.getAbstractNumId());
  });

  it('should set a startOverride on the new instance (default level 0, value 1)', () => {
    const newNumId = manager.restartNumbering(originalNumId);
    const newInstance = manager.getInstance(newNumId)!;

    const overrides = newInstance.getLevelOverrides();
    expect(overrides.get(0)).toBe(1);
  });

  it('should support custom level and startValue', () => {
    const newNumId = manager.restartNumbering(originalNumId, 2, 5);
    const newInstance = manager.getInstance(newNumId)!;

    const overrides = newInstance.getLevelOverrides();
    expect(overrides.get(2)).toBe(5);
    // Level 0 should not have an override
    expect(overrides.get(0)).toBeUndefined();
  });

  it('should support multiple restarts from the same original', () => {
    const restart1 = manager.restartNumbering(originalNumId);
    const restart2 = manager.restartNumbering(originalNumId);
    const restart3 = manager.restartNumbering(originalNumId, 1, 10);

    // All should be unique
    const ids = [originalNumId, restart1, restart2, restart3];
    expect(new Set(ids).size).toBe(4);

    // All reference the same abstractNumId
    const abstractNumId = manager.getInstance(originalNumId)!.getAbstractNumId();
    expect(manager.getInstance(restart1)!.getAbstractNumId()).toBe(abstractNumId);
    expect(manager.getInstance(restart2)!.getAbstractNumId()).toBe(abstractNumId);
    expect(manager.getInstance(restart3)!.getAbstractNumId()).toBe(abstractNumId);
  });

  it('should throw for non-existent numId', () => {
    expect(() => manager.restartNumbering(9999)).toThrow('Numbering instance 9999 does not exist');
  });

  it('should throw for invalid level (< 0)', () => {
    expect(() => manager.restartNumbering(originalNumId, -1)).toThrow(
      'Invalid level -1. Level must be between 0 and 8.'
    );
  });

  it('should throw for invalid level (> 8)', () => {
    expect(() => manager.restartNumbering(originalNumId, 9)).toThrow(
      'Invalid level 9. Level must be between 0 and 8.'
    );
  });

  it('should throw for invalid startValue (< 1)', () => {
    expect(() => manager.restartNumbering(originalNumId, 0, 0)).toThrow(
      'Invalid startValue 0. Start value must be at least 1.'
    );
  });

  it('should generate XML with lvlOverride and startOverride', () => {
    const newNumId = manager.restartNumbering(originalNumId);
    const newInstance = manager.getInstance(newNumId)!;

    const xml = newInstance.toXML();
    const children = filterXMLElements(xml.children);

    // Should have abstractNumId reference + lvlOverride
    const abstractNumRef = children.find((c) => c.name === 'w:abstractNumId');
    expect(abstractNumRef).toBeDefined();

    const lvlOverride = children.find((c) => c.name === 'w:lvlOverride');
    expect(lvlOverride).toBeDefined();
    expect(lvlOverride!.attributes?.['w:ilvl']).toBe('0');

    const overrideChildren = filterXMLElements(lvlOverride!.children);
    const startOverride = overrideChildren.find((c) => c.name === 'w:startOverride');
    expect(startOverride).toBeDefined();
    expect(startOverride!.attributes?.['w:val']).toBe('1');
  });

  it('should generate XML with custom level and startValue', () => {
    const newNumId = manager.restartNumbering(originalNumId, 3, 42);
    const newInstance = manager.getInstance(newNumId)!;

    const xml = newInstance.toXML();
    const children = filterXMLElements(xml.children);

    const lvlOverride = children.find((c) => c.name === 'w:lvlOverride');
    expect(lvlOverride).toBeDefined();
    expect(lvlOverride!.attributes?.['w:ilvl']).toBe('3');

    const overrideChildren = filterXMLElements(lvlOverride!.children);
    const startOverride = overrideChildren.find((c) => c.name === 'w:startOverride');
    expect(startOverride!.attributes?.['w:val']).toBe('42');
  });

  it('should mark numbering as modified', () => {
    manager.resetModified();
    expect(manager.isModified()).toBe(false);

    const newNumId = manager.restartNumbering(originalNumId);

    expect(manager.isModified()).toBe(true);
    expect(manager.getModifiedNumIds().has(newNumId)).toBe(true);
  });
});

describe('Document.restartNumbering()', () => {
  it('should delegate to numberingManager and return new numId', () => {
    const doc = Document.create();
    const listId = doc.createNumberedList();

    const restartId = doc.restartNumbering(listId);

    expect(restartId).not.toBe(listId);
    expect(typeof restartId).toBe('number');

    doc.dispose();
  });

  it('should work with paragraphs using setNumbering', async () => {
    const doc = Document.create();
    const listId = doc.createNumberedList();

    doc.createParagraph('Item 1').setNumbering(listId, 0);
    doc.createParagraph('Item 2').setNumbering(listId, 0);

    const restartId = doc.restartNumbering(listId);
    doc.createParagraph('Restarted item 1').setNumbering(restartId, 0);

    // Verify the document can be saved without errors
    const buffer = await doc.toBuffer();
    expect(buffer).toBeDefined();
    expect(buffer.length).toBeGreaterThan(0);

    doc.dispose();
  });

  it('should round-trip through save and load', async () => {
    const doc = Document.create();
    const listId = doc.createNumberedList();

    doc.createParagraph('Item 1').setNumbering(listId, 0);
    const restartId = doc.restartNumbering(listId);
    doc.createParagraph('Restarted 1').setNumbering(restartId, 0);

    const buffer = await doc.toBuffer();
    doc.dispose();

    // Load and verify the restart numbering instance is preserved
    const loaded = await Document.loadFromBuffer(buffer);
    const instances = loaded.getNumberingManager().getAllInstances();

    // Find the restart instance (should have a level override)
    const restartInstance = instances.find((inst) => {
      const overrides = inst.getLevelOverrides();
      return overrides.size > 0 && overrides.get(0) === 1;
    });
    expect(restartInstance).toBeDefined();

    loaded.dispose();
  });
});
