/**
 * Tests for numbering definition consolidation
 *
 * Covers:
 * - setAbstractNumId() on NumberingInstance
 * - cleanupUnusedNumbering() dirty-tracking fix
 * - consolidateNumbering() core logic
 * - Protected IDs
 * - End-to-end via Document
 * - Fingerprint correctness
 */

import { NumberingManager } from '../../src/formatting/NumberingManager';
import { NumberingInstance } from '../../src/formatting/NumberingInstance';
import { AbstractNumbering } from '../../src/formatting/AbstractNumbering';
import { NumberingLevel } from '../../src/formatting/NumberingLevel';
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: create a bullet abstractNum with standard levels
 */
function createBulletAbstractNum(id: number, name?: string): AbstractNumbering {
  return AbstractNumbering.createBulletList(id);
}

/**
 * Helper: create a numbered abstractNum with standard levels
 */
function createNumberedAbstractNum(id: number): AbstractNumbering {
  return AbstractNumbering.createNumberedList(id);
}

/**
 * Helper: create a bullet abstractNum with a custom font on level 0
 */
function createCustomBulletAbstractNum(id: number, font: string): AbstractNumbering {
  const abstractNum = new AbstractNumbering({ abstractNumId: id, multiLevelType: 1 });
  for (let i = 0; i < 9; i++) {
    abstractNum.addLevel(NumberingLevel.createBulletLevel(i, undefined, i === 0 ? font : undefined));
  }
  return abstractNum;
}

describe('NumberingInstance.setAbstractNumId()', () => {
  it('should remap instance to a different abstractNum', () => {
    const instance = NumberingInstance.create({ numId: 1, abstractNumId: 5 });
    expect(instance.getAbstractNumId()).toBe(5);

    instance.setAbstractNumId(10);
    expect(instance.getAbstractNumId()).toBe(10);
  });

  it('should reject negative IDs', () => {
    const instance = NumberingInstance.create({ numId: 1, abstractNumId: 5 });
    expect(() => instance.setAbstractNumId(-1)).toThrow('Abstract numbering ID must be non-negative');
  });

  it('should support method chaining', () => {
    const instance = NumberingInstance.create({ numId: 1, abstractNumId: 5 });
    const result = instance.setAbstractNumId(10);
    expect(result).toBe(instance);
  });
});

describe('cleanupUnusedNumbering() dirty-tracking fix', () => {
  it('should set isModified() after cleanup removes items', () => {
    const manager = new NumberingManager();

    // Add two abstractNums with instances
    const abs0 = createBulletAbstractNum(0);
    const abs1 = createBulletAbstractNum(1);
    manager.addAbstractNumbering(abs0);
    manager.addAbstractNumbering(abs1);
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));

    // Reset modified (simulating post-parse state)
    manager.resetModified();
    expect(manager.isModified()).toBe(false);

    // Only numId 1 is used; numId 2 and abstractNum 1 are unused
    const usedNumIds = new Set([1]);
    manager.cleanupUnusedNumbering(usedNumIds);

    expect(manager.isModified()).toBe(true);
  });

  it('should populate removedAbstractNumIds and removedNumIds', () => {
    const manager = new NumberingManager();

    const abs0 = createBulletAbstractNum(0);
    const abs1 = createBulletAbstractNum(1);
    manager.addAbstractNumbering(abs0);
    manager.addAbstractNumbering(abs1);
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
    manager.resetModified();

    // Only numId 1 is used
    manager.cleanupUnusedNumbering(new Set([1]));

    expect(manager.getRemovedNumIds().has(2)).toBe(true);
    expect(manager.getRemovedAbstractNumIds().has(1)).toBe(true);
  });

  it('should not set modified when nothing is removed', () => {
    const manager = new NumberingManager();

    const abs0 = createBulletAbstractNum(0);
    manager.addAbstractNumbering(abs0);
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.resetModified();

    // Both are used
    manager.cleanupUnusedNumbering(new Set([1]));
    expect(manager.isModified()).toBe(false);
  });
});

describe('consolidateNumbering()', () => {
  it('should merge 3 identical bullet abstractNums into 1', () => {
    const manager = new NumberingManager();

    // Create 3 identical bullet abstractNums
    for (let i = 0; i < 3; i++) {
      manager.addAbstractNumbering(createBulletAbstractNum(i));
      manager.addInstance(NumberingInstance.create({ numId: i + 1, abstractNumId: i }));
    }
    manager.resetModified();

    const result = manager.consolidateNumbering();

    expect(result.abstractNumsRemoved).toBe(2);
    expect(result.instancesRemapped).toBe(2);
    expect(result.groupsConsolidated).toBe(1);

    // All instances should now point to abstractNum 0 (lowest ID)
    for (const instance of manager.getAllInstances()) {
      expect(instance.getAbstractNumId()).toBe(0);
    }

    // Only abstractNum 0 should remain
    expect(manager.getAbstractNumberingCount()).toBe(1);
    expect(manager.hasAbstractNumbering(0)).toBe(true);
  });

  it('should consolidate bullet and numbered groups independently', () => {
    const manager = new NumberingManager();

    // 2 identical bullet abstractNums
    manager.addAbstractNumbering(createBulletAbstractNum(0));
    manager.addAbstractNumbering(createBulletAbstractNum(1));
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));

    // 2 identical numbered abstractNums
    manager.addAbstractNumbering(createNumberedAbstractNum(2));
    manager.addAbstractNumbering(createNumberedAbstractNum(3));
    manager.addInstance(NumberingInstance.create({ numId: 3, abstractNumId: 2 }));
    manager.addInstance(NumberingInstance.create({ numId: 4, abstractNumId: 3 }));

    manager.resetModified();

    const result = manager.consolidateNumbering();

    // 2 groups consolidated, 1 removed from each
    expect(result.groupsConsolidated).toBe(2);
    expect(result.abstractNumsRemoved).toBe(2);
    expect(manager.getAbstractNumberingCount()).toBe(2);
  });

  it('should leave single-member groups untouched', () => {
    const manager = new NumberingManager();

    // 1 bullet, 1 numbered — different fingerprints
    manager.addAbstractNumbering(createBulletAbstractNum(0));
    manager.addAbstractNumbering(createNumberedAbstractNum(1));
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
    manager.resetModified();

    const result = manager.consolidateNumbering();

    expect(result.abstractNumsRemoved).toBe(0);
    expect(result.instancesRemapped).toBe(0);
    expect(result.groupsConsolidated).toBe(0);
    expect(manager.getAbstractNumberingCount()).toBe(2);
  });

  it('should return correct result counts', () => {
    const manager = new NumberingManager();

    // 5 identical bullet abstractNums, each with an instance
    for (let i = 0; i < 5; i++) {
      manager.addAbstractNumbering(createBulletAbstractNum(i));
      manager.addInstance(NumberingInstance.create({ numId: i + 1, abstractNumId: i }));
    }
    manager.resetModified();

    const result = manager.consolidateNumbering();

    expect(result.abstractNumsRemoved).toBe(4);
    expect(result.instancesRemapped).toBe(4);
    expect(result.groupsConsolidated).toBe(1);
  });

  it('should track _removedAbstractNumIds and _modifiedNumIds correctly', () => {
    const manager = new NumberingManager();

    manager.addAbstractNumbering(createBulletAbstractNum(0));
    manager.addAbstractNumbering(createBulletAbstractNum(1));
    manager.addAbstractNumbering(createBulletAbstractNum(2));
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
    manager.addInstance(NumberingInstance.create({ numId: 3, abstractNumId: 2 }));
    manager.resetModified();

    manager.consolidateNumbering();

    // AbstractNums 1 and 2 should be removed
    const removedAbstracts = manager.getRemovedAbstractNumIds();
    expect(removedAbstracts.has(1)).toBe(true);
    expect(removedAbstracts.has(2)).toBe(true);
    expect(removedAbstracts.has(0)).toBe(false);

    // Instances 2 and 3 should be marked modified (remapped)
    const modifiedNums = manager.getModifiedNumIds();
    expect(modifiedNums.has(2)).toBe(true);
    expect(modifiedNums.has(3)).toBe(true);

    expect(manager.isModified()).toBe(true);
  });
});

describe('Protected IDs', () => {
  it('should exclude protected abstractNums from consolidation', () => {
    const manager = new NumberingManager();

    // 3 identical bullet abstractNums, but id 1 is protected
    for (let i = 0; i < 3; i++) {
      manager.addAbstractNumbering(createBulletAbstractNum(i));
      manager.addInstance(NumberingInstance.create({ numId: i + 1, abstractNumId: i }));
    }
    manager.resetModified();

    const result = manager.consolidateNumbering({
      protectedAbstractNumIds: new Set([1]),
    });

    // Only ids 0 and 2 can consolidate; id 1 is excluded
    expect(result.abstractNumsRemoved).toBe(1);
    expect(result.instancesRemapped).toBe(1);

    // AbstractNum 1 should still exist
    expect(manager.hasAbstractNumbering(1)).toBe(true);
    // Instance 2 (pointed to abstractNum 1) should not be remapped
    expect(manager.getInstance(2)?.getAbstractNumId()).toBe(1);
  });

  it('should still consolidate non-protected identical ones', () => {
    const manager = new NumberingManager();

    // 4 identical, protect id 0
    for (let i = 0; i < 4; i++) {
      manager.addAbstractNumbering(createBulletAbstractNum(i));
      manager.addInstance(NumberingInstance.create({ numId: i + 1, abstractNumId: i }));
    }
    manager.resetModified();

    const result = manager.consolidateNumbering({
      protectedAbstractNumIds: new Set([0]),
    });

    // ids 1,2,3 consolidate → canonical 1, remove 2,3
    expect(result.abstractNumsRemoved).toBe(2);
    expect(manager.hasAbstractNumbering(0)).toBe(true);
    expect(manager.hasAbstractNumbering(1)).toBe(true);
    expect(manager.hasAbstractNumbering(2)).toBe(false);
    expect(manager.hasAbstractNumbering(3)).toBe(false);
  });
});

describe('End-to-end via Document', () => {
  it('should consolidate duplicate abstractNums, save, and reload with fewer', async () => {
    const doc = Document.create();

    const manager = doc.getNumberingManager();

    // Create 3 identical bullet lists
    const numId1 = manager.createBulletList();
    const numId2 = manager.createBulletList();
    const numId3 = manager.createBulletList();

    // Add paragraphs referencing each
    const p1 = new Paragraph();
    p1.addText('Item A');
    p1.setNumbering(numId1, 0);
    doc.addParagraph(p1);

    const p2 = new Paragraph();
    p2.addText('Item B');
    p2.setNumbering(numId2, 0);
    doc.addParagraph(p2);

    const p3 = new Paragraph();
    p3.addText('Item C');
    p3.setNumbering(numId3, 0);
    doc.addParagraph(p3);

    // Before consolidation: 3 abstractNums
    expect(manager.getAbstractNumberingCount()).toBe(3);

    // Consolidate
    const result = doc.consolidateNumbering();
    expect(result.abstractNumsRemoved).toBe(2);

    // After: 1 abstractNum
    expect(manager.getAbstractNumberingCount()).toBe(1);

    // Save and reload
    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    const reloadedManager = reloaded.getNumberingManager();

    // Reloaded should have only 1 abstractNum
    expect(reloadedManager.getAbstractNumberingCount()).toBe(1);

    doc.dispose();
    reloaded.dispose();
  });

  it('should work correctly with consolidation + cleanupUnusedNumbering in sequence', () => {
    const doc = Document.create();
    const manager = doc.getNumberingManager();

    // Create 3 identical bullets + 1 unused numbered
    const numId1 = manager.createBulletList();
    const numId2 = manager.createBulletList();
    const numId3 = manager.createBulletList();
    const _unusedNumId = manager.createNumberedList();

    // Only reference numId1 and numId2
    const p1 = new Paragraph();
    p1.addText('Item A');
    p1.setNumbering(numId1, 0);
    doc.addParagraph(p1);

    const p2 = new Paragraph();
    p2.addText('Item B');
    p2.setNumbering(numId2, 0);
    doc.addParagraph(p2);

    // Consolidate first (merges 3 bullet abstractNums → 1)
    const consolidateResult = doc.consolidateNumbering();
    expect(consolidateResult.abstractNumsRemoved).toBe(2);

    // Now cleanup unused (removes the unreferenced numbered list + numId3's instance)
    doc.cleanupUnusedNumbering();

    // Should be left with 1 abstractNum and 2 instances
    expect(manager.getAbstractNumberingCount()).toBe(1);
    expect(manager.getInstanceCount()).toBe(2);

    doc.dispose();
  });
});

describe('Fingerprint correctness', () => {
  it('should merge abstractNums with same levels but different names', () => {
    const manager = new NumberingManager();

    // Two identical definitions with different names
    const abs0 = createBulletAbstractNum(0);
    abs0.setName('My Custom Bullets');
    const abs1 = createBulletAbstractNum(1);
    abs1.setName('Another Bullet List');

    manager.addAbstractNumbering(abs0);
    manager.addAbstractNumbering(abs1);
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
    manager.resetModified();

    const result = manager.consolidateNumbering();

    // Names are excluded from fingerprint → should consolidate
    expect(result.abstractNumsRemoved).toBe(1);
    expect(result.groupsConsolidated).toBe(1);
  });

  it('should NOT merge abstractNums with different level properties', () => {
    const manager = new NumberingManager();

    // Standard bullet
    manager.addAbstractNumbering(createBulletAbstractNum(0));
    // Custom bullet with different font on level 0
    manager.addAbstractNumbering(createCustomBulletAbstractNum(1, 'Arial'));

    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
    manager.resetModified();

    const result = manager.consolidateNumbering();

    // Different fonts → different fingerprints → no consolidation
    expect(result.abstractNumsRemoved).toBe(0);
    expect(result.groupsConsolidated).toBe(0);
    expect(manager.getAbstractNumberingCount()).toBe(2);
  });

  it('should distinguish bullet vs numbered even with same level count', () => {
    const manager = new NumberingManager();

    manager.addAbstractNumbering(createBulletAbstractNum(0));
    manager.addAbstractNumbering(createNumberedAbstractNum(1));

    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
    manager.resetModified();

    const result = manager.consolidateNumbering();

    expect(result.abstractNumsRemoved).toBe(0);
    expect(manager.getAbstractNumberingCount()).toBe(2);
  });
});

/**
 * Helper: Creates a DOCX buffer with custom numbering.xml and optional numbering.xml.rels
 */
async function createDocxWithNumbering(
  numberingXml: string,
  numberingRels?: string
): Promise<Buffer> {
  const doc = Document.create();
  const para = new Paragraph().addText('Test item');
  doc.addParagraph(para);
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  // Use addFile (not updateFile) because the base doc may not have these files yet
  zipHandler.addFile(DOCX_PATHS.NUMBERING, numberingXml);
  if (numberingRels) {
    zipHandler.addFile(DOCX_PATHS.NUMBERING_RELS, numberingRels);
    // Add minimal 1x1 PNG files so the validator doesn't report missing parts
    const minimalPng = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAB' +
      'Nl7BcQAAAABJRU5ErkJggg==',
      'base64'
    );
    zipHandler.addFile('word/media/image1.png', minimalPng);
    zipHandler.addFile('word/media/image2.png', minimalPng);
  }
  return await zipHandler.toBuffer();
}

describe('numPicBullet cleanup', () => {
  /** Numbering XML with 2 numPicBullets, 2 abstractNums (one uses picBullet, one doesn't), 2 nums */
  const NUMBERING_WITH_PIC_BULLETS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
             xmlns:v="urn:schemas-microsoft-com:vml">
  <w:numPicBullet w:numPicBulletId="0">
    <w:pict><v:shape><v:imagedata r:id="rId1"/></v:shape></w:pict>
  </w:numPicBullet>
  <w:numPicBullet w:numPicBulletId="1">
    <w:pict><v:shape><v:imagedata r:id="rId2"/></v:shape></w:pict>
  </w:numPicBullet>
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlPicBulletId w:val="0"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

  const NUMBERING_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
</Relationships>`;

  it('should remove orphaned numPicBullets when their abstractNum is removed', async () => {
    const buffer = await createDocxWithNumbering(NUMBERING_WITH_PIC_BULLETS, NUMBERING_RELS);

    // Load doc — only numId 2 (abstractNum 1, no picBullet) is used by a paragraph
    const doc = await Document.loadFromBuffer(buffer);
    const p = new Paragraph().addText('Bullet item');
    p.setNumbering(2, 0);
    doc.addParagraph(p);

    // Cleanup: numId 1 (abstractNum 0, which uses picBullet 0) is unused
    doc.cleanupUnusedNumbering();

    // Save and inspect
    const output = await doc.toBuffer();
    const zip = new ZipHandler();
    await zip.loadFromBuffer(output);
    const numbXml = zip.getFileAsString(DOCX_PATHS.NUMBERING)!;

    // AbstractNum 0 should be gone (orphaned)
    expect(numbXml).not.toContain('w:abstractNumId="0"');
    // numPicBullet 0 should be gone (referenced only by removed abstractNum 0)
    expect(numbXml).not.toContain('w:numPicBulletId="0"');
    // numPicBullet 1 was never referenced — should also be gone
    expect(numbXml).not.toContain('w:numPicBulletId="1"');
    // AbstractNum 1 and num 2 should still be present
    expect(numbXml).toContain('w:abstractNumId="1"');

    doc.dispose();
  });

  it('should preserve numPicBullets that are still referenced', async () => {
    // Add a third unused abstractNum/num so cleanupUnusedNumbering() removes something,
    // triggering the merge path where removeOrphanedNumPicBullets() runs
    const xmlWithExtra = NUMBERING_WITH_PIC_BULLETS.replace(
      '</w:numbering>',
      `  <w:abstractNum w:abstractNumId="99">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="99"><w:abstractNumId w:val="99"/></w:num>
</w:numbering>`
    );
    const buffer = await createDocxWithNumbering(xmlWithExtra, NUMBERING_RELS);

    // Load doc — numIds 1 and 2 are used; numId 99 is NOT used
    const doc = await Document.loadFromBuffer(buffer);
    const p1 = new Paragraph().addText('Pic bullet item');
    p1.setNumbering(1, 0);
    doc.addParagraph(p1);

    const p2 = new Paragraph().addText('Symbol bullet item');
    p2.setNumbering(2, 0);
    doc.addParagraph(p2);

    // Cleanup removes unused numId 99 → triggers merge path
    doc.cleanupUnusedNumbering();

    const output = await doc.toBuffer();
    const zip = new ZipHandler();
    await zip.loadFromBuffer(output);
    const numbXml = zip.getFileAsString(DOCX_PATHS.NUMBERING)!;

    // picBullet 0 is referenced by abstractNum 0 which is used → should be preserved
    expect(numbXml).toContain('w:numPicBulletId="0"');
    // picBullet 1 is NOT referenced by any abstractNum → should be removed
    expect(numbXml).not.toContain('w:numPicBulletId="1"');
    // Unused abstractNum 99 should be gone
    expect(numbXml).not.toContain('w:abstractNumId="99"');

    doc.dispose();
  });
});

describe('numbering.xml.rels cleanup', () => {
  const NUMBERING_WITH_PIC_BULLETS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:v="urn:schemas-microsoft-com:vml">
  <w:numPicBullet w:numPicBulletId="0">
    <w:pict><v:shape><v:imagedata r:id="rId1"/></v:shape></w:pict>
  </w:numPicBullet>
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

  const NUMBERING_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

  it('should remove orphaned relationships when numPicBullets are removed', async () => {
    const buffer = await createDocxWithNumbering(NUMBERING_WITH_PIC_BULLETS, NUMBERING_RELS);

    // Load and don't use numId 1 → abstractNum 0 and its picBullet become orphaned
    const doc = await Document.loadFromBuffer(buffer);
    doc.cleanupUnusedNumbering();

    const output = await doc.toBuffer();
    const zip = new ZipHandler();
    await zip.loadFromBuffer(output);

    // numbering.xml.rels should be removed (all relationships were orphaned)
    const relsXml = zip.getFileAsString(DOCX_PATHS.NUMBERING_RELS);
    // Either the file is gone or it has no Relationship entries
    if (relsXml) {
      expect(relsXml).not.toContain('rId1');
    }

    doc.dispose();
  });

  it('should not touch numbering.xml.rels when numPicBullets are preserved', async () => {
    const buffer = await createDocxWithNumbering(NUMBERING_WITH_PIC_BULLETS, NUMBERING_RELS);

    // Load and USE numId 1 → abstractNum 0 is preserved → its picBullet stays
    const doc = await Document.loadFromBuffer(buffer);
    const p = new Paragraph().addText('Bullet item');
    p.setNumbering(1, 0);
    doc.addParagraph(p);

    // No cleanup — everything is in use
    const output = await doc.toBuffer();
    const zip = new ZipHandler();
    await zip.loadFromBuffer(output);

    // rels should still have rId1 since the picBullet referencing it was NOT removed
    // (No cleanup was called, so numbering is unmodified and uses passthrough)
    const relsXml = zip.getFileAsString(DOCX_PATHS.NUMBERING_RELS);
    expect(relsXml).toBeTruthy();
    expect(relsXml).toContain('rId1');

    doc.dispose();
  });
});

describe('Full cleanup pipeline (cleanup → consolidate → validate)', () => {
  it('should produce clean numbering after full pipeline', async () => {
    const doc = Document.create();
    const manager = doc.getNumberingManager();

    // Create 5 identical bullet lists and 2 identical numbered lists
    const bulletIds: number[] = [];
    for (let i = 0; i < 5; i++) {
      bulletIds.push(manager.createBulletList());
    }
    const numIds: number[] = [];
    for (let i = 0; i < 2; i++) {
      numIds.push(manager.createNumberedList());
    }

    // Only reference 2 bullet lists and 1 numbered list
    const p1 = new Paragraph().addText('Bullet A');
    p1.setNumbering(bulletIds[0]!, 0);
    doc.addParagraph(p1);

    const p2 = new Paragraph().addText('Bullet B');
    p2.setNumbering(bulletIds[1]!, 0);
    doc.addParagraph(p2);

    const p3 = new Paragraph().addText('Number 1');
    p3.setNumbering(numIds[0]!, 0);
    doc.addParagraph(p3);

    // Before: 7 abstractNums, 7 instances
    expect(manager.getAbstractNumberingCount()).toBe(7);

    // Phase 1: Remove orphaned
    doc.cleanupUnusedNumbering();
    // After cleanup: 3 abstractNums (2 bullet + 1 numbered), 3 instances
    expect(manager.getAbstractNumberingCount()).toBe(3);
    expect(manager.getInstanceCount()).toBe(3);

    // Phase 2: Consolidate duplicates
    const result = doc.consolidateNumbering();
    // The 2 bullet abstractNums are identical → consolidated to 1
    expect(result.abstractNumsRemoved).toBe(1);
    expect(manager.getAbstractNumberingCount()).toBe(2); // 1 bullet + 1 numbered

    // Phase 3: Validate references
    const orphaned = doc.validateNumberingReferences();
    expect(orphaned).toBe(0); // No orphaned refs expected

    // Save and reload
    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    const reloadedMgr = reloaded.getNumberingManager();

    expect(reloadedMgr.getAbstractNumberingCount()).toBe(2);
    expect(reloadedMgr.getInstanceCount()).toBe(3); // 2 bullet instances + 1 numbered

    doc.dispose();
    reloaded.dispose();
  });
});

describe('cleanupUnusedNumbering() header/footer/footnote/endnote scanning', () => {
  it('should preserve numIds referenced only in a header', () => {
    const doc = Document.create();
    const manager = doc.getNumberingManager();

    // Create two lists — one used in body, one only in header
    const bodyNumId = manager.createBulletList();
    const headerNumId = manager.createNumberedList();

    // Body paragraph uses bodyNumId
    const bodyPara = new Paragraph().addText('Body item');
    bodyPara.setNumbering(bodyNumId, 0);
    doc.addParagraph(bodyPara);

    // Header paragraph uses headerNumId
    const header = new Header();
    const headerPara = new Paragraph().addText('Header item');
    headerPara.setNumbering(headerNumId, 0);
    header.addParagraph(headerPara);
    doc.setHeader(header);

    // Before: 2 abstractNums, 2 instances
    expect(manager.getAbstractNumberingCount()).toBe(2);

    // Cleanup should preserve both — headerNumId is in use via header
    doc.cleanupUnusedNumbering();

    expect(manager.getInstanceCount()).toBe(2);
    expect(manager.getAbstractNumberingCount()).toBe(2);
    doc.dispose();
  });

  it('should preserve numIds referenced only in a footer', () => {
    const doc = Document.create();
    const manager = doc.getNumberingManager();

    const bodyNumId = manager.createBulletList();
    const footerNumId = manager.createNumberedList();

    const bodyPara = new Paragraph().addText('Body item');
    bodyPara.setNumbering(bodyNumId, 0);
    doc.addParagraph(bodyPara);

    const footer = new Footer();
    const footerPara = new Paragraph().addText('Footer item');
    footerPara.setNumbering(footerNumId, 0);
    footer.addParagraph(footerPara);
    doc.setFooter(footer);

    doc.cleanupUnusedNumbering();

    expect(manager.getInstanceCount()).toBe(2);
    expect(manager.getAbstractNumberingCount()).toBe(2);
    doc.dispose();
  });

  it('should preserve numIds referenced only in footnotes', () => {
    const doc = Document.create();
    const manager = doc.getNumberingManager();

    const bodyNumId = manager.createBulletList();
    const footnoteNumId = manager.createNumberedList();

    const bodyPara = new Paragraph().addText('Body item');
    bodyPara.setNumbering(bodyNumId, 0);
    doc.addParagraph(bodyPara);

    // Create a footnote and manually add a numbered paragraph
    const footnote = doc.createFootnote('Footnote text');
    const footnotePara = new Paragraph().addText('Footnote list item');
    footnotePara.setNumbering(footnoteNumId, 0);
    footnote.addParagraph(footnotePara);

    doc.cleanupUnusedNumbering();

    expect(manager.getInstanceCount()).toBe(2);
    expect(manager.getAbstractNumberingCount()).toBe(2);
    doc.dispose();
  });

  it('should preserve numIds referenced only in endnotes', () => {
    const doc = Document.create();
    const manager = doc.getNumberingManager();

    const bodyNumId = manager.createBulletList();
    const endnoteNumId = manager.createNumberedList();

    const bodyPara = new Paragraph().addText('Body item');
    bodyPara.setNumbering(bodyNumId, 0);
    doc.addParagraph(bodyPara);

    // Create an endnote and add a numbered paragraph
    const endnote = doc.createEndnote('Endnote text');
    const endnotePara = new Paragraph().addText('Endnote list item');
    endnotePara.setNumbering(endnoteNumId, 0);
    endnote.addParagraph(endnotePara);

    doc.cleanupUnusedNumbering();

    expect(manager.getInstanceCount()).toBe(2);
    expect(manager.getAbstractNumberingCount()).toBe(2);
    doc.dispose();
  });
});

describe('numStyleLink / styleLink support', () => {
  it('should parse numStyleLink from XML', () => {
    const xml = `<w:abstractNum w:abstractNumId="5">
      <w:multiLevelType w:val="multilevel"/>
      <w:numStyleLink w:val="ListBullet"/>
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    </w:abstractNum>`;
    const abstractNum = AbstractNumbering.fromXML(xml);

    expect(abstractNum.getNumStyleLink()).toBe('ListBullet');
    expect(abstractNum.getStyleLink()).toBeUndefined();
    expect(abstractNum.getAbstractNumId()).toBe(5);
  });

  it('should parse styleLink from XML', () => {
    const xml = `<w:abstractNum w:abstractNumId="7">
      <w:multiLevelType w:val="multilevel"/>
      <w:styleLink w:val="MyListStyle"/>
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    </w:abstractNum>`;
    const abstractNum = AbstractNumbering.fromXML(xml);

    expect(abstractNum.getStyleLink()).toBe('MyListStyle');
    expect(abstractNum.getNumStyleLink()).toBeUndefined();
  });

  it('should serialize numStyleLink and styleLink in toXML()', () => {
    const abstractNum = new AbstractNumbering({
      abstractNumId: 10,
      numStyleLink: 'ListBullet',
      styleLink: 'MyListStyle',
      multiLevelType: 1,
    });
    abstractNum.addLevel(NumberingLevel.createDecimalLevel(0));

    const xmlElement = abstractNum.toXML();
    // Convert to string for assertion — use XMLBuilder to render
    const builder = new XMLBuilder();
    builder.element(xmlElement.name, xmlElement.attributes, xmlElement.children);
    const xmlStr = builder.build(false);

    expect(xmlStr).toContain('w:numStyleLink');
    expect(xmlStr).toContain('w:val="ListBullet"');
    expect(xmlStr).toContain('w:styleLink');
    expect(xmlStr).toContain('w:val="MyListStyle"');
  });

  it('should NOT consolidate abstractNums with different numStyleLink', () => {
    const manager = new NumberingManager();

    // Two otherwise identical bullet abstractNums with different numStyleLink
    const abs0 = createBulletAbstractNum(0);
    abs0.setNumStyleLink('ListBullet');
    const abs1 = createBulletAbstractNum(1);
    abs1.setNumStyleLink('ListBullet2');

    manager.addAbstractNumbering(abs0);
    manager.addAbstractNumbering(abs1);
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
    manager.resetModified();

    const result = manager.consolidateNumbering();

    // Different numStyleLink → different fingerprints → no consolidation
    expect(result.abstractNumsRemoved).toBe(0);
    expect(result.groupsConsolidated).toBe(0);
    expect(manager.getAbstractNumberingCount()).toBe(2);
  });

  it('should consolidate abstractNums with same numStyleLink', () => {
    const manager = new NumberingManager();

    // Two identical bullet abstractNums with the same numStyleLink
    const abs0 = createBulletAbstractNum(0);
    abs0.setNumStyleLink('ListBullet');
    const abs1 = createBulletAbstractNum(1);
    abs1.setNumStyleLink('ListBullet');

    manager.addAbstractNumbering(abs0);
    manager.addAbstractNumbering(abs1);
    manager.addInstance(NumberingInstance.create({ numId: 1, abstractNumId: 0 }));
    manager.addInstance(NumberingInstance.create({ numId: 2, abstractNumId: 1 }));
    manager.resetModified();

    const result = manager.consolidateNumbering();

    // Same numStyleLink and same levels → consolidate
    expect(result.abstractNumsRemoved).toBe(1);
    expect(result.groupsConsolidated).toBe(1);
    expect(manager.getAbstractNumberingCount()).toBe(1);
  });
});
