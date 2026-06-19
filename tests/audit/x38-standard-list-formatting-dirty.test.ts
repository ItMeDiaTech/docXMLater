/**
 * Tests that applyStandardListFormatting/applyStandardNumberedListFormatting
 * mark the mutated abstract numberings as modified, so the level changes made
 * to a LOADED document survive the save pipeline instead of being replaced by
 * the preserved original numbering.xml.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: build a DOCX buffer containing one list paragraph (bullet or
 * numbered) and return the buffer plus the numId of the list.
 */
async function createListDocBuffer(
  kind: 'bullet' | 'numbered'
): Promise<{ buffer: Buffer; numId: number }> {
  const doc = Document.create();
  try {
    const manager = doc.getNumberingManager();
    const numId = kind === 'bullet' ? manager.createBulletList() : manager.createNumberedList();

    const para = new Paragraph().addText('Item 1');
    para.setNumbering(numId, 0);
    doc.addParagraph(para);

    const buffer = await doc.toBuffer();
    return { buffer, numId };
  } finally {
    doc.dispose();
  }
}

async function readNumberingXml(buffer: Buffer): Promise<string> {
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  const xml = zip.getFileAsString(DOCX_PATHS.NUMBERING);
  expect(xml).toBeTruthy();
  return xml!;
}

describe('Standard list formatting dirty tracking', () => {
  it('applyStandardListFormatting marks the bullet abstract numbering modified', async () => {
    const { buffer, numId } = await createListDocBuffer('bullet');
    const doc = await Document.loadFromBuffer(buffer);
    try {
      const manager = doc.getNumberingManager();
      expect(manager.isModified()).toBe(false);

      const count = doc.applyStandardListFormatting();
      expect(count).toBeGreaterThanOrEqual(1);

      expect(manager.isModified()).toBe(true);
      const instance = manager.getInstance(numId)!;
      expect(manager.getModifiedAbstractNumIds().has(instance.getAbstractNumId())).toBe(true);
    } finally {
      doc.dispose();
    }
  });

  it('persists applyStandardListFormatting level changes when saving a loaded document', async () => {
    const { buffer, numId } = await createListDocBuffer('bullet');

    const doc = await Document.loadFromBuffer(buffer);
    let output: Buffer;
    try {
      const count = doc.applyStandardListFormatting();
      expect(count).toBeGreaterThanOrEqual(1);
      output = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    // Standard formatting sets the bullet font to Arial and level 1 indent to
    // 1440 twips; neither value exists in the original numbering.xml
    const numberingXml = await readNumberingXml(output);
    expect(numberingXml).toContain('Arial');
    expect(numberingXml).toContain('w:left="1440"');

    // Reload and confirm the parsed model reflects the new level formatting
    const reloaded = await Document.loadFromBuffer(output);
    try {
      const manager = reloaded.getNumberingManager();
      const instance = manager.getInstance(numId)!;
      const level0 = manager.getAbstractNumbering(instance.getAbstractNumId())!.getLevel(0)!;
      expect(level0.getProperties().font).toBe('Arial');
      expect(level0.getProperties().text).toBe('•');
      expect(level0.getProperties().leftIndent).toBe(720);
      expect(level0.getProperties().hangingIndent).toBe(360);
    } finally {
      reloaded.dispose();
    }
  });

  it('persists applyStandardNumberedListFormatting level changes when saving a loaded document', async () => {
    const { buffer, numId } = await createListDocBuffer('numbered');

    const doc = await Document.loadFromBuffer(buffer);
    let output: Buffer;
    try {
      const count = doc.applyStandardNumberedListFormatting();
      expect(count).toBeGreaterThanOrEqual(1);
      output = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    // Standard formatting uses 720 * (level + 1) indents; level 8 becomes
    // 6480 twips, a value that cannot appear in the default 720 + level * 360
    // scheme of the original numbering.xml
    const numberingXml = await readNumberingXml(output);
    expect(numberingXml).toContain('w:left="6480"');

    const reloaded = await Document.loadFromBuffer(output);
    try {
      const manager = reloaded.getNumberingManager();
      const instance = manager.getInstance(numId)!;
      const abstractNum = manager.getAbstractNumbering(instance.getAbstractNumId())!;
      // Defaults are 1080 (level 1) and 3600 (level 8); standard formatting
      // replaces them with 1440 and 6480
      expect(abstractNum.getLevel(1)!.getProperties().leftIndent).toBe(1440);
      expect(abstractNum.getLevel(8)!.getProperties().leftIndent).toBe(6480);
      expect(abstractNum.getLevel(0)!.getProperties().font).toBe('Verdana');
      expect(abstractNum.getLevel(0)!.getProperties().hangingIndent).toBe(360);
    } finally {
      reloaded.dispose();
    }
  });
});
