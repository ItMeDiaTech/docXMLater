/**
 * Tests that setListIndentation/normalizeListIndentation mark the affected
 * abstract numbering as modified, so indentation changes made to a LOADED
 * document survive the save pipeline instead of being replaced by the
 * preserved original numbering.xml.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: build a DOCX buffer containing one numbered list paragraph
 * and return the buffer plus the numId of the list.
 */
async function createNumberedDocBuffer(
  leftIndent?: number,
  hangingIndent?: number
): Promise<{ buffer: Buffer; numId: number }> {
  const doc = Document.create();
  try {
    const manager = doc.getNumberingManager();
    const numId = manager.createNumberedList();

    if (leftIndent !== undefined) {
      manager.setListIndentation(numId, 0, leftIndent, hangingIndent);
    }

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

describe('NumberingManager dirty tracking for indentation APIs', () => {
  it('setListIndentation marks the abstract numbering modified', async () => {
    const { buffer, numId } = await createNumberedDocBuffer();
    const doc = await Document.loadFromBuffer(buffer);
    try {
      const manager = doc.getNumberingManager();
      expect(manager.isModified()).toBe(false);

      const ok = manager.setListIndentation(numId, 0, 1440, 720);
      expect(ok).toBe(true);

      expect(manager.isModified()).toBe(true);
      const instance = manager.getInstance(numId)!;
      expect(manager.getModifiedAbstractNumIds().has(instance.getAbstractNumId())).toBe(true);
    } finally {
      doc.dispose();
    }
  });

  it('persists setListIndentation changes when saving a loaded document', async () => {
    const { buffer, numId } = await createNumberedDocBuffer();

    const doc = await Document.loadFromBuffer(buffer);
    let output: Buffer;
    try {
      doc.setListIndentation(numId, 0, 1440, 720);
      output = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const numberingXml = await readNumberingXml(output);
    expect(numberingXml).toContain('w:left="1440"');
    expect(numberingXml).toContain('w:hanging="720"');

    // Reload and confirm the parsed model reflects the new indentation
    const reloaded = await Document.loadFromBuffer(output);
    try {
      const manager = reloaded.getNumberingManager();
      const instance = manager.getInstance(numId)!;
      const level = manager.getAbstractNumbering(instance.getAbstractNumId())!.getLevel(0)!;
      expect(level.getProperties().leftIndent).toBe(1440);
      expect(level.getProperties().hangingIndent).toBe(720);
    } finally {
      reloaded.dispose();
    }
  });

  it('persists normalizeAllListIndentation changes when saving a loaded document', async () => {
    // Start from a list with non-standard indentation
    const { buffer } = await createNumberedDocBuffer(2000, 500);

    const doc = await Document.loadFromBuffer(buffer);
    let output: Buffer;
    try {
      const count = doc.normalizeAllListIndentation();
      expect(count).toBeGreaterThanOrEqual(1);
      output = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const numberingXml = await readNumberingXml(output);
    // Level 0 standard indentation is 720/360; the non-standard values must be gone
    expect(numberingXml).toContain('w:left="720"');
    expect(numberingXml).not.toContain('w:left="2000"');
    expect(numberingXml).not.toContain('w:hanging="500"');
  });
});
