/**
 * Tests that consolidateNumbering adds startOverrides when remapping
 * instances onto a shared canonical abstractNum. Numbering counters belong
 * to the abstract definition (ECMA-376 §17.9.27 startOverride), so without
 * an override two previously independent numbered lists would merge into a
 * single continuing sequence (1,2,3 then 4,5,6) after consolidation.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: build a document with two identical decimal lists,
 * each with one paragraph, and return it (caller disposes).
 */
function createDocWithTwoNumberedLists(): { doc: Document; numId1: number; numId2: number } {
  const doc = Document.create();
  const manager = doc.getNumberingManager();

  const numId1 = manager.createNumberedList();
  const numId2 = manager.createNumberedList();

  const p1 = new Paragraph().addText('List one item');
  p1.setNumbering(numId1, 0);
  doc.addParagraph(p1);

  const p2 = new Paragraph().addText('List two item');
  p2.setNumbering(numId2, 0);
  doc.addParagraph(p2);

  return { doc, numId1, numId2 };
}

async function readNumberingXml(buffer: Buffer): Promise<string> {
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  const xml = zip.getFileAsString(DOCX_PATHS.NUMBERING);
  expect(xml).toBeTruthy();
  return xml!;
}

/** Extracts the <w:num> element for the given numId from numbering.xml */
function extractNumElement(numberingXml: string, numId: number): string {
  const match = new RegExp(`<w:num [^>]*w:numId="${numId}"[^>]*>[\\s\\S]*?</w:num>`).exec(
    numberingXml
  );
  expect(match).toBeTruthy();
  return match![0];
}

describe('consolidateNumbering startOverride for remapped instances', () => {
  it('saves a startOverride on the remapped num so list 2 still begins at 1', async () => {
    const { doc, numId1, numId2 } = createDocWithTwoNumberedLists();
    let output: Buffer;
    try {
      const result = doc.consolidateNumbering();
      expect(result.abstractNumsRemoved).toBe(1);
      expect(result.instancesRemapped).toBe(1);

      output = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const numberingXml = await readNumberingXml(output);

    // The remapped instance restarts independently at the level start value
    const num2Xml = extractNumElement(numberingXml, numId2);
    expect(num2Xml).toContain('<w:lvlOverride w:ilvl="0">');
    expect(num2Xml).toContain('<w:startOverride w:val="1"/>');

    // The canonical instance keeps its original counter — no override
    const num1Xml = extractNumElement(numberingXml, numId1);
    expect(num1Xml).not.toContain('startOverride');
  });

  it('round-trips the restart through the merge path of a loaded document', async () => {
    const { doc, numId2 } = createDocWithTwoNumberedLists();
    let buffer: Buffer;
    try {
      buffer = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    // Consolidate AFTER reload so the save goes through original-XML merging
    const loaded = await Document.loadFromBuffer(buffer);
    let output: Buffer;
    try {
      const result = loaded.consolidateNumbering();
      expect(result.instancesRemapped).toBe(1);
      output = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const numberingXml = await readNumberingXml(output);
    const num2Xml = extractNumElement(numberingXml, numId2);
    expect(num2Xml).toContain('<w:startOverride w:val="1"/>');

    // Reload once more and confirm the parsed model retains the override
    const reloaded = await Document.loadFromBuffer(output);
    try {
      const instance = reloaded.getNumberingManager().getInstance(numId2)!;
      expect(instance.getLevelOverride(0)).toBe(1);
    } finally {
      reloaded.dispose();
    }
  });

  it('does not add startOverrides when consolidating bullet lists', async () => {
    const doc = Document.create();
    let output: Buffer;
    let bulletNumId2: number;
    try {
      const manager = doc.getNumberingManager();
      const bulletNumId1 = manager.createBulletList();
      bulletNumId2 = manager.createBulletList();

      const p1 = new Paragraph().addText('Bullet one');
      p1.setNumbering(bulletNumId1, 0);
      doc.addParagraph(p1);

      const p2 = new Paragraph().addText('Bullet two');
      p2.setNumbering(bulletNumId2, 0);
      doc.addParagraph(p2);

      const result = doc.consolidateNumbering();
      expect(result.instancesRemapped).toBe(1);

      output = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const numberingXml = await readNumberingXml(output);
    const num2Xml = extractNumElement(numberingXml, bulletNumId2);
    expect(num2Xml).not.toContain('startOverride');
  });
});
