/**
 * Document.removeStyle() returning true must mean the style is gone from the
 * saved word/styles.xml. For loaded documents the save path merges into the
 * preserved original XML, which previously had no removal handling — the
 * "removed" style was written back verbatim while the API reported success.
 */
import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';

const JSZip = require('jszip');

async function stylesXmlFromBuffer(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file('word/styles.xml')!.async('string');
}

describe('X39: removeStyle removes the style from saved styles.xml', () => {
  it('strips the removed style block from the saved XML of a loaded document', async () => {
    const doc = Document.create();
    doc.addStyle(
      Style.create({
        styleId: 'DoomedStyle',
        name: 'Doomed Style',
        type: 'paragraph',
        runFormatting: { bold: true, size: 14 },
      })
    );
    doc.createParagraph('Body text').setStyle('Normal');
    const buffer1 = await doc.toBuffer();
    doc.dispose();

    // Sanity: the style survives the first round trip
    expect(await stylesXmlFromBuffer(buffer1)).toContain('w:styleId="DoomedStyle"');

    const loaded = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      expect(loaded.removeStyle('DoomedStyle')).toBe(true);
      buffer2 = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const stylesXml = await stylesXmlFromBuffer(buffer2);
    expect(stylesXml).not.toContain('w:styleId="DoomedStyle"');
    expect(stylesXml).not.toContain('Doomed Style');
    // Other styles remain intact
    expect(stylesXml).toContain('w:styleId="Normal"');
  });

  it('keeps the removal applied when the same document is saved twice', async () => {
    const doc = Document.create();
    doc.addStyle(Style.create({ styleId: 'DoomedStyle', name: 'Doomed Style', type: 'paragraph' }));
    const buffer1 = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer1);
    try {
      expect(loaded.removeStyle('DoomedStyle')).toBe(true);

      const firstSave = await loaded.toBuffer();
      expect(await stylesXmlFromBuffer(firstSave)).not.toContain('w:styleId="DoomedStyle"');

      // Each save merges against the preserved original XML, so the removal
      // must still apply on subsequent saves
      const secondSave = await loaded.toBuffer();
      expect(await stylesXmlFromBuffer(secondSave)).not.toContain('w:styleId="DoomedStyle"');
    } finally {
      loaded.dispose();
    }
  });
});
