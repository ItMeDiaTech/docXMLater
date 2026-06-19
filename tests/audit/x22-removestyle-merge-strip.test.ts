/**
 * StylesManager must track removals so the styles merge can strip them from
 * the preserved original XML — mirroring NumberingManager's removed-ID
 * tracking. Previously removeStyle() only deleted from the in-memory map
 * without setting the modified flag or recording the ID, so the merge wrote
 * the original styles.xml back verbatim.
 */
import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';

const JSZip = require('jszip');

async function stylesXmlFromBuffer(buffer: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file('word/styles.xml')!.async('string');
}

describe('X22: styles merge strips removed styles from original XML', () => {
  it('records removal state in StylesManager', async () => {
    const doc = Document.create();
    doc.addStyle(Style.create({ styleId: 'Tracked', name: 'Tracked', type: 'paragraph' }));
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer);
    try {
      const manager = loaded.getStylesManager();
      expect(manager.getRemovedStyleIds().size).toBe(0);

      expect(loaded.removeStyle('Tracked')).toBe(true);
      expect(manager.isModified()).toBe(true);
      expect(manager.getRemovedStyleIds().has('Tracked')).toBe(true);
    } finally {
      loaded.dispose();
    }
  });

  it('applies removals alongside modified-style merging in saved styles.xml', async () => {
    const doc = Document.create();
    doc.addStyle(Style.create({ styleId: 'Survivor', name: 'Survivor', type: 'paragraph' }));
    doc.addStyle(Style.create({ styleId: 'Victim', name: 'Victim', type: 'paragraph' }));
    const buffer1 = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      // Modify one style and remove another in the same session
      loaded.addStyle(
        Style.create({
          styleId: 'Survivor',
          name: 'Survivor Updated',
          type: 'paragraph',
        })
      );
      expect(loaded.removeStyle('Victim')).toBe(true);
      buffer2 = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const stylesXml = await stylesXmlFromBuffer(buffer2);
    expect(stylesXml).toContain('w:styleId="Survivor"');
    expect(stylesXml).toContain('Survivor Updated');
    expect(stylesXml).not.toContain('w:styleId="Victim"');
  });

  it('does not resurrect a style that was added then removed in the same session', async () => {
    const doc = Document.create();
    const buffer1 = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      loaded.addStyle(Style.create({ styleId: 'Transient', name: 'Transient', type: 'paragraph' }));
      expect(loaded.removeStyle('Transient')).toBe(true);
      buffer2 = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const stylesXml = await stylesXmlFromBuffer(buffer2);
    expect(stylesXml).not.toContain('w:styleId="Transient"');
  });

  it('re-adding a removed style cancels the pending removal', async () => {
    const doc = Document.create();
    doc.addStyle(Style.create({ styleId: 'Phoenix', name: 'Phoenix', type: 'paragraph' }));
    const buffer1 = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      expect(loaded.removeStyle('Phoenix')).toBe(true);
      loaded.addStyle(
        Style.create({ styleId: 'Phoenix', name: 'Phoenix Reborn', type: 'paragraph' })
      );
      buffer2 = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const stylesXml = await stylesXmlFromBuffer(buffer2);
    expect(stylesXml).toContain('w:styleId="Phoenix"');
    expect(stylesXml).toContain('Phoenix Reborn');
  });
});
