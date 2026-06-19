/**
 * removeStyle() must persist for documents loaded from file. Loaded
 * documents save styles.xml by merging modified styles into the preserved
 * original XML — previously the merge only replaced or appended styles, so
 * removals never reached the saved file and reappeared on reload.
 */
import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';

async function buildDocWithCustomStyle(styleId: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addStyle(
    Style.create({
      styleId,
      name: `${styleId} Name`,
      type: 'paragraph',
    })
  );
  doc.createParagraph('Body text');
  try {
    return await doc.toBuffer();
  } finally {
    doc.dispose();
  }
}

describe('X2: removeStyle persists across save/reload for loaded documents', () => {
  it('removed style is absent after saving and reloading a loaded document', async () => {
    const buffer1 = await buildDocWithCustomStyle('ObsoleteStyle');

    const loaded = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      expect(loaded.hasStyle('ObsoleteStyle')).toBe(true);
      expect(loaded.removeStyle('ObsoleteStyle')).toBe(true);
      expect(loaded.hasStyle('ObsoleteStyle')).toBe(false);
      buffer2 = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const reloaded = await Document.loadFromBuffer(buffer2);
    try {
      expect(reloaded.hasStyle('ObsoleteStyle')).toBe(false);
      expect(reloaded.getStyle('ObsoleteStyle')).toBeUndefined();
    } finally {
      reloaded.dispose();
    }
  });

  it('untouched styles survive the merge when another style is removed', async () => {
    const doc = Document.create();
    doc.addStyle(Style.create({ styleId: 'KeepMe', name: 'Keep Me', type: 'paragraph' }));
    doc.addStyle(Style.create({ styleId: 'DropMe', name: 'Drop Me', type: 'paragraph' }));
    doc.createParagraph('Body text');
    let buffer1: Buffer;
    try {
      buffer1 = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const loaded = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      expect(loaded.removeStyle('DropMe')).toBe(true);
      buffer2 = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const reloaded = await Document.loadFromBuffer(buffer2);
    try {
      expect(reloaded.hasStyle('KeepMe')).toBe(true);
      expect(reloaded.hasStyle('DropMe')).toBe(false);
      expect(reloaded.hasStyle('Normal')).toBe(true);
    } finally {
      reloaded.dispose();
    }
  });
});
