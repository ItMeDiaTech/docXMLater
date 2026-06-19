/**
 * setDefaultFont()/setDefaultFontSize() must persist for documents loaded
 * from file. Loaded documents save styles.xml by merging modified styles
 * into the preserved original XML, and only styles registered as modified
 * are merged — previously the setters mutated the Normal style in place
 * without registering it, so the change was silently dropped on save.
 */
import { Document } from '../../src/core/Document';

const JSZip = require('jszip');

async function buildDocWithCalibriNormal(): Promise<Buffer> {
  const doc = Document.create();
  doc.setDefaultFont('Calibri', 11);
  doc.createParagraph('Body text');
  try {
    return await doc.toBuffer();
  } finally {
    doc.dispose();
  }
}

describe('C27: default font changes persist on loaded documents', () => {
  it('writes setDefaultFont changes into styles.xml when saving a loaded document', async () => {
    const buffer1 = await buildDocWithCalibriNormal();

    const loaded = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      loaded.setDefaultFont('Times New Roman', 14);
      buffer2 = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const zip = await JSZip.loadAsync(buffer2);
    const stylesXml = await zip.file('word/styles.xml')!.async('string');
    // The font must land inside the Normal style block specifically, not merely
    // appear somewhere in styles.xml (a theme override or unrelated style).
    const normalBlock =
      stylesXml.match(/<w:style[^>]*w:styleId="Normal"[\s\S]*?<\/w:style>/)?.[0] ?? '';
    expect(normalBlock).toContain('Times New Roman');
    expect(normalBlock).toMatch(/<w:sz w:val="28"\/>/); // 14pt = 28 half-points

    const reloaded = await Document.loadFromBuffer(buffer2);
    try {
      const fmt = reloaded.getStylesManager().getStyle('Normal')!.getRunFormatting()!;
      expect(fmt.font).toBe('Times New Roman');
      expect(fmt.size).toBe(14);
    } finally {
      reloaded.dispose();
    }
  });

  it('writes setDefaultFontSize changes into styles.xml when saving a loaded document', async () => {
    const buffer1 = await buildDocWithCalibriNormal();

    const loaded = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      loaded.setDefaultFontSize(18);
      buffer2 = await loaded.toBuffer();
    } finally {
      loaded.dispose();
    }

    const reloaded = await Document.loadFromBuffer(buffer2);
    try {
      const fmt = reloaded.getStylesManager().getStyle('Normal')!.getRunFormatting()!;
      expect(fmt.size).toBe(18);
      expect(fmt.font).toBe('Calibri');
    } finally {
      reloaded.dispose();
    }
  });
});
