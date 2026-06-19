/**
 * w14:checkedState / w14:uncheckedState (CT_SdtCheckboxSymbol) pair a
 * character code with a specific glyph font. The parser discarded
 * w14:font and toXML() unconditionally re-emitted 'MS Gothic', so a
 * checkbox authored with Wingdings codes (e.g. F0FE) was re-pointed at a
 * font where those code points map to wrong or missing glyphs — visible
 * corruption the next time the checkbox is toggled. The font now
 * round-trips; MS Gothic remains the default only for programmatically
 * created checkboxes.
 */
import { Document } from '../../src/core/Document';
import { StructuredDocumentTag } from '../../src/elements/StructuredDocumentTag';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLElement } from '../../src/xml/XMLBuilder';

function findChild(parent: XMLElement, name: string): XMLElement | undefined {
  return parent.children?.find(
    (child): child is XMLElement => typeof child !== 'string' && child.name === name
  );
}

const WINGDINGS_SDT =
  `<w:sdt>` +
  `<w:sdtPr>` +
  `<w:id w:val="77"/>` +
  `<w14:checkbox>` +
  `<w14:checked w14:val="1"/>` +
  `<w14:checkedState w14:val="F0FE" w14:font="Wingdings"/>` +
  `<w14:uncheckedState w14:val="F0A8" w14:font="Wingdings 2"/>` +
  `</w14:checkbox>` +
  `</w:sdtPr>` +
  `<w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent>` +
  `</w:sdt>`;

async function buildDocxWithCheckbox(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('checkbox font round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);
  const docXml = zip.getFileAsString('word/document.xml')!;
  const updated = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${WINGDINGS_SDT}<w:sectPr`)
    : docXml.replace('</w:body>', `${WINGDINGS_SDT}</w:body>`);
  zip.updateFile('word/document.xml', updated);
  return zip.toBuffer();
}

describe('checkbox SDT state-symbol w14:font round-trip', () => {
  it('parses w14:font into checkedFont/uncheckedFont', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithCheckbox());
    try {
      const sdt = doc
        .getBodyElements()
        .find((el): el is StructuredDocumentTag => el instanceof StructuredDocumentTag);
      expect(sdt).toBeDefined();
      const props = sdt!.getCheckboxProperties();
      expect(props?.checkedFont).toBe('Wingdings');
      expect(props?.uncheckedFont).toBe('Wingdings 2');
    } finally {
      doc.dispose();
    }
  });

  it('re-emits the parsed font instead of MS Gothic when sdtPr is rebuilt', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithCheckbox());
    try {
      const sdt = doc
        .getBodyElements()
        .find((el): el is StructuredDocumentTag => el instanceof StructuredDocumentTag);
      expect(sdt).toBeDefined();
      // Toggling through the API invalidates the raw sdtPr passthrough,
      // forcing toXML() to rebuild the checkbox from the model.
      sdt!.setCheckboxProperties({ ...sdt!.getCheckboxProperties()!, checked: false });

      const out = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const docXml = zip.getFileAsString('word/document.xml')!;
      expect(docXml).toContain('w14:font="Wingdings"');
      expect(docXml).toContain('w14:font="Wingdings 2"');
      expect(docXml).not.toContain('MS Gothic');
    } finally {
      doc.dispose();
    }
  });

  it('still defaults to MS Gothic for programmatically created checkboxes', () => {
    const sdt = StructuredDocumentTag.createCheckbox(true, [new Paragraph().addText('☒')]);

    const sdtPr = findChild(sdt.toXML(), 'w:sdtPr');
    expect(sdtPr).toBeDefined();
    const checkbox = findChild(sdtPr!, 'w14:checkbox');
    expect(checkbox).toBeDefined();
    const checkedState = findChild(checkbox!, 'w14:checkedState');
    const uncheckedState = findChild(checkbox!, 'w14:uncheckedState');
    expect(checkedState!.attributes!['w14:font']).toBe('MS Gothic');
    expect(uncheckedState!.attributes!['w14:font']).toBe('MS Gothic');
  });
});
