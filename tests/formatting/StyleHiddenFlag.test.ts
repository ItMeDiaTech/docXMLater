/**
 * Style metadata — `<w:hidden>` CT_OnOff flag round-trip.
 *
 * Per ECMA-376 Part 1 §17.7.4 CT_Style, the `<w:hidden>` element (between
 * `autoRedefine` and `uiPriority`) marks the style as completely hidden
 * (stronger than `semiHidden`). Previously the library had no model field
 * for this — it was neither parsed nor emitted, so a document importing
 * a hidden style lost the hidden status on programmatic resave.
 *
 * Like other CT_Style metadata flags, `<w:hidden>` is OnOffOnlyType:
 * explicit-false is `w:val="off"`, not `"0"` / `"false"`.
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleHidden(hiddenInner: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`
  );
  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="HiddenTest">
    <w:name w:val="HiddenTest"/>${hiddenInner}
  </w:style>
</w:styles>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('Style <w:hidden> flag (§17.7.4, OnOffOnlyType)', () => {
  it('parses <w:hidden/> (bare) as hidden: true', async () => {
    const buffer = await makeDocxWithStyleHidden('<w:hidden/>');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('HiddenTest');
    expect(style?.getProperties().hidden).toBe(true);
    doc.dispose();
  });

  it('parses <w:hidden w:val="on"/> as hidden: true', async () => {
    const buffer = await makeDocxWithStyleHidden('<w:hidden w:val="on"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getStylesManager().getStyle('HiddenTest')?.getProperties().hidden).toBe(true);
    doc.dispose();
  });

  it('parses <w:hidden w:val="off"/> as hidden: false', async () => {
    const buffer = await makeDocxWithStyleHidden('<w:hidden w:val="off"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getStylesManager().getStyle('HiddenTest')?.getProperties().hidden).toBe(false);
    doc.dispose();
  });

  it('leaves hidden undefined when absent', async () => {
    const buffer = await makeDocxWithStyleHidden('');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getStylesManager().getStyle('HiddenTest')?.getProperties().hidden).toBeUndefined();
    doc.dispose();
  });

  it('emits bare <w:hidden/> for explicit true via toXML()', () => {
    const style = new Style({
      styleId: 'HiddenTest',
      type: 'paragraph',
      name: 'HiddenTest',
      hidden: true,
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toContain('<w:hidden/>');
  });

  it('emits <w:hidden w:val="off"/> for explicit false (OnOffOnlyType)', () => {
    const style = new Style({
      styleId: 'HiddenTest',
      type: 'paragraph',
      name: 'HiddenTest',
      hidden: false,
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toContain('<w:hidden w:val="off"/>');
  });

  it('omits <w:hidden> when the field is undefined', () => {
    const style = new Style({
      styleId: 'HiddenTest',
      type: 'paragraph',
      name: 'HiddenTest',
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).not.toContain('<w:hidden');
  });

  it('emits <w:hidden> between autoRedefine and uiPriority per CT_Style order', () => {
    const style = new Style({
      styleId: 'HiddenTest',
      type: 'paragraph',
      name: 'HiddenTest',
      autoRedefine: true,
      hidden: true,
      uiPriority: 99,
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    const autoIdx = xml.indexOf('<w:autoRedefine');
    const hiddenIdx = xml.indexOf('<w:hidden');
    const uiIdx = xml.indexOf('<w:uiPriority');
    expect(autoIdx).toBeGreaterThan(-1);
    expect(hiddenIdx).toBeGreaterThan(autoIdx);
    expect(uiIdx).toBeGreaterThan(hiddenIdx);
  });
});
