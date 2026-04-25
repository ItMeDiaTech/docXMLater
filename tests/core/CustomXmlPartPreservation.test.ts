/**
 * Regression: custom XML parts (`customXml/item*.xml`,
 * `customXml/itemProps*.xml`) plus their per-part relationship and
 * Content_Types entries must survive load → save round-trip even when
 * the document body is modified between load and save.
 *
 * Word stores SDT data-binding payloads here; losing them on round-trip
 * silently breaks Word/Power Automate field bindings without producing
 * a parse error.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const ITEM1_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<myCustom><answer>42</answer></myCustom>`;
const ITEM_PROPS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{12345678-ABCD-EF01-2345-6789ABCDEF01}"/>`;
const ITEM1_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>
</Relationships>`;

async function buildDocxWithCustomXml(): Promise<Buffer> {
  // Build a real, valid base document via Document.create(), then inject
  // custom XML parts into the saved buffer's ZIP layer. This avoids
  // hand-crafting all required OOXML parts (styles.xml, fontTable, etc.)
  // and lets the test focus on what it cares about: customXml/* survival.
  const seed = Document.create();
  seed.createParagraph('Original');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);
  zip.addFile('customXml/item1.xml', ITEM1_XML);
  zip.addFile('customXml/itemProps1.xml', ITEM_PROPS_XML);
  zip.addFile('customXml/_rels/item1.xml.rels', ITEM1_RELS_XML);

  // Add Content_Types overrides for the new parts.
  const ct = zip.getFileAsString('[Content_Types].xml')!;
  const updated = ct.replace(
    '</Types>',
    `<Override PartName="/customXml/item1.xml" ContentType="application/xml"/>` +
      `<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>` +
      `</Types>`
  );
  zip.updateFile('[Content_Types].xml', updated);

  // Wire the document → customXml relationship.
  const docRels = zip.getFileAsString('word/_rels/document.xml.rels')!;
  const updatedRels = docRels.replace(
    '</Relationships>',
    `<Relationship Id="rId900" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>` +
      `</Relationships>`
  );
  zip.updateFile('word/_rels/document.xml.rels', updatedRels);

  return zip.toBuffer();
}

describe('Custom XML part preservation', () => {
  it('preserves customXml/* files on unmodified round-trip', async () => {
    const buffer1 = await buildDocxWithCustomXml();
    const doc = await Document.loadFromBuffer(buffer1);
    const buffer2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buffer2);
    expect(out.hasFile('customXml/item1.xml')).toBe(true);
    expect(out.hasFile('customXml/itemProps1.xml')).toBe(true);
    expect(out.hasFile('customXml/_rels/item1.xml.rels')).toBe(true);
    expect(out.getFileAsString('customXml/item1.xml')).toContain('<answer>42</answer>');
    expect(out.getFileAsString('customXml/itemProps1.xml')).toContain(
      '{12345678-ABCD-EF01-2345-6789ABCDEF01}'
    );
  });

  it('preserves customXml/* files when the body is modified between load and save', async () => {
    const buffer1 = await buildDocxWithCustomXml();
    const doc = await Document.loadFromBuffer(buffer1);
    doc.createParagraph('Added after load');
    const buffer2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buffer2);
    expect(out.hasFile('customXml/item1.xml')).toBe(true);
    expect(out.hasFile('customXml/itemProps1.xml')).toBe(true);
    expect(out.getFileAsString('customXml/item1.xml')).toContain('<answer>42</answer>');
  });

  it('preserves Content_Types overrides for customXml parts', async () => {
    const buffer1 = await buildDocxWithCustomXml();
    const doc = await Document.loadFromBuffer(buffer1);
    const buffer2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buffer2);
    const ct = out.getFileAsString('[Content_Types].xml')!;
    expect(ct).toContain('/customXml/item1.xml');
    expect(ct).toContain('/customXml/itemProps1.xml');
  });

  it('preserves the document → customXml relationship', async () => {
    const buffer1 = await buildDocxWithCustomXml();
    const doc = await Document.loadFromBuffer(buffer1);
    const buffer2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buffer2);
    const rels = out.getFileAsString('word/_rels/document.xml.rels')!;
    expect(rels).toContain('customXml/item1.xml');
    expect(rels).toMatch(/Type="[^"]*relationships\/customXml"/);
  });

  it('survives a double round-trip', async () => {
    const buffer1 = await buildDocxWithCustomXml();
    const doc1 = await Document.loadFromBuffer(buffer1);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    const buffer3 = await doc2.toBuffer();
    doc2.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buffer3);
    expect(out.hasFile('customXml/item1.xml')).toBe(true);
    expect(out.hasFile('customXml/itemProps1.xml')).toBe(true);
  });
});
