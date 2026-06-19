/**
 * Regression: content-type generation must recognize custom XML parts under
 * the OPC-conventional 'customXml/' (lowercase 'ml') prefix that Word writes.
 * The scan previously matched only 'customXML/', so itemProps parts added
 * programmatically via setPart() to a new document got no
 * customXmlProperties+xml override and fell back to the generic xml Default.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const ITEM1_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<myCustom><answer>42</answer></myCustom>`;
const ITEM_PROPS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{12345678-ABCD-EF01-2345-6789ABCDEF01}"/>`;

describe('Content types for programmatically added customXml parts', () => {
  it('emits the customXmlProperties+xml override for customXml/itemProps1.xml added via setPart()', async () => {
    const doc = Document.create();
    try {
      doc.createParagraph('Body');
      await doc.setPart('customXml/item1.xml', ITEM1_XML);
      await doc.setPart('customXml/itemProps1.xml', ITEM_PROPS_XML);
      const buffer = await doc.toBuffer();

      const out = new ZipHandler();
      await out.loadFromBuffer(buffer);
      const ct = out.getFileAsString('[Content_Types].xml')!;

      expect(ct).toContain(
        '<Override PartName="/customXml/itemProps1.xml" ' +
          'ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>'
      );
      expect(ct).toContain('<Override PartName="/customXml/item1.xml"');
    } finally {
      doc.dispose();
    }
  });
});
