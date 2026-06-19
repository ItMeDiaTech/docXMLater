/**
 * CT_SdtComboBox / CT_SdtDropDownList (ECMA-376 §17.5.2.5 / §17.5.2.15)
 * record the control's last selected list entry in the optional
 * w:lastValue attribute. The parser captured it into ComboBoxProperties /
 * DropDownListProperties, but toXML() emitted <w:comboBox> /
 * <w:dropDownList> with an empty attribute set, so the stored selection
 * was silently dropped whenever sdtPr was rebuilt — programmatic
 * creation, or any modeled-property mutation on a loaded SDT.
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

const COMBO_SDT =
  `<w:sdt>` +
  `<w:sdtPr>` +
  `<w:id w:val="42"/>` +
  `<w:tag w:val="orig-tag"/>` +
  `<w:comboBox w:lastValue="beta">` +
  `<w:listItem w:displayText="Alpha" w:value="alpha"/>` +
  `<w:listItem w:displayText="Beta" w:value="beta"/>` +
  `</w:comboBox>` +
  `</w:sdtPr>` +
  `<w:sdtContent><w:p><w:r><w:t>Beta</w:t></w:r></w:p></w:sdtContent>` +
  `</w:sdt>`;

async function buildDocxWithComboSdt(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('combo lastValue round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);
  const docXml = zip.getFileAsString('word/document.xml')!;
  const updated = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${COMBO_SDT}<w:sectPr`)
    : docXml.replace('</w:body>', `${COMBO_SDT}</w:body>`);
  zip.updateFile('word/document.xml', updated);
  return zip.toBuffer();
}

describe('comboBox/dropDownList w:lastValue emission', () => {
  it('emits w:lastValue on a programmatically configured combo box', () => {
    const sdt = StructuredDocumentTag.createComboBox(
      [{ displayText: 'Alpha', value: 'alpha' }],
      [new Paragraph().addText('Alpha')]
    );
    sdt.setComboBoxProperties({
      items: [{ displayText: 'Alpha', value: 'alpha' }],
      lastValue: 'alpha',
    });

    const sdtPr = findChild(sdt.toXML(), 'w:sdtPr');
    expect(sdtPr).toBeDefined();
    const comboBox = findChild(sdtPr!, 'w:comboBox');
    expect(comboBox).toBeDefined();
    expect(comboBox!.attributes!['w:lastValue']).toBe('alpha');
  });

  it('emits w:lastValue on a programmatically configured dropdown list', () => {
    const sdt = StructuredDocumentTag.createDropDownList(
      [{ displayText: 'One', value: '1' }],
      [new Paragraph().addText('One')]
    );
    sdt.setDropDownListProperties({
      items: [{ displayText: 'One', value: '1' }],
      lastValue: '1',
    });

    const sdtPr = findChild(sdt.toXML(), 'w:sdtPr');
    expect(sdtPr).toBeDefined();
    const dropDown = findChild(sdtPr!, 'w:dropDownList');
    expect(dropDown).toBeDefined();
    expect(dropDown!.attributes!['w:lastValue']).toBe('1');
  });

  it('omits w:lastValue when no selection has been recorded', () => {
    const sdt = StructuredDocumentTag.createComboBox(
      [{ displayText: 'Alpha', value: 'alpha' }],
      [new Paragraph().addText('Alpha')]
    );

    const sdtPr = findChild(sdt.toXML(), 'w:sdtPr');
    const comboBox = findChild(sdtPr!, 'w:comboBox');
    expect(comboBox!.attributes?.['w:lastValue']).toBeUndefined();
  });

  it('preserves a parsed w:lastValue when sdtPr is rebuilt after a mutation', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithComboSdt());
    try {
      const sdt = doc
        .getBodyElements()
        .find((el): el is StructuredDocumentTag => el instanceof StructuredDocumentTag);
      expect(sdt).toBeDefined();
      expect(sdt!.getComboBoxProperties()?.lastValue).toBe('beta');

      // Invalidate the raw sdtPr passthrough so toXML() rebuilds from the model
      sdt!.setTag('mutated-tag');

      const out = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const docXml = zip.getFileAsString('word/document.xml')!;
      expect(docXml).toContain('w:lastValue="beta"');
    } finally {
      doc.dispose();
    }
  });
});
