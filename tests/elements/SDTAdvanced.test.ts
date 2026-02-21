/**
 * Advanced SDT tests: placeholder, data binding, showingPlcHdr,
 * citation, bibliography, equation, docPartList types
 */

import { StructuredDocumentTag } from '../../src/elements/StructuredDocumentTag';
import { Paragraph } from '../../src/elements/Paragraph';
import { Document } from '../../src/core/Document';
import { XMLElement } from '../../src/xml/XMLBuilder';

function filterXMLElements(children?: (XMLElement | string)[]): XMLElement[] {
  return (children || []).filter((c): c is XMLElement => typeof c !== 'string');
}

function findInSdtPr(xml: XMLElement, name: string): XMLElement | undefined {
  const sdtPr = filterXMLElements(xml.children).find(c => c.name === 'w:sdtPr');
  return filterXMLElements(sdtPr?.children).find(c => c.name === name);
}

describe('SDT Placeholder', () => {
  test('should set and get placeholder', () => {
    const sdt = StructuredDocumentTag.createRichText();
    sdt.setPlaceholder('DefaultPlaceholder_-1854013440');
    expect(sdt.getPlaceholder()?.docPart).toBe('DefaultPlaceholder_-1854013440');
  });

  test('should generate w:placeholder in XML', () => {
    const sdt = StructuredDocumentTag.createRichText(
      [new Paragraph().addText('Content')],
    );
    sdt.setPlaceholder('DefaultPlaceholder_-1854013440');

    const xml = sdt.toXML();
    const placeholder = findInSdtPr(xml, 'w:placeholder');
    expect(placeholder).toBeDefined();

    const docPart = filterXMLElements(placeholder?.children).find(c => c.name === 'w:docPart');
    expect(docPart).toBeDefined();
    expect(docPart?.attributes?.['w:val']).toBe('DefaultPlaceholder_-1854013440');
  });

  test('should round-trip placeholder', async () => {
    const doc = Document.create();
    const sdt = StructuredDocumentTag.createPlainText(
      [new Paragraph().addText('Enter text here')],
      false,
    );
    sdt.setPlaceholder('DefaultPlaceholder_-1854013440');
    doc.addBodyElement(sdt);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedSdt = loaded.getBodyElements()
      .find(el => el instanceof StructuredDocumentTag) as StructuredDocumentTag;

    expect(loadedSdt).toBeDefined();
    expect(loadedSdt.getPlaceholder()?.docPart).toBe('DefaultPlaceholder_-1854013440');

    doc.dispose();
    loaded.dispose();
  });
});

describe('SDT Data Binding', () => {
  test('should set and get data binding', () => {
    const sdt = StructuredDocumentTag.createPlainText();
    sdt.setDataBinding(
      '/root/element',
      'xmlns:ns="http://example.com"',
      '{12345678-ABCD-1234-5678-ABCDEF012345}'
    );

    const db = sdt.getDataBinding();
    expect(db?.xpath).toBe('/root/element');
    expect(db?.prefixMappings).toBe('xmlns:ns="http://example.com"');
    expect(db?.storeItemId).toBe('{12345678-ABCD-1234-5678-ABCDEF012345}');
  });

  test('should generate w:dataBinding in XML', () => {
    const sdt = StructuredDocumentTag.createPlainText(
      [new Paragraph().addText('Bound value')],
    );
    sdt.setDataBinding(
      '/root/name',
      'xmlns:ns="http://example.com"',
      '{AABBCCDD-1234-5678-9ABC-DDEEFF001122}'
    );

    const xml = sdt.toXML();
    const dataBinding = findInSdtPr(xml, 'w:dataBinding');
    expect(dataBinding).toBeDefined();
    expect(dataBinding?.attributes?.['w:xpath']).toBe('/root/name');
    expect(dataBinding?.attributes?.['w:prefixMappings']).toBe('xmlns:ns="http://example.com"');
    expect(dataBinding?.attributes?.['w:storeItemID']).toBe('{AABBCCDD-1234-5678-9ABC-DDEEFF001122}');
  });

  test('should generate data binding without optional attributes', () => {
    const sdt = StructuredDocumentTag.createPlainText(
      [new Paragraph().addText('Simple binding')],
    );
    sdt.setDataBinding('/root/simple');

    const xml = sdt.toXML();
    const dataBinding = findInSdtPr(xml, 'w:dataBinding');
    expect(dataBinding).toBeDefined();
    expect(dataBinding?.attributes?.['w:xpath']).toBe('/root/simple');
    expect(dataBinding?.attributes?.['w:prefixMappings']).toBeUndefined();
    expect(dataBinding?.attributes?.['w:storeItemID']).toBeUndefined();
  });

  test('should round-trip data binding', async () => {
    const doc = Document.create();
    const sdt = StructuredDocumentTag.createPlainText(
      [new Paragraph().addText('Data')],
      false,
    );
    sdt.setDataBinding(
      '/root/value',
      undefined,
      '{11111111-2222-3333-4444-555555555555}'
    );
    doc.addBodyElement(sdt);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedSdt = loaded.getBodyElements()
      .find(el => el instanceof StructuredDocumentTag) as StructuredDocumentTag;

    expect(loadedSdt).toBeDefined();
    const db = loadedSdt.getDataBinding();
    expect(db?.xpath).toBe('/root/value');
    expect(db?.storeItemId).toBe('{11111111-2222-3333-4444-555555555555}');

    doc.dispose();
    loaded.dispose();
  });
});

describe('SDT Showing Placeholder', () => {
  test('should set and get showingPlcHdr', () => {
    const sdt = StructuredDocumentTag.createRichText();
    sdt.setShowingPlaceholder(true);
    expect(sdt.isShowingPlaceholder()).toBe(true);

    sdt.setShowingPlaceholder(false);
    expect(sdt.isShowingPlaceholder()).toBe(false);
  });

  test('should generate w:showingPlcHdr in XML', () => {
    const sdt = StructuredDocumentTag.createRichText(
      [new Paragraph().addText('Placeholder text')],
    );
    sdt.setShowingPlaceholder(true);

    const xml = sdt.toXML();
    const showingPlcHdr = findInSdtPr(xml, 'w:showingPlcHdr');
    expect(showingPlcHdr).toBeDefined();
    expect(showingPlcHdr?.attributes?.['w:val']).toBe('true');
  });

  test('should round-trip showingPlcHdr', async () => {
    const doc = Document.create();
    const sdt = StructuredDocumentTag.createRichText(
      [new Paragraph().addText('Placeholder')],
    );
    sdt.setPlaceholder('DefaultPlaceholder_-1854013440');
    sdt.setShowingPlaceholder(true);
    doc.addBodyElement(sdt);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedSdt = loaded.getBodyElements()
      .find(el => el instanceof StructuredDocumentTag) as StructuredDocumentTag;

    expect(loadedSdt).toBeDefined();
    expect(loadedSdt.isShowingPlaceholder()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('SDT Combined Placeholder + Data Binding', () => {
  test('should support placeholder + data binding + showingPlcHdr together', () => {
    const sdt = StructuredDocumentTag.createPlainText(
      [new Paragraph().addText('Click to enter text')],
      false,
    );
    sdt.setPlaceholder('DefaultPlaceholder_-1854013440');
    sdt.setDataBinding('/company/name');
    sdt.setShowingPlaceholder(true);

    const xml = sdt.toXML();
    expect(findInSdtPr(xml, 'w:placeholder')).toBeDefined();
    expect(findInSdtPr(xml, 'w:dataBinding')).toBeDefined();
    expect(findInSdtPr(xml, 'w:showingPlcHdr')).toBeDefined();
    expect(findInSdtPr(xml, 'w:text')).toBeDefined();
  });

  test('should round-trip all three properties', async () => {
    const doc = Document.create();
    const sdt = StructuredDocumentTag.createPlainText(
      [new Paragraph().addText('Placeholder text')],
      false,
    );
    sdt.setPlaceholder('DefaultPlaceholder_-1854013440');
    sdt.setDataBinding('/company/name', undefined, '{AABB1234-5678-9012-3456-789ABCDEF012}');
    sdt.setShowingPlaceholder(true);
    doc.addBodyElement(sdt);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedSdt = loaded.getBodyElements()
      .find(el => el instanceof StructuredDocumentTag) as StructuredDocumentTag;

    expect(loadedSdt.getPlaceholder()?.docPart).toBe('DefaultPlaceholder_-1854013440');
    expect(loadedSdt.getDataBinding()?.xpath).toBe('/company/name');
    expect(loadedSdt.getDataBinding()?.storeItemId).toBe('{AABB1234-5678-9012-3456-789ABCDEF012}');
    expect(loadedSdt.isShowingPlaceholder()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('SDT Citation Type', () => {
  test('should create citation control', () => {
    const sdt = StructuredDocumentTag.createCitation(
      [new Paragraph().addText('(Author, 2024)')],
    );
    expect(sdt.getControlType()).toBe('citation');
  });

  test('should generate w:citation in XML', () => {
    const sdt = StructuredDocumentTag.createCitation(
      [new Paragraph().addText('(Author, 2024)')],
    );

    const xml = sdt.toXML();
    expect(findInSdtPr(xml, 'w:citation')).toBeDefined();
  });

  test('should round-trip citation control', async () => {
    const doc = Document.create();
    const sdt = StructuredDocumentTag.createCitation(
      [new Paragraph().addText('(Smith, 2024)')],
    );
    doc.addBodyElement(sdt);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedSdt = loaded.getBodyElements()
      .find(el => el instanceof StructuredDocumentTag) as StructuredDocumentTag;

    expect(loadedSdt).toBeDefined();
    expect(loadedSdt.getControlType()).toBe('citation');

    doc.dispose();
    loaded.dispose();
  });
});

describe('SDT Bibliography Type', () => {
  test('should create bibliography control', () => {
    const sdt = StructuredDocumentTag.createBibliography(
      [new Paragraph().addText('Bibliography entries...')],
    );
    expect(sdt.getControlType()).toBe('bibliography');
  });

  test('should generate w:bibliography in XML', () => {
    const sdt = StructuredDocumentTag.createBibliography(
      [new Paragraph().addText('References')],
    );

    const xml = sdt.toXML();
    expect(findInSdtPr(xml, 'w:bibliography')).toBeDefined();
  });

  test('should round-trip bibliography control', async () => {
    const doc = Document.create();
    const sdt = StructuredDocumentTag.createBibliography(
      [new Paragraph().addText('Smith, J. (2024). Title.')],
    );
    doc.addBodyElement(sdt);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedSdt = loaded.getBodyElements()
      .find(el => el instanceof StructuredDocumentTag) as StructuredDocumentTag;

    expect(loadedSdt).toBeDefined();
    expect(loadedSdt.getControlType()).toBe('bibliography');

    doc.dispose();
    loaded.dispose();
  });
});

describe('SDT DocPartList Type', () => {
  test('should create docPartList control', () => {
    const sdt = StructuredDocumentTag.createDocPartList(
      'Table of Contents',
      'Built-In',
      [new Paragraph().addText('TOC placeholder')],
    );
    expect(sdt.getControlType()).toBe('docPartList');
    expect(sdt.getBuildingBlockProperties()?.gallery).toBe('Table of Contents');
    expect(sdt.getBuildingBlockProperties()?.isList).toBe(true);
  });

  test('should generate w:docPartList in XML', () => {
    const sdt = StructuredDocumentTag.createDocPartList(
      'Table of Contents',
      'Built-In',
      [new Paragraph().addText('TOC')],
    );

    const xml = sdt.toXML();
    const docPartList = findInSdtPr(xml, 'w:docPartList');
    expect(docPartList).toBeDefined();

    const gallery = filterXMLElements(docPartList?.children).find(
      c => c.name === 'w:docPartGallery'
    );
    expect(gallery?.attributes?.['w:val']).toBe('Table of Contents');
  });

  test('should round-trip docPartList control', async () => {
    const doc = Document.create();
    const sdt = StructuredDocumentTag.createDocPartList(
      'Table of Contents',
      'Built-In',
      [new Paragraph().addText('Contents here')],
    );
    doc.addBodyElement(sdt);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedSdt = loaded.getBodyElements()
      .find(el => el instanceof StructuredDocumentTag) as StructuredDocumentTag;

    expect(loadedSdt).toBeDefined();
    expect(loadedSdt.getControlType()).toBe('docPartList');
    expect(loadedSdt.getBuildingBlockProperties()?.gallery).toBe('Table of Contents');

    doc.dispose();
    loaded.dispose();
  });
});

describe('SDT Equation Type', () => {
  test('should create equation control', () => {
    const sdt = StructuredDocumentTag.createEquation(
      [new Paragraph().addText('E = mc²')],
    );
    expect(sdt.getControlType()).toBe('equation');
  });

  test('should generate w:equation in XML', () => {
    const sdt = StructuredDocumentTag.createEquation(
      [new Paragraph().addText('x² + y² = r²')],
    );

    const xml = sdt.toXML();
    expect(findInSdtPr(xml, 'w:equation')).toBeDefined();
  });
});

describe('SDT Group with Nested Content', () => {
  test('should create group with multiple content types', () => {
    const para1 = new Paragraph().addText('Section title');
    const para2 = new Paragraph().addText('Section body');
    const innerSdt = StructuredDocumentTag.createPlainText(
      [new Paragraph().addText('Editable field')],
      false,
    );

    const group = StructuredDocumentTag.createGroup([para1, innerSdt, para2]);
    expect(group.getControlType()).toBe('group');
    expect(group.getLock()).toBe('sdtContentLocked');
    expect(group.getContent()).toHaveLength(3);
  });

  test('should round-trip group with nested SDT', async () => {
    const doc = Document.create();
    const innerSdt = StructuredDocumentTag.createPlainText(
      [new Paragraph().addText('Inner')],
      false,
    );
    const group = StructuredDocumentTag.createGroup([
      new Paragraph().addText('Before'),
      innerSdt,
      new Paragraph().addText('After'),
    ]);
    doc.addBodyElement(group);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    const loadedGroup = loaded.getBodyElements()
      .find(el => el instanceof StructuredDocumentTag) as StructuredDocumentTag;

    expect(loadedGroup).toBeDefined();
    expect(loadedGroup.getControlType()).toBe('group');
    expect(loadedGroup.getContent().length).toBeGreaterThanOrEqual(3);

    // Find nested SDT
    const nestedSdt = loadedGroup.getContent()
      .find(c => c instanceof StructuredDocumentTag) as StructuredDocumentTag;
    expect(nestedSdt).toBeDefined();
    expect(nestedSdt.getControlType()).toBe('plainText');

    doc.dispose();
    loaded.dispose();
  });
});
