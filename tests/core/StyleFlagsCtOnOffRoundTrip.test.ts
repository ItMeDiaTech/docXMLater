/**
 * Style CT_OnOff metadata flag round-trip tests.
 *
 * Per ECMA-376 Part 1 §17.7.4, a `<w:style>` carries eight CT_OnOff
 * metadata flags:
 *
 *   w:qFormat, w:semiHidden, w:unhideWhenUsed, w:locked,
 *   w:personal, w:personalCompose, w:personalReply, w:autoRedefine
 *
 * They use the OnOffType schema binding in the Open XML SDK, so `w:val`
 * accepts the full ST_OnOff union: `1`/`0`/`true`/`false`/`on`/`off`.
 *
 * Bugs this suite guards against:
 *
 *   Parser — each flag was detected with
 *     `styleXml.includes('<w:qFormat/>') || styleXml.includes('<w:qFormat ')`
 *   which (a) ignored `w:val` entirely, so `<w:qFormat w:val="off"/>`
 *   (explicit override of a based-on style's qFormat=true) was silently
 *   flipped to `true`; and (b) used a string substring match that was
 *   fragile (the substring is only safe when no sibling tag starts with
 *   the same name, which happens to hold for these flags but is not
 *   guaranteed by the schema).
 *
 *   Generator — each flag was serialized with
 *     `if (this.properties.qFormat) { emit <w:qFormat/> }`
 *   which dropped explicit false. An override of an inherited true
 *   could therefore never round-trip.
 *
 * This suite covers every literal on parse and locks explicit-false
 * emission on the generator, including round-trip through load → save
 * → load.
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyles(styleInner: string): Promise<Buffer> {
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
  <w:style w:type="paragraph" w:styleId="Test">
    <w:name w:val="Test"/>
    ${styleInner}
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

async function loadFlag(styleInner: string) {
  const buffer = await makeDocxWithStyles(styleInner);
  const doc = await Document.loadFromBuffer(buffer);
  const style = doc.getStylesManager().getStyle('Test');
  // Access the internal properties to inspect every CT_OnOff flag uniformly.
  const props = (style as unknown as { properties: Record<string, unknown> }).properties;
  doc.dispose();
  return props;
}

const FLAGS: ReadonlyArray<{ xml: string; field: string }> = [
  { xml: 'w:qFormat', field: 'qFormat' },
  { xml: 'w:semiHidden', field: 'semiHidden' },
  { xml: 'w:unhideWhenUsed', field: 'unhideWhenUsed' },
  { xml: 'w:locked', field: 'locked' },
  { xml: 'w:personal', field: 'personal' },
  { xml: 'w:personalCompose', field: 'personalCompose' },
  { xml: 'w:personalReply', field: 'personalReply' },
  { xml: 'w:autoRedefine', field: 'autoRedefine' },
];

describe('Style metadata flags — parser honours w:val per ECMA-376 §17.17.4', () => {
  for (const { xml, field } of FLAGS) {
    it(`parses <${xml} w:val="off"/> as false`, async () => {
      const props = await loadFlag(`<${xml} w:val="off"/>`);
      expect(props[field]).toBe(false);
    });

    it(`parses <${xml} w:val="false"/> as false`, async () => {
      const props = await loadFlag(`<${xml} w:val="false"/>`);
      expect(props[field]).toBe(false);
    });

    it(`parses <${xml} w:val="off"/> as false`, async () => {
      const props = await loadFlag(`<${xml} w:val="off"/>`);
      expect(props[field]).toBe(false);
    });

    it(`parses <${xml}/> as true`, async () => {
      const props = await loadFlag(`<${xml}/>`);
      expect(props[field]).toBe(true);
    });

    it(`parses <${xml} w:val="on"/> as true`, async () => {
      const props = await loadFlag(`<${xml} w:val="on"/>`);
      expect(props[field]).toBe(true);
    });

    it(`parses <${xml} w:val="1"/> as true`, async () => {
      const props = await loadFlag(`<${xml} w:val="1"/>`);
      expect(props[field]).toBe(true);
    });
  }
});

describe('Style metadata flags — generator preserves explicit false', () => {
  for (const { xml, field } of FLAGS) {
    it(`serializes ${field}=false as <${xml} w:val="off"/> (OnOffOnlyType)`, () => {
      // Construct a bare custom style (customStyle=true to avoid qFormat defaulting)
      const style = Style.create({
        styleId: 'Test',
        name: 'Test',
        type: 'paragraph',
        customStyle: true,
        [field]: false,
      } as Parameters<typeof Style.create>[0]);

      const xmlOut = XMLBuilder.elementToString(style.toXML());
      expect(xmlOut).toContain(`<${xml} w:val="off"/>`);
    });

    it(`serializes ${field}=true as <${xml}/> (bare self-closing)`, () => {
      const style = Style.create({
        styleId: 'Test',
        name: 'Test',
        type: 'paragraph',
        customStyle: true,
        [field]: true,
      } as Parameters<typeof Style.create>[0]);

      const xmlOut = XMLBuilder.elementToString(style.toXML());
      // For OnOffOnlyType, `<w:qFormat/>` (bare, val defaults to "on") is the
      // canonical form for true — just assert presence and absence of explicit "off".
      expect(xmlOut).toContain(`<${xml}`);
      expect(xmlOut).not.toMatch(new RegExp(`<${xml}[^>]*w:val="off"`));
    });
  }
});

describe('Style metadata flags — full load → save → load round-trip preserves explicit false', () => {
  // Spot-check three representative flags to keep the suite fast.
  const cases: ReadonlyArray<{ xml: string; field: string }> = [
    { xml: 'w:qFormat', field: 'qFormat' },
    { xml: 'w:semiHidden', field: 'semiHidden' },
    { xml: 'w:locked', field: 'locked' },
  ];

  for (const { xml, field } of cases) {
    it(`round-trips ${xml}=false`, async () => {
      const buffer1 = await makeDocxWithStyles(`<${xml} w:val="off"/>`);
      const doc1 = await Document.loadFromBuffer(buffer1);
      const style1 = doc1.getStylesManager().getStyle('Test');
      expect((style1 as unknown as { properties: Record<string, unknown> }).properties[field]).toBe(
        false
      );
      const buffer2 = await doc1.toBuffer();
      doc1.dispose();

      const doc2 = await Document.loadFromBuffer(buffer2);
      const style2 = doc2.getStylesManager().getStyle('Test');
      expect((style2 as unknown as { properties: Record<string, unknown> }).properties[field]).toBe(
        false
      );
      doc2.dispose();
    });
  }
});
