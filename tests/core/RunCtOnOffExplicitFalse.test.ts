/**
 * Run CT_OnOff explicit-false round-trip tests.
 *
 * Ten `w:rPr` boolean elements were dropping the explicit-false case on
 * both sides of the pipeline:
 *
 *   w:outline, w:shadow, w:emboss, w:imprint, w:noProof,
 *   w:vanish, w:specVanish, w:webHidden, w:rtl, w:cs
 *
 * Parser side (`parseRunFromObject` for the main rPr):
 *   if (parseOoxmlBoolean(rPrObj['w:outline'])) run.setOutline(true);
 * — correctly parses the ST_OnOff literal but only sets the formatting
 * field when the result is true. Explicit `<w:outline w:val="0"/>`
 * (override an inherited true) therefore left `formatting.outline` as
 * `undefined`, collapsing into "inherit from style".
 *
 * Generator side (`Run.generateRunPropertiesXML`, used for BOTH the main
 * rPr AND for rPrChange previousProperties):
 *   if (formatting.outline) { emit <w:outline w:val="1"/> }
 * — only emits when true, so a model value of `false` became invisible
 * in the output and could not round-trip. That corrupted tracked
 * `w:rPrChange` history: a recorded "previous: outline=false" silently
 * dropped its override intent.
 *
 * All ten elements are plain `CT_OnOff` per ECMA-376 §17.3.2 (OnOffType
 * in the Open XML SDK, NOT OnOffOnlyType) and accept every ST_OnOff
 * literal. This suite locks parse-and-emit symmetry for explicit false
 * on all ten, using the main rPr path (the rPrChange generator shares
 * the same `generateRunPropertiesXML` function, so fixing it covers
 * tracked changes too).
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithRunRPr(rPrInner: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr>${rPrInner}</w:rPr><w:t>test</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );

  return await zipHandler.toBuffer();
}

describe('Run main-path parser honours explicit-false CT_OnOff', () => {
  const cases: Array<{ xml: string; field: string }> = [
    { xml: 'w:outline', field: 'outline' },
    { xml: 'w:shadow', field: 'shadow' },
    { xml: 'w:emboss', field: 'emboss' },
    { xml: 'w:imprint', field: 'imprint' },
    { xml: 'w:noProof', field: 'noProof' },
    { xml: 'w:vanish', field: 'vanish' },
    { xml: 'w:specVanish', field: 'specVanish' },
    { xml: 'w:webHidden', field: 'webHidden' },
    { xml: 'w:rtl', field: 'rtl' },
    { xml: 'w:cs', field: 'complexScript' },
  ];

  for (const { xml, field } of cases) {
    it(`parses <${xml} w:val="0"/> as false`, async () => {
      const buffer = await makeDocxWithRunRPr(`<${xml} w:val="0"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const run = doc.getParagraphs()[0]!.getRuns()[0]!;
      const fmt = run.getFormatting() as Record<string, unknown>;
      expect(fmt[field]).toBe(false);
      doc.dispose();
    });

    it(`parses <${xml} w:val="false"/> as false`, async () => {
      const buffer = await makeDocxWithRunRPr(`<${xml} w:val="false"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const run = doc.getParagraphs()[0]!.getRuns()[0]!;
      const fmt = run.getFormatting() as Record<string, unknown>;
      expect(fmt[field]).toBe(false);
      doc.dispose();
    });

    it(`parses <${xml}/> as true`, async () => {
      const buffer = await makeDocxWithRunRPr(`<${xml}/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const run = doc.getParagraphs()[0]!.getRuns()[0]!;
      const fmt = run.getFormatting() as Record<string, unknown>;
      expect(fmt[field]).toBe(true);
      doc.dispose();
    });
  }
});

describe('Run generator emits w:val="0" for explicit-false CT_OnOff', () => {
  const cases = [
    { xml: 'w:outline', field: 'outline' as const, setter: 'setOutline' as const },
    { xml: 'w:shadow', field: 'shadow' as const, setter: 'setShadow' as const },
    { xml: 'w:emboss', field: 'emboss' as const, setter: 'setEmboss' as const },
    { xml: 'w:imprint', field: 'imprint' as const, setter: 'setImprint' as const },
    { xml: 'w:noProof', field: 'noProof' as const, setter: 'setNoProof' as const },
    { xml: 'w:vanish', field: 'vanish' as const, setter: 'setVanish' as const },
    { xml: 'w:specVanish', field: 'specVanish' as const, setter: 'setSpecVanish' as const },
    { xml: 'w:webHidden', field: 'webHidden' as const, setter: 'setWebHidden' as const },
    { xml: 'w:rtl', field: 'rtl' as const, setter: 'setRTL' as const },
    { xml: 'w:cs', field: 'complexScript' as const, setter: 'setComplexScript' as const },
  ];

  for (const { xml, setter } of cases) {
    it(`serializes ${String(setter)}(false) as <${xml} w:val="0"/>`, () => {
      const run = new Run('test');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (run as any)[setter](false);
      const xmlOut = XMLBuilder.elementToString(run.toXML());
      expect(xmlOut).toContain(`<${xml} w:val="0"/>`);
    });

    it(`serializes ${String(setter)}(true) without w:val="0"`, () => {
      const run = new Run('test');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (run as any)[setter](true);
      const xmlOut = XMLBuilder.elementToString(run.toXML());
      expect(xmlOut).toContain(`<${xml}`);
      expect(xmlOut).not.toMatch(new RegExp(`<${xml}[^>]*w:val="0"`));
    });
  }
});

describe('Run CT_OnOff full round-trip (load → save → load preserves explicit false)', () => {
  const cases = [
    { xml: 'w:outline', field: 'outline' as const },
    { xml: 'w:shadow', field: 'shadow' as const },
    { xml: 'w:vanish', field: 'vanish' as const },
    { xml: 'w:rtl', field: 'rtl' as const },
    { xml: 'w:cs', field: 'complexScript' as const },
  ];

  for (const { xml, field } of cases) {
    it(`round-trips ${xml}=false`, async () => {
      const buffer1 = await makeDocxWithRunRPr(`<${xml} w:val="0"/>`);
      const doc1 = await Document.loadFromBuffer(buffer1);
      expect((doc1.getParagraphs()[0]!.getRuns()[0]!.getFormatting() as any)[field]).toBe(false);

      const buffer2 = await doc1.toBuffer();
      doc1.dispose();

      const doc2 = await Document.loadFromBuffer(buffer2);
      expect((doc2.getParagraphs()[0]!.getRuns()[0]!.getFormatting() as any)[field]).toBe(false);
      doc2.dispose();
    });
  }
});

describe('rPrChange (tracked run property change) preserves explicit false', () => {
  it('emits <w:outline w:val="0"/> inside rPrChange when previous was explicitly false', () => {
    const run = new Run('test');
    run.setOutline(true);
    run.setPropertyChangeRevision({
      id: 1,
      author: 'Tester',
      date: new Date('2024-01-01T00:00:00Z'),
      previousProperties: { outline: false },
    });

    const xml = XMLBuilder.elementToString(run.toXML());
    const changeStart = xml.indexOf('<w:rPrChange');
    expect(changeStart).toBeGreaterThan(-1);
    const changeXml = xml.substring(changeStart);
    expect(changeXml).toContain('<w:outline w:val="0"/>');
  });

  it('emits <w:rtl w:val="0"/> inside rPrChange when previous was explicitly false', () => {
    const run = new Run('test');
    run.setRTL(true);
    run.setPropertyChangeRevision({
      id: 2,
      author: 'Tester',
      date: new Date('2024-01-01T00:00:00Z'),
      previousProperties: { rtl: false },
    });

    const xml = XMLBuilder.elementToString(run.toXML());
    const changeStart = xml.indexOf('<w:rPrChange');
    const changeXml = xml.substring(changeStart);
    expect(changeXml).toContain('<w:rtl w:val="0"/>');
  });
});

// Silence unused-variable lint for the Paragraph import (kept for future uses)
void Paragraph;
