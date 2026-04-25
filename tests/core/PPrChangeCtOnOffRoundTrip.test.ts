/**
 * Tracked Changes — CT_OnOff Round-Trip Tests for w:pPrChange
 *
 * Per ECMA-376 Part 1 §17.17.4, the ST_OnOff simple type accepts:
 *   - "1", "true", "on"   → true
 *   - "0", "false", "off" → false
 *   - Absent w:val attribute on a present element → true (presence = on)
 *
 * The pPrChange parser/generator must honour all of these so that tracked
 * paragraph-property changes survive a load → serialize → reload cycle
 * regardless of which textual form was used in the source document.
 *
 * Bugs this suite guards against:
 *   - Parser: the pattern `obj['@_w:val'] !== '0'` returns TRUE for the
 *     literal strings "false" and "off", silently flipping the recorded
 *     previous value.
 *   - Generator: `suppressAutoHyphens` / `suppressLineNumbers` inside
 *     pPrChange only emitted `<w:x/>` when true — an explicit `false`
 *     was dropped, collapsing the distinction between "explicitly off"
 *     and "inherited".
 *   - Parser: `<w:spacing w:beforeAutospacing="on"/>` was parsed as
 *     false because only "1"/"true" were accepted.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithPPrChange(prevPPrXml: string): Promise<Buffer> {
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
    <w:p>
      <w:pPr>
        <w:pPrChange w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z">
          <w:pPr>${prevPPrXml}</w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );

  return await zipHandler.toBuffer();
}

async function loadPreservedProperties(
  prevPPrXml: string
): Promise<Record<string, unknown> | undefined> {
  const buffer = await makeDocxWithPPrChange(prevPPrXml);
  const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
  const change = doc.getParagraphs()[0]!.formatting.pPrChange;
  const props = change?.previousProperties as Record<string, unknown> | undefined;
  doc.dispose();
  return props;
}

describe('pPrChange CT_OnOff parsing — accepts all ST_OnOff forms', () => {
  // CT_OnOff properties parsed by DocumentParser for pPrChange previousProperties.
  // (keepNext is the canonical example; the table below exercises one form per property.)
  const ctOnOffProps: ReadonlyArray<{
    xmlName: string;
    field: string;
  }> = [
    { xmlName: 'w:keepNext', field: 'keepNext' },
    { xmlName: 'w:keepLines', field: 'keepLines' },
    { xmlName: 'w:pageBreakBefore', field: 'pageBreakBefore' },
    { xmlName: 'w:widowControl', field: 'widowControl' },
    { xmlName: 'w:suppressAutoHyphens', field: 'suppressAutoHyphens' },
    { xmlName: 'w:contextualSpacing', field: 'contextualSpacing' },
    { xmlName: 'w:mirrorIndents', field: 'mirrorIndents' },
    { xmlName: 'w:bidi', field: 'bidi' },
    { xmlName: 'w:suppressLineNumbers', field: 'suppressLineNumbers' },
    { xmlName: 'w:adjustRightInd', field: 'adjustRightInd' },
    { xmlName: 'w:snapToGrid', field: 'snapToGrid' },
    { xmlName: 'w:wordWrap', field: 'wordWrap' },
    { xmlName: 'w:autoSpaceDE', field: 'autoSpaceDE' },
    { xmlName: 'w:autoSpaceDN', field: 'autoSpaceDN' },
  ];

  for (const { xmlName, field } of ctOnOffProps) {
    it(`parses <${xmlName} w:val="false"/> as false`, async () => {
      const props = await loadPreservedProperties(`<${xmlName} w:val="false"/>`);
      expect(props?.[field]).toBe(false);
    });

    it(`parses <${xmlName} w:val="off"/> as false`, async () => {
      const props = await loadPreservedProperties(`<${xmlName} w:val="off"/>`);
      expect(props?.[field]).toBe(false);
    });

    it(`parses <${xmlName} w:val="true"/> as true`, async () => {
      const props = await loadPreservedProperties(`<${xmlName} w:val="true"/>`);
      expect(props?.[field]).toBe(true);
    });

    it(`parses <${xmlName} w:val="on"/> as true`, async () => {
      const props = await loadPreservedProperties(`<${xmlName} w:val="on"/>`);
      expect(props?.[field]).toBe(true);
    });

    it(`parses <${xmlName}/> (no attribute) as true`, async () => {
      const props = await loadPreservedProperties(`<${xmlName}/>`);
      expect(props?.[field]).toBe(true);
    });

    it(`parses <${xmlName} w:val="0"/> as false`, async () => {
      const props = await loadPreservedProperties(`<${xmlName} w:val="0"/>`);
      expect(props?.[field]).toBe(false);
    });

    it(`parses <${xmlName} w:val="1"/> as true`, async () => {
      const props = await loadPreservedProperties(`<${xmlName} w:val="1"/>`);
      expect(props?.[field]).toBe(true);
    });
  }
});

describe('pPrChange spacing autospacing — accepts all ST_OnOff forms', () => {
  it('parses beforeAutospacing="on" as true', async () => {
    const props = await loadPreservedProperties(`<w:spacing w:beforeAutospacing="on"/>`);
    const spacing = props?.spacing as { beforeAutospacing?: boolean } | undefined;
    expect(spacing?.beforeAutospacing).toBe(true);
  });

  it('parses beforeAutospacing="off" as false', async () => {
    const props = await loadPreservedProperties(`<w:spacing w:beforeAutospacing="off"/>`);
    const spacing = props?.spacing as { beforeAutospacing?: boolean } | undefined;
    expect(spacing?.beforeAutospacing).toBe(false);
  });

  it('parses afterAutospacing="on" as true', async () => {
    const props = await loadPreservedProperties(`<w:spacing w:afterAutospacing="on"/>`);
    const spacing = props?.spacing as { afterAutospacing?: boolean } | undefined;
    expect(spacing?.afterAutospacing).toBe(true);
  });

  it('parses afterAutospacing="false" as false', async () => {
    const props = await loadPreservedProperties(`<w:spacing w:afterAutospacing="false"/>`);
    const spacing = props?.spacing as { afterAutospacing?: boolean } | undefined;
    expect(spacing?.afterAutospacing).toBe(false);
  });
});

describe('pPrChange generator — preserves explicit false for CT_OnOff', () => {
  // These two properties previously generated nothing when false,
  // collapsing "explicitly off" and "inherited" into the same output.
  it('serializes suppressAutoHyphens=false as <w:suppressAutoHyphens w:val="0"/>', () => {
    const para = new Paragraph();
    para.addText('test');
    para.setParagraphPropertiesChange({
      id: '1',
      author: 'Tester',
      date: '2024-01-01T00:00:00Z',
      previousProperties: { suppressAutoHyphens: false },
    });

    const xml = XMLBuilder.elementToString(para.toXML());
    const changeStart = xml.indexOf('<w:pPrChange');
    expect(changeStart).toBeGreaterThan(-1);
    const changeXml = xml.substring(changeStart);

    expect(changeXml).toContain('<w:suppressAutoHyphens');
    expect(changeXml).toMatch(/<w:suppressAutoHyphens[^>]*w:val="0"/);
  });

  it('serializes suppressLineNumbers=false as <w:suppressLineNumbers w:val="0"/>', () => {
    const para = new Paragraph();
    para.addText('test');
    para.setParagraphPropertiesChange({
      id: '2',
      author: 'Tester',
      date: '2024-01-01T00:00:00Z',
      previousProperties: { suppressLineNumbers: false },
    });

    const xml = XMLBuilder.elementToString(para.toXML());
    const changeStart = xml.indexOf('<w:pPrChange');
    expect(changeStart).toBeGreaterThan(-1);
    const changeXml = xml.substring(changeStart);

    expect(changeXml).toContain('<w:suppressLineNumbers');
    expect(changeXml).toMatch(/<w:suppressLineNumbers[^>]*w:val="0"/);
  });

  it('serializes suppressAutoHyphens=true as <w:suppressAutoHyphens w:val="1"/>', () => {
    const para = new Paragraph();
    para.addText('test');
    para.setParagraphPropertiesChange({
      id: '3',
      author: 'Tester',
      date: '2024-01-01T00:00:00Z',
      previousProperties: { suppressAutoHyphens: true },
    });

    const xml = XMLBuilder.elementToString(para.toXML());
    const changeStart = xml.indexOf('<w:pPrChange');
    const changeXml = xml.substring(changeStart);

    expect(changeXml).toMatch(/<w:suppressAutoHyphens[^>]*w:val="1"/);
  });

  it('omits suppressAutoHyphens entirely when undefined', () => {
    const para = new Paragraph();
    para.addText('test');
    para.setParagraphPropertiesChange({
      id: '4',
      author: 'Tester',
      date: '2024-01-01T00:00:00Z',
      previousProperties: { keepNext: true }, // something else so pPrChange emits
    });

    const xml = XMLBuilder.elementToString(para.toXML());
    const changeStart = xml.indexOf('<w:pPrChange');
    const changeXml = xml.substring(changeStart);

    expect(changeXml).not.toContain('<w:suppressAutoHyphens');
    expect(changeXml).not.toContain('<w:suppressLineNumbers');
  });
});

describe('pPrChange CT_OnOff full round-trip', () => {
  it('round-trips suppressAutoHyphens=false through generate → parse', async () => {
    const doc = Document.create();
    const para = new Paragraph();
    para.addText('test');
    para.setParagraphPropertiesChange({
      id: '10',
      author: 'Tester',
      date: '2024-01-01T00:00:00Z',
      previousProperties: { suppressAutoHyphens: false, keepNext: false },
    });
    doc.addParagraph(para);

    const buffer = await doc.toBuffer();
    doc.dispose();

    const reloaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = reloaded.getParagraphs()[0]!.formatting.pPrChange;
    const props = change?.previousProperties as Record<string, unknown> | undefined;

    expect(props?.suppressAutoHyphens).toBe(false);
    expect(props?.keepNext).toBe(false);
    reloaded.dispose();
  });

  it('round-trips suppressLineNumbers=false through generate → parse', async () => {
    const doc = Document.create();
    const para = new Paragraph();
    para.addText('test');
    para.setParagraphPropertiesChange({
      id: '11',
      author: 'Tester',
      date: '2024-01-01T00:00:00Z',
      previousProperties: { suppressLineNumbers: false },
    });
    doc.addParagraph(para);

    const buffer = await doc.toBuffer();
    doc.dispose();

    const reloaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = reloaded.getParagraphs()[0]!.formatting.pPrChange;
    const props = change?.previousProperties as Record<string, unknown> | undefined;

    expect(props?.suppressLineNumbers).toBe(false);
    reloaded.dispose();
  });
});
