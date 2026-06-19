/**
 * XMLParser.parseToObject must decode XML entities in attribute values the
 * same way it already decodes them in text nodes (and extractAttribute does
 * for single attributes). Storing the raw escaped substring means every
 * serializer that escapes on output (XMLBuilder.escapeXmlAttribute, the
 * raw-XML revision acceptor) double-escapes: w:instr=" HYPERLINK
 * &quot;...&amp;...&quot; " becomes &amp;quot;/&amp;amp; on save, silently
 * corrupting HYPERLINK field instructions, tooltips, and author names on
 * every load/save cycle — compounding with each round-trip.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLParser } from '../../src/xml/XMLParser';

const FLD_SIMPLE_PARAGRAPH =
  `<w:p>` +
  `<w:fldSimple w:instr=" HYPERLINK &quot;https://example.com/?a=1&amp;b=2&quot; ">` +
  `<w:r><w:t>example link</w:t></w:r>` +
  `</w:fldSimple>` +
  `</w:p>`;

async function buildDocxWithFldSimple(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Entity attribute round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  const docXml = zip.getFileAsString('word/document.xml')!;
  const updated = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${FLD_SIMPLE_PARAGRAPH}<w:sectPr`)
    : docXml.replace('</w:body>', `${FLD_SIMPLE_PARAGRAPH}</w:body>`);
  zip.updateFile('word/document.xml', updated);

  return zip.toBuffer();
}

describe('parseToObject attribute entity decoding', () => {
  it('decodes XML entities in attribute values to the actual characters', () => {
    const parsed = XMLParser.parseToObject(
      '<w:fldSimple w:instr=" HYPERLINK &quot;http://x.com/?a=1&amp;b=2&quot; "/>',
      { trimValues: false }
    ) as any;

    expect(parsed['w:fldSimple']['@_w:instr']).toBe(' HYPERLINK "http://x.com/?a=1&b=2" ');
  });

  it('decodes attribute values and text nodes consistently', () => {
    const parsed = XMLParser.parseToObject('<w:x w:val="a &amp; b">a &amp; b</w:x>') as any;

    expect(parsed['w:x']['@_w:val']).toBe('a & b');
    expect(parsed['w:x']['#text']).toBe('a & b');
  });

  it('does not double-escape w:instr entities on a plain load/save round-trip', async () => {
    const buf = await buildDocxWithFldSimple();
    const doc = await Document.loadFromBuffer(buf);
    try {
      const out = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const savedXml = zip.getFileAsString('word/document.xml')!;

      // Single-escaped form survives; double-escaped corruption does not appear.
      expect(savedXml).toContain('&quot;https://example.com/?a=1&amp;b=2&quot;');
      expect(savedXml).not.toContain('&amp;quot;');
      expect(savedXml).not.toContain('&amp;amp;');
    } finally {
      doc.dispose();
    }
  });

  it('does not double-escape w:instr entities after a mutation forces full regeneration', async () => {
    const buf = await buildDocxWithFldSimple();
    const doc = await Document.loadFromBuffer(buf);
    try {
      // Mutating the document forces document.xml to be regenerated from the
      // model rather than passed through verbatim — the path that re-escapes
      // if attribute values were stored already-escaped.
      doc.createParagraph('forces document.xml regeneration');
      const out = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const savedXml = zip.getFileAsString('word/document.xml')!;

      expect(savedXml).toContain('&quot;https://example.com/?a=1&amp;b=2&quot;');
      expect(savedXml).not.toContain('&amp;quot;');
      expect(savedXml).not.toContain('&amp;amp;');
    } finally {
      doc.dispose();
    }
  });
});
