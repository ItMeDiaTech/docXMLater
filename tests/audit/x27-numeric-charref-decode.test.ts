/**
 * XMLBuilder.unescapeXml must expand numeric character references
 * (&#8217; / &#xA0;) — XML 1.0 §4.1 makes CharRef expansion mandatory for
 * every conforming processor, and third-party DOCX producers emit them.
 * Leaving them raw means escapeXmlText re-escapes the ampersand on save
 * (&amp;#8217;), so Word renders the literal text "&#8217;" instead of the
 * intended character. Code points outside the XML 1.0 Char production
 * (NUL, surrogates) must stay as raw references, never decoded.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { XMLParser } from '../../src/xml/XMLParser';

const CHARREF_PARAGRAPH = `<w:p><w:r><w:t xml:space="preserve">It&#8217;s a test&#xA0;here</w:t></w:r></w:p>`;

async function buildDocxWithCharRefs(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Numeric character reference round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  const docXml = zip.getFileAsString('word/document.xml')!;
  const updated = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${CHARREF_PARAGRAPH}<w:sectPr`)
    : docXml.replace('</w:body>', `${CHARREF_PARAGRAPH}</w:body>`);
  zip.updateFile('word/document.xml', updated);

  return zip.toBuffer();
}

describe('unescapeXml numeric character references', () => {
  it('decodes decimal and hexadecimal character references', () => {
    expect(XMLBuilder.unescapeXml('curly &#8217; quote')).toBe('curly ’ quote');
    expect(XMLBuilder.unescapeXml('nbsp&#xA0;here')).toBe('nbsp here');
    expect(XMLBuilder.unescapeXml('astral &#x1F600;')).toBe('astral \u{1F600}');
  });

  it('still decodes the five named entities, with &amp; not double-decoded', () => {
    expect(XMLBuilder.unescapeXml('&lt;a&gt; &quot;b&quot; &apos;c&apos; &amp;')).toBe(
      '<a> "b" \'c\' &'
    );
    // Entity text produced by one decode must not be re-decoded.
    expect(XMLBuilder.unescapeXml('&amp;lt;')).toBe('&lt;');
    expect(XMLBuilder.unescapeXml('&#38;amp;')).toBe('&amp;');
  });

  it('keeps references outside the XML 1.0 Char production raw', () => {
    expect(XMLBuilder.unescapeXml('&#0;')).toBe('&#0;');
    expect(XMLBuilder.unescapeXml('&#x0;')).toBe('&#x0;');
    expect(XMLBuilder.unescapeXml('&#xD800;')).toBe('&#xD800;');
    expect(XMLBuilder.unescapeXml('&#xDFFF;')).toBe('&#xDFFF;');
    expect(XMLBuilder.unescapeXml('&#x110000;')).toBe('&#x110000;');
  });

  it('does not double-escape a decoded reference on re-serialization', () => {
    const escaped = XMLBuilder.escapeXmlText(XMLBuilder.unescapeXml('curly &#8217; quote'));
    expect(escaped).toBe('curly ’ quote');
    expect(escaped).not.toContain('&amp;#');
  });

  it('decodes numeric references in attribute values via parseToObject', () => {
    const parsed = XMLParser.parseToObject('<w:fldSimple w:instr="a&#8217;b&#xA0;c"/>') as any;
    expect(parsed['w:fldSimple']['@_w:instr']).toBe('a’b c');
  });

  it('round-trips numeric references as real characters, not literal &#...; text', async () => {
    const buf = await buildDocxWithCharRefs();
    const doc = await Document.loadFromBuffer(buf);
    try {
      const allText = doc
        .getParagraphs()
        .map((p) => p.getText())
        .join('\n');
      expect(allText).toContain('It’s a test here');

      const out = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const savedXml = zip.getFileAsString('word/document.xml')!;

      // Double-escaped corruption (rendered by Word as literal "&#8217;") never appears.
      expect(savedXml).not.toContain('&amp;#8217;');
      expect(savedXml).not.toContain('&amp;#xA0;');
    } finally {
      doc.dispose();
    }
  });
});
