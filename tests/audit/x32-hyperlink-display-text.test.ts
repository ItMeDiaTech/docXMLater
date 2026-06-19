/**
 * Field.createHyperlink display text — fldSimple cached result.
 *
 * A HYPERLINK field's visible text is its cached result; it is NOT
 * derivable from the field instruction. Field.createHyperlink accepted a
 * displayText parameter (documented as "The text to display") but never
 * used it: the emitted <w:fldSimple> run text came from the hardcoded
 * 'Link' placeholder, so the caller's text was silently lost and the
 * document rendered 'Link' until a manual field refresh (F9).
 *
 * These tests pin the fixed behavior: the cached result run text is the
 * caller's displayText, falling back to the URL when displayText is
 * omitted.
 */

import { Document } from '../../src/core/Document';
import { Field } from '../../src/elements/Field';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLElement } from '../../src/xml/XMLBuilder';

function getFldSimpleRunText(xml: XMLElement): string {
  const run = xml.children!.find(
    (child): child is XMLElement => typeof child !== 'string' && child.name === 'w:r'
  );
  const textElement = run?.children?.find(
    (child): child is XMLElement => typeof child !== 'string' && child.name === 'w:t'
  );
  return (textElement?.children ?? []).filter((c): c is string => typeof c === 'string').join('');
}

describe('Field.createHyperlink display text', () => {
  it('emits the caller displayText as the fldSimple run text', () => {
    const field = Field.createHyperlink('https://example.com', 'Click here');

    const xml = field.toXML();
    expect(xml.name).toBe('w:fldSimple');
    expect(xml.attributes!['w:instr']).toContain('HYPERLINK "https://example.com"');
    expect(getFldSimpleRunText(xml)).toBe('Click here');
  });

  it('falls back to the URL when displayText is omitted', () => {
    const field = Field.createHyperlink('https://example.com');

    expect(getFldSimpleRunText(field.toXML())).toBe('https://example.com');
  });

  it('keeps the displayText alongside a tooltip and formatting', () => {
    const field = Field.createHyperlink('https://example.com', 'Docs', 'Visit Example', {
      bold: true,
    });

    const xml = field.toXML();
    expect(xml.attributes!['w:instr']).toContain('\\o "Visit Example"');
    expect(getFldSimpleRunText(xml)).toBe('Docs');
  });

  it('saves the displayText into word/document.xml', async () => {
    const doc = Document.create();
    try {
      const para = doc.createParagraph();
      para.addField(Field.createHyperlink('https://example.com', 'Click here'));

      const buffer = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const content = zip.getFile('word/document.xml')?.content;
      const documentXml = content instanceof Buffer ? content.toString('utf8') : String(content);

      const fldSimple = documentXml.match(/<w:fldSimple[\s\S]*?<\/w:fldSimple>/)?.[0] ?? '';
      expect(fldSimple).toContain('Click here');
      expect(fldSimple).not.toContain('>Link<');
    } finally {
      doc.dispose();
    }
  });
});
