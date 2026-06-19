/**
 * CDATA sections (XML 1.0 §2.7) are legal anywhere character data is allowed
 * and their payload is by definition unescaped.
 *
 * The parser special-cased only comments; '<![CDATA[' failed the
 * element-name match and was consumed one character at a time, dropping
 * every '<' in the payload, leaking the '![CDATA[' / ']]>' markers into text,
 * and applying entity decoding to content that must be taken verbatim.
 * Processing instructions other than the stripped XML declaration hit the
 * same one-char-skip fallback.
 */

import { XMLParser } from '../../src/xml/XMLParser';

describe('XMLParser.parseToObject — CDATA and processing instructions', () => {
  it("preserves a CDATA payload containing '<', '&', and '>' verbatim", () => {
    const result: any = XMLParser.parseToObject('<root><![CDATA[5 < 6 & 7 > 2]]></root>');

    expect(result.root).toBe('5 < 6 & 7 > 2');
  });

  it('does not apply entity decoding to CDATA content', () => {
    const result: any = XMLParser.parseToObject('<root><![CDATA[a &lt; b]]></root>');

    // '&lt;' inside CDATA is literal text, not an entity reference
    expect(result.root).toBe('a &lt; b');
  });

  it('combines CDATA with surrounding text content', () => {
    const result: any = XMLParser.parseToObject(
      '<root>before <![CDATA[<raw & stuff>]]> after</root>'
    );

    expect(result.root).toBe('before <raw & stuff> after');
  });

  it('skips a processing instruction before the root element', () => {
    const result: any = XMLParser.parseToObject(
      '<?mso-application progid="Word.Document"?><root><a/></root>'
    );

    expect(result.root.a).toEqual({});
  });

  it('skips a processing instruction inside element content', () => {
    const result: any = XMLParser.parseToObject('<root><?pi data?><a/></root>');

    expect(result.root.a).toEqual({});
    // The PI body must not leak into text content
    expect(JSON.stringify(result)).not.toContain('pi data');
  });
});
