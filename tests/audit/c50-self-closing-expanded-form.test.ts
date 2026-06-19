/**
 * extractSelfClosingTag — the result must be bounded to the named element's
 * own tag header.
 *
 * Expanded-form empty elements (<tag a="b"></tag>) are well-formed XML 1.0
 * and infoset-identical to the self-closing form; several non-Word
 * serializers emit them. A global indexOf('/>') search lands inside a later
 * sibling whenever the target is written in expanded form, returning a
 * fragment that spans multiple elements and attributing a sibling's
 * attributes (e.g. w:themeColor of w:u) to the target. This helper backs the
 * rPr/pPr/tblPr string-parsing paths in DocumentParser, so the misparse
 * propagates into the model and the saved document.
 */

import { XMLParser } from '../../src/xml/XMLParser';

describe('XMLParser.extractSelfClosingTag — expanded-form empty elements', () => {
  it('returns only the target attributes when the tag uses expanded form', () => {
    const xml = '<w:rPr><w:highlight w:val="yellow"></w:highlight><w:sz w:val="24"/></w:rPr>';
    const fragment = XMLParser.extractSelfClosingTag(xml, 'w:highlight');

    expect(fragment).toBe(' w:val="yellow"');
    expect(XMLParser.extractAttribute(fragment ?? '', 'w:val')).toBe('yellow');
  });

  it('does not attribute a later sibling attribute to an expanded-form empty element', () => {
    const xml = '<w:highlight></w:highlight><w:sz w:val="24"/>';
    const fragment = XMLParser.extractSelfClosingTag(xml, 'w:highlight');

    expect(fragment).toBe('');
    expect(XMLParser.extractAttribute(fragment ?? '', 'w:val')).toBeUndefined();
  });

  it('does not steal theme attributes from a following sibling', () => {
    const xml = '<w:color w:val="FF0000"></w:color><w:u w:val="single" w:themeColor="accent1"/>';
    const fragment = XMLParser.extractSelfClosingTag(xml, 'w:color');

    expect(XMLParser.extractAttribute(fragment ?? '', 'w:val')).toBe('FF0000');
    expect(XMLParser.extractAttribute(fragment ?? '', 'w:themeColor')).toBeUndefined();
  });

  it('still parses self-closing form correctly (control)', () => {
    const xml = '<w:sz w:val="36"/><w:color w:val="FF0000"/>';

    expect(XMLParser.extractSelfClosingTag(xml, 'w:sz')).toBe(' w:val="36"');
    expect(XMLParser.extractSelfClosingTag(xml, 'w:color')).toBe(' w:val="FF0000"');
  });

  it('handles attribute-less self-closing tags (control)', () => {
    expect(XMLParser.extractSelfClosingTag('<w:rPr><w:b/></w:rPr>', 'w:b')).toBe('');
  });

  it("does not end the tag header at a '/>' inside a quoted attribute value", () => {
    const xml = '<w:color w:val="a/>b"></w:color>';
    const fragment = XMLParser.extractSelfClosingTag(xml, 'w:color');

    expect(XMLParser.extractAttribute(fragment ?? '', 'w:val')).toBe('a/>b');
  });
});
