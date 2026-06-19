/**
 * Mixed-content text must survive regardless of attribute presence.
 *
 * When an element had text plus element children but no attributes, the
 * element value was first set to the bare text string; the children-merge
 * block then took the non-object branch and overwrote it, silently dropping
 * the character data. With attributes present the same text survived under
 * the text-node key — an asymmetry, not a format limitation.
 */

import { XMLParser } from '../../src/xml/XMLParser';

describe('XMLParser.parseToObject — mixed content without attributes', () => {
  it('keeps text under the text-node key when an attribute-less element has children', () => {
    const result: any = XMLParser.parseToObject('<w:p>orphan text<w:r><w:t>x</w:t></w:r></w:p>', {
      trimValues: false,
    });

    expect(result['w:p']['#text']).toBe('orphan text');
    expect(result['w:p']['w:r']['w:t']).toBe('x');
  });

  it('matches the with-attributes form (asymmetry removed)', () => {
    const withAttr: any = XMLParser.parseToObject('<note id="1">lead <b>x</b></note>');
    const withoutAttr: any = XMLParser.parseToObject('<note>lead <b>x</b></note>');

    expect(withAttr.note['#text']).toBe('lead');
    expect(withoutAttr.note['#text']).toBe('lead');
    expect(withoutAttr.note.b).toBe('x');
  });

  it('honors a custom textNodeName for the folded text', () => {
    const result: any = XMLParser.parseToObject('<note>lead <b>x</b></note>', {
      textNodeName: '$value',
    });

    expect(result.note['$value']).toBe('lead');
  });
});
