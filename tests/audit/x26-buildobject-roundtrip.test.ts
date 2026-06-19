/**
 * XMLBuilder.buildObject is documented as the reverse of
 * XMLParser.parseToObject, so feeding it real parser output must reproduce
 * the original markup. That requires three behaviors: primitive child values
 * (text-only elements collapsed by the parser) are re-wrapped in their tag
 * instead of leaking bare text into the parent; the parser's
 * _orderedChildren metadata key is never serialized as literal elements; and
 * that metadata drives emission order so interleaved same-named siblings
 * (w:t / w:tab / w:t) keep their document order.
 */
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { XMLParser } from '../../src/xml/XMLParser';

describe('XMLBuilder.buildObject round-trip with parseToObject output', () => {
  it('round-trips mixed run content preserving tags and interleaved order', () => {
    const xml = '<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r></w:p>';
    const parsed = XMLParser.parseToObject(xml) as any;

    const rebuilt = XMLBuilder.buildObject(parsed['w:p'], 'w:p');

    expect(rebuilt).toBe(xml);
  });

  it('does not emit _orderedChildren metadata as literal elements', () => {
    const parsed = XMLParser.parseToObject('<w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r>') as any;

    const rebuilt = XMLBuilder.buildObject(parsed['w:r'], 'w:r');

    expect(rebuilt).not.toContain('_orderedChildren');
  });

  it('wraps primitive text values in their element tag', () => {
    // The parser collapses <w:t>A</w:t> to the string 'A'; rebuilding must
    // restore the tag — bare text directly inside w:r is invalid OOXML.
    const parsed = XMLParser.parseToObject('<w:r><w:t>Hello</w:t></w:r>') as any;

    const rebuilt = XMLBuilder.buildObject(parsed['w:r'], 'w:r');

    expect(rebuilt).toBe('<w:r><w:t>Hello</w:t></w:r>');
  });

  it('wraps #text-only objects in the root tag', () => {
    const rebuilt = XMLBuilder.buildObject({ '#text': 'Hello' }, 'w:t');

    expect(rebuilt).toBe('<w:t>Hello</w:t>');
  });

  it('round-trips attributes alongside text and ordered properties', () => {
    const xml =
      '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>' +
      '<w:t xml:space="preserve">Hello </w:t></w:r>';
    const parsed = XMLParser.parseToObject(xml, { trimValues: false }) as any;

    const rebuilt = XMLBuilder.buildObject(parsed['w:r'], 'w:r');

    expect(rebuilt).toBe(xml);
  });

  it('preserves attributes on interleaved same-name siblings in document order', () => {
    // Two <w:t> siblings, one carrying xml:space, interleaved with <w:tab/>:
    // ordered-children reconstruction must keep order AND the per-element
    // attribute (a regression here would drop xml:space or reorder the runs).
    const xml = '<w:r><w:t xml:space="preserve">A </w:t><w:tab/><w:t>B</w:t></w:r>';
    const parsed = XMLParser.parseToObject(xml, { trimValues: false }) as any;

    const rebuilt = XMLBuilder.buildObject(parsed['w:r'], 'w:r');

    expect(rebuilt).toBe(xml);
  });

  it('round-trips namespace-prefixed elements interleaved with standard content', () => {
    const xml =
      '<w:p><w:r><w:t>x</w:t></w:r><w14:something w14:val="1"/><w:r><w:t>y</w:t></w:r></w:p>';
    const parsed = XMLParser.parseToObject(xml) as any;

    const rebuilt = XMLBuilder.buildObject(parsed['w:p'], 'w:p');

    expect(rebuilt).toBe(xml);
  });
});
