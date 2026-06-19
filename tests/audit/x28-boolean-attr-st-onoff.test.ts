/**
 * XMLBuilder.elementToString must serialize boolean attribute values as
 * ST_OnOff literals (ECMA-376 §22.9.2.13: 1/0/true/false/on/off). The
 * XMLElement attribute type allows booleans and parseToObject coerces
 * w:val="true"/"false" to JS booleans, so HTML-style minimization
 * (true -> attr="attr") emits an invalid literal, and dropping explicit
 * false inverts an explicit style-override off to on under CT_OnOff
 * presence semantics. Absence is signalled with undefined, never false.
 */
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { XMLParser } from '../../src/xml/XMLParser';

describe('elementToString boolean attribute serialization', () => {
  it('serializes true as the ST_OnOff literal "1", not the attribute name', () => {
    const xml = new XMLBuilder().selfClosingElement('w:b', { 'w:val': true }).build();
    expect(xml).toBe('<w:b w:val="1"/>');
  });

  it('serializes explicit false as "0" instead of dropping the attribute', () => {
    const xml = new XMLBuilder().selfClosingElement('w:i', { 'w:val': false }).build();
    expect(xml).toBe('<w:i w:val="0"/>');
  });

  it('still omits undefined and null attribute values', () => {
    const xml = new XMLBuilder()
      .selfClosingElement('w:b', {
        'w:val': undefined,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any -- null exercises the runtime guard
        'w:x': null as any,
      })
      .build();
    expect(xml).toBe('<w:b/>');
  });

  it('round-trips parser-coerced boolean attributes through buildObject as valid ST_OnOff', () => {
    const parsedTrue = XMLParser.parseToObject('<w:b w:val="true"/>') as any;
    expect(parsedTrue['w:b']['@_w:val']).toBe(true);
    expect(XMLBuilder.buildObject(parsedTrue['w:b'], 'w:b')).toBe('<w:b w:val="1"/>');

    const parsedFalse = XMLParser.parseToObject('<w:i w:val="false"/>') as any;
    expect(parsedFalse['w:i']['@_w:val']).toBe(false);
    // Dropping w:val here would mean TRUE per CT_OnOff presence semantics.
    expect(XMLBuilder.buildObject(parsedFalse['w:i'], 'w:i')).toBe('<w:i w:val="0"/>');
  });
});
