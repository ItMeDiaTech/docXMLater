/**
 * Field.createCustom field type.
 *
 * createCustom previously hardcoded type 'PAGE' as a placeholder, so any
 * custom field (e.g. STYLEREF) serialized with a cached fldSimple result
 * of '1' — wrong visible content until the user refreshes fields — and
 * getType() misreported the type. The FieldType union has 'CUSTOM',
 * whose placeholder text is the empty string.
 */

import { Field } from '../../src/elements/Field';
import { XMLElement } from '../../src/xml/XMLBuilder';

function getResultText(fldSimple: XMLElement): string[] | undefined {
  const run = fldSimple.children!.find(
    (child): child is XMLElement => typeof child !== 'string' && child.name === 'w:r'
  );
  const text = (run?.children || []).find(
    (child): child is XMLElement => typeof child !== 'string' && child.name === 'w:t'
  );
  return text?.children as string[] | undefined;
}

describe('Field.createCustom uses the CUSTOM field type', () => {
  it('reports CUSTOM from getType()', () => {
    const field = Field.createCustom('STYLEREF "Heading 1"');

    expect(field.getType()).toBe('CUSTOM');
  });

  it('keeps the provided instruction unchanged', () => {
    const field = Field.createCustom('STYLEREF "Heading 1"');

    expect(field.getInstruction()).toBe('STYLEREF "Heading 1"');
  });

  it('does not cache the PAGE placeholder "1" as the field result', () => {
    const field = Field.createCustom('STYLEREF "Heading 1"');
    const xml = field.toXML();

    expect(xml.name).toBe('w:fldSimple');
    expect(xml.attributes!['w:instr']).toBe('STYLEREF "Heading 1"');
    // CUSTOM placeholder is the empty string, not the PAGE placeholder '1'
    expect(getResultText(xml)).toEqual(['']);
  });
});
