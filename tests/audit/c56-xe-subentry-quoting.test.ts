/**
 * XE field subentry quoting (ECMA-376 §17.16.5.75).
 *
 * Word's XE field syntax requires the main-entry/subentry colon separator
 * inside the quoted field-argument: { XE "Main entry:Subentry" }. Emitting
 * the subentry after the closing quote ('XE "Main":Sub') leaves stray
 * unquoted text, so Word indexes only the main entry and drops or
 * misparses the subentry.
 */

import { Field } from '../../src/elements/Field';

describe('Field.createXEEntry subentry quoting (ECMA-376 §17.16.5.75)', () => {
  it('places the subentry colon separator inside the quoted argument', () => {
    const field = Field.createXEEntry('Main Entry', 'Sub Entry');

    expect(field.getInstruction()).toBe('XE "Main Entry:Sub Entry"');
  });

  it('does not leave stray text after the closing quote', () => {
    const field = Field.createXEEntry('Main Entry', 'Sub Entry');

    expect(field.getInstruction()).not.toMatch(/":/);
    expect(field.getInstruction().endsWith('"')).toBe(true);
  });

  it('keeps the plain single-entry form unchanged', () => {
    const field = Field.createXEEntry('Index Term');

    expect(field.getInstruction()).toBe('XE "Index Term"');
  });

  it('emits the quoted form into the fldSimple instruction attribute', () => {
    const field = Field.createXEEntry('Main', 'Sub');
    const xml = field.toXML();

    expect(xml.name).toBe('w:fldSimple');
    expect(xml.attributes!['w:instr']).toBe('XE "Main:Sub"');
  });
});
