/**
 * CT_Sym — required attribute handling (§17.3.3.30).
 *
 * Per ECMA-376 Part 1 §17.3.3.30 CT_Sym declares BOTH
 *   w:font (ST_String)
 *   w:char (ST_ShortHexNumber)
 * as REQUIRED attributes. Omitting either produces a `<w:sym/>` element
 * that fails strict OOXML schema validation.
 *
 * The Run generator emitted each attribute only when the corresponding
 * `RunContent` field was truthy — a symbol constructed without both
 * would emit malformed XML (e.g. `<w:sym w:font="Wingdings"/>` with no
 * char, or an empty `<w:sym/>`).
 */

import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('CT_Sym — required attributes (§17.3.3.30)', () => {
  it('emits both attributes when both are provided', () => {
    const run = Run.createFromContent([
      { type: 'symbol', symbolFont: 'Wingdings', symbolChar: 'F0A0' },
    ]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('w:font="Wingdings"');
    expect(xml).toContain('w:char="F0A0"');
  });

  it('skips the element entirely when w:font is missing (required by schema)', () => {
    const run = Run.createFromContent([{ type: 'symbol', symbolChar: 'F0A0' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    // Schema-required attribute is missing — skip emission to avoid
    // producing invalid OOXML. Consumers should always set both fields.
    expect(xml).not.toContain('<w:sym ');
    expect(xml).not.toContain('<w:sym/>');
  });

  it('skips the element entirely when w:char is missing (required by schema)', () => {
    const run = Run.createFromContent([{ type: 'symbol', symbolFont: 'Wingdings' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).not.toContain('<w:sym ');
    expect(xml).not.toContain('<w:sym/>');
  });

  it('skips the element entirely when both attributes are missing', () => {
    const run = Run.createFromContent([{ type: 'symbol' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).not.toContain('<w:sym ');
    expect(xml).not.toContain('<w:sym/>');
  });
});
