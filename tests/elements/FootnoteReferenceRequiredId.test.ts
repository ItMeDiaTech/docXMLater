/**
 * CT_FtnEdnRef — required `w:id` attribute handling.
 *
 * Per ECMA-376 Part 1 §17.11.12 CT_FtnEdnRef (the schema backing both
 * `<w:footnoteReference>` §17.11.13 and `<w:endnoteReference>` §17.11.2),
 * the `w:id` attribute is REQUIRED. Omitting it produces a reference
 * that fails strict OOXML schema validation.
 *
 * The Run generator emitted `w:id` only when the corresponding
 * `footnoteId` / `endnoteId` field was defined — a footnote/endnote
 * reference constructed without an id would emit malformed XML
 * (`<w:footnoteReference/>` with no attributes).
 *
 * Correct behavior: skip emission entirely if the id is missing —
 * matches the iter-59 `<w:sym/>` fix for the same class of required-
 * attribute bugs.
 */

import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('CT_FtnEdnRef — required w:id (§17.11.12)', () => {
  it('emits <w:footnoteReference> with the id when provided', () => {
    const run = Run.createFromContent([{ type: 'footnoteReference', footnoteId: 3 }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('<w:footnoteReference w:id="3"/>');
  });

  it('emits <w:endnoteReference> with the id when provided', () => {
    const run = Run.createFromContent([{ type: 'endnoteReference', endnoteId: 7 }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('<w:endnoteReference w:id="7"/>');
  });

  it('preserves w:customMarkFollows alongside the id', () => {
    const run = Run.createFromContent([
      { type: 'footnoteReference', footnoteId: 3, customMarkFollows: true },
    ]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('w:id="3"');
    expect(xml).toContain('w:customMarkFollows="1"');
  });

  it('skips <w:footnoteReference> entirely when the id is missing (schema-required)', () => {
    const run = Run.createFromContent([{ type: 'footnoteReference' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).not.toContain('<w:footnoteReference');
  });

  it('skips <w:endnoteReference> entirely when the id is missing (schema-required)', () => {
    const run = Run.createFromContent([{ type: 'endnoteReference' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).not.toContain('<w:endnoteReference');
  });

  it('accepts id=0 as a valid explicit value (not treated as "missing")', () => {
    const run = Run.createFromContent([{ type: 'footnoteReference', footnoteId: 0 }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('<w:footnoteReference w:id="0"/>');
  });
});
