/**
 * CT_PTab — required attribute defaults.
 *
 * Per ECMA-376 Part 1 §17.3.3.23 CT_PTab has three REQUIRED attributes:
 *   w:alignment  (ST_PTabAlignment)
 *   w:relativeTo (ST_PTabRelativeTo)
 *   w:leader     (ST_PTabLeader)
 *
 * The RunContent → XML generator emitted each attribute only when the
 * corresponding field was truthy. If a consumer constructed a
 * `positionTab` RunContent without setting all three (or parsed
 * `<w:ptab/>` without them — technically malformed input), the output
 * XML would omit one or more required attributes, failing strict OOXML
 * schema validation.
 */

import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('CT_PTab — required-attribute defaults (§17.3.3.23)', () => {
  it('emits all three required attributes when only alignment is set', () => {
    const run = Run.createFromContent([{ type: 'positionTab', ptabAlignment: 'center' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('w:alignment="center"');
    // Spec defaults: relativeTo → margin, leader → none (both REQUIRED
    // attributes; absence fails schema validation).
    expect(xml).toMatch(/w:relativeTo="[^"]+"/);
    expect(xml).toMatch(/w:leader="[^"]+"/);
  });

  it('emits default w:relativeTo="margin" when not specified', () => {
    const run = Run.createFromContent([{ type: 'positionTab', ptabAlignment: 'left' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('w:relativeTo="margin"');
  });

  it('emits default w:leader="none" when not specified', () => {
    const run = Run.createFromContent([{ type: 'positionTab', ptabAlignment: 'right' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('w:leader="none"');
  });

  it('emits default w:alignment="left" when not specified (edge case — technically malformed input)', () => {
    const run = Run.createFromContent([{ type: 'positionTab' }]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('w:alignment="left"');
    expect(xml).toContain('w:relativeTo="margin"');
    expect(xml).toContain('w:leader="none"');
  });

  it('preserves explicit values when provided', () => {
    const run = Run.createFromContent([
      {
        type: 'positionTab',
        ptabAlignment: 'right',
        ptabRelativeTo: 'indent',
        ptabLeader: 'dot',
      },
    ]);
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('w:alignment="right"');
    expect(xml).toContain('w:relativeTo="indent"');
    expect(xml).toContain('w:leader="dot"');
  });
});
