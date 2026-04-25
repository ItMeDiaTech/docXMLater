/**
 * `<w:footnoteReference>` / `<w:endnoteReference>` — w:customMarkFollows
 * attribute round-trip.
 *
 * Per ECMA-376 Part 1 §17.11.13 (w:footnoteReference) and §17.11.2
 * (w:endnoteReference), both elements carry:
 *
 *   w:id (required)          — references the footnote/endnote definition
 *   w:customMarkFollows      — ST_OnOff. When true, tells Word "a custom
 *                              glyph (rendered as a normal run that
 *                              immediately follows this reference) is
 *                              the display mark; do NOT also emit the
 *                              auto-numbered footnote/endnote mark."
 *
 * Bug this suite guards against:
 *   - Parser reads only `w:id`. Generator emits only `w:id`. Any
 *     document that carried `<w:footnoteReference w:id="1"
 *     w:customMarkFollows="1"/><w:r><w:t>†</w:t></w:r>` would on
 *     round-trip drop the `customMarkFollows` flag and Word would
 *     then render BOTH the automatic number "1" AND the "†" glyph
 *     (doubled mark). That's a visible, incorrect rendering change.
 *
 * The attribute is ST_OnOff, so parse must honour every literal
 * ("1"/"0"/"true"/"false"/"on"/"off") per the cross-iteration pattern.
 */

import { Run, type RunContent } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { XMLParser } from '../../src/xml/XMLParser';

// Convenience: instantiate a Run and push a raw RunContent entry via a
// narrow cast — there's no public add method for footnote references, so
// tests route through the internal `content` array directly.
function pushContent(run: Run, entry: RunContent): void {
  (run as unknown as { content: RunContent[] }).content.push(entry);
}

describe('Run generator — w:customMarkFollows emission', () => {
  it('emits w:customMarkFollows="1" on footnoteReference when set', () => {
    const run = new Run('');
    pushContent(run, {
      type: 'footnoteReference',
      footnoteId: 1,
      customMarkFollows: true,
    });
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toMatch(/<w:footnoteReference[^>]*w:id="1"/);
    expect(xml).toMatch(/<w:footnoteReference[^>]*w:customMarkFollows="1"/);
  });

  it('omits w:customMarkFollows on footnoteReference when false', () => {
    const run = new Run('');
    pushContent(run, {
      type: 'footnoteReference',
      footnoteId: 1,
      customMarkFollows: false,
    });
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toMatch(/<w:footnoteReference[^>]*w:id="1"/);
    expect(xml).not.toMatch(/w:customMarkFollows=/);
  });

  it('omits w:customMarkFollows on footnoteReference when undefined', () => {
    const run = new Run('');
    pushContent(run, { type: 'footnoteReference', footnoteId: 1 });
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).not.toMatch(/w:customMarkFollows=/);
  });

  it('emits w:customMarkFollows="1" on endnoteReference when set', () => {
    const run = new Run('');
    pushContent(run, {
      type: 'endnoteReference',
      endnoteId: 2,
      customMarkFollows: true,
    });
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toMatch(/<w:endnoteReference[^>]*w:id="2"/);
    expect(xml).toMatch(/<w:endnoteReference[^>]*w:customMarkFollows="1"/);
  });

  it('omits w:customMarkFollows on endnoteReference when undefined', () => {
    const run = new Run('');
    pushContent(run, { type: 'endnoteReference', endnoteId: 2 });
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).not.toMatch(/w:customMarkFollows=/);
  });
});

// Parser side: verify parseRunFromObject preserves the attribute from source.
// We call into parseRunFromObject via the DocumentParser to avoid reconstructing
// the full run-level parse context.
import { DocumentParser } from '../../src/core/DocumentParser';

describe('Parser — w:customMarkFollows extraction', () => {
  it('parses w:customMarkFollows="1" on footnoteReference', () => {
    const runXml =
      '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnoteReference w:id="1" w:customMarkFollows="1"/></w:r>';
    const runObj = XMLParser.parseToObject(runXml) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = (new DocumentParser() as any).parseRunFromObject(runObj['w:r']);
    const content = (parsed as Run).getContent();
    const fnRef = content.find((c) => c.type === 'footnoteReference');
    expect(fnRef?.footnoteId).toBe(1);
    expect(fnRef?.customMarkFollows).toBe(true);
  });

  it('parses w:customMarkFollows="0" on endnoteReference as false', () => {
    const runXml =
      '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnoteReference w:id="2" w:customMarkFollows="0"/></w:r>';
    const runObj = XMLParser.parseToObject(runXml) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = (new DocumentParser() as any).parseRunFromObject(runObj['w:r']);
    const content = (parsed as Run).getContent();
    const enRef = content.find((c) => c.type === 'endnoteReference');
    expect(enRef?.endnoteId).toBe(2);
    expect(enRef?.customMarkFollows).toBe(false);
  });

  it('parses w:customMarkFollows="true" as true', () => {
    const runXml =
      '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnoteReference w:id="3" w:customMarkFollows="true"/></w:r>';
    const runObj = XMLParser.parseToObject(runXml) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = (new DocumentParser() as any).parseRunFromObject(runObj['w:r']);
    const content = (parsed as Run).getContent();
    const fnRef = content.find((c) => c.type === 'footnoteReference');
    expect(fnRef?.customMarkFollows).toBe(true);
  });

  it('leaves customMarkFollows undefined when attribute absent', () => {
    const runXml =
      '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnoteReference w:id="1"/></w:r>';
    const runObj = XMLParser.parseToObject(runXml) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = (new DocumentParser() as any).parseRunFromObject(runObj['w:r']);
    const content = (parsed as Run).getContent();
    const fnRef = content.find((c) => c.type === 'footnoteReference');
    expect(fnRef?.customMarkFollows).toBeUndefined();
  });
});
