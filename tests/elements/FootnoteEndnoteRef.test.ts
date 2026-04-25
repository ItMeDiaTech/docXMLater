/**
 * `<w:footnoteRef/>` and `<w:endnoteRef/>` — auto-numbered mark
 * elements inside footnote/endnote definition bodies.
 *
 * These are **distinct** from the body-side `<w:footnoteReference>` /
 * `<w:endnoteReference>` (which are references in the main document
 * body that POINT TO a footnote/endnote definition). The -Ref elements
 * appear INSIDE the footnote/endnote content to render the
 * auto-generated number (1, 2, 3...) as the displayed mark.
 *
 *   ECMA-376 Part 1 §17.11.14 (`<w:footnoteRef>`)
 *   ECMA-376 Part 1 §17.11.3  (`<w:endnoteRef>`)
 *
 * Both are empty self-closing elements with NO attributes.
 *
 * Canonical shape of a footnote body (what Word writes):
 *   <w:footnote w:id="1">
 *     <w:p>
 *       <w:r>
 *         <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
 *         <w:footnoteRef/>   <!-- renders as "1" -->
 *       </w:r>
 *       <w:r><w:t> Footnote text here.</w:t></w:r>
 *     </w:p>
 *   </w:footnote>
 *
 * Bug this suite guards against:
 *   - The RunContentType union had no `footnoteRef` / `endnoteRef`
 *     members. Neither parser nor generator handled either element,
 *     so any footnote/endnote definition that used the canonical
 *     Word shape lost the auto-numbered mark on round-trip — and
 *     once lost, Word has no explicit mark element to render, so the
 *     displayed footnote number could disappear or render at a
 *     different position than the source document specified.
 */

import { Run, type RunContent } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { XMLParser } from '../../src/xml/XMLParser';
import { DocumentParser } from '../../src/core/DocumentParser';

function pushContent(run: Run, entry: RunContent): void {
  (run as unknown as { content: RunContent[] }).content.push(entry);
}

describe('Run generator — emits <w:footnoteRef/> and <w:endnoteRef/>', () => {
  it('emits <w:footnoteRef/> when content entry is footnoteRef', () => {
    const run = new Run('');
    pushContent(run, { type: 'footnoteRef' });
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('<w:footnoteRef/>');
  });

  it('emits <w:endnoteRef/> when content entry is endnoteRef', () => {
    const run = new Run('');
    pushContent(run, { type: 'endnoteRef' });
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toContain('<w:endnoteRef/>');
  });

  it('preserves ordering of footnoteRef alongside text in the same run', () => {
    const run = new Run('');
    pushContent(run, { type: 'footnoteRef' });
    pushContent(run, { type: 'text', value: ' Footnote text.' });
    const xml = XMLBuilder.elementToString(run.toXML());
    const refIdx = xml.indexOf('<w:footnoteRef/>');
    const textIdx = xml.indexOf('Footnote text.');
    expect(refIdx).toBeGreaterThan(-1);
    expect(textIdx).toBeGreaterThan(refIdx);
  });
});

describe('Parser — extracts <w:footnoteRef/> and <w:endnoteRef/>', () => {
  it('parses <w:footnoteRef/> inside a run', () => {
    const runXml =
      '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnoteRef/></w:r>';
    const runObj = XMLParser.parseToObject(runXml) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = (new DocumentParser() as any).parseRunFromObject(runObj['w:r']);
    const content = (parsed as Run).getContent();
    expect(content.some((c) => c.type === 'footnoteRef')).toBe(true);
  });

  it('parses <w:endnoteRef/> inside a run', () => {
    const runXml =
      '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnoteRef/></w:r>';
    const runObj = XMLParser.parseToObject(runXml) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = (new DocumentParser() as any).parseRunFromObject(runObj['w:r']);
    const content = (parsed as Run).getContent();
    expect(content.some((c) => c.type === 'endnoteRef')).toBe(true);
  });

  it('parses a canonical footnote-body run (rStyle + footnoteRef)', () => {
    const runXml =
      '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>';
    const runObj = XMLParser.parseToObject(runXml) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = (new DocumentParser() as any).parseRunFromObject(runObj['w:r']);
    const content = (parsed as Run).getContent();
    expect(content.some((c) => c.type === 'footnoteRef')).toBe(true);
    // And the run carries the FootnoteReference character-style reference.
    expect((parsed as Run).getFormatting().characterStyle).toBe('FootnoteReference');
  });
});

describe('Full round-trip of <w:footnoteRef/>', () => {
  it('round-trips through parse → serialize → parse intact', () => {
    const runXml =
      '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnoteRef/></w:r>';
    const runObj = XMLParser.parseToObject(runXml) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parsed = (new DocumentParser() as any).parseRunFromObject(runObj['w:r']);
    const serialized = XMLBuilder.elementToString((parsed as Run).toXML());
    expect(serialized).toContain('<w:footnoteRef/>');
    // Re-parse what we serialized
    const re = XMLParser.parseToObject(serialized) as { 'w:r': Record<string, unknown> };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const reparsed = (new DocumentParser() as any).parseRunFromObject(re['w:r']);
    expect((reparsed as Run).getContent().some((c) => c.type === 'footnoteRef')).toBe(true);
  });
});
