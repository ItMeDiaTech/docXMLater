/**
 * `<w:lvl><w:pPr><w:ind w:start="…" …/>` — numbering level indent
 * parser must honour the bidi-aware `w:start` / `w:end` attribute
 * aliases, not just the legacy `w:left` / `w:right`.
 *
 * Per ECMA-376 Part 1 §17.3.1.12 CT_Ind, the indentation element
 * accepts both legacy LTR-centric attributes (w:left, w:right) and
 * the bidi-aware aliases (w:start, w:end). Modern authoring tools
 * (Word 2013+ defaults, Google Docs, most web-based editors) emit
 * the bidi-aware form by default.
 *
 * Bug: `NumberingLevel.fromXML` (the `<w:lvl>` parser) used
 *
 *     const leftMatch = /w:left="([^"]+)"/.exec(indElement);
 *
 * which matched only `w:left`. A numbering level whose indent was
 * authored with `<w:ind w:start="720" w:hanging="360"/>` silently
 * fell back to the default `720 + level * 360` / `360` values on
 * load — losing the author's indentation entirely for every level
 * in the list.
 *
 * Iteration 109 extends the regex to prefer `w:start` and fall back
 * to `w:left`, matching the parse precedence used on the main
 * paragraph-indent path.
 */

import { NumberingLevel } from '../../src/formatting/NumberingLevel';

describe('NumberingLevel <w:ind> bidi-aware w:start parse', () => {
  it('parses w:start on the numbering level indent', () => {
    const xml = `
      <w:lvl w:ilvl="0" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
          <w:ind w:start="1440" w:hanging="360"/>
        </w:pPr>
      </w:lvl>`;
    const level = NumberingLevel.fromXML(xml);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const props = (level as any).properties;
    expect(props.leftIndent).toBe(1440);
    expect(props.hangingIndent).toBe(360);
  });

  it('still parses w:left on the numbering level indent (regression guard)', () => {
    const xml = `
      <w:lvl w:ilvl="0" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
          <w:ind w:left="1080" w:hanging="360"/>
        </w:pPr>
      </w:lvl>`;
    const level = NumberingLevel.fromXML(xml);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const props = (level as any).properties;
    expect(props.leftIndent).toBe(1080);
    expect(props.hangingIndent).toBe(360);
  });

  it('prefers w:start over w:left when both are present (spec precedence)', () => {
    // Per ECMA-376 §17.3.1.12, `w:start` is the bidi-aware value and
    // takes precedence over the legacy `w:left` when both appear.
    const xml = `
      <w:lvl w:ilvl="0" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:start w:val="1"/>
        <w:numFmt w:val="decimal"/>
        <w:lvlText w:val="%1."/>
        <w:lvlJc w:val="left"/>
        <w:pPr>
          <w:ind w:start="2000" w:left="999" w:hanging="360"/>
        </w:pPr>
      </w:lvl>`;
    const level = NumberingLevel.fromXML(xml);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const props = (level as any).properties;
    expect(props.leftIndent).toBe(2000);
  });
});
