/**
 * Built-in styles must specify w:hAnsi and w:cs explicitly (not rely on the
 * removed serialization back-fill from C44).
 *
 * C44 stopped the rFonts serializer from synthesizing w:hAnsi/w:cs from the
 * ascii font, which is correct for parsed/direct-formatting runs. The framework's
 * built-in Verdana styles (Normal, Heading*, Title, Subtitle, List Paragraph,
 * TOC Heading) are documented as rendering Verdana across all character ranges;
 * since docDefaults resolves the high-ANSI (w:hAnsi) slot to Calibri, these
 * styles must carry w:hAnsi/w:cs explicitly or high-ANSI characters (smart
 * quotes, dashes, accented Latin) would render in Calibri instead of Verdana.
 */

import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

function rFontsOf(style: Style): string {
  const xml = XMLBuilder.elementToString(style.toXML());
  return xml.match(/<w:rFonts[^>]*\/?>/)?.[0] ?? '';
}

describe('built-in Verdana styles carry all rFonts slots explicitly', () => {
  const cases: Array<[string, () => Style]> = [
    ['Normal', () => Style.createNormalStyle()],
    ['Heading1', () => Style.createHeadingStyle(1)],
    ['Heading1 Char', () => Style.createHeadingCharStyle(1)],
    ['Title', () => Style.createTitleStyle()],
    ['Subtitle', () => Style.createSubtitleStyle()],
    ['List Paragraph', () => Style.createListParagraphStyle()],
    ['TOC Heading', () => Style.createTOCHeadingStyle()],
  ];

  it.each(cases)('%s emits w:ascii, w:hAnsi and w:cs Verdana', (_name, factory) => {
    const rFonts = rFontsOf(factory());

    expect(rFonts).toContain('w:ascii="Verdana"');
    expect(rFonts).toContain('w:hAnsi="Verdana"');
    expect(rFonts).toContain('w:cs="Verdana"');
  });
});
