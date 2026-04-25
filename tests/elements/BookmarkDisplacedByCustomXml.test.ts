/**
 * Bookmark — w:displacedByCustomXml round-trip tests.
 *
 * Per ECMA-376 Part 1 §17.16.5, a `<w:bookmarkStart>` is CT_Bookmark,
 * and CT_Bookmark extends CT_BookmarkRange which extends CT_MarkupRange:
 *
 *   <xsd:complexType name="CT_MarkupRange">
 *     <xsd:attribute name="id"/>
 *     <xsd:attribute name="displacedByCustomXml" type="ST_DisplacedByCustomXml"/>
 *   </xsd:complexType>
 *
 * `<w:bookmarkEnd>` is just CT_MarkupRange directly — same base, same
 * permitted attributes.
 *
 * Bugs this suite guards against:
 *   - `Bookmark` has no `displacedByCustomXml` field at all. The
 *     attribute is silently dropped on parse and cannot be emitted on
 *     serialize, even though the sibling class `RangeMarker` just
 *     gained support for exactly this attribute via iterations 14–15.
 *     Any Word-authored document with custom-XML-displaced bookmarks
 *     loses the disambiguator on any round-trip through docxmlater.
 *   - `DocumentParser.parseBookmark` reads `w:colFirst` / `w:colLast`
 *     but not `w:displacedByCustomXml`, so the parse side drops the
 *     attribute before the model ever sees it.
 */

import { Bookmark } from '../../src/elements/Bookmark';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('Bookmark — w:displacedByCustomXml emission on toStartXML', () => {
  it('emits w:displacedByCustomXml="next" when set', () => {
    const bookmark = new Bookmark({
      id: 1,
      name: 'test_bookmark',
      displacedByCustomXml: 'next',
    });
    const xml = XMLBuilder.elementToString(bookmark.toStartXML());
    expect(xml).toContain('<w:bookmarkStart');
    expect(xml).toMatch(/w:displacedByCustomXml="next"/);
  });

  it('emits w:displacedByCustomXml="prev" when set', () => {
    const bookmark = new Bookmark({
      id: 2,
      name: 'test_bookmark',
      displacedByCustomXml: 'prev',
    });
    const xml = XMLBuilder.elementToString(bookmark.toStartXML());
    expect(xml).toMatch(/w:displacedByCustomXml="prev"/);
  });

  it('omits w:displacedByCustomXml when undefined', () => {
    const bookmark = new Bookmark({
      id: 3,
      name: 'test_bookmark',
    });
    const xml = XMLBuilder.elementToString(bookmark.toStartXML());
    expect(xml).not.toMatch(/w:displacedByCustomXml=/);
  });

  it('emits w:displacedByCustomXml alongside w:colFirst/w:colLast', () => {
    const bookmark = new Bookmark({
      id: 4,
      name: 'test_bookmark',
      colFirst: 0,
      colLast: 3,
      displacedByCustomXml: 'next',
    });
    const xml = XMLBuilder.elementToString(bookmark.toStartXML());
    expect(xml).toMatch(/w:colFirst="0"/);
    expect(xml).toMatch(/w:colLast="3"/);
    expect(xml).toMatch(/w:displacedByCustomXml="next"/);
  });
});

describe('Bookmark — w:displacedByCustomXml emission on toEndXML', () => {
  it('emits w:displacedByCustomXml on bookmarkEnd when set', () => {
    // bookmarkEnd is also CT_MarkupRange per ECMA-376 §17.13.6.1, so it too
    // permits displacedByCustomXml (w:colFirst/w:colLast are NOT valid on end,
    // since end inherits directly from MarkupRange without the BookmarkRange layer).
    const bookmark = new Bookmark({
      id: 5,
      name: 'test_bookmark',
      displacedByCustomXml: 'prev',
    });
    const xml = XMLBuilder.elementToString(bookmark.toEndXML());
    expect(xml).toContain('<w:bookmarkEnd');
    expect(xml).toMatch(/w:displacedByCustomXml="prev"/);
  });

  it('omits w:displacedByCustomXml on bookmarkEnd when undefined', () => {
    const bookmark = new Bookmark({ id: 6, name: 'test_bookmark' });
    const xml = XMLBuilder.elementToString(bookmark.toEndXML());
    expect(xml).not.toMatch(/w:displacedByCustomXml=/);
  });
});

describe('Bookmark — displacedByCustomXml getter/setter', () => {
  it('getDisplacedByCustomXml() returns undefined by default', () => {
    const bookmark = new Bookmark({ id: 7, name: 'test_bookmark' });
    expect(bookmark.getDisplacedByCustomXml()).toBeUndefined();
  });

  it('getDisplacedByCustomXml() returns the value set in constructor', () => {
    const bookmark = new Bookmark({
      id: 8,
      name: 'test_bookmark',
      displacedByCustomXml: 'next',
    });
    expect(bookmark.getDisplacedByCustomXml()).toBe('next');
  });
});
