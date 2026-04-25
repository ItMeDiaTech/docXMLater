/**
 * RangeMarker w:name scoping tests.
 *
 * Per ECMA-376 Part 1 §17.13.5, the twelve range-marker elements the
 * codebase knows about fall into three distinct OOXML types:
 *
 *   moveFromRangeStart / moveToRangeStart        → CT_MoveBookmark
 *       attrs: w:id, w:displacedByCustomXml, w:colFirst, w:colLast,
 *              w:name (required), w:author (required), w:date
 *
 *   customXml*RangeStart (Ins/Del/MoveFrom/MoveTo) → CT_TrackChange
 *       attrs: w:id (required), w:author (required), w:date
 *       NOTE: NO w:name
 *
 *   moveFromRangeEnd / moveToRangeEnd            → CT_MarkupRange
 *       attrs: w:id, w:displacedByCustomXml
 *       NOTE: NO w:name
 *
 *   customXml*RangeEnd                           → CT_Markup
 *       attrs: w:id only
 *
 * Bug this suite guards against:
 *   - `RangeMarker.toXML` emitted `w:name` unconditionally whenever the
 *     instance carried a name, regardless of the marker's OOXML type.
 *     That wrote `w:name` onto CT_TrackChange and CT_MarkupRange/Markup
 *     elements where the attribute is not in the schema, producing
 *     invalid XML. The fault surfaces because `createCustomXmlMoveFromStart`
 *     and `createCustomXmlMoveToStart` accept a `name` parameter, so
 *     pipeline code calling the straightforward factory path generates
 *     documents that fail the Open XML SDK validator.
 */

import { RangeMarker } from '../../src/elements/RangeMarker';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('RangeMarker toXML — w:name is only valid on CT_MoveBookmark elements', () => {
  it('emits w:name on moveFromRangeStart (CT_MoveBookmark)', () => {
    const marker = RangeMarker.createMoveFromStart(1, 'move_1', 'Alice');
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:moveFromRangeStart');
    expect(xml).toMatch(/w:name="move_1"/);
  });

  it('emits w:name on moveToRangeStart (CT_MoveBookmark)', () => {
    const marker = RangeMarker.createMoveToStart(2, 'move_1', 'Alice');
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:moveToRangeStart');
    expect(xml).toMatch(/w:name="move_1"/);
  });

  it('does NOT emit w:name on moveFromRangeEnd (CT_MarkupRange)', () => {
    // If someone manually constructs an end marker with a name, it must
    // still be dropped on serialize because CT_MarkupRange has no w:name.
    const marker = new RangeMarker({
      id: 3,
      type: 'moveFromRangeEnd',
      name: 'move_1',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:moveFromRangeEnd');
    expect(xml).not.toMatch(/w:name=/);
  });

  it('does NOT emit w:name on moveToRangeEnd (CT_MarkupRange)', () => {
    const marker = new RangeMarker({
      id: 4,
      type: 'moveToRangeEnd',
      name: 'move_1',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:moveToRangeEnd');
    expect(xml).not.toMatch(/w:name=/);
  });

  it('does NOT emit w:name on customXmlMoveFromRangeStart (CT_TrackChange)', () => {
    // The factory accepts a `name` parameter, but CT_TrackChange has no
    // w:name attribute — so toXML must still omit it on serialize.
    const marker = RangeMarker.createCustomXmlMoveFromStart(5, 'cx_move_1', 'Bob');
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:customXmlMoveFromRangeStart');
    expect(xml).not.toMatch(/w:name=/);
    // But author/date should still be emitted (those ARE in CT_TrackChange)
    expect(xml).toMatch(/w:author="Bob"/);
    expect(xml).toMatch(/w:date=/);
  });

  it('does NOT emit w:name on customXmlMoveToRangeStart (CT_TrackChange)', () => {
    const marker = RangeMarker.createCustomXmlMoveToStart(6, 'cx_move_1', 'Bob');
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:customXmlMoveToRangeStart');
    expect(xml).not.toMatch(/w:name=/);
    expect(xml).toMatch(/w:author="Bob"/);
  });

  it('does NOT emit w:name on customXmlMoveFromRangeEnd (CT_Markup)', () => {
    const marker = new RangeMarker({
      id: 7,
      type: 'customXmlMoveFromRangeEnd',
      name: 'cx_move_1',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:customXmlMoveFromRangeEnd');
    expect(xml).not.toMatch(/w:name=/);
  });

  it('does NOT emit w:name on customXmlInsRangeStart (CT_TrackChange)', () => {
    const marker = RangeMarker.createCustomXmlInsStart(8, 'Carol');
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:customXmlInsRangeStart');
    expect(xml).not.toMatch(/w:name=/);
    expect(xml).toMatch(/w:author="Carol"/);
  });

  it('does NOT emit w:name on customXmlDelRangeStart (CT_TrackChange)', () => {
    const marker = RangeMarker.createCustomXmlDelStart(9, 'Carol');
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toContain('<w:customXmlDelRangeStart');
    expect(xml).not.toMatch(/w:name=/);
    expect(xml).toMatch(/w:author="Carol"/);
  });
});

describe('RangeMarker toXML — w:id always emitted', () => {
  it('emits w:id on every marker type', () => {
    const markers: RangeMarker[] = [
      RangeMarker.createMoveFromStart(1, 'n', 'A'),
      RangeMarker.createMoveFromEnd(1),
      RangeMarker.createMoveToStart(2, 'n', 'A'),
      RangeMarker.createMoveToEnd(2),
      RangeMarker.createCustomXmlInsStart(3, 'A'),
      RangeMarker.createCustomXmlInsEnd(3),
      RangeMarker.createCustomXmlDelStart(4, 'A'),
      RangeMarker.createCustomXmlDelEnd(4),
      RangeMarker.createCustomXmlMoveFromStart(5, 'n', 'A'),
      RangeMarker.createCustomXmlMoveFromEnd(5),
      RangeMarker.createCustomXmlMoveToStart(6, 'n', 'A'),
      RangeMarker.createCustomXmlMoveToEnd(6),
    ];
    for (const m of markers) {
      const xml = XMLBuilder.elementToString(m.toXML());
      expect(xml).toMatch(/w:id="\d+"/);
    }
  });
});
