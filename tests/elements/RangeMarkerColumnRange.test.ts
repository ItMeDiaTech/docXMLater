/**
 * RangeMarker — w:colFirst / w:colLast column-range attributes.
 *
 * Per ECMA-376 Part 1 §17.13.5 (CT_BookmarkRange, from which
 * CT_MoveBookmark inherits), start markers for moves in a table carry
 * an optional column-range pair:
 *
 *   <xsd:complexType name="CT_BookmarkRange">
 *     <xsd:attribute name="colFirst" type="ST_DecimalNumber"/>
 *     <xsd:attribute name="colLast"  type="ST_DecimalNumber"/>
 *   </xsd:complexType>
 *
 * Word writes these when a tracked move begins or ends on a range of
 * table columns rather than on in-cell text. Only the two
 * CT_MoveBookmark-derived start markers carry them:
 *
 *   moveFromRangeStart / moveToRangeStart  → YES
 *   moveFromRangeEnd   / moveToRangeEnd    → NO (CT_MarkupRange)
 *   customXml* (any)                        → NO (CT_TrackChange / CT_Markup)
 *
 * Bug this suite guards against:
 *   - `RangeMarker` had no `colFirst` / `colLast` support at all. Any
 *     Word-authored document that recorded a column-scoped tracked
 *     move through these markers lost the column range on round-trip.
 *     `Bookmark` (the non-tracked-move sibling class) already supports
 *     these — so tracked moves were the asymmetric odd one out.
 */

import { RangeMarker } from '../../src/elements/RangeMarker';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('RangeMarker — w:colFirst / w:colLast emission', () => {
  it('emits w:colFirst and w:colLast on moveFromRangeStart when set', () => {
    const marker = new RangeMarker({
      id: 1,
      type: 'moveFromRangeStart',
      name: 'move_1',
      author: 'Alice',
      date: new Date('2024-01-01T00:00:00Z'),
      colFirst: 2,
      colLast: 5,
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toMatch(/w:colFirst="2"/);
    expect(xml).toMatch(/w:colLast="5"/);
  });

  it('emits w:colFirst and w:colLast on moveToRangeStart when set', () => {
    const marker = new RangeMarker({
      id: 2,
      type: 'moveToRangeStart',
      name: 'move_1',
      author: 'Alice',
      date: new Date('2024-01-01T00:00:00Z'),
      colFirst: 0,
      colLast: 3,
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    // Zero is a valid column index (first column); must emit it rather than drop.
    expect(xml).toMatch(/w:colFirst="0"/);
    expect(xml).toMatch(/w:colLast="3"/);
  });

  it('emits only the set attribute when only one of the pair is provided', () => {
    // Per schema, both are independently optional. A pipeline that only
    // knows colFirst (say, because it's constructing a half-open range)
    // should be able to emit just that.
    const marker = new RangeMarker({
      id: 3,
      type: 'moveFromRangeStart',
      name: 'move_1',
      author: 'Alice',
      colFirst: 1,
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toMatch(/w:colFirst="1"/);
    expect(xml).not.toMatch(/w:colLast=/);
  });
});

describe('RangeMarker — w:colFirst / w:colLast scoped to CT_MoveBookmark types', () => {
  // On every other marker type the schema has no colFirst/colLast, so even
  // if the model carries values toXML must drop them to keep output valid.
  it('does NOT emit w:colFirst/w:colLast on moveFromRangeEnd', () => {
    const marker = new RangeMarker({
      id: 4,
      type: 'moveFromRangeEnd',
      colFirst: 1,
      colLast: 3,
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:colFirst=/);
    expect(xml).not.toMatch(/w:colLast=/);
  });

  it('does NOT emit w:colFirst/w:colLast on moveToRangeEnd', () => {
    const marker = new RangeMarker({
      id: 5,
      type: 'moveToRangeEnd',
      colFirst: 1,
      colLast: 3,
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:colFirst=/);
    expect(xml).not.toMatch(/w:colLast=/);
  });

  it('does NOT emit w:colFirst/w:colLast on customXmlMoveFromRangeStart', () => {
    const marker = new RangeMarker({
      id: 6,
      type: 'customXmlMoveFromRangeStart',
      author: 'Bob',
      colFirst: 1,
      colLast: 3,
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:colFirst=/);
    expect(xml).not.toMatch(/w:colLast=/);
  });

  it('does NOT emit w:colFirst/w:colLast on customXmlInsRangeEnd', () => {
    const marker = new RangeMarker({
      id: 7,
      type: 'customXmlInsRangeEnd',
      colFirst: 1,
      colLast: 3,
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:colFirst=/);
    expect(xml).not.toMatch(/w:colLast=/);
  });
});

describe('RangeMarker — w:colFirst / w:colLast absent when undefined', () => {
  it('omits both attributes entirely when neither is set', () => {
    const marker = RangeMarker.createMoveFromStart(10, 'move_1', 'Alice');
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:colFirst=/);
    expect(xml).not.toMatch(/w:colLast=/);
  });

  it('getColFirst() / getColLast() return undefined by default', () => {
    const marker = RangeMarker.createMoveFromStart(11, 'move_1', 'Alice');
    expect(marker.getColFirst()).toBeUndefined();
    expect(marker.getColLast()).toBeUndefined();
  });

  it('getColFirst() / getColLast() return the set values', () => {
    const marker = new RangeMarker({
      id: 12,
      type: 'moveFromRangeStart',
      name: 'move_1',
      author: 'Alice',
      colFirst: 4,
      colLast: 7,
    });
    expect(marker.getColFirst()).toBe(4);
    expect(marker.getColLast()).toBe(7);
  });
});
