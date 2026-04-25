/**
 * RangeMarker — w:displacedByCustomXml attribute support.
 *
 * Per ECMA-376 Part 1 §17.13.5, the CT_MarkupRange base type (and
 * CT_MoveBookmark which inherits from it) carries an optional
 * `w:displacedByCustomXml` attribute of type `ST_DisplacedByCustomXml`
 * (enum: "next" | "prev"). Word emits this when a range-marker had to
 * be displaced because a custom-XML node boundary fell at the same
 * position; it tells a downstream consumer which side of the custom-XML
 * the marker semantically belongs to.
 *
 * Four range markers are allowed to carry this attribute per spec:
 *
 *   moveFromRangeStart / moveToRangeStart  (CT_MoveBookmark → CT_MarkupRange)
 *   moveFromRangeEnd   / moveToRangeEnd    (CT_MarkupRange)
 *
 * The four customXml* start markers (CT_TrackChange) and the four
 * customXml* end markers (CT_Markup) do NOT have it.
 *
 * Bug this suite guards against:
 *   - `RangeMarker.toXML` had zero support for `w:displacedByCustomXml`.
 *     The attribute was neither settable on the model nor emitted on
 *     serialize, so any pipeline that needed to preserve this
 *     tracked-change disambiguator (Word-authored docs with
 *     custom-XML-displaced move boundaries, typically) couldn't.
 */

import { RangeMarker } from '../../src/elements/RangeMarker';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('RangeMarker — w:displacedByCustomXml on CT_MarkupRange-derived types', () => {
  it('emits w:displacedByCustomXml="next" on moveFromRangeStart when set', () => {
    const marker = new RangeMarker({
      id: 1,
      type: 'moveFromRangeStart',
      name: 'move_1',
      author: 'Alice',
      date: new Date('2024-01-01T00:00:00Z'),
      displacedByCustomXml: 'next',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toMatch(/w:displacedByCustomXml="next"/);
  });

  it('emits w:displacedByCustomXml="prev" on moveToRangeStart when set', () => {
    const marker = new RangeMarker({
      id: 2,
      type: 'moveToRangeStart',
      name: 'move_1',
      author: 'Alice',
      date: new Date('2024-01-01T00:00:00Z'),
      displacedByCustomXml: 'prev',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toMatch(/w:displacedByCustomXml="prev"/);
  });

  it('emits w:displacedByCustomXml on moveFromRangeEnd when set', () => {
    const marker = new RangeMarker({
      id: 3,
      type: 'moveFromRangeEnd',
      displacedByCustomXml: 'next',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toMatch(/w:displacedByCustomXml="next"/);
  });

  it('emits w:displacedByCustomXml on moveToRangeEnd when set', () => {
    const marker = new RangeMarker({
      id: 4,
      type: 'moveToRangeEnd',
      displacedByCustomXml: 'prev',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).toMatch(/w:displacedByCustomXml="prev"/);
  });
});

describe('RangeMarker — w:displacedByCustomXml NOT emitted on types lacking it', () => {
  // CT_TrackChange (customXml*RangeStart) and CT_Markup (customXml*RangeEnd)
  // have no displacedByCustomXml attribute. Even if the model carries a value,
  // toXML must drop it there to keep the output schema-valid.
  it('does NOT emit w:displacedByCustomXml on customXmlMoveFromRangeStart', () => {
    const marker = new RangeMarker({
      id: 5,
      type: 'customXmlMoveFromRangeStart',
      author: 'Bob',
      displacedByCustomXml: 'next',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:displacedByCustomXml=/);
  });

  it('does NOT emit w:displacedByCustomXml on customXmlInsRangeStart', () => {
    const marker = new RangeMarker({
      id: 6,
      type: 'customXmlInsRangeStart',
      author: 'Carol',
      displacedByCustomXml: 'prev',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:displacedByCustomXml=/);
  });

  it('does NOT emit w:displacedByCustomXml on customXmlDelRangeEnd', () => {
    const marker = new RangeMarker({
      id: 7,
      type: 'customXmlDelRangeEnd',
      displacedByCustomXml: 'next',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:displacedByCustomXml=/);
  });

  it('does NOT emit w:displacedByCustomXml on customXmlMoveToRangeEnd', () => {
    const marker = new RangeMarker({
      id: 8,
      type: 'customXmlMoveToRangeEnd',
      displacedByCustomXml: 'prev',
    });
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:displacedByCustomXml=/);
  });
});

describe('RangeMarker — w:displacedByCustomXml absent when undefined', () => {
  it('omits the attribute entirely when displacedByCustomXml is undefined', () => {
    const marker = RangeMarker.createMoveFromStart(10, 'move_1', 'Alice');
    const xml = XMLBuilder.elementToString(marker.toXML());
    expect(xml).not.toMatch(/w:displacedByCustomXml=/);
  });

  it('exposes getDisplacedByCustomXml() returning undefined by default', () => {
    const marker = RangeMarker.createMoveFromStart(11, 'move_1', 'Alice');
    expect(marker.getDisplacedByCustomXml()).toBeUndefined();
  });

  it('exposes getDisplacedByCustomXml() returning the set value', () => {
    const marker = new RangeMarker({
      id: 12,
      type: 'moveFromRangeStart',
      name: 'move_1',
      author: 'Alice',
      displacedByCustomXml: 'next',
    });
    expect(marker.getDisplacedByCustomXml()).toBe('next');
  });
});
