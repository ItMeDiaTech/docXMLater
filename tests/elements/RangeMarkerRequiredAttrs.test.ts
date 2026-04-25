/**
 * Range marker emission — required-attribute compliance per ECMA-376.
 *
 * `<w:moveFromRangeStart>` / `<w:moveToRangeStart>` are CT_MoveBookmark
 * (§17.13.5.24) which extends CT_BookmarkRange → CT_MarkupRange → CT_Markup.
 * REQUIRED attributes on the start markers:
 *   - w:id     (from CT_Markup)
 *   - w:author (from CT_MoveBookmark's extension of CT_TrackChange ancestry)
 *   - w:date   (from CT_MoveBookmark)
 *   - w:name   (from CT_MoveBookmark)
 *
 * `<w:customXmlMoveFromRangeStart>` / similar Ins/Del variants are
 * CT_TrackChangeRange (§17.13.5.5) extending CT_TrackChange — REQUIRE:
 *   - w:id
 *   - w:author
 *   - w:date
 *
 * End markers (RangeEnd variants) are CT_MarkupRange — require only w:id.
 *
 * The previous `RangeMarker.toXML` used truthy-gated `if (this.author)` /
 * `if (this.date)` / `if (this.name)` checks. If a caller constructed a
 * start marker via the public constructor without supplying one of the
 * required fields, the emitted XML silently omitted the REQUIRED attribute
 * and produced schema-invalid output that strict OOXML validators reject
 * with "The required attribute 'author' is missing".
 *
 * Fix: always emit every required attribute, supplying sensible defaults
 * when a caller didn't specify one. This keeps the document schema-valid
 * while preserving backward-compat with every `toXML()` caller that
 * already assumes an XMLElement return.
 */

import { RangeMarker } from '../../src/elements/RangeMarker';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('RangeMarker required-attribute compliance', () => {
  describe('moveFromRangeStart / moveToRangeStart (CT_MoveBookmark)', () => {
    it('emits all required attributes for a properly-constructed start marker', () => {
      const marker = RangeMarker.createMoveFromStart(
        1,
        'moveAlpha',
        'Tester',
        new Date('2026-01-01T00:00:00Z')
      );
      const xml = XMLBuilder.elementToString(marker.toXML());
      expect(xml).toMatch(/w:id="1"/);
      expect(xml).toMatch(/w:name="moveAlpha"/);
      expect(xml).toMatch(/w:author="Tester"/);
      expect(xml).toMatch(/w:date="2026-01-01T00:00:00Z"/);
    });

    it('supplies default w:author="Unknown" when missing', () => {
      const marker = new RangeMarker({
        id: 2,
        type: 'moveFromRangeStart',
        name: 'moveBeta',
        date: new Date('2026-01-01T00:00:00Z'),
      });
      const xml = XMLBuilder.elementToString(marker.toXML());
      expect(xml).toMatch(/w:author="Unknown"/);
      // Must NOT emit the start marker missing w:author entirely.
      expect(xml).not.toMatch(/<w:moveFromRangeStart\s+w:id="2"\s+w:name="moveBeta"\s*\/>/);
    });

    it('synthesizes w:name="_move_${id}" when missing', () => {
      const marker = new RangeMarker({
        id: 3,
        type: 'moveFromRangeStart',
        author: 'Tester',
        date: new Date('2026-01-01T00:00:00Z'),
      });
      const xml = XMLBuilder.elementToString(marker.toXML());
      expect(xml).toMatch(/w:name="_move_3"/);
    });

    it('supplies default date (current time) when missing', () => {
      const marker = new RangeMarker({
        id: 4,
        type: 'moveToRangeStart',
        name: 'moveGamma',
        author: 'Tester',
      });
      const xml = XMLBuilder.elementToString(marker.toXML());
      // ISO 8601 with 'Z' suffix per formatDateForXml.
      expect(xml).toMatch(/w:date="\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z"/);
    });
  });

  describe('customXml*RangeStart (CT_TrackChangeRange)', () => {
    it('supplies default w:author when missing on customXmlInsRangeStart', () => {
      const marker = new RangeMarker({
        id: 5,
        type: 'customXmlInsRangeStart',
      });
      const xml = XMLBuilder.elementToString(marker.toXML());
      expect(xml).toMatch(/w:author="Unknown"/);
      expect(xml).toMatch(/w:date=/);
      // customXml* markers do NOT carry w:name in the schema.
      expect(xml).not.toMatch(/w:name=/);
    });

    it('emits when all required attrs supplied for customXmlInsRangeStart', () => {
      const marker = new RangeMarker({
        id: 6,
        type: 'customXmlInsRangeStart',
        author: 'Tester',
        date: new Date('2026-01-01T00:00:00Z'),
      });
      const xml = XMLBuilder.elementToString(marker.toXML());
      expect(xml).toMatch(/<w:customXmlInsRangeStart[^>]*w:id="6"/);
      expect(xml).toMatch(/w:author="Tester"/);
      expect(xml).not.toMatch(/w:name=/);
    });
  });

  describe('end markers (CT_MarkupRange) need only w:id', () => {
    it('emits moveFromRangeEnd with just w:id', () => {
      const marker = RangeMarker.createMoveFromEnd(7);
      const xml = XMLBuilder.elementToString(marker.toXML());
      expect(xml).toMatch(/<w:moveFromRangeEnd[^>]*w:id="7"/);
      expect(xml).not.toMatch(/w:author=|w:date=|w:name=/);
    });
  });
});
