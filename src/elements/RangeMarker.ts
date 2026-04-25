/**
 * RangeMarker - Represents range markers for tracked changes
 *
 * Range markers are used to mark the boundaries of moved or deleted content spans.
 * They enable Word to properly display and manage complex revision regions.
 */

import { XMLElement } from '../xml/XMLBuilder.js';
import { formatDateForXml } from '../utils/dateFormatting.js';

/**
 * Range marker type
 */
export type RangeMarkerType =
  | 'moveFromRangeStart'
  | 'moveFromRangeEnd'
  | 'moveToRangeStart'
  | 'moveToRangeEnd'
  | 'customXmlMoveFromRangeStart'
  | 'customXmlMoveFromRangeEnd'
  | 'customXmlMoveToRangeStart'
  | 'customXmlMoveToRangeEnd'
  | 'customXmlInsRangeStart'
  | 'customXmlInsRangeEnd'
  | 'customXmlDelRangeStart'
  | 'customXmlDelRangeEnd';

/**
 * `w:displacedByCustomXml` attribute value per ECMA-376 §17.13.5 — ST_DisplacedByCustomXml.
 * Indicates which side of a custom-XML boundary a displaced range marker
 * semantically belongs to.
 */
export type DisplacedByCustomXml = 'next' | 'prev';

/**
 * Range marker properties
 */
export interface RangeMarkerProperties {
  /** Unique ID for this range marker */
  id: number;
  /** Type of range marker */
  type: RangeMarkerType;
  /** Name linking start and end markers (for move operations) */
  name?: string;
  /** Author who made the change (for start markers only) */
  author?: string;
  /** Date when the change was made (for start markers only) */
  date?: Date;
  /**
   * `w:displacedByCustomXml` per ECMA-376 CT_MarkupRange / CT_MoveBookmark.
   * Only valid on moveFromRangeStart, moveToRangeStart, moveFromRangeEnd,
   * and moveToRangeEnd (the CT_MarkupRange-derived marker types).
   * Ignored on serialize for other types. "next" / "prev".
   */
  displacedByCustomXml?: DisplacedByCustomXml;
  /**
   * `w:colFirst` per ECMA-376 CT_BookmarkRange — first column index for
   * table-column-scoped moves. Only valid on moveFromRangeStart /
   * moveToRangeStart (CT_MoveBookmark-derived). Ignored on serialize for
   * other types. 0-based, inclusive.
   */
  colFirst?: number;
  /**
   * `w:colLast` per ECMA-376 CT_BookmarkRange — last column index for
   * table-column-scoped moves. Only valid on moveFromRangeStart /
   * moveToRangeStart. 0-based, inclusive.
   */
  colLast?: number;
}

/**
 * Represents a range marker for tracked changes
 * Range markers mark the boundaries of moved, inserted, or deleted content
 */
export class RangeMarker {
  private id: number;
  private type: RangeMarkerType;
  private name?: string;
  private author?: string;
  private date?: Date;
  private displacedByCustomXml?: DisplacedByCustomXml;
  private colFirst?: number;
  private colLast?: number;

  /**
   * Creates a new RangeMarker
   * @param properties - Range marker properties
   */
  constructor(properties: RangeMarkerProperties) {
    this.id = properties.id;
    this.type = properties.type;
    this.name = properties.name;
    this.author = properties.author;
    this.date = properties.date;
    this.displacedByCustomXml = properties.displacedByCustomXml;
    this.colFirst = properties.colFirst;
    this.colLast = properties.colLast;
  }

  /**
   * Gets the range marker ID
   */
  getId(): number {
    return this.id;
  }

  /**
   * Sets the range marker ID (used by RevisionManager)
   * @internal
   */
  setId(id: number): void {
    this.id = id;
  }

  /**
   * Gets the range marker type
   */
  getType(): RangeMarkerType {
    return this.type;
  }

  /**
   * Gets the range marker name
   */
  getName(): string | undefined {
    return this.name;
  }

  /**
   * Gets the author
   */
  getAuthor(): string | undefined {
    return this.author;
  }

  /**
   * Gets the date
   */
  getDate(): Date | undefined {
    return this.date;
  }

  /**
   * Gets the w:displacedByCustomXml attribute value. Only meaningful on
   * moveFrom/moveTo range start/end markers; undefined on all other types.
   */
  getDisplacedByCustomXml(): DisplacedByCustomXml | undefined {
    return this.displacedByCustomXml;
  }

  /**
   * Gets the w:colFirst column index. Only meaningful on CT_MoveBookmark-
   * derived start markers; undefined otherwise.
   */
  getColFirst(): number | undefined {
    return this.colFirst;
  }

  /**
   * Gets the w:colLast column index. Only meaningful on CT_MoveBookmark-
   * derived start markers; undefined otherwise.
   */
  getColLast(): number | undefined {
    return this.colLast;
  }

  /**
   * Checks if this is a start marker
   */
  isStartMarker(): boolean {
    return this.type.endsWith('RangeStart');
  }

  /**
   * Checks if this is an end marker
   */
  isEndMarker(): boolean {
    return this.type.endsWith('RangeEnd');
  }

  /**
   * Gets the XML element name for this range marker type
   */
  private getElementName(): string {
    switch (this.type) {
      case 'moveFromRangeStart':
        return 'w:moveFromRangeStart';
      case 'moveFromRangeEnd':
        return 'w:moveFromRangeEnd';
      case 'moveToRangeStart':
        return 'w:moveToRangeStart';
      case 'moveToRangeEnd':
        return 'w:moveToRangeEnd';
      case 'customXmlMoveFromRangeStart':
        return 'w:customXmlMoveFromRangeStart';
      case 'customXmlMoveFromRangeEnd':
        return 'w:customXmlMoveFromRangeEnd';
      case 'customXmlMoveToRangeStart':
        return 'w:customXmlMoveToRangeStart';
      case 'customXmlMoveToRangeEnd':
        return 'w:customXmlMoveToRangeEnd';
      case 'customXmlInsRangeStart':
        return 'w:customXmlInsRangeStart';
      case 'customXmlInsRangeEnd':
        return 'w:customXmlInsRangeEnd';
      case 'customXmlDelRangeStart':
        return 'w:customXmlDelRangeStart';
      case 'customXmlDelRangeEnd':
        return 'w:customXmlDelRangeEnd';
      default:
        return 'w:moveFromRangeStart';
    }
  }

  /**
   * Formats a date to ISO 8601 format for XML
   * Uses formatDateForXml() to strip milliseconds which Word does not accept.
   */
  private formatDate(date: Date): string {
    return formatDateForXml(date);
  }

  /**
   * Returns true for marker types whose OOXML schema base is CT_MoveBookmark
   * (ECMA-376 §17.13.5.15 / 19). Only these carry a w:name attribute.
   * moveFromRangeEnd/moveToRangeEnd derive from CT_MarkupRange, and all
   * customXml* markers derive from CT_TrackChange or CT_Markup — none of
   * those have w:name in the schema, so emitting it produces invalid XML
   * that fails Open XML SDK validation.
   */
  private supportsName(): boolean {
    return this.type === 'moveFromRangeStart' || this.type === 'moveToRangeStart';
  }

  /**
   * Returns true for marker types derived from CT_MarkupRange, which are the
   * only ones whose schema permits `w:displacedByCustomXml` (ECMA-376 §17.13.5).
   * customXml* markers (CT_TrackChange / CT_Markup) do NOT have this attribute,
   * so emitting it there would fail schema validation.
   */
  private supportsDisplacedByCustomXml(): boolean {
    return (
      this.type === 'moveFromRangeStart' ||
      this.type === 'moveToRangeStart' ||
      this.type === 'moveFromRangeEnd' ||
      this.type === 'moveToRangeEnd'
    );
  }

  /**
   * Returns true for marker types derived from CT_BookmarkRange, the only
   * ones whose schema permits `w:colFirst` / `w:colLast` (ECMA-376 §17.13.5).
   * The end markers and customXml* markers do NOT have these attributes.
   */
  private supportsColumnRange(): boolean {
    return this.type === 'moveFromRangeStart' || this.type === 'moveToRangeStart';
  }

  /**
   * Generates XML for this range marker.
   *
   * Required-attribute compliance per ECMA-376:
   *   - moveFromRangeStart / moveToRangeStart (CT_MoveBookmark §17.13.5.24):
   *     REQUIRE w:id, w:author, w:date, w:name. If any is missing, supply a
   *     sentinel default so the emitted XML stays schema-valid:
   *       - author → "Unknown"     (matches CT_TrackChange fallback elsewhere)
   *       - date   → new Date()    (best-effort; the revision couldn't be
   *                                 recorded without any time information)
   *       - name   → "_move_${id}" (synthesized link name; both halves of a
   *                                 move operation must use the same name,
   *                                 so construct it deterministically from
   *                                 the id)
   *   - customXml*RangeStart (CT_TrackChangeRange §17.13.5.5): REQUIRE
   *     w:id, w:author, w:date. Same defaults for author/date.
   *   - End markers (CT_MarkupRange §17.13.5.4): only w:id required.
   *
   * Defaults are preferred over throwing or returning null because they
   * keep the surrounding tracked-change context intact and leave the
   * document openable. Factory methods (`createMoveFromStart`, etc.)
   * already enforce the required fields at construction, so the defaults
   * only fire when a caller directly invokes the public constructor with
   * incomplete properties.
   *
   * @returns XMLElement representing the range marker.
   */
  toXML(): XMLElement {
    const attributes: Record<string, string> = {
      'w:id': this.id.toString(),
    };

    const isStart = this.isStartMarker();
    const isMoveBookmark = this.supportsName();

    // Emit w:name only on CT_MoveBookmark-derived elements (moveFromRangeStart /
    // moveToRangeStart). CT_MoveBookmark requires w:name (§17.13.5.24); supply a
    // synthesized fallback if the caller constructed the marker without one.
    // Other range markers — customXml*Range* (CT_TrackChange / CT_Markup) and
    // moveFrom/ToRangeEnd (CT_MarkupRange) — have no w:name in their schema.
    if (isMoveBookmark) {
      attributes['w:name'] = this.name || `_move_${this.id}`;
    }

    // Author and date are required on both CT_MoveBookmark and CT_TrackChange
    // start markers. Supply defaults if missing so the emission stays valid.
    if (isStart) {
      attributes['w:author'] = this.author || 'Unknown';
      attributes['w:date'] = this.formatDate(this.date ?? new Date());
    }

    // Emit w:displacedByCustomXml only on CT_MarkupRange-derived types
    // (move*RangeStart / move*RangeEnd). On CT_TrackChange and CT_Markup
    // elements the attribute is not defined by the schema.
    if (this.displacedByCustomXml && this.supportsDisplacedByCustomXml()) {
      attributes['w:displacedByCustomXml'] = this.displacedByCustomXml;
    }

    // Emit w:colFirst / w:colLast only on CT_BookmarkRange-derived start
    // markers (moveFromRangeStart / moveToRangeStart). Explicit-undefined
    // check so valid column index 0 is not dropped.
    if (this.supportsColumnRange()) {
      if (this.colFirst !== undefined) {
        attributes['w:colFirst'] = this.colFirst.toString();
      }
      if (this.colLast !== undefined) {
        attributes['w:colLast'] = this.colLast.toString();
      }
    }

    const elementName = this.getElementName();

    return {
      name: elementName,
      attributes,
      children: [], // Range markers are self-closing
    };
  }

  /**
   * Creates a moveFromRangeStart marker
   * @param id - Unique ID
   * @param name - Move operation name (links start/end markers)
   * @param author - Author who made the move
   * @param date - Optional date (defaults to now)
   * @returns New RangeMarker instance
   */
  static createMoveFromStart(id: number, name: string, author: string, date?: Date): RangeMarker {
    return new RangeMarker({
      id,
      type: 'moveFromRangeStart',
      name,
      author,
      date: date || new Date(),
    });
  }

  /**
   * Creates a moveFromRangeEnd marker
   * @param id - Unique ID (must match the start marker)
   * @returns New RangeMarker instance
   */
  static createMoveFromEnd(id: number): RangeMarker {
    return new RangeMarker({
      id,
      type: 'moveFromRangeEnd',
    });
  }

  /**
   * Creates a moveToRangeStart marker
   * @param id - Unique ID
   * @param name - Move operation name (links to moveFrom markers)
   * @param author - Author who made the move
   * @param date - Optional date (defaults to now)
   * @returns New RangeMarker instance
   */
  static createMoveToStart(id: number, name: string, author: string, date?: Date): RangeMarker {
    return new RangeMarker({
      id,
      type: 'moveToRangeStart',
      name,
      author,
      date: date || new Date(),
    });
  }

  /**
   * Creates a moveToRangeEnd marker
   * @param id - Unique ID (must match the start marker)
   * @returns New RangeMarker instance
   */
  static createMoveToEnd(id: number): RangeMarker {
    return new RangeMarker({
      id,
      type: 'moveToRangeEnd',
    });
  }

  /**
   * Creates a customXmlInsRangeStart marker
   * @param id - Unique ID
   * @param author - Author who made the insertion
   * @param date - Optional date (defaults to now)
   * @returns New RangeMarker instance
   */
  static createCustomXmlInsStart(id: number, author: string, date?: Date): RangeMarker {
    return new RangeMarker({
      id,
      type: 'customXmlInsRangeStart',
      author,
      date: date || new Date(),
    });
  }

  /**
   * Creates a customXmlInsRangeEnd marker
   * @param id - Unique ID (must match the start marker)
   * @returns New RangeMarker instance
   */
  static createCustomXmlInsEnd(id: number): RangeMarker {
    return new RangeMarker({
      id,
      type: 'customXmlInsRangeEnd',
    });
  }

  /**
   * Creates a customXmlDelRangeStart marker
   * @param id - Unique ID
   * @param author - Author who made the deletion
   * @param date - Optional date (defaults to now)
   * @returns New RangeMarker instance
   */
  static createCustomXmlDelStart(id: number, author: string, date?: Date): RangeMarker {
    return new RangeMarker({
      id,
      type: 'customXmlDelRangeStart',
      author,
      date: date || new Date(),
    });
  }

  /**
   * Creates a customXmlDelRangeEnd marker
   * @param id - Unique ID (must match the start marker)
   * @returns New RangeMarker instance
   */
  static createCustomXmlDelEnd(id: number): RangeMarker {
    return new RangeMarker({
      id,
      type: 'customXmlDelRangeEnd',
    });
  }

  /**
   * Creates a customXmlMoveFromRangeStart marker
   * @param id - Unique ID
   * @param name - Move operation name
   * @param author - Author who made the move
   * @param date - Optional date (defaults to now)
   * @returns New RangeMarker instance
   */
  static createCustomXmlMoveFromStart(
    id: number,
    name: string,
    author: string,
    date?: Date
  ): RangeMarker {
    return new RangeMarker({
      id,
      type: 'customXmlMoveFromRangeStart',
      name,
      author,
      date: date || new Date(),
    });
  }

  /**
   * Creates a customXmlMoveFromRangeEnd marker
   * @param id - Unique ID (must match the start marker)
   * @returns New RangeMarker instance
   */
  static createCustomXmlMoveFromEnd(id: number): RangeMarker {
    return new RangeMarker({
      id,
      type: 'customXmlMoveFromRangeEnd',
    });
  }

  /**
   * Creates a customXmlMoveToRangeStart marker
   * @param id - Unique ID
   * @param name - Move operation name
   * @param author - Author who made the move
   * @param date - Optional date (defaults to now)
   * @returns New RangeMarker instance
   */
  static createCustomXmlMoveToStart(
    id: number,
    name: string,
    author: string,
    date?: Date
  ): RangeMarker {
    return new RangeMarker({
      id,
      type: 'customXmlMoveToRangeStart',
      name,
      author,
      date: date || new Date(),
    });
  }

  /**
   * Creates a customXmlMoveToRangeEnd marker
   * @param id - Unique ID (must match the start marker)
   * @returns New RangeMarker instance
   */
  static createCustomXmlMoveToEnd(id: number): RangeMarker {
    return new RangeMarker({
      id,
      type: 'customXmlMoveToRangeEnd',
    });
  }
}
