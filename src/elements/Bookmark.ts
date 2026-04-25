/**
 * Bookmark - Represents a bookmark in a Word document
 *
 * Bookmarks mark specific locations in a document for internal navigation.
 * They consist of a start marker and an end marker with matching IDs.
 */

import { XMLElement } from '../xml/XMLBuilder.js';

/**
 * `w:displacedByCustomXml` attribute value per ECMA-376 CT_MarkupRange —
 * ST_DisplacedByCustomXml enum: "next" | "prev". Indicates which side of
 * a custom-XML boundary the marker semantically belongs to when the
 * marker had to be displaced because a custom-XML node boundary fell at
 * the same position.
 */
export type DisplacedByCustomXml = 'next' | 'prev';

/**
 * Bookmark properties
 */
export interface BookmarkProperties {
  /** Unique bookmark ID (generated automatically if not provided) */
  id?: number;
  /** Bookmark name (must be unique within document) */
  name: string;
  /** Skip name normalization (used when loading from existing documents) */
  skipNormalization?: boolean;
  /** First column in table bookmark range (ECMA-376 §17.16.5) */
  colFirst?: number;
  /** Last column in table bookmark range (ECMA-376 §17.16.5) */
  colLast?: number;
  /**
   * `w:displacedByCustomXml` per ECMA-376 CT_MarkupRange / CT_Bookmark.
   * Valid on both `<w:bookmarkStart>` (CT_Bookmark) and `<w:bookmarkEnd>`
   * (CT_MarkupRange). Preserves the custom-XML boundary disambiguator
   * that Word emits when a bookmark was displaced by a custom-XML node.
   */
  displacedByCustomXml?: DisplacedByCustomXml;
}

/**
 * Represents a bookmark location in a document
 */
export class Bookmark {
  private id: number;
  private name: string;
  private colFirst?: number;
  private colLast?: number;
  private displacedByCustomXml?: DisplacedByCustomXml;

  /**
   * Creates a new Bookmark
   * @param properties - Bookmark properties
   */
  constructor(properties: BookmarkProperties) {
    this.id = properties.id ?? 0; // ID will be assigned by BookmarkManager
    // Preserve exact bookmark names when loading from documents (Word allows =, ., etc.)
    // Only normalize when creating new bookmarks programmatically
    this.name = properties.skipNormalization
      ? properties.name
      : this.normalizeName(properties.name);
    this.colFirst = properties.colFirst;
    this.colLast = properties.colLast;
    this.displacedByCustomXml = properties.displacedByCustomXml;
  }

  /**
   * Normalizes a bookmark name to be valid
   * - Must start with a letter or underscore
   * - Can only contain letters, numbers, and underscores
   * - Maximum 40 characters
   * - Case-insensitive (Word converts to lowercase)
   * @param name - Raw bookmark name
   * @returns Normalized bookmark name
   */
  private normalizeName(name: string): string {
    // Remove invalid characters
    let normalized = name.replace(/[^a-zA-Z0-9_]/g, '_');

    // Ensure it starts with a letter or underscore
    if (normalized.length > 0 && /^\d/.test(normalized)) {
      normalized = '_' + normalized;
    }

    // Limit to 40 characters
    if (normalized.length > 40) {
      normalized = normalized.substring(0, 40);
    }

    // Default if empty
    if (normalized.length === 0) {
      normalized = '_bookmark';
    }

    return normalized;
  }

  /**
   * Gets the bookmark ID
   */
  getId(): number {
    return this.id;
  }

  /**
   * Sets the bookmark ID (used by BookmarkManager)
   * @internal
   */
  setId(id: number): void {
    this.id = id;
  }

  /**
   * Gets the bookmark name
   */
  getName(): string {
    return this.name;
  }

  /**
   * Sets the bookmark name (normalizes the name)
   *
   * The name will be normalized:
   * - Invalid characters replaced with underscores
   * - Leading digits prefixed with underscore
   * - Limited to 40 characters
   *
   * @param name - New bookmark name
   * @returns This bookmark for chaining
   */
  setName(name: string): this {
    this.name = this.normalizeName(name);
    return this;
  }

  /**
   * Sets the bookmark name without normalization
   *
   * Use this method when:
   * - Preserving exact names from imported documents
   * - Setting names with special characters that Word allows (=, ., etc.)
   * - Round-trip fidelity is required
   *
   * **Warning:** Setting invalid bookmark names may cause issues in Word.
   * Only use this if you know the name is valid or needs to match an existing document.
   *
   * @param name - Raw bookmark name (not normalized)
   * @returns This bookmark for chaining
   *
   * @example
   * ```typescript
   * // Preserve exact name from Word document
   * bookmark.setRawName('SECTION=II.MNKE7E8NA385_');
   *
   * // For new bookmarks, prefer setName() which normalizes
   * bookmark.setName('My Heading'); // Becomes 'My_Heading'
   * ```
   */
  setRawName(name: string): this {
    this.name = name;
    return this;
  }

  /**
   * Gets the first column in a table bookmark range
   */
  getColFirst(): number | undefined {
    return this.colFirst;
  }

  /**
   * Gets the last column in a table bookmark range
   */
  getColLast(): number | undefined {
    return this.colLast;
  }

  /**
   * Sets the column range for a table bookmark (ECMA-376 §17.16.5)
   */
  setColumnRange(colFirst: number, colLast: number): this {
    this.colFirst = colFirst;
    this.colLast = colLast;
    return this;
  }

  /**
   * Gets the `w:displacedByCustomXml` attribute value.
   * Preserved on both bookmarkStart and bookmarkEnd per ECMA-376
   * CT_MarkupRange (§17.13.5) / CT_Bookmark (§17.16.5).
   */
  getDisplacedByCustomXml(): DisplacedByCustomXml | undefined {
    return this.displacedByCustomXml;
  }

  /**
   * Sets the `w:displacedByCustomXml` attribute value.
   * @param value "next" | "prev" | undefined
   */
  setDisplacedByCustomXml(value: DisplacedByCustomXml | undefined): this {
    this.displacedByCustomXml = value;
    return this;
  }

  /**
   * Generates XML for the bookmark start marker
   * @returns XMLElement for bookmarkStart
   */
  toStartXML(): XMLElement {
    const attrs: Record<string, string | number> = {
      'w:id': this.id.toString(),
      'w:name': this.name,
    };
    // Table bookmark column range per ECMA-376 §17.16.5
    if (this.colFirst !== undefined) attrs['w:colFirst'] = this.colFirst.toString();
    if (this.colLast !== undefined) attrs['w:colLast'] = this.colLast.toString();
    // Custom-XML displacement marker per CT_MarkupRange (§17.13.5) —
    // the same attribute also carried by the RangeMarker sibling class.
    if (this.displacedByCustomXml) {
      attrs['w:displacedByCustomXml'] = this.displacedByCustomXml;
    }
    return {
      name: 'w:bookmarkStart',
      attributes: attrs,
      selfClosing: true,
    };
  }

  /**
   * Generates XML for the bookmark end marker
   * @returns XMLElement for bookmarkEnd
   */
  toEndXML(): XMLElement {
    const attrs: Record<string, string | number> = {
      'w:id': this.id.toString(),
    };
    // bookmarkEnd is CT_MarkupRange per ECMA-376 §17.13.6.1 — w:displacedByCustomXml
    // is valid here too (the BookmarkRange column attributes are NOT, so those
    // stay scoped to toStartXML above).
    if (this.displacedByCustomXml) {
      attrs['w:displacedByCustomXml'] = this.displacedByCustomXml;
    }
    return {
      name: 'w:bookmarkEnd',
      attributes: attrs,
      selfClosing: true,
    };
  }

  /**
   * Generates both start and end XML elements as an array
   * @returns Array of XMLElements [start, end]
   */
  toXML(): [XMLElement, XMLElement] {
    return [this.toStartXML(), this.toEndXML()];
  }

  /**
   * Creates a new Bookmark
   * @param name - Bookmark name
   * @returns New Bookmark instance
   */
  static create(name: string): Bookmark {
    return new Bookmark({ name });
  }

  /**
   * Creates a bookmark for a heading
   * Useful for table of contents internal links
   * @param headingText - The text of the heading
   * @returns New Bookmark instance with normalized name
   */
  static createForHeading(headingText: string): Bookmark {
    // Create a bookmark name from heading text
    // Example: "Chapter 1: Introduction" -> "_Chapter_1_Introduction"
    const name = headingText
      .trim()
      .replace(/[^a-zA-Z0-9]+/g, '_')
      .substring(0, 40);
    return new Bookmark({ name: name || '_heading' });
  }

  /**
   * Creates a bookmark with an auto-generated unique name
   * @param prefix - Optional prefix for the name (default: 'bookmark')
   * @returns New Bookmark instance
   */
  static createAuto(prefix = 'bookmark'): Bookmark {
    const timestamp = Date.now().toString(36);
    const random = Math.random().toString(36).substring(2, 7);
    const name = `${prefix}_${timestamp}_${random}`;
    return new Bookmark({ name });
  }
}
