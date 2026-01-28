/**
 * Bookmark - Represents a bookmark in a Word document
 *
 * Bookmarks mark specific locations in a document for internal navigation.
 * They consist of a start marker and an end marker with matching IDs.
 */

import { XMLElement } from '../xml/XMLBuilder';

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
}

/**
 * Represents a bookmark location in a document
 */
export class Bookmark {
  private id: number;
  private name: string;

  /**
   * Creates a new Bookmark
   * @param properties - Bookmark properties
   */
  constructor(properties: BookmarkProperties) {
    this.id = properties.id ?? 0; // ID will be assigned by BookmarkManager
    // Preserve exact bookmark names when loading from documents (Word allows =, ., etc.)
    // Only normalize when creating new bookmarks programmatically
    this.name = properties.skipNormalization ? properties.name : this.normalizeName(properties.name);
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
   * Generates XML for the bookmark start marker
   * @returns XMLElement for bookmarkStart
   */
  toStartXML(): XMLElement {
    return {
      name: 'w:bookmarkStart',
      attributes: {
        'w:id': this.id.toString(),
        'w:name': this.name,
      },
      selfClosing: true,
    };
  }

  /**
   * Generates XML for the bookmark end marker
   * @returns XMLElement for bookmarkEnd
   */
  toEndXML(): XMLElement {
    return {
      name: 'w:bookmarkEnd',
      attributes: {
        'w:id': this.id.toString(),
      },
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
  static createAuto(prefix: string = 'bookmark'): Bookmark {
    const timestamp = Date.now().toString(36);
    const random = Math.random().toString(36).substring(2, 7);
    const name = `${prefix}_${timestamp}_${random}`;
    return new Bookmark({ name });
  }
}
