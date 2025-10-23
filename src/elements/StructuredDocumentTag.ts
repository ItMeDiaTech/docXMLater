/**
 * Structured Document Tag (SDT) - Content control wrapper
 *
 * SDTs are used by applications like Google Docs to wrap content
 * with metadata and control settings. They can contain paragraphs,
 * tables, or other block-level elements.
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { Paragraph } from './Paragraph';
import { Table } from './Table';

/**
 * Type of content lock for SDT
 */
export type SDTLockType = 'unlocked' | 'sdtLocked' | 'contentLocked' | 'sdtContentLocked';

/**
 * Properties for a Structured Document Tag
 */
export interface SDTProperties {
  /** Unique ID for this SDT */
  id?: number;
  /** Tag value (used by applications) */
  tag?: string;
  /** Lock type */
  lock?: SDTLockType;
  /** Alias (display name) */
  alias?: string;
}

/**
 * Content that can be wrapped by an SDT
 */
export type SDTContent = Table | Paragraph | StructuredDocumentTag;

/**
 * Structured Document Tag class
 * Wraps content with metadata and control settings
 *
 * Example XML:
 * ```xml
 * <w:sdt>
 *   <w:sdtPr>
 *     <w:lock w:val="contentLocked"/>
 *     <w:id w:val="-258490443"/>
 *     <w:tag w:val="goog_rdk_0"/>
 *   </w:sdtPr>
 *   <w:sdtContent>
 *     <!-- content here -->
 *   </w:sdtContent>
 * </w:sdt>
 * ```
 */
export class StructuredDocumentTag {
  private properties: SDTProperties;
  private content: SDTContent[];

  /**
   * Create a new Structured Document Tag
   * @param properties - SDT properties
   * @param content - Content elements to wrap
   */
  constructor(properties: SDTProperties = {}, content: SDTContent[] = []) {
    this.properties = properties;
    this.content = content;
  }

  /**
   * Get the SDT ID
   * @returns SDT ID or undefined
   */
  getId(): number | undefined {
    return this.properties.id;
  }

  /**
   * Set the SDT ID
   * @param id - Unique ID
   */
  setId(id: number): this {
    this.properties.id = id;
    return this;
  }

  /**
   * Get the SDT tag
   * @returns Tag value or undefined
   */
  getTag(): string | undefined {
    return this.properties.tag;
  }

  /**
   * Set the SDT tag
   * @param tag - Tag value
   */
  setTag(tag: string): this {
    this.properties.tag = tag;
    return this;
  }

  /**
   * Get the lock type
   * @returns Lock type or undefined
   */
  getLock(): SDTLockType | undefined {
    return this.properties.lock;
  }

  /**
   * Set the lock type
   * @param lock - Lock type
   */
  setLock(lock: SDTLockType): this {
    this.properties.lock = lock;
    return this;
  }

  /**
   * Get the alias (display name)
   * @returns Alias or undefined
   */
  getAlias(): string | undefined {
    return this.properties.alias;
  }

  /**
   * Set the alias (display name)
   * @param alias - Display name
   */
  setAlias(alias: string): this {
    this.properties.alias = alias;
    return this;
  }

  /**
   * Get all content elements
   * @returns Array of content elements
   */
  getContent(): SDTContent[] {
    return [...this.content];
  }

  /**
   * Add content element
   * @param element - Element to add
   */
  addContent(element: SDTContent): this {
    this.content.push(element);
    return this;
  }

  /**
   * Clear all content
   */
  clearContent(): this {
    this.content = [];
    return this;
  }

  /**
   * Generate XML for this SDT
   * @returns XML element
   */
  toXML(): XMLElement {
    const children: XMLElement[] = [];

    // Build sdtPr (properties)
    const sdtPrChildren: XMLElement[] = [];

    if (this.properties.lock) {
      sdtPrChildren.push(
        XMLBuilder.wSelf('lock', { 'w:val': this.properties.lock })
      );
    }

    if (this.properties.id !== undefined) {
      sdtPrChildren.push(
        XMLBuilder.wSelf('id', { 'w:val': this.properties.id.toString() })
      );
    }

    if (this.properties.tag) {
      sdtPrChildren.push(
        XMLBuilder.wSelf('tag', { 'w:val': this.properties.tag })
      );
    }

    if (this.properties.alias) {
      sdtPrChildren.push(
        XMLBuilder.wSelf('alias', { 'w:val': this.properties.alias })
      );
    }

    if (sdtPrChildren.length > 0) {
      children.push(XMLBuilder.w('sdtPr', {}, sdtPrChildren));
    }

    // Build sdtContent
    const sdtContentChildren: XMLElement[] = [];
    for (const element of this.content) {
      sdtContentChildren.push(element.toXML());
    }

    children.push(XMLBuilder.w('sdtContent', {}, sdtContentChildren));

    return XMLBuilder.w('sdt', {}, children);
  }

  /**
   * Create an SDT wrapping a table (common for Google Docs)
   * @param table - Table to wrap
   * @param tag - Optional tag value
   * @returns New SDT instance
   */
  static wrapTable(table: Table, tag: string = 'goog_rdk_0'): StructuredDocumentTag {
    return new StructuredDocumentTag(
      {
        id: Date.now() % 1000000000, // Generate a reasonable ID
        tag,
        lock: 'contentLocked',
      },
      [table]
    );
  }

  /**
   * Create an SDT wrapping a paragraph
   * @param paragraph - Paragraph to wrap
   * @param tag - Optional tag value
   * @returns New SDT instance
   */
  static wrapParagraph(paragraph: Paragraph, tag?: string): StructuredDocumentTag {
    return new StructuredDocumentTag(
      {
        id: Date.now() % 1000000000,
        tag,
      },
      [paragraph]
    );
  }

  /**
   * Create an empty SDT
   * @param properties - SDT properties
   * @returns New SDT instance
   */
  static create(properties: SDTProperties = {}): StructuredDocumentTag {
    return new StructuredDocumentTag(properties);
  }
}
