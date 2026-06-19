/**
 * Header - Represents a document header
 *
 * Headers appear at the top of pages and can contain text, tables, images, and fields.
 * Different headers can be defined for first page, odd pages, and even pages.
 */

import { XMLBuilder } from '../xml/XMLBuilder.js';
import { XMLParser } from '../xml/XMLParser.js';
import { Paragraph } from './Paragraph.js';
import { RunFormatting } from './Run.js';
import { Table } from './Table.js';

/**
 * Header type
 */
export type HeaderType = 'default' | 'first' | 'even';

/**
 * Header content element
 */
type HeaderElement = Paragraph | Table;

/**
 * Header properties
 */
export interface HeaderProperties {
  /** Header type (default, first page, or even page) */
  type?: HeaderType;
}

/**
 * Represents a document header
 */
export class Header {
  private elements: HeaderElement[] = [];
  private type: HeaderType;
  private headerId?: string;
  private rawXML?: string; // Store original XML for preservation
  // Top-level w:p/w:tbl count of the preserved raw XML. DocumentParser calls
  // setRawXML() before repopulating elements from that same XML, so mutators
  // can't simply clear rawXML — they would fire during parsing and lose the
  // preserved bytes. Instead, growth beyond this count is the signal that
  // content was added after load and the preserved XML is stale.
  private rawXmlBlockCount = Number.MAX_SAFE_INTEGER;

  /**
   * Creates a new header
   * @param properties Header properties
   */
  constructor(properties: HeaderProperties = {}) {
    this.type = properties.type || 'default';
  }

  /**
   * Sets the raw XML content (used when loading existing headers)
   * @param xml Raw XML content
   */
  setRawXML(xml: string): this {
    this.rawXML = xml;
    this.rawXmlBlockCount = Header.countTopLevelBlocks(xml);
    return this;
  }

  /**
   * Counts the top-level w:p/w:tbl children of a raw header part.
   * Parsing populates this.elements from exactly those blocks, so the count
   * is the baseline against which mutators detect post-load additions.
   */
  private static countTopLevelBlocks(xml: string): number {
    try {
      const parsed = XMLParser.parseToObject(xml) as Record<string, unknown>;
      const root = parsed?.['w:hdr'];
      if (!root || typeof root !== 'object') {
        return 0;
      }
      const count = (node: unknown): number =>
        node === undefined ? 0 : Array.isArray(node) ? node.length : 1;
      const rootObj = root as Record<string, unknown>;
      return count(rootObj['w:p']) + count(rootObj['w:tbl']);
    } catch {
      // Unparseable part: never invalidate, keep the preserved bytes
      return Number.MAX_SAFE_INTEGER;
    }
  }

  /**
   * Drops the preserved raw XML once content grows beyond what that XML
   * already contains, so toXML() regenerates from elements and post-load
   * edits are not silently discarded on save.
   */
  private invalidateRawXmlOnGrowth(): void {
    if (this.rawXML !== undefined && this.elements.length > this.rawXmlBlockCount) {
      this.rawXML = undefined;
    }
  }

  /**
   * Gets the raw XML content if available
   */
  getRawXML(): string | undefined {
    return this.rawXML;
  }

  /**
   * Gets the header type
   */
  getType(): HeaderType {
    return this.type;
  }

  /**
   * Sets the header ID (used for relationships)
   * @param id Header ID
   */
  setHeaderId(id: string): this {
    this.headerId = id;
    return this;
  }

  /**
   * Gets the header ID
   */
  getHeaderId(): string | undefined {
    return this.headerId;
  }

  /**
   * Adds a paragraph to the header
   * @param paragraph Paragraph to add
   */
  addParagraph(paragraph: Paragraph): this {
    this.elements.push(paragraph);
    this.invalidateRawXmlOnGrowth();
    return this;
  }

  /**
   * Creates and adds a new paragraph
   * @param text Optional text content
   */
  createParagraph(text?: string): Paragraph {
    const para = new Paragraph();
    if (text) {
      para.addText(text);
    }
    this.elements.push(para);
    this.invalidateRawXmlOnGrowth();
    return para;
  }

  /**
   * Adds formatted text to the header as a new paragraph
   *
   * Convenience method that creates a paragraph with a single formatted run.
   * For headers with multiple runs or complex content, use `createParagraph()`
   * and build content manually.
   *
   * @param text - Text content
   * @param formatting - Optional run formatting (bold, font, size, etc.)
   * @returns The created Paragraph for further customization
   *
   * @example
   * ```typescript
   * header.addText('Company Name', { bold: true, font: 'Arial', size: 10 });
   * header.addText('Confidential', { italic: true, color: 'FF0000' });
   * ```
   */
  addText(text: string, formatting?: RunFormatting): Paragraph {
    const para = new Paragraph();
    para.addText(text, formatting);
    this.elements.push(para);
    this.invalidateRawXmlOnGrowth();
    return para;
  }

  /**
   * Adds a table to the header
   * @param table Table to add
   */
  addTable(table: Table): this {
    this.elements.push(table);
    this.invalidateRawXmlOnGrowth();
    return this;
  }

  /**
   * Creates and adds a new table
   * @param rows Number of rows
   * @param columns Number of columns
   */
  createTable(rows: number, columns: number): Table {
    const table = new Table(rows, columns);
    this.elements.push(table);
    this.invalidateRawXmlOnGrowth();
    return table;
  }

  /**
   * Gets all elements in the header
   */
  getElements(): HeaderElement[] {
    return [...this.elements];
  }

  /**
   * Gets the number of elements
   */
  getElementCount(): number {
    return this.elements.length;
  }

  /**
   * Clears all elements and resets raw XML so toXML() regenerates an empty header
   */
  clear(): this {
    this.elements = [];
    this.rawXML = undefined;
    return this;
  }

  /**
   * Generates the complete header XML file content
   * This creates a complete header document (header1.xml, etc.)
   */
  toXML(): string {
    // If we have raw XML preserved from loading, use it
    if (this.rawXML) {
      return this.rawXML;
    }

    // Serialize through XMLBuilder so text and attribute values are
    // XML-escaped and rawXml passthrough children are honored
    const children = this.elements.map((el) => el.toXML());

    // ECMA-376 requires block-level content in w:hdr
    if (children.length === 0) {
      children.push(XMLBuilder.w('p'));
    }

    const namespaces: Record<string, string> = {
      ...XMLBuilder.createNamespaces(),
      // Extension namespaces (w14:paraId etc. on regenerated paragraphs)
      // must be declared ignorable or Word rejects the part
      'mc:Ignorable': 'w14 w15 wp14 asvg',
    };

    const builder = new XMLBuilder();
    builder.element('w:hdr', namespaces, children);
    return builder.build(true);
  }

  /**
   * Gets the filename for this header
   * @param number Header number (1, 2, 3, etc.)
   */
  getFilename(number: number): string {
    return `header${number}.xml`;
  }

  /**
   * Creates a default header
   */
  static createDefault(): Header {
    return new Header({ type: 'default' });
  }

  /**
   * Creates a first page header
   */
  static createFirst(): Header {
    return new Header({ type: 'first' });
  }

  /**
   * Creates an even page header
   */
  static createEven(): Header {
    return new Header({ type: 'even' });
  }

  /**
   * Creates a header with properties
   * @param properties Header properties
   */
  static create(properties?: HeaderProperties): Header {
    return new Header(properties);
  }
}
