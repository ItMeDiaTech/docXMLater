/**
 * HeaderFooterManager - Manages headers and footers in a document
 *
 * Tracks all headers and footers, assigns unique IDs, generates filenames,
 * and manages relationships.
 */

import { Header } from './Header.js';
import { Footer } from './Footer.js';

/**
 * Header entry with metadata
 */
interface HeaderEntry {
  header: Header;
  filename: string;
  relationshipId: string;
  number: number;
}

/**
 * Footer entry with metadata
 */
interface FooterEntry {
  footer: Footer;
  filename: string;
  relationshipId: string;
  number: number;
}

/**
 * Manages all headers and footers in a document
 */
export class HeaderFooterManager {
  private headers: Map<Header, HeaderEntry>;
  private footers: Map<Footer, FooterEntry>;
  private nextHeaderNumber: number;
  private nextFooterNumber: number;

  /**
   * Creates a new header/footer manager
   */
  constructor() {
    this.headers = new Map();
    this.footers = new Map();
    this.nextHeaderNumber = 1;
    this.nextFooterNumber = 1;
  }

  /**
   * Registers a header with the manager
   * @param header The header to register
   * @param relationshipId The relationship ID for this header
   * @param originalFilename Part name from the relationship target for loaded
   *   headers — kept as-is so saved content lands where the relationship points
   * @returns The filename assigned to this header
   */
  registerHeader(header: Header, relationshipId: string, originalFilename?: string): string {
    // Check if already registered
    const existing = this.headers.get(header);
    if (existing) {
      return existing.filename;
    }

    let number: number;
    let filename: string;
    if (originalFilename) {
      filename = originalFilename;
      const digits = /(\d+)\.xml$/i.exec(originalFilename)?.[1];
      number = digits ? parseInt(digits, 10) : this.nextHeaderNumber;
      // Keep the counter ahead of parsed part names so generated names never collide
      if (number >= this.nextHeaderNumber) {
        this.nextHeaderNumber = number + 1;
      }
    } else {
      number = this.nextHeaderNumber++;
      filename = header.getFilename(number);
    }

    // Set header ID
    header.setHeaderId(relationshipId);

    // Store entry
    const entry: HeaderEntry = {
      header,
      filename,
      relationshipId,
      number,
    };

    this.headers.set(header, entry);

    return filename;
  }

  /**
   * Registers a footer with the manager
   * @param footer The footer to register
   * @param relationshipId The relationship ID for this footer
   * @param originalFilename Part name from the relationship target for loaded
   *   footers — kept as-is so saved content lands where the relationship points
   * @returns The filename assigned to this footer
   */
  registerFooter(footer: Footer, relationshipId: string, originalFilename?: string): string {
    // Check if already registered
    const existing = this.footers.get(footer);
    if (existing) {
      return existing.filename;
    }

    let number: number;
    let filename: string;
    if (originalFilename) {
      filename = originalFilename;
      const digits = /(\d+)\.xml$/i.exec(originalFilename)?.[1];
      number = digits ? parseInt(digits, 10) : this.nextFooterNumber;
      // Keep the counter ahead of parsed part names so generated names never collide
      if (number >= this.nextFooterNumber) {
        this.nextFooterNumber = number + 1;
      }
    } else {
      number = this.nextFooterNumber++;
      filename = footer.getFilename(number);
    }

    // Set footer ID
    footer.setFooterId(relationshipId);

    // Store entry
    const entry: FooterEntry = {
      footer,
      filename,
      relationshipId,
      number,
    };

    this.footers.set(footer, entry);

    return filename;
  }

  /**
   * Returns the part filename this header has, or would receive when
   * registered. Does not consume a sequence number — callers can create the
   * relationship targeting the same part name registerHeader() will assign.
   * @param header The header
   * @returns The current or upcoming filename
   */
  peekHeaderFilename(header: Header): string {
    return this.headers.get(header)?.filename ?? header.getFilename(this.nextHeaderNumber);
  }

  /**
   * Returns the part filename this footer has, or would receive when
   * registered. Does not consume a sequence number — callers can create the
   * relationship targeting the same part name registerFooter() will assign.
   * @param footer The footer
   * @returns The current or upcoming filename
   */
  peekFooterFilename(footer: Footer): string {
    return this.footers.get(footer)?.filename ?? footer.getFilename(this.nextFooterNumber);
  }

  /**
   * Gets the filename for a header
   * @param header The header
   * @returns The filename, or undefined if not registered
   */
  getHeaderFilename(header: Header): string | undefined {
    return this.headers.get(header)?.filename;
  }

  /**
   * Gets the filename for a footer
   * @param footer The footer
   * @returns The filename, or undefined if not registered
   */
  getFooterFilename(footer: Footer): string | undefined {
    return this.footers.get(footer)?.filename;
  }

  /**
   * Gets the relationship ID for a header
   * @param header The header
   * @returns The relationship ID, or undefined if not registered
   */
  getHeaderRelationshipId(header: Header): string | undefined {
    return this.headers.get(header)?.relationshipId;
  }

  /**
   * Gets the relationship ID for a footer
   * @param footer The footer
   * @returns The relationship ID, or undefined if not registered
   */
  getFooterRelationshipId(footer: Footer): string | undefined {
    return this.footers.get(footer)?.relationshipId;
  }

  /**
   * Gets all registered headers
   * @returns Array of header entries
   */
  getAllHeaders(): HeaderEntry[] {
    return Array.from(this.headers.values());
  }

  /**
   * Gets all registered footers
   * @returns Array of footer entries
   */
  getAllFooters(): FooterEntry[] {
    return Array.from(this.footers.values());
  }

  /**
   * Gets the number of headers
   * @returns Number of registered headers
   */
  getHeaderCount(): number {
    return this.headers.size;
  }

  /**
   * Gets the number of footers
   * @returns Number of registered footers
   */
  getFooterCount(): number {
    return this.footers.size;
  }

  /**
   * Checks if a header is registered
   * @param header The header
   * @returns True if registered
   */
  hasHeader(header: Header): boolean {
    return this.headers.has(header);
  }

  /**
   * Checks if a footer is registered
   * @param footer The footer
   * @returns True if registered
   */
  hasFooter(footer: Footer): boolean {
    return this.footers.has(footer);
  }

  /**
   * Clears all headers and footers
   */
  clear(): this {
    this.headers.clear();
    this.footers.clear();
    this.nextHeaderNumber = 1;
    this.nextFooterNumber = 1;
    return this;
  }

  /**
   * Creates a new header/footer manager
   */
  static create(): HeaderFooterManager {
    return new HeaderFooterManager();
  }
}
