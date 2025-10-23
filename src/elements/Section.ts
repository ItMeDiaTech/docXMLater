/**
 * Section - Represents a document section with page setup properties
 *
 * Sections allow different page setups within a single document (margins, orientation, etc.)
 * Each section can have its own headers, footers, and page numbering.
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { PAGE_SIZES } from '../utils/units';

/**
 * Page orientation
 */
export type PageOrientation = 'portrait' | 'landscape';

/**
 * Section break type
 */
export type SectionType = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';

/**
 * Page numbering format
 */
export type PageNumberFormat = 'decimal' | 'lowerRoman' | 'upperRoman' | 'lowerLetter' | 'upperLetter';

/**
 * Page size properties
 */
export interface PageSize {
  /** Width in twips */
  width: number;
  /** Height in twips */
  height: number;
  /** Orientation */
  orientation?: PageOrientation;
}

/**
 * Margin properties
 */
export interface Margins {
  /** Top margin in twips */
  top: number;
  /** Bottom margin in twips */
  bottom: number;
  /** Left margin in twips */
  left: number;
  /** Right margin in twips */
  right: number;
  /** Header distance from top in twips */
  header?: number;
  /** Footer distance from bottom in twips */
  footer?: number;
  /** Gutter margin in twips */
  gutter?: number;
}

/**
 * Column properties
 */
export interface Columns {
  /** Number of columns */
  count: number;
  /** Space between columns in twips */
  space?: number;
  /** Equal column widths */
  equalWidth?: boolean;
}

/**
 * Page numbering properties
 */
export interface PageNumbering {
  /** Starting page number */
  start?: number;
  /** Number format */
  format?: PageNumberFormat;
}

/**
 * Section properties
 */
export interface SectionProperties {
  /** Page size */
  pageSize?: PageSize;
  /** Margins */
  margins?: Margins;
  /** Column layout */
  columns?: Columns;
  /** Section break type */
  type?: SectionType;
  /** Page numbering */
  pageNumbering?: PageNumbering;
  /** Header reference IDs */
  headers?: {
    default?: string;  // rId for default header
    first?: string;    // rId for first page header
    even?: string;     // rId for even page header
  };
  /** Footer reference IDs */
  footers?: {
    default?: string;  // rId for default footer
    first?: string;    // rId for first page footer
    even?: string;     // rId for even page footer
  };
  /** Title page (different first page) */
  titlePage?: boolean;
}

/**
 * Represents a document section
 */
export class Section {
  private properties: SectionProperties;

  /**
   * Creates a new section
   * @param properties Section properties
   */
  constructor(properties: SectionProperties = {}) {
    // Set defaults only where necessary
    this.properties = {
      pageSize: properties.pageSize || {
        width: PAGE_SIZES.LETTER.width,
        height: PAGE_SIZES.LETTER.height,
        orientation: 'portrait',
      },
      margins: properties.margins || {
        top: 1440,    // 1 inch
        bottom: 1440,
        left: 1440,
        right: 1440,
        header: 720,  // 0.5 inch
        footer: 720,
      },
      // Default to single column layout
      columns: properties.columns || {
        count: 1,
      },
      // Default to next page section break
      type: properties.type || 'nextPage',
      pageNumbering: properties.pageNumbering,
      headers: properties.headers,
      footers: properties.footers,
      titlePage: properties.titlePage,
    };
  }

  /**
   * Gets the section properties
   */
  getProperties(): SectionProperties {
    return { ...this.properties };
  }

  /**
   * Sets page size
   * @param width Width in twips
   * @param height Height in twips
   * @param orientation Page orientation
   */
  setPageSize(width: number, height: number, orientation: PageOrientation = 'portrait'): this {
    this.properties.pageSize = { width, height, orientation };
    return this;
  }

  /**
   * Sets page orientation
   * @param orientation Page orientation
   */
  setOrientation(orientation: PageOrientation): this {
    if (!this.properties.pageSize) {
      this.properties.pageSize = {
        width: PAGE_SIZES.LETTER.width,
        height: PAGE_SIZES.LETTER.height,
      };
    }
    this.properties.pageSize.orientation = orientation;

    // Swap width/height for landscape
    if (orientation === 'landscape' && this.properties.pageSize.width < this.properties.pageSize.height) {
      const temp = this.properties.pageSize.width;
      this.properties.pageSize.width = this.properties.pageSize.height;
      this.properties.pageSize.height = temp;
    }

    return this;
  }

  /**
   * Sets margins
   * @param margins Margin properties
   */
  setMargins(margins: Margins): this {
    this.properties.margins = { ...margins };
    return this;
  }

  /**
   * Sets column layout
   * @param count Number of columns
   * @param space Space between columns in twips
   */
  setColumns(count: number, space: number = 720): this {
    this.properties.columns = {
      count,
      space,
      equalWidth: true,
    };
    return this;
  }

  /**
   * Sets section type
   * @param type Section break type
   */
  setSectionType(type: SectionType): this {
    this.properties.type = type;
    return this;
  }

  /**
   * Sets page numbering
   * @param start Starting page number
   * @param format Number format
   */
  setPageNumbering(start?: number, format?: PageNumberFormat): this {
    this.properties.pageNumbering = { start, format };
    return this;
  }

  /**
   * Sets title page flag (different first page)
   * @param titlePage Whether this section has a different first page
   */
  setTitlePage(titlePage: boolean = true): this {
    this.properties.titlePage = titlePage;
    return this;
  }

  /**
   * Sets header reference
   * @param type Header type (default, first, even)
   * @param rId Relationship ID
   */
  setHeaderReference(type: 'default' | 'first' | 'even', rId: string): this {
    if (!this.properties.headers) {
      this.properties.headers = {};
    }
    this.properties.headers[type] = rId;
    return this;
  }

  /**
   * Sets footer reference
   * @param type Footer type (default, first, even)
   * @param rId Relationship ID
   */
  setFooterReference(type: 'default' | 'first' | 'even', rId: string): this {
    if (!this.properties.footers) {
      this.properties.footers = {};
    }
    this.properties.footers[type] = rId;
    return this;
  }

  /**
   * Generates WordprocessingML XML for section properties
   */
  toXML(): XMLElement {
    const children: XMLElement[] = [];

    // Header references
    if (this.properties.headers) {
      if (this.properties.headers.first) {
        children.push(
          XMLBuilder.wSelf('headerReference', {
            'w:type': 'first',
            'r:id': this.properties.headers.first,
          })
        );
      }
      if (this.properties.headers.even) {
        children.push(
          XMLBuilder.wSelf('headerReference', {
            'w:type': 'even',
            'r:id': this.properties.headers.even,
          })
        );
      }
      if (this.properties.headers.default) {
        children.push(
          XMLBuilder.wSelf('headerReference', {
            'w:type': 'default',
            'r:id': this.properties.headers.default,
          })
        );
      }
    }

    // Footer references
    if (this.properties.footers) {
      if (this.properties.footers.first) {
        children.push(
          XMLBuilder.wSelf('footerReference', {
            'w:type': 'first',
            'r:id': this.properties.footers.first,
          })
        );
      }
      if (this.properties.footers.even) {
        children.push(
          XMLBuilder.wSelf('footerReference', {
            'w:type': 'even',
            'r:id': this.properties.footers.even,
          })
        );
      }
      if (this.properties.footers.default) {
        children.push(
          XMLBuilder.wSelf('footerReference', {
            'w:type': 'default',
            'r:id': this.properties.footers.default,
          })
        );
      }
    }

    // Section type
    if (this.properties.type) {
      children.push(
        XMLBuilder.wSelf('type', { 'w:val': this.properties.type })
      );
    }

    // Page size
    if (this.properties.pageSize) {
      const attrs: Record<string, string> = {
        'w:w': this.properties.pageSize.width.toString(),
        'w:h': this.properties.pageSize.height.toString(),
      };
      if (this.properties.pageSize.orientation === 'landscape') {
        attrs['w:orient'] = 'landscape';
      }
      children.push(XMLBuilder.wSelf('pgSz', attrs));
    }

    // Margins
    if (this.properties.margins) {
      const attrs: Record<string, string> = {
        'w:top': this.properties.margins.top.toString(),
        'w:right': this.properties.margins.right.toString(),
        'w:bottom': this.properties.margins.bottom.toString(),
        'w:left': this.properties.margins.left.toString(),
      };
      if (this.properties.margins.header !== undefined) {
        attrs['w:header'] = this.properties.margins.header.toString();
      }
      if (this.properties.margins.footer !== undefined) {
        attrs['w:footer'] = this.properties.margins.footer.toString();
      }
      if (this.properties.margins.gutter !== undefined) {
        attrs['w:gutter'] = this.properties.margins.gutter.toString();
      }
      children.push(XMLBuilder.wSelf('pgMar', attrs));
    }

    // Columns - output when set (including single column)
    if (this.properties.columns) {
      const attrs: Record<string, string> = {
        'w:num': this.properties.columns.count.toString(),
      };
      if (this.properties.columns.space !== undefined) {
        attrs['w:space'] = this.properties.columns.space.toString();
      }
      if (this.properties.columns.equalWidth !== undefined) {
        attrs['w:equalWidth'] = this.properties.columns.equalWidth ? '1' : '0';
      }
      children.push(XMLBuilder.wSelf('cols', attrs));
    }

    // Title page
    if (this.properties.titlePage) {
      children.push(XMLBuilder.wSelf('titlePg', { 'w:val': '1' }));
    }

    // Page numbering
    if (this.properties.pageNumbering) {
      const attrs: Record<string, string> = {};
      if (this.properties.pageNumbering.start !== undefined) {
        attrs['w:start'] = this.properties.pageNumbering.start.toString();
      }
      if (this.properties.pageNumbering.format) {
        attrs['w:fmt'] = this.properties.pageNumbering.format;
      }
      if (Object.keys(attrs).length > 0) {
        children.push(XMLBuilder.wSelf('pgNumType', attrs));
      }
    }

    return XMLBuilder.w('sectPr', undefined, children);
  }

  /**
   * Creates a section with default properties
   */
  static create(properties?: SectionProperties): Section {
    return new Section(properties);
  }

  /**
   * Creates a letter-sized section (8.5" x 11")
   */
  static createLetter(): Section {
    return new Section({
      pageSize: {
        width: PAGE_SIZES.LETTER.width,
        height: PAGE_SIZES.LETTER.height,
        orientation: 'portrait',
      },
    });
  }

  /**
   * Creates an A4-sized section (21cm x 29.7cm)
   */
  static createA4(): Section {
    return new Section({
      pageSize: {
        width: PAGE_SIZES.A4.width,
        height: PAGE_SIZES.A4.height,
        orientation: 'portrait',
      },
    });
  }

  /**
   * Creates a landscape section
   * @param pageSize Page size (default: Letter)
   */
  static createLandscape(pageSize: 'letter' | 'a4' = 'letter'): Section {
    const size = pageSize === 'a4' ? PAGE_SIZES.A4 : PAGE_SIZES.LETTER;
    return new Section({
      pageSize: {
        width: size.height,  // Swap for landscape
        height: size.width,
        orientation: 'landscape',
      },
    });
  }
}
