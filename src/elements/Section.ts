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
/**
 * Per ECMA-376 Part 1 §17.18.59 (ST_NumberFormat)
 * Comprehensive set of page number formats
 */
export type PageNumberFormat =
  | 'decimal'
  | 'lowerRoman'
  | 'upperRoman'
  | 'lowerLetter'
  | 'upperLetter'
  | 'ordinal'
  | 'cardinalText'
  | 'ordinalText'
  | 'hex'
  | 'chicago'
  | 'decimalZero'
  | 'decimalEnclosedCircle'
  | 'decimalEnclosedFullstop'
  | 'decimalEnclosedParen'
  | 'ideographDigital'
  | 'ideographTraditional'
  | 'ideographLegalTraditional'
  | 'ideographEnclosedCircle'
  | 'ideographZodiac'
  | 'ideographZodiacTraditional'
  | 'chineseCounting'
  | 'chineseCountingThousand'
  | 'chineseLegalSimplified'
  | 'japaneseCounting'
  | 'japaneseDigitalTenThousand'
  | 'japaneseLegal'
  | 'koreanCounting'
  | 'koreanDigital'
  | 'koreanDigital2'
  | 'koreanLegal'
  | 'taiwaneseCounting'
  | 'taiwaneseCountingThousand'
  | 'taiwaneseDigital'
  | 'aiueo'
  | 'aiueoFullWidth'
  | 'iroha'
  | 'irohaFullWidth'
  | 'arabicAbjad'
  | 'arabicAlpha'
  | 'hebrew1'
  | 'hebrew2'
  | 'hindiConsonants'
  | 'hindiCounting'
  | 'hindiNumbers'
  | 'hindiVowels'
  | 'thaiCounting'
  | 'thaiLetters'
  | 'thaiNumbers'
  | 'vietnameseCounting'
  | 'russianLower'
  | 'russianUpper'
  | 'numberInDash'
  | 'dollarText'
  | 'bullet'
  | 'bahtText'
  | 'ganada'
  | 'chosung'
  | 'none'
  | string; // Allow any format string for forward compatibility

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
  /** Show column separator line */
  separator?: boolean;
  /** Individual column widths (for unequal columns) in twips */
  columnWidths?: number[];
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
 * Paper source (printer tray) properties
 */
export interface PaperSource {
  /** First page tray number */
  first?: number;
  /** Other pages tray number */
  other?: number;
}

/**
 * Vertical page alignment
 */
export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'both';

/**
 * Text direction
 */
export type TextDirection = 'ltr' | 'rtl' | 'tbRl' | 'btLr';

/**
 * Document grid type for East Asian typography
 */
export type DocumentGridType = 'default' | 'lines' | 'linesAndChars' | 'snapToChars';

/**
 * Document grid properties
 */
export interface DocumentGrid {
  /** Grid type */
  type?: DocumentGridType;
  /** Lines per page */
  linePitch?: number;
  /** Characters per line */
  charSpace?: number;
}

/**
 * Line numbering restart mode
 */
export type LineNumberingRestart = 'newPage' | 'newSection' | 'continuous';

/**
 * Line numbering properties
 * Per ECMA-376 Part 1, Section 17.6.8 (w:lnNumType element)
 */
export interface LineNumbering {
  /** Display line number every N lines (countBy attribute) */
  countBy?: number;
  /** Starting line number */
  start?: number;
  /** Distance from text margin in twips */
  distance?: number;
  /** When to restart line numbering */
  restart?: LineNumberingRestart;
}

/**
 * Footnote/endnote position
 * Per ECMA-376 Part 1 §17.11.6 and §17.11.7
 */
export type NotePosition = 'pageBottom' | 'beneathText' | 'sectEnd' | 'docEnd';

/**
 * Note numbering restart mode
 */
export type NoteNumberRestart = 'continuous' | 'eachSect' | 'eachPage';

/**
 * Chapter separator character for page numbering
 * Per ECMA-376 Part 1 §17.18.5
 */
export type ChapterSeparator = 'colon' | 'emDash' | 'enDash' | 'hyphen' | 'period';

/**
 * Footnote/endnote section-level properties
 */
export interface NoteProperties {
  /** Note positioning */
  position?: NotePosition;
  /** Number format */
  numberFormat?: PageNumberFormat;
  /** Starting number */
  startNumber?: number;
  /** Restart numbering mode */
  restart?: NoteNumberRestart;
}

/**
 * Page border definition for a single side
 * Per ECMA-376 Part 1 §17.6.10
 */
export interface PageBorderDef {
  /** Border style (single, double, dashed, etc.) */
  style?: string;
  /** Border size in eighths of a point */
  size?: number;
  /** Border color in hex format (without #) */
  color?: string;
  /** Space between border and page edge/text in points */
  space?: number;
  /** Whether to show shadow */
  shadow?: boolean;
  /** Whether to include frame around page */
  frame?: boolean;
  /** Theme color reference */
  themeColor?: string;
  /** Art border ID (for decorative borders) */
  artId?: number;
}

/**
 * Page borders configuration
 * Per ECMA-376 Part 1 §17.6.10
 */
export interface PageBorders {
  /** Top page border */
  top?: PageBorderDef;
  /** Bottom page border */
  bottom?: PageBorderDef;
  /** Left page border */
  left?: PageBorderDef;
  /** Right page border */
  right?: PageBorderDef;
  /** Whether border is measured from page edge or text edge */
  offsetFrom?: 'page' | 'text';
  /** Display on all pages, first page only, or not first page */
  display?: 'allPages' | 'firstPage' | 'notFirstPage';
  /** Z-ordering relative to text */
  zOrder?: 'front' | 'back';
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
    default?: string; // rId for default header
    first?: string; // rId for first page header
    even?: string; // rId for even page header
  };
  /** Footer reference IDs */
  footers?: {
    default?: string; // rId for default footer
    first?: string; // rId for first page footer
    even?: string; // rId for even page footer
  };
  /** Title page (different first page) */
  titlePage?: boolean;
  /** Vertical page alignment */
  verticalAlignment?: VerticalAlignment;
  /** Paper source (printer tray) */
  paperSource?: PaperSource;
  /** Text direction (LTR/RTL support) */
  textDirection?: TextDirection;
  /** Right-to-left section layout */
  bidi?: boolean;
  /** Gutter on right side (for RTL) */
  rtlGutter?: boolean;
  /** Document grid for snapping text to grid */
  docGrid?: DocumentGrid;
  /** Line numbering configuration */
  lineNumbering?: LineNumbering;
  /** Section-level footnote properties (w:footnotePr) */
  footnotePr?: NoteProperties;
  /** Section-level endnote properties (w:endnotePr) */
  endnotePr?: NoteProperties;
  /** Suppress endnotes in this section (w:noEndnote) */
  noEndnote?: boolean;
  /** Form protection for this section (w:formProt) */
  formProt?: boolean;
  /** Page borders per ECMA-376 Part 1 §17.6.10 */
  pageBorders?: PageBorders;
  /** Printer settings relationship ID (w:printerSettings r:id) */
  printerSettingsId?: string;
  /** Chapter style heading level for page numbering (w:pgNumType w:chapStyle) */
  chapStyle?: number;
  /** Chapter separator for page numbering (w:pgNumType w:chapSep) */
  chapSep?: ChapterSeparator;
}

/**
 * Represents a document section
 */
/**
 * Section property change tracking (w:sectPrChange)
 * Per ECMA-376 Part 1 §17.13.5.32
 */
export interface SectPrChange {
  author: string;
  date: string;
  id: string;
  previousProperties: Record<string, any>;
}

export class Section {
  private properties: SectionProperties;
  /** Tracking context for automatic change tracking */
  private trackingContext?: import('../tracking/TrackingContext').TrackingContext;
  /** Section property change tracking (w:sectPrChange) */
  private sectPrChange?: SectPrChange;

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
        top: 1440, // 1 inch
        bottom: 1440,
        left: 1440,
        right: 1440,
        header: 720, // 0.5 inch
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
      // Phase 4.5 - New properties
      verticalAlignment: properties.verticalAlignment,
      paperSource: properties.paperSource,
      textDirection: properties.textDirection,
      pageBorders: properties.pageBorders,
    };
  }

  /**
   * Sets the tracking context for automatic change tracking.
   * Called by Document when track changes is enabled.
   * @internal
   */
  _setTrackingContext(context: import('../tracking/TrackingContext').TrackingContext): void {
    this.trackingContext = context;
  }

  /**
   * Gets the section property change tracking info
   */
  getSectPrChange(): SectPrChange | undefined {
    return this.sectPrChange;
  }

  /**
   * Sets the section property change tracking info
   */
  setSectPrChange(change: SectPrChange | undefined): void {
    this.sectPrChange = change;
  }

  /**
   * Clears the section property change tracking
   */
  clearSectPrChange(): void {
    this.sectPrChange = undefined;
  }

  /**
   * Gets the section properties
   */
  getProperties(): SectionProperties {
    return { ...this.properties };
  }

  // ============================================================================
  // Individual Property Getters
  // ============================================================================

  /**
   * Gets the page size settings
   * @returns Page size object with width, height, and orientation, or undefined
   */
  getPageSize(): PageSize | undefined {
    return this.properties.pageSize ? { ...this.properties.pageSize } : undefined;
  }

  /**
   * Gets the page orientation
   * @returns 'portrait' or 'landscape', or undefined if not set
   */
  getOrientation(): PageOrientation | undefined {
    return this.properties.pageSize?.orientation;
  }

  /**
   * Gets the margin settings
   * @returns Margins object or undefined if not set
   */
  getMargins(): Margins | undefined {
    return this.properties.margins ? { ...this.properties.margins } : undefined;
  }

  /**
   * Gets the column layout settings
   * @returns Columns object or undefined if not set
   */
  getColumns(): Columns | undefined {
    return this.properties.columns ? { ...this.properties.columns } : undefined;
  }

  /**
   * Gets the section type (break type)
   * @returns Section type or undefined if not set
   */
  getSectionType(): SectionType | undefined {
    return this.properties.type;
  }

  /**
   * Gets the page numbering settings
   * @returns PageNumbering object or undefined if not set
   */
  getPageNumbering(): PageNumbering | undefined {
    return this.properties.pageNumbering ? { ...this.properties.pageNumbering } : undefined;
  }

  /**
   * Gets whether this section has a title page (different first page)
   * @returns true if title page is enabled, false otherwise
   */
  getTitlePage(): boolean {
    return this.properties.titlePage ?? false;
  }

  /**
   * Gets header references
   * @returns Object with default, first, and even header relationship IDs, or undefined
   */
  getHeaderReferences(): { default?: string; first?: string; even?: string } | undefined {
    return this.properties.headers ? { ...this.properties.headers } : undefined;
  }

  /**
   * Gets a specific header reference
   * @param type Header type (default, first, even)
   * @returns Relationship ID or undefined if not set
   */
  getHeaderReference(type: 'default' | 'first' | 'even'): string | undefined {
    return this.properties.headers?.[type];
  }

  /**
   * Gets footer references
   * @returns Object with default, first, and even footer relationship IDs, or undefined
   */
  getFooterReferences(): { default?: string; first?: string; even?: string } | undefined {
    return this.properties.footers ? { ...this.properties.footers } : undefined;
  }

  /**
   * Gets a specific footer reference
   * @param type Footer type (default, first, even)
   * @returns Relationship ID or undefined if not set
   */
  getFooterReference(type: 'default' | 'first' | 'even'): string | undefined {
    return this.properties.footers?.[type];
  }

  /**
   * Gets the vertical page alignment
   * @returns Vertical alignment or undefined if not set
   */
  getVerticalAlignment(): VerticalAlignment | undefined {
    return this.properties.verticalAlignment;
  }

  /**
   * Gets the paper source (printer tray) settings
   * @returns PaperSource object or undefined if not set
   */
  getPaperSource(): PaperSource | undefined {
    return this.properties.paperSource ? { ...this.properties.paperSource } : undefined;
  }

  /**
   * Gets whether column separator is enabled
   * @returns true if separator is enabled, false otherwise
   */
  getColumnSeparator(): boolean {
    return this.properties.columns?.separator ?? false;
  }

  /**
   * Gets custom column widths
   * @returns Array of column widths in twips, or undefined if not set
   */
  getColumnWidths(): number[] | undefined {
    return this.properties.columns?.columnWidths
      ? [...this.properties.columns.columnWidths]
      : undefined;
  }

  /**
   * Gets the text direction
   * @returns Text direction or undefined if not set
   */
  getTextDirection(): TextDirection | undefined {
    return this.properties.textDirection;
  }

  /**
   * Gets whether the section is bidirectional (RTL)
   * @returns true if bidi is enabled, false otherwise
   */
  getBidi(): boolean {
    return this.properties.bidi ?? false;
  }

  /**
   * Gets whether RTL gutter is enabled (gutter on right side)
   * @returns true if RTL gutter is enabled, false otherwise
   */
  getRtlGutter(): boolean {
    return this.properties.rtlGutter ?? false;
  }

  /**
   * Gets the document grid settings
   * @returns DocumentGrid object or undefined if not set
   */
  getDocGrid(): DocumentGrid | undefined {
    return this.properties.docGrid ? { ...this.properties.docGrid } : undefined;
  }

  // ============================================================================
  // Setters
  // ============================================================================

  /**
   * Sets page size
   * @param width Width in twips
   * @param height Height in twips
   * @param orientation Page orientation
   */
  setPageSize(width: number, height: number, orientation: PageOrientation = 'portrait'): this {
    const prev = this.properties.pageSize ? { ...this.properties.pageSize } : undefined;
    this.properties.pageSize = { width, height, orientation };
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(this, 'pageSize', prev, this.properties.pageSize);
    }
    return this;
  }

  /**
   * Sets page orientation
   * @param orientation Page orientation
   */
  setOrientation(orientation: PageOrientation): this {
    const prev = this.properties.pageSize?.orientation;
    if (!this.properties.pageSize) {
      this.properties.pageSize = {
        width: PAGE_SIZES.LETTER.width,
        height: PAGE_SIZES.LETTER.height,
      };
    }
    this.properties.pageSize.orientation = orientation;

    // Swap width/height for landscape
    if (
      orientation === 'landscape' &&
      this.properties.pageSize.width < this.properties.pageSize.height
    ) {
      const temp = this.properties.pageSize.width;
      this.properties.pageSize.width = this.properties.pageSize.height;
      this.properties.pageSize.height = temp;
    }

    if (this.trackingContext?.isEnabled() && prev !== orientation) {
      this.trackingContext.trackSectionChange(this, 'orientation', prev, orientation);
    }
    return this;
  }

  /**
   * Sets margins
   * @param margins Margin properties
   */
  setMargins(margins: Margins): this {
    const prev = this.properties.margins ? { ...this.properties.margins } : undefined;
    const existing = this.properties.margins;
    this.properties.margins = {
      top: margins.top,
      bottom: margins.bottom,
      left: margins.left,
      right: margins.right,
      header: margins.header ?? existing?.header ?? 720,
      footer: margins.footer ?? existing?.footer ?? 720,
      gutter: margins.gutter ?? existing?.gutter,
    };
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(this, 'margins', prev, this.properties.margins);
    }
    return this;
  }

  /**
   * Sets column layout
   * @param count Number of columns
   * @param space Space between columns in twips
   */
  setColumns(count: number, space = 720): this {
    const prev = this.properties.columns ? { ...this.properties.columns } : undefined;
    this.properties.columns = {
      count,
      space,
      equalWidth: true,
    };
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(this, 'columns', prev, this.properties.columns);
    }
    return this;
  }

  /**
   * Sets section type
   * @param type Section break type
   */
  setSectionType(type: SectionType): this {
    const prev = this.properties.type;
    this.properties.type = type;
    if (this.trackingContext?.isEnabled() && prev !== type) {
      this.trackingContext.trackSectionChange(this, 'type', prev, type);
    }
    return this;
  }

  /**
   * Sets page numbering
   * @param start Starting page number
   * @param format Number format
   */
  setPageNumbering(start?: number, format?: PageNumberFormat): this {
    const prev = this.properties.pageNumbering ? { ...this.properties.pageNumbering } : undefined;
    this.properties.pageNumbering = { start, format };
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(
        this,
        'pageNumbering',
        prev,
        this.properties.pageNumbering
      );
    }
    return this;
  }

  /**
   * Sets title page flag (different first page)
   * @param titlePage Whether this section has a different first page
   */
  setTitlePage(titlePage = true): this {
    const prev = this.properties.titlePage;
    this.properties.titlePage = titlePage;
    if (this.trackingContext?.isEnabled() && prev !== titlePage) {
      this.trackingContext.trackSectionChange(this, 'titlePage', prev, titlePage);
    }
    return this;
  }

  /**
   * Sets header reference
   * @param type Header type (default, first, even)
   * @param rId Relationship ID
   */
  setHeaderReference(type: 'default' | 'first' | 'even', rId: string): this {
    const prev = this.properties.headers?.[type];
    if (!this.properties.headers) {
      this.properties.headers = {};
    }
    this.properties.headers[type] = rId;
    if (this.trackingContext?.isEnabled() && prev !== rId) {
      this.trackingContext.trackSectionChange(this, `headerReference:${type}`, prev, rId);
    }
    return this;
  }

  /**
   * Sets footer reference
   * @param type Footer type (default, first, even)
   * @param rId Relationship ID
   */
  setFooterReference(type: 'default' | 'first' | 'even', rId: string): this {
    const prev = this.properties.footers?.[type];
    if (!this.properties.footers) {
      this.properties.footers = {};
    }
    this.properties.footers[type] = rId;
    if (this.trackingContext?.isEnabled() && prev !== rId) {
      this.trackingContext.trackSectionChange(this, `footerReference:${type}`, prev, rId);
    }
    return this;
  }

  /**
   * Sets vertical page alignment
   * Controls how content is vertically aligned on the page
   * @param alignment Vertical alignment (top, center, bottom, both=justified)
   */
  setVerticalAlignment(alignment: VerticalAlignment): this {
    const prev = this.properties.verticalAlignment;
    this.properties.verticalAlignment = alignment;
    if (this.trackingContext?.isEnabled() && prev !== alignment) {
      this.trackingContext.trackSectionChange(this, 'verticalAlignment', prev, alignment);
    }
    return this;
  }

  /**
   * Sets paper source (printer tray selection)
   * @param first First page tray number
   * @param other Other pages tray number
   */
  setPaperSource(first?: number, other?: number): this {
    const prev = this.properties.paperSource ? { ...this.properties.paperSource } : undefined;
    this.properties.paperSource = { first, other };
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(
        this,
        'paperSource',
        prev,
        this.properties.paperSource
      );
    }
    return this;
  }

  /**
   * Sets column separator line
   * Shows a vertical line between columns
   * @param separator Whether to show column separator line
   */
  setColumnSeparator(separator = true): this {
    const prev = this.properties.columns?.separator;
    if (!this.properties.columns) {
      this.properties.columns = { count: 1 };
    }
    this.properties.columns.separator = separator;
    if (this.trackingContext?.isEnabled() && prev !== separator) {
      this.trackingContext.trackSectionChange(this, 'columnSeparator', prev, separator);
    }
    return this;
  }

  /**
   * Sets custom column widths (for unequal columns)
   * @param widths Array of column widths in twips
   */
  setColumnWidths(widths: number[]): this {
    const prev = this.properties.columns ? { ...this.properties.columns } : undefined;
    if (!this.properties.columns) {
      this.properties.columns = { count: widths.length };
    }
    this.properties.columns.columnWidths = widths;
    this.properties.columns.equalWidth = false;
    this.properties.columns.count = widths.length;
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(this, 'columns', prev, {
        ...this.properties.columns,
      });
    }
    return this;
  }

  /**
   * Sets text direction for the section
   * @param direction Text direction (ltr=left-to-right, rtl=right-to-left, tbRl=top-to-bottom-right-to-left, btLr=bottom-to-top-left-to-right)
   */
  setTextDirection(direction: TextDirection): this {
    const prev = this.properties.textDirection;
    this.properties.textDirection = direction;
    if (this.trackingContext?.isEnabled() && prev !== direction) {
      this.trackingContext.trackSectionChange(this, 'textDirection', prev, direction);
    }
    return this;
  }

  /**
   * Sets line numbering for the section
   * Per ECMA-376 Part 1, Section 17.6.8 (w:lnNumType)
   * @param options Line numbering configuration
   */
  setLineNumbering(options: LineNumbering): this {
    const prev = this.properties.lineNumbering ? { ...this.properties.lineNumbering } : undefined;
    this.properties.lineNumbering = { ...options };
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(
        this,
        'lineNumbering',
        prev,
        this.properties.lineNumbering
      );
    }
    return this;
  }

  /**
   * Gets line numbering configuration
   * @returns Line numbering settings or undefined if not set
   */
  getLineNumbering(): LineNumbering | undefined {
    return this.properties.lineNumbering ? { ...this.properties.lineNumbering } : undefined;
  }

  /**
   * Clears line numbering for the section
   */
  clearLineNumbering(): this {
    this.properties.lineNumbering = undefined;
    return this;
  }

  /**
   * Sets section-level footnote properties
   * Per ECMA-376 Part 1 §17.11.6
   */
  setFootnoteProperties(props: NoteProperties): this {
    const prev = this.properties.footnotePr ? { ...this.properties.footnotePr } : undefined;
    this.properties.footnotePr = { ...props };
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(this, 'footnotePr', prev, this.properties.footnotePr);
    }
    return this;
  }

  /**
   * Sets section-level endnote properties
   * Per ECMA-376 Part 1 §17.11.7
   */
  setEndnoteProperties(props: NoteProperties): this {
    const prev = this.properties.endnotePr ? { ...this.properties.endnotePr } : undefined;
    this.properties.endnotePr = { ...props };
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(this, 'endnotePr', prev, this.properties.endnotePr);
    }
    return this;
  }

  /**
   * Sets whether endnotes are suppressed in this section
   * Per ECMA-376 Part 1 §17.6.14
   */
  setNoEndnote(noEndnote = true): this {
    const prev = this.properties.noEndnote;
    this.properties.noEndnote = noEndnote;
    if (this.trackingContext?.isEnabled() && prev !== noEndnote) {
      this.trackingContext.trackSectionChange(this, 'noEndnote', prev, noEndnote);
    }
    return this;
  }

  /**
   * Sets form protection for this section
   * Per ECMA-376 Part 1 §17.6.4
   */
  setFormProtection(formProt = true): this {
    const prev = this.properties.formProt;
    this.properties.formProt = formProt;
    if (this.trackingContext?.isEnabled() && prev !== formProt) {
      this.trackingContext.trackSectionChange(this, 'formProt', prev, formProt);
    }
    return this;
  }

  /**
   * Sets printer settings relationship ID
   * Per ECMA-376 Part 1 §17.6.6
   */
  setPrinterSettings(rId: string): this {
    const prev = this.properties.printerSettingsId;
    this.properties.printerSettingsId = rId;
    if (this.trackingContext?.isEnabled() && prev !== rId) {
      this.trackingContext.trackSectionChange(this, 'printerSettings', prev, rId);
    }
    return this;
  }

  /**
   * Sets chapter numbering for page numbers
   * Per ECMA-376 Part 1 §17.6.12
   * @param chapStyle - Heading level (1-9) to use for chapter numbering
   * @param chapSep - Separator between chapter and page number
   */
  setChapterNumbering(chapStyle: number, chapSep?: ChapterSeparator): this {
    const prevStyle = this.properties.chapStyle;
    const prevSep = this.properties.chapSep;
    this.properties.chapStyle = chapStyle;
    if (chapSep) this.properties.chapSep = chapSep;
    if (this.trackingContext?.isEnabled()) {
      this.trackingContext.trackSectionChange(
        this,
        'chapterNumbering',
        { chapStyle: prevStyle, chapSep: prevSep },
        { chapStyle, chapSep: chapSep || prevSep }
      );
    }
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

    // CT_SectPr element order per ECMA-376:
    // footnotePr → endnotePr → type → pgSz → pgMar → paperSrc → pgBorders →
    // lnNumType → pgNumType → cols → formProt → vAlign → noEndnote → titlePg →
    // textDirection → bidi → rtlGutter → docGrid → printerSettings → sectPrChange

    // Footnote properties (w:footnotePr)
    if (this.properties.footnotePr) {
      const fnChildren: XMLElement[] = [];
      if (this.properties.footnotePr.position) {
        fnChildren.push(XMLBuilder.wSelf('pos', { 'w:val': this.properties.footnotePr.position }));
      }
      if (this.properties.footnotePr.numberFormat) {
        fnChildren.push(
          XMLBuilder.wSelf('numFmt', { 'w:val': this.properties.footnotePr.numberFormat })
        );
      }
      if (this.properties.footnotePr.startNumber !== undefined) {
        fnChildren.push(
          XMLBuilder.wSelf('numStart', {
            'w:val': this.properties.footnotePr.startNumber.toString(),
          })
        );
      }
      if (this.properties.footnotePr.restart) {
        fnChildren.push(
          XMLBuilder.wSelf('numRestart', { 'w:val': this.properties.footnotePr.restart })
        );
      }
      if (fnChildren.length > 0) {
        children.push(XMLBuilder.w('footnotePr', undefined, fnChildren));
      }
    }

    // Endnote properties (w:endnotePr)
    if (this.properties.endnotePr) {
      const enChildren: XMLElement[] = [];
      if (this.properties.endnotePr.position) {
        enChildren.push(XMLBuilder.wSelf('pos', { 'w:val': this.properties.endnotePr.position }));
      }
      if (this.properties.endnotePr.numberFormat) {
        enChildren.push(
          XMLBuilder.wSelf('numFmt', { 'w:val': this.properties.endnotePr.numberFormat })
        );
      }
      if (this.properties.endnotePr.startNumber !== undefined) {
        enChildren.push(
          XMLBuilder.wSelf('numStart', {
            'w:val': this.properties.endnotePr.startNumber.toString(),
          })
        );
      }
      if (this.properties.endnotePr.restart) {
        enChildren.push(
          XMLBuilder.wSelf('numRestart', { 'w:val': this.properties.endnotePr.restart })
        );
      }
      if (enChildren.length > 0) {
        children.push(XMLBuilder.w('endnotePr', undefined, enChildren));
      }
    }

    // Section type
    if (this.properties.type) {
      children.push(XMLBuilder.wSelf('type', { 'w:val': this.properties.type }));
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
      attrs['w:header'] = (this.properties.margins.header ?? 720).toString();
      attrs['w:footer'] = (this.properties.margins.footer ?? 720).toString();
      if (this.properties.margins.gutter !== undefined) {
        attrs['w:gutter'] = this.properties.margins.gutter.toString();
      }
      children.push(XMLBuilder.wSelf('pgMar', attrs));
    }

    // Paper source
    if (this.properties.paperSource) {
      const attrs: Record<string, string> = {};
      if (this.properties.paperSource.first !== undefined) {
        attrs['w:first'] = this.properties.paperSource.first.toString();
      }
      if (this.properties.paperSource.other !== undefined) {
        attrs['w:other'] = this.properties.paperSource.other.toString();
      }
      if (Object.keys(attrs).length > 0) {
        children.push(XMLBuilder.wSelf('paperSrc', attrs));
      }
    }

    // Page borders per ECMA-376 Part 1 §17.6.10
    if (this.properties.pageBorders) {
      const pgBorders = this.properties.pageBorders;
      const pgBordersAttrs: Record<string, string> = {};
      if (pgBorders.offsetFrom) pgBordersAttrs['w:offsetFrom'] = pgBorders.offsetFrom;
      if (pgBorders.display) pgBordersAttrs['w:display'] = pgBorders.display;
      if (pgBorders.zOrder) pgBordersAttrs['w:zOrder'] = pgBorders.zOrder;

      const borderChildren: XMLElement[] = [];
      const buildBorder = (side: string, def: PageBorderDef) => {
        const bAttrs: Record<string, string | number> = {};
        if (def.style) bAttrs['w:val'] = def.style;
        // ECMA-376 Part 1 §17.18.2: sz valid range 2-96 (eighths of a point)
        if (def.size !== undefined) bAttrs['w:sz'] = Math.max(2, Math.min(96, def.size));
        if (def.color) bAttrs['w:color'] = def.color;
        // ECMA-376 Part 1 §17.18.88: space valid range 0-31680 (points)
        if (def.space !== undefined) bAttrs['w:space'] = Math.max(0, Math.min(31680, def.space));
        if (def.shadow) bAttrs['w:shadow'] = '1';
        if (def.frame) bAttrs['w:frame'] = '1';
        if (def.themeColor) bAttrs['w:themeColor'] = def.themeColor;
        if (def.artId !== undefined) bAttrs['w:id'] = def.artId;
        borderChildren.push(XMLBuilder.wSelf(side, bAttrs));
      };

      if (pgBorders.top) buildBorder('top', pgBorders.top);
      if (pgBorders.left) buildBorder('left', pgBorders.left);
      if (pgBorders.bottom) buildBorder('bottom', pgBorders.bottom);
      if (pgBorders.right) buildBorder('right', pgBorders.right);

      children.push(XMLBuilder.w('pgBorders', pgBordersAttrs, borderChildren));
    }

    // Line numbering (w:lnNumType)
    if (this.properties.lineNumbering) {
      const attrs: Record<string, string> = {};
      if (this.properties.lineNumbering.countBy !== undefined) {
        attrs['w:countBy'] = this.properties.lineNumbering.countBy.toString();
      }
      if (this.properties.lineNumbering.start !== undefined) {
        attrs['w:start'] = this.properties.lineNumbering.start.toString();
      }
      if (this.properties.lineNumbering.distance !== undefined) {
        attrs['w:distance'] = this.properties.lineNumbering.distance.toString();
      }
      if (this.properties.lineNumbering.restart) {
        attrs['w:restart'] = this.properties.lineNumbering.restart;
      }
      if (Object.keys(attrs).length > 0) {
        children.push(XMLBuilder.wSelf('lnNumType', attrs));
      }
    }

    // Page numbering
    if (this.properties.pageNumbering || this.properties.chapStyle !== undefined) {
      const attrs: Record<string, string> = {};
      if (this.properties.pageNumbering?.start !== undefined) {
        attrs['w:start'] = this.properties.pageNumbering.start.toString();
      }
      if (this.properties.pageNumbering?.format) {
        attrs['w:fmt'] = this.properties.pageNumbering.format;
      }
      if (this.properties.chapStyle !== undefined) {
        attrs['w:chapStyle'] = this.properties.chapStyle.toString();
      }
      if (this.properties.chapSep) {
        attrs['w:chapSep'] = this.properties.chapSep;
      }
      if (Object.keys(attrs).length > 0) {
        children.push(XMLBuilder.wSelf('pgNumType', attrs));
      }
    }

    // Columns
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
      if (this.properties.columns.separator !== undefined) {
        attrs['w:sep'] = this.properties.columns.separator ? '1' : '0';
      }

      const colChildren: XMLElement[] = [];
      if (this.properties.columns.columnWidths) {
        for (const width of this.properties.columns.columnWidths) {
          colChildren.push(XMLBuilder.wSelf('col', { 'w:w': width.toString() }));
        }
      }

      children.push(
        colChildren.length > 0
          ? XMLBuilder.w('cols', attrs, colChildren)
          : XMLBuilder.wSelf('cols', attrs)
      );
    }

    // Form protection (w:formProt)
    if (this.properties.formProt) {
      children.push(XMLBuilder.wSelf('formProt'));
    }

    // Vertical alignment
    if (this.properties.verticalAlignment) {
      children.push(XMLBuilder.wSelf('vAlign', { 'w:val': this.properties.verticalAlignment }));
    }

    // Suppress endnotes (w:noEndnote)
    if (this.properties.noEndnote) {
      children.push(XMLBuilder.wSelf('noEndnote'));
    }

    // Title page
    if (this.properties.titlePage) {
      children.push(XMLBuilder.wSelf('titlePg', { 'w:val': '1' }));
    }

    // Text direction (map to valid ST_TextDirection values per ECMA-376 §17.18.93)
    if (this.properties.textDirection) {
      const textDirMap: Record<string, string> = {
        ltr: 'lrTb',
        rtl: 'tbRl',
        tbRl: 'tbRl',
        btLr: 'btLr',
        lrTb: 'lrTb',
        lrTbV: 'lrTbV',
        tbRlV: 'tbRlV',
        tbLrV: 'tbLrV',
      };
      const val = textDirMap[this.properties.textDirection] || this.properties.textDirection;
      children.push(XMLBuilder.wSelf('textDirection', { 'w:val': val }));
    }

    // Bidirectional section (RTL)
    if (this.properties.bidi) {
      children.push(XMLBuilder.wSelf('bidi'));
    }

    // RTL gutter (gutter on right side)
    if (this.properties.rtlGutter) {
      children.push(XMLBuilder.wSelf('rtlGutter'));
    }

    // Document grid
    if (this.properties.docGrid) {
      const attrs: Record<string, string> = {};
      if (this.properties.docGrid.type) {
        attrs['w:type'] = this.properties.docGrid.type;
      }
      if (this.properties.docGrid.linePitch !== undefined) {
        attrs['w:linePitch'] = this.properties.docGrid.linePitch.toString();
      }
      if (this.properties.docGrid.charSpace !== undefined) {
        attrs['w:charSpace'] = this.properties.docGrid.charSpace.toString();
      }
      if (Object.keys(attrs).length > 0) {
        children.push(XMLBuilder.wSelf('docGrid', attrs));
      }
    }

    // Printer settings (w:printerSettings)
    if (this.properties.printerSettingsId) {
      children.push(
        XMLBuilder.wSelf('printerSettings', {
          'r:id': this.properties.printerSettingsId,
        })
      );
    }

    // Add section property change (w:sectPrChange) per ECMA-376 Part 1 §17.13.5.32
    // Must be last child of w:sectPr
    if (this.sectPrChange) {
      const changeAttrs: Record<string, string | number> = {
        'w:id': this.sectPrChange.id,
        'w:author': this.sectPrChange.author,
        'w:date': this.sectPrChange.date,
      };
      const prevChildren: XMLElement[] = [];
      const prev = this.sectPrChange.previousProperties;
      if (prev) {
        // Ordered per CT_SectPrBase:
        // type → pgSz → pgMar → lnNumType → pgNumType → cols → formProt → vAlign → titlePg → textDirection
        if (prev.type) {
          prevChildren.push(XMLBuilder.wSelf('type', { 'w:val': prev.type }));
        }
        if (prev.pageSize) {
          const pgSzAttrs: Record<string, string> = {
            'w:w': prev.pageSize.width?.toString() || '12240',
            'w:h': prev.pageSize.height?.toString() || '15840',
          };
          if (prev.pageSize.orientation === 'landscape') {
            pgSzAttrs['w:orient'] = 'landscape';
          }
          prevChildren.push(XMLBuilder.wSelf('pgSz', pgSzAttrs));
        }
        if (prev.margins) {
          const pgMarAttrs: Record<string, string> = {};
          if (prev.margins.top !== undefined) pgMarAttrs['w:top'] = prev.margins.top.toString();
          if (prev.margins.bottom !== undefined)
            pgMarAttrs['w:bottom'] = prev.margins.bottom.toString();
          if (prev.margins.left !== undefined) pgMarAttrs['w:left'] = prev.margins.left.toString();
          if (prev.margins.right !== undefined)
            pgMarAttrs['w:right'] = prev.margins.right.toString();
          if (prev.margins.header !== undefined)
            pgMarAttrs['w:header'] = prev.margins.header.toString();
          if (prev.margins.footer !== undefined)
            pgMarAttrs['w:footer'] = prev.margins.footer.toString();
          prevChildren.push(XMLBuilder.wSelf('pgMar', pgMarAttrs));
        }
        if (prev.lineNumbering) {
          const lnAttrs: Record<string, string> = {};
          if (prev.lineNumbering.countBy !== undefined)
            lnAttrs['w:countBy'] = prev.lineNumbering.countBy.toString();
          if (prev.lineNumbering.start !== undefined)
            lnAttrs['w:start'] = prev.lineNumbering.start.toString();
          if (prev.lineNumbering.restart) lnAttrs['w:restart'] = prev.lineNumbering.restart;
          if (prev.lineNumbering.distance !== undefined)
            lnAttrs['w:distance'] = prev.lineNumbering.distance.toString();
          if (Object.keys(lnAttrs).length > 0) {
            prevChildren.push(XMLBuilder.wSelf('lnNumType', lnAttrs));
          }
        }
        if (prev.pageNumbering) {
          const pnAttrs: Record<string, string> = {};
          if (prev.pageNumbering.start !== undefined)
            pnAttrs['w:start'] = prev.pageNumbering.start.toString();
          if (prev.pageNumbering.format) pnAttrs['w:fmt'] = prev.pageNumbering.format;
          if (Object.keys(pnAttrs).length > 0) {
            prevChildren.push(XMLBuilder.wSelf('pgNumType', pnAttrs));
          }
        }
        if (prev.columns) {
          const colAttrs: Record<string, string> = {
            'w:num': prev.columns.count?.toString() || '1',
          };
          if (prev.columns.space !== undefined) colAttrs['w:space'] = prev.columns.space.toString();
          prevChildren.push(XMLBuilder.wSelf('cols', colAttrs));
        }
        if (prev.formProt) {
          prevChildren.push(XMLBuilder.wSelf('formProt'));
        }
        if (prev.verticalAlignment) {
          prevChildren.push(XMLBuilder.wSelf('vAlign', { 'w:val': prev.verticalAlignment }));
        }
        if (prev.titlePage) {
          prevChildren.push(XMLBuilder.wSelf('titlePg'));
        }
        if (prev.textDirection) {
          const tdMap: Record<string, string> = {
            ltr: 'lrTb',
            rtl: 'tbRl',
            tbRl: 'tbRl',
            btLr: 'btLr',
            lrTb: 'lrTb',
            lrTbV: 'lrTbV',
            tbRlV: 'tbRlV',
            tbLrV: 'tbLrV',
          };
          prevChildren.push(
            XMLBuilder.wSelf('textDirection', {
              'w:val': tdMap[prev.textDirection] || prev.textDirection,
            })
          );
        }
      }
      const prevSectPr = XMLBuilder.w('sectPr', undefined, prevChildren);
      children.push(XMLBuilder.w('sectPrChange', changeAttrs, [prevSectPr]));
    }

    return XMLBuilder.w('sectPr', undefined, children);
  }

  /**
   * Creates a deep clone of this section
   * @returns New Section instance with copied properties
   */
  clone(): Section {
    // Deep clone all nested objects
    const clonedProperties: SectionProperties = {};

    if (this.properties.pageSize) {
      clonedProperties.pageSize = { ...this.properties.pageSize };
    }

    if (this.properties.margins) {
      clonedProperties.margins = { ...this.properties.margins };
    }

    if (this.properties.columns) {
      clonedProperties.columns = {
        ...this.properties.columns,
        columnWidths: this.properties.columns.columnWidths
          ? [...this.properties.columns.columnWidths]
          : undefined,
      };
    }

    if (this.properties.pageNumbering) {
      clonedProperties.pageNumbering = { ...this.properties.pageNumbering };
    }

    if (this.properties.headers) {
      clonedProperties.headers = { ...this.properties.headers };
    }

    if (this.properties.footers) {
      clonedProperties.footers = { ...this.properties.footers };
    }

    if (this.properties.paperSource) {
      clonedProperties.paperSource = { ...this.properties.paperSource };
    }

    if (this.properties.docGrid) {
      clonedProperties.docGrid = { ...this.properties.docGrid };
    }

    if (this.properties.lineNumbering) {
      clonedProperties.lineNumbering = { ...this.properties.lineNumbering };
    }

    // Copy primitive properties
    clonedProperties.type = this.properties.type;
    clonedProperties.titlePage = this.properties.titlePage;
    clonedProperties.verticalAlignment = this.properties.verticalAlignment;
    clonedProperties.textDirection = this.properties.textDirection;
    clonedProperties.bidi = this.properties.bidi;
    clonedProperties.rtlGutter = this.properties.rtlGutter;

    return new Section(clonedProperties);
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
        width: size.height, // Swap for landscape
        height: size.width,
        orientation: 'landscape',
      },
    });
  }
}
