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
  /** Paper size code per ECMA-376 ST_PaperSize (1=Letter, 9=A4, 5=Legal, etc.) */
  code?: number;
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
  /** Individual column spacings (space after each column) in twips — per ECMA-376 CT_Column */
  columnSpaces?: number[];
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
  /**
   * Theme tint (2-hex-digit string, ST_UcharHexNumber per §17.18.82) applied
   * to themeColor lookup. CT_TopBorder/CT_BottomBorder extend CT_Border so
   * inherit the full attribute set of §17.18.2.
   */
  themeTint?: string;
  /** Theme shade (2-hex-digit string) applied to themeColor lookup */
  themeShade?: string;
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
    // Spread the full input first so every SectionProperties field is preserved
    // (bidi, rtlGutter, docGrid, lineNumbering, footnotePr, endnotePr, noEndnote,
    // formProt, printerSettingsId, chapStyle, chapSep, etc.), then apply defaults
    // only for the handful of fields that must always have a value.
    // Before this change the constructor enumerated an allow-list and silently
    // dropped every other SectionProperties field, corrupting round-trip fidelity
    // for any source document using those properties.
    this.properties = {
      ...properties,
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
   * Gets per-column spacing (space after each column) in twips
   * @returns Array of column spacings, or undefined if not set
   */
  getColumnSpaces(): number[] | undefined {
    return this.properties.columns?.columnSpaces
      ? [...this.properties.columns.columnSpaces]
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
    const prevOrientation = this.properties.pageSize?.orientation;
    // Capture full pageSize BEFORE mutation for tracking
    const prevPageSize = this.properties.pageSize ? { ...this.properties.pageSize } : undefined;
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

    if (this.trackingContext?.isEnabled() && prevOrientation !== orientation) {
      // Track the full pageSize change (not just orientation) since width/height are swapped
      this.trackingContext.trackSectionChange(this, 'pageSize', prevPageSize, {
        ...this.properties.pageSize,
      });
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
      if (this.properties.pageSize.code !== undefined) {
        attrs['w:code'] = this.properties.pageSize.code.toString();
      }
      children.push(XMLBuilder.wSelf('pgSz', attrs));
    }

    // Margins — CT_PageMar §17.6.11 declares ALL seven attributes
    // (top/right/bottom/left/header/footer/gutter) as `use="required"`.
    // Emit every one so output is spec-compliant for strict validators
    // (the Open XML SDK validator is lenient here but third-party / XSD
    // tooling is not). header/footer default to 720 twips (0.5"), and
    // gutter to 0 (no gutter — the usual value for non-book-bound docs).
    if (this.properties.margins) {
      const attrs: Record<string, string> = {
        'w:top': this.properties.margins.top.toString(),
        'w:right': this.properties.margins.right.toString(),
        'w:bottom': this.properties.margins.bottom.toString(),
        'w:left': this.properties.margins.left.toString(),
        'w:header': (this.properties.margins.header ?? 720).toString(),
        'w:footer': (this.properties.margins.footer ?? 720).toString(),
        'w:gutter': (this.properties.margins.gutter ?? 0).toString(),
      };
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
        // CT_PageBorder §17.6.2 extends CT_Border (§17.18.2) — w:val is
        // REQUIRED. Default to "nil" when consumer set only size/color.
        const bAttrs: Record<string, string | number> = { 'w:val': def.style ?? 'nil' };
        // ECMA-376 Part 1 §17.18.2: sz valid range 2-96 (eighths of a point)
        if (def.size !== undefined) bAttrs['w:sz'] = Math.max(2, Math.min(96, def.size));
        if (def.color) bAttrs['w:color'] = def.color;
        // ECMA-376 Part 1 §17.18.88: space valid range 0-31680 (points)
        if (def.space !== undefined) bAttrs['w:space'] = Math.max(0, Math.min(31680, def.space));
        // shadow / frame are CT_OnOff — use `!== undefined` so explicit false
        // (w:shadow="0") round-trips without being dropped.
        if (def.shadow !== undefined) bAttrs['w:shadow'] = def.shadow ? '1' : '0';
        if (def.frame !== undefined) bAttrs['w:frame'] = def.frame ? '1' : '0';
        if (def.themeColor) bAttrs['w:themeColor'] = def.themeColor;
        if (def.themeTint) bAttrs['w:themeTint'] = def.themeTint;
        if (def.themeShade) bAttrs['w:themeShade'] = def.themeShade;
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
        const widths = this.properties.columns.columnWidths;
        const spaces = this.properties.columns.columnSpaces;
        for (let i = 0; i < widths.length; i++) {
          const colAttrs: Record<string, string | number> = {
            'w:w': widths[i]!.toString(),
          };
          const spaceVal = spaces?.[i];
          if (spaceVal !== undefined) {
            colAttrs['w:space'] = spaceVal.toString();
          }
          colChildren.push(XMLBuilder.wSelf('col', colAttrs));
        }
      }

      children.push(
        colChildren.length > 0
          ? XMLBuilder.w('cols', attrs, colChildren)
          : XMLBuilder.wSelf('cols', attrs)
      );
    }

    // Form protection (w:formProt) — CT_OnOff per §17.6.9. Preserve the
    // explicit-false distinction on round-trip (parser already does so).
    if (this.properties.formProt !== undefined) {
      children.push(
        this.properties.formProt
          ? XMLBuilder.wSelf('formProt')
          : XMLBuilder.wSelf('formProt', { 'w:val': '0' })
      );
    }

    // Vertical alignment
    if (this.properties.verticalAlignment) {
      children.push(XMLBuilder.wSelf('vAlign', { 'w:val': this.properties.verticalAlignment }));
    }

    // Suppress endnotes (w:noEndnote) — CT_OnOff per §17.11.14.
    if (this.properties.noEndnote !== undefined) {
      children.push(
        this.properties.noEndnote
          ? XMLBuilder.wSelf('noEndnote')
          : XMLBuilder.wSelf('noEndnote', { 'w:val': '0' })
      );
    }

    // Title page (w:titlePg) — CT_OnOff per §17.10.6.
    if (this.properties.titlePage !== undefined) {
      children.push(
        this.properties.titlePage
          ? XMLBuilder.wSelf('titlePg', { 'w:val': '1' })
          : XMLBuilder.wSelf('titlePg', { 'w:val': '0' })
      );
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

    // Bidirectional section (RTL) — CT_OnOff per §17.6.3.
    if (this.properties.bidi !== undefined) {
      children.push(
        this.properties.bidi ? XMLBuilder.wSelf('bidi') : XMLBuilder.wSelf('bidi', { 'w:val': '0' })
      );
    }

    // RTL gutter — CT_OnOff per §17.6.16.
    if (this.properties.rtlGutter !== undefined) {
      children.push(
        this.properties.rtlGutter
          ? XMLBuilder.wSelf('rtlGutter')
          : XMLBuilder.wSelf('rtlGutter', { 'w:val': '0' })
      );
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
        // Ordered per CT_SectPrBase (ECMA-376 §17.6.18):
        // footnotePr → endnotePr → type → pgSz → pgMar → paperSrc → pgBorders →
        // lnNumType → pgNumType → cols → formProt → vAlign → noEndnote → titlePg →
        // textDirection → bidi → rtlGutter → docGrid
        if (prev.footnotePr) {
          const fnChildren: XMLElement[] = [];
          if (prev.footnotePr.position) {
            fnChildren.push(XMLBuilder.wSelf('pos', { 'w:val': prev.footnotePr.position }));
          }
          if (prev.footnotePr.numberFormat) {
            fnChildren.push(XMLBuilder.wSelf('numFmt', { 'w:val': prev.footnotePr.numberFormat }));
          }
          if (prev.footnotePr.startNumber !== undefined) {
            fnChildren.push(
              XMLBuilder.wSelf('numStart', { 'w:val': prev.footnotePr.startNumber.toString() })
            );
          }
          if (prev.footnotePr.restart) {
            fnChildren.push(XMLBuilder.wSelf('numRestart', { 'w:val': prev.footnotePr.restart }));
          }
          if (fnChildren.length > 0) {
            prevChildren.push(XMLBuilder.w('footnotePr', undefined, fnChildren));
          }
        }
        if (prev.endnotePr) {
          const enChildren: XMLElement[] = [];
          if (prev.endnotePr.position) {
            enChildren.push(XMLBuilder.wSelf('pos', { 'w:val': prev.endnotePr.position }));
          }
          if (prev.endnotePr.numberFormat) {
            enChildren.push(XMLBuilder.wSelf('numFmt', { 'w:val': prev.endnotePr.numberFormat }));
          }
          if (prev.endnotePr.startNumber !== undefined) {
            enChildren.push(
              XMLBuilder.wSelf('numStart', { 'w:val': prev.endnotePr.startNumber.toString() })
            );
          }
          if (prev.endnotePr.restart) {
            enChildren.push(XMLBuilder.wSelf('numRestart', { 'w:val': prev.endnotePr.restart }));
          }
          if (enChildren.length > 0) {
            prevChildren.push(XMLBuilder.w('endnotePr', undefined, enChildren));
          }
        }
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
          if (prev.pageSize.code !== undefined) {
            pgSzAttrs['w:code'] = prev.pageSize.code.toString();
          }
          prevChildren.push(XMLBuilder.wSelf('pgSz', pgSzAttrs));
        }
        if (prev.margins) {
          // CT_PageMar §17.6.11 requires all 7 attributes. Supply defaults
          // (top/right/bottom/left → 1440 twips = 1" following Word's default,
          // header/footer → 720, gutter → 0) for any missing from the
          // sectPrChange previous-properties snapshot so the emitted
          // previous-state is still schema-compliant.
          const pgMarAttrs: Record<string, string> = {
            'w:top': (prev.margins.top ?? 1440).toString(),
            'w:right': (prev.margins.right ?? 1440).toString(),
            'w:bottom': (prev.margins.bottom ?? 1440).toString(),
            'w:left': (prev.margins.left ?? 1440).toString(),
            'w:header': (prev.margins.header ?? 720).toString(),
            'w:footer': (prev.margins.footer ?? 720).toString(),
            'w:gutter': (prev.margins.gutter ?? 0).toString(),
          };
          prevChildren.push(XMLBuilder.wSelf('pgMar', pgMarAttrs));
        }
        if (prev.paperSource) {
          const psAttrs: Record<string, string> = {};
          if (prev.paperSource.first !== undefined) {
            psAttrs['w:first'] = prev.paperSource.first.toString();
          }
          if (prev.paperSource.other !== undefined) {
            psAttrs['w:other'] = prev.paperSource.other.toString();
          }
          if (Object.keys(psAttrs).length > 0) {
            prevChildren.push(XMLBuilder.wSelf('paperSrc', psAttrs));
          }
        }
        if (prev.pageBorders) {
          const pgb = prev.pageBorders;
          const pgbAttrs: Record<string, string> = {};
          if (pgb.offsetFrom) pgbAttrs['w:offsetFrom'] = pgb.offsetFrom;
          if (pgb.display) pgbAttrs['w:display'] = pgb.display;
          if (pgb.zOrder) pgbAttrs['w:zOrder'] = pgb.zOrder;
          const borderChildren: XMLElement[] = [];
          const buildPrevBorder = (side: string, def: PageBorderDef) => {
            // CT_PageBorder §17.6.2: w:val required. Default to "nil".
            // Full CT_Border attribute set preserved for sectPrChange history.
            const bAttrs: Record<string, string | number> = { 'w:val': def.style ?? 'nil' };
            if (def.size !== undefined) bAttrs['w:sz'] = Math.max(2, Math.min(96, def.size));
            if (def.color) bAttrs['w:color'] = def.color;
            if (def.space !== undefined)
              bAttrs['w:space'] = Math.max(0, Math.min(31680, def.space));
            if (def.shadow !== undefined) bAttrs['w:shadow'] = def.shadow ? '1' : '0';
            if (def.frame !== undefined) bAttrs['w:frame'] = def.frame ? '1' : '0';
            if (def.themeColor) bAttrs['w:themeColor'] = def.themeColor;
            if (def.themeTint) bAttrs['w:themeTint'] = def.themeTint;
            if (def.themeShade) bAttrs['w:themeShade'] = def.themeShade;
            if (def.artId !== undefined) bAttrs['w:id'] = def.artId;
            borderChildren.push(XMLBuilder.wSelf(side, bAttrs));
          };
          if (pgb.top) buildPrevBorder('top', pgb.top);
          if (pgb.left) buildPrevBorder('left', pgb.left);
          if (pgb.bottom) buildPrevBorder('bottom', pgb.bottom);
          if (pgb.right) buildPrevBorder('right', pgb.right);
          if (borderChildren.length > 0) {
            prevChildren.push(XMLBuilder.w('pgBorders', pgbAttrs, borderChildren));
          }
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
        if (prev.pageNumbering || prev.chapStyle !== undefined || prev.chapSep) {
          // CT_PageNumber §17.6.12 has four attributes: fmt / start /
          // chapStyle / chapSep. chapStyle/chapSep live on the root of
          // SectionProperties (not inside pageNumbering) — mirror the
          // main-sectPr emitter at line ~1215.
          const pnAttrs: Record<string, string> = {};
          if (prev.pageNumbering?.start !== undefined)
            pnAttrs['w:start'] = prev.pageNumbering.start.toString();
          if (prev.pageNumbering?.format) pnAttrs['w:fmt'] = prev.pageNumbering.format;
          if (prev.chapStyle !== undefined) pnAttrs['w:chapStyle'] = prev.chapStyle.toString();
          if (prev.chapSep) pnAttrs['w:chapSep'] = prev.chapSep;
          if (Object.keys(pnAttrs).length > 0) {
            prevChildren.push(XMLBuilder.wSelf('pgNumType', pnAttrs));
          }
        }
        if (prev.columns) {
          const colAttrs: Record<string, string> = {
            'w:num': prev.columns.count?.toString() || '1',
          };
          if (prev.columns.space !== undefined) colAttrs['w:space'] = prev.columns.space.toString();
          if (prev.columns.equalWidth !== undefined) {
            colAttrs['w:equalWidth'] = prev.columns.equalWidth ? '1' : '0';
          }
          if (prev.columns.separator !== undefined) {
            colAttrs['w:sep'] = prev.columns.separator ? '1' : '0';
          }
          const colChildren: XMLElement[] = [];
          if (prev.columns.columnWidths) {
            const widths = prev.columns.columnWidths;
            const spaces = prev.columns.columnSpaces;
            for (let i = 0; i < widths.length; i++) {
              const cAttrs: Record<string, string | number> = {
                'w:w': widths[i]!.toString(),
              };
              const spaceVal = spaces?.[i];
              if (spaceVal !== undefined) {
                cAttrs['w:space'] = spaceVal.toString();
              }
              colChildren.push(XMLBuilder.wSelf('col', cAttrs));
            }
          }
          prevChildren.push(
            colChildren.length > 0
              ? XMLBuilder.w('cols', colAttrs, colChildren)
              : XMLBuilder.wSelf('cols', colAttrs)
          );
        }
        // Section-level CT_OnOff flags (formProt §17.6.9 / noEndnote
        // §17.11.14 / titlePg §17.10.6 / bidi §17.6.1 / rtlGutter
        // §17.6.16). The main emitter already preserves the explicit-
        // false distinction via `!== undefined` (see §1286 formProt),
        // but the sectPrChange emitter used a truthy gate — so a
        // tracked change capturing a "previous state = false" (e.g.,
        // form protection was OFF before the revision that turned it
        // ON) silently dropped the `<w:formProt w:val="0"/>` marker on
        // save. Parity-fix: guard on presence and emit w:val="0" for
        // explicit false.
        if (prev.formProt !== undefined) {
          prevChildren.push(
            prev.formProt
              ? XMLBuilder.wSelf('formProt')
              : XMLBuilder.wSelf('formProt', { 'w:val': '0' })
          );
        }
        if (prev.verticalAlignment) {
          prevChildren.push(XMLBuilder.wSelf('vAlign', { 'w:val': prev.verticalAlignment }));
        }
        if (prev.noEndnote !== undefined) {
          prevChildren.push(
            prev.noEndnote
              ? XMLBuilder.wSelf('noEndnote')
              : XMLBuilder.wSelf('noEndnote', { 'w:val': '0' })
          );
        }
        if (prev.titlePage !== undefined) {
          prevChildren.push(XMLBuilder.wSelf('titlePg', { 'w:val': prev.titlePage ? '1' : '0' }));
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
        if (prev.bidi !== undefined) {
          prevChildren.push(
            prev.bidi ? XMLBuilder.wSelf('bidi') : XMLBuilder.wSelf('bidi', { 'w:val': '0' })
          );
        }
        if (prev.rtlGutter !== undefined) {
          prevChildren.push(
            prev.rtlGutter
              ? XMLBuilder.wSelf('rtlGutter')
              : XMLBuilder.wSelf('rtlGutter', { 'w:val': '0' })
          );
        }
        if (prev.docGrid) {
          const dgAttrs: Record<string, string> = {};
          if (prev.docGrid.type) dgAttrs['w:type'] = prev.docGrid.type;
          if (prev.docGrid.linePitch !== undefined)
            dgAttrs['w:linePitch'] = prev.docGrid.linePitch.toString();
          if (prev.docGrid.charSpace !== undefined)
            dgAttrs['w:charSpace'] = prev.docGrid.charSpace.toString();
          if (Object.keys(dgAttrs).length > 0) {
            prevChildren.push(XMLBuilder.wSelf('docGrid', dgAttrs));
          }
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
        columnSpaces: this.properties.columns.columnSpaces
          ? [...this.properties.columns.columnSpaces]
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
   * Creates a legal-sized section (8.5" x 14")
   */
  static createLegal(): Section {
    return new Section({
      pageSize: {
        width: PAGE_SIZES.LEGAL.width,
        height: PAGE_SIZES.LEGAL.height,
        orientation: 'portrait',
      },
    });
  }

  /**
   * Creates a tabloid-sized section (11" x 17")
   */
  static createTabloid(): Section {
    return new Section({
      pageSize: {
        width: PAGE_SIZES.TABLOID.width,
        height: PAGE_SIZES.TABLOID.height,
        orientation: 'portrait',
      },
    });
  }

  /**
   * Creates an A3-sized section (29.7cm x 42cm)
   */
  static createA3(): Section {
    return new Section({
      pageSize: {
        width: PAGE_SIZES.A3.width,
        height: PAGE_SIZES.A3.height,
        orientation: 'portrait',
      },
    });
  }

  /**
   * Creates a landscape section
   *
   * @param pageSize - Page size name (default: 'letter')
   * @returns Section with swapped width/height and landscape orientation
   *
   * @example
   * ```typescript
   * const landscape = Section.createLandscape('a4');
   * const legalLandscape = Section.createLandscape('legal');
   * ```
   */
  static createLandscape(
    pageSize: 'letter' | 'a4' | 'legal' | 'tabloid' | 'a3' = 'letter'
  ): Section {
    const sizeMap: Record<string, { width: number; height: number }> = {
      letter: PAGE_SIZES.LETTER,
      a4: PAGE_SIZES.A4,
      legal: PAGE_SIZES.LEGAL,
      tabloid: PAGE_SIZES.TABLOID,
      a3: PAGE_SIZES.A3,
    };
    const size = sizeMap[pageSize] ?? PAGE_SIZES.LETTER;
    return new Section({
      pageSize: {
        width: size.height, // Swap for landscape
        height: size.width,
        orientation: 'landscape',
      },
    });
  }
}
