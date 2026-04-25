/**
 * Paragraph - Represents a paragraph in a Word document
 * Contains one or more runs of formatted text
 */

import { deepClone } from '../utils/deepClone';
import { deepEqual } from '../utils/deepEqual';
import { formatDateForXml } from '../utils/dateFormatting';
import { logParagraphContent, logTextDirection } from '../utils/diagnostics';
import { isEqualFormatting } from '../utils/formatting';
import { defaultLogger } from '../utils/logger';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { Bookmark } from './Bookmark';
import type { Comment } from './Comment';
import {
  // Import common types
  ParagraphAlignment as CommonParagraphAlignment,
  BorderStyle as CommonBorderStyle,
  FullBorderStyle as CommonFullBorderStyle,
  BasicShadingPattern,
  TabAlignment as CommonTabAlignment,
  TabLeader as CommonTabLeader,
  TextDirection as CommonTextDirection,
  TextVerticalAlignment,
  ShadingConfig,
  buildShadingAttributes,
} from './CommonTypes';
import { ComplexField, Field } from './Field';
import { Hyperlink } from './Hyperlink';
import { RangeMarker } from './RangeMarker';
import { Revision } from './Revision';
import { Run, RunFormatting } from './Run';
import { Shape } from './Shape';
import { TextBox } from './TextBox';
import { PreservedElement } from './PreservedElement';

// ============================================================================
// RE-EXPORTED TYPES (for backward compatibility)
// These types are now defined in CommonTypes.ts but re-exported here
// to maintain backward compatibility with existing imports.
// ============================================================================

/**
 * Paragraph alignment options
 * @see CommonTypes.ParagraphAlignment
 */
export type ParagraphAlignment = CommonParagraphAlignment;

/**
 * Type to indicate ComplexField support in paragraph content
 */
export type FieldLike = Field | ComplexField;

/**
 * Border style types for paragraph borders
 * @see CommonTypes.BorderStyle
 */
export type BorderStyle = CommonBorderStyle;

/**
 * Shading pattern types (basic patterns without percentage fills)
 * @see CommonTypes.BasicShadingPattern for basic patterns
 * @see CommonTypes.ShadingPattern for full pattern set including percentages
 */
export type ShadingPattern = BasicShadingPattern;

/**
 * Tab stop alignment types
 * @see CommonTypes.TabAlignment
 */
export type TabAlignment = CommonTabAlignment;

/**
 * Tab stop leader types
 * @see CommonTypes.TabLeader
 */
export type TabLeader = CommonTabLeader;

/**
 * Text direction types for paragraphs
 * @see CommonTypes.TextDirection
 */
export type TextDirection = CommonTextDirection;

/**
 * Text vertical alignment types
 * @see CommonTypes.TextVerticalAlignment
 */
export type TextAlignment = TextVerticalAlignment;

/**
 * Textbox tight wrap modes
 */
export type TextboxTightWrap =
  | 'none'
  | 'allLines'
  | 'firstAndLastLine'
  | 'firstLineOnly'
  | 'lastLineOnly';

/**
 * Paragraph property change tracking (for revision history)
 */
export interface ParagraphPropertiesChange {
  /** Author of the change */
  author?: string;
  /** Date of the change */
  date?: string;
  /** Unique ID for this revision */
  id?: string;
  /** Previous paragraph properties before the change (stored as object) */
  previousProperties?: Partial<ParagraphFormatting>;
}

/**
 * Frame/text box properties
 */
export interface FrameProperties {
  /** Width in twips */
  w?: number;
  /** Height in twips */
  h?: number;
  /** Height rule */
  hRule?: 'auto' | 'atLeast' | 'exact';
  /** Absolute horizontal position in twips */
  x?: number;
  /** Absolute vertical position in twips */
  y?: number;
  /** Relative horizontal alignment */
  xAlign?: 'left' | 'center' | 'right' | 'inside' | 'outside';
  /**
   * Relative vertical alignment per ECMA-376 Part 1 §17.18.110 ST_YAlign.
   * `inline` anchors the frame in line with the surrounding text rather
   * than offsetting it vertically.
   */
  yAlign?: 'top' | 'center' | 'bottom' | 'inline' | 'inside' | 'outside';
  /** Horizontal anchor/positioning base */
  hAnchor?: 'page' | 'margin' | 'text';
  /** Vertical anchor/positioning base */
  vAnchor?: 'page' | 'margin' | 'text';
  /** Horizontal padding in twips */
  hSpace?: number;
  /** Vertical padding in twips */
  vSpace?: number;
  /** Text wrapping around frame per ECMA-376 ST_Wrap (§17.18.104) */
  wrap?: 'around' | 'auto' | 'none' | 'notBeside' | 'through' | 'tight';
  /** Drop cap style */
  dropCap?: 'none' | 'drop' | 'margin';
  /** Drop cap height in lines */
  lines?: number;
  /** Lock frame anchor to paragraph */
  anchorLock?: boolean;
}

/**
 * Single paragraph-border definition. `style` uses FullBorderStyle (25+
 * ST_Border values per ECMA-376 §17.18.2) — paragraph borders support
 * the full multi-line-gap / triple / inset / outset set, not just the
 * narrow 6-value BorderStyle subset that this field previously used.
 */
export interface BorderDefinition {
  /** Border style (full ST_Border enumeration) */
  style?: CommonFullBorderStyle;
  /** Border width in eighths of a point (1-96) */
  size?: number;
  /** Border color (hex without #) */
  color?: string;
  /** Space between border and text in points (0-31) */
  space?: number;
  /** Theme color reference (ST_ThemeColor per §17.18.97) */
  themeColor?: string;
  /** Theme tint (2-hex-digit string) */
  themeTint?: string;
  /** Theme shade (2-hex-digit string) */
  themeShade?: string;
  /** Border casts a shadow (ST_OnOff attribute on CT_Border §17.18.2) */
  shadow?: boolean;
  /** Border is part of a frame around the content (ST_OnOff) */
  frame?: boolean;
}

/**
 * Tab stop definition
 */
export interface TabStop {
  /** Position in twips */
  position: number;
  /** Alignment type */
  val?: TabAlignment;
  /** Leader character */
  leader?: TabLeader;
}

/**
 * Paragraph formatting options
 */
export interface ParagraphFormatting {
  /** Text alignment */
  alignment?: ParagraphAlignment;
  /** Indentation in twips (1/20th of a point) */
  indentation?: {
    left?: number;
    right?: number;
    firstLine?: number;
    hanging?: number;
    /**
     * CJK character-unit indentation attributes per ECMA-376 §17.3.1.12.
     * Values are in hundredths of a character unit (200 = 2 character widths).
     * Used alongside (not instead of) the twips attributes: Word authors
     * specify one or the other based on document language, and both are
     * preserved on round-trip so downstream consumers keep the author's
     * original intent.
     */
    leftChars?: number;
    rightChars?: number;
    firstLineChars?: number;
    hangingChars?: number;
  };
  /** Spacing per ECMA-376 §17.3.1.33 (CT_Spacing) */
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: 'auto' | 'exact' | 'atLeast';
    /** Spacing before in hundredths of a line */
    beforeLines?: number;
    /** Spacing after in hundredths of a line */
    afterLines?: number;
    /** Auto-calculate before spacing (overrides w:before when true) */
    beforeAutospacing?: boolean;
    /** Auto-calculate after spacing (overrides w:after when true) */
    afterAutospacing?: boolean;
  };
  /** Keep with next paragraph */
  keepNext?: boolean;
  /** Keep lines together */
  keepLines?: boolean;
  /** Page break before */
  pageBreakBefore?: boolean;
  /** Paragraph style ID */
  style?: string;
  /** Numbering properties */
  numbering?: {
    numId: number;
    level: number;
  };
  /** Contextual spacing - removes spacing between paragraphs of same style */
  contextualSpacing?: boolean;
  /** Paragraph ID (Word 2010+) - required by modern Word for change tracking */
  paraId?: string;
  /** Text ID (Word 2010+) - tracks text modifications for merge conflict resolution */
  textId?: string;
  /** Paragraph borders (top, bottom, left, right, between, bar) */
  borders?: {
    top?: BorderDefinition;
    bottom?: BorderDefinition;
    left?: BorderDefinition;
    right?: BorderDefinition;
    between?: BorderDefinition;
    bar?: BorderDefinition;
  };
  /** Paragraph shading (background color and pattern) */
  shading?: ShadingConfig;
  /** Tab stops */
  tabs?: TabStop[];
  /** Widow/orphan control - prevents single lines at top/bottom of pages */
  widowControl?: boolean;
  /** Outline level (0-9) for table of contents hierarchy */
  outlineLevel?: number;
  /** Suppress line numbers for this paragraph */
  suppressLineNumbers?: boolean;
  /** Right-to-left paragraph layout (for Arabic, Hebrew, etc.) */
  bidi?: boolean;
  /** Text flow direction */
  textDirection?: TextDirection;
  /** Vertical text alignment */
  textAlignment?: TextAlignment;
  /** Use inside/outside indents instead of left/right (for double-sided printing) */
  mirrorIndents?: boolean;
  /** Auto-adjust right indent when document grid is defined */
  adjustRightInd?: boolean;
  /** Text frame/box properties (positioning, wrapping, drop cap) */
  framePr?: FrameProperties;
  /** Suppress automatic hyphenation for this paragraph */
  suppressAutoHyphens?: boolean;
  /** Kinsoku rules - CJK line-breaking rules per ECMA-376 Part 1 §17.3.1.16 */
  kinsoku?: boolean;
  /** Word wrap - allow CJK text to wrap mid-word per ECMA-376 Part 1 §17.3.1.45 */
  wordWrap?: boolean;
  /** Overflow punctuation - allow CJK punctuation to overhang margins per ECMA-376 Part 1 §17.3.1.24 */
  overflowPunct?: boolean;
  /** Top line punctuation - compress CJK punctuation at start of line per ECMA-376 Part 1 §17.3.1.43 */
  topLinePunct?: boolean;
  /** Auto space between East Asian and numeric text per ECMA-376 Part 1 §17.3.1.2 */
  autoSpaceDE?: boolean;
  /** Auto space between East Asian and Western text per ECMA-376 Part 1 §17.3.1.3 */
  autoSpaceDN?: boolean;
  /** Prevent text frames from overlapping */
  suppressOverlap?: boolean;
  /** Tight wrapping mode for text boxes */
  textboxTightWrap?: TextboxTightWrap;
  /** Associated HTML div ID (for HTML round-trip) */
  divId?: number;
  /** Conditional table style formatting (bitmask string, e.g., "101000000100") */
  cnfStyle?: string;
  /** Section properties at paragraph level (for section breaks) */
  sectPr?: string | Record<string, unknown>;
  /** Paragraph property change tracking (revision history) */
  pPrChange?: ParagraphPropertiesChange;
  /** Run properties for the paragraph mark (¶ symbol formatting) */
  paragraphMarkRunProperties?: RunFormatting;
  /**
   * Paragraph mark run-properties change tracking (w:rPrChange inside
   * <w:pPr><w:rPr>), per ECMA-376 Part 1 §17.3.1.30 / CT_ParaRPrChange.
   * Stores the previous rPr of the pilcrow glyph when its formatting has
   * been modified under track changes. Emitted as the LAST child of the
   * paragraph-mark <w:rPr> per CT_ParaRPr schema.
   */
  paragraphMarkRunPropertiesChange?: {
    id: number;
    author: string;
    date: Date;
    previousProperties: Partial<RunFormatting>;
  };
  /** Paragraph mark deletion tracking (for deleted ¶ symbols) */
  paragraphMarkDeletion?: {
    /** Unique revision ID */
    id: number;
    /** Author who deleted the paragraph mark */
    author: string;
    /** Date when the paragraph mark was deleted */
    date: Date;
  };
  /** Paragraph mark insertion tracking (for inserted ¶ symbols) */
  paragraphMarkInsertion?: {
    /** Unique revision ID */
    id: number;
    /** Author who inserted the paragraph mark */
    author: string;
    /** Date when the paragraph mark was inserted */
    date: Date;
  };
  /** True when the original XML had numId=0 (explicitly suppressed numbering) */
  numberingSuppressed?: boolean;
}

/**
 * Paragraph content (runs, fields, hyperlinks, revisions, range markers, shapes, text boxes)
 */
export type ParagraphContent =
  | Run
  | FieldLike
  | Hyperlink
  | Revision
  | RangeMarker
  | Shape
  | TextBox
  | PreservedElement;

// ============================================================================
// TYPE GUARDS FOR ParagraphContent
// These functions help with type narrowing when working with paragraph content
// ============================================================================

/**
 * Type guard: Check if content is a Run
 * @param content - Paragraph content item to check
 * @returns True if the content is a Run instance
 */
export function isRun(content: ParagraphContent): content is Run {
  return content instanceof Run;
}

/**
 * Type guard: Check if content is a Field (simple or complex)
 * @param content - Paragraph content item to check
 * @returns True if the content is a Field or ComplexField instance
 */
export function isField(content: ParagraphContent): content is FieldLike {
  return content instanceof Field || content instanceof ComplexField;
}

/**
 * Type guard: Check if content is a simple Field
 * @param content - Paragraph content item to check
 * @returns True if the content is a Field instance (not ComplexField)
 */
export function isSimpleField(content: ParagraphContent): content is Field {
  return content instanceof Field && !(content instanceof ComplexField);
}

/**
 * Type guard: Check if content is a ComplexField
 * @param content - Paragraph content item to check
 * @returns True if the content is a ComplexField instance
 */
export function isComplexField(content: ParagraphContent): content is ComplexField {
  return content instanceof ComplexField;
}

/**
 * Type guard: Check if content is a Hyperlink
 * @param content - Paragraph content item to check
 * @returns True if the content is a Hyperlink instance
 */
export function isHyperlink(content: ParagraphContent): content is Hyperlink {
  return content instanceof Hyperlink;
}

/**
 * Type guard: Check if content is a Revision
 * @param content - Paragraph content item to check
 * @returns True if the content is a Revision instance
 */
export function isRevision(content: ParagraphContent): content is Revision {
  return content instanceof Revision;
}

/**
 * Type guard: Check if content is a RangeMarker (bookmark start/end, comment start/end)
 * @param content - Paragraph content item to check
 * @returns True if the content is a RangeMarker instance
 */
export function isRangeMarker(content: ParagraphContent): content is RangeMarker {
  return content instanceof RangeMarker;
}

/**
 * Type guard: Check if content is a Shape
 * @param content - Paragraph content item to check
 * @returns True if the content is a Shape instance
 */
export function isShape(content: ParagraphContent): content is Shape {
  return content instanceof Shape;
}

/**
 * Type guard: Check if content is a TextBox
 * @param content - Paragraph content item to check
 * @returns True if the content is a TextBox instance
 */
export function isTextBox(content: ParagraphContent): content is TextBox {
  return content instanceof TextBox;
}

/**
 * Represents a paragraph in a document
 */
export class Paragraph {
  private content: ParagraphContent[] = [];
  public formatting: ParagraphFormatting;
  private bookmarksStart: Bookmark[] = [];
  private bookmarksEnd: Bookmark[] = [];
  private commentsStart: Comment[] = [];
  private commentsEnd: Comment[] = [];
  /** Internal flag to mark paragraph as preserved from removal operations */
  private _isPreserved = false;
  /** Tracking context for automatic change tracking */
  private trackingContext?: import('../tracking/TrackingContext').TrackingContext;
  /** Parent table cell reference (if paragraph is inside a table cell) */
  private _parentCell?: import('./TableCell').TableCell;
  /** StylesManager reference for conditional formatting resolution */
  private _stylesManager?: import('../formatting/StylesManager').StylesManager;
  /**
   * Internal flag to mark paragraph as part of a multi-paragraph field (e.g., TOC)
   * When true, assembleComplexFields() will skip processing this paragraph
   * to preserve the original field structure across paragraphs.
   * @internal
   */
  _isPartOfMultiParagraphField?: boolean;

  /**
   * Creates a new Paragraph
   * @param formatting - Paragraph formatting options
   */
  constructor(formatting: ParagraphFormatting = {}) {
    this.formatting = formatting;
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
   * Sets the parent cell reference for this paragraph.
   * Called by TableCell when adding paragraphs.
   * @internal
   */
  _setParentCell(cell: import('./TableCell').TableCell | undefined): void {
    this._parentCell = cell;
  }

  /**
   * Gets the parent cell reference for this paragraph.
   * @internal
   */
  _getParentCell(): import('./TableCell').TableCell | undefined {
    return this._parentCell;
  }

  /**
   * Checks if this paragraph is inside a table cell.
   * @returns True if paragraph has a parent cell
   */
  isInTableCell(): boolean {
    return this._parentCell !== undefined;
  }

  /**
   * Gets the table's cnfStyle (conditional formatting flags) for this paragraph.
   * Returns the cell's cnfStyle if in a table, or the paragraph's own cnfStyle.
   * @returns The cnfStyle string or undefined
   */
  getTableConditionalStyle(): string | undefined {
    // Cell cnfStyle takes precedence
    if (this._parentCell) {
      const cellFormatting = this._parentCell.getFormatting();
      if (cellFormatting.cnfStyle) {
        return cellFormatting.cnfStyle;
      }
    }
    // Fall back to paragraph's own cnfStyle
    return this.formatting.cnfStyle;
  }

  /**
   * Sets the StylesManager reference for conditional formatting resolution.
   * Called by Document/Table when adding paragraphs.
   * @internal
   */
  _setStylesManager(manager: import('../formatting/StylesManager').StylesManager): void {
    this._stylesManager = manager;
  }

  /**
   * Gets the StylesManager reference for conditional formatting resolution.
   * @internal
   */
  _getStylesManager(): import('../formatting/StylesManager').StylesManager | undefined {
    return this._stylesManager;
  }

  /**
   * Creates an empty detached paragraph
   * @returns New Paragraph instance
   * @example
   * const para = Paragraph.create();
   * para.addText('Added later');
   */
  static create(): Paragraph;

  /**
   * Creates a detached paragraph with formatting
   * @param formatting - Paragraph formatting options
   * @returns New Paragraph instance
   * @example
   * const para = Paragraph.create({ alignment: 'center' });
   */
  static create(formatting: ParagraphFormatting): Paragraph;

  /**
   * Creates a detached paragraph with text and optional formatting
   * @param text - Text content
   * @param formatting - Optional paragraph formatting
   * @returns New Paragraph instance
   * @example
   * const para = Paragraph.create('Hello World', { alignment: 'center' });
   */
  static create(text: string, formatting?: ParagraphFormatting): Paragraph;

  /**
   * Creates a detached paragraph (not yet added to a document)
   * @param textOrFormatting - Optional text content or paragraph formatting
   * @param formatting - Optional paragraph formatting (only used if first param is text)
   * @returns New Paragraph instance
   * @example
   * // Create with text and formatting
   * const para1 = Paragraph.create('Hello World', { alignment: 'center' });
   *
   * // Create with just formatting
   * const para2 = Paragraph.create({ alignment: 'right' });
   *
   * // Create empty
   * const para3 = Paragraph.create();
   *
   * // Add to document later
   * doc.addParagraph(para1);
   */
  static create(
    textOrFormatting?: string | ParagraphFormatting,
    formatting?: ParagraphFormatting
  ): Paragraph {
    // Handle overloaded parameters
    if (typeof textOrFormatting === 'string') {
      // First param is text
      const paragraph = new Paragraph(formatting);
      paragraph.addText(textOrFormatting);
      return paragraph;
    } else {
      // First param is formatting (or undefined)
      return new Paragraph(textOrFormatting);
    }
  }

  /**
   * Creates a detached paragraph with a specific style
   * @param text - Text content
   * @param styleId - Style ID (e.g., 'Heading1', 'Title')
   * @returns New Paragraph instance
   * @example
   * const heading = Paragraph.createWithStyle('Chapter 1', 'Heading1');
   * doc.addParagraph(heading);
   */
  static createWithStyle(text: string, styleId: string): Paragraph {
    const paragraph = new Paragraph({ style: styleId });
    paragraph.addText(text);
    return paragraph;
  }

  /**
   * Creates a detached empty paragraph
   * Useful for adding blank lines or spacing
   * @returns New empty Paragraph instance
   * @example
   * const blank = Paragraph.createEmpty();
   * doc.addParagraph(blank);
   */
  static createEmpty(): Paragraph {
    return new Paragraph();
  }

  /**
   * Creates a detached paragraph with formatted text
   * @param text - Text content
   * @param runFormatting - Run formatting (bold, italic, etc.)
   * @param paragraphFormatting - Paragraph formatting (alignment, spacing, etc.)
   * @returns New Paragraph instance
   * @example
   * const para = Paragraph.createFormatted(
   *   'Important Text',
   *   { bold: true, color: 'FF0000' },
   *   { alignment: 'center' }
   * );
   */
  static createFormatted(
    text: string,
    runFormatting?: RunFormatting,
    paragraphFormatting?: ParagraphFormatting
  ): Paragraph {
    const paragraph = new Paragraph(paragraphFormatting);
    paragraph.addText(text, runFormatting);
    return paragraph;
  }

  /**
   * Adds a Run to the paragraph
   *
   * Appends a Run instance to the paragraph's content. Runs are sequences
   * of text with uniform formatting.
   *
   * @param run - The Run instance to add
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * const para = new Paragraph();
   * const run = new Run('Bold text', { bold: true });
   * para.addRun(run);
   * ```
   */
  addRun(run: Run): this {
    // Set parent reference for setText tracking
    run._setParentParagraph(this);

    if (this.trackingContext?.isEnabled()) {
      // Wrap the run in an insert revision when tracking is enabled
      const revision = Revision.createInsertion(this.trackingContext.getAuthor(), run, new Date());
      this.trackingContext.getRevisionManager().register(revision);
      this.content.push(revision);
    } else {
      this.content.push(run);
    }
    return this;
  }

  /**
   * Adds a field to the paragraph (supports both Field and ComplexField)
   * @param field - Field or ComplexField to add
   * @returns This paragraph for chaining
   */
  addField(field: FieldLike): this {
    if (this.trackingContext?.isEnabled()) {
      // Fields need special handling - wrap in revision if the field has runs
      // For now, just add directly as complex fields have their own structure
      this.content.push(field);
    } else {
      this.content.push(field);
    }
    return this;
  }

  /**
   * Adds a complex field to the paragraph
   * @param field - ComplexField to add
   * @returns This paragraph for chaining
   */
  addComplexField(field: ComplexField): this {
    this.content.push(field);
    return this;
  }

  /**
   * Adds a hyperlink to the paragraph
   * @param urlOrHyperlink - URL string for new hyperlink, or existing Hyperlink object
   * @returns Hyperlink object for fluent chaining (when creating new), or this paragraph (when adding existing)
   *
   * @example
   * // Fluent API (new signature)
   * const link = para.addHyperlink('https://example.com');
   * link.setText('Visit Example');
   *
   * // Or use without URL
   * const link2 = para.addHyperlink();
   * link2.setUrl('https://example.com').setText('Link');
   *
   * // Legacy API (still supported)
   * const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
   * para.addHyperlink(hyperlink);
   */
  /**
   * Adds a hyperlink to the paragraph
   * NOTE: This method has overloaded return types:
   * - `addHyperlink(url: string)` returns the `Hyperlink` object (for configuring the link)
   * - `addHyperlink(hyperlink: Hyperlink)` returns `this` (for paragraph chaining)
   *
   * @example
   * ```typescript
   * // Pattern 1: Create and configure a new hyperlink
   * const link = para.addHyperlink('https://example.com');
   * link.setText('Visit Example');
   *
   * // Pattern 2: Add pre-built hyperlink (returns paragraph for chaining)
   * para.addHyperlink(new Hyperlink({ url: 'https://example.com', text: 'Link' }))
   *     .addText(' more text after the link');
   * ```
   *
   * @param url - URL string to create a new hyperlink
   * @returns Hyperlink object for configuring the link
   */
  addHyperlink(url?: string): Hyperlink;
  /**
   * @param hyperlink - Existing Hyperlink object to add
   * @returns This paragraph for chaining
   */
  addHyperlink(hyperlink: Hyperlink): this;
  addHyperlink(urlOrHyperlink?: string | Hyperlink): Hyperlink | this {
    if (typeof urlOrHyperlink === 'string') {
      // New fluent API: create hyperlink from URL
      const hyperlink = new Hyperlink({ url: urlOrHyperlink, text: urlOrHyperlink });
      hyperlink._setParentParagraph(this);
      if (this.trackingContext?.isEnabled()) {
        const revision = Revision.createInsertion(
          this.trackingContext.getAuthor(),
          hyperlink,
          new Date()
        );
        this.trackingContext.getRevisionManager().register(revision);
        this.content.push(revision);
      } else {
        this.content.push(hyperlink);
      }
      return hyperlink;
    } else if (urlOrHyperlink instanceof Hyperlink) {
      // Legacy API: add existing hyperlink
      urlOrHyperlink._setParentParagraph(this);
      if (this.trackingContext?.isEnabled()) {
        const revision = Revision.createInsertion(
          this.trackingContext.getAuthor(),
          urlOrHyperlink,
          new Date()
        );
        this.trackingContext.getRevisionManager().register(revision);
        this.content.push(revision);
      } else {
        this.content.push(urlOrHyperlink);
      }
      return this;
    } else {
      // No argument: create empty hyperlink for fluent building
      const hyperlink = new Hyperlink({ text: 'Link' });
      hyperlink._setParentParagraph(this);
      if (this.trackingContext?.isEnabled()) {
        const revision = Revision.createInsertion(
          this.trackingContext.getAuthor(),
          hyperlink,
          new Date()
        );
        this.trackingContext.getRevisionManager().register(revision);
        this.content.push(revision);
      } else {
        this.content.push(hyperlink);
      }
      return hyperlink;
    }
  }

  /**
   * Adds a revision (tracked change) to the paragraph
   * @param revision - Revision to add
   * @returns This paragraph for chaining
   */
  addRevision(revision: Revision): this {
    this.content.push(revision);
    return this;
  }

  /**
   * Adds a range marker to the paragraph
   * Range markers mark the boundaries of moved, inserted, or deleted content
   * @param rangeMarker - Range marker to add
   * @returns This paragraph for chaining
   */
  addRangeMarker(rangeMarker: RangeMarker): this {
    this.content.push(rangeMarker);
    return this;
  }

  /**
   * Adds a shape to the paragraph
   * @param shape - Shape to add
   * @returns This paragraph for chaining
   * @example
   * const rect = Shape.createRectangle(inchesToEmus(2), inchesToEmus(1));
   * paragraph.addShape(rect);
   */
  addShape(shape: Shape): this {
    this.content.push(shape);
    return this;
  }

  /**
   * Adds a text box to the paragraph
   * @param textbox - TextBox to add
   * @returns This paragraph for chaining
   * @example
   * const textbox = TextBox.create(inchesToEmus(3), inchesToEmus(2));
   * paragraph.addTextBox(textbox);
   */
  addTextBox(textbox: TextBox): this {
    this.content.push(textbox);
    return this;
  }

  /**
   * Adds any ParagraphContent item to the paragraph
   * Used for preserved elements (proofErr, permStart/End, etc.)
   * @param item - Content item to add
   * @returns This paragraph for chaining
   */
  addContent(item: ParagraphContent): this {
    this.content.push(item);
    return this;
  }

  /**
   * Adds a bookmark start marker at the beginning of this paragraph
   * @param bookmark - Bookmark to add
   * @returns This paragraph for chaining
   */
  addBookmarkStart(bookmark: Bookmark): this {
    this.bookmarksStart.push(bookmark);
    return this;
  }

  /**
   * Adds a bookmark end marker at the end of this paragraph
   * @param bookmark - Bookmark to add (must have matching start marker)
   * @returns This paragraph for chaining
   */
  addBookmarkEnd(bookmark: Bookmark): this {
    this.bookmarksEnd.push(bookmark);
    return this;
  }

  /**
   * Adds both start and end bookmark markers (wraps entire paragraph)
   * @param bookmark - Bookmark to add
   * @returns This paragraph for chaining
   */
  addBookmark(bookmark: Bookmark): this {
    this.addBookmarkStart(bookmark);
    this.addBookmarkEnd(bookmark);
    return this;
  }

  /**
   * Gets all bookmarks that start in this paragraph
   * @returns Array of bookmarks
   */
  getBookmarksStart(): Bookmark[] {
    return [...this.bookmarksStart];
  }

  /**
   * Gets all bookmarks that end in this paragraph
   * @returns Array of bookmarks
   */
  getBookmarksEnd(): Bookmark[] {
    return [...this.bookmarksEnd];
  }

  /**
   * Removes a bookmark end marker by its ID
   * @param id - The bookmark ID to remove
   * @returns true if a bookmark was removed, false if not found
   */
  removeBookmarkEnd(id: number): boolean {
    const index = this.bookmarksEnd.findIndex((bm) => bm.getId() === id);
    if (index !== -1) {
      this.bookmarksEnd.splice(index, 1);
      return true;
    }
    return false;
  }

  /**
   * Adds a comment range start marker at the beginning of this paragraph
   * @param comment - Comment to start
   * @returns This paragraph for chaining
   */
  addCommentStart(comment: Comment): this {
    this.commentsStart.push(comment);
    return this;
  }

  /**
   * Adds a comment range end marker at the end of this paragraph
   * @param comment - Comment to end (must have matching start marker)
   * @returns This paragraph for chaining
   */
  addCommentEnd(comment: Comment): this {
    this.commentsEnd.push(comment);
    return this;
  }

  /**
   * Adds both start and end comment range markers (comments entire paragraph)
   * @param comment - Comment to add
   * @returns This paragraph for chaining
   */
  addComment(comment: Comment): this {
    this.addCommentStart(comment);
    this.addCommentEnd(comment);
    return this;
  }

  /**
   * Gets all comments that start in this paragraph
   * @returns Array of comments
   */
  getCommentsStart(): Comment[] {
    return [...this.commentsStart];
  }

  /**
   * Gets all comments that end in this paragraph
   * @returns Array of comments
   */
  getCommentsEnd(): Comment[] {
    return [...this.commentsEnd];
  }

  /**
   * Adds a page number field
   * @param formatting - Optional run formatting for the page number
   * @returns This paragraph for chaining
   * @example
   * paragraph.addPageNumber();
   * paragraph.addPageNumber({ bold: true, size: 12 });
   */
  addPageNumber(formatting?: RunFormatting): this {
    return this.addField(Field.createPageNumber(formatting));
  }

  /**
   * Adds a total pages field (NUMPAGES)
   * @param formatting - Optional run formatting
   * @returns This paragraph for chaining
   * @example
   * paragraph.addTotalPages();
   */
  addTotalPages(formatting?: RunFormatting): this {
    return this.addField(Field.createTotalPages(formatting));
  }

  /**
   * Adds a date field
   * @param format - Date format (e.g., 'MMMM d, yyyy', 'M/d/yyyy')
   * @param formatting - Optional run formatting
   * @returns This paragraph for chaining
   * @example
   * paragraph.addDate();
   * paragraph.addDate('MMMM d, yyyy');
   * paragraph.addDate('M/d/yyyy', { italic: true });
   */
  addDate(format?: string, formatting?: RunFormatting): this {
    return this.addField(Field.createDate(format, formatting));
  }

  /**
   * Adds a time field
   * @param format - Time format
   * @param formatting - Optional run formatting
   * @returns This paragraph for chaining
   * @example
   * paragraph.addTime();
   * paragraph.addTime('h:mm:ss tt');
   */
  addTime(format?: string, formatting?: RunFormatting): this {
    return this.addField(Field.createTime(format, formatting));
  }

  /**
   * Adds a filename field
   * @param includePath - Whether to include full path
   * @param formatting - Optional run formatting
   * @returns This paragraph for chaining
   * @example
   * paragraph.addFilename();
   * paragraph.addFilename(true); // Includes full path
   */
  addFilename(includePath = false, formatting?: RunFormatting): this {
    return this.addField(Field.createFilename(includePath, formatting));
  }

  /**
   * Adds an author field
   * @param formatting - Optional run formatting
   * @returns This paragraph for chaining
   * @example
   * paragraph.addAuthor();
   */
  addAuthor(formatting?: RunFormatting): this {
    return this.addField(Field.createAuthor(formatting));
  }

  /**
   * Adds a title field
   * @param formatting - Optional run formatting
   * @returns This paragraph for chaining
   * @example
   * paragraph.addTitle();
   */
  addTitle(formatting?: RunFormatting): this {
    return this.addField(Field.createTitle(formatting));
  }

  /**
   * Adds text to the paragraph with optional formatting
   *
   * Creates a new Run with the specified text and formatting, then adds it
   * to the paragraph. This is a convenience method combining Run creation
   * and addition in one call.
   *
   * @param text - The text content to add
   * @param formatting - Optional formatting to apply to the text
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * const para = new Paragraph();
   * para.addText('Normal text');
   * para.addText('Bold text', { bold: true });
   * para.addText('Red italic', { italic: true, color: 'FF0000' });
   * ```
   */
  addText(text: string, formatting?: RunFormatting): this {
    const run = new Run(text, formatting);
    // Set parent reference for setText tracking
    run._setParentParagraph(this);

    if (this.trackingContext?.isEnabled()) {
      // Wrap the run in an insert revision when tracking is enabled
      const revision = Revision.createInsertion(this.trackingContext.getAuthor(), run, new Date());
      this.trackingContext.getRevisionManager().register(revision);
      this.content.push(revision);
    } else {
      this.content.push(run);
    }
    return this;
  }

  /**
   * Adds a line break to the paragraph
   *
   * Creates a Run containing a line break element (`w:br`). This produces
   * a soft return (line break within the same paragraph), unlike creating
   * a new paragraph which produces a hard return.
   *
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * para.addText('Line one');
   * para.addLineBreak();
   * para.addText('Line two (same paragraph)');
   * ```
   */
  addLineBreak(): this {
    const run = new Run('');
    run.addBreak();
    this.content.push(run);
    return this;
  }

  /**
   * Adds a column break to the paragraph
   *
   * Inserts a column break element (`w:br w:type="column"`). In multi-column
   * layouts, forces subsequent content to the next column.
   *
   * @returns This paragraph for chaining
   */
  addColumnBreak(): this {
    const run = new Run('');
    run.addBreak('column');
    this.content.push(run);
    return this;
  }

  /**
   * Sets the paragraph text content (replaces all existing content)
   *
   * Clears all existing runs, fields, hyperlinks, and other content,
   * then adds a single Run with the specified text and formatting.
   *
   * @param text - The new text content
   * @param formatting - Optional formatting to apply to the text
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * const para = new Paragraph();
   * para.addText('First text');
   * para.addText('More text');
   * para.setText('Replace all', { bold: true }); // Replaces everything
   * ```
   */
  setText(text: string, formatting?: RunFormatting): this {
    this.content = [new Run(text, formatting)];
    return this;
  }

  /**
   * Gets all Run instances in the paragraph
   *
   * Returns only Run objects, excluding other content types like fields,
   * hyperlinks, revisions, and range markers.
   *
   * @returns Array of Run instances
   *
   * @example
   * ```typescript
   * const runs = para.getRuns();
   * console.log(`Paragraph has ${runs.length} text runs`);
   *
   * // Apply formatting to all runs
   * for (const run of runs) {
   *   run.setBold(true);
   * }
   * ```
   */
  getRuns(): Run[] {
    const runs: Run[] = [];
    for (const item of this.content) {
      if (item instanceof Run) {
        runs.push(item);
      } else if (item instanceof Revision) {
        // Extract runs from inside revisions (track changes)
        runs.push(...item.getRuns());
      } else if (item instanceof Hyperlink) {
        // Extract run from inside hyperlink
        runs.push(item.getRun());
      }
    }
    return runs;
  }

  /**
   * Gets all Revision instances in the paragraph
   *
   * Returns only Revision objects (tracked changes), excluding other content types.
   *
   * @returns Array of Revision instances
   *
   * @example
   * ```typescript
   * const revisions = para.getRevisions();
   * console.log(`Paragraph has ${revisions.length} tracked changes`);
   *
   * // Check each revision
   * for (const rev of revisions) {
   *   console.log(`${rev.getType()} by ${rev.getAuthor()}`);
   * }
   * ```
   */
  getRevisions(): Revision[] {
    return this.content.filter((item): item is Revision => item instanceof Revision);
  }

  /**
   * Gets all content in the paragraph
   *
   * Returns all content items including runs, fields, hyperlinks,
   * revisions, range markers, shapes, and text boxes.
   *
   * @returns Array of all content items in order
   *
   * @example
   * ```typescript
   * const content = para.getContent();
   * for (const item of content) {
   *   if (item instanceof Run) {
   *     console.log('Text:', item.getText());
   *   } else if (item instanceof Hyperlink) {
   *     console.log('Link:', item.getUrl());
   *   }
   * }
   * ```
   */
  getContent(): ParagraphContent[] {
    return [...this.content];
  }

  /**
   * Clears all content from the paragraph
   *
   * Removes all runs, hyperlinks, fields, revisions, and other content items.
   * The paragraph formatting is preserved.
   *
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * para.clearContent();
   * para.addText('Fresh start');
   * ```
   */
  clearContent(): this {
    if (this.trackingContext?.isEnabled() && this.content.length > 0) {
      // Wrap all existing content in delete revisions instead of removing
      const deletedContent: ParagraphContent[] = [];
      for (const item of this.content) {
        // Skip items that are already revisions (don't double-wrap)
        if (item instanceof Revision) {
          deletedContent.push(item);
        } else if (item instanceof Run || item instanceof Hyperlink) {
          const revision = Revision.createDeletion(
            this.trackingContext.getAuthor(),
            item,
            new Date()
          );
          this.trackingContext.getRevisionManager().register(revision);
          deletedContent.push(revision);
        } else {
          // For other content types (fields, etc.), just keep them
          // as they may have complex structures
          deletedContent.push(item);
        }
      }
      this.content = deletedContent;
    } else {
      this.content = [];
    }
    return this;
  }

  /**
   * Replaces a content item with one or more new items
   *
   * This is useful for tracked changes where a hyperlink needs to be replaced
   * with a deletion revision (containing old hyperlink) and insertion revision
   * (containing new hyperlink).
   *
   * @param oldItem - The item to replace (must be the exact same object reference)
   * @param newItems - The new items to insert in place of the old item
   * @returns true if the item was found and replaced, false if not found
   *
   * @example
   * ```typescript
   * // Replace a hyperlink with tracked changes
   * const deletion = Revision.createDeletion(author, [oldHyperlink]);
   * const insertion = Revision.createInsertion(author, [newHyperlink]);
   * para.replaceContent(hyperlink, [deletion, insertion]);
   * ```
   */
  replaceContent(oldItem: ParagraphContent, newItems: ParagraphContent[]): boolean {
    const index = this.content.indexOf(oldItem);
    if (index === -1) {
      return false;
    }

    // Replace the item with the new items
    this.content.splice(index, 1, ...newItems);
    return true;
  }

  /**
   * Sets all content of the paragraph, replacing existing content
   *
   * Used for bulk content operations like regrouping field runs
   * into ComplexField objects during parsing.
   *
   * @param content - Array of content items to set
   *
   * @example
   * ```typescript
   * // Replace all content with grouped items
   * para.setContent([run1, complexField, run2]);
   * ```
   */
  setContent(content: ParagraphContent[]): void {
    // If tracking is enabled and we have existing content, create delete revisions
    if (this.trackingContext?.isEnabled() && this.content.length > 0) {
      const deletedContent: ParagraphContent[] = [];
      for (const item of this.content) {
        if (item instanceof Revision) {
          deletedContent.push(item);
        } else if (item instanceof Run || item instanceof Hyperlink) {
          const revision = Revision.createDeletion(
            this.trackingContext.getAuthor(),
            item,
            new Date()
          );
          this.trackingContext.getRevisionManager().register(revision);
          deletedContent.push(revision);
        }
      }
      // New content array starts with deleted items followed by new items
      // But actually, for setContent we probably want to just replace
      // Keep the deletion revisions and add new content
      this.content = [...deletedContent, ...content];
    } else {
      this.content = [...content];
    }
    // Set parent reference for runs and hyperlinks in the content
    for (const item of this.content) {
      if (item instanceof Hyperlink) {
        item._setParentParagraph(this);
      } else if (item instanceof Run) {
        item._setParentParagraph(this);
      }
    }
  }

  /**
   * Removes a content item at the specified index
   *
   * When track changes is enabled, wraps the removed item in a delete revision
   * instead of actually removing it.
   *
   * @param index - Index of the item to remove
   * @returns The removed item, or undefined if index is out of bounds
   *
   * @example
   * ```typescript
   * const para = new Paragraph();
   * para.addText('First');
   * para.addText('Second');
   * para.removeContentAt(0); // Removes 'First'
   * ```
   */
  removeContentAt(index: number): ParagraphContent | undefined {
    if (index < 0 || index >= this.content.length) {
      return undefined;
    }

    const item = this.content[index];

    if (this.trackingContext?.isEnabled()) {
      // Wrap the item in a delete revision instead of removing
      if (item instanceof Run || item instanceof Hyperlink) {
        const revision = Revision.createDeletion(
          this.trackingContext.getAuthor(),
          item,
          new Date()
        );
        this.trackingContext.getRevisionManager().register(revision);
        this.content[index] = revision;
      }
      // For revisions and other types, leave them as-is
    } else {
      this.content.splice(index, 1);
    }

    return item;
  }

  /**
   * Gets the combined text content of the paragraph
   *
   * Concatenates text from all Runs and Hyperlinks in the paragraph,
   * excluding other content types like fields or revisions.
   *
   * @returns Combined text string from all text-bearing elements
   *
   * @example
   * ```typescript
   * const para = new Paragraph();
   * para.addText('Hello ');
   * para.addText('World');
   * console.log(para.getText()); // "Hello World"
   * ```
   */
  getText(): string {
    return this.content
      .filter((item): item is Run | Hyperlink => item instanceof Run || item instanceof Hyperlink)
      .map((item) => item.getText())
      .join('');
  }

  /**
   * Gets all fields in the paragraph (both Field and ComplexField)
   *
   * Returns all field instances including simple fields (`<w:fldSimple>`)
   * and complex fields (begin/separate/end structure).
   *
   * @returns Array of fields
   *
   * @example
   * ```typescript
   * const fields = para.getFields();
   * console.log(`P aragraph has ${fields.length} fields`);
   * for (const field of fields) {
   *   console.log(`Instruction: ${field.getInstruction()}`);
   * }
   * ```
   */
  getFields(): FieldLike[] {
    return this.content.filter(
      (item): item is FieldLike => item instanceof Field || item instanceof ComplexField
    );
  }

  /**
   * Finds fields matching an instruction pattern
   *
   * Searches all fields and returns those whose instruction matches
   * the specified pattern (string or regex).
   *
   * @param pattern - Regex pattern or string to match against instruction
   * @returns Array of matching fields
   *
   * @example
   * ```typescript
   * // Find all PAGE fields
   * const pageFields = para.findFieldsByInstruction('PAGE');
   *
   * // Find all TOC fields
   * const tocFields = para.findFieldsByInstruction(/^TOC/i);
   *
   * // Find fields with specific switches
   * const hyperlinkedFields = para.findFieldsByInstruction(/\\h/);
   * ```
   */
  findFieldsByInstruction(pattern: string | RegExp): FieldLike[] {
    const regex = typeof pattern === 'string' ? new RegExp(pattern, 'i') : pattern;

    return this.getFields().filter((field) => {
      const instruction = field.getInstruction();
      return regex.test(instruction);
    });
  }

  /**
   * Removes all fields from the paragraph
   *
   * Filters out all Field and ComplexField instances, converting them
   * to plain text if they have result text.
   *
   * @returns Count of fields removed
   *
   * @example
   * ```typescript
   * const count = para.removeAllFields();
   * console.log(`Removed ${count} fields`);
   * ```
   */
  removeAllFields(): number {
    const originalLength = this.content.length;
    this.content = this.content.filter(
      (item) => !(item instanceof Field || item instanceof ComplexField)
    );
    return originalLength - this.content.length;
  }

  /**
   * Replaces a field with another field or text
   *
   * Swaps out an existing field with a replacement. If replacement is a string,
   * converts it to a Run.
   *
   * @param oldField - Field to replace
   * @param replacement - New field or text to insert
   * @returns True if replacement successful, false if field not found
   *
   * @example
   * ```typescript
   * const pageField = para.getFields()[0];
   * if (pageField) {
   *   // Replace with text
   *   para.replaceField(pageField, 'Page 1');
   *
   *   // Or replace with another field
   *   para.replaceField(pageField, Field.createDate());
   * }
   * ```
   */
  replaceField(oldField: FieldLike, replacement: FieldLike | string): boolean {
    const index = this.content.indexOf(oldField);
    if (index === -1) return false;

    if (typeof replacement === 'string') {
      this.content[index] = new Run(replacement);
    } else {
      this.content[index] = replacement;
    }
    return true;
  }

  /**
   * Gets a copy of the paragraph formatting
   *
   * Returns a copy of all formatting properties including alignment,
   * indentation, spacing, style, numbering, borders, shading, etc.
   *
   * @returns Copy of the paragraph formatting object
   *
   * @example
   * ```typescript
   * const formatting = para.getFormatting();
   * console.log(`Style: ${formatting.style}`);
   * console.log(`Alignment: ${formatting.alignment}`);
   * if (formatting.spacing) {
   *   console.log(`Spacing before: ${formatting.spacing.before} twips`);
   * }
   * ```
   */
  getFormatting(): ParagraphFormatting {
    return { ...this.formatting };
  }

  // ============================================================================
  // Individual Formatting Getters
  // ============================================================================

  /**
   * Gets the left indentation in twips
   * @returns Left indent in twips or undefined if not set
   */
  getLeftIndent(): number | undefined {
    return this.formatting.indentation?.left;
  }

  /**
   * Gets the right indentation in twips
   * @returns Right indent in twips or undefined if not set
   */
  getRightIndent(): number | undefined {
    return this.formatting.indentation?.right;
  }

  /**
   * Gets the first line indentation in twips
   * @returns First line indent in twips or undefined if not set
   */
  getFirstLineIndent(): number | undefined {
    return this.formatting.indentation?.firstLine;
  }

  /**
   * Gets the hanging indentation in twips
   * @returns Hanging indent in twips or undefined if not set
   */
  getHangingIndent(): number | undefined {
    return this.formatting.indentation?.hanging;
  }

  /**
   * Gets the spacing before the paragraph in twips
   * @returns Space before in twips or undefined if not set
   */
  getSpaceBefore(): number | undefined {
    return this.formatting.spacing?.before;
  }

  /**
   * Gets the spacing after the paragraph in twips
   * @returns Space after in twips or undefined if not set
   */
  getSpaceAfter(): number | undefined {
    return this.formatting.spacing?.after;
  }

  /**
   * Gets the line spacing value
   * @returns Line spacing value or undefined if not set
   */
  getLineSpacing(): number | undefined {
    return this.formatting.spacing?.line;
  }

  /**
   * Gets the paragraph alignment
   * @returns Alignment ('left', 'center', 'right', 'justify') or undefined
   */
  getAlignment(): string | undefined {
    return this.formatting.alignment;
  }

  /**
   * Gets the keepNext property (keep with next paragraph)
   * @returns True if keepNext is set
   */
  getKeepNext(): boolean {
    return this.formatting.keepNext ?? false;
  }

  /**
   * Gets the keepLines property (keep lines together)
   * @returns True if keepLines is set
   */
  getKeepLines(): boolean {
    return this.formatting.keepLines ?? false;
  }

  /**
   * Gets the pageBreakBefore property
   * @returns True if page break before is set
   */
  getPageBreakBefore(): boolean {
    return this.formatting.pageBreakBefore ?? false;
  }

  /**
   * Gets the outline level for TOC headings
   * @returns Outline level (0-8) or undefined if not set
   */
  getOutlineLevel(): number | undefined {
    return this.formatting.outlineLevel;
  }

  /**
   * Gets the text direction
   * @returns Text direction or undefined if not set
   */
  getTextDirection(): string | undefined {
    return this.formatting.textDirection;
  }

  /**
   * Gets the widow/orphan control setting
   * @returns True if widow control is enabled
   */
  getWidowControl(): boolean {
    return this.formatting.widowControl ?? true; // Word defaults to true
  }

  /**
   * Gets the contextual spacing setting
   * @returns True if contextual spacing is enabled
   */
  getContextualSpacing(): boolean {
    return this.formatting.contextualSpacing ?? false;
  }

  // ============================================================================
  // Checker Methods (has*, is*, isEmpty)
  // ============================================================================

  /**
   * Checks if the paragraph has list numbering applied
   * @returns True if paragraph is part of a list
   */
  hasNumbering(): boolean {
    return this.formatting.numbering?.numId !== undefined && this.formatting.numbering.numId !== 0;
  }

  /**
   * Checks if numbering was explicitly suppressed in the original XML (numId=0).
   * This is different from having no numbering — it means the document author
   * intentionally overrode style-inherited numbering.
   * @returns True if numbering is explicitly suppressed
   */
  isNumberingSuppressed(): boolean {
    return this.formatting.numberingSuppressed === true;
  }

  /**
   * Checks if the paragraph contains any fields
   * @returns True if paragraph has fields
   */
  hasFields(): boolean {
    return this.getFields().length > 0;
  }

  /**
   * Checks if the paragraph has any bookmark start markers
   * @returns True if paragraph has bookmarks
   */
  hasBookmarks(): boolean {
    return this.getBookmarksStart().length > 0;
  }

  /**
   * Checks if the paragraph has any comment start markers
   * @returns True if paragraph has comments
   */
  hasComments(): boolean {
    return this.getCommentsStart().length > 0;
  }

  /**
   * Checks if the paragraph contains any revisions
   * @returns True if paragraph has revisions
   */
  hasRevisions(): boolean {
    return this.getRevisions().length > 0;
  }

  /**
   * Consolidates adjacent revisions of the same type, author, and within a time window.
   *
   * This addresses the "random insertions and deletions" problem where Word displays
   * many small revisions instead of consolidated ones. The method merges:
   * - Adjacent insertions from the same author within 1 second
   * - Adjacent deletions from the same author within 1 second
   *
   * **How it works:**
   * - Iterates through paragraph content
   * - When finding consecutive Revisions of same type/author within time window:
   *   - Merges their content (Runs/Hyperlinks) into a single Revision
   *   - Uses the earliest timestamp for the merged revision
   * - Non-revision content acts as a boundary (stops merging)
   *
   * **Why this matters:**
   * Microsoft Word typically consolidates edits made in quick succession by the same
   * author. Without consolidation, programmatic edits create many tiny revisions that
   * clutter the document and confuse users when reviewing changes.
   *
   * @param timeWindowMs - Time window in milliseconds for consolidation (default: 1000ms)
   * @returns Number of revisions that were consolidated (merged)
   *
   * @example
   * ```typescript
   * // Before: Multiple separate insertions
   * para.addText('Hello ');  // Creates w:ins #1
   * para.addText('World');   // Creates w:ins #2
   *
   * // Consolidate into single revision
   * const merged = para.consolidateRevisions();
   * console.log(`Consolidated ${merged} revisions`);
   * // Result: Single w:ins containing "Hello World"
   * ```
   */
  consolidateRevisions(timeWindowMs = 1000): number {
    if (!this.hasRevisions()) {
      return 0;
    }

    const consolidatedContent: ParagraphContent[] = [];
    let mergeCount = 0;
    let currentMergeGroup: Revision | null = null;

    for (const item of this.content) {
      if (item instanceof Revision) {
        // Check if we can merge with current group
        if (currentMergeGroup && this.canMergeRevisions(currentMergeGroup, item, timeWindowMs)) {
          // Merge this revision into the current group
          for (const content of item.getContent()) {
            currentMergeGroup.addContent(content);
          }
          mergeCount++;
        } else {
          // Start a new merge group
          if (currentMergeGroup) {
            consolidatedContent.push(currentMergeGroup);
          }
          // Clone the revision to avoid modifying the original
          currentMergeGroup = this.cloneRevision(item);
        }
      } else {
        // Non-revision content - acts as boundary
        if (currentMergeGroup) {
          consolidatedContent.push(currentMergeGroup);
          currentMergeGroup = null;
        }
        consolidatedContent.push(item);
      }
    }

    // Don't forget the last merge group
    if (currentMergeGroup) {
      consolidatedContent.push(currentMergeGroup);
    }

    // Only update content if we actually merged something
    if (mergeCount > 0) {
      this.content = consolidatedContent;
    }

    return mergeCount;
  }

  /**
   * Checks if two revisions can be merged based on type, author, and time window.
   * @private
   */
  private canMergeRevisions(rev1: Revision, rev2: Revision, timeWindowMs: number): boolean {
    // Must be same type (both insertions or both deletions)
    if (rev1.getType() !== rev2.getType()) {
      return false;
    }

    // Only merge insert/delete content revisions, not property changes
    const mergeableTypes: string[] = ['insert', 'delete', 'moveFrom', 'moveTo'];
    if (!mergeableTypes.includes(rev1.getType())) {
      return false;
    }

    // Must be same author
    if (rev1.getAuthor() !== rev2.getAuthor()) {
      return false;
    }

    // Must be within time window
    const timeDiff = Math.abs(rev1.getDate().getTime() - rev2.getDate().getTime());
    if (timeDiff > timeWindowMs) {
      return false;
    }

    return true;
  }

  /**
   * Creates a shallow clone of a revision with cloned content array.
   * @private
   */
  private cloneRevision(revision: Revision): Revision {
    const cloned = new Revision({
      id: revision.getId(),
      author: revision.getAuthor(),
      date: revision.getDate(),
      type: revision.getType(),
      content: [...revision.getContent()],
      previousProperties: revision.getPreviousProperties(),
      newProperties: revision.getNewProperties(),
      moveId: revision.getMoveId(),
      moveLocation: revision.getMoveLocation(),
      location: revision.getLocation(),
      fieldContext: revision.getFieldContext(),
    });
    return cloned;
  }

  /**
   * Checks if the paragraph is empty (no text content)
   * @returns True if paragraph has no text
   */
  isEmpty(): boolean {
    return this.getText().trim().length === 0;
  }

  /**
   * Checks whether the paragraph text contains a substring
   *
   * Case-insensitive by default. A simpler alternative to `findText()`
   * when you just need a boolean check.
   *
   * @param text - Substring to search for
   * @param caseSensitive - Match case exactly (default: false)
   * @returns True if the paragraph contains the text
   *
   * @example
   * ```typescript
   * if (para.contains('TODO')) {
   *   console.log('Found a TODO item');
   * }
   * ```
   */
  contains(text: string, caseSensitive = false): boolean {
    const paraText = caseSensitive ? this.getText() : this.getText().toLowerCase();
    const search = caseSensitive ? text : text.toLowerCase();
    return paraText.includes(search);
  }

  /**
   * Gets the paragraph style ID
   *
   * Returns the style identifier if a style is applied to this paragraph.
   *
   * @returns Style ID (e.g., 'Heading1', 'Normal') or undefined if no style is set
   *
   * @example
   * ```typescript
   * const style = para.getStyle();
   * if (style === 'Heading1') {
   *   console.log('This is a heading paragraph');
   * }
   * ```
   */
  getStyle(): string | undefined {
    return this.formatting.style;
  }

  /**
   * Detects if this paragraph is a heading and returns its level (1-9)
   *
   * Detection is performed using multiple methods in order of reliability:
   * 1. **Style ID matching** - Checks for 'Heading1', 'Heading2', etc. (most reliable)
   * 2. **Outline level** - Uses ECMA-376 outlineLevel property (0-based, where 0 = H1)
   * 3. **Formatting heuristics** - Analyzes text size, boldness, and patterns
   *
   * @returns Heading level (1-9) if detected, null if not a heading
   *
   * @example
   * ```typescript
   * const para = doc.createParagraph('Chapter 1');
   * para.setStyle('Heading1');
   *
   * const level = para.detectHeadingLevel();
   * console.log(level); // 1
   *
   * // Use for TOC generation
   * for (const para of doc.getParagraphs()) {
   *   const level = para.detectHeadingLevel();
   *   if (level) {
   *     console.log(`H${level}: ${para.getText()}`);
   *   }
   * }
   * ```
   */
  detectHeadingLevel(): number | null {
    // Method 1: Check style ID (most reliable)
    const style = this.formatting.style;
    if (style) {
      const match = /^Heading(\d)$/i.exec(style);
      if (match?.[1]) {
        const level = parseInt(match[1], 10);
        // Word supports Heading1-Heading9
        return level >= 1 && level <= 9 ? level : null;
      }
    }

    // Method 2: Check outline level (ECMA-376 property)
    // Per ECMA-376 Part 1 §17.3.1.20, outlineLevel is 0-based (0 = Level 1)
    if (this.formatting.outlineLevel !== undefined) {
      const level = this.formatting.outlineLevel + 1;
      return level >= 1 && level <= 9 ? level : null;
    }

    // Method 3: Formatting heuristics (least reliable, use as fallback)
    // Only attempt if paragraph has text
    const text = this.getText().trim();
    if (!text) return null;

    // Get formatting from first run (headings typically have uniform formatting)
    const runs = this.getRuns();
    if (runs.length === 0) return null;

    const firstRun = runs[0];
    if (!firstRun) return null;

    const fmt = firstRun.getFormatting();
    if (!fmt) return null;

    // Heuristic: Large bold text suggests a heading
    if (fmt.bold && fmt.size) {
      // H1: Very large (>= 24pt)
      if (fmt.size >= 24) return 1;
      // H2: Large (>= 20pt)
      if (fmt.size >= 20) return 2;
      // H3: Medium-large (>= 16pt)
      if (fmt.size >= 16) return 3;
      // H4-H6: Moderate size (14pt+)
      if (fmt.size >= 14) return 4;
    }

    // Additional heuristic: All caps bold text (common for headings)
    if (fmt.bold && fmt.allCaps && text.length < 100) {
      return 2; // Often used for section headings
    }

    // Not detected as a heading
    return null;
  }

  /**
   * Sets paragraph text alignment
   *
   * Controls how text is aligned within the paragraph boundaries.
   *
   * @param alignment - Alignment value ('left' | 'center' | 'right' | 'justify' | 'both')
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * para.setAlignment('center');  // Center-aligned
   * para.setAlignment('justify'); // Justified text
   * ```
   */
  setAlignment(alignment: ParagraphAlignment): this {
    const previousValue = this.formatting.alignment;
    this.formatting.alignment = alignment;
    if (this.trackingContext?.isEnabled() && previousValue !== alignment) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'alignment',
        previousValue,
        alignment
      );
    }
    return this;
  }

  /**
   * Sets left indentation
   *
   * WARNING: If this paragraph has numbering (is part of a list), setting
   * left indentation will be ignored as numbering controls indentation.
   * Use setNumbering() with different levels to change list indentation.
   *
   * @param twips - Indentation in twips (1/20th of a point)
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * // For regular paragraphs
   * paragraph.setLeftIndent(720); // 0.5 inch indent
   *
   * // For list items, use numbering levels instead
   * paragraph.setNumbering(listId, 1); // Increases indent to level 1
   * ```
   */
  setLeftIndent(twips: number): this {
    if (this.formatting.numbering) {
      // Note: This will be cleared when setNumbering() was called or will be on next call
      // Still allow setting for edge cases, but it will have no effect
      defaultLogger.warn(
        'Setting left indentation on a numbered paragraph has no effect. ' +
          'Numbering controls indentation. Use different numbering levels to change indent.'
      );
    }
    const previousValue = this.formatting.indentation?.left;
    if (!this.formatting.indentation) {
      this.formatting.indentation = {};
    }
    this.formatting.indentation.left = twips;
    if (this.trackingContext?.isEnabled() && previousValue !== twips) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'indentation.left',
        previousValue,
        twips
      );
    }
    return this;
  }

  /**
   * Sets right indentation
   * @param twips - Indentation in twips
   * @returns This paragraph for chaining
   */
  setRightIndent(twips: number): this {
    const previousValue = this.formatting.indentation?.right;
    if (!this.formatting.indentation) {
      this.formatting.indentation = {};
    }
    this.formatting.indentation.right = twips;
    if (this.trackingContext?.isEnabled() && previousValue !== twips) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'indentation.right',
        previousValue,
        twips
      );
    }
    return this;
  }

  /**
   * Sets first line indentation
   *
   * WARNING: If this paragraph has numbering (is part of a list), setting
   * first line indentation will be ignored as numbering controls indentation.
   * Numbered lists use hanging indentation for proper alignment.
   *
   * @param twips - Indentation in twips
   * @returns This paragraph for chaining
   */
  setFirstLineIndent(twips: number): this {
    if (this.formatting.numbering) {
      defaultLogger.warn(
        'Setting first line indentation on a numbered paragraph has no effect. ' +
          'Numbering controls indentation using hanging indent.'
      );
    }
    const previousValue = this.formatting.indentation?.firstLine;
    if (!this.formatting.indentation) {
      this.formatting.indentation = {};
    }
    this.formatting.indentation.firstLine = twips;
    if (this.trackingContext?.isEnabled() && previousValue !== twips) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'indentation.firstLine',
        previousValue,
        twips
      );
    }
    return this;
  }

  /**
   * Sets hanging indentation
   *
   * Creates a hanging indent where the first line starts at the left margin
   * and subsequent lines are indented. Common for bulleted/numbered lists
   * and bibliographies.
   *
   * WARNING: If this paragraph has numbering (is part of a list), setting
   * hanging indentation manually will be ignored as numbering controls indentation.
   *
   * @param twips - Indentation in twips (must be non-negative)
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * // Create hanging indent of 0.5 inch (720 twips)
   * paragraph.setHangingIndent(720);
   * ```
   */
  setHangingIndent(twips: number): this {
    if (twips < 0) {
      throw new Error('Hanging indent must be non-negative');
    }
    if (this.formatting.numbering) {
      defaultLogger.warn(
        'Setting hanging indentation on a numbered paragraph has no effect. ' +
          'Numbering controls indentation.'
      );
    }
    const previousValue = this.formatting.indentation?.hanging;
    if (!this.formatting.indentation) {
      this.formatting.indentation = {};
    }
    this.formatting.indentation.hanging = twips;
    if (this.trackingContext?.isEnabled() && previousValue !== twips) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'indentation.hanging',
        previousValue,
        twips
      );
    }
    return this;
  }

  /**
   * Sets spacing before the paragraph
   *
   * Controls the vertical space above the paragraph.
   * 1 point = 20 twips, so 120 twips = 6pt spacing.
   *
   * @param twips - Spacing value in twips (1/20th of a point)
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * para.setSpaceBefore(240);  // 12pt (240 twips) before paragraph
   * para.setSpaceBefore(0);    // No space before
   * ```
   */
  setSpaceBefore(twips: number): this {
    const previousValue = this.formatting.spacing?.before;
    if (!this.formatting.spacing) {
      this.formatting.spacing = {};
    }
    this.formatting.spacing.before = twips;
    if (this.trackingContext?.isEnabled() && previousValue !== twips) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'spacing.before',
        previousValue,
        twips
      );
    }
    return this;
  }

  /**
   * Sets spacing after the paragraph
   *
   * Controls the vertical space below the paragraph.
   * 1 point = 20 twips, so 120 twips = 6pt spacing.
   *
   * @param twips - Spacing value in twips (1/20th of a point)
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * para.setSpaceAfter(120);  // 6pt (120 twips) after paragraph
   * para.setSpaceAfter(0);    // No space after
   * ```
   */
  setSpaceAfter(twips: number): this {
    const previousValue = this.formatting.spacing?.after;
    if (!this.formatting.spacing) {
      this.formatting.spacing = {};
    }
    this.formatting.spacing.after = twips;
    if (this.trackingContext?.isEnabled() && previousValue !== twips) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'spacing.after',
        previousValue,
        twips
      );
    }
    return this;
  }

  /**
   * Sets line spacing within the paragraph
   *
   * Controls the vertical space between lines of text within the paragraph.
   * Per ECMA-376 Part 1 §17.3.1.33 (w:spacing/@w:line)
   *
   * @param twips - Line spacing value in twips (1/20th of a point)
   * @param rule - Line spacing rule (default: 'auto')
   *   - 'auto': Line spacing is based on font size (240 twips = single spacing)
   *   - 'exact': Exact line height regardless of font size
   *   - 'atLeast': Minimum line height, expands if content is larger
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * para.setLineSpacing(240, 'auto');    // Single spacing
   * para.setLineSpacing(360, 'auto');    // 1.5 spacing
   * para.setLineSpacing(480, 'auto');    // Double spacing
   * para.setLineSpacing(300, 'exact');   // Exactly 15pt line height
   * ```
   */
  setLineSpacing(twips: number, rule: 'auto' | 'exact' | 'atLeast' = 'auto'): this {
    const previousLine = this.formatting.spacing?.line;
    const previousRule = this.formatting.spacing?.lineRule;
    if (!this.formatting.spacing) {
      this.formatting.spacing = {};
    }
    this.formatting.spacing.line = twips;
    this.formatting.spacing.lineRule = rule;
    if (this.trackingContext?.isEnabled()) {
      if (previousLine !== twips) {
        this.trackingContext.trackParagraphPropertyChange(
          this,
          'spacing.line',
          previousLine,
          twips
        );
      }
      if (previousRule !== rule) {
        this.trackingContext.trackParagraphPropertyChange(
          this,
          'spacing.lineRule',
          previousRule,
          rule
        );
      }
    }
    return this;
  }

  /**
   * Clears all direct spacing properties from the paragraph.
   *
   * Removes the spacing object (before, after, line, lineRule) so the paragraph
   * inherits spacing from its style definition. This is different from setting
   * spacing to 0 — clearing means "use style value", while 0 means "no spacing".
   *
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * // Clear direct spacing so Normal style's 3pt/3pt takes effect
   * para.clearSpacing();
   * ```
   */
  clearSpacing(): this {
    delete this.formatting.spacing;
    return this;
  }

  /**
   * Sets the paragraph style
   *
   * Applies a style definition to the paragraph. The style must exist
   * in the document's StylesManager.
   *
   * @param styleId - The style identifier (e.g., 'Heading1', 'Normal', 'ListParagraph')
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * para.setStyle('Heading1');  // Apply Heading1 style
   * para.setStyle('Normal');    // Apply Normal style
   * ```
   */
  setStyle(styleId: string): this {
    const previousValue = this.formatting.style;
    this.formatting.style = styleId;
    if (this.trackingContext?.isEnabled() && previousValue !== styleId) {
      this.trackingContext.trackParagraphPropertyChange(this, 'style', previousValue, styleId);
    }
    return this;
  }

  /**
   * Sets keep with next
   *
   * **Automatic Conflict Resolution:**
   * When setting keepNext to true, pageBreakBefore is automatically cleared to prevent
   * layout conflicts. The pageBreakBefore property creates contradictory constraints
   * (breaking the page vs. keeping content together) that cause massive whitespace
   * in Word as the layout engine tries to satisfy both. Keep properties take priority
   * as they represent explicit user intent to keep content together.
   *
   * @param keepNext - Whether to keep with next paragraph
   * @returns This paragraph for chaining
   */
  setKeepNext(keepNext = true): this {
    const previousValue = this.formatting.keepNext;
    this.formatting.keepNext = keepNext;

    // Resolve property conflicts: keepNext contradicts pageBreakBefore
    if (keepNext) {
      this.formatting.pageBreakBefore = undefined;
    }

    if (this.trackingContext?.isEnabled() && previousValue !== keepNext) {
      this.trackingContext.trackParagraphPropertyChange(this, 'keepNext', previousValue, keepNext);
    }
    return this;
  }

  /**
   * Sets keep lines together
   *
   * **Automatic Conflict Resolution:**
   * When setting keepLines to true, pageBreakBefore is automatically cleared to prevent
   * layout conflicts. The pageBreakBefore property creates contradictory constraints
   * (breaking the page vs. keeping lines together) that cause massive whitespace
   * in Word as the layout engine tries to satisfy both. Keep properties take priority
   * as they represent explicit user intent to keep content together.
   *
   * @param keepLines - Whether to keep lines together
   * @returns This paragraph for chaining
   */
  setKeepLines(keepLines = true): this {
    const previousValue = this.formatting.keepLines;
    this.formatting.keepLines = keepLines;

    // Resolve property conflicts: keepLines contradicts pageBreakBefore
    if (keepLines) {
      this.formatting.pageBreakBefore = undefined;
    }

    if (this.trackingContext?.isEnabled() && previousValue !== keepLines) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'keepLines',
        previousValue,
        keepLines
      );
    }
    return this;
  }

  /**
   * Sets page break before
   * @param pageBreakBefore - Whether to insert page break before
   * @returns This paragraph for chaining
   */
  setPageBreakBefore(pageBreakBefore = true): this {
    const previousValue = this.formatting.pageBreakBefore;
    this.formatting.pageBreakBefore = pageBreakBefore;
    if (this.trackingContext?.isEnabled() && previousValue !== pageBreakBefore) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'pageBreakBefore',
        previousValue,
        pageBreakBefore
      );
    }
    return this;
  }

  /**
   * Marks this paragraph as preserved to prevent automatic removal by document processing operations
   * (e.g., removing extra blank paragraphs). Useful for spacing paragraphs that should remain
   * even if they appear to be "extra" blank lines.
   * @param preserved - Whether to preserve this paragraph
   * @returns This paragraph for chaining
   */
  setPreserved(preserved = true): this {
    this._isPreserved = preserved;
    return this;
  }

  /**
   * Checks if this paragraph is marked as preserved from automatic removal
   * @returns True if paragraph should be preserved from removal operations
   */
  isPreserved(): boolean {
    return this._isPreserved;
  }

  /**
   * Sets numbering for this paragraph (adds to a list)
   *
   * When numbering is applied, any conflicting paragraph indentation
   * (left, firstLine, hanging) is automatically cleared to prevent
   * override issues. Right indentation is preserved as it doesn't
   * conflict with list formatting.
   *
   * This matches Microsoft Word behavior where numbering controls
   * the indentation, not paragraph-level formatting.
   *
   * @param numId - The numbering instance ID
   * @param level - The level (0-8, where 0 is the outermost level)
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * const listId = doc.createBulletList();
   * paragraph.setNumbering(listId, 0); // Level 0, indent controlled by numbering
   * paragraph.setNumbering(listId, 1); // Level 1, deeper indent
   * ```
   */
  setNumbering(numId: number, level = 0): this {
    if (numId < 0) {
      throw new Error('Numbering ID must be non-negative');
    }
    if (level < 0 || level > 8) {
      throw new Error('Level must be between 0 and 8');
    }

    const previousValue = this.formatting.numbering;
    this.formatting.numbering = { numId, level };
    delete this.formatting.numberingSuppressed;

    // Clear conflicting indentation properties
    // Per ECMA-376 §17.3.1.12, paragraph indentation overrides numbering indentation
    // To prevent unexpected behavior, we clear left/firstLine/hanging when numbering is applied
    // This matches Microsoft Word behavior where numbering controls indentation
    if (this.formatting.indentation) {
      const { right } = this.formatting.indentation;
      // Preserve right indent only (doesn't conflict with numbering)
      this.formatting.indentation = right !== undefined ? { right } : undefined;
    }

    if (this.trackingContext?.isEnabled()) {
      const newValue = { numId, level };
      if (previousValue?.numId !== newValue.numId || previousValue?.level !== newValue.level) {
        this.trackingContext.trackParagraphPropertyChange(
          this,
          'numbering',
          previousValue,
          newValue
        );
      }
    }
    return this;
  }

  /**
   * Sets contextual spacing for this paragraph
   * When enabled, removes spacing between consecutive paragraphs of the same style
   * Per ECMA-376 Part 1 §17.3.1.8
   * @param enable - Whether to enable contextual spacing
   * @returns This paragraph for chaining
   */
  setContextualSpacing(enable = true): this {
    const previousValue = this.formatting.contextualSpacing;
    this.formatting.contextualSpacing = enable;
    if (this.trackingContext?.isEnabled() && previousValue !== enable) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'contextualSpacing',
        previousValue,
        enable
      );
    }
    return this;
  }

  /**
   * Removes numbering from this paragraph
   * @returns This paragraph for chaining
   */
  removeNumbering(): this {
    delete this.formatting.numbering;
    delete this.formatting.numberingSuppressed;
    return this;
  }

  /**
   * Gets the numbering properties
   * @returns Numbering properties or undefined
   */
  getNumbering(): { numId: number; level: number } | undefined {
    return this.formatting.numbering ? { ...this.formatting.numbering } : undefined;
  }

  /**
   * Sets widow/orphan control for this paragraph
   * Controls whether to prevent single lines at the top or bottom of a page.
   * Word's default is true - set to false to allow widows/orphans.
   * Per ECMA-376 Part 1 §17.3.1.40
   * @param enable - Whether to enable widow/orphan control
   * @returns This paragraph for chaining
   */
  setWidowControl(enable = true): this {
    const previousValue = this.formatting.widowControl;
    this.formatting.widowControl = enable;
    if (this.trackingContext?.isEnabled() && previousValue !== enable) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'widowControl',
        previousValue,
        enable
      );
    }
    return this;
  }

  /**
   * Sets outline level for this paragraph (for table of contents)
   * Level 0-9 indicates hierarchy in document structure.
   * Level 0 = highest level (like Heading 1)
   * Level 9 = lowest level
   * Per ECMA-376 Part 1 §17.3.1.19
   * @param level - Outline level (0-9)
   * @returns This paragraph for chaining
   */
  setOutlineLevel(level: number): this {
    if (level < 0 || level > 9) {
      throw new Error('Outline level must be between 0 and 9');
    }
    const previousValue = this.formatting.outlineLevel;
    this.formatting.outlineLevel = level;
    if (this.trackingContext?.isEnabled() && previousValue !== level) {
      this.trackingContext.trackParagraphPropertyChange(this, 'outlineLevel', previousValue, level);
    }
    return this;
  }

  /**
   * Sets whether to suppress line numbers for this paragraph
   * Per ECMA-376 Part 1 §17.3.1.34
   * @param suppress - Whether to suppress line numbers
   * @returns This paragraph for chaining
   */
  setSuppressLineNumbers(suppress = true): this {
    const previousValue = this.formatting.suppressLineNumbers;
    this.formatting.suppressLineNumbers = suppress;
    if (this.trackingContext?.isEnabled() && previousValue !== suppress) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'suppressLineNumbers',
        previousValue,
        suppress
      );
    }
    return this;
  }

  /**
   * Sets bidirectional text layout (right-to-left)
   * Enables right-to-left paragraph layout for languages like Arabic and Hebrew.
   * Per ECMA-376 Part 1 §17.3.1.6
   * @param enable - Whether to enable bidirectional (RTL) layout
   * @returns This paragraph for chaining
   */
  setBidi(enable = true): this {
    const previousValue = this.formatting.bidi;
    this.formatting.bidi = enable;
    if (this.trackingContext?.isEnabled() && previousValue !== enable) {
      this.trackingContext.trackParagraphPropertyChange(this, 'bidi', previousValue, enable);
    }
    return this;
  }

  /**
   * Sets text flow direction for this paragraph
   * Per ECMA-376 Part 1 §17.3.1.36
   * @param direction - Text flow direction
   *   - 'lrTb': Left-to-right, top-to-bottom (default for English)
   *   - 'tbRl': Top-to-bottom, right-to-left (traditional Chinese/Japanese)
   *   - 'btLr': Bottom-to-top, left-to-right (Mongolian)
   *   - 'lrTbV': Left-to-right, top-to-bottom vertical
   *   - 'tbRlV': Top-to-bottom, right-to-left vertical
   *   - 'tbLrV': Top-to-bottom, left-to-right vertical
   * @returns This paragraph for chaining
   */
  setTextDirection(direction: TextDirection): this {
    const previousValue = this.formatting.textDirection;
    this.formatting.textDirection = direction;
    if (this.trackingContext?.isEnabled() && previousValue !== direction) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'textDirection',
        previousValue,
        direction
      );
    }
    return this;
  }

  /**
   * Sets vertical text alignment for this paragraph
   * Per ECMA-376 Part 1 §17.3.1.35
   * @param alignment - Vertical alignment
   *   - 'top': Align to top of line
   *   - 'center': Align to center of line
   *   - 'baseline': Align to baseline
   *   - 'bottom': Align to bottom of line
   *   - 'auto': Automatic alignment
   * @returns This paragraph for chaining
   */
  setTextAlignment(alignment: TextAlignment): this {
    const previousValue = this.formatting.textAlignment;
    this.formatting.textAlignment = alignment;
    if (this.trackingContext?.isEnabled() && previousValue !== alignment) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'textAlignment',
        previousValue,
        alignment
      );
    }
    return this;
  }

  /**
   * Sets mirror indents for this paragraph
   * When enabled, uses inside/outside indents instead of left/right for double-sided printing.
   * Per ECMA-376 Part 1 §17.3.1.18
   * @param enable - Whether to enable mirror indents
   * @returns This paragraph for chaining
   */
  setMirrorIndents(enable = true): this {
    const previousValue = this.formatting.mirrorIndents;
    this.formatting.mirrorIndents = enable;
    if (this.trackingContext?.isEnabled() && previousValue !== enable) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'mirrorIndents',
        previousValue,
        enable
      );
    }
    return this;
  }

  /**
   * Sets auto-adjust right indent for this paragraph
   * When enabled, automatically adjusts right indent when a document grid is defined.
   * Per ECMA-376 Part 1 §17.3.1.1
   * @param enable - Whether to enable auto-adjust right indent
   * @returns This paragraph for chaining
   */
  setAdjustRightInd(enable = true): this {
    const previousValue = this.formatting.adjustRightInd;
    this.formatting.adjustRightInd = enable;
    if (this.trackingContext?.isEnabled() && previousValue !== enable) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'adjustRightInd',
        previousValue,
        enable
      );
    }
    return this;
  }

  /**
   * Sets text frame/box properties for this paragraph
   * Text frames allow positioning paragraphs in specific locations with text wrapping.
   * Per ECMA-376 Part 1 §17.3.1.11
   * @param props - Frame properties
   *   - w, h: Width and height in twips
   *   - hRule: Height rule ('auto', 'atLeast', 'exact')
   *   - x, y: Absolute positioning in twips
   *   - xAlign, yAlign: Relative alignment
   *   - hAnchor, vAnchor: Positioning base ('page', 'margin', 'text')
   *   - hSpace, vSpace: Padding around frame in twips
   *   - wrap: Text wrapping ('around', 'notBeside', 'none', 'tight')
   *   - dropCap: Drop cap style ('none', 'drop', 'margin')
   *   - lines: Drop cap height in lines
   *   - anchorLock: Lock frame anchor to paragraph
   * @returns This paragraph for chaining
   */
  setFrameProperties(props: FrameProperties): this {
    const previousValue = this.formatting.framePr;
    this.formatting.framePr = props;
    if (this.trackingContext?.isEnabled() && previousValue !== props) {
      this.trackingContext.trackParagraphPropertyChange(this, 'framePr', previousValue, props);
    }
    return this;
  }

  /**
   * Suppress automatic hyphenation for this paragraph
   * Per ECMA-376 Part 1 §17.3.1.33
   * @param suppress - Whether to suppress hyphenation (default: true)
   * @returns This paragraph for chaining
   */
  setSuppressAutoHyphens(suppress = true): this {
    const previousValue = this.formatting.suppressAutoHyphens;
    this.formatting.suppressAutoHyphens = suppress;
    if (this.trackingContext?.isEnabled() && previousValue !== suppress) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'suppressAutoHyphens',
        previousValue,
        suppress
      );
    }
    return this;
  }

  /**
   * Sets CJK kinsoku line-breaking rules
   * Per ECMA-376 Part 1 §17.3.1.16
   * @param enable - Whether to enable kinsoku rules (default: true)
   * @returns This paragraph for chaining
   */
  setKinsoku(enable = true): this {
    this.formatting.kinsoku = enable;
    return this;
  }

  /**
   * Sets CJK word wrap behavior
   * Per ECMA-376 Part 1 §17.3.1.45
   * @param enable - Whether to allow wrapping mid-word (default: true)
   * @returns This paragraph for chaining
   */
  setWordWrap(enable = true): this {
    this.formatting.wordWrap = enable;
    return this;
  }

  /**
   * Sets CJK overflow punctuation
   * Per ECMA-376 Part 1 §17.3.1.24
   * @param enable - Whether to allow punctuation overhang (default: true)
   * @returns This paragraph for chaining
   */
  setOverflowPunct(enable = true): this {
    this.formatting.overflowPunct = enable;
    return this;
  }

  /**
   * Sets CJK top line punctuation compression
   * Per ECMA-376 Part 1 §17.3.1.43
   * @param enable - Whether to compress punctuation at line start (default: true)
   * @returns This paragraph for chaining
   */
  setTopLinePunct(enable = true): this {
    this.formatting.topLinePunct = enable;
    return this;
  }

  /**
   * Sets auto space between East Asian and numeric text
   * Per ECMA-376 Part 1 §17.3.1.2
   * @param enable - Whether to auto-space (default: true)
   * @returns This paragraph for chaining
   */
  setAutoSpaceDE(enable = true): this {
    this.formatting.autoSpaceDE = enable;
    return this;
  }

  /**
   * Sets auto space between East Asian and Western text
   * Per ECMA-376 Part 1 §17.3.1.3
   * @param enable - Whether to auto-space (default: true)
   * @returns This paragraph for chaining
   */
  setAutoSpaceDN(enable = true): this {
    this.formatting.autoSpaceDN = enable;
    return this;
  }

  /**
   * Prevent text frames from overlapping with this paragraph
   * Per ECMA-376 Part 1 §17.3.1.34
   * @param suppress - Whether to prevent overlap (default: true)
   * @returns This paragraph for chaining
   */
  setSuppressOverlap(suppress = true): this {
    const previousValue = this.formatting.suppressOverlap;
    this.formatting.suppressOverlap = suppress;
    if (this.trackingContext?.isEnabled() && previousValue !== suppress) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'suppressOverlap',
        previousValue,
        suppress
      );
    }
    return this;
  }

  /**
   * Sets tight wrapping mode for text boxes
   * Controls how tightly surrounding text wraps around text box content.
   * Per ECMA-376 Part 1 §17.3.1.37
   * @param wrap - Tight wrap mode
   *   - 'none': No tight wrapping
   *   - 'allLines': Tight wrap all lines
   *   - 'firstAndLastLine': Tight wrap first and last lines only
   *   - 'firstLineOnly': Tight wrap first line only
   *   - 'lastLineOnly': Tight wrap last line only
   * @returns This paragraph for chaining
   */
  setTextboxTightWrap(wrap: TextboxTightWrap): this {
    const previousValue = this.formatting.textboxTightWrap;
    this.formatting.textboxTightWrap = wrap;
    if (this.trackingContext?.isEnabled() && previousValue !== wrap) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'textboxTightWrap',
        previousValue,
        wrap
      );
    }
    return this;
  }

  /**
   * Sets the HTML div ID associated with this paragraph
   * Used for HTML round-trip conversion to preserve div structure.
   * Per ECMA-376 Part 1 §17.3.1.9
   * @param id - Decimal ID referencing a div in the web settings part
   * @returns This paragraph for chaining
   */
  setDivId(id: number): this {
    const previousValue = this.formatting.divId;
    this.formatting.divId = id;
    if (this.trackingContext?.isEnabled() && previousValue !== id) {
      this.trackingContext.trackParagraphPropertyChange(this, 'divId', previousValue, id);
    }
    return this;
  }

  /**
   * Sets conditional table style formatting for this paragraph
   * Used to apply conditional formatting based on table position (first row, last column, etc.).
   * Per ECMA-376 Part 1 §17.3.1.8
   * @param bitmask - Bitmask string (e.g., "101000000100")
   *   Each bit represents a conditional formatting property:
   *   - Bit 0: First row
   *   - Bit 1: Last row
   *   - Bit 2: First column
   *   - Bit 3: Last column
   *   - Bit 4: Band 1 vertical
   *   - Bit 5: Band 2 vertical
   *   - Bit 6: Band 1 horizontal
   *   - Bit 7: Band 2 horizontal
   *   - Bit 8-11: Corner cells (NE, NW, SE, SW)
   * @returns This paragraph for chaining
   */
  setConditionalFormatting(bitmask: string): this {
    const previousValue = this.formatting.cnfStyle;
    this.formatting.cnfStyle = bitmask;
    if (this.trackingContext?.isEnabled() && previousValue !== bitmask) {
      this.trackingContext.trackParagraphPropertyChange(this, 'cnfStyle', previousValue, bitmask);
    }
    return this;
  }

  /**
   * Sets section properties for this paragraph
   * Used to define section breaks and section-specific formatting.
   * Per ECMA-376 Part 1 §17.3.1.30
   * @param properties - Section properties object
   * @returns This paragraph for chaining
   */
  setSectionProperties(properties: string | Record<string, unknown>): this {
    const previousValue = this.formatting.sectPr;
    this.formatting.sectPr = properties;
    if (this.trackingContext?.isEnabled() && previousValue !== properties) {
      this.trackingContext.trackParagraphPropertyChange(this, 'sectPr', previousValue, properties);
    }
    return this;
  }

  /**
   * Sets paragraph property change tracking information
   * Used for revision history and change tracking.
   * Per ECMA-376 Part 1 §17.3.1.27
   * @param change - Change tracking information
   * @returns This paragraph for chaining
   */
  setParagraphPropertiesChange(change: ParagraphPropertiesChange): this {
    const previousValue = this.formatting.pPrChange;
    this.formatting.pPrChange = change;
    if (this.trackingContext?.isEnabled() && previousValue !== change) {
      this.trackingContext.trackParagraphPropertyChange(this, 'pPrChange', previousValue, change);
    }
    return this;
  }

  /**
   * Clears paragraph property change tracking information.
   * Used when accepting revisions to remove the w:pPrChange element.
   * @returns This paragraph for chaining
   */
  clearParagraphPropertiesChange(): this {
    delete this.formatting.pPrChange;
    return this;
  }

  /**
   * Sets run properties for the paragraph mark (¶ symbol)
   *
   * The paragraph mark is the invisible character at the end of every paragraph.
   * It can have its own formatting independent of the text runs in the paragraph.
   * This is useful for controlling formatting behavior when text is inserted after
   * the paragraph mark.
   *
   * Per ECMA-376 Part 1 §17.3.1.29 (Run Properties for the Paragraph Mark)
   *
   * Common use cases:
   * - Set default font for new text added to paragraph
   * - Control paragraph mark visibility in "Show/Hide ¶" mode
   * - Apply highlighting to paragraph mark for visual consistency
   *
   * @param properties - Run formatting properties for the paragraph mark
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * // Set paragraph mark to be red and bold
   * paragraph.setParagraphMarkFormatting({ bold: true, color: 'FF0000' });
   *
   * // Hide paragraph mark
   * paragraph.setParagraphMarkFormatting({ vanish: true });
   *
   * // Set default font for new text in this paragraph
   * paragraph.setParagraphMarkFormatting({ font: 'Arial', size: 12 });
   * ```
   */
  setParagraphMarkFormatting(properties: RunFormatting): this {
    const previousValue = this.formatting.paragraphMarkRunProperties;
    this.formatting.paragraphMarkRunProperties = properties;
    if (this.trackingContext?.isEnabled() && previousValue !== properties) {
      this.trackingContext.trackParagraphPropertyChange(
        this,
        'paragraphMarkRunProperties',
        previousValue,
        properties
      );
    }
    return this;
  }

  /**
   * Marks the paragraph mark as deleted (tracked change)
   *
   * When a paragraph mark is deleted, it indicates that the paragraph
   * was joined with the next paragraph. Word shows this as a deletion
   * of the ¶ symbol.
   *
   * @param id - Unique revision ID
   * @param author - Author who deleted the paragraph mark
   * @param date - Date when the deletion occurred (defaults to now)
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * paragraph.markParagraphMarkAsDeleted(1, 'Alice', new Date());
   * ```
   */
  markParagraphMarkAsDeleted(id: number, author: string, date?: Date): this {
    this.formatting.paragraphMarkDeletion = {
      id,
      author,
      date: date || new Date(),
    };
    return this;
  }

  /**
   * Clears the paragraph mark deletion marker
   * @returns This paragraph for chaining
   */
  clearParagraphMarkDeletion(): this {
    delete this.formatting.paragraphMarkDeletion;
    return this;
  }

  /**
   * Checks if the paragraph mark is marked as deleted
   * @returns True if the paragraph mark is deleted
   */
  isParagraphMarkDeleted(): boolean {
    return !!this.formatting.paragraphMarkDeletion;
  }

  /**
   * Marks the paragraph mark (¶ symbol) as inserted via tracked changes.
   * This adds a `w:ins` element inside `w:pPr/w:rPr` indicating the insertion
   * of the ¶ symbol.
   *
   * @param id - Unique revision ID
   * @param author - Author who inserted the paragraph mark
   * @param date - Date when the insertion occurred (defaults to now)
   * @returns This paragraph for chaining
   */
  markParagraphMarkAsInserted(id: number, author: string, date?: Date): this {
    this.formatting.paragraphMarkInsertion = {
      id,
      author,
      date: date || new Date(),
    };
    return this;
  }

  /**
   * Clears the paragraph mark insertion marker
   * @returns This paragraph for chaining
   */
  clearParagraphMarkInsertion(): this {
    delete this.formatting.paragraphMarkInsertion;
    return this;
  }

  /**
   * Checks if the paragraph mark is marked as inserted
   * @returns True if the paragraph mark is inserted
   */
  isParagraphMarkInserted(): boolean {
    return !!this.formatting.paragraphMarkInsertion;
  }

  /**
   * Converts the paragraph to WordprocessingML XML element
   *
   * **ECMA-376 Compliance:** Properties are generated in the order specified by
   * ECMA-376 Part 1 §17.3.1.26 to ensure strict OpenXML conformance.
   *
   * Per spec, the order includes (partial list):
   * 1. pStyle (style reference)
   * 2. keepNext (keep with next paragraph)
   * 3. keepLines (keep lines together)
   * 4. pageBreakBefore (page break before)
   * 5. widowControl (widow/orphan control)
   * 6. numPr (numbering properties)
   * 7. suppressLineNumbers (suppress line numbers)
   * 8-10. borders, shading, tabs
   * 11-19. East Asian typography properties
   * 20. bidi (bidirectional layout)
   * 21. adjustRightInd (auto-adjust right indent)
   * 22. spacing, indentation, contextualSpacing
   * 23. mirrorIndents (mirror indents)
   * 24. jc (justification/alignment)
   * 25. textAlignment (vertical text alignment)
   * 26. textDirection (text flow direction)
   * 27. outlineLvl (outline level)
   *
   * @returns XMLElement representing the paragraph
   */
  toXML(): XMLElement {
    // Diagnostic logging before serialization
    const runData = this.getRuns().map((run) => ({
      text: run.getText(),
      rtl: run.isRTL(),
    }));
    logParagraphContent('serialization', -1, runData, this.formatting.bidi);

    if (this.formatting.bidi) {
      logTextDirection(`Serializing paragraph with BiDi enabled`);
    }

    const pPrChildren: XMLElement[] = [];

    // 1. Paragraph style (must be first per ECMA-376 §17.3.1.26)
    if (this.formatting.style) {
      pPrChildren.push(XMLBuilder.wSelf('pStyle', { 'w:val': this.formatting.style }));
    }

    // 2. Keep with next paragraph (CT_OnOff — emit val="0" to override style inheritance)
    if (this.formatting.keepNext !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('keepNext', { 'w:val': this.formatting.keepNext ? '1' : '0' })
      );
    }

    // 3. Keep lines together
    if (this.formatting.keepLines !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('keepLines', { 'w:val': this.formatting.keepLines ? '1' : '0' })
      );
    }

    // CT_PPrBase element order per ECMA-376:
    // pStyle → keepNext → keepLines → pageBreakBefore → framePr → widowControl →
    // numPr → suppressLineNumbers → pBdr → shd → tabs → suppressAutoHyphens →
    // kinsoku → wordWrap → overflowPunct → topLinePunct → autoSpaceDE → autoSpaceDN →
    // bidi → adjustRightInd → spacing → ind → contextualSpacing → mirrorIndents →
    // suppressOverlap → jc → textDirection → textAlignment → textboxTightWrap →
    // outlineLvl → divId → cnfStyle

    // 4. Page break before
    if (this.formatting.pageBreakBefore !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('pageBreakBefore', {
          'w:val': this.formatting.pageBreakBefore ? '1' : '0',
        })
      );
    }

    // 5. Text frame properties (framePr)
    if (this.formatting.framePr) {
      const attrs: Record<string, string> = {};
      const f = this.formatting.framePr;
      if (f.w !== undefined) attrs['w:w'] = f.w.toString();
      if (f.h !== undefined) attrs['w:h'] = f.h.toString();
      if (f.hRule) attrs['w:hRule'] = f.hRule;
      if (f.x !== undefined) attrs['w:x'] = f.x.toString();
      if (f.y !== undefined) attrs['w:y'] = f.y.toString();
      if (f.xAlign) attrs['w:xAlign'] = f.xAlign;
      if (f.yAlign) attrs['w:yAlign'] = f.yAlign;
      if (f.hAnchor) attrs['w:hAnchor'] = f.hAnchor;
      if (f.vAnchor) attrs['w:vAnchor'] = f.vAnchor;
      if (f.hSpace !== undefined) attrs['w:hSpace'] = f.hSpace.toString();
      if (f.vSpace !== undefined) attrs['w:vSpace'] = f.vSpace.toString();
      if (f.wrap) attrs['w:wrap'] = f.wrap;
      if (f.dropCap) attrs['w:dropCap'] = f.dropCap;
      if (f.lines !== undefined) attrs['w:lines'] = f.lines.toString();
      if (f.anchorLock !== undefined) attrs['w:anchorLock'] = f.anchorLock ? '1' : '0';
      if (Object.keys(attrs).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('framePr', attrs));
      }
    }

    // 6. Widow/orphan control
    if (this.formatting.widowControl !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('widowControl', {
          'w:val': this.formatting.widowControl ? '1' : '0',
        })
      );
    }

    // 7. Numbering properties
    if (this.formatting.numbering) {
      const numPr = XMLBuilder.w('numPr', undefined, [
        XMLBuilder.wSelf('ilvl', {
          'w:val': this.formatting.numbering.level.toString(),
        }),
        XMLBuilder.wSelf('numId', {
          'w:val': this.formatting.numbering.numId.toString(),
        }),
      ]);
      pPrChildren.push(numPr);
    } else if (this.formatting.numberingSuppressed) {
      // Per ECMA-376 §17.3.1.19, numId=0 explicitly removes numbering inherited from style
      const numPr = XMLBuilder.w('numPr', undefined, [
        XMLBuilder.wSelf('ilvl', { 'w:val': '0' }),
        XMLBuilder.wSelf('numId', { 'w:val': '0' }),
      ]);
      pPrChildren.push(numPr);
    }

    // 8. Suppress line numbers
    if (this.formatting.suppressLineNumbers !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('suppressLineNumbers', {
          'w:val': this.formatting.suppressLineNumbers ? '1' : '0',
        })
      );
    }

    // 9. Paragraph borders
    if (this.formatting.borders) {
      const borderChildren: XMLElement[] = [];
      const borders = this.formatting.borders;

      const createBorder = (
        borderType: string,
        border: BorderDefinition | undefined
      ): XMLElement | null => {
        if (!border) return null;
        // CT_Border §17.18.2: `w:val` (ST_Border) is REQUIRED. Default to
        // "nil" (the "no visible border" sentinel) when the consumer set
        // only size/color/space — otherwise strict OOXML validators reject
        // the document with "The required attribute 'val' is missing".
        const attributes: Record<string, string | number> = {
          'w:val': border.style ?? 'nil',
        };
        if (border.size !== undefined) attributes['w:sz'] = border.size;
        if (border.color) attributes['w:color'] = border.color;
        if (border.space !== undefined) attributes['w:space'] = border.space;
        // Full CT_Border attribute set (§17.18.2): themeColor / themeTint /
        // themeShade / shadow / frame. Round-trips themed paragraph borders
        // authored by Word — previously all five were silently stripped.
        if (border.themeColor) attributes['w:themeColor'] = border.themeColor;
        if (border.themeTint) attributes['w:themeTint'] = border.themeTint;
        if (border.themeShade) attributes['w:themeShade'] = border.themeShade;
        if (border.shadow !== undefined) attributes['w:shadow'] = border.shadow ? '1' : '0';
        if (border.frame !== undefined) attributes['w:frame'] = border.frame ? '1' : '0';
        return XMLBuilder.wSelf(borderType, attributes);
      };

      // Add borders in order: top, left, bottom, right, between, bar
      const topBorder = createBorder('top', borders.top);
      if (topBorder) borderChildren.push(topBorder);

      const leftBorder = createBorder('left', borders.left);
      if (leftBorder) borderChildren.push(leftBorder);

      const bottomBorder = createBorder('bottom', borders.bottom);
      if (bottomBorder) borderChildren.push(bottomBorder);

      const rightBorder = createBorder('right', borders.right);
      if (rightBorder) borderChildren.push(rightBorder);

      const betweenBorder = createBorder('between', borders.between);
      if (betweenBorder) borderChildren.push(betweenBorder);

      const barBorder = createBorder('bar', borders.bar);
      if (barBorder) borderChildren.push(barBorder);

      if (borderChildren.length > 0) {
        pPrChildren.push(XMLBuilder.w('pBdr', undefined, borderChildren));
      }
    }

    // 10. Paragraph shading
    if (this.formatting.shading) {
      const shdAttrs = buildShadingAttributes(this.formatting.shading);
      if (Object.keys(shdAttrs).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('shd', shdAttrs));
      }
    }

    // 11. Tab stops. CT_TabStop §17.3.1.37 declares BOTH `w:val` (ST_TabJc)
    // and `w:pos` (ST_SignedTwipsMeasure) as REQUIRED attributes. Default
    // `w:val` to "left" (ST_TabJc default and Word's authored default) when
    // caller didn't specify — otherwise strict OOXML validation rejects
    // the output with "The required attribute 'val' is missing".
    if (this.formatting.tabs && this.formatting.tabs.length > 0) {
      const tabChildren: XMLElement[] = [];
      for (const tab of this.formatting.tabs) {
        const attributes: Record<string, string | number> = {
          'w:val': tab.val ?? 'left',
          'w:pos': tab.position,
        };
        if (tab.leader) attributes['w:leader'] = tab.leader;
        tabChildren.push(XMLBuilder.wSelf('tab', attributes));
      }
      if (tabChildren.length > 0) {
        pPrChildren.push(XMLBuilder.w('tabs', undefined, tabChildren));
      }
    }

    // 12. Suppress automatic hyphenation
    if (this.formatting.suppressAutoHyphens !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('suppressAutoHyphens', {
          'w:val': this.formatting.suppressAutoHyphens ? '1' : '0',
        })
      );
    }

    // 13. CJK paragraph properties
    if (this.formatting.kinsoku !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('kinsoku', { 'w:val': this.formatting.kinsoku ? '1' : '0' })
      );
    }
    if (this.formatting.wordWrap !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('wordWrap', { 'w:val': this.formatting.wordWrap ? '1' : '0' })
      );
    }
    if (this.formatting.overflowPunct !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('overflowPunct', { 'w:val': this.formatting.overflowPunct ? '1' : '0' })
      );
    }
    if (this.formatting.topLinePunct !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('topLinePunct', { 'w:val': this.formatting.topLinePunct ? '1' : '0' })
      );
    }
    if (this.formatting.autoSpaceDE !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('autoSpaceDE', { 'w:val': this.formatting.autoSpaceDE ? '1' : '0' })
      );
    }
    if (this.formatting.autoSpaceDN !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('autoSpaceDN', { 'w:val': this.formatting.autoSpaceDN ? '1' : '0' })
      );
    }

    // 14. Bidirectional layout
    if (this.formatting.bidi !== undefined) {
      pPrChildren.push(XMLBuilder.wSelf('bidi', { 'w:val': this.formatting.bidi ? '1' : '0' }));
    }

    // 15. Auto-adjust right indent
    if (this.formatting.adjustRightInd !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('adjustRightInd', {
          'w:val': this.formatting.adjustRightInd ? '1' : '0',
        })
      );
    }

    // 16. Spacing (before/after/line + autospacing/lines per ECMA-376 §17.3.1.33)
    if (this.formatting.spacing) {
      const spc = this.formatting.spacing;
      const attributes: Record<string, number | string> = {};
      if (spc.before !== undefined) attributes['w:before'] = spc.before;
      if (spc.beforeLines !== undefined) attributes['w:beforeLines'] = spc.beforeLines;
      if (spc.beforeAutospacing !== undefined)
        attributes['w:beforeAutospacing'] = spc.beforeAutospacing ? '1' : '0';
      if (spc.after !== undefined) attributes['w:after'] = spc.after;
      if (spc.afterLines !== undefined) attributes['w:afterLines'] = spc.afterLines;
      if (spc.afterAutospacing !== undefined)
        attributes['w:afterAutospacing'] = spc.afterAutospacing ? '1' : '0';
      if (spc.line !== undefined) attributes['w:line'] = spc.line;
      if (spc.lineRule) attributes['w:lineRule'] = spc.lineRule;
      if (Object.keys(attributes).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('spacing', attributes));
      }
    }

    // 17. Indentation (left/right/firstLine/hanging + CJK character-unit variants)
    // Per ECMA-376 §17.3.1.12, firstLine and hanging are mutually exclusive —
    // hanging takes precedence. The *Chars attributes (ST_DecimalNumber, in
    // hundredths of a character unit) are independent of the twips variants;
    // Word authors in CJK locales emit them alongside — preserve both on save.
    if (this.formatting.indentation) {
      const ind = this.formatting.indentation;
      const attributes: Record<string, number> = {};
      if (ind.left !== undefined) attributes['w:left'] = ind.left;
      if (ind.right !== undefined) attributes['w:right'] = ind.right;
      if (ind.hanging !== undefined) {
        attributes['w:hanging'] = ind.hanging;
      } else if (ind.firstLine !== undefined) {
        attributes['w:firstLine'] = ind.firstLine;
      }
      if (ind.leftChars !== undefined) attributes['w:leftChars'] = ind.leftChars;
      if (ind.rightChars !== undefined) attributes['w:rightChars'] = ind.rightChars;
      // Same mutual-exclusive treatment: hangingChars wins over firstLineChars
      // when both are set, mirroring the twips behaviour above.
      if (ind.hangingChars !== undefined) {
        attributes['w:hangingChars'] = ind.hangingChars;
      } else if (ind.firstLineChars !== undefined) {
        attributes['w:firstLineChars'] = ind.firstLineChars;
      }
      if (Object.keys(attributes).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('ind', attributes));
      }
    }

    // 18. Contextual spacing
    if (this.formatting.contextualSpacing !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('contextualSpacing', {
          'w:val': this.formatting.contextualSpacing ? '1' : '0',
        })
      );
    }

    // 19. Mirror indents
    if (this.formatting.mirrorIndents !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('mirrorIndents', { 'w:val': this.formatting.mirrorIndents ? '1' : '0' })
      );
    }

    // 20. Suppress text frame overlap
    if (this.formatting.suppressOverlap !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('suppressOverlap', {
          'w:val': this.formatting.suppressOverlap ? '1' : '0',
        })
      );
    }

    // 21. Justification/Alignment
    if (this.formatting.alignment) {
      const alignmentValue =
        this.formatting.alignment === 'justify' ? 'both' : this.formatting.alignment;
      pPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': alignmentValue }));
    }

    // 22. Text direction
    if (this.formatting.textDirection) {
      pPrChildren.push(
        XMLBuilder.wSelf('textDirection', {
          'w:val': this.formatting.textDirection,
        })
      );
    }

    // 23. Text vertical alignment
    if (this.formatting.textAlignment) {
      pPrChildren.push(
        XMLBuilder.wSelf('textAlignment', {
          'w:val': this.formatting.textAlignment,
        })
      );
    }

    // 24. Textbox tight wrap
    if (this.formatting.textboxTightWrap) {
      pPrChildren.push(
        XMLBuilder.wSelf('textboxTightWrap', {
          'w:val': this.formatting.textboxTightWrap,
        })
      );
    }

    // 25. Outline level
    if (this.formatting.outlineLevel !== undefined) {
      pPrChildren.push(
        XMLBuilder.wSelf('outlineLvl', {
          'w:val': this.formatting.outlineLevel.toString(),
        })
      );
    }

    // 17. HTML div ID per ECMA-376 Part 1 §17.3.1.9
    if (this.formatting.divId !== undefined) {
      pPrChildren.push(XMLBuilder.wSelf('divId', { 'w:val': this.formatting.divId.toString() }));
    }

    // 18. Conditional table style formatting per ECMA-376 Part 1 §17.3.1.8
    if (this.formatting.cnfStyle) {
      pPrChildren.push(XMLBuilder.wSelf('cnfStyle', { 'w:val': this.formatting.cnfStyle }));
    }

    // 19. Paragraph mark run properties per ECMA-376 Part 1 §17.3.1.29
    // Per CT_PPr, w:rPr comes after all CT_PPrBase elements and before w:sectPr/w:pPrChange.
    //
    // CT_ParaRPr content model (ECMA-376 Part 1, Annex A / wml.xsd):
    //   <xsd:sequence>
    //     <xsd:group ref="EG_ParaRPrTrackChanges" minOccurs="0"/>   <!-- ins, del, moveFrom, moveTo -->
    //     <xsd:group ref="EG_RPrBase" minOccurs="0" maxOccurs="unbounded"/>  <!-- rStyle, rFonts, b, bCs, i, ... -->
    //     <xsd:element name="rPrChange" type="CT_ParaRPrChange" minOccurs="0"/>
    //   </xsd:sequence>
    //
    // The track-change markers (w:ins / w:del) must precede the EG_RPrBase
    // run-property children. Earlier revisions of this code emitted run
    // properties first and then ins/del, which is a schema violation — strict
    // validators reject the inverted order, and tracked-change-aware tools
    // may misinterpret the revision state of the paragraph mark.
    if (
      this.formatting.paragraphMarkRunProperties ||
      this.formatting.paragraphMarkDeletion ||
      this.formatting.paragraphMarkInsertion ||
      this.formatting.paragraphMarkRunPropertiesChange
    ) {
      const rPrChildren: XMLElement[] = [];

      // EG_ParaRPrTrackChanges (ins → del → moveFrom → moveTo) — FIRST per CT_ParaRPr.
      if (this.formatting.paragraphMarkInsertion) {
        const ins = this.formatting.paragraphMarkInsertion;
        rPrChildren.push(
          XMLBuilder.wSelf('ins', {
            'w:id': ins.id.toString(),
            'w:author': ins.author,
            'w:date': formatDateForXml(ins.date),
          })
        );
      }

      if (this.formatting.paragraphMarkDeletion) {
        const del = this.formatting.paragraphMarkDeletion;
        rPrChildren.push(
          XMLBuilder.wSelf('del', {
            'w:id': del.id.toString(),
            'w:author': del.author,
            'w:date': formatDateForXml(del.date),
          })
        );
      }

      // EG_RPrBase run properties — AFTER the track-change markers.
      if (this.formatting.paragraphMarkRunProperties) {
        const rPr = Run.generateRunPropertiesXML(this.formatting.paragraphMarkRunProperties);
        if (rPr?.children) {
          // Filter to only XMLElement types (children can be XMLElement or string)
          for (const child of rPr.children) {
            if (typeof child !== 'string') {
              rPrChildren.push(child);
            }
          }
        }
      }

      // <w:rPrChange> (CT_ParaRPrChange, §17.3.1.30) — LAST per CT_ParaRPr.
      // Contains a single <w:rPr> child with the previous run properties of
      // the paragraph mark. Reuses Run.generateRunPropertiesXML for full
      // CT_RPr coverage in the previous-properties block.
      const paraRPrChange = this.formatting.paragraphMarkRunPropertiesChange;
      if (paraRPrChange) {
        const prevRPr = Run.generateRunPropertiesXML(
          paraRPrChange.previousProperties as RunFormatting
        );
        const prevRPrElement: XMLElement = prevRPr ?? {
          name: 'w:rPr',
          attributes: {},
          children: [],
        };
        rPrChildren.push({
          name: 'w:rPrChange',
          attributes: {
            'w:id': paraRPrChange.id.toString(),
            'w:author': paraRPrChange.author,
            'w:date': formatDateForXml(paraRPrChange.date),
          },
          children: [prevRPrElement],
        });
      }

      // Add w:rPr element if there are any run properties
      if (rPrChildren.length > 0) {
        pPrChildren.push(XMLBuilder.w('rPr', undefined, rPrChildren));
      }
    }

    // 20. Paragraph property change tracking per ECMA-376 Part 1 §17.3.1.27
    /**
     * Per OOXML spec, w:pPrChange contains:
     * - Attributes: w:id (required), w:author (required), w:date (optional)
     * - Child element: w:pPr containing the PREVIOUS paragraph properties before the change
     *
     * Structure:
     * <w:pPrChange w:id="1" w:author="Author" w:date="2024-01-01T12:00:00Z">
     *   <w:pPr>
     *     <!-- Previous paragraph properties -->
     *   </w:pPr>
     * </w:pPrChange>
     */
    // Per CT_PPr: sectPr comes BEFORE pPrChange
    if (this.formatting.sectPr) {
      if (typeof this.formatting.sectPr === 'string') {
        // Raw XML passthrough for inline sectPr (preserves exact structure)
        pPrChildren.push({
          name: '__rawXml',
          rawXml: this.formatting.sectPr,
        } as XMLElement);
      }
      // Non-string (parsed object) is skipped to prevent corruption from
      // XMLBuilder.wSelf treating complex objects as flat attributes
    }

    if (this.formatting.pPrChange) {
      const change = this.formatting.pPrChange;
      const attrs: Record<string, string> = {};
      // ECMA-376 attribute order: w:id (required), w:author (required), w:date (optional)
      if (change.id) attrs['w:id'] = change.id;
      if (change.author) attrs['w:author'] = change.author;
      if (change.date) attrs['w:date'] = change.date;

      // Build child w:pPr element with previous properties
      // Per CT_PPrBase schema order: pStyle, keepNext, keepLines, pageBreakBefore,
      // framePr, widowControl, numPr, suppressLineNumbers, pBdr, shd, tabs,
      // suppressAutoHyphens, kinsoku, wordWrap, overflowPunct, topLinePunct,
      // autoSpaceDE, autoSpaceDN, bidi, adjustRightInd, spacing, ind,
      // contextualSpacing, mirrorIndents, suppressOverlap, jc, textDirection,
      // textAlignment, outlineLvl
      const prevPPrChildren: XMLElement[] = [];
      if (change.previousProperties) {
        const prev = change.previousProperties;

        // 1. pStyle
        if (prev.style) {
          prevPPrChildren.push(XMLBuilder.wSelf('pStyle', { 'w:val': prev.style }));
        }

        // 2. keepNext
        if (prev.keepNext !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('keepNext', prev.keepNext ? { 'w:val': '1' } : { 'w:val': '0' })
          );
        }

        // 3. keepLines
        if (prev.keepLines !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('keepLines', prev.keepLines ? { 'w:val': '1' } : { 'w:val': '0' })
          );
        }

        // 4. pageBreakBefore
        if (prev.pageBreakBefore !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf(
              'pageBreakBefore',
              prev.pageBreakBefore ? { 'w:val': '1' } : { 'w:val': '0' }
            )
          );
        }

        // 5. framePr (text frame properties)
        if (prev.framePr) {
          const fAttrs: Record<string, string> = {};
          const f = prev.framePr;
          if (f.w !== undefined) fAttrs['w:w'] = f.w.toString();
          if (f.h !== undefined) fAttrs['w:h'] = f.h.toString();
          if (f.hRule) fAttrs['w:hRule'] = f.hRule;
          if (f.x !== undefined) fAttrs['w:x'] = f.x.toString();
          if (f.y !== undefined) fAttrs['w:y'] = f.y.toString();
          if (f.xAlign) fAttrs['w:xAlign'] = f.xAlign;
          if (f.yAlign) fAttrs['w:yAlign'] = f.yAlign;
          if (f.hAnchor) fAttrs['w:hAnchor'] = f.hAnchor;
          if (f.vAnchor) fAttrs['w:vAnchor'] = f.vAnchor;
          if (f.hSpace !== undefined) fAttrs['w:hSpace'] = f.hSpace.toString();
          if (f.vSpace !== undefined) fAttrs['w:vSpace'] = f.vSpace.toString();
          if (f.wrap) fAttrs['w:wrap'] = f.wrap;
          if (f.dropCap) fAttrs['w:dropCap'] = f.dropCap;
          if (f.lines !== undefined) fAttrs['w:lines'] = f.lines.toString();
          if (f.anchorLock !== undefined) fAttrs['w:anchorLock'] = f.anchorLock ? '1' : '0';
          if (Object.keys(fAttrs).length > 0) {
            prevPPrChildren.push(XMLBuilder.wSelf('framePr', fAttrs));
          }
        }

        // 6. widowControl
        if (prev.widowControl !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('widowControl', {
              'w:val': prev.widowControl ? '1' : '0',
            })
          );
        }

        // 6. numPr
        if (prev.numbering) {
          const numPrChildren: XMLElement[] = [];
          if (prev.numbering.level !== undefined) {
            numPrChildren.push(
              XMLBuilder.wSelf('ilvl', { 'w:val': prev.numbering.level.toString() })
            );
          }
          if (prev.numbering.numId !== undefined) {
            numPrChildren.push(
              XMLBuilder.wSelf('numId', { 'w:val': prev.numbering.numId.toString() })
            );
          }
          if (numPrChildren.length > 0) {
            prevPPrChildren.push(XMLBuilder.w('numPr', undefined, numPrChildren));
          }
        } else if (prev.numberingSuppressed) {
          prevPPrChildren.push(
            XMLBuilder.w('numPr', undefined, [
              XMLBuilder.wSelf('ilvl', { 'w:val': '0' }),
              XMLBuilder.wSelf('numId', { 'w:val': '0' }),
            ])
          );
        }

        // 7. suppressLineNumbers — CT_OnOff. Emit w:val="0" for explicit false so the
        // tracked "previous" value round-trips losslessly (see main generator line ~3154).
        if (prev.suppressLineNumbers !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('suppressLineNumbers', {
              'w:val': prev.suppressLineNumbers ? '1' : '0',
            })
          );
        }

        // 8. pBdr (paragraph borders) — CT_Border §17.18.2 requires w:val.
        // Default to "nil" when style undefined so pPrChange's previous
        // border set round-trips through strict validation. Full CT_Border
        // attribute set (themeColor/themeTint/themeShade/shadow/frame) is
        // preserved so Word's themed borders survive tracked-change history.
        if (prev.borders) {
          const borderChildren: XMLElement[] = [];
          const borderSides = ['top', 'left', 'bottom', 'right', 'between', 'bar'] as const;
          for (const side of borderSides) {
            const border = prev.borders[side];
            if (border) {
              const attrs: Record<string, string> = { 'w:val': border.style ?? 'nil' };
              if (border.size !== undefined) attrs['w:sz'] = border.size.toString();
              if (border.color) attrs['w:color'] = border.color;
              if (border.space !== undefined) attrs['w:space'] = border.space.toString();
              if (border.themeColor) attrs['w:themeColor'] = border.themeColor;
              if (border.themeTint) attrs['w:themeTint'] = border.themeTint;
              if (border.themeShade) attrs['w:themeShade'] = border.themeShade;
              if (border.shadow !== undefined) attrs['w:shadow'] = border.shadow ? '1' : '0';
              if (border.frame !== undefined) attrs['w:frame'] = border.frame ? '1' : '0';
              borderChildren.push(XMLBuilder.wSelf(side, attrs));
            }
          }
          if (borderChildren.length > 0) {
            prevPPrChildren.push(XMLBuilder.w('pBdr', undefined, borderChildren));
          }
        }

        // 9. shd (paragraph shading)
        if (prev.shading) {
          const shdAttrs = buildShadingAttributes(prev.shading);
          if (Object.keys(shdAttrs).length > 0) {
            prevPPrChildren.push(XMLBuilder.wSelf('shd', shdAttrs));
          }
        }

        // 10. tabs — CT_TabStop §17.3.1.37 requires w:val AND w:pos. Default
        // w:val to "left" when pPrChange's previous state didn't record one.
        if (prev.tabs && prev.tabs.length > 0) {
          const tabChildren: XMLElement[] = prev.tabs.map((tab) => {
            const tabAttrs: Record<string, string> = {
              'w:val': tab.val ?? 'left',
              'w:pos': tab.position.toString(),
            };
            if (tab.leader) tabAttrs['w:leader'] = tab.leader;
            return XMLBuilder.wSelf('tab', tabAttrs);
          });
          prevPPrChildren.push(XMLBuilder.w('tabs', undefined, tabChildren));
        }

        // 11. suppressAutoHyphens — CT_OnOff. Emit w:val="0" for explicit false so the
        // tracked "previous" value round-trips losslessly (see main generator line ~3231).
        if (prev.suppressAutoHyphens !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('suppressAutoHyphens', {
              'w:val': prev.suppressAutoHyphens ? '1' : '0',
            })
          );
        }

        // 12. CJK paragraph properties
        if (prev.kinsoku !== undefined) {
          prevPPrChildren.push(XMLBuilder.wSelf('kinsoku', { 'w:val': prev.kinsoku ? '1' : '0' }));
        }
        if (prev.wordWrap !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('wordWrap', { 'w:val': prev.wordWrap ? '1' : '0' })
          );
        }
        if (prev.overflowPunct !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('overflowPunct', { 'w:val': prev.overflowPunct ? '1' : '0' })
          );
        }
        if (prev.topLinePunct !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('topLinePunct', { 'w:val': prev.topLinePunct ? '1' : '0' })
          );
        }
        if (prev.autoSpaceDE !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('autoSpaceDE', { 'w:val': prev.autoSpaceDE ? '1' : '0' })
          );
        }
        if (prev.autoSpaceDN !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('autoSpaceDN', { 'w:val': prev.autoSpaceDN ? '1' : '0' })
          );
        }

        // 13. bidi
        if (prev.bidi !== undefined) {
          prevPPrChildren.push(XMLBuilder.wSelf('bidi', { 'w:val': prev.bidi ? '1' : '0' }));
        }

        // 13. adjustRightInd
        if (prev.adjustRightInd !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('adjustRightInd', {
              'w:val': prev.adjustRightInd ? '1' : '0',
            })
          );
        }

        // 14. spacing (all 8 CT_Spacing attributes)
        if (prev.spacing) {
          const spacingAttrs: Record<string, string> = {};
          if (prev.spacing.before !== undefined)
            spacingAttrs['w:before'] = prev.spacing.before.toString();
          if (prev.spacing.beforeLines !== undefined)
            spacingAttrs['w:beforeLines'] = prev.spacing.beforeLines.toString();
          if (prev.spacing.beforeAutospacing !== undefined)
            spacingAttrs['w:beforeAutospacing'] = prev.spacing.beforeAutospacing ? '1' : '0';
          if (prev.spacing.after !== undefined)
            spacingAttrs['w:after'] = prev.spacing.after.toString();
          if (prev.spacing.afterLines !== undefined)
            spacingAttrs['w:afterLines'] = prev.spacing.afterLines.toString();
          if (prev.spacing.afterAutospacing !== undefined)
            spacingAttrs['w:afterAutospacing'] = prev.spacing.afterAutospacing ? '1' : '0';
          if (prev.spacing.line !== undefined)
            spacingAttrs['w:line'] = prev.spacing.line.toString();
          if (prev.spacing.lineRule) spacingAttrs['w:lineRule'] = prev.spacing.lineRule;
          if (Object.keys(spacingAttrs).length > 0) {
            prevPPrChildren.push(XMLBuilder.wSelf('spacing', spacingAttrs));
          }
        }

        // 15. ind (indentation + CJK character-unit variants per ECMA-376 §17.3.1.12)
        // hanging / firstLine are conceptually mutually exclusive (opposite-
        // direction first-line indents). Mirror the direct paragraph generator's
        // precedence: hanging wins when both are set. Previously this path emitted
        // BOTH attributes, producing ambiguous tracked-change history that Word
        // renders inconsistently across versions.
        if (prev.indentation) {
          const indAttrs: Record<string, string> = {};
          const indentation = prev.indentation as Record<string, number | undefined>;
          if (indentation.left !== undefined) indAttrs['w:left'] = indentation.left.toString();
          if (indentation.right !== undefined) indAttrs['w:right'] = indentation.right.toString();
          if (indentation.hanging !== undefined) {
            indAttrs['w:hanging'] = indentation.hanging.toString();
          } else if (indentation.firstLine !== undefined) {
            indAttrs['w:firstLine'] = indentation.firstLine.toString();
          }
          if (indentation.leftChars !== undefined)
            indAttrs['w:leftChars'] = indentation.leftChars.toString();
          if (indentation.rightChars !== undefined)
            indAttrs['w:rightChars'] = indentation.rightChars.toString();
          if (indentation.hangingChars !== undefined) {
            indAttrs['w:hangingChars'] = indentation.hangingChars.toString();
          } else if (indentation.firstLineChars !== undefined) {
            indAttrs['w:firstLineChars'] = indentation.firstLineChars.toString();
          }
          prevPPrChildren.push(XMLBuilder.wSelf('ind', indAttrs));
        }

        // 16. contextualSpacing
        if (prev.contextualSpacing !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('contextualSpacing', {
              'w:val': prev.contextualSpacing ? '1' : '0',
            })
          );
        }

        // 17. mirrorIndents
        if (prev.mirrorIndents !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('mirrorIndents', {
              'w:val': prev.mirrorIndents ? '1' : '0',
            })
          );
        }

        // 18. suppressOverlap
        if (prev.suppressOverlap !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('suppressOverlap', {
              'w:val': prev.suppressOverlap ? '1' : '0',
            })
          );
        }

        // 19. jc (alignment)
        if (prev.alignment) {
          // Map 'justify' to 'both' per ECMA-376 ST_Jc enumeration
          const alignmentValue = prev.alignment === 'justify' ? 'both' : prev.alignment;
          prevPPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': alignmentValue }));
        }

        // 19. textDirection
        if (prev.textDirection) {
          prevPPrChildren.push(XMLBuilder.wSelf('textDirection', { 'w:val': prev.textDirection }));
        }

        // 20. textAlignment
        if (prev.textAlignment) {
          prevPPrChildren.push(XMLBuilder.wSelf('textAlignment', { 'w:val': prev.textAlignment }));
        }

        // 21. textboxTightWrap
        if (prev.textboxTightWrap) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('textboxTightWrap', { 'w:val': prev.textboxTightWrap })
          );
        }

        // 22. outlineLvl
        if (prev.outlineLevel !== undefined) {
          prevPPrChildren.push(
            XMLBuilder.wSelf('outlineLvl', { 'w:val': prev.outlineLevel.toString() })
          );
        }

        // 23. divId
        if (prev.divId !== undefined) {
          prevPPrChildren.push(XMLBuilder.wSelf('divId', { 'w:val': prev.divId.toString() }));
        }

        // 24. cnfStyle
        if (prev.cnfStyle) {
          prevPPrChildren.push(XMLBuilder.wSelf('cnfStyle', { 'w:val': prev.cnfStyle }));
        }
      }

      // Create w:pPrChange element with child w:pPr
      // Per ECMA-376 Part 1 §17.13.5.29, w:pPrChange MUST contain a w:pPr child element.
      // Only output pPrChange if we have properties to include in w:pPr.
      // Empty pPrChange elements cause Word to report "unreadable content" corruption.
      if (prevPPrChildren.length > 0) {
        const pPrChangeChildren: XMLElement[] = [
          {
            name: 'w:pPr',
            attributes: {},
            children: prevPPrChildren,
          },
        ];

        pPrChildren.push({
          name: 'w:pPrChange',
          attributes: attrs,
          children: pPrChangeChildren,
        });
      }
      // If no previous properties to record, skip w:pPrChange entirely to avoid corruption
    }

    // Build paragraph element
    const paragraphChildren: XMLElement[] = [];

    // Add paragraph properties if there are any
    if (pPrChildren.length > 0) {
      paragraphChildren.push(XMLBuilder.w('pPr', undefined, pPrChildren));
    }

    // Add bookmark start markers
    for (const bookmark of this.bookmarksStart) {
      paragraphChildren.push(bookmark.toStartXML());
    }

    // Add comment range start markers
    for (const comment of this.commentsStart) {
      paragraphChildren.push(comment.toRangeStartXML());
    }

    // Add content (runs, fields, hyperlinks, revisions, range markers, shapes, text boxes)
    for (let i = 0; i < this.content.length; i++) {
      const item = this.content[i];
      if (item instanceof Field) {
        // Simple Field — fldSimple is a paragraph-level element containing w:r children
        paragraphChildren.push(item.toXML());
      } else if (item instanceof ComplexField) {
        // ComplexField returns array of runs - spread them directly into paragraphChildren
        const fieldXml = item.toXML();
        if (Array.isArray(fieldXml)) {
          paragraphChildren.push(...fieldXml);
        } else {
          // Fallback if toXML() doesn't return array
          paragraphChildren.push(XMLBuilder.w('r', undefined, [fieldXml]));
        }
      } else if (item instanceof Hyperlink) {
        // Hyperlinks are their own element
        paragraphChildren.push(item.toXML());
      } else if (item instanceof Revision) {
        // Revisions (track changes) are their own element
        // Property change types (rPrChange, pPrChange, tblPrChange, etc.) are only valid
        // inside their respective property elements (w:rPr, w:pPr, etc.), not as direct
        // children of w:p. Skip them here to avoid producing invalid XML.
        const revType = item.getType();
        if (
          revType === 'runPropertiesChange' ||
          revType === 'paragraphPropertiesChange' ||
          revType === 'tablePropertiesChange' ||
          revType === 'tableExceptionPropertiesChange' ||
          revType === 'tableRowPropertiesChange' ||
          revType === 'tableCellPropertiesChange' ||
          revType === 'sectionPropertiesChange' ||
          revType === 'numberingChange'
        ) {
          continue;
        }
        // Note: toXML() returns null for internal-only revision types (e.g., hyperlinkChange)
        const revisionXml = item.toXML();
        if (revisionXml) {
          paragraphChildren.push(revisionXml);
        }
      } else if (item instanceof RangeMarker) {
        // Range markers are their own element (mark boundaries of moves, etc.)
        paragraphChildren.push(item.toXML());
      } else if (item instanceof Shape) {
        // Shapes are wrapped in a run
        paragraphChildren.push(XMLBuilder.w('r', undefined, [item.toXML()]));
      } else if (item instanceof TextBox) {
        // Text boxes are wrapped in a run
        paragraphChildren.push(XMLBuilder.w('r', undefined, [item.toXML()]));
      } else if (item) {
        paragraphChildren.push(item.toXML());
      }
    }

    // If no content, add empty run to prevent invalid XML
    if (this.content.length === 0) {
      paragraphChildren.push(new Run('').toXML());
    }

    // Add comment range end markers
    for (const comment of this.commentsEnd) {
      paragraphChildren.push(comment.toRangeEndXML());
    }

    // Add comment references (must come after range end)
    for (const comment of this.commentsEnd) {
      paragraphChildren.push(comment.toReferenceXML());
    }

    // Add bookmark end markers
    for (const bookmark of this.bookmarksEnd) {
      paragraphChildren.push(bookmark.toEndXML());
    }

    // Add paragraph-level attributes (Word 2010+ w14:paraId and w14:textId)
    const paragraphAttributes: Record<string, string> = {};
    if (this.formatting.paraId) {
      paragraphAttributes['w14:paraId'] = this.formatting.paraId;
    }
    if (this.formatting.textId) {
      paragraphAttributes['w14:textId'] = this.formatting.textId;
    }

    return XMLBuilder.w(
      'p',
      Object.keys(paragraphAttributes).length > 0 ? paragraphAttributes : undefined,
      paragraphChildren
    );
  }

  /**
   * Gets the word count for this paragraph
   *
   * Counts words by splitting text on whitespace and filtering empty strings.
   *
   * @returns Number of words in the paragraph
   *
   * @example
   * ```typescript
   * const count = para.getWordCount();
   * console.log(`Paragraph has ${count} words`);
   * ```
   */
  getWordCount(): number {
    const text = this.getText().trim();
    if (!text) return 0;

    // Split by whitespace and filter out empty strings
    const words = text.split(/\s+/).filter((word) => word.length > 0);
    return words.length;
  }

  /**
   * Gets the character count for this paragraph
   *
   * Counts all characters including or excluding whitespace.
   *
   * @param includeSpaces - If true, includes spaces; if false, excludes them (default: true)
   * @returns Number of characters in the paragraph
   *
   * @example
   * ```typescript
   * const withSpaces = para.getLength();       // Includes spaces
   * const noSpaces = para.getLength(false);    // Excludes spaces
   * console.log(`${withSpaces} chars (${noSpaces} without spaces)`);
   * ```
   */
  getLength(includeSpaces = true): number {
    const text = this.getText();
    if (includeSpaces) {
      return text.length;
    } else {
      return text.replace(/\s/g, '').length;
    }
  }

  /**
   * Creates a deep clone of this paragraph
   *
   * Creates a new Paragraph with copies of all content, formatting,
   * bookmarks, and comments. The clone is independent of the original.
   *
   * @returns A new Paragraph instance with the same content and formatting
   *
   * @example
   * ```typescript
   * const original = new Paragraph();
   * original.addText('Template text', { bold: true });
   * original.setStyle('Heading1');
   *
   * const copy = original.clone();
   * copy.addText(' - modified'); // Original unchanged
   * ```
   */
  clone(): Paragraph {
    // Clone the formatting
    const clonedFormatting: ParagraphFormatting = deepClone(this.formatting);

    // Create new paragraph with cloned formatting
    const clonedParagraph = new Paragraph(clonedFormatting);

    // Clone all content (runs, fields, hyperlinks, revisions)
    for (const item of this.content) {
      if (item instanceof Run) {
        // Clone the run with its text and formatting
        const runFormatting = item.getFormatting();
        const clonedRun = new Run(item.getText(), deepClone(runFormatting));
        clonedParagraph.addRun(clonedRun);
      } else {
        // For other content types, add them as-is (shallow copy for now)
        // In a more complete implementation, we'd clone these too
        clonedParagraph.content.push(item);
      }
    }

    // Clone bookmark and comment markers
    clonedParagraph.bookmarksStart = [...this.bookmarksStart];
    clonedParagraph.bookmarksEnd = [...this.bookmarksEnd];
    clonedParagraph.commentsStart = [...this.commentsStart];
    clonedParagraph.commentsEnd = [...this.commentsEnd];

    return clonedParagraph;
  }

  /**
   * Serializes the paragraph to a portable JSON object
   *
   * Returns a plain object representing the paragraph's text content, inline
   * formatting, style, and paragraph-level formatting. Only Run content is
   * serialized — hyperlinks, revisions, and other non-Run elements are
   * represented by their text only.
   *
   * The output is designed for data pipelines, API responses, storage,
   * and debugging. Use `Paragraph.fromJSON()` to reconstruct.
   *
   * @returns JSON-serializable paragraph representation
   *
   * @example
   * ```typescript
   * const para = new Paragraph().addText('Hello ', { bold: true }).addText('World');
   * para.setStyle('Heading1').setAlignment('center');
   *
   * const json = para.toJSON();
   * // {
   * //   text: 'Hello World',
   * //   style: 'Heading1',
   * //   alignment: 'center',
   * //   runs: [
   * //     { text: 'Hello ', formatting: { bold: true } },
   * //     { text: 'World', formatting: {} }
   * //   ]
   * // }
   * ```
   */
  toJSON(): {
    text: string;
    style?: string;
    alignment?: string;
    runs: { text: string; formatting: RunFormatting }[];
    numbering?: { numId: number; level: number };
    indentation?: { left?: number; right?: number; firstLine?: number; hanging?: number };
    spacing?: { before?: number; after?: number; line?: number; lineRule?: string };
  } {
    const runs = this.getRuns().map((run) => ({
      text: run.getText(),
      formatting: run.getFormatting(),
    }));

    const result: ReturnType<Paragraph['toJSON']> = {
      text: this.getText(),
      runs,
    };

    if (this.formatting.style) result.style = this.formatting.style;
    if (this.formatting.alignment) result.alignment = this.formatting.alignment;
    if (this.formatting.numbering) result.numbering = { ...this.formatting.numbering };
    if (this.formatting.indentation) result.indentation = { ...this.formatting.indentation };
    if (this.formatting.spacing) {
      result.spacing = {
        before: this.formatting.spacing.before,
        after: this.formatting.spacing.after,
        line: this.formatting.spacing.line,
        lineRule: this.formatting.spacing.lineRule,
      };
    }

    return result;
  }

  /**
   * Creates a Paragraph from a JSON object (as produced by `toJSON()`)
   *
   * Reconstructs a paragraph with runs, formatting, style, and layout
   * from a serialized representation. This enables round-trip serialization.
   *
   * @param data - JSON object matching the `toJSON()` output shape
   * @returns New Paragraph instance
   *
   * @example
   * ```typescript
   * const json = { text: 'Hello', style: 'Heading1', runs: [{ text: 'Hello', formatting: { bold: true } }] };
   * const para = Paragraph.fromJSON(json);
   * para.getText(); // 'Hello'
   * para.getStyle(); // 'Heading1'
   * ```
   */
  static fromJSON(data: {
    text?: string;
    style?: string;
    alignment?: string;
    runs?: { text: string; formatting?: RunFormatting }[];
    numbering?: { numId: number; level: number };
    indentation?: { left?: number; right?: number; firstLine?: number; hanging?: number };
    spacing?: { before?: number; after?: number; line?: number; lineRule?: string };
  }): Paragraph {
    const para = new Paragraph();

    if (data.runs && data.runs.length > 0) {
      for (const runData of data.runs) {
        para.addText(runData.text, runData.formatting);
      }
    } else if (data.text) {
      para.addText(data.text);
    }

    if (data.style) para.setStyle(data.style);
    if (data.alignment) para.setAlignment(data.alignment as ParagraphAlignment);
    if (data.numbering) para.setNumbering(data.numbering.numId, data.numbering.level);
    if (data.indentation) {
      if (data.indentation.left !== undefined) para.setLeftIndent(data.indentation.left);
      if (data.indentation.right !== undefined) para.setRightIndent(data.indentation.right);
      if (data.indentation.firstLine !== undefined)
        para.setFirstLineIndent(data.indentation.firstLine);
      if (data.indentation.hanging !== undefined) para.setHangingIndent(data.indentation.hanging);
    }
    if (data.spacing) {
      if (data.spacing.before !== undefined) para.setSpaceBefore(data.spacing.before);
      if (data.spacing.after !== undefined) para.setSpaceAfter(data.spacing.after);
      if (data.spacing.line !== undefined) {
        para.setLineSpacing(
          data.spacing.line,
          (data.spacing.lineRule as 'auto' | 'exact' | 'atLeast') ?? 'auto'
        );
      }
    }

    return para;
  }

  /**
   * Splits this paragraph at a character offset, returning the tail as a new paragraph
   *
   * Content from the offset onward is moved to a new Paragraph. The new paragraph
   * inherits this paragraph's style and formatting (deep-cloned). If the split point
   * falls within a Run, that run is split using `Run.splitAt()`.
   *
   * Only Run elements contribute to the character offset count. Non-Run content
   * (hyperlinks, revisions, fields) that appears after all affected runs is moved
   * to the new paragraph as-is.
   *
   * @param offset - Character position to split at (0-based). Content from this
   *   position onward moves to the new paragraph.
   * @returns A new Paragraph containing content from `offset` onward
   *
   * @example
   * ```typescript
   * const para = new Paragraph().addText('Hello World');
   * const tail = para.splitAt(5);
   * para.getText();   // "Hello"
   * tail.getText();   // " World"
   * ```
   *
   * @example
   * ```typescript
   * // Insert a table between two halves of a paragraph
   * const tail = para.splitAt(offset);
   * const paraIndex = doc.getParagraphIndex(para);
   * doc.insertTableAt(paraIndex + 1, table);
   * doc.insertParagraphAt(paraIndex + 2, tail);
   * ```
   */
  splitAt(offset: number): Paragraph {
    // Create new paragraph with cloned formatting and style
    const tailPara = new Paragraph(deepClone(this.formatting));

    // Edge case: split at or past end — return empty paragraph
    const totalLength = this.getText().length;
    if (offset >= totalLength) {
      return tailPara;
    }

    // Edge case: split at or before start — move everything
    if (offset <= 0) {
      tailPara.content = this.content;
      this.content = [];
      return tailPara;
    }

    // Walk content to find the split point
    let charPos = 0;
    let splitContentIndex = -1;

    for (let i = 0; i < this.content.length; i++) {
      const item = this.content[i]!;

      if (item instanceof Run) {
        const runLen = item.getText().length;
        const runEnd = charPos + runLen;

        if (runEnd <= offset) {
          // Entire run stays in this paragraph
          charPos = runEnd;
          continue;
        }

        if (charPos >= offset) {
          // Entire run goes to tail paragraph
          splitContentIndex = i;
          break;
        }

        // Split falls within this run
        const localOffset = offset - charPos;
        const tailRun = item.splitAt(localOffset);

        // Insert the tail run and mark everything after it for the new paragraph
        this.content.splice(i + 1, 0, tailRun);
        splitContentIndex = i + 1;
        break;
      } else {
        // Non-Run content: if we haven't reached the offset yet, keep it here
        // Once we've passed the offset, it moves to tail
        if (charPos >= offset) {
          splitContentIndex = i;
          break;
        }
      }
    }

    // If no split point found (shouldn't happen given offset < totalLength),
    // but guard against it
    if (splitContentIndex === -1) {
      return tailPara;
    }

    // Move content from splitContentIndex onward to the tail paragraph
    tailPara.content = this.content.splice(splitContentIndex);

    return tailPara;
  }

  /**
   * Sets paragraph borders
   * @param borders - Border definitions for each side
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * para.setBorder({
   *   top: { style: 'single', size: 4, color: '000000', space: 1 },
   *   bottom: { style: 'single', size: 4, color: '000000', space: 1 }
   * });
   * ```
   */
  setBorder(borders: {
    top?: BorderDefinition;
    bottom?: BorderDefinition;
    left?: BorderDefinition;
    right?: BorderDefinition;
    between?: BorderDefinition;
    bar?: BorderDefinition;
  }): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    this.formatting.borders = borders;

    return this;
  }

  /**
   * Sets paragraph shading (background color and pattern)
   * @param shading - Shading options
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * // Solid background
   * para.setShading({ fill: 'FFFF00', pattern: 'solid' });
   *
   * // Pattern with colors
   * para.setShading({
   *   fill: 'FFFF00',
   *   color: '000000',
   *   pattern: 'diagStripe'
   * });
   * ```
   */
  setShading(shading: ShadingConfig): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    const previousValue = this.formatting.shading;
    this.formatting.shading = shading;

    if (this.trackingContext?.isEnabled() && previousValue !== shading) {
      this.trackingContext.trackParagraphPropertyChange(this, 'shading', previousValue, shading);
    }

    return this;
  }

  /**
   * Sets tab stops for the paragraph
   * @param tabs - Array of tab stop definitions
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * para.setTabs([
   *   { position: 720, val: 'left' },
   *   { position: 1440, val: 'center', leader: 'dot' },
   *   { position: 2160, val: 'right' }
   * ]);
   * ```
   */
  setTabs(tabs: TabStop[]): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    this.formatting.tabs = tabs;

    return this;
  }

  /**
   * Inserts a run at a specific position
   * @param index - Position to insert at (0-based)
   * @param run - Run to insert
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * const para = new Paragraph();
   * const run = new Run('Inserted', { bold: true });
   * para.insertRunAt(0, run);
   * ```
   */
  insertRunAt(index: number, run: Run): this {
    if (index < 0) index = 0;
    if (index > this.content.length) index = this.content.length;

    this.content.splice(index, 0, run);
    return this;
  }

  /**
   * Removes a run at a specific position
   * @param index - Position to remove (0-based)
   * @returns True if removed, false if index invalid
   * @example
   * ```typescript
   * para.removeRunAt(2);  // Remove third run
   * ```
   */
  removeRunAt(index: number): boolean {
    if (index >= 0 && index < this.content.length) {
      const item = this.content[index];
      if (item instanceof Run) {
        this.content.splice(index, 1);
        return true;
      }
    }
    return false;
  }

  /**
   * Replaces a run at a specific position
   * @param index - Position to replace (0-based)
   * @param run - New run
   * @returns True if replaced, false if index invalid or not a run
   * @example
   * ```typescript
   * const newRun = new Run('Replacement', { italic: true });
   * para.replaceRunAt(1, newRun);
   * ```
   */
  replaceRunAt(index: number, run: Run): boolean {
    if (index >= 0 && index < this.content.length) {
      const item = this.content[index];
      if (item instanceof Run) {
        this.content[index] = run;
        return true;
      }
    }
    return false;
  }

  /**
   * Finds text within the paragraph and returns run indices
   *
   * Searches through all runs in the paragraph and returns the indices
   * of runs that contain the search text.
   *
   * @param text - The text to search for
   * @param options - Optional search configuration
   * @param options.caseSensitive - If true, match case exactly (default: false)
   * @param options.wholeWord - If true, match whole words only (default: false)
   * @returns Array of run indices (0-based) that contain the search text
   *
   * @example
   * ```typescript
   * const indices = para.findText('important');
   * console.log(`Found in runs: ${indices.join(', ')}`);
   *
   * // Highlight all matching runs
   * for (const idx of indices) {
   *   const run = para.getRuns()[idx];
   *   run?.setHighlight('yellow');
   * }
   * ```
   */
  findText(text: string, options?: { caseSensitive?: boolean; wholeWord?: boolean }): number[] {
    const indices: number[] = [];
    const caseSensitive = options?.caseSensitive || false;
    const wholeWord = options?.wholeWord || false;

    const searchText = caseSensitive ? text : text.toLowerCase();

    for (let i = 0; i < this.content.length; i++) {
      const item = this.content[i];
      if (item instanceof Run) {
        const runText = caseSensitive ? item.getText() : item.getText().toLowerCase();

        if (wholeWord) {
          // Use word boundary regex
          const wordPattern = new RegExp(
            `\\b${searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`
          );
          if (wordPattern.test(runText)) {
            indices.push(i);
          }
        } else {
          // Simple substring search
          if (runText.includes(searchText)) {
            indices.push(i);
          }
        }
      }
    }

    return indices;
  }

  /**
   * Replaces text within paragraph runs
   *
   * Searches through all runs and replaces matching text while preserving
   * the original formatting of each run.
   *
   * @param find - The text to search for
   * @param replace - The replacement text
   * @param options - Optional search configuration
   * @param options.caseSensitive - If true, match case exactly (default: false)
   * @param options.wholeWord - If true, match whole words only (default: false)
   * @returns Number of replacements made
   *
   * @example
   * ```typescript
   * const count = para.replaceText('color', 'colour');
   * console.log(`Replaced ${count} occurrences`);
   * ```
   *
   * @example
   * ```typescript
   * // Case-sensitive whole word replacement
   * const count = para.replaceText('Error', 'Warning', {
   *   caseSensitive: true,
   *   wholeWord: true
   * });
   * ```
   */
  replaceText(
    find: string,
    replace: string,
    options?: { caseSensitive?: boolean; wholeWord?: boolean }
  ): number {
    let replacementCount = 0;
    const caseSensitive = options?.caseSensitive || false;
    const wholeWord = options?.wholeWord || false;

    for (const item of this.content) {
      if (item instanceof Run) {
        const originalText = item.getText();
        let newText = originalText;

        if (wholeWord) {
          // Use word boundary regex
          const wordPattern = new RegExp(
            `\\b${find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`,
            caseSensitive ? 'g' : 'gi'
          );
          const matches = originalText.match(wordPattern);
          if (matches) {
            replacementCount += matches.length;
            newText = originalText.replace(wordPattern, replace);
          }
        } else {
          // Simple substring replacement
          const searchPattern = new RegExp(
            find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
            caseSensitive ? 'g' : 'gi'
          );
          const matches = originalText.match(searchPattern);
          if (matches) {
            replacementCount += matches.length;
            newText = originalText.replace(searchPattern, replace);
          }
        }

        if (newText !== originalText) {
          item.setText(newText);
        }
      }
    }

    return replacementCount;
  }

  /**
   * Replaces all occurrences of a string, searching across run boundaries
   *
   * A simplified alias for `replaceTextCrossRun()` with no options — just
   * find and replace. Case-insensitive. Handles Word-fragmented text.
   *
   * @param find - Text to search for
   * @param replace - Replacement text
   * @returns Number of replacements made
   *
   * @example
   * ```typescript
   * para.replaceAll('old', 'new');
   * para.replaceAll('{{name}}', 'Alice');
   * ```
   */
  replaceAll(find: string, replace: string): number {
    return this.replaceTextCrossRun(find, replace);
  }

  /**
   * Searches for text across run boundaries, returning match positions
   *
   * Concatenates text from all runs and searches the combined string,
   * finding matches that may span multiple runs (e.g., `{{name}}` split
   * across runs as `{{` + `name` + `}}`). This is the read-only
   * counterpart to `replaceTextCrossRun()`.
   *
   * @param find - Text to search for (plain string)
   * @param options - Search options
   * @param options.caseSensitive - Match case exactly (default: false)
   * @returns Array of match objects with offset and matched text
   *
   * @example
   * ```typescript
   * // Find all placeholders, even if fragmented across runs
   * const matches = para.findTextCrossRun('{{name}}');
   * console.log(`Found ${matches.length} matches`);
   * for (const m of matches) {
   *   console.log(`  at offset ${m.offset}: "${m.text}"`);
   * }
   * ```
   */
  findTextCrossRun(
    find: string,
    options?: { caseSensitive?: boolean }
  ): { offset: number; text: string }[] {
    const caseSensitive = options?.caseSensitive ?? false;

    // Build full text from runs only
    let fullText = '';
    for (const item of this.content) {
      if (item instanceof Run) {
        fullText += item.getText();
      }
    }

    if (fullText.length === 0) return [];

    const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped, caseSensitive ? 'g' : 'gi');
    const results: { offset: number; text: string }[] = [];
    let m: RegExpExecArray | null;

    while ((m = regex.exec(fullText)) !== null) {
      results.push({ offset: m.index, text: m[0]! });
    }

    return results;
  }

  /**
   * Finds and replaces text that may span across multiple runs
   *
   * Unlike `replaceText()` which searches each run independently, this method
   * concatenates text across all runs and searches the combined string. This
   * handles the common case where Word splits text like `{{name}}` across
   * multiple runs (e.g., `{{` in one run, `name` in another, `}}` in a third).
   *
   * The replacement text inherits formatting from the first run of the match.
   * Runs that are fully consumed by the match are removed; partially consumed
   * runs are trimmed.
   *
   * @param find - Text to search for (plain string, case-insensitive by default)
   * @param replace - Replacement text
   * @param options - Search options
   * @param options.caseSensitive - Match case exactly (default: false)
   * @returns Number of replacements made
   *
   * @example
   * ```typescript
   * // Handle Word-fragmented placeholders
   * // Paragraph has runs: ["{{", "name", "}}"]
   * para.replaceTextCrossRun('{{name}}', 'Alice');
   * // Now has single run: "Alice"
   * ```
   *
   * @example
   * ```typescript
   * // Works with normal (non-fragmented) text too
   * para.replaceTextCrossRun('old text', 'new text');
   * ```
   */
  replaceTextCrossRun(
    find: string,
    replace: string,
    options?: { caseSensitive?: boolean }
  ): number {
    const caseSensitive = options?.caseSensitive ?? false;

    // Build run map: [{run, contentIndex, startOffset, endOffset}]
    const runMap: { run: Run; contentIndex: number; start: number; end: number }[] = [];
    let charPos = 0;
    for (let i = 0; i < this.content.length; i++) {
      const item = this.content[i]!;
      if (item instanceof Run) {
        const len = item.getText().length;
        if (len > 0) {
          runMap.push({ run: item, contentIndex: i, start: charPos, end: charPos + len });
        }
        charPos += len;
      }
    }

    if (runMap.length === 0) return 0;

    // Get full concatenated text
    const fullText = runMap.map((r) => r.run.getText()).join('');

    // Find all match positions
    const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escaped, caseSensitive ? 'g' : 'gi');
    const matches: { start: number; end: number }[] = [];
    let m: RegExpExecArray | null;
    while ((m = regex.exec(fullText)) !== null) {
      matches.push({ start: m.index, end: m.index + m[0]!.length });
    }

    if (matches.length === 0) return 0;

    // Process matches in reverse order to preserve earlier offsets
    for (let mi = matches.length - 1; mi >= 0; mi--) {
      const match = matches[mi]!;

      // Find all runs that overlap this match
      const affectedRuns = runMap.filter((r) => r.end > match.start && r.start < match.end);

      if (affectedRuns.length === 0) continue;

      const firstAffected = affectedRuns[0]!;
      const lastAffected = affectedRuns[affectedRuns.length - 1]!;

      // Calculate what to keep from the first and last runs
      const keepBefore =
        match.start > firstAffected.start
          ? firstAffected.run.getText().slice(0, match.start - firstAffected.start)
          : '';
      const keepAfter =
        match.end < lastAffected.end
          ? lastAffected.run.getText().slice(match.end - lastAffected.start)
          : '';

      // Set the first affected run to: kept-before + replacement + kept-after
      firstAffected.run.setText(keepBefore + replace + keepAfter);

      // Remove any middle/last runs that were fully or partially consumed
      // (iterate in reverse to preserve indices)
      for (let r = affectedRuns.length - 1; r >= 1; r--) {
        const runEntry = affectedRuns[r]!;
        const idx = this.content.indexOf(runEntry.run);
        if (idx !== -1) {
          this.content.splice(idx, 1);
        }
      }
    }

    return matches.length;
  }

  /**
   * Deletes a character range from the paragraph
   *
   * Removes text in the [start, end) range. Runs fully within the range
   * are removed entirely; boundary runs are trimmed. Non-Run content is
   * not affected.
   *
   * @param start - Start character offset (0-based, inclusive)
   * @param end - End character offset (0-based, exclusive)
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * const para = new Paragraph().addText('Hello Beautiful World');
   * para.deleteRange(5, 15);
   * para.getText(); // "Hello World"
   * ```
   */
  deleteRange(start: number, end: number): this {
    if (start >= end) return this;

    // Build run map
    const runMap: { index: number; start: number; end: number }[] = [];
    let charPos = 0;
    for (let i = 0; i < this.content.length; i++) {
      const item = this.content[i]!;
      if (item instanceof Run) {
        const len = item.getText().length;
        if (len > 0) {
          runMap.push({ index: i, start: charPos, end: charPos + len });
        }
        charPos += len;
      }
    }

    // Process in reverse to preserve indices
    for (let m = runMap.length - 1; m >= 0; m--) {
      const entry = runMap[m]!;
      const run = this.content[entry.index] as Run;

      // Skip runs entirely outside the range
      if (entry.end <= start || entry.start >= end) continue;

      const runStart = entry.start;
      const runEnd = entry.end;

      // How much of this run is in the delete range?
      const delStart = Math.max(0, start - runStart);
      const delEnd = Math.min(runEnd - runStart, end - runStart);
      const runLen = runEnd - runStart;

      if (delStart === 0 && delEnd === runLen) {
        // Entire run is deleted
        this.content.splice(entry.index, 1);
      } else {
        // Partial deletion — keep the parts outside the range
        const textBefore = run.getText().slice(0, delStart);
        const textAfter = run.getText().slice(delEnd);
        run.setText(textBefore + textAfter);
      }
    }

    return this;
  }

  /**
   * Truncates paragraph text to a maximum length
   *
   * If the paragraph text exceeds `maxLength`, content beyond that point
   * is removed and an optional suffix (default `'...'`) is appended.
   * The suffix counts toward the total length, so the final text is at
   * most `maxLength` characters.
   *
   * @param maxLength - Maximum total character length (including suffix)
   * @param suffix - Text to append when truncated (default: '...')
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * const para = new Paragraph().addText('The quick brown fox jumps over the lazy dog');
   * para.truncate(20);
   * para.getText(); // "The quick brown f..."
   *
   * para.truncate(10, ' [more]');
   * // getText() => "The [more]"
   * ```
   */
  truncate(maxLength: number, suffix = '...'): this {
    const text = this.getText();
    if (text.length <= maxLength) return this;

    const cutoff = Math.max(0, maxLength - suffix.length);
    this.deleteRange(cutoff, text.length);

    // Append suffix as a new run (inherits no formatting — plain text indicator)
    if (suffix) {
      this.addText(suffix);
    }

    return this;
  }

  /**
   * Wraps existing paragraph content with prefix and/or suffix text
   *
   * Inserts a Run with the prefix text before existing content and
   * a Run with the suffix text after it. Both inherit no formatting
   * by default, but optional formatting can be provided.
   *
   * @param prefix - Text to prepend (empty string to skip)
   * @param suffix - Text to append (empty string to skip)
   * @param formatting - Optional formatting for both prefix and suffix runs
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * const para = new Paragraph().addText('important');
   * para.wrap('[', ']');
   * para.getText(); // "[important]"
   *
   * // With formatting
   * para.wrap('NOTE: ', '', { bold: true });
   * ```
   */
  wrap(prefix: string, suffix: string, formatting?: RunFormatting): this {
    if (prefix) {
      const prefixRun = new Run(prefix, formatting);
      this.content.unshift(prefixRun);
    }
    if (suffix) {
      const suffixRun = new Run(suffix, formatting);
      this.content.push(suffixRun);
    }
    return this;
  }

  /**
   * Returns the Run that contains the character at the given offset
   *
   * Walks through Run elements in the paragraph, counting characters,
   * and returns the Run that contains the specified offset along with
   * the local offset within that run. Non-Run content is skipped.
   *
   * @param offset - Character position (0-based)
   * @returns Object with the Run and the local offset within it, or
   *   undefined if the offset is out of range or hits non-Run content
   *
   * @example
   * ```typescript
   * const para = new Paragraph();
   * para.addRun(new Run('Hello ', { bold: true }));
   * para.addRun(new Run('World'));
   *
   * const result = para.getRunAtOffset(8);
   * // result.run.getText() === 'World'
   * // result.localOffset === 2 (the 'r' in 'World')
   * ```
   */
  getRunAtOffset(offset: number): { run: Run; localOffset: number } | undefined {
    if (offset < 0) return undefined;

    let charPos = 0;
    for (const item of this.content) {
      if (!(item instanceof Run)) continue;

      const len = item.getText().length;
      if (charPos + len > offset) {
        return { run: item, localOffset: offset - charPos };
      }
      charPos += len;
    }

    return undefined;
  }

  /**
   * Returns the RunFormatting active at a given character position
   *
   * Finds the Run containing the character at `offset` and returns its
   * formatting. Useful for format inspection, format painting, and building
   * rich text cursors.
   *
   * @param offset - Character position (0-based)
   * @returns RunFormatting at that position, or undefined if out of range
   *
   * @example
   * ```typescript
   * const para = new Paragraph();
   * para.addRun(new Run('Bold text', { bold: true, color: 'FF0000' }));
   * para.addRun(new Run(' plain text'));
   *
   * const fmt = para.getFormattingAtOffset(3);
   * // fmt.bold === true
   * // fmt.color === 'FF0000'
   *
   * const fmt2 = para.getFormattingAtOffset(12);
   * // fmt2.bold === undefined
   * ```
   */
  getFormattingAtOffset(offset: number): RunFormatting | undefined {
    const result = this.getRunAtOffset(offset);
    return result ? result.run.getFormatting() : undefined;
  }

  /**
   * Applies formatting to a character range across runs
   *
   * Splits runs at the range boundaries as needed, then applies the given
   * formatting properties to every run that falls within [start, end).
   * Only modifies Run elements — hyperlinks, revisions, fields, and other
   * non-Run content are skipped (their characters still count toward offsets).
   *
   * Each property in the formatting object is applied via the corresponding
   * setter method (e.g. `bold` → `setBold()`, `color` → `setColor()`).
   *
   * @param start - Start character offset (0-based, inclusive)
   * @param end - End character offset (0-based, exclusive)
   * @param formatting - Partial RunFormatting to apply to the range
   * @returns This paragraph for chaining
   *
   * @example
   * ```typescript
   * // Bold "World" in "Hello World"
   * const para = new Paragraph().addText('Hello World');
   * para.applyFormattingToRange(6, 11, { bold: true });
   * ```
   *
   * @example
   * ```typescript
   * // Highlight a middle section
   * const para = new Paragraph().addText('The quick brown fox');
   * para.applyFormattingToRange(4, 9, {
   *   bold: true,
   *   color: 'FF0000',
   *   highlight: 'yellow',
   * });
   * // Results in 3 runs: "The " | "quick" (bold red highlighted) | " brown fox"
   * ```
   */
  applyFormattingToRange(start: number, end: number, formatting: Partial<RunFormatting>): this {
    if (start >= end) return this;

    // Build a map of content indices to their character offset ranges,
    // then work backwards to avoid index shifting during splicing.
    const contentMap: { index: number; start: number; end: number }[] = [];
    let charPos = 0;
    for (let i = 0; i < this.content.length; i++) {
      const item = this.content[i]!;
      let len = 0;
      if (item instanceof Run) {
        len = item.getText().length;
      }
      // Non-Run items don't contribute characters for our purposes
      // (they're skipped but still occupy content array slots)
      if (len > 0) {
        contentMap.push({ index: i, start: charPos, end: charPos + len });
      }
      charPos += len;
    }

    // Process runs in reverse order so splice indices stay valid
    for (let m = contentMap.length - 1; m >= 0; m--) {
      const entry = contentMap[m]!;
      const run = this.content[entry.index] as Run;

      // Skip runs entirely outside the range
      if (entry.end <= start || entry.start >= end) continue;

      const runStart = entry.start;
      const runEnd = entry.end;

      // Determine split points relative to this run
      const splitStart = Math.max(0, start - runStart);
      const splitEnd = Math.min(runEnd - runStart, end - runStart);
      const runLen = runEnd - runStart;

      if (splitStart === 0 && splitEnd === runLen) {
        // Entire run is within range — just apply formatting
        this.applyFormattingToRun(run, formatting);
      } else if (splitStart === 0) {
        // Range covers beginning of run — split off the tail
        const tail = run.splitAt(splitEnd);
        this.applyFormattingToRun(run, formatting);
        this.content.splice(entry.index + 1, 0, tail);
      } else if (splitEnd === runLen) {
        // Range covers end of run — split off the head
        const tail = run.splitAt(splitStart);
        this.applyFormattingToRun(tail, formatting);
        this.content.splice(entry.index + 1, 0, tail);
      } else {
        // Range is in the middle — three-way split
        const mid = run.splitAt(splitStart);
        const tail = mid.splitAt(splitEnd - splitStart);
        this.applyFormattingToRun(mid, formatting);
        this.content.splice(entry.index + 1, 0, mid, tail);
      }
    }

    return this;
  }

  /**
   * Applies partial RunFormatting to a run via setter methods.
   * @internal
   */
  private applyFormattingToRun(run: Run, formatting: Partial<RunFormatting>): void {
    if (formatting.bold !== undefined) run.setBold(formatting.bold);
    if (formatting.italic !== undefined) run.setItalic(formatting.italic);
    if (formatting.underline !== undefined) run.setUnderline(formatting.underline);
    if (formatting.strike !== undefined) run.setStrike(formatting.strike);
    if (formatting.subscript !== undefined) run.setSubscript(formatting.subscript);
    if (formatting.superscript !== undefined) run.setSuperscript(formatting.superscript);
    if (formatting.smallCaps !== undefined) run.setSmallCaps(formatting.smallCaps);
    if (formatting.allCaps !== undefined) run.setAllCaps(formatting.allCaps);
    if (formatting.font !== undefined) run.setFont(formatting.font);
    if (formatting.size !== undefined) run.setSize(formatting.size);
    if (formatting.color !== undefined) run.setColor(formatting.color);
    if (formatting.highlight !== undefined) run.setHighlight(formatting.highlight);
    if (formatting.characterSpacing !== undefined)
      run.setCharacterSpacing(formatting.characterSpacing);
    if (formatting.vanish !== undefined) run.setVanish(formatting.vanish);
  }

  /**
   * Merges adjacent runs that have identical formatting
   *
   * Walks the content array and combines consecutive Run elements whose
   * formatting is deeply equal. Non-Run content (hyperlinks, revisions,
   * fields, etc.) acts as a boundary — runs on opposite sides are never
   * merged even if their formatting matches.
   *
   * This is useful after operations that fragment runs, such as
   * `applyFormattingToRange()` or `Run.splitAt()`, to reduce file size
   * and simplify the content model.
   *
   * @returns Number of runs that were eliminated by merging
   *
   * @example
   * ```typescript
   * const para = new Paragraph();
   * para.addRun(new Run('Hello ', { bold: true }));
   * para.addRun(new Run('World', { bold: true }));
   * para.addRun(new Run('!'));
   *
   * const merged = para.consolidateRuns();
   * // merged === 1 (two bold runs became one)
   * // para now has 2 runs: "Hello World" (bold) and "!"
   * ```
   *
   * @example
   * ```typescript
   * // Clean up after applyFormattingToRange
   * para.applyFormattingToRange(0, 5, { bold: true });
   * para.applyFormattingToRange(5, 10, { bold: true });
   * para.consolidateRuns(); // merges the two bold fragments
   * ```
   */
  consolidateRuns(): number {
    if (this.content.length < 2) return 0;

    const consolidated: ParagraphContent[] = [];
    let eliminated = 0;

    for (const item of this.content) {
      if (!(item instanceof Run)) {
        // Non-run content acts as a merge boundary
        consolidated.push(item);
        continue;
      }

      const prev = consolidated.length > 0 ? consolidated[consolidated.length - 1] : undefined;

      if (
        prev instanceof Run &&
        isEqualFormatting(
          prev.getFormatting() as unknown as Record<string, unknown>,
          item.getFormatting() as unknown as Record<string, unknown>
        )
      ) {
        // Merge: append item's content to prev run
        const mergedContent = [...prev.getContent(), ...item.getContent()];
        const mergedRun = Run.createFromContent(mergedContent, deepClone(prev.getFormatting()));
        consolidated[consolidated.length - 1] = mergedRun;
        eliminated++;
      } else {
        consolidated.push(item);
      }
    }

    if (eliminated > 0) {
      this.content = consolidated;
    }

    return eliminated;
  }

  /**
   * Merges another paragraph's content into this one
   *
   * Appends all content (runs, fields, hyperlinks), bookmarks, and comments
   * from another paragraph to this one. The other paragraph is not modified.
   *
   * @param otherPara - The paragraph whose content to merge
   * @returns This paragraph instance for method chaining
   *
   * @example
   * ```typescript
   * const para1 = new Paragraph().addText('Hello ');
   * const para2 = new Paragraph().addText('World');
   * para1.mergeWith(para2);  // para1 now contains "Hello World"
   * ```
   */
  mergeWith(otherPara: Paragraph): this {
    // Add all content from other paragraph
    for (const item of otherPara.content) {
      if (item instanceof Run) {
        this.content.push(item.clone());
      } else {
        this.content.push(item);
      }
    }

    // Merge bookmarks
    this.bookmarksStart.push(...otherPara.bookmarksStart);
    this.bookmarksEnd.push(...otherPara.bookmarksEnd);

    // Merge comments
    this.commentsStart.push(...otherPara.commentsStart);
    this.commentsEnd.push(...otherPara.commentsEnd);

    return this;
  }

  /**
   * Clears direct run formatting from all runs in this paragraph
   *
   * This is useful when applying styles, as direct run formatting takes precedence
   * over style formatting in Word. By clearing direct formatting, the style's
   * formatting can take effect.
   *
   * @param properties - Optional array of specific properties to clear.
   *                     If not specified, clears ALL direct formatting.
   *                     Valid properties: 'bold', 'italic', 'underline', 'strike',
   *                     'font', 'size', 'color', 'highlight', 'subscript', 'superscript',
   *                     'smallCaps', 'allCaps', 'dstrike'
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * // Clear all direct formatting
   * para.clearDirectRunFormatting();
   *
   * // Clear only font and color
   * para.clearDirectRunFormatting(['font', 'color']);
   *
   * // Apply style and clear all formatting
   * para.setStyle('Heading1').clearDirectRunFormatting();
   * ```
   */
  clearDirectRunFormatting(properties?: string[]): this {
    const runs = this.getRuns();

    for (const run of runs) {
      const formatting = run.getFormatting();

      if (properties && properties.length > 0) {
        // Clear only specified properties
        const newFormatting: RunFormatting = { ...formatting };
        for (const prop of properties) {
          if (prop in newFormatting) {
            delete (newFormatting as any)[prop];
          }
        }

        // Create new run with cleared formatting
        const text = run.getText();
        const newRun = new Run(text, newFormatting);

        // Replace in content array
        const index = this.content.indexOf(run);
        if (index !== -1) {
          this.content[index] = newRun;
        }
      } else {
        // Clear ALL direct formatting - replace with plain text run
        const text = run.getText();
        const newRun = new Run(text, {});

        // Replace in content array
        const index = this.content.indexOf(run);
        if (index !== -1) {
          this.content[index] = newRun;
        }
      }
    }

    return this;
  }

  /**
   * Applies a style to this paragraph and optionally clears direct run formatting
   *
   * In Word, direct run formatting takes precedence over style formatting. This method
   * provides a way to apply a style while ensuring it takes effect by clearing
   * conflicting direct formatting.
   *
   * @param styleId - Style ID to apply (e.g., 'Normal', 'Heading1')
   * @param clearProperties - Optional array of run properties to clear.
   *                          If not specified, does NOT clear any formatting (style overlay behavior).
   *                          Pass empty array [] to clear ALL formatting.
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * // Apply Heading1 and clear all direct formatting
   * para.applyStyleAndClearFormatting('Heading1', []);
   *
   * // Apply Normal and clear only font and color
   * para.applyStyleAndClearFormatting('Normal', ['font', 'color']);
   *
   * // Apply Title but keep existing run formatting (overlay style)
   * para.applyStyleAndClearFormatting('Title');
   * ```
   */
  applyStyleAndClearFormatting(styleId: string, clearProperties?: string[]): this {
    // Apply the style
    this.setStyle(styleId);

    // Clear direct formatting if requested
    if (clearProperties !== undefined) {
      this.clearDirectRunFormatting(clearProperties.length === 0 ? undefined : clearProperties);
    }

    return this;
  }

  /**
   * Clears paragraph and run formatting that conflicts with a style definition.
   * Uses smart clearing per ECMA-376 §17.7.2 formatting hierarchy.
   *
   * This is critical because direct formatting in document.xml ALWAYS overrides
   * style definitions in styles.xml. To make style modifications take effect,
   * we must remove conflicting direct formatting.
   *
   * Strategy:
   * - Compare paragraph properties with style's paragraph properties
   * - Clear only properties that DIFFER from the style
   * - For each run, call run.clearFormattingConflicts() with style's run formatting
   * - Preserve style reference (pStyle element)
   *
   * @param styleDefinition - Style object containing both paragraph and run formatting
   * @returns This paragraph for method chaining
   * @example
   * ```typescript
   * // Style defines: left alignment, 14pt black Verdana, 6pt spacing
   * // Paragraph has: center alignment (conflict!), 120 twips spacing (conflict!)
   * // Runs have: red color (conflict!), 12pt size (conflict!), bold (not in style - keep!)
   * const style = stylesManager.getStyle('Heading2');
   * paragraph.clearDirectFormattingConflicts(style);
   * // Result: Alignment cleared, spacing cleared, runs' color/size cleared, bold preserved
   * ```
   */
  clearDirectFormattingConflicts(styleDefinition: {
    getProperties(): {
      paragraphFormatting?: ParagraphFormatting;
      runFormatting?: RunFormatting;
    };
  }): this {
    const styleProperties = styleDefinition.getProperties();
    const styleParagraphFormatting = styleProperties.paragraphFormatting || {};
    const styleRunFormatting = styleProperties.runFormatting || {};

    // Clear conflicting paragraph-level properties
    // Keep only pStyle (style reference) - clear everything else that conflicts
    const conflictingParaProps: (keyof ParagraphFormatting)[] = [];

    for (const key in this.formatting) {
      const propKey = key as keyof ParagraphFormatting;

      // Always preserve style reference
      if (propKey === 'style') {
        continue;
      }

      // Skip if style doesn't define this property
      if (styleParagraphFormatting[propKey] === undefined) {
        continue;
      }

      // Special handling for indentation - strip direct formatting so style/numbering values take effect
      if (propKey === 'indentation') {
        const paraIndent = this.formatting.indentation;
        const styleIndent = styleParagraphFormatting.indentation;

        // Preserve intentional zero-indent overrides
        // When style has left indent > 0 and paragraph explicitly sets left to 0,
        // this is an intentional override to remove the indent - preserve it
        if (paraIndent?.left === 0 && styleIndent?.left && styleIndent.left > 0) {
          continue;
        }

        // Clear direct indentation so style value takes effect (consistent with other properties)
        conflictingParaProps.push(propKey);
        continue;
      }

      // Handle complex objects (spacing, borders, etc.)
      if (
        propKey === 'spacing' ||
        propKey === 'borders' ||
        propKey === 'shading' ||
        propKey === 'numbering'
      ) {
        // Deep comparison for objects
        const paraValue = this.formatting[propKey];
        const styleValue = styleParagraphFormatting[propKey];

        if (!deepEqual(paraValue, styleValue)) {
          conflictingParaProps.push(propKey);
        }
      } else {
        // Simple value comparison
        if (this.formatting[propKey] !== styleParagraphFormatting[propKey]) {
          conflictingParaProps.push(propKey);
        }
      }
    }

    // Clear conflicting paragraph properties
    for (const prop of conflictingParaProps) {
      delete this.formatting[prop];
    }

    // Clear conflicting run-level properties from all runs
    for (const item of this.content) {
      if (item instanceof Run) {
        item.clearFormattingConflicts(styleRunFormatting);
      }
    }

    return this;
  }

  /**
   * Clears all direct formatting from this paragraph and its runs
   *
   * Removes all direct formatting properties from the paragraph and all its runs,
   * leaving only the style reference and text content. This is useful for ensuring
   * paragraphs match their defined style without formatting overrides.
   *
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * paragraph.clearDirectFormatting();
   * ```
   */
  clearDirectFormatting(): this {
    // Clear paragraph-level formatting (keep only style, numbering, and preserved flag)
    const style = this.formatting.style;
    const numbering = this.formatting.numbering;

    this.formatting = {};

    // Restore essential properties
    if (style) this.formatting.style = style;
    if (numbering) this.formatting.numbering = numbering;

    // Clear run-level formatting
    for (const item of this.content) {
      if (item instanceof Run) {
        item.clearFormatting();
      }
    }

    return this;
  }
}
