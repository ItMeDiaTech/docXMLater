/**
 * TableCell - Represents a cell in a table
 */

import { formatDateForXml } from "../utils/dateFormatting";
import { XMLBuilder, XMLElement } from "../xml/XMLBuilder";
import { Paragraph, TextDirection } from "./Paragraph";
import { Revision } from "./Revision";
import {
  BorderStyle as CommonBorderStyle,
  BorderDefinition,
  FourSidedBorders,
  CellVerticalAlignment as CommonCellVerticalAlignment,
  ShadingConfig,
  ShadingPattern,
  buildShadingAttributes,
  WidthType,
} from "./CommonTypes";

// ============================================================================
// RE-EXPORTED TYPES (for backward compatibility)
// ============================================================================

/**
 * Cell border style
 * @see CommonTypes.BorderStyle
 */
export type BorderStyle = CommonBorderStyle;

/**
 * Cell border definition
 * @see CommonTypes.BorderDefinition
 */
export interface CellBorder {
  style?: BorderStyle;
  size?: number; // Size in eighths of a point
  color?: string; // Hex color without #
}

/**
 * Cell borders
 * @see CommonTypes.FourSidedBorders
 */
export interface CellBorders {
  top?: CellBorder;
  bottom?: CellBorder;
  left?: CellBorder;
  right?: CellBorder;
  /** Diagonal border from top-left to bottom-right per ECMA-376 Part 1 §17.4.84 */
  tl2br?: CellBorder;
  /** Diagonal border from top-right to bottom-left per ECMA-376 Part 1 §17.4.85 */
  tr2bl?: CellBorder;
}

/**
 * Cell shading/background
 * @see ShadingConfig in CommonTypes.ts for the canonical definition
 */
export type CellShading = ShadingConfig;

/**
 * Vertical alignment in cell
 * @see CommonTypes.CellVerticalAlignment
 */
export type CellVerticalAlignment = CommonCellVerticalAlignment;

/**
 * Cell margins (spacing inside cell borders)
 * Per ECMA-376 Part 1 §17.4.43
 */
export interface CellMargins {
  top?: number; // Margin in twips
  bottom?: number; // Margin in twips
  left?: number; // Margin in twips
  right?: number; // Margin in twips
}

/**
 * Cell width type
 * Per ECMA-376 Part 1 §17.18.87
 */
export type CellWidthType = "auto" | "dxa" | "pct";

/**
 * Vertical merge type for cells
 * Per ECMA-376 Part 1 §17.4.85
 */
export type VerticalMerge = "restart" | "continue";

/**
 * Cell formatting options
 */
export interface CellFormatting {
  width?: number; // Width in twips
  widthType?: CellWidthType; // Width type (auto, dxa, pct)
  borders?: CellBorders;
  shading?: CellShading;
  verticalAlignment?: CellVerticalAlignment;
  columnSpan?: number; // Number of columns to span (gridSpan)
  rowSpan?: number; // Number of rows to span (gridSpan)
  margins?: CellMargins; // Cell margins (spacing inside borders)
  textDirection?: TextDirection; // Text flow direction
  fitText?: boolean; // Fit text to cell width
  noWrap?: boolean; // Prevent text wrapping
  hideMark?: boolean; // Hide end-of-cell mark
  cnfStyle?: string; // Conditional formatting style (14-char binary string)
  vMerge?: VerticalMerge; // Vertical cell merge
  hMerge?: 'restart' | 'continue'; // Legacy horizontal merge (w:hMerge) per ECMA-376 Part 1 §17.4.22
  /** Cell headers attribute for accessibility per ECMA-376 Part 1 §17.4.26 */
  headers?: string;
}

/**
 * Raw nested content stored as XML to preserve nested tables/SDTs
 * Position indicates after which paragraph index the content appears
 */
export interface RawNestedContent {
  position: number; // Content appears after this paragraph index
  xml: string; // Raw XML content
  type: "table" | "sdt"; // Type of nested content
}

/**
 * Table cell property change tracking (w:tcPrChange)
 * Per ECMA-376 Part 1 §17.13.5.37
 */
export interface TcPrChange {
  author: string;
  date: string;
  id: string;
  previousProperties: Record<string, any>;
}

/**
 * Represents a table cell
 */
export class TableCell {
  private paragraphs: Paragraph[] = [];
  private formatting: CellFormatting;
  /** Raw nested content (tables, SDTs) stored as XML for passthrough */
  private rawNestedContent: RawNestedContent[] = [];
  /** Parent row reference (if cell is inside a table row) */
  private _parentRow?: import('./TableRow').TableRow;
  /** Table cell revision (w:cellIns, w:cellDel, w:cellMerge) per ECMA-376 Part 1 §17.13.5.4-5.6 */
  private cellRevision?: Revision;
  /** Tracking context for automatic change tracking */
  private trackingContext?: import('../tracking/TrackingContext').TrackingContext;
  /** Table cell property change tracking (w:tcPrChange) */
  private tcPrChange?: TcPrChange;

  /**
   * Creates a new TableCell
   * @param formatting - Cell formatting options
   */
  constructor(formatting: CellFormatting = {}) {
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
   * Gets the table cell property change tracking info
   * @returns tcPrChange or undefined
   */
  getTcPrChange(): TcPrChange | undefined {
    return this.tcPrChange;
  }

  /**
   * Sets the table cell property change tracking info
   * @param change - Property change info or undefined to clear
   */
  setTcPrChange(change: TcPrChange | undefined): void {
    this.tcPrChange = change;
  }

  /**
   * Clears the table cell property change tracking
   */
  clearTcPrChange(): void {
    this.tcPrChange = undefined;
  }

  /**
   * Adds a paragraph to the cell
   * @param paragraph - Paragraph to add
   * @returns This cell for chaining
   */
  addParagraph(paragraph: Paragraph): this {
    this.paragraphs.push(paragraph);
    paragraph._setParentCell(this);
    // Propagate StylesManager from table if available
    const stylesManager = this._parentRow?._getParentTable()?._getStylesManager();
    if (stylesManager) {
      paragraph._setStylesManager(stylesManager);
    }
    return this;
  }

  /**
   * Creates and adds a new paragraph with text
   * @param text - Text content
   * @returns The created paragraph
   */
  createParagraph(text?: string): Paragraph {
    const para = new Paragraph();
    if (text) {
      para.addText(text);
    }
    this.paragraphs.push(para);
    para._setParentCell(this);
    // Propagate StylesManager from table if available
    const stylesManager = this._parentRow?._getParentTable()?._getStylesManager();
    if (stylesManager) {
      para._setStylesManager(stylesManager);
    }
    return para;
  }

  /**
   * Gets all paragraphs in the cell
   * @returns Array of paragraphs
   */
  getParagraphs(): Paragraph[] {
    return [...this.paragraphs];
  }

  /**
   * Removes a paragraph at the specified index
   * @param index - Index of paragraph to remove
   * @returns True if removed, false if index out of bounds
   */
  removeParagraph(index: number): boolean {
    if (index < 0 || index >= this.paragraphs.length) {
      return false;
    }

    // When tracking enabled, wrap content in w:del instead of removing
    if (this.trackingContext?.isEnabled()) {
      const paragraph = this.paragraphs[index]!;
      const runs = paragraph.getRuns();
      if (runs.length > 0) {
        const author = this.trackingContext.getAuthor();
        const deletion = Revision.createDeletion(author, runs);
        this.trackingContext.getRevisionManager().register(deletion);
        paragraph.addRevision(deletion);
      }
      return true;
    }

    const removed = this.paragraphs.splice(index, 1);
    const removedPara = removed[0];
    if (removedPara) {
      removedPara._setParentCell(undefined);
    }

    // Update raw nested content positions
    // Any nested content positioned AFTER the removed paragraph needs its position decremented
    // This maintains correct relative positioning for nested tables, SDTs, etc.
    for (const item of this.rawNestedContent) {
      if (item.position > index) {
        item.position--;
      }
    }

    return true;
  }

  /**
   * Adds a paragraph at the specified index
   * @param index - Index to insert at
   * @param paragraph - Paragraph to add
   * @returns This cell for chaining
   */
  addParagraphAt(index: number, paragraph: Paragraph): this {
    if (index < 0) {
      index = 0;
    }

    // Determine actual insertion index
    const actualIndex = index >= this.paragraphs.length ? this.paragraphs.length : index;

    if (index >= this.paragraphs.length) {
      this.paragraphs.push(paragraph);
    } else {
      this.paragraphs.splice(index, 0, paragraph);
    }

    // Update raw nested content positions
    // Any nested content positioned AT OR AFTER the insertion point needs its position incremented
    // This maintains correct relative positioning for nested tables, SDTs, etc.
    for (const item of this.rawNestedContent) {
      if (item.position >= actualIndex) {
        item.position++;
      }
    }

    paragraph._setParentCell(this);
    // Propagate StylesManager from table if available
    const stylesManager = this._parentRow?._getParentTable()?._getStylesManager();
    if (stylesManager) {
      paragraph._setStylesManager(stylesManager);
    }

    // When tracking enabled, wrap paragraph content in w:ins revision
    if (this.trackingContext?.isEnabled()) {
      const runs = paragraph.getRuns();
      if (runs.length > 0) {
        const author = this.trackingContext.getAuthor();
        const insertion = Revision.createInsertion(author, runs);
        this.trackingContext.getRevisionManager().register(insertion);
        paragraph.addRevision(insertion);
      }
    }

    return this;
  }

  /**
   * Gets the text content of all paragraphs
   * @returns Combined text
   */
  getText(): string {
    return this.paragraphs.map((p) => p.getText()).join("\n");
  }

  /**
   * Gets all fields from paragraphs in this cell
   *
   * Collects all Field and ComplexField instances from every paragraph
   * in the table cell.
   *
   * @returns Array of fields (Field and ComplexField instances)
   *
   * @example
   * ```typescript
   * const cell = table.getCell(0, 0);
   * const fields = cell?.getFields() || [];
   * console.log(`Cell has ${fields.length} fields`);
   * ```
   */
  getFields(): (import("./Field").Field | import("./Field").ComplexField)[] {
    const fields: (import("./Field").Field | import("./Field").ComplexField)[] = [];
    for (const para of this.paragraphs) {
      fields.push(...para.getFields());
    }
    return fields;
  }

  /**
   * Finds fields matching a predicate function
   *
   * Searches through all fields in the cell and returns those matching
   * the specified criteria.
   *
   * @param predicate - Function to test each field
   * @returns Array of matching fields
   *
   * @example
   * ```typescript
   * // Find all PAGE fields
   * const pageFields = cell.findFields(f =>
   *   f.getInstruction().startsWith('PAGE')
   * );
   *
   * // Find fields with specific switches
   * const mergeFields = cell.findFields(f =>
   *   f.getInstruction().includes('MERGEFIELD')
   * );
   * ```
   */
  findFields(
    predicate: (
      field: import("./Field").Field | import("./Field").ComplexField
    ) => boolean
  ): (import("./Field").Field | import("./Field").ComplexField)[] {
    return this.getFields().filter(predicate);
  }

  /**
   * Removes all fields from paragraphs in this cell
   *
   * Iterates through each paragraph and removes all fields,
   * preserving text content.
   *
   * @returns Count of fields removed
   *
   * @example
   * ```typescript
   * const count = cell.removeAllFields();
   * console.log(`Removed ${count} fields from cell`);
   * ```
   */
  removeAllFields(): number {
    let count = 0;
    for (const para of this.paragraphs) {
      count += para.removeAllFields();
    }
    return count;
  }

  /**
   * Sets cell width
   * @param twips - Width in twips
   * @returns This cell for chaining
   */
  setWidth(twips: number): this {
    const prev = this.formatting.width;
    this.formatting.width = twips;
    if (this.trackingContext?.isEnabled() && prev !== twips) {
      this.trackingContext.trackTableChange(this, 'width', prev, twips);
    }
    return this;
  }

  /**
   * Sets cell borders
   * @param borders - Border definitions
   * @returns This cell for chaining
   */
  setBorders(borders: CellBorders): this {
    const prev = this.formatting.borders;
    this.formatting.borders = borders;
    if (this.trackingContext?.isEnabled() && prev !== borders) {
      this.trackingContext.trackTableChange(this, 'borders', prev, borders);
    }
    return this;
  }

  /**
   * Sets cell shading/background
   * @param shading - Shading definition
   * @returns This cell for chaining
   */
  setShading(shading: CellShading): this {
    const prev = this.formatting.shading;
    this.formatting.shading = shading;
    if (this.trackingContext?.isEnabled() && prev !== shading) {
      this.trackingContext.trackTableChange(this, 'shading', prev, shading);
    }
    return this;
  }

  /**
   * Sets vertical alignment
   * @param alignment - Vertical alignment
   * @returns This cell for chaining
   */
  setVerticalAlignment(alignment: CellVerticalAlignment): this {
    const prev = this.formatting.verticalAlignment;
    this.formatting.verticalAlignment = alignment;
    if (this.trackingContext?.isEnabled() && prev !== alignment) {
      this.trackingContext.trackTableChange(this, 'verticalAlignment', prev, alignment);
    }
    return this;
  }

  /**
   * Sets column span (merge cells horizontally)
   * @param span - Number of columns to span
   * @returns This cell for chaining
   */
  setColumnSpan(span: number): this {
    const prev = this.formatting.columnSpan;
    this.formatting.columnSpan = span;
    if (this.trackingContext?.isEnabled() && prev !== span) {
      this.trackingContext.trackTableChange(this, 'columnSpan', prev, span);
    }
    return this;
  }

  /**
   * Gets the number of columns this cell spans (gridSpan)
   * @returns Column span, defaults to 1 if not set
   */
  getColumnSpan(): number {
    return this.formatting.columnSpan || 1;
  }

  /**
   * Gets whether to fit text to cell width
   * @returns true if fit text is enabled, false otherwise
   */
  getFitText(): boolean {
    return this.formatting.fitText ?? false;
  }

  /**
   * Gets whether text wrapping is prevented in cell
   * @returns true if no-wrap is enabled, false otherwise
   */
  getNoWrap(): boolean {
    return this.formatting.noWrap ?? false;
  }

  /**
   * Gets whether the end-of-cell mark is hidden
   * @returns true if hidden, false otherwise
   */
  getHideMark(): boolean {
    return this.formatting.hideMark ?? false;
  }

  /**
   * Gets the conditional formatting style (cnfStyle) for this cell
   * @returns 14-character binary string or undefined if not set
   */
  getCnfStyle(): string | undefined {
    return this.formatting.cnfStyle;
  }

  /**
   * Validates a margin value
   * @param value - Margin value in twips
   * @param side - Name of the margin side (for error messages)
   * @throws {Error} If margin is negative or exceeds maximum
   * @private
   */
  private validateMargin(value: number | undefined, side: string): void {
    if (value === undefined) return;

    // Margins must be non-negative
    if (value < 0) {
      throw new Error(
        `Invalid ${side} margin: ${value} twips. Cell margins cannot be negative.`
      );
    }

    // Maximum reasonable margin (1 inch = 1440 twips)
    // Word typically allows up to several inches, but we set a reasonable limit
    const MAX_MARGIN_TWIPS = 14400; // 10 inches
    if (value > MAX_MARGIN_TWIPS) {
      throw new Error(
        `Invalid ${side} margin: ${value} twips exceeds maximum of ${MAX_MARGIN_TWIPS} twips (10 inches).`
      );
    }
  }

  /**
   * Sets cell margins (spacing inside cell borders)
   * Per ECMA-376 Part 1 §17.4.43
   * @param margins - Margin definitions for each side
   * @returns This cell for chaining
   * @throws {Error} If any margin value is negative or exceeds maximum
   */
  setMargins(margins: CellMargins): this {
    // Validate each margin
    this.validateMargin(margins.top, "top");
    this.validateMargin(margins.bottom, "bottom");
    this.validateMargin(margins.left, "left");
    this.validateMargin(margins.right, "right");

    const prev = this.formatting.margins;
    this.formatting.margins = margins;
    if (this.trackingContext?.isEnabled() && prev !== margins) {
      this.trackingContext.trackTableChange(this, 'margins', prev, margins);
    }
    return this;
  }

  /**
   * Sets all cell margins to the same value
   * @param margin - Margin in twips to apply to all sides
   * @returns This cell for chaining
   * @throws {Error} If margin value is negative or exceeds maximum
   */
  setAllMargins(margin: number): this {
    this.validateMargin(margin, "all");
    return this.setMargins({ top: margin, bottom: margin, left: margin, right: margin });
  }

  /**
   * Sets text direction for cell content
   * Per ECMA-376 Part 1 §17.4.72
   * @param direction - Text flow direction
   *   - 'lrTb': Left-to-right, top-to-bottom (default)
   *   - 'tbRl': Top-to-bottom, right-to-left (vertical text, East Asian)
   *   - 'btLr': Bottom-to-top, left-to-right (vertical text)
   *   - 'lrTbV': Left-to-right, top-to-bottom, vertical
   *   - 'tbRlV': Top-to-bottom, right-to-left, vertical
   *   - 'tbLrV': Top-to-bottom, left-to-right, vertical
   * @returns This cell for chaining
   */
  setTextDirection(direction: TextDirection): this {
    const prev = this.formatting.textDirection;
    this.formatting.textDirection = direction;
    if (this.trackingContext?.isEnabled() && prev !== direction) {
      this.trackingContext.trackTableChange(this, 'textDirection', prev, direction);
    }
    return this;
  }

  /**
   * Sets whether to fit text to cell width
   * Per ECMA-376 Part 1 §17.4.68
   * @param fit - Whether to expand/compress text to fit cell width
   * @returns This cell for chaining
   */
  setFitText(fit = true): this {
    const prev = this.formatting.fitText;
    this.formatting.fitText = fit;
    if (this.trackingContext?.isEnabled() && prev !== fit) {
      this.trackingContext.trackTableChange(this, 'fitText', prev, fit);
    }
    return this;
  }

  /**
   * Sets whether to prevent text wrapping in cell
   * Per ECMA-376 Part 1 §17.4.34
   * @param noWrap - Whether to prevent wrapping (default: true)
   * @returns This cell for chaining
   */
  setNoWrap(noWrap = true): this {
    const prev = this.formatting.noWrap;
    this.formatting.noWrap = noWrap;
    if (this.trackingContext?.isEnabled() && prev !== noWrap) {
      this.trackingContext.trackTableChange(this, 'noWrap', prev, noWrap);
    }
    return this;
  }

  /**
   * Sets whether to hide the end-of-cell mark
   * Per ECMA-376 Part 1 §17.4.24
   * @param hide - Whether to ignore cell end mark in height calculations (default: true)
   * @returns This cell for chaining
   */
  setHideMark(hide = true): this {
    const prev = this.formatting.hideMark;
    this.formatting.hideMark = hide;
    if (this.trackingContext?.isEnabled() && prev !== hide) {
      this.trackingContext.trackTableChange(this, 'hideMark', prev, hide);
    }
    return this;
  }

  /**
   * Sets conditional formatting style for this cell
   * Per ECMA-376 Part 1 §17.4.7
   * @param cnfStyle - 14-character binary string representing which conditional formats to apply
   *   Each bit position controls a different conditional format (e.g., "100000000000" for first row)
   * @returns This cell for chaining
   */
  setConditionalStyle(cnfStyle: string): this {
    const prev = this.formatting.cnfStyle;
    this.formatting.cnfStyle = cnfStyle;
    if (this.trackingContext?.isEnabled() && prev !== cnfStyle) {
      this.trackingContext.trackTableChange(this, 'cnfStyle', prev, cnfStyle);
    }
    return this;
  }

  /**
   * Sets cell width with type specification
   * Per ECMA-376 Part 1 §17.4.81
   * @param width - Width value
   * @param type - Width type: 'auto' (automatic), 'dxa' (twips), or 'pct' (percentage * 50)
   * @returns This cell for chaining
   */
  setWidthType(width: number, type: CellWidthType = "dxa"): this {
    const prevWidth = this.formatting.width;
    const prevType = this.formatting.widthType;
    this.formatting.width = width;
    this.formatting.widthType = type;
    if (this.trackingContext?.isEnabled()) {
      if (prevWidth !== width) {
        this.trackingContext.trackTableChange(this, 'width', prevWidth, width);
      }
      if (prevType !== type) {
        this.trackingContext.trackTableChange(this, 'widthType', prevType, type);
      }
    }
    return this;
  }

  /**
   * Sets vertical merge for this cell
   * Per ECMA-376 Part 1 §17.4.85
   * @param merge - Vertical merge type:
   *   - 'restart': Start a new vertically merged region (top cell)
   *   - 'continue': Continue the current vertically merged region (cells below)
   * @returns This cell for chaining
   */
  setVerticalMerge(merge: VerticalMerge | undefined): this {
    const prev = this.formatting.vMerge;
    this.formatting.vMerge = merge;
    if (this.trackingContext?.isEnabled() && prev !== merge) {
      this.trackingContext.trackTableChange(this, 'vMerge', prev, merge);
    }
    return this;
  }

  /**
   * Sets the legacy horizontal merge state per ECMA-376 Part 1 §17.4.22
   * @param merge - 'restart' to start a new merge region, 'continue' to continue, or undefined to clear
   * @returns This cell for chaining
   */
  setHorizontalMerge(merge: 'restart' | 'continue' | undefined): this {
    const prev = this.formatting.hMerge;
    this.formatting.hMerge = merge;
    if (this.trackingContext?.isEnabled() && prev !== merge) {
      this.trackingContext.trackTableChange(this, 'hMerge', prev, merge);
    }
    return this;
  }

  // ============================================================================
  // CELL REVISIONS (w:cellIns, w:cellDel, w:cellMerge)
  // ============================================================================

  /**
   * Sets the revision marker for this cell
   * Per ECMA-376 Part 1 §17.13.5.4-5.6
   *
   * Table cell revisions track structural changes to table cells:
   * - tableCellInsert (w:cellIns): Cell was inserted
   * - tableCellDelete (w:cellDel): Cell was deleted
   * - tableCellMerge (w:cellMerge): Cell merge/split operation
   *
   * @param revision - Revision marker for this cell
   * @returns This cell for chaining
   *
   * @example
   * ```typescript
   * const revision = new Revision({
   *   id: 1,
   *   author: 'Alice',
   *   date: new Date(),
   *   type: 'tableCellInsert',
   *   content: [],
   * });
   * cell.setCellRevision(revision);
   * ```
   */
  setCellRevision(revision: Revision): this {
    this.cellRevision = revision;
    return this;
  }

  /**
   * Gets the revision marker for this cell
   *
   * @returns The cell revision if present, undefined otherwise
   *
   * @example
   * ```typescript
   * const revision = cell.getCellRevision();
   * if (revision) {
   *   console.log(`Cell ${revision.getType()} by ${revision.getAuthor()}`);
   * }
   * ```
   */
  getCellRevision(): Revision | undefined {
    return this.cellRevision;
  }

  /**
   * Checks if this cell has a revision marker
   *
   * @returns True if cell has a revision (insert, delete, or merge)
   */
  hasCellRevision(): boolean {
    return this.cellRevision !== undefined;
  }

  /**
   * Clears the revision marker for this cell
   *
   * @returns This cell for chaining
   */
  clearCellRevision(): this {
    this.cellRevision = undefined;
    return this;
  }

  // ============================================================================
  // CONVENIENCE METHODS (for easier paragraph manipulation)
  // ============================================================================

  /**
   * Sets text alignment for all paragraphs in the cell
   *
   * Applies the specified horizontal alignment to every paragraph
   * in this cell.
   *
   * @param alignment - Paragraph alignment (left, center, right, both)
   * @returns This cell for chaining
   *
   * @example
   * ```typescript
   * cell.setTextAlignment('center');
   * ```
   */
  setTextAlignment(alignment: import('./Paragraph').ParagraphAlignment): this {
    for (const para of this.paragraphs) {
      para.setAlignment(alignment);
    }
    return this;
  }

  /**
   * Sets the style for all paragraphs in the cell
   *
   * Applies the specified style ID to every paragraph in this cell.
   *
   * @param styleId - Style ID to apply
   * @returns This cell for chaining
   *
   * @example
   * ```typescript
   * cell.setAllParagraphsStyle('TableContent');
   * ```
   */
  setAllParagraphsStyle(styleId: string): this {
    for (const para of this.paragraphs) {
      para.setStyle(styleId);
    }
    return this;
  }

  /**
   * Sets font for all runs in the cell
   *
   * Applies the specified font to every text run in every paragraph
   * in this cell.
   *
   * @param fontName - Font name to apply
   * @returns Number of runs modified
   *
   * @example
   * ```typescript
   * const count = cell.setAllRunsFont('Arial');
   * ```
   */
  setAllRunsFont(fontName: string): number {
    let count = 0;
    for (const para of this.paragraphs) {
      for (const run of para.getRuns()) {
        run.setFont(fontName);
        count++;
      }
    }
    return count;
  }

  /**
   * Sets font size for all runs in the cell
   *
   * Applies the specified font size to every text run in every paragraph
   * in this cell.
   *
   * @param size - Font size in half-points (e.g., 24 = 12pt)
   * @returns Number of runs modified
   *
   * @example
   * ```typescript
   * const count = cell.setAllRunsSize(22); // 11pt
   * ```
   */
  setAllRunsSize(size: number): number {
    let count = 0;
    for (const para of this.paragraphs) {
      for (const run of para.getRuns()) {
        run.setSize(size);
        count++;
      }
    }
    return count;
  }

  /**
   * Sets color for all runs in the cell
   *
   * Applies the specified color to every text run in every paragraph
   * in this cell.
   *
   * @param color - Hex color code (e.g., 'FF0000', '#0000FF')
   * @returns Number of runs modified
   *
   * @example
   * ```typescript
   * const count = cell.setAllRunsColor('000000'); // Black
   * ```
   */
  setAllRunsColor(color: string): number {
    let count = 0;
    for (const para of this.paragraphs) {
      for (const run of para.getRuns()) {
        run.setColor(color);
        count++;
      }
    }
    return count;
  }

  /**
   * Gets the cell formatting
   * @returns Cell formatting
   */
  getFormatting(): CellFormatting {
    return { ...this.formatting };
  }

  // ============================================================================
  // Individual Formatting Getters
  // ============================================================================

  /**
   * Gets the cell width in twips
   * @returns Width in twips or undefined if not set
   */
  getWidth(): number | undefined {
    return this.formatting.width;
  }

  /**
   * Gets the cell width type
   * @returns Width type ('auto', 'dxa', 'pct', 'nil') or undefined
   */
  getWidthType(): string | undefined {
    return this.formatting.widthType;
  }

  /**
   * Gets the vertical alignment of cell content
   * @returns Vertical alignment ('top', 'center', 'bottom') or undefined
   */
  getVerticalAlignment(): string | undefined {
    return this.formatting.verticalAlignment;
  }

  /**
   * Gets the vertical merge state
   * @returns Vertical merge state ('restart', 'continue') or undefined
   */
  getVerticalMerge(): VerticalMerge | undefined {
    return this.formatting.vMerge;
  }

  /**
   * Gets the legacy horizontal merge state
   * @returns Horizontal merge state ('restart', 'continue') or undefined
   */
  getHorizontalMerge(): 'restart' | 'continue' | undefined {
    return this.formatting.hMerge;
  }

  /**
   * Sets the cell headers attribute for accessibility
   * Links data cells to header cells per ECMA-376 Part 1 §17.4.26
   * @param headers - Space-separated list of header cell IDs
   * @returns This cell for chaining
   */
  setHeaders(headers: string): this {
    this.formatting.headers = headers;
    return this;
  }

  /**
   * Gets the cell headers attribute
   * @returns Headers string or undefined
   */
  getHeaders(): string | undefined {
    return this.formatting.headers;
  }

  /**
   * Gets the cell margins
   * @returns Margins object with top, right, bottom, left or undefined
   */
  getMargins(): CellMargins | undefined {
    return this.formatting.margins;
  }

  /**
   * Gets the cell borders
   * @returns Borders object or undefined
   */
  getBorders(): CellBorders | undefined {
    return this.formatting.borders;
  }

  /**
   * Gets the cell shading/background
   * @returns Shading object or undefined
   */
  getShading(): CellShading | undefined {
    return this.formatting.shading;
  }

  /**
   * Gets the text direction for the cell
   * @returns Text direction or undefined
   */
  getTextDirection(): string | undefined {
    return this.formatting.textDirection;
  }

  // ============================================================================
  // RAW NESTED CONTENT (Tables, SDTs preserved as XML)
  // ============================================================================

  /**
   * Adds raw nested content (table or SDT) to the cell
   * Used during parsing to preserve nested tables that cannot be fully modeled
   * @param position - Paragraph index after which this content appears (0 = before first paragraph)
   * @param xml - Raw XML content
   * @param type - Type of content ('table' or 'sdt')
   * @returns This cell for chaining
   */
  addRawNestedContent(
    position: number,
    xml: string,
    type: "table" | "sdt" = "table"
  ): this {
    this.rawNestedContent.push({ position, xml, type });
    return this;
  }

  /**
   * Gets all raw nested content in this cell
   * @returns Array of raw nested content items
   */
  getRawNestedContent(): RawNestedContent[] {
    return [...this.rawNestedContent];
  }

  /**
   * Checks if this cell has any nested tables
   * @returns True if cell contains nested tables stored as raw XML
   */
  hasNestedTables(): boolean {
    return this.rawNestedContent.some((c) => c.type === "table");
  }

  /**
   * Checks if this cell has any raw nested content (tables or SDTs)
   * @returns True if cell contains any raw nested content
   */
  hasRawNestedContent(): boolean {
    return this.rawNestedContent.length > 0;
  }

  /**
   * Clears all raw nested content from this cell
   * @returns This cell for chaining
   */
  clearRawNestedContent(): this {
    this.rawNestedContent = [];
    return this;
  }

  /**
   * Updates raw nested content at a specific index
   * Used for revision acceptance in nested tables
   * @param index - Index in the rawNestedContent array
   * @param xml - New XML content
   * @returns True if updated, false if index out of bounds
   */
  updateRawNestedContent(index: number, xml: string): boolean {
    if (index < 0 || index >= this.rawNestedContent.length) {
      return false;
    }
    const item = this.rawNestedContent[index];
    if (item) {
      item.xml = xml;
      return true;
    }
    return false;
  }

  // ============================================================================
  // TRAILING BLANK PARAGRAPH REMOVAL
  // ============================================================================

  /**
   * Removes trailing blank paragraphs from this cell.
   * A trailing blank is a blank paragraph at the end of the cell, after all content.
   * This respects ECMA-376 requirement of at least one paragraph per cell.
   *
   * @param options.ignorePreserveFlag - If true, removes trailing blanks even if marked preserved (default: false)
   * @returns Number of paragraphs removed
   *
   * @example
   * ```typescript
   * // Remove trailing blanks, respecting preserve flags
   * const removed = cell.removeTrailingBlankParagraphs();
   *
   * // Remove all trailing blanks, ignoring preserve flags
   * const removed = cell.removeTrailingBlankParagraphs({ ignorePreserveFlag: true });
   * ```
   */
  removeTrailingBlankParagraphs(options?: { ignorePreserveFlag?: boolean }): number {
    let removed = 0;
    const ignorePreserve = options?.ignorePreserveFlag ?? false;

    // Work backwards from end of paragraphs array
    while (this.paragraphs.length > 1) {
      const lastIndex = this.paragraphs.length - 1;
      const lastPara = this.paragraphs[lastIndex];

      if (!lastPara) break;

      // Check if this is a blank paragraph
      const isBlank = this.isParaBlank(lastPara);

      // Stop if not blank
      if (!isBlank) break;

      // Stop if preserved and we're respecting preserve flags
      if (!ignorePreserve && lastPara.isPreserved()) break;

      // Check if there's raw nested content positioned after this paragraph
      // If so, we should NOT remove this trailing blank as it maintains structure
      const hasContentAfter = this.rawNestedContent.some(
        (item) => item.position >= lastIndex
      );
      if (hasContentAfter) break;

      this.removeParagraph(lastIndex);
      removed++;
    }

    return removed;
  }

  /**
   * Checks if a paragraph is blank (no meaningful content).
   * A paragraph is considered blank if it has no text, images, shapes, hyperlinks, fields,
   * cnfStyle (conditional formatting), or other structural elements.
   *
   * IMPORTANT: cnfStyle preservation is critical! When a paragraph with cnfStyle is removed,
   * Word may apply default table style conditional formatting to the cell, causing unexpected
   * shading changes. A paragraph with cnfStyle="000000000000" (no conditionals) keeps the cell
   * from matching table style conditionals like firstRow or band1Horz.
   *
   * @private
   */
  private isParaBlank(para: Paragraph): boolean {
    // Check for text content
    const text = para.getText().trim();
    if (text !== "") return false;

    // Check for cnfStyle (conditional formatting) - critical for shading preservation
    // Even a "blank" cnfStyle like "000000000000" is meaningful as it prevents
    // the cell from inheriting table style conditional formatting
    const cnfStyle = para.getTableConditionalStyle();
    if (cnfStyle && cnfStyle !== "") {
      return false;
    }

    // Check all content items for non-text elements
    const content = para.getContent();
    for (const item of content) {
      // Cast to unknown first for safe duck-typing checks
      const itemAny = item as unknown as Record<string, unknown>;

      // ImageRun check - ImageRun has getImageElement method
      if (item && typeof itemAny.getImageElement === "function") {
        return false;
      }

      // Shape check - Shape has getShapeType method
      if (item && typeof itemAny.getShapeType === "function") {
        return false;
      }

      // TextBox check - TextBox has getTextContent method
      if (item && typeof itemAny.getTextContent === "function") {
        return false;
      }

      // Hyperlink check - Hyperlink has getUrl method
      if (item && typeof itemAny.getUrl === "function") {
        return false;
      }

      // Field check - Field has getInstruction method
      if (item && typeof itemAny.getInstruction === "function") {
        return false;
      }

      // Revision check - check if revision contains meaningful content
      if (item && typeof itemAny.getText === "function") {
        const itemText = (itemAny.getText as () => string)().trim();
        if (itemText !== "") return false;

        // Also check if revision contains non-text elements (hyperlinks, images, shapes, textboxes)
        if (typeof itemAny.getContent === "function") {
          const revContent = (itemAny.getContent as () => unknown[])();
          for (const revItem of revContent) {
            const revItemAny = revItem as Record<string, unknown>;
            // Check if revision content is a Hyperlink using duck typing (getUrl method)
            if (revItem && typeof revItemAny.getUrl === "function") {
              return false; // Revision contains hyperlink - not blank
            }
            // Check if revision content is an ImageRun (getImageElement method)
            if (revItem && typeof revItemAny.getImageElement === "function") {
              return false; // Revision contains image - not blank
            }
            // Check if revision content is a Shape (getShapeType method)
            if (revItem && typeof revItemAny.getShapeType === "function") {
              return false; // Revision contains shape - not blank
            }
            // Check if revision content is a TextBox (getTextContent method)
            if (revItem && typeof revItemAny.getTextContent === "function") {
              return false; // Revision contains textbox - not blank
            }
          }
        }
      }
    }

    // Check for bookmarks (they count as content)
    if (
      para.getBookmarksStart().length > 0 ||
      para.getBookmarksEnd().length > 0
    ) {
      return false;
    }

    // Check for comments (start/end markers)
    if (typeof para.getCommentsStart === "function") {
      const commentsStart = para.getCommentsStart();
      if (commentsStart && commentsStart.length > 0) {
        return false;
      }
    }
    if (typeof para.getCommentsEnd === "function") {
      const commentsEnd = para.getCommentsEnd();
      if (commentsEnd && commentsEnd.length > 0) {
        return false;
      }
    }

    return true;
  }

  /**
   * Sets the parent row reference for this cell.
   * Called by TableRow when adding cells.
   * @internal
   */
  _setParentRow(row: import('./TableRow').TableRow | undefined): void {
    this._parentRow = row;
  }

  /**
   * Gets the parent row reference for this cell.
   * @internal
   */
  _getParentRow(): import('./TableRow').TableRow | undefined {
    return this._parentRow;
  }

  /**
   * Gets the table style ID by traversing up the parent chain.
   * @returns Table style ID or undefined if not in a table or no style set
   */
  getTableStyleId(): string | undefined {
    const row = this._parentRow;
    if (!row) return undefined;

    const table = row._getParentTable();
    if (!table) return undefined;

    return table.getFormatting().style;
  }

  /**
   * Converts the cell to WordprocessingML XML element
   * @returns XMLElement representing the cell
   */
  toXML(): XMLElement {
    const tcPrChildren: XMLElement[] = [];

    // tcPr children ordered per ECMA-376 CT_TcPr:
    // cnfStyle → tcW → gridSpan → hMerge → vMerge → tcBorders → shd → noWrap → tcMar →
    // textDirection → tcFitText → vAlign → hideMark →
    // cellIns/cellDel/cellMerge → tcPrChange

    // cnfStyle - conditional formatting style (MUST be first child per CT_TcPr)
    if (this.formatting.cnfStyle) {
      tcPrChildren.push(
        XMLBuilder.wSelf("cnfStyle", { "w:val": this.formatting.cnfStyle })
      );
    }

    // tcW - cell width
    if (this.formatting.width !== undefined) {
      const widthAttrs: Record<string, string | number> = {
        "w:w": this.formatting.width,
        "w:type": this.formatting.widthType || "dxa",
      };
      tcPrChildren.push(XMLBuilder.wSelf("tcW", widthAttrs));
    }

    // gridSpan - column span
    if (this.formatting.columnSpan && this.formatting.columnSpan > 1) {
      tcPrChildren.push(
        XMLBuilder.wSelf("gridSpan", { "w:val": this.formatting.columnSpan })
      );
    }

    // hMerge - legacy horizontal merge
    if (this.formatting.hMerge) {
      tcPrChildren.push(XMLBuilder.wSelf("hMerge", { "w:val": this.formatting.hMerge }));
    }

    // vMerge - vertical merge
    // Per OOXML, <w:vMerge/> without val means "continue" (default).
    // Only "restart" needs an explicit w:val attribute.
    if (this.formatting.vMerge) {
      if (this.formatting.vMerge === "restart") {
        tcPrChildren.push(XMLBuilder.wSelf("vMerge", { "w:val": "restart" }));
      } else {
        tcPrChildren.push(XMLBuilder.wSelf("vMerge"));
      }
    }

    // tcBorders - cell borders
    if (this.formatting.borders) {
      const borderElements: XMLElement[] = [];
      const borders = this.formatting.borders;

      // Ordered per ECMA-376 CT_TcBorders: top, left, bottom, right, insideH, insideV, tl2br, tr2bl
      if (borders.top) {
        borderElements.push(XMLBuilder.createBorder("top", borders.top));
      }
      if (borders.left) {
        borderElements.push(XMLBuilder.createBorder("left", borders.left));
      }
      if (borders.bottom) {
        borderElements.push(XMLBuilder.createBorder("bottom", borders.bottom));
      }
      if (borders.right) {
        borderElements.push(XMLBuilder.createBorder("right", borders.right));
      }
      if (borders.tl2br) {
        borderElements.push(XMLBuilder.createBorder("tl2br", borders.tl2br));
      }
      if (borders.tr2bl) {
        borderElements.push(XMLBuilder.createBorder("tr2bl", borders.tr2bl));
      }

      if (borderElements.length > 0) {
        tcPrChildren.push(XMLBuilder.w("tcBorders", undefined, borderElements));
      }
    }

    // shd - shading
    if (this.formatting.shading) {
      const shadingAttrs = buildShadingAttributes(this.formatting.shading);
      if (Object.keys(shadingAttrs).length > 0) {
        tcPrChildren.push(XMLBuilder.wSelf("shd", shadingAttrs));
      }
    }

    // noWrap
    if (this.formatting.noWrap) {
      tcPrChildren.push(XMLBuilder.wSelf("noWrap"));
    }

    // tcMar - cell margins (ordered per CT_TcMar: top, left, bottom, right)
    if (this.formatting.margins) {
      const margins = this.formatting.margins;
      const marginChildren: XMLElement[] = [];

      if (margins.top !== undefined) {
        marginChildren.push(
          XMLBuilder.wSelf("top", {
            "w:w": margins.top.toString(),
            "w:type": "dxa",
          })
        );
      }
      if (margins.left !== undefined) {
        marginChildren.push(
          XMLBuilder.wSelf("left", {
            "w:w": margins.left.toString(),
            "w:type": "dxa",
          })
        );
      }
      if (margins.bottom !== undefined) {
        marginChildren.push(
          XMLBuilder.wSelf("bottom", {
            "w:w": margins.bottom.toString(),
            "w:type": "dxa",
          })
        );
      }
      if (margins.right !== undefined) {
        marginChildren.push(
          XMLBuilder.wSelf("right", {
            "w:w": margins.right.toString(),
            "w:type": "dxa",
          })
        );
      }

      if (marginChildren.length > 0) {
        tcPrChildren.push(XMLBuilder.w("tcMar", undefined, marginChildren));
      }
    }

    // textDirection
    if (this.formatting.textDirection) {
      tcPrChildren.push(
        XMLBuilder.wSelf("textDirection", {
          "w:val": this.formatting.textDirection,
        })
      );
    }

    // tcFitText
    if (this.formatting.fitText) {
      tcPrChildren.push(XMLBuilder.wSelf("tcFitText"));
    }

    // vAlign - vertical alignment
    if (this.formatting.verticalAlignment) {
      tcPrChildren.push(
        XMLBuilder.wSelf("vAlign", {
          "w:val": this.formatting.verticalAlignment,
        })
      );
    }

    // hideMark
    if (this.formatting.hideMark) {
      tcPrChildren.push(XMLBuilder.wSelf("hideMark"));
    }

    // Note: w:headers (cell headers for accessibility) is defined in ECMA-376 Part 1 §17.4.26
    // but is NOT included in the Transitional schema and fails OOXML validation.
    // The property is preserved in memory for reading but not generated in XML.

    // Add cell revision markers (w:cellIns, w:cellDel, w:cellMerge) per ECMA-376 Part 1 §17.13.5.4-5.6
    if (this.cellRevision) {
      const revType = this.cellRevision.getType();
      const attrs: Record<string, string | number> = {
        "w:id": this.cellRevision.getId(),
        "w:author": this.cellRevision.getAuthor(),
        "w:date": formatDateForXml(this.cellRevision.getDate()),
      };

      if (revType === "tableCellInsert") {
        tcPrChildren.push(XMLBuilder.wSelf("cellIns", attrs));
      } else if (revType === "tableCellDelete") {
        tcPrChildren.push(XMLBuilder.wSelf("cellDel", attrs));
      } else if (revType === "tableCellMerge") {
        // Add vMerge and vMergeOrig attributes if present
        // ST_AnnotationVMerge: "cont" | "rest" (not "continue" / "restart")
        const mergeMap: Record<string, string> = { continue: "cont", restart: "rest" };
        const prevProps = this.cellRevision.getPreviousProperties();
        if (prevProps?.vMerge) {
          attrs["w:vMerge"] = mergeMap[prevProps.vMerge] || prevProps.vMerge;
        }
        if (prevProps?.vMergeOrig) {
          attrs["w:vMergeOrig"] = mergeMap[prevProps.vMergeOrig] || prevProps.vMergeOrig;
        }
        tcPrChildren.push(XMLBuilder.wSelf("cellMerge", attrs));
      }
    }

    // Add table cell property change (w:tcPrChange) per ECMA-376 Part 1 §17.13.5.37
    // Must be last child of w:tcPr
    if (this.tcPrChange) {
      const changeAttrs: Record<string, string | number> = {
        "w:id": this.tcPrChange.id,
        "w:author": this.tcPrChange.author,
        "w:date": this.tcPrChange.date,
      };
      const prevTcPrChildren: XMLElement[] = [];
      const prev = this.tcPrChange.previousProperties;
      if (prev) {
        // Ordered per CT_TcPr: tcW → tcBorders → shd → noWrap → tcMar →
        // textDirection → tcFitText → vAlign → hideMark → cnfStyle
        if (prev.width !== undefined) {
          prevTcPrChildren.push(XMLBuilder.wSelf("tcW", {
            "w:w": prev.width,
            "w:type": prev.widthType || "dxa",
          }));
        }
        if (prev.borders) {
          const borderElements: XMLElement[] = [];
          if (prev.borders.top) borderElements.push(XMLBuilder.createBorder("top", prev.borders.top));
          if (prev.borders.left) borderElements.push(XMLBuilder.createBorder("left", prev.borders.left));
          if (prev.borders.bottom) borderElements.push(XMLBuilder.createBorder("bottom", prev.borders.bottom));
          if (prev.borders.right) borderElements.push(XMLBuilder.createBorder("right", prev.borders.right));
          if (borderElements.length > 0) {
            prevTcPrChildren.push(XMLBuilder.w("tcBorders", undefined, borderElements));
          }
        }
        if (prev.shading) {
          const shadingAttrs = buildShadingAttributes(prev.shading);
          if (Object.keys(shadingAttrs).length > 0) {
            prevTcPrChildren.push(XMLBuilder.wSelf("shd", shadingAttrs));
          }
        }
        if (prev.noWrap) {
          prevTcPrChildren.push(XMLBuilder.wSelf("noWrap"));
        }
        if (prev.margins) {
          const marginChildren: XMLElement[] = [];
          if (prev.margins.top !== undefined) {
            marginChildren.push(XMLBuilder.wSelf("top", { "w:w": prev.margins.top.toString(), "w:type": "dxa" }));
          }
          if (prev.margins.left !== undefined) {
            marginChildren.push(XMLBuilder.wSelf("left", { "w:w": prev.margins.left.toString(), "w:type": "dxa" }));
          }
          if (prev.margins.bottom !== undefined) {
            marginChildren.push(XMLBuilder.wSelf("bottom", { "w:w": prev.margins.bottom.toString(), "w:type": "dxa" }));
          }
          if (prev.margins.right !== undefined) {
            marginChildren.push(XMLBuilder.wSelf("right", { "w:w": prev.margins.right.toString(), "w:type": "dxa" }));
          }
          if (marginChildren.length > 0) {
            prevTcPrChildren.push(XMLBuilder.w("tcMar", undefined, marginChildren));
          }
        }
        if (prev.textDirection) {
          prevTcPrChildren.push(XMLBuilder.wSelf("textDirection", { "w:val": prev.textDirection }));
        }
        if (prev.fitText) {
          prevTcPrChildren.push(XMLBuilder.wSelf("tcFitText"));
        }
        if (prev.verticalAlignment) {
          prevTcPrChildren.push(XMLBuilder.wSelf("vAlign", { "w:val": prev.verticalAlignment }));
        }
        if (prev.hideMark) {
          prevTcPrChildren.push(XMLBuilder.wSelf("hideMark"));
        }
        if (prev.cnfStyle) {
          prevTcPrChildren.push(XMLBuilder.wSelf("cnfStyle", { "w:val": prev.cnfStyle }));
        }
      }
      const prevTcPr = prevTcPrChildren.length > 0
        ? XMLBuilder.w("tcPr", undefined, prevTcPrChildren)
        : XMLBuilder.w("tcPr", undefined, []);
      tcPrChildren.push(XMLBuilder.w("tcPrChange", changeAttrs, [prevTcPr]));
    }

    // Build cell element
    const cellChildren: XMLElement[] = [];

    // Add cell properties if there are any
    if (tcPrChildren.length > 0) {
      cellChildren.push(XMLBuilder.w("tcPr", undefined, tcPrChildren));
    }

    // Add paragraphs and raw nested content in correct order
    // Raw nested content (tables, SDTs) are interspersed with paragraphs
    if (this.paragraphs.length > 0 || this.rawNestedContent.length > 0) {
      // Sort raw content by position
      const sortedRaw = [...this.rawNestedContent].sort(
        (a, b) => a.position - b.position
      );
      let rawIndex = 0;

      for (let i = 0; i < this.paragraphs.length; i++) {
        // Insert any raw content that comes before this paragraph (position <= i)
        let rawItem = sortedRaw[rawIndex];
        while (rawIndex < sortedRaw.length && rawItem && rawItem.position <= i) {
          // Use __rawXml element for passthrough (supported by XMLBuilder)
          cellChildren.push({
            name: "__rawXml",
            rawXml: rawItem.xml,
          });
          rawIndex++;
          rawItem = sortedRaw[rawIndex];
        }
        const para = this.paragraphs[i];
        if (para) {
          cellChildren.push(para.toXML());
        }
      }

      // Insert any remaining raw content after all paragraphs
      while (rawIndex < sortedRaw.length) {
        const rawItem = sortedRaw[rawIndex];
        if (rawItem) {
          cellChildren.push({
            name: "__rawXml",
            rawXml: rawItem.xml,
          });
        }
        rawIndex++;
      }

      // If we only have raw content and no paragraphs, we need at least one empty paragraph
      // per ECMA-376 (table cell must contain at least one block-level element)
      if (this.paragraphs.length === 0) {
        cellChildren.push(new Paragraph().toXML());
      }
    } else {
      // Empty cell needs at least one empty paragraph
      cellChildren.push(new Paragraph().toXML());
    }

    return XMLBuilder.w("tc", undefined, cellChildren);
  }

  /**
   * Creates a new TableCell
   * @param formatting - Cell formatting
   * @returns New TableCell instance
   */
  static create(formatting?: CellFormatting): TableCell {
    return new TableCell(formatting);
  }
}
