/**
 * TableRow - Represents a row in a table
 */

import { TableCell } from './TableCell';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { TableBorder, TableBorders } from './Table';
import {
  BasicShadingPattern,
  RowJustification as CommonRowJustification,
  ShadingConfig,
  buildShadingAttributes,
} from './CommonTypes';
import { defaultLogger } from '../utils/logger';

// ============================================================================
// RE-EXPORTED TYPES (for backward compatibility)
// ============================================================================

/**
 * Row justification/alignment options
 * @see CommonTypes.RowJustification
 */
export type RowJustification = CommonRowJustification;

/**
 * Shading pattern values per ECMA-376
 * @see CommonTypes.BasicShadingPattern for the canonical definition
 */
export type ShadingPattern = BasicShadingPattern;

/**
 * Shading configuration
 * @see ShadingConfig in CommonTypes.ts for the canonical definition
 */
export type Shading = ShadingConfig;

/**
 * Table property exceptions - overrides table-level properties for this row
 * Per ECMA-376 Part 1 §17.4.61 (w:tblPrEx)
 */
export interface TablePropertyExceptions {
  /** Border overrides for this row */
  borders?: TableBorders;
  /** Shading override for this row */
  shading?: Shading;
  /** Cell spacing override in twips */
  cellSpacing?: number;
  /** Table width override in twips */
  width?: number;
  /** Table indentation override in twips */
  indentation?: number;
  /** Table justification override */
  justification?: RowJustification;
}

/**
 * Row formatting options
 */
export interface RowFormatting {
  height?: number; // Height in twips
  heightRule?: 'auto' | 'exact' | 'atLeast';
  isHeader?: boolean; // Whether this is a header row
  cantSplit?: boolean; // Prevent row from breaking across pages
  justification?: RowJustification; // Row justification/alignment
  hidden?: boolean; // Hide row
  gridBefore?: number; // Grid columns before first cell
  gridAfter?: number; // Grid columns after last cell
  tablePropertyExceptions?: TablePropertyExceptions; // Table property exceptions for this row
  wBefore?: number; // Width before row in twips (per ECMA-376 §17.4.83)
  wBeforeType?: string; // Width before type (dxa, pct, auto)
  wAfter?: number; // Width after row in twips (per ECMA-376 §17.4.82)
  wAfterType?: string; // Width after type (dxa, pct, auto)
  cellSpacing?: number; // Row-level cell spacing override in twips
  cellSpacingType?: string; // Cell spacing type (dxa, pct)
  cnfStyle?: string; // Conditional formatting bitmask (per ECMA-376 §17.3.1.8)
  divId?: number; // HTML div association (per ECMA-376 §17.4.9)
}

/**
 * Table row property change tracking (w:trPrChange)
 * Per ECMA-376 Part 1 §17.13.5.38
 */
export interface TrPrChange {
  author: string;
  date: string;
  id: string;
  previousProperties: Record<string, any>;
}

/**
 * Represents a table row
 */
export class TableRow {
  private cells: TableCell[] = [];
  private formatting: RowFormatting;
  /** Parent table reference (if row is inside a table) */
  private _parentTable?: import('./Table').Table;
  /** Tracking context for automatic change tracking */
  private trackingContext?: import('../tracking/TrackingContext').TrackingContext;
  /** Table row property change tracking (w:trPrChange) */
  private trPrChange?: TrPrChange;

  /**
   * Creates a new TableRow
   * @param cellCount - Number of cells to create (optional)
   * @param formatting - Row formatting options
   */
  constructor(cellCount?: number, formatting: RowFormatting = {}) {
    this.formatting = formatting;

    if (cellCount !== undefined && cellCount > 0) {
      for (let i = 0; i < cellCount; i++) {
        const cell = new TableCell();
        cell._setParentRow(this);
        this.cells.push(cell);
      }
    }
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
   * Gets the table row property change tracking info
   */
  getTrPrChange(): TrPrChange | undefined {
    return this.trPrChange;
  }

  /**
   * Sets the table row property change tracking info
   */
  setTrPrChange(change: TrPrChange | undefined): void {
    this.trPrChange = change;
  }

  /**
   * Clears the table row property change tracking
   */
  clearTrPrChange(): void {
    this.trPrChange = undefined;
  }

  /**
   * Adds a cell to the row
   * @param cell - Cell to add
   * @returns This row for chaining
   */
  addCell(cell: TableCell): this {
    this.cells.push(cell);
    cell._setParentRow(this);
    return this;
  }

  /**
   * Creates and adds a new cell
   * @param text - Optional text content for the cell
   * @returns The created cell
   */
  createCell(text?: string): TableCell {
    const cell = new TableCell();
    if (text) {
      cell.createParagraph(text);
    }
    this.cells.push(cell);
    cell._setParentRow(this);
    return cell;
  }

  /**
   * Gets a cell by index
   * @param index - Cell index (0-based)
   * @returns The cell at the index, or undefined
   */
  getCell(index: number): TableCell | undefined {
    return this.cells[index];
  }

  /**
   * Inserts a cell at the specified index
   * @param index - Position to insert (0-based)
   * @param cell - Cell to insert
   */
  insertCellAt(index: number, cell: TableCell): void {
    this.cells.splice(index, 0, cell);
    cell._setParentRow(this);
  }

  /**
   * Removes and returns the cell at the specified index
   * @param index - Position to remove (0-based)
   * @returns The removed cell, or undefined if index is out of bounds
   */
  removeCellAt(index: number): TableCell | undefined {
    if (index < 0 || index >= this.cells.length) return undefined;
    const removed = this.cells.splice(index, 1);
    const cell = removed[0];
    if (cell) cell._setParentRow(undefined);
    return cell;
  }

  /**
   * Gets all cells in the row
   * @returns Array of cells
   */
  getCells(): TableCell[] {
    return [...this.cells];
  }

  /**
   * Gets the number of cells in the row
   * @returns Number of cells
   */
  getCellCount(): number {
    return this.cells.length;
  }

  /**
   * Calculates the total grid span of the row (considering column spans)
   *
   * For tables with merged cells, this returns the number of logical columns
   * this row spans based on the columnSpan values of each cell.
   *
   * @returns Total grid span (sum of all cell spans, where unspanned cells count as 1)
   *
   * @example
   * ```typescript
   * // Row with 3 cells, middle one spanning 2 columns
   * const row = new TableRow();
   * row.createCell('A');                              // span = 1
   * row.createCell('B').setColumnSpan(2);             // span = 2
   * row.createCell('C');                              // span = 1
   * row.getTotalGridSpan();                           // Returns 4
   * ```
   */
  getTotalGridSpan(): number {
    let totalSpan = 0;
    for (const cell of this.cells) {
      const formatting = cell.getFormatting();
      totalSpan += formatting.columnSpan || 1;
    }
    return totalSpan;
  }

  /**
   * Validates the row's grid alignment
   *
   * Checks if the total grid span matches the expected column count.
   * Logs a warning if there's a mismatch, which can indicate:
   * - Missing cells (grid span < expected)
   * - Extra cells (grid span > expected)
   * - Incorrect columnSpan values
   *
   * @param expectedColumns - Expected number of columns in the table grid
   * @returns Object with validation result and details
   *
   * @example
   * ```typescript
   * const result = row.validateGridAlignment(4);
   * if (!result.isValid) {
   *   console.log(result.message); // "Row has 3 grid columns but expected 4"
   * }
   * ```
   */
  validateGridAlignment(expectedColumns: number): {
    isValid: boolean;
    actualSpan: number;
    message?: string;
  } {
    const actualSpan = this.getTotalGridSpan();

    if (actualSpan === expectedColumns) {
      return { isValid: true, actualSpan };
    }

    const message =
      `Row grid alignment mismatch: has ${actualSpan} grid columns but expected ${expectedColumns}. ` +
      `Cell count: ${this.cells.length}. ` +
      (actualSpan < expectedColumns
        ? 'Missing cells or incorrect columnSpan values.'
        : 'Extra cells or excessive columnSpan values.');

    defaultLogger.warn(`[TableRow] ${message}`);

    return {
      isValid: false,
      actualSpan,
      message,
    };
  }

  /**
   * Sets row height
   * @param twips - Height in twips
   * @param rule - Height rule
   * @returns This row for chaining
   */
  setHeight(twips: number, rule: RowFormatting['heightRule'] = 'atLeast'): this {
    const prevHeight = this.formatting.height;
    const prevRule = this.formatting.heightRule;
    this.formatting.height = twips;
    this.formatting.heightRule = rule;
    if (this.trackingContext?.isEnabled()) {
      if (prevHeight !== twips) {
        this.trackingContext.trackTableChange(this, 'height', prevHeight, twips);
      }
      if (prevRule !== rule) {
        this.trackingContext.trackTableChange(this, 'heightRule', prevRule, rule);
      }
    }
    return this;
  }

  /**
   * Clears the row height, allowing Word to auto-size the row based on content
   * @returns This row for chaining
   */
  clearHeight(): this {
    delete this.formatting.height;
    delete this.formatting.heightRule;
    return this;
  }

  /**
   * Sets whether this is a header row
   * @param isHeader - Whether this is a header row
   * @returns This row for chaining
   */
  setHeader(isHeader = true): this {
    const prev = this.formatting.isHeader;
    this.formatting.isHeader = isHeader;
    if (this.trackingContext?.isEnabled() && prev !== isHeader) {
      this.trackingContext.trackTableChange(this, 'isHeader', prev, isHeader);
    }
    return this;
  }

  /**
   * Sets whether row can split across pages
   * @param cantSplit - Whether to prevent splitting
   * @returns This row for chaining
   */
  setCantSplit(cantSplit = true): this {
    const prev = this.formatting.cantSplit;
    this.formatting.cantSplit = cantSplit;
    if (this.trackingContext?.isEnabled() && prev !== cantSplit) {
      this.trackingContext.trackTableChange(this, 'cantSplit', prev, cantSplit);
    }
    return this;
  }

  /**
   * Sets row justification/alignment
   * Per ECMA-376 Part 1 §17.4.79 (w:jc)
   * @param alignment - Row justification ('left' | 'center' | 'right' | 'start' | 'end')
   * @returns This row for chaining
   * @example
   * ```typescript
   * row.setJustification('center'); // Center-align the entire row
   * ```
   */
  setJustification(alignment: RowJustification): this {
    const prev = this.formatting.justification;
    this.formatting.justification = alignment;
    if (this.trackingContext?.isEnabled() && prev !== alignment) {
      this.trackingContext.trackTableChange(this, 'justification', prev, alignment);
    }
    return this;
  }

  /**
   * Sets whether row is hidden
   * Per ECMA-376 Part 1 §17.4.23 (w:hidden)
   * @param hidden - Whether to hide the row
   * @returns This row for chaining
   * @example
   * ```typescript
   * row.setHidden(true); // Hide this row from display
   * ```
   */
  setHidden(hidden = true): this {
    const prev = this.formatting.hidden;
    this.formatting.hidden = hidden;
    if (this.trackingContext?.isEnabled() && prev !== hidden) {
      this.trackingContext.trackTableChange(this, 'hidden', prev, hidden);
    }
    return this;
  }

  /**
   * Sets grid columns before first cell
   * Per ECMA-376 Part 1 §17.4.15 (w:gridBefore)
   * Specifies number of grid columns that must be skipped before the first cell
   * @param columns - Number of grid columns to skip before first cell
   * @returns This row for chaining
   * @example
   * ```typescript
   * row.setGridBefore(2); // Skip 2 columns before first cell
   * ```
   */
  setGridBefore(columns: number): this {
    const prev = this.formatting.gridBefore;
    this.formatting.gridBefore = columns;
    if (this.trackingContext?.isEnabled() && prev !== columns) {
      this.trackingContext.trackTableChange(this, 'gridBefore', prev, columns);
    }
    return this;
  }

  /**
   * Sets grid columns after last cell
   * Per ECMA-376 Part 1 §17.4.14 (w:gridAfter)
   * Specifies number of grid columns that must be left after the last cell
   * @param columns - Number of grid columns to leave after last cell
   * @returns This row for chaining
   * @example
   * ```typescript
   * row.setGridAfter(1); // Leave 1 column after last cell
   * ```
   */
  setGridAfter(columns: number): this {
    const prev = this.formatting.gridAfter;
    this.formatting.gridAfter = columns;
    if (this.trackingContext?.isEnabled() && prev !== columns) {
      this.trackingContext.trackTableChange(this, 'gridAfter', prev, columns);
    }
    return this;
  }

  /**
   * Sets table property exceptions for this row
   * Per ECMA-376 Part 1 §17.4.61 (w:tblPrEx)
   *
   * Allows this row to override table-level properties like borders, shading, and cell spacing.
   * Typically used when merging tables or preserving formatting from legacy documents.
   *
   * @param exceptions - Table property exceptions configuration
   * @returns This row for chaining
   * @example
   * ```typescript
   * // Override borders for this row
   * row.setTablePropertyExceptions({
   *   borders: {
   *     top: { style: 'single', size: 8, color: 'FF0000' },
   *     bottom: { style: 'single', size: 8, color: 'FF0000' }
   *   },
   *   shading: { fill: 'FFFF00', pattern: 'clear' }
   * });
   * ```
   */
  setTablePropertyExceptions(exceptions: TablePropertyExceptions): this {
    const prev = this.formatting.tablePropertyExceptions;
    this.formatting.tablePropertyExceptions = exceptions;
    if (this.trackingContext?.isEnabled() && prev !== exceptions) {
      this.trackingContext.trackTableChange(this, 'tablePropertyExceptions', prev, exceptions);
    }
    return this;
  }

  /**
   * Sets width before row (w:wBefore) per ECMA-376 Part 1 §17.4.83
   * @param width - Width in twips
   * @param type - Width type (dxa, pct, auto)
   * @returns This row for chaining
   */
  setWBefore(width: number, type = 'dxa'): this {
    const prevWidth = this.formatting.wBefore;
    const prevType = this.formatting.wBeforeType;
    this.formatting.wBefore = width;
    this.formatting.wBeforeType = type;
    if (this.trackingContext?.isEnabled()) {
      if (prevWidth !== width) {
        this.trackingContext.trackTableChange(this, 'wBefore', prevWidth, width);
      }
      if (prevType !== type) {
        this.trackingContext.trackTableChange(this, 'wBeforeType', prevType, type);
      }
    }
    return this;
  }

  /**
   * Sets width after row (w:wAfter) per ECMA-376 Part 1 §17.4.82
   * @param width - Width in twips
   * @param type - Width type (dxa, pct, auto)
   * @returns This row for chaining
   */
  setWAfter(width: number, type = 'dxa'): this {
    const prevWidth = this.formatting.wAfter;
    const prevType = this.formatting.wAfterType;
    this.formatting.wAfter = width;
    this.formatting.wAfterType = type;
    if (this.trackingContext?.isEnabled()) {
      if (prevWidth !== width) {
        this.trackingContext.trackTableChange(this, 'wAfter', prevWidth, width);
      }
      if (prevType !== type) {
        this.trackingContext.trackTableChange(this, 'wAfterType', prevType, type);
      }
    }
    return this;
  }

  /**
   * Sets row-level cell spacing override (w:tblCellSpacing on row)
   * @param spacing - Cell spacing in twips
   * @param type - Spacing type (dxa, pct)
   * @returns This row for chaining
   */
  setRowCellSpacing(spacing: number, type = 'dxa'): this {
    const prevSpacing = this.formatting.cellSpacing;
    const prevType = this.formatting.cellSpacingType;
    this.formatting.cellSpacing = spacing;
    this.formatting.cellSpacingType = type;
    if (this.trackingContext?.isEnabled()) {
      if (prevSpacing !== spacing) {
        this.trackingContext.trackTableChange(this, 'cellSpacing', prevSpacing, spacing);
      }
      if (prevType !== type) {
        this.trackingContext.trackTableChange(this, 'cellSpacingType', prevType, type);
      }
    }
    return this;
  }

  /**
   * Sets conditional formatting bitmask for this row (w:cnfStyle)
   * Per ECMA-376 Part 1 §17.3.1.8
   * @param cnfStyle - Binary string (e.g., '100000000000' for firstRow)
   * @returns This row for chaining
   */
  setCnfStyle(cnfStyle: string): this {
    const prev = this.formatting.cnfStyle;
    this.formatting.cnfStyle = cnfStyle;
    if (this.trackingContext?.isEnabled() && prev !== cnfStyle) {
      this.trackingContext.trackTableChange(this, 'cnfStyle', prev, cnfStyle);
    }
    return this;
  }

  /**
   * Sets the HTML div ID for web round-trip
   * Per ECMA-376 Part 1 §17.4.9
   * @param id - Div ID number
   * @returns This row for chaining
   */
  setDivId(id: number): this {
    this.formatting.divId = id;
    return this;
  }

  /**
   * Gets the HTML div ID
   * @returns Div ID or undefined
   */
  getDivId(): number | undefined {
    return this.formatting.divId;
  }

  /**
   * Gets table property exceptions
   * @returns Table property exceptions or undefined
   */
  getTablePropertyExceptions(): TablePropertyExceptions | undefined {
    return this.formatting.tablePropertyExceptions;
  }

  /**
   * Gets the row formatting
   * @returns Row formatting
   */
  getFormatting(): RowFormatting {
    return { ...this.formatting };
  }

  // ============================================================================
  // Individual Formatting Getters
  // ============================================================================

  /**
   * Gets the row height in twips
   * @returns Height in twips or undefined if not set
   */
  getHeight(): number | undefined {
    return this.formatting.height;
  }

  /**
   * Gets the row height rule
   * @returns Height rule ('auto', 'exact', 'atLeast') or undefined
   */
  getHeightRule(): string | undefined {
    return this.formatting.heightRule;
  }

  /**
   * Checks if this row is marked as a header row
   * @returns True if this is a header row
   */
  getIsHeader(): boolean {
    return this.formatting.isHeader ?? false;
  }

  /**
   * Gets whether the row can split across pages
   * @returns True if row cannot split
   */
  getCantSplit(): boolean {
    return this.formatting.cantSplit ?? false;
  }

  /**
   * Gets the row justification (alignment)
   * @returns Justification ('left', 'center', 'right') or undefined
   */
  getJustification(): string | undefined {
    return this.formatting.justification;
  }

  /**
   * Checks if this row is hidden
   * @returns True if row is hidden
   */
  isHidden(): boolean {
    return this.formatting.hidden ?? false;
  }

  /**
   * Sets the parent table reference for this row.
   * Called by Table when adding rows.
   * @internal
   */
  _setParentTable(table: import('./Table').Table | undefined): void {
    this._parentTable = table;
  }

  /**
   * Gets the parent table reference for this row.
   * @internal
   */
  _getParentTable(): import('./Table').Table | undefined {
    return this._parentTable;
  }

  /**
   * Builds XML for table property exceptions
   * Per ECMA-376 Part 1 §17.4.61
   * @private
   */
  private buildTablePropertyExceptionsXML(exceptions: TablePropertyExceptions): XMLElement[] {
    const children: XMLElement[] = [];

    // Add table width exception (w:tblW)
    if (exceptions.width !== undefined) {
      children.push(
        XMLBuilder.wSelf('tblW', {
          'w:w': exceptions.width,
          'w:type': 'dxa',
        })
      );
    }

    // Add table justification exception (w:jc)
    if (exceptions.justification) {
      children.push(XMLBuilder.wSelf('jc', { 'w:val': exceptions.justification }));
    }

    // Add cell spacing exception (w:tblCellSpacing)
    if (exceptions.cellSpacing !== undefined) {
      children.push(
        XMLBuilder.wSelf('tblCellSpacing', {
          'w:w': exceptions.cellSpacing,
          'w:type': 'dxa',
        })
      );
    }

    // Add table indentation exception (w:tblInd)
    if (exceptions.indentation !== undefined) {
      children.push(
        XMLBuilder.wSelf('tblInd', {
          'w:w': exceptions.indentation,
          'w:type': 'dxa',
        })
      );
    }

    // Add table borders exception (w:tblBorders)
    if (exceptions.borders) {
      const borderChildren = this.buildBordersXML(exceptions.borders);
      if (borderChildren.length > 0) {
        children.push(XMLBuilder.w('tblBorders', undefined, borderChildren));
      }
    }

    // Add shading exception (w:shd)
    if (exceptions.shading) {
      const shdAttrs: Record<string, string> = {};
      if (exceptions.shading.pattern) shdAttrs['w:val'] = exceptions.shading.pattern;
      if (exceptions.shading.color) shdAttrs['w:color'] = exceptions.shading.color;
      if (exceptions.shading.fill) shdAttrs['w:fill'] = exceptions.shading.fill;
      if (exceptions.shading.themeColor) shdAttrs['w:themeColor'] = exceptions.shading.themeColor;
      if (exceptions.shading.themeFill) shdAttrs['w:themeFill'] = exceptions.shading.themeFill;
      if (exceptions.shading.themeFillShade)
        shdAttrs['w:themeFillShade'] = exceptions.shading.themeFillShade;
      if (exceptions.shading.themeFillTint)
        shdAttrs['w:themeFillTint'] = exceptions.shading.themeFillTint;
      if (exceptions.shading.themeShade) shdAttrs['w:themeShade'] = exceptions.shading.themeShade;
      if (exceptions.shading.themeTint) shdAttrs['w:themeTint'] = exceptions.shading.themeTint;
      children.push(XMLBuilder.w('shd', shdAttrs));
    }

    return children;
  }

  /**
   * Builds XML for table borders
   * @private
   */
  private buildBordersXML(borders: TableBorders): XMLElement[] {
    const children: XMLElement[] = [];

    const borderNames: (keyof TableBorders)[] = [
      'top',
      'left',
      'bottom',
      'right',
      'insideH',
      'insideV',
    ];

    for (const name of borderNames) {
      const border = borders[name];
      if (border) {
        const attrs: Record<string, string | number> = {};
        if (border.style) attrs['w:val'] = border.style;
        if (border.size !== undefined) attrs['w:sz'] = border.size;
        if (border.space !== undefined) attrs['w:space'] = border.space;
        if (border.color) attrs['w:color'] = border.color;

        if (Object.keys(attrs).length > 0) {
          children.push(XMLBuilder.wSelf(name, attrs));
        }
      }
    }

    return children;
  }

  /**
   * Converts the row to WordprocessingML XML element
   * @returns XMLElement representing the row
   */
  toXML(): XMLElement {
    const trPrChildren: XMLElement[] = [];

    // Ordered per CT_TrPr (ECMA-376 §17.4.79):
    // cnfStyle → divId → gridBefore → gridAfter → wBefore → wAfter →
    // cantSplit → trHeight → tblHeader → tblCellSpacing → jc → hidden

    // 1. cnfStyle (conditional formatting bitmask)
    if (this.formatting.cnfStyle) {
      trPrChildren.push(XMLBuilder.wSelf('cnfStyle', { 'w:val': this.formatting.cnfStyle }));
    }

    // 2. divId
    if (this.formatting.divId !== undefined) {
      trPrChildren.push(XMLBuilder.wSelf('divId', { 'w:val': this.formatting.divId }));
    }

    // 3. gridBefore
    if (this.formatting.gridBefore !== undefined) {
      trPrChildren.push(XMLBuilder.wSelf('gridBefore', { 'w:val': this.formatting.gridBefore }));
    }

    // 4. gridAfter
    if (this.formatting.gridAfter !== undefined) {
      trPrChildren.push(XMLBuilder.wSelf('gridAfter', { 'w:val': this.formatting.gridAfter }));
    }

    // 5. wBefore
    if (this.formatting.wBefore !== undefined) {
      trPrChildren.push(
        XMLBuilder.wSelf('wBefore', {
          'w:w': this.formatting.wBefore,
          'w:type': this.formatting.wBeforeType || 'dxa',
        })
      );
    }

    // 6. wAfter
    if (this.formatting.wAfter !== undefined) {
      trPrChildren.push(
        XMLBuilder.wSelf('wAfter', {
          'w:w': this.formatting.wAfter,
          'w:type': this.formatting.wAfterType || 'dxa',
        })
      );
    }

    // 7. cantSplit
    if (this.formatting.cantSplit) {
      trPrChildren.push(XMLBuilder.wSelf('cantSplit'));
    }

    // 8. trHeight
    if (this.formatting.height !== undefined) {
      const attrs: Record<string, string | number> = {
        'w:val': this.formatting.height,
      };
      if (this.formatting.heightRule) {
        attrs['w:hRule'] = this.formatting.heightRule;
      }
      trPrChildren.push(XMLBuilder.wSelf('trHeight', attrs));
    }

    // 9. tblHeader
    if (this.formatting.isHeader) {
      trPrChildren.push(XMLBuilder.wSelf('tblHeader'));
    }

    // 10. tblCellSpacing
    if (this.formatting.cellSpacing !== undefined) {
      trPrChildren.push(
        XMLBuilder.wSelf('tblCellSpacing', {
          'w:w': this.formatting.cellSpacing,
          'w:type': this.formatting.cellSpacingType || 'dxa',
        })
      );
    }

    // 11. jc (map 'start'/'end' to valid ST_JcTable values)
    if (this.formatting.justification) {
      const jcMap: Record<string, string> = { start: 'left', end: 'right' };
      const jcVal = jcMap[this.formatting.justification] || this.formatting.justification;
      trPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': jcVal }));
    }

    // 12. hidden
    if (this.formatting.hidden) {
      trPrChildren.push(XMLBuilder.wSelf('hidden'));
    }

    // Add table row property change (w:trPrChange) per ECMA-376 Part 1 §17.13.5.38
    // Must be last child of w:trPr
    if (this.trPrChange) {
      const changeAttrs: Record<string, string | number> = {
        'w:id': this.trPrChange.id,
        'w:author': this.trPrChange.author,
        'w:date': this.trPrChange.date,
      };
      const prevTrPrChildren: XMLElement[] = [];
      const prev = this.trPrChange.previousProperties;
      if (prev) {
        // Ordered per CT_TrPr: cnfStyle → gridBefore → gridAfter → wBefore → wAfter →
        // cantSplit → trHeight → tblHeader → tblCellSpacing → jc → hidden
        if (prev.cnfStyle) {
          prevTrPrChildren.push(XMLBuilder.wSelf('cnfStyle', { 'w:val': prev.cnfStyle }));
        }
        if (prev.gridBefore !== undefined) {
          prevTrPrChildren.push(XMLBuilder.wSelf('gridBefore', { 'w:val': prev.gridBefore }));
        }
        if (prev.gridAfter !== undefined) {
          prevTrPrChildren.push(XMLBuilder.wSelf('gridAfter', { 'w:val': prev.gridAfter }));
        }
        if (prev.wBefore !== undefined) {
          prevTrPrChildren.push(
            XMLBuilder.wSelf('wBefore', {
              'w:w': prev.wBefore,
              'w:type': prev.wBeforeType || 'dxa',
            })
          );
        }
        if (prev.wAfter !== undefined) {
          prevTrPrChildren.push(
            XMLBuilder.wSelf('wAfter', {
              'w:w': prev.wAfter,
              'w:type': prev.wAfterType || 'dxa',
            })
          );
        }
        if (prev.cantSplit) {
          prevTrPrChildren.push(XMLBuilder.wSelf('cantSplit'));
        }
        if (prev.height !== undefined) {
          const heightAttrs: Record<string, string | number> = { 'w:val': prev.height };
          if (prev.heightRule) heightAttrs['w:hRule'] = prev.heightRule;
          prevTrPrChildren.push(XMLBuilder.wSelf('trHeight', heightAttrs));
        }
        if (prev.isHeader) {
          prevTrPrChildren.push(XMLBuilder.wSelf('tblHeader'));
        }
        if (prev.cellSpacing !== undefined) {
          prevTrPrChildren.push(
            XMLBuilder.wSelf('tblCellSpacing', {
              'w:w': prev.cellSpacing,
              'w:type': prev.cellSpacingType || 'dxa',
            })
          );
        }
        if (prev.justification) {
          const jcPrevMap: Record<string, string> = { start: 'left', end: 'right' };
          prevTrPrChildren.push(
            XMLBuilder.wSelf('jc', { 'w:val': jcPrevMap[prev.justification] || prev.justification })
          );
        }
        if (prev.hidden) {
          prevTrPrChildren.push(XMLBuilder.wSelf('hidden'));
        }
      }
      const prevTrPr = XMLBuilder.w('trPr', undefined, prevTrPrChildren);
      trPrChildren.push(XMLBuilder.w('trPrChange', changeAttrs, [prevTrPr]));
    }

    // Build row element
    const rowChildren: XMLElement[] = [];

    // Add row properties if there are any
    if (trPrChildren.length > 0) {
      rowChildren.push(XMLBuilder.w('trPr', undefined, trPrChildren));
    }

    // Add table property exceptions (tblPrEx) if present
    if (this.formatting.tablePropertyExceptions) {
      const tblPrExChildren = this.buildTablePropertyExceptionsXML(
        this.formatting.tablePropertyExceptions
      );
      if (tblPrExChildren.length > 0) {
        rowChildren.push(XMLBuilder.w('tblPrEx', undefined, tblPrExChildren));
      }
    }

    // Add all cells - each cell is independent
    // Note: gridSpan (columnSpan) means a single cell spans multiple columns in the grid,
    // it does NOT mean subsequent cells should be skipped. Each cell in the array
    // represents a distinct cell that should be output to the XML.
    for (const cell of this.cells) {
      rowChildren.push(cell.toXML());
    }

    return XMLBuilder.w('tr', undefined, rowChildren);
  }

  /**
   * Creates a new TableRow
   * @param cellCount - Number of cells to create
   * @param formatting - Row formatting
   * @returns New TableRow instance
   */
  static create(cellCount?: number, formatting?: RowFormatting): TableRow {
    return new TableRow(cellCount, formatting);
  }
}
