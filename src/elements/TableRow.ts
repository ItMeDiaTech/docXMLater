/**
 * TableRow - Represents a row in a table
 */

import { TableCell } from './TableCell';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { TableBorder, TableBorders } from './Table';

/**
 * Row justification/alignment options
 */
export type RowJustification = 'left' | 'center' | 'right' | 'start' | 'end';

/**
 * Shading pattern values per ECMA-376
 */
export type ShadingPattern = 'clear' | 'solid' | 'horzStripe' | 'vertStripe' | 'reverseDiagStripe' |
  'diagStripe' | 'horzCross' | 'diagCross' | 'thinHorzStripe' | 'thinVertStripe' |
  'thinReverseDiagStripe' | 'thinDiagStripe' | 'thinHorzCross' | 'thinDiagCross';

/**
 * Shading configuration
 */
export interface Shading {
  /** Fill color in hex (e.g., 'FFFF00' for yellow) */
  fill?: string;
  /** Pattern color in hex */
  color?: string;
  /** Shading pattern type */
  pattern?: ShadingPattern;
}

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
}

/**
 * Represents a table row
 */
export class TableRow {
  private cells: TableCell[] = [];
  private formatting: RowFormatting;

  /**
   * Creates a new TableRow
   * @param cellCount - Number of cells to create (optional)
   * @param formatting - Row formatting options
   */
  constructor(cellCount?: number, formatting: RowFormatting = {}) {
    this.formatting = formatting;

    if (cellCount !== undefined && cellCount > 0) {
      for (let i = 0; i < cellCount; i++) {
        this.cells.push(new TableCell());
      }
    }
  }

  /**
   * Adds a cell to the row
   * @param cell - Cell to add
   * @returns This row for chaining
   */
  addCell(cell: TableCell): this {
    this.cells.push(cell);
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
   * Sets row height
   * @param twips - Height in twips
   * @param rule - Height rule
   * @returns This row for chaining
   */
  setHeight(twips: number, rule: RowFormatting['heightRule'] = 'atLeast'): this {
    this.formatting.height = twips;
    this.formatting.heightRule = rule;
    return this;
  }

  /**
   * Sets whether this is a header row
   * @param isHeader - Whether this is a header row
   * @returns This row for chaining
   */
  setHeader(isHeader: boolean = true): this {
    this.formatting.isHeader = isHeader;
    return this;
  }

  /**
   * Sets whether row can split across pages
   * @param cantSplit - Whether to prevent splitting
   * @returns This row for chaining
   */
  setCantSplit(cantSplit: boolean = true): this {
    this.formatting.cantSplit = cantSplit;
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
    this.formatting.justification = alignment;
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
  setHidden(hidden: boolean = true): this {
    this.formatting.hidden = hidden;
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
    this.formatting.gridBefore = columns;
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
    this.formatting.gridAfter = columns;
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
    this.formatting.tablePropertyExceptions = exceptions;
    return this;
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

  /**
   * Builds XML for table property exceptions
   * Per ECMA-376 Part 1 §17.4.61
   * @private
   */
  private buildTablePropertyExceptionsXML(exceptions: TablePropertyExceptions): XMLElement[] {
    const children: XMLElement[] = [];

    // Add table width exception (w:tblW)
    if (exceptions.width !== undefined) {
      children.push(XMLBuilder.wSelf('tblW', {
        'w:w': exceptions.width,
        'w:type': 'dxa'
      }));
    }

    // Add table justification exception (w:jc)
    if (exceptions.justification) {
      children.push(XMLBuilder.wSelf('jc', { 'w:val': exceptions.justification }));
    }

    // Add cell spacing exception (w:tblCellSpacing)
    if (exceptions.cellSpacing !== undefined) {
      children.push(XMLBuilder.wSelf('tblCellSpacing', {
        'w:w': exceptions.cellSpacing,
        'w:type': 'dxa'
      }));
    }

    // Add table indentation exception (w:tblInd)
    if (exceptions.indentation !== undefined) {
      children.push(XMLBuilder.wSelf('tblInd', {
        'w:w': exceptions.indentation,
        'w:type': 'dxa'
      }));
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
      const attrs: Record<string, string> = {};
      if (exceptions.shading.fill) attrs['w:fill'] = exceptions.shading.fill;
      if (exceptions.shading.color) attrs['w:color'] = exceptions.shading.color;
      if (exceptions.shading.pattern) attrs['w:val'] = exceptions.shading.pattern;

      if (Object.keys(attrs).length > 0) {
        children.push(XMLBuilder.wSelf('shd', attrs));
      }
    }

    return children;
  }

  /**
   * Builds XML for table borders
   * @private
   */
  private buildBordersXML(borders: TableBorders): XMLElement[] {
    const children: XMLElement[] = [];

    const borderNames: Array<keyof TableBorders> = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];

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

    // Add row height
    if (this.formatting.height !== undefined) {
      const attrs: Record<string, string | number> = {
        'w:val': this.formatting.height,
      };

      if (this.formatting.heightRule) {
        attrs['w:hRule'] = this.formatting.heightRule;
      }

      trPrChildren.push(XMLBuilder.wSelf('trHeight', attrs));
    }

    // Add header row flag
    if (this.formatting.isHeader) {
      trPrChildren.push(XMLBuilder.wSelf('tblHeader'));
    }

    // Add can't split flag
    if (this.formatting.cantSplit) {
      trPrChildren.push(XMLBuilder.wSelf('cantSplit'));
    }

    // Add row justification
    if (this.formatting.justification) {
      trPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': this.formatting.justification }));
    }

    // Add hidden flag
    if (this.formatting.hidden) {
      trPrChildren.push(XMLBuilder.wSelf('hidden'));
    }

    // Add grid before
    if (this.formatting.gridBefore !== undefined) {
      trPrChildren.push(XMLBuilder.wSelf('gridBefore', { 'w:val': this.formatting.gridBefore }));
    }

    // Add grid after
    if (this.formatting.gridAfter !== undefined) {
      trPrChildren.push(XMLBuilder.wSelf('gridAfter', { 'w:val': this.formatting.gridAfter }));
    }

    // Build row element
    const rowChildren: XMLElement[] = [];

    // Add row properties if there are any
    if (trPrChildren.length > 0) {
      rowChildren.push(XMLBuilder.w('trPr', undefined, trPrChildren));
    }

    // Add table property exceptions (tblPrEx) if present
    if (this.formatting.tablePropertyExceptions) {
      const tblPrExChildren = this.buildTablePropertyExceptionsXML(this.formatting.tablePropertyExceptions);
      if (tblPrExChildren.length > 0) {
        rowChildren.push(XMLBuilder.w('tblPrEx', undefined, tblPrExChildren));
      }
    }

    // Add cells, skipping those covered by gridSpan (column merging)
    let skipCount = 0;
    for (let i = 0; i < this.cells.length; i++) {
      if (skipCount > 0) {
        // This cell is covered by a previous cell's gridSpan, skip it
        skipCount--;
        continue;
      }

      const cell = this.cells[i];
      rowChildren.push(cell!.toXML());

      // If this cell has a gridSpan, skip the next (span - 1) cells
      const formatting = cell!.getFormatting();
      if (formatting.columnSpan && formatting.columnSpan > 1) {
        skipCount = formatting.columnSpan - 1;
      }
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
