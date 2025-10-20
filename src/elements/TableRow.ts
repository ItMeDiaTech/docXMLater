/**
 * TableRow - Represents a row in a table
 */

import { TableCell } from './TableCell';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';

/**
 * Row formatting options
 */
export interface RowFormatting {
  height?: number; // Height in twips
  heightRule?: 'auto' | 'exact' | 'atLeast';
  isHeader?: boolean; // Whether this is a header row
  cantSplit?: boolean; // Prevent row from breaking across pages
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
   * Gets the row formatting
   * @returns Row formatting
   */
  getFormatting(): RowFormatting {
    return { ...this.formatting };
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

    // Build row element
    const rowChildren: XMLElement[] = [];

    // Add row properties if there are any
    if (trPrChildren.length > 0) {
      rowChildren.push(XMLBuilder.w('trPr', undefined, trPrChildren));
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
