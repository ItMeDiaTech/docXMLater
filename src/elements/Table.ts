/**
 * Table - Represents a table in a document
 */

import { TableRow, RowFormatting } from './TableRow';
import { TableCell } from './TableCell';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';

/**
 * Table alignment
 */
export type TableAlignment = 'left' | 'center' | 'right';

/**
 * Table layout type
 */
export type TableLayout = 'auto' | 'fixed';

/**
 * Table border definition (same as cell borders)
 */
export interface TableBorder {
  style?: 'none' | 'single' | 'double' | 'dashed' | 'dotted' | 'thick';
  size?: number;
  color?: string;
}

/**
 * Table borders
 */
export interface TableBorders {
  top?: TableBorder;
  bottom?: TableBorder;
  left?: TableBorder;
  right?: TableBorder;
  insideH?: TableBorder; // Inside horizontal borders
  insideV?: TableBorder; // Inside vertical borders
}

/**
 * Table formatting options
 */
export interface TableFormatting {
  width?: number; // Table width in twips
  alignment?: TableAlignment;
  layout?: TableLayout;
  borders?: TableBorders;
  cellSpacing?: number; // Cell spacing in twips
  indent?: number; // Left indent in twips
}

/**
 * Represents a table
 */
export class Table {
  private rows: TableRow[] = [];
  private formatting: TableFormatting;

  /**
   * Creates a new Table
   * @param rows - Number of rows to create (optional)
   * @param columns - Number of columns per row (optional)
   * @param formatting - Table formatting options
   */
  constructor(rows?: number, columns?: number, formatting: TableFormatting = {}) {
    this.formatting = formatting;

    if (rows !== undefined && rows > 0 && columns !== undefined && columns > 0) {
      for (let i = 0; i < rows; i++) {
        this.rows.push(new TableRow(columns));
      }
    }
  }

  /**
   * Adds a row to the table
   * @param row - Row to add
   * @returns This table for chaining
   */
  addRow(row: TableRow): this {
    this.rows.push(row);
    return this;
  }

  /**
   * Creates and adds a new row
   * @param cellCount - Number of cells in the row
   * @param formatting - Row formatting
   * @returns The created row
   */
  createRow(cellCount?: number, formatting?: RowFormatting): TableRow {
    const row = new TableRow(cellCount, formatting);
    this.rows.push(row);
    return row;
  }

  /**
   * Gets a row by index
   * @param index - Row index (0-based)
   * @returns The row at the index, or undefined
   */
  getRow(index: number): TableRow | undefined {
    return this.rows[index];
  }

  /**
   * Gets all rows in the table
   * @returns Array of rows
   */
  getRows(): TableRow[] {
    return [...this.rows];
  }

  /**
   * Gets the number of rows
   * @returns Number of rows
   */
  getRowCount(): number {
    return this.rows.length;
  }

  /**
   * Gets a cell by row and column index
   * @param rowIndex - Row index (0-based)
   * @param columnIndex - Column index (0-based)
   * @returns The cell, or undefined
   */
  getCell(rowIndex: number, columnIndex: number): TableCell | undefined {
    const row = this.getRow(rowIndex);
    return row ? row.getCell(columnIndex) : undefined;
  }

  /**
   * Sets table width
   * @param twips - Width in twips
   * @returns This table for chaining
   */
  setWidth(twips: number): this {
    this.formatting.width = twips;
    return this;
  }

  /**
   * Sets table alignment
   * @param alignment - Table alignment
   * @returns This table for chaining
   */
  setAlignment(alignment: TableAlignment): this {
    this.formatting.alignment = alignment;
    return this;
  }

  /**
   * Sets table layout
   * @param layout - Layout type
   * @returns This table for chaining
   */
  setLayout(layout: TableLayout): this {
    this.formatting.layout = layout;
    return this;
  }

  /**
   * Sets table borders
   * @param borders - Border definitions
   * @returns This table for chaining
   */
  setBorders(borders: TableBorders): this {
    this.formatting.borders = borders;
    return this;
  }

  /**
   * Sets all borders to the same style (convenience method)
   * @param border - Border definition to apply to all sides
   * @returns This table for chaining
   */
  setAllBorders(border: TableBorder): this {
    this.formatting.borders = {
      top: border,
      bottom: border,
      left: border,
      right: border,
      insideH: border,
      insideV: border,
    };
    return this;
  }

  /**
   * Sets cell spacing
   * @param twips - Cell spacing in twips
   * @returns This table for chaining
   */
  setCellSpacing(twips: number): this {
    this.formatting.cellSpacing = twips;
    return this;
  }

  /**
   * Sets left indent
   * @param twips - Indent in twips
   * @returns This table for chaining
   */
  setIndent(twips: number): this {
    this.formatting.indent = twips;
    return this;
  }

  /**
   * Gets the table formatting
   * @returns Table formatting
   */
  getFormatting(): TableFormatting {
    return { ...this.formatting };
  }

  /**
   * Converts the table to WordprocessingML XML element
   * @returns XMLElement representing the table
   */
  toXML(): XMLElement {
    const tblPrChildren: XMLElement[] = [];

    // Add table width
    if (this.formatting.width !== undefined) {
      tblPrChildren.push(
        XMLBuilder.wSelf('tblW', {
          'w:w': this.formatting.width,
          'w:type': 'dxa',
        })
      );
    }

    // Add table alignment (jc = justification/alignment)
    if (this.formatting.alignment) {
      tblPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': this.formatting.alignment }));
    }

    // Add table layout
    if (this.formatting.layout) {
      tblPrChildren.push(
        XMLBuilder.wSelf('tblLayout', { 'w:type': this.formatting.layout })
      );
    }

    // Add table borders
    if (this.formatting.borders) {
      const borderElements: XMLElement[] = [];
      const borders = this.formatting.borders;

      if (borders.top) {
        borderElements.push(this.createBorderElement('top', borders.top));
      }
      if (borders.bottom) {
        borderElements.push(this.createBorderElement('bottom', borders.bottom));
      }
      if (borders.left) {
        borderElements.push(this.createBorderElement('left', borders.left));
      }
      if (borders.right) {
        borderElements.push(this.createBorderElement('right', borders.right));
      }
      if (borders.insideH) {
        borderElements.push(this.createBorderElement('insideH', borders.insideH));
      }
      if (borders.insideV) {
        borderElements.push(this.createBorderElement('insideV', borders.insideV));
      }

      if (borderElements.length > 0) {
        tblPrChildren.push(XMLBuilder.w('tblBorders', undefined, borderElements));
      }
    }

    // Add cell spacing
    if (this.formatting.cellSpacing !== undefined) {
      tblPrChildren.push(
        XMLBuilder.wSelf('tblCellSpacing', {
          'w:w': this.formatting.cellSpacing,
          'w:type': 'dxa',
        })
      );
    }

    // Add table indent
    if (this.formatting.indent !== undefined) {
      tblPrChildren.push(
        XMLBuilder.wSelf('tblInd', {
          'w:w': this.formatting.indent,
          'w:type': 'dxa',
        })
      );
    }

    // Build table element
    const tableChildren: XMLElement[] = [];

    // Add table properties
    tableChildren.push(XMLBuilder.w('tblPr', undefined, tblPrChildren));

    // Add table grid (column definitions) - simplified for now
    // In a more complete implementation, we'd calculate proper column widths
    const maxColumns = Math.max(...this.rows.map(row => row.getCellCount()), 0);
    if (maxColumns > 0) {
      const tblGridChildren: XMLElement[] = [];
      for (let i = 0; i < maxColumns; i++) {
        tblGridChildren.push(XMLBuilder.wSelf('gridCol', { 'w:w': 2880 })); // Default width
      }
      tableChildren.push(XMLBuilder.w('tblGrid', undefined, tblGridChildren));
    }

    // Add rows
    for (const row of this.rows) {
      tableChildren.push(row.toXML());
    }

    return XMLBuilder.w('tbl', undefined, tableChildren);
  }

  /**
   * Creates a border element
   */
  private createBorderElement(side: string, border: TableBorder): XMLElement {
    const attrs: Record<string, string | number> = {
      'w:val': border.style || 'single',
    };

    if (border.size !== undefined) {
      attrs['w:sz'] = border.size;
    }
    if (border.color) {
      attrs['w:color'] = border.color;
    }

    return XMLBuilder.wSelf(side, attrs);
  }

  /**
   * Creates a new Table
   * @param rows - Number of rows
   * @param columns - Number of columns
   * @param formatting - Table formatting
   * @returns New Table instance
   */
  static create(rows?: number, columns?: number, formatting?: TableFormatting): Table {
    return new Table(rows, columns, formatting);
  }
}
