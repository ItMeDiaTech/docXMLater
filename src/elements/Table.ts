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
  space?: number; // Border spacing (padding) in points
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
  style?: string; // Table style ID (e.g., 'Table1', 'TableGrid')
  width?: number; // Table width in twips
  alignment?: TableAlignment;
  layout?: TableLayout;
  borders?: TableBorders;
  cellSpacing?: number; // Cell spacing in twips
  indent?: number; // Left indent in twips
  tblLook?: string; // Table look flags (appearance settings)
}

/**
 * First row formatting options for table headers
 */
export interface FirstRowFormattingOptions {
  /** Text alignment in cells */
  alignment?: 'left' | 'center' | 'right';
  /** Bold text */
  bold?: boolean;
  /** Italic text */
  italic?: boolean;
  /** Underline text */
  underline?: boolean | 'single' | 'double' | 'thick' | 'dotted' | 'dash';
  /** Spacing before paragraph (in twips) */
  spacingBefore?: number;
  /** Spacing after paragraph (in twips) */
  spacingAfter?: number;
  /** Background color (hex without #) */
  shading?: string;
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
    // Set default width if not specified
    // Per ECMA-376, tables require <w:tblW> element for Word compatibility
    // Default: Letter page width (12240 twips) minus standard margins (2*1440 twips) = 9360 twips
    if (formatting.width === undefined) {
      formatting.width = 9360; // ~6.5 inches
    }

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
   * Sets shading (background color) for the first row of the table
   * Useful for header rows in data tables
   * @param color - Hex color without # (e.g., 'DFDFDF' for light gray)
   * @returns This table for chaining
   * @example
   * ```typescript
   * table.setFirstRowShading('DFDFDF'); // Light gray header
   * ```
   */
  setFirstRowShading(color: string): this {
    if (this.rows.length > 0) {
      const firstRow = this.rows[0];
      if (firstRow) {
        for (const cell of firstRow.getCells()) {
          cell.setShading({ fill: color });
        }
      }
    }
    return this;
  }

  /**
   * Sets comprehensive formatting for the first row of the table
   *
   * This is ideal for formatting table headers with alignment, text formatting,
   * spacing, and background color in a single call.
   *
   * @param options - Formatting options for the first row
   * @returns This table for chaining
   * @example
   * ```typescript
   * // Create a formatted table header
   * table.setFirstRowFormatting({
   *   alignment: 'center',
   *   bold: true,
   *   spacingBefore: 120,
   *   spacingAfter: 120,
   *   shading: 'DFDFDF'
   * });
   *
   * // Format header with underline and custom spacing
   * table.setFirstRowFormatting({
   *   alignment: 'left',
   *   bold: true,
   *   underline: 'single',
   *   spacingAfter: 80
   * });
   * ```
   */
  setFirstRowFormatting(options: FirstRowFormattingOptions): this {
    if (this.rows.length === 0) {
      return this; // No rows to format
    }

    const firstRow = this.rows[0];
    if (!firstRow) {
      return this;
    }

    // Apply formatting to each cell in the first row
    for (const cell of firstRow.getCells()) {
      const paragraphs = cell.getParagraphs();

      // Apply shading to the cell if specified
      if (options.shading) {
        cell.setShading({ fill: options.shading });
      }

      // Apply formatting to each paragraph in the cell
      for (const para of paragraphs) {
        // Apply paragraph alignment
        if (options.alignment) {
          para.setAlignment(options.alignment as any);
        }

        // Apply spacing
        if (options.spacingBefore !== undefined) {
          para.setSpaceBefore(options.spacingBefore);
        }
        if (options.spacingAfter !== undefined) {
          para.setSpaceAfter(options.spacingAfter);
        }

        // Apply run formatting to all runs in the paragraph
        const runs = para.getRuns();
        for (const run of runs) {
          // Apply formatting using run's setter methods
          if (options.bold !== undefined) {
            run.setBold(options.bold);
          }
          if (options.italic !== undefined) {
            run.setItalic(options.italic);
          }
          if (options.underline !== undefined) {
            run.setUnderline(options.underline);
          }
        }
      }
    }

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
   * Sets table style reference
   * @param style - Table style ID (e.g., 'Table1', 'TableGrid')
   * @returns This table for chaining
   */
  setStyle(style: string): this {
    this.formatting.style = style;
    return this;
  }

  /**
   * Sets table look flags (appearance settings)
   * @param tblLook - Table look value (e.g., '0000', '04A0')
   * @returns This table for chaining
   */
  setTblLook(tblLook: string): this {
    this.formatting.tblLook = tblLook;
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

    // Add table style (must come first per ECMA-376)
    if (this.formatting.style) {
      tblPrChildren.push(XMLBuilder.wSelf('tblStyle', { 'w:val': this.formatting.style }));
    }

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

    // Add table look (appearance flags)
    if (this.formatting.tblLook) {
      tblPrChildren.push(XMLBuilder.wSelf('tblLook', { 'w:val': this.formatting.tblLook }));
    }

    // Build table element
    const tableChildren: XMLElement[] = [];

    // Add table properties
    tableChildren.push(XMLBuilder.w('tblPr', undefined, tblPrChildren));

    // Add table grid (column definitions)
    const maxColumns = Math.max(...this.rows.map(row => row.getCellCount()), 0);
    if (maxColumns > 0) {
      const tblGridChildren: XMLElement[] = [];
      const columnWidths = (this.formatting as any).columnWidths;

      for (let i = 0; i < maxColumns; i++) {
        // Use specified width if available, otherwise default
        const width = columnWidths && columnWidths[i] !== null ? columnWidths[i] : 2880;
        if (width !== null) {
          tblGridChildren.push(XMLBuilder.wSelf('gridCol', { 'w:w': width }));
        } else {
          // Auto width (no w:w attribute)
          tblGridChildren.push(XMLBuilder.wSelf('gridCol', {}));
        }
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

    if (border.color) {
      attrs['w:color'] = border.color;
    }
    if (border.space !== undefined) {
      attrs['w:space'] = border.space;
    }
    if (border.size !== undefined) {
      attrs['w:sz'] = border.size;
    }

    return XMLBuilder.wSelf(side, attrs);
  }

  /**
   * Removes a row from the table
   * @param index - Row index to remove (0-based)
   * @returns True if the row was removed, false if index was invalid
   */
  removeRow(index: number): boolean {
    if (index >= 0 && index < this.rows.length) {
      this.rows.splice(index, 1);
      return true;
    }
    return false;
  }

  /**
   * Inserts a row at the specified position
   * @param index - Position to insert at (0-based)
   * @param row - Row to insert (optional, creates empty row if not provided)
   * @returns The inserted row
   */
  insertRow(index: number, row?: TableRow): TableRow {
    // Clamp index to valid range
    if (index < 0) index = 0;
    if (index > this.rows.length) index = this.rows.length;

    // Create new row if not provided, matching the column count
    if (!row) {
      const columnCount = this.getColumnCount();
      row = new TableRow(columnCount);
    }

    // Insert the row
    this.rows.splice(index, 0, row);
    return row;
  }

  /**
   * Adds a column to all rows in the table
   * @param index - Optional position to insert column (defaults to end)
   * @returns This table for chaining
   */
  addColumn(index?: number): this {
    for (const row of this.rows) {
      const cell = new TableCell();
      const cells = row.getCells();

      if (index === undefined || index >= cells.length) {
        // Add to end
        row.addCell(cell);
      } else {
        // Insert at specific position
        const idx = Math.max(0, index);
        // We need to rebuild the row with cells in the correct order
        const newCells = [...cells.slice(0, idx), cell, ...cells.slice(idx)];

        // Clear existing cells and add in new order
        (row as any).cells = newCells;
      }
    }
    return this;
  }

  /**
   * Removes a column from all rows in the table
   * @param index - Column index to remove (0-based)
   * @returns True if the column was removed, false if index was invalid
   */
  removeColumn(index: number): boolean {
    if (index < 0 || this.rows.length === 0) {
      return false;
    }

    let removed = false;
    for (const row of this.rows) {
      const cells = row.getCells();
      if (index < cells.length) {
        // Remove the cell at the specified index
        (row as any).cells.splice(index, 1);
        removed = true;
      }
    }

    return removed;
  }

  /**
   * Gets the maximum column count across all rows
   * @returns Maximum number of columns
   */
  getColumnCount(): number {
    if (this.rows.length === 0) {
      return 0;
    }
    return Math.max(...this.rows.map(row => row.getCellCount()));
  }

  /**
   * Sets column widths for the table
   * @param widths - Array of widths in twips (null for auto width)
   * @returns This table for chaining
   */
  setColumnWidths(widths: (number | null)[]): this {
    // Store column widths in formatting for use in toXML
    (this.formatting as any).columnWidths = widths;
    return this;
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

  /**
   * Merges cells into a single cell (uses columnSpan and rowSpan)
   * @param startRow - Starting row index (0-based)
   * @param startCol - Starting column index (0-based)
   * @param endRow - Ending row index (0-based, inclusive)
   * @param endCol - Ending column index (0-based, inclusive)
   * @returns This table for chaining
   * @example
   * ```typescript
   * table.mergeCells(0, 0, 0, 2);  // Merge cells across columns in first row
   * table.mergeCells(0, 0, 2, 0);  // Merge cells down rows in first column
   * ```
   */
  mergeCells(startRow: number, startCol: number, endRow: number, endCol: number): this {
    if (startRow < 0 || endRow >= this.rows.length || startCol < 0 || endCol < 0) {
      return this;
    }

    const cell = this.getCell(startRow, startCol);
    if (!cell) {
      return this;
    }

    // Set column span if merging horizontally
    if (endCol > startCol) {
      cell.setColumnSpan(endCol - startCol + 1);
    }

    // Note: Row spanning is more complex in Word and requires gridSpan and vMerge
    // For now, we only handle column spanning properly
    // TODO: Implement full row spanning with vMerge in future

    return this;
  }

  /**
   * Splits a cell (removes column/row span)
   * @param row - Row index (0-based)
   * @param col - Column index (0-based)
   * @returns This table for chaining
   * @example
   * ```typescript
   * table.splitCell(0, 0);  // Remove any spanning from cell
   * ```
   */
  splitCell(row: number, col: number): this {
    const cell = this.getCell(row, col);
    if (cell) {
      cell.setColumnSpan(1);  // Reset to single cell
    }
    return this;
  }

  /**
   * Moves cell contents from one position to another
   * @param fromRow - Source row index
   * @param fromCol - Source column index
   * @param toRow - Target row index
   * @param toCol - Target column index
   * @returns This table for chaining
   * @example
   * ```typescript
   * table.moveCell(0, 0, 1, 1);  // Move cell from [0,0] to [1,1]
   * ```
   */
  moveCell(fromRow: number, fromCol: number, toRow: number, toCol: number): this {
    const fromCell = this.getCell(fromRow, fromCol);
    const toCell = this.getCell(toRow, toCol);

    if (!fromCell || !toCell) {
      return this;
    }

    // Copy all paragraphs from source to target
    const paragraphs = fromCell.getParagraphs();
    for (const para of paragraphs) {
      toCell.addParagraph(para);
    }

    // Copy formatting
    const formatting = fromCell.getFormatting();
    if (formatting.shading) toCell.setShading(formatting.shading);
    if (formatting.borders) toCell.setBorders(formatting.borders);
    if (formatting.width) toCell.setWidth(formatting.width);
    if (formatting.verticalAlignment) toCell.setVerticalAlignment(formatting.verticalAlignment);

    // Clear source cell (replace with empty paragraph)
    const row = this.getRow(fromRow);
    if (row) {
      const cells = row.getCells();
      cells[fromCol] = new TableCell();
    }

    return this;
  }

  /**
   * Swaps contents of two cells
   * @param row1 - First cell row index
   * @param col1 - First cell column index
   * @param row2 - Second cell row index
   * @param col2 - Second cell column index
   * @returns This table for chaining
   * @example
   * ```typescript
   * table.swapCells(0, 0, 1, 1);  // Swap cells at [0,0] and [1,1]
   * ```
   */
  swapCells(row1: number, col1: number, row2: number, col2: number): this {
    const row1Obj = this.getRow(row1);
    const row2Obj = this.getRow(row2);

    if (!row1Obj || !row2Obj) {
      return this;
    }

    const cells1 = row1Obj.getCells();
    const cells2 = row2Obj.getCells();

    if (col1 >= cells1.length || col2 >= cells2.length) {
      return this;
    }

    // Swap cells
    const temp = cells1[col1];
    cells1[col1] = cells2[col2]!;
    cells2[col2] = temp!;

    return this;
  }

  /**
   * Sets width for a specific column
   * @param columnIndex - Column index (0-based)
   * @param width - Width in twips
   * @returns This table for chaining
   * @example
   * ```typescript
   * table.setColumnWidth(0, 2000);  // Set first column to 2000 twips
   * ```
   */
  setColumnWidth(columnIndex: number, width: number): this {
    const columnWidths = (this.formatting as any).columnWidths || [];
    columnWidths[columnIndex] = width;
    (this.formatting as any).columnWidths = columnWidths;
    return this;
  }

  /**
   * Inserts multiple rows at once
   * @param startIndex - Position to insert at (0-based)
   * @param count - Number of rows to insert
   * @returns Array of inserted rows
   * @example
   * ```typescript
   * const rows = table.insertRows(2, 3);  // Insert 3 rows at position 2
   * ```
   */
  insertRows(startIndex: number, count: number): TableRow[] {
    const insertedRows: TableRow[] = [];
    const columnCount = this.getColumnCount();

    for (let i = 0; i < count; i++) {
      const row = new TableRow(columnCount);
      this.insertRow(startIndex + i, row);
      insertedRows.push(row);
    }

    return insertedRows;
  }

  /**
   * Removes multiple rows at once
   * @param startIndex - Starting position (0-based)
   * @param count - Number of rows to remove
   * @returns True if rows were removed
   * @example
   * ```typescript
   * table.removeRows(2, 3);  // Remove 3 rows starting at position 2
   * ```
   */
  removeRows(startIndex: number, count: number): boolean {
    if (startIndex < 0 || startIndex >= this.rows.length || count <= 0) {
      return false;
    }

    const actualCount = Math.min(count, this.rows.length - startIndex);
    this.rows.splice(startIndex, actualCount);
    return actualCount > 0;
  }

  /**
   * Creates a deep clone of this table
   * @returns New Table instance with copied rows and formatting
   * @example
   * ```typescript
   * const original = new Table(2, 3);
   * const copy = original.clone();
   * ```
   */
  clone(): Table {
    // Clone formatting
    const clonedFormatting: TableFormatting = JSON.parse(JSON.stringify(this.formatting));

    // Create new table with same structure
    const clonedTable = new Table(0, 0, clonedFormatting);

    // Clone all rows
    for (const row of this.rows) {
      // Clone row by creating new cells with same content
      const cells = row.getCells();
      const clonedRow = new TableRow(0);

      for (const cell of cells) {
        const cellFormatting = cell.getFormatting();
        const clonedCell = new TableCell(JSON.parse(JSON.stringify(cellFormatting)));

        // Clone paragraphs in cell
        for (const para of cell.getParagraphs()) {
          clonedCell.addParagraph(para.clone());
        }

        clonedRow.addCell(clonedCell);
      }

      clonedTable.addRow(clonedRow);
    }

    return clonedTable;
  }
}
