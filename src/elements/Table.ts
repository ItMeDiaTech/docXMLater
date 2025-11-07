/**
 * Table - Represents a table in a document
 */

import { TableRow, RowFormatting } from './TableRow';
import { TableCell, CellFormatting } from './TableCell';
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
 * Horizontal anchor for table positioning
 */
export type TableHorizontalAnchor = 'text' | 'margin' | 'page';

/**
 * Vertical anchor for table positioning
 */
export type TableVerticalAnchor = 'text' | 'margin' | 'page';

/**
 * Horizontal alignment for relative table positioning
 */
export type TableHorizontalAlignment = 'left' | 'center' | 'right' | 'inside' | 'outside';

/**
 * Vertical alignment for relative table positioning
 */
export type TableVerticalAlignment = 'top' | 'center' | 'bottom' | 'inside' | 'outside';

/**
 * Table positioning properties (for floating tables)
 * Per ECMA-376 Part 1 §17.4.57 (tblpPr)
 */
export interface TablePositionProperties {
  /** Horizontal position in twips (absolute positioning) */
  x?: number;
  /** Vertical position in twips (absolute positioning) */
  y?: number;
  /** Horizontal anchor/positioning base */
  horizontalAnchor?: TableHorizontalAnchor;
  /** Vertical anchor/positioning base */
  verticalAnchor?: TableVerticalAnchor;
  /** Horizontal alignment (relative positioning) */
  horizontalAlignment?: TableHorizontalAlignment;
  /** Vertical alignment (relative positioning) */
  verticalAlignment?: TableVerticalAlignment;
  /** Left padding from anchor in twips */
  leftFromText?: number;
  /** Right padding from anchor in twips */
  rightFromText?: number;
  /** Top padding from anchor in twips */
  topFromText?: number;
  /** Bottom padding from anchor in twips */
  bottomFromText?: number;
}

/**
 * Table width type
 */
export type TableWidthType = 'auto' | 'dxa' | 'pct';

/**
 * Table formatting options
 */
export interface TableFormatting {
  style?: string; // Table style ID (e.g., 'Table1', 'TableGrid')
  width?: number; // Table width in twips
  widthType?: TableWidthType; // Width type (auto, dxa=twips, pct=percentage)
  alignment?: TableAlignment;
  layout?: TableLayout;
  borders?: TableBorders;
  cellSpacing?: number; // Cell spacing in twips
  cellSpacingType?: TableWidthType; // Cell spacing type
  indent?: number; // Left indent in twips
  tblLook?: string; // Table look flags (appearance settings)
  // Batch 1 properties
  position?: TablePositionProperties; // Floating table positioning
  overlap?: boolean; // Allow table overlap with other floating tables
  bidiVisual?: boolean; // Right-to-left table layout
  tableGrid?: number[]; // Column widths in twips
  caption?: string; // Table caption for accessibility
  description?: string; // Table description for accessibility
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
   * Applies conditional formatting to table cells based on rules
   *
   * Enables advanced table formatting including:
   * - Automatic header row detection and styling
   * - Alternating row colors (zebra striping)
   * - Content-based formatting rules
   *
   * @param rules - Conditional formatting rules
   * @returns This table for chaining
   *
   * @example
   * ```typescript
   * // Apply header formatting and alternating rows
   * table.applyConditionalFormatting({
   *   headerRow: true,
   *   alternatingRows: {
   *     even: { shading: { fill: 'F0F0F0' } },
   *     odd: { shading: { fill: 'FFFFFF' } }
   *   }
   * });
   *
   * // Custom header formatting
   * table.applyConditionalFormatting({
   *   headerRow: {
   *     shading: { fill: '4472C4' },
   *     textColor: 'FFFFFF'
   *   }
   * });
   *
   * // Content-based formatting
   * table.applyConditionalFormatting({
   *   contentRules: [
   *     {
   *       condition: (text, row, col) => parseFloat(text) > 1000,
   *       formatting: { shading: { fill: 'FFD700' } } // Gold for large numbers
   *     },
   *     {
   *       condition: (text) => text.toLowerCase().includes('error'),
   *       formatting: { shading: { fill: 'FF0000' } } // Red for errors
   *     }
   *   ]
   * });
   *
   * // Combined rules
   * table.applyConditionalFormatting({
   *   headerRow: { shading: { fill: '2F5496' } },
   *   alternatingRows: {
   *     even: { shading: { fill: 'D9E1F2' } }
   *   },
   *   contentRules: [
   *     {
   *       condition: (text, row, col) => col === 0 && row > 0,
   *       formatting: { textColor: '000000' }
   *     }
   *   ]
   * });
   * ```
   */
  applyConditionalFormatting(rules: {
    headerRow?: boolean | Partial<CellFormatting>;
    alternatingRows?: {
      even?: Partial<CellFormatting>;
      odd?: Partial<CellFormatting>;
    };
    contentRules?: Array<{
      condition: (cellText: string, rowIndex: number, colIndex: number) => boolean;
      formatting: Partial<CellFormatting>;
    }>;
  }): this {
    const rows = this.getRows();

    // Apply header row formatting
    if (rules.headerRow && rows.length > 0) {
      const headerFormatting: Partial<CellFormatting> =
        rules.headerRow === true
          ? { shading: { fill: '4472C4' } } // Default blue header
          : rules.headerRow;

      const headerRow = rows[0];
      if (headerRow) {
        for (const cell of headerRow.getCells()) {
          this.applyCellFormatting(cell, headerFormatting);
        }
      }
    }

    // Apply alternating rows
    if (rules.alternatingRows) {
      rows.forEach((row, index) => {
        const isEven = index % 2 === 0;
        const formatting = isEven
          ? rules.alternatingRows?.even
          : rules.alternatingRows?.odd;

        if (formatting) {
          for (const cell of row.getCells()) {
            this.applyCellFormatting(cell, formatting);
          }
        }
      });
    }

    // Apply content-based rules
    if (rules.contentRules && rules.contentRules.length > 0) {
      rows.forEach((row, rowIndex) => {
        row.getCells().forEach((cell, colIndex) => {
          const cellText = cell
            .getParagraphs()
            .map((p) => p.getText())
            .join('');

          for (const rule of rules.contentRules!) {
            if (rule.condition(cellText, rowIndex, colIndex)) {
              this.applyCellFormatting(cell, rule.formatting);
            }
          }
        });
      });
    }

    return this;
  }

  /**
   * Helper method to apply formatting to a cell
   * @private
   */
  private applyCellFormatting(
    cell: TableCell,
    formatting: Partial<CellFormatting>
  ): void {
    // Apply shading
    if (formatting.shading) {
      cell.setShading(formatting.shading);
    }

    // Apply borders
    if (formatting.borders) {
      cell.setBorders(formatting.borders);
    }

    // Apply text direction
    if (formatting.textDirection) {
      cell.setTextDirection(formatting.textDirection);
    }

    // Apply vertical alignment
    if (formatting.verticalAlignment) {
      cell.setVerticalAlignment(formatting.verticalAlignment);
    }

    // Apply width
    if (formatting.width !== undefined) {
      cell.setWidth(formatting.width);
    }

    // Apply margins
    if (formatting.margins) {
      cell.setMargins(formatting.margins);
    }
  }

  /**
   * Sets table positioning properties for floating tables
   * Per ECMA-376 Part 1 §17.4.57
   * @param position - Table position properties
   * @returns This table for chaining
   * @example
   * ```typescript
   * // Position table at absolute coordinates
   * table.setPosition({
   *   x: 1440, // 1 inch from left
   *   y: 1440, // 1 inch from top
   *   horizontalAnchor: 'page',
   *   verticalAnchor: 'page'
   * });
   *
   * // Position table with relative alignment
   * table.setPosition({
   *   horizontalAlignment: 'center',
   *   verticalAlignment: 'top',
   *   horizontalAnchor: 'margin',
   *   verticalAnchor: 'page'
   * });
   * ```
   */
  setPosition(position: TablePositionProperties): this {
    this.formatting.position = position;
    return this;
  }

  /**
   * Sets whether table can overlap with other floating tables
   * Per ECMA-376 Part 1 §17.4.30
   * @param overlap - True to allow overlap, false to prevent
   * @returns This table for chaining
   */
  setOverlap(overlap: boolean): this {
    this.formatting.overlap = overlap;
    return this;
  }

  /**
   * Sets bidirectional (right-to-left) visual layout
   * Per ECMA-376 Part 1 §17.4.1
   * @param bidi - True for RTL layout, false for LTR
   * @returns This table for chaining
   */
  setBidiVisual(bidi: boolean): this {
    this.formatting.bidiVisual = bidi;
    return this;
  }

  /**
   * Sets table grid column widths
   * Per ECMA-376 Part 1 §17.4.49
   * @param widths - Array of column widths in twips
   * @returns This table for chaining
   * @example
   * ```typescript
   * // 3 columns: 2 inches, 3 inches, 2 inches
   * table.setTableGrid([2880, 4320, 2880]);
   * ```
   */
  setTableGrid(widths: number[]): this {
    this.formatting.tableGrid = widths;
    return this;
  }

  /**
   * Sets table caption for accessibility
   * Per ECMA-376 Part 1 §17.4.58
   * @param caption - Table caption text
   * @returns This table for chaining
   */
  setCaption(caption: string): this {
    this.formatting.caption = caption;
    return this;
  }

  /**
   * Sets table description for accessibility
   * Per ECMA-376 Part 1 §17.4.63
   * @param description - Table description text
   * @returns This table for chaining
   */
  setDescription(description: string): this {
    this.formatting.description = description;
    return this;
  }

  /**
   * Sets table width type
   * Per ECMA-376 Part 1 §17.4.64
   * @param type - Width type ('auto', 'dxa' for twips, 'pct' for percentage)
   * @returns This table for chaining
   */
  setWidthType(type: TableWidthType): this {
    this.formatting.widthType = type;
    return this;
  }

  /**
   * Sets cell spacing type
   * @param type - Cell spacing type
   * @returns This table for chaining
   */
  setCellSpacingType(type: TableWidthType): this {
    this.formatting.cellSpacingType = type;
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

    // Add table positioning properties (tblpPr) - for floating tables
    if (this.formatting.position) {
      const pos = this.formatting.position;
      const posAttrs: Record<string, string | number> = {};

      if (pos.x !== undefined) posAttrs['w:tblpX'] = pos.x;
      if (pos.y !== undefined) posAttrs['w:tblpY'] = pos.y;
      if (pos.horizontalAnchor) posAttrs['w:tblpXSpec'] = pos.horizontalAnchor;
      if (pos.verticalAnchor) posAttrs['w:tblpYSpec'] = pos.verticalAnchor;
      if (pos.horizontalAlignment) posAttrs['w:tblpXAlign'] = pos.horizontalAlignment;
      if (pos.verticalAlignment) posAttrs['w:tblpYAlign'] = pos.verticalAlignment;
      if (pos.leftFromText !== undefined) posAttrs['w:leftFromText'] = pos.leftFromText;
      if (pos.rightFromText !== undefined) posAttrs['w:rightFromText'] = pos.rightFromText;
      if (pos.topFromText !== undefined) posAttrs['w:topFromText'] = pos.topFromText;
      if (pos.bottomFromText !== undefined) posAttrs['w:bottomFromText'] = pos.bottomFromText;

      if (Object.keys(posAttrs).length > 0) {
        tblPrChildren.push(XMLBuilder.wSelf('tblpPr', posAttrs));
      }
    }

    // Add table overlap
    if (this.formatting.overlap !== undefined) {
      tblPrChildren.push(XMLBuilder.wSelf('tblOverlap', {
        'w:val': this.formatting.overlap ? 'overlap' : 'never'
      }));
    }

    // Add bidirectional visual layout
    if (this.formatting.bidiVisual) {
      tblPrChildren.push(XMLBuilder.wSelf('bidiVisual'));
    }

    // Add table width
    if (this.formatting.width !== undefined) {
      const widthType = this.formatting.widthType || 'dxa';
      tblPrChildren.push(
        XMLBuilder.wSelf('tblW', {
          'w:w': this.formatting.width,
          'w:type': widthType,
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
        borderElements.push(XMLBuilder.createBorder('top', borders.top));
      }
      if (borders.bottom) {
        borderElements.push(XMLBuilder.createBorder('bottom', borders.bottom));
      }
      if (borders.left) {
        borderElements.push(XMLBuilder.createBorder('left', borders.left));
      }
      if (borders.right) {
        borderElements.push(XMLBuilder.createBorder('right', borders.right));
      }
      if (borders.insideH) {
        borderElements.push(XMLBuilder.createBorder('insideH', borders.insideH));
      }
      if (borders.insideV) {
        borderElements.push(XMLBuilder.createBorder('insideV', borders.insideV));
      }

      if (borderElements.length > 0) {
        tblPrChildren.push(XMLBuilder.w('tblBorders', undefined, borderElements));
      }
    }

    // Add cell spacing
    if (this.formatting.cellSpacing !== undefined) {
      const cellSpacingType = this.formatting.cellSpacingType || 'dxa';
      tblPrChildren.push(
        XMLBuilder.wSelf('tblCellSpacing', {
          'w:w': this.formatting.cellSpacing,
          'w:type': cellSpacingType,
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

    // Add table caption (accessibility)
    if (this.formatting.caption) {
      tblPrChildren.push(XMLBuilder.wSelf('tblCaption', { 'w:val': this.formatting.caption }));
    }

    // Add table description (accessibility)
    if (this.formatting.description) {
      tblPrChildren.push(XMLBuilder.wSelf('tblDescription', { 'w:val': this.formatting.description }));
    }

    // Build table element
    const tableChildren: XMLElement[] = [];

    // Add table properties
    tableChildren.push(XMLBuilder.w('tblPr', undefined, tblPrChildren));

    // Add table grid (column definitions)
    // Use custom tableGrid if specified, otherwise auto-generate
    const gridWidths = this.formatting.tableGrid;
    const maxColumns = gridWidths ? gridWidths.length : Math.max(...this.rows.map(row => row.getCellCount()), 0);

    if (maxColumns > 0) {
      const tblGridChildren: XMLElement[] = [];

      for (let i = 0; i < maxColumns; i++) {
        if (gridWidths && gridWidths[i] !== undefined) {
          // Use specified grid width
          tblGridChildren.push(XMLBuilder.wSelf('gridCol', { 'w:w': gridWidths[i] }));
        } else {
          // Auto width (default to 2880 twips = 2 inches)
          tblGridChildren.push(XMLBuilder.wSelf('gridCol', { 'w:w': 2880 }));
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

    // Set vertical merge if merging vertically
    if (endRow > startRow) {
      // First cell starts the merge region
      cell.setVerticalMerge('restart');

      // Subsequent cells continue the merge
      for (let row = startRow + 1; row <= endRow; row++) {
        const mergeCell = this.getCell(row, startCol);
        if (mergeCell) {
          mergeCell.setVerticalMerge('continue');
          // If also merging horizontally, set column span on all merged cells
          if (endCol > startCol) {
            mergeCell.setColumnSpan(endCol - startCol + 1);
          }
        }
      }
    }

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
