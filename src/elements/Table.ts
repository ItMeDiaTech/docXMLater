/**
 * Table - Represents a table in a document
 */

import { Paragraph } from './Paragraph';
import { TableRow, RowFormatting } from './TableRow';
import { TableCell, CellFormatting } from './TableCell';
import { Revision } from './Revision';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { deepClone } from '../utils/deepClone';
import { TableGridChange } from './TableGridChange';
import {
  TableAlignment as CommonTableAlignment,
  FullBorderStyle,
  HorizontalAnchor,
  VerticalAnchor,
  HorizontalAlignment,
  VerticalAlignment,
  ShadingConfig,
  buildShadingAttributes,
} from './CommonTypes';

// ============================================================================
// RE-EXPORTED TYPES (for backward compatibility)
// ============================================================================

/**
 * Table alignment
 * @see CommonTypes.TableAlignment
 */
export type TableAlignment = CommonTableAlignment;

/**
 * Table layout type
 */
export type TableLayout = 'auto' | 'autofit' | 'fixed';

/**
 * Table border definition (same as cell borders)
 * Table border per ECMA-376 §17.4.66 (tblBorders) / §17.4.66 / etc.
 * `style` accepts the full ST_Border enumeration (§17.18.2) — table
 * borders support the 25+ multi-line / dot-dash / inset / outset
 * variants, not just the narrow 6-value BorderStyle subset.
 *
 * @see CommonTypes.BorderDefinition
 */
export interface TableBorder {
  style?: FullBorderStyle;
  size?: number;
  space?: number; // Border spacing (padding) in points
  color?: string;
  /** Theme color reference (ST_ThemeColor per §17.18.97) */
  themeColor?: string;
  /** Theme tint (2-hex-digit string) */
  themeTint?: string;
  /** Theme shade (2-hex-digit string) */
  themeShade?: string;
  /** Border casts a shadow (CT_OnOff attribute on CT_Border §17.18.2) */
  shadow?: boolean;
  /** Border is part of a frame around the content (CT_OnOff) */
  frame?: boolean;
}

/**
 * Table borders
 * @see CommonTypes.TableBorderDefinitions
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
 * @see CommonTypes.HorizontalAnchor
 */
export type TableHorizontalAnchor = HorizontalAnchor;

/**
 * Vertical anchor for table positioning
 * @see CommonTypes.VerticalAnchor
 */
export type TableVerticalAnchor = VerticalAnchor;

/**
 * Horizontal alignment for relative table positioning
 * @see CommonTypes.HorizontalAlignment
 */
export type TableHorizontalAlignment = HorizontalAlignment;

/**
 * Vertical alignment for relative table positioning
 * @see CommonTypes.VerticalAlignment
 */
export type TableVerticalAlignment = VerticalAlignment;

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
export type TableWidthType = 'auto' | 'dxa' | 'nil' | 'pct';

/**
 * Table shading/background
 * @see ShadingConfig in CommonTypes.ts for the canonical definition
 */
export type TableShading = ShadingConfig;

/**
 * Table formatting options
 */
/**
 * Table cell margins (padding inside cells)
 * Per ECMA-376 Part 1 §17.4.42 (tblCellMar)
 */
export interface TableCellMargins {
  /** Top margin in twips */
  top?: number;
  /** Bottom margin in twips */
  bottom?: number;
  /** Left margin in twips */
  left?: number;
  /** Right margin in twips */
  right?: number;
}

/**
 * Builds tblLook attributes including extended attributes (firstRow, lastRow, etc.)
 * Per ECMA-376 Part 1 §17.4.57, Word expects both w:val and the extended boolean attributes.
 *
 * CT_TblLook's `w:val` is `ST_ShortHexNumber` (XSD `hexBinary`, 2 bytes → 4
 * hex digits). Normalise upstream inputs to 4 uppercase hex digits so we
 * never emit invalid forms like `w:val="620"` (pre-existing upstream
 * issue: XMLParser coerces purely-numeric hex strings like "0620" to the
 * number 620, which would otherwise flow through unpadded).
 */
function buildTblLookAttributes(hex: string | number): Record<string, string> {
  const raw = typeof hex === 'number' ? hex.toString(16) : String(hex ?? '');
  const normalized = raw.toUpperCase().padStart(4, '0');
  const value = parseInt(normalized, 16) || 0;
  return {
    'w:val': normalized,
    'w:firstRow': (value & 0x0020) !== 0 ? '1' : '0',
    'w:lastRow': (value & 0x0040) !== 0 ? '1' : '0',
    'w:firstColumn': (value & 0x0080) !== 0 ? '1' : '0',
    'w:lastColumn': (value & 0x0100) !== 0 ? '1' : '0',
    'w:noHBand': (value & 0x0200) !== 0 ? '1' : '0',
    'w:noVBand': (value & 0x0400) !== 0 ? '1' : '0',
  };
}

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
  cellMargins?: TableCellMargins; // Default cell margins (padding) for all cells
  indent?: number; // Left indent in twips
  indentType?: TableWidthType; // Indent type (auto, dxa, nil, pct)
  tblLook?: string; // Table look flags (appearance settings)
  shading?: TableShading; // Table background shading
  // Batch 1 properties
  position?: TablePositionProperties; // Floating table positioning
  overlap?: boolean; // Allow table overlap with other floating tables
  bidiVisual?: boolean; // Right-to-left table layout
  tableGrid?: number[]; // Column widths in twips
  caption?: string; // Table caption for accessibility
  description?: string; // Table description for accessibility
  tblStyleRowBandSize?: number; // Number of rows in each row band for style alternation
  tblStyleColBandSize?: number; // Number of columns in each column band for style alternation
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
/**
 * Table property change tracking (w:tblPrChange)
 * Per ECMA-376 Part 1 §17.13.5.36
 */
export interface TblPrChange {
  author: string;
  date: string;
  id: string;
  previousProperties: Record<string, any>;
}

export class Table {
  private rows: TableRow[] = [];
  private formatting: TableFormatting;
  /** StylesManager reference for conditional formatting resolution */
  private _stylesManager?: import('../formatting/StylesManager').StylesManager;
  /** Tracking context for automatic change tracking */
  private trackingContext?: import('../tracking/TrackingContext').TrackingContext;
  /** Table property change tracking (w:tblPrChange) */
  private tblPrChange?: TblPrChange;
  /** Table grid change tracking (w:tblGridChange) per ECMA-376 §17.13.5.35 */
  private tblGridChange?: TableGridChange;

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
        const row = new TableRow(columns);
        row._setParentTable(this);
        this.rows.push(row);
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
   * Gets the table property change tracking info
   */
  getTblPrChange(): TblPrChange | undefined {
    return this.tblPrChange;
  }

  /**
   * Sets the table property change tracking info
   */
  setTblPrChange(change: TblPrChange | undefined): void {
    this.tblPrChange = change;
  }

  /**
   * Clears the table property change tracking
   */
  clearTblPrChange(): void {
    this.tblPrChange = undefined;
  }

  /**
   * Gets the table grid change tracking info
   */
  getTblGridChange(): TableGridChange | undefined {
    return this.tblGridChange;
  }

  /**
   * Sets the table grid change tracking info
   */
  setTblGridChange(change: TableGridChange | undefined): void {
    this.tblGridChange = change;
  }

  /**
   * Adds a row to the table
   *
   * Appends a TableRow instance to the end of the table.
   *
   * @param row - The TableRow instance to add
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * const table = new Table();
   * const row = new TableRow(3);
   * table.addRow(row);
   * ```
   */
  addRow(row: TableRow): this {
    this.rows.push(row);
    row._setParentTable(this);
    return this;
  }

  /**
   * Adds a row from an array of cell text values
   *
   * Creates a TableRow with cells populated from the string array and appends
   * it to the table. This is the single-row counterpart to `Table.fromArray()`.
   *
   * @param cells - Array of text values, one per cell
   * @returns The created TableRow for further customization
   *
   * @example
   * ```typescript
   * const table = Table.fromArray([['Name', 'Age', 'City']]);
   * table.addRowFromArray(['Alice', '30', 'NYC']);
   * table.addRowFromArray(['Bob', '25', 'London']);
   * ```
   */
  addRowFromArray(cells: string[]): TableRow {
    const row = new TableRow();
    for (const text of cells) {
      row.createCell(text || undefined);
    }
    this.rows.push(row);
    row._setParentTable(this);
    return row;
  }

  /**
   * Adds a summary/totals row computed from column data
   *
   * Appends a new row where each cell's value is computed by applying an
   * aggregation function to that column's data rows (header row excluded).
   * By default, numeric columns are summed and non-numeric columns show a
   * count. The first column shows a label.
   *
   * @param options - Summary configuration
   * @param options.label - Label for the first cell (default: 'Total')
   * @param options.startRow - First data row index to include (default: 1, skipping header)
   * @param options.compute - Custom function to compute a cell value from its column's text values.
   *   Receives `(values: string[], colIndex: number)`. If not provided, defaults to
   *   numeric sum for parseable columns, empty string otherwise.
   * @returns The created summary TableRow
   *
   * @example
   * ```typescript
   * const table = Table.fromArray([
   *   ['Product', 'Price', 'Qty'],
   *   ['Widget', '10', '5'],
   *   ['Gadget', '25', '2'],
   * ]);
   * table.addSummaryRow();
   * // Adds: ['Total', '35', '7']
   * ```
   *
   * @example
   * ```typescript
   * // Custom compute: average instead of sum
   * table.addSummaryRow({
   *   label: 'Average',
   *   compute: (values) => {
   *     const nums = values.map(Number).filter(n => !isNaN(n));
   *     return nums.length ? (nums.reduce((a, b) => a + b, 0) / nums.length).toFixed(1) : '';
   *   },
   * });
   * ```
   */
  addSummaryRow(options?: {
    label?: string;
    startRow?: number;
    compute?: (values: string[], colIndex: number) => string;
  }): TableRow {
    const label = options?.label ?? 'Total';
    const startRow = options?.startRow ?? 1;
    const colCount = this.getColumnCount();

    const defaultCompute = (values: string[]): string => {
      const nums = values.map((v) => parseFloat(v)).filter((n) => !isNaN(n));
      if (nums.length === 0) return '';
      return String(nums.reduce((a, b) => a + b, 0));
    };

    const compute = options?.compute ?? defaultCompute;
    const cells: string[] = [];

    for (let c = 0; c < colCount; c++) {
      if (c === 0) {
        cells.push(label);
        continue;
      }

      const colTexts = this.getColumnTexts(c);
      const dataValues = colTexts.slice(startRow);
      cells.push(compute(dataValues, c));
    }

    return this.addRowFromArray(cells);
  }

  /**
   * Creates a new row and adds it to the table
   *
   * Convenience method that creates a TableRow and appends it in one operation.
   *
   * @param cellCount - Number of cells to create in the row
   * @param formatting - Optional row formatting properties
   * @returns The created TableRow instance for further customization
   *
   * @example
   * ```typescript
   * const table = new Table();
   * const row = table.createRow(4);
   * row.getCell(0)?.addParagraph(new Paragraph().addText('Cell 1'));
   * ```
   */
  createRow(cellCount?: number, formatting?: RowFormatting): TableRow {
    const row = new TableRow(cellCount, formatting);
    this.rows.push(row);
    row._setParentTable(this);
    return row;
  }

  /**
   * Gets a row by its index
   *
   * @param index - The row index (0-based, where 0 is the first row)
   * @returns The TableRow at the specified index, or undefined if index is out of bounds
   *
   * @example
   * ```typescript
   * const firstRow = table.getRow(0);
   * const secondRow = table.getRow(1);
   * if (firstRow) {
   *   console.log(`First row has ${firstRow.getCellCount()} cells`);
   * }
   * ```
   */
  getRow(index: number): TableRow | undefined {
    return this.rows[index];
  }

  /**
   * Gets all rows in the table
   *
   * Returns a copy of the rows array to prevent external modification.
   *
   * @returns Array of all TableRow instances
   *
   * @example
   * ```typescript
   * const rows = table.getRows();
   * console.log(`Table has ${rows.length} rows`);
   * for (const row of rows) {
   *   console.log(`Row has ${row.getCellCount()} cells`);
   * }
   * ```
   */
  getRows(): TableRow[] {
    return [...this.rows];
  }

  /**
   * Gets the first paragraph in the table (first cell of first row)
   * @returns The first paragraph, or null if the table is empty
   */
  getFirstParagraph(): Paragraph | null {
    if (this.rows.length === 0) return null;
    const firstRow = this.rows[0];
    if (!firstRow) return null;
    const cells = firstRow.getCells();
    if (cells.length === 0) return null;
    const paras = cells[0]?.getParagraphs();
    return paras && paras.length > 0 ? (paras[0] ?? null) : null;
  }

  /**
   * Gets the last paragraph in the table (last cell of last row)
   * @returns The last paragraph, or null if the table is empty
   */
  getLastParagraph(): Paragraph | null {
    for (let r = this.rows.length - 1; r >= 0; r--) {
      const row = this.rows[r];
      if (!row) continue;
      const cells = row.getCells();
      for (let c = cells.length - 1; c >= 0; c--) {
        const cell = cells[c];
        if (!cell) continue;
        const paras = cell.getParagraphs();
        if (paras.length > 0) {
          return paras[paras.length - 1] ?? null;
        }
      }
    }
    return null;
  }

  /**
   * Gets the total number of rows in the table
   *
   * @returns Number of rows
   *
   * @example
   * ```typescript
   * console.log(`Table has ${table.getRowCount()} rows`);
   * ```
   */
  getRowCount(): number {
    return this.rows.length;
  }

  /**
   * Gets a specific cell by row and column indices
   *
   * @param rowIndex - The row index (0-based)
   * @param columnIndex - The column index (0-based)
   * @returns The TableCell at the specified position, or undefined if indices are out of bounds
   *
   * @example
   * ```typescript
   * const cell = table.getCell(0, 0); // Top-left cell
   * if (cell) {
   *   cell.addParagraph(new Paragraph().addText('A1'));
   * }
   *
   * // Access cell in third row, second column
   * const cell2 = table.getCell(2, 1);
   * ```
   */
  getCell(rowIndex: number, columnIndex: number): TableCell | undefined {
    const row = this.getRow(rowIndex);
    return row ? row.getCell(columnIndex) : undefined;
  }

  /**
   * Sets the text content of a cell, replacing any existing content
   *
   * Convenience method that locates the cell at (row, col) and sets its
   * text. If the cell has existing paragraphs, the first paragraph's text
   * is replaced and extra paragraphs are removed. If the cell is empty,
   * a new paragraph is created.
   *
   * @param rowIndex - Row index (0-based)
   * @param colIndex - Column index (0-based)
   * @param text - New text content for the cell
   * @returns This table for chaining, or this table (no-op if cell not found)
   *
   * @example
   * ```typescript
   * const table = new Table(3, 3);
   * table.setCell(0, 0, 'Name');
   * table.setCell(0, 1, 'Age');
   * table.setCell(1, 0, 'Alice');
   * table.setCell(1, 1, '30');
   * ```
   */
  setCell(rowIndex: number, colIndex: number, text: string): this {
    const cell = this.getCell(rowIndex, colIndex);
    if (!cell) return this;

    const paragraphs = cell.getParagraphs();
    if (paragraphs.length > 0) {
      paragraphs[0]!.setText(text);
      // Remove extra paragraphs
      for (let p = paragraphs.length - 1; p >= 1; p--) {
        cell.removeParagraph(p);
      }
    } else {
      cell.createParagraph(text);
    }

    return this;
  }

  /**
   * Sets the table width
   *
   * Defines the total width of the table. Use with {@link setWidthType}
   * to specify if width is in twips, percentage, or auto.
   *
   * @param twips - Width value (interpretation depends on widthType)
   *   - For 'dxa' (default): Width in twips (1/20th of a point)
   *   - For 'pct': Percentage * 50 (e.g., 5000 = 100%)
   *   - For 'auto': Value is ignored
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * table.setWidth(9360);              // 6.5 inches in twips
   * table.setWidth(5000).setWidthType('pct');  // 100% width
   * ```
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
   * Sets table horizontal alignment
   *
   * Controls where the table is positioned horizontally on the page.
   *
   * @param alignment - Alignment value ('left' |'center' | 'right')
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * table.setAlignment('center');  // Center the table on page
   * table.setAlignment('right');   // Align table to right margin
   * ```
   */
  setAlignment(alignment: TableAlignment): this {
    const prev = this.formatting.alignment;
    this.formatting.alignment = alignment;
    if (this.trackingContext?.isEnabled() && prev !== alignment) {
      this.trackingContext.trackTableChange(this, 'alignment', prev, alignment);
    }
    return this;
  }

  /**
   * Sets table layout algorithm
   *
   * Controls how table column widths are calculated.
   *
   * @param layout - Layout type
   *   - 'auto': Columns auto-fit to content and window width
   *   - 'fixed': Columns use fixed widths (faster rendering)
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * table.setLayout('auto');   // Auto-fit to window
   * table.setLayout('fixed');  // Use fixed column widths
   * ```
   */
  setLayout(layout: TableLayout): this {
    const prev = this.formatting.layout;
    // Normalize "auto" to "autofit" per ECMA-376 ST_TblLayoutType (§17.18.87)
    this.formatting.layout = layout === 'auto' ? 'autofit' : layout;
    if (this.trackingContext?.isEnabled() && prev !== this.formatting.layout) {
      this.trackingContext.trackTableChange(this, 'layout', prev, this.formatting.layout);
    }
    return this;
  }

  /**
   * Sets table borders
   *
   * Defines borders for all sides of the table and interior borders.
   * By default, also applies borders to all cells within the table for consistency.
   *
   * @param borders - Border definitions for each edge
   * @param borders.top - Top border of table
   * @param borders.bottom - Bottom border of table
   * @param borders.left - Left border of table
   * @param borders.right - Right border of table
   * @param borders.insideH - Horizontal borders between rows
   * @param borders.insideV - Vertical borders between columns
   * @param options - Additional options
   * @param options.applyToCells - Whether to also set borders on all cells (default: true)
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * // Set borders on table AND all cells (default behavior)
   * table.setBorders({
   *   top: { style: 'single', size: 4, color: '000000' },
   *   bottom: { style: 'single', size: 4, color: '000000' },
   *   insideH: { style: 'single', size: 2, color: 'CCCCCC' },
   *   insideV: { style: 'single', size: 2, color: 'CCCCCC' }
   * });
   *
   * // Set only table-level borders (no cell borders)
   * table.setBorders({ top: { style: 'single', size: 4 } }, { applyToCells: false });
   * ```
   */
  setBorders(borders: TableBorders, options?: { applyToCells?: boolean }): this {
    const prev = this.formatting.borders;
    // Set table-level borders (w:tblBorders)
    this.formatting.borders = borders;
    if (this.trackingContext?.isEnabled() && prev !== borders) {
      this.trackingContext.trackTableChange(this, 'borders', prev, borders);
    }

    // Also apply to all cells for consistency (default: true)
    if (options?.applyToCells !== false) {
      const cellBorders = {
        top: borders.top,
        bottom: borders.bottom,
        left: borders.left,
        right: borders.right,
      };

      for (const row of this.rows) {
        // Update row-level table property exceptions if they have borders
        // Per ECMA-376, w:tblPrEx can contain border overrides that must also be updated
        const exceptions = row.getTablePropertyExceptions();
        if (exceptions?.borders) {
          row.setTablePropertyExceptions({
            ...exceptions,
            borders: borders, // Use full table borders for row-level exceptions
          });
        }

        // Update cell borders
        for (const cell of row.getCells()) {
          cell.setBorders(cellBorders);
        }
      }
    }

    return this;
  }

  /**
   * Sets all borders to the same style
   *
   * Convenience method that applies identical borders to all edges
   * (top, bottom, left, right, insideH, insideV) and all cells.
   *
   * @param border - Border definition to apply uniformly
   * @param options - Additional options
   * @param options.applyToCells - Whether to also set borders on all cells (default: true)
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * // Apply single black border to all edges and cells
   * table.setAllBorders({
   *   style: 'single',
   *   size: 4,
   *   color: '000000'
   * });
   *
   * // Apply only to table, not individual cells
   * table.setAllBorders({ style: 'single', size: 4 }, { applyToCells: false });
   * ```
   */
  setAllBorders(border: TableBorder, options?: { applyToCells?: boolean }): this {
    return this.setBorders(
      {
        top: border,
        bottom: border,
        left: border,
        right: border,
        insideH: border,
        insideV: border,
      },
      options
    );
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
    const prev = this.formatting.cellSpacing;
    this.formatting.cellSpacing = twips;
    if (this.trackingContext?.isEnabled() && prev !== twips) {
      this.trackingContext.trackTableChange(this, 'cellSpacing', prev, twips);
    }
    return this;
  }

  /**
   * Sets left indent
   * @param twips - Indent in twips
   * @returns This table for chaining
   */
  setIndent(twips: number): this {
    const prev = this.formatting.indent;
    this.formatting.indent = twips;
    if (this.trackingContext?.isEnabled() && prev !== twips) {
      this.trackingContext.trackTableChange(this, 'indent', prev, twips);
    }
    return this;
  }

  /**
   * Sets the indent width type per ECMA-376 ST_TblWidth
   */
  setIndentType(type: TableWidthType): this {
    const prev = this.formatting.indentType;
    this.formatting.indentType = type;
    if (this.trackingContext?.isEnabled() && prev !== type) {
      this.trackingContext.trackTableChange(this, 'indentType', prev, type);
    }
    return this;
  }

  /**
   * Gets the indent width type
   */
  getIndentType(): TableWidthType | undefined {
    return this.formatting.indentType;
  }

  /**
   * Sets table style reference
   * @param style - Table style ID (e.g., 'Table1', 'TableGrid')
   * @returns This table for chaining
   */
  setStyle(style: string): this {
    const prev = this.formatting.style;
    this.formatting.style = style;
    if (this.trackingContext?.isEnabled() && prev !== style) {
      this.trackingContext.trackTableChange(this, 'style', prev, style);
    }
    return this;
  }

  /**
   * Sets table look flags (appearance settings)
   * @param tblLook - Table look value (e.g., '0000', '04A0')
   * @returns This table for chaining
   */
  setTblLook(tblLook: string): this {
    const prev = this.formatting.tblLook;
    this.formatting.tblLook = tblLook;
    if (this.trackingContext?.isEnabled() && prev !== tblLook) {
      this.trackingContext.trackTableChange(this, 'tblLook', prev, tblLook);
    }
    return this;
  }

  /**
   * Sets the number of rows in each row band for table style alternation
   * Per ECMA-376 Part 1 §17.4.52
   * @param size - Number of rows per band (default: 1)
   * @returns This table for chaining
   */
  setStyleRowBandSize(size: number): this {
    const prev = this.formatting.tblStyleRowBandSize;
    this.formatting.tblStyleRowBandSize = size;
    if (this.trackingContext?.isEnabled() && prev !== size) {
      this.trackingContext.trackTableChange(this, 'tblStyleRowBandSize', prev, size);
    }
    return this;
  }

  /**
   * Sets the number of columns in each column band for table style alternation
   * Per ECMA-376 Part 1 §17.4.51
   * @param size - Number of columns per band (default: 1)
   * @returns This table for chaining
   */
  setStyleColBandSize(size: number): this {
    const prev = this.formatting.tblStyleColBandSize;
    this.formatting.tblStyleColBandSize = size;
    if (this.trackingContext?.isEnabled() && prev !== size) {
      this.trackingContext.trackTableChange(this, 'tblStyleColBandSize', prev, size);
    }
    return this;
  }

  /**
   * Decodes tblLook hex string into boolean flags
   * Per ECMA-376 Part 1 Section 17.4.56
   * @returns Object with boolean flags for each tblLook property
   */
  getTblLookFlags(): {
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    noHBand: boolean;
    noVBand: boolean;
  } {
    const hex = this.formatting.tblLook || '0000';
    const value = parseInt(hex, 16) || 0;

    return {
      firstRow: (value & 0x0020) !== 0,
      lastRow: (value & 0x0040) !== 0,
      firstColumn: (value & 0x0080) !== 0,
      lastColumn: (value & 0x0100) !== 0,
      noHBand: (value & 0x0200) !== 0,
      noVBand: (value & 0x0400) !== 0,
    };
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
    contentRules?: {
      condition: (cellText: string, rowIndex: number, colIndex: number) => boolean;
      formatting: Partial<CellFormatting>;
    }[];
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
        const formatting = isEven ? rules.alternatingRows?.even : rules.alternatingRows?.odd;

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
  private applyCellFormatting(cell: TableCell, formatting: Partial<CellFormatting>): void {
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
    const prev = this.formatting.position;
    this.formatting.position = position;
    if (this.trackingContext?.isEnabled() && prev !== position) {
      this.trackingContext.trackTableChange(this, 'position', prev, position);
    }
    return this;
  }

  /**
   * Sets whether table can overlap with other floating tables
   * Per ECMA-376 Part 1 §17.4.30
   * @param overlap - True to allow overlap, false to prevent
   * @returns This table for chaining
   */
  setOverlap(overlap: boolean): this {
    const prev = this.formatting.overlap;
    this.formatting.overlap = overlap;
    if (this.trackingContext?.isEnabled() && prev !== overlap) {
      this.trackingContext.trackTableChange(this, 'overlap', prev, overlap);
    }
    return this;
  }

  /**
   * Sets bidirectional (right-to-left) visual layout
   * Per ECMA-376 Part 1 §17.4.1
   * @param bidi - True for RTL layout, false for LTR
   * @returns This table for chaining
   */
  setBidiVisual(bidi: boolean): this {
    const prev = this.formatting.bidiVisual;
    this.formatting.bidiVisual = bidi;
    if (this.trackingContext?.isEnabled() && prev !== bidi) {
      this.trackingContext.trackTableChange(this, 'bidiVisual', prev, bidi);
    }
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
    const prev = this.formatting.tableGrid;
    this.formatting.tableGrid = widths;
    if (this.trackingContext?.isEnabled() && prev) {
      // Create tblGridChange (not tblPrChange) per ECMA-376 §17.13.5.35
      // Only create if there was a previous grid to record
      if (!this.tblGridChange) {
        const revisionManager = this.trackingContext.getRevisionManager();
        this.tblGridChange = TableGridChange.create(
          revisionManager.consumeNextId(),
          prev.map((w) => ({ width: w })),
          this.trackingContext.getAuthor(),
          new Date()
        );
      }
    }
    return this;
  }

  /**
   * Sets table caption for accessibility
   * Per ECMA-376 Part 1 §17.4.58
   * @param caption - Table caption text
   * @returns This table for chaining
   */
  setCaption(caption: string): this {
    const prev = this.formatting.caption;
    this.formatting.caption = caption;
    if (this.trackingContext?.isEnabled() && prev !== caption) {
      this.trackingContext.trackTableChange(this, 'caption', prev, caption);
    }
    return this;
  }

  /**
   * Sets table description for accessibility
   * Per ECMA-376 Part 1 §17.4.63
   * @param description - Table description text
   * @returns This table for chaining
   */
  setDescription(description: string): this {
    const prev = this.formatting.description;
    this.formatting.description = description;
    if (this.trackingContext?.isEnabled() && prev !== description) {
      this.trackingContext.trackTableChange(this, 'description', prev, description);
    }
    return this;
  }

  /**
   * Sets table-level shading (background)
   *
   * Per ECMA-376 Part 1 §17.4.56 (w:shd), this sets the default
   * background shading for the entire table. Individual cells can
   * override this with their own shading.
   *
   * @param shading - Table shading configuration
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * table.setShading({ fill: 'F0F0F0' }); // Light gray background
   * table.setShading({ fill: 'FFFFFF', pattern: 'pct12', color: '000000' });
   * ```
   */
  setShading(shading: TableShading): this {
    const prev = this.formatting.shading;
    this.formatting.shading = shading;
    if (this.trackingContext?.isEnabled() && prev !== shading) {
      this.trackingContext.trackTableChange(this, 'shading', prev, shading);
    }
    return this;
  }

  /**
   * Gets table-level shading configuration
   * @returns Table shading or undefined if not set
   */
  getShading(): TableShading | undefined {
    return this.formatting.shading;
  }

  /**
   * Sets table width type
   *
   * Defines how the table width value should be interpreted.
   * Per ECMA-376 Part 1 §17.4.64
   *
   * @param type - Width interpretation type
   *   - 'auto': Automatic width (ignores width value)
   *   - 'dxa': Width in twips (1/20th of a point)
   *   - 'pct': Width as percentage (value * 50 = percentage, e.g., 5000 = 100%)
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * table.setWidth(5000).setWidthType('pct');  // 100% page width
   * table.setWidth(9360).setWidthType('dxa');  // 6.5 inches (absolute)
   * table.setWidthType('auto');                // Auto-fit content
   * ```
   */
  setWidthType(type: TableWidthType): this {
    const prev = this.formatting.widthType;
    this.formatting.widthType = type;
    if (this.trackingContext?.isEnabled() && prev !== type) {
      this.trackingContext.trackTableChange(this, 'widthType', prev, type);
    }
    return this;
  }

  /**
   * Sets cell spacing type
   * @param type - Cell spacing type
   * @returns This table for chaining
   */
  setCellSpacingType(type: TableWidthType): this {
    const prev = this.formatting.cellSpacingType;
    this.formatting.cellSpacingType = type;
    if (this.trackingContext?.isEnabled() && prev !== type) {
      this.trackingContext.trackTableChange(this, 'cellSpacingType', prev, type);
    }
    return this;
  }

  /**
   * Gets a copy of the table formatting
   *
   * Returns a copy of all formatting properties including width, alignment,
   * layout, borders, and other table-level settings.
   *
   * @returns Copy of the table formatting object
   *
   * @example
   * ```typescript
   * const formatting = table.getFormatting();
   * console.log(`Width: ${formatting.width} twips`);
   * console.log(`Layout: ${formatting.layout}`);
   * ```
   */
  getFormatting(): TableFormatting {
    return { ...this.formatting };
  }

  // ============================================================================
  // Individual Formatting Getters
  // ============================================================================

  /**
   * Gets the table width in twips
   * @returns Width in twips or undefined if not set
   */
  getWidth(): number | undefined {
    return this.formatting.width;
  }

  /**
   * Gets the table width type
   * @returns Width type ('auto', 'dxa', 'pct', 'nil') or undefined
   */
  getWidthType(): string | undefined {
    return this.formatting.widthType;
  }

  /**
   * Gets the table horizontal alignment
   * @returns Alignment ('left', 'center', 'right') or undefined
   */
  getAlignment(): string | undefined {
    return this.formatting.alignment;
  }

  /**
   * Gets the table layout type
   * @returns Layout ('autofit', 'fixed') or undefined
   */
  getLayout(): string | undefined {
    return this.formatting.layout;
  }

  /**
   * Gets the table left indentation in twips
   * @returns Indent in twips or undefined if not set
   */
  getIndent(): number | undefined {
    return this.formatting.indent;
  }

  /**
   * Gets the table borders
   * @returns Borders object or undefined
   */
  getBorders(): TableFormatting['borders'] | undefined {
    return this.formatting.borders;
  }

  /**
   * Gets the column widths (table grid)
   * @returns Array of column widths in twips
   */
  getColumnWidths(): number[] {
    return [...(this.formatting.tableGrid || [])];
  }

  /**
   * Gets the cell spacing value
   * @returns Cell spacing in twips or undefined
   */
  getCellSpacing(): number | undefined {
    return this.formatting.cellSpacing;
  }

  /**
   * Gets the default cell margins (padding) for all cells
   * Per ECMA-376 Part 1 §17.4.42 (tblCellMar)
   * @returns Cell margins object or undefined if not set
   */
  getCellMargins(): TableCellMargins | undefined {
    return this.formatting.cellMargins;
  }

  /**
   * Sets the default cell margins (padding) for all cells
   * Per ECMA-376 Part 1 §17.4.42 (tblCellMar)
   * @param margins - Cell margins in twips
   * @returns This table for chaining
   * @example
   * ```typescript
   * table.setCellMargins({ top: 43, left: 115, bottom: 43, right: 115 });
   * ```
   */
  setCellMargins(margins: TableCellMargins): this {
    const prev = this.formatting.cellMargins;
    this.formatting.cellMargins = margins;
    if (this.trackingContext?.isEnabled() && prev !== margins) {
      this.trackingContext.trackTableChange(this, 'cellMargins', prev, margins);
    }
    return this;
  }

  /**
   * Gets the table style ID
   * @returns Style ID or undefined if not set
   */
  getStyle(): string | undefined {
    return this.formatting.style;
  }

  /**
   * Gets the raw table look (tblLook) value
   * @returns Table look hex string or undefined if not set
   */
  getTblLook(): string | undefined {
    return this.formatting.tblLook;
  }

  /**
   * Gets the table position properties for floating tables
   * @returns Table position properties or undefined if not set
   */
  getPosition(): TablePositionProperties | undefined {
    return this.formatting.position ? { ...this.formatting.position } : undefined;
  }

  /**
   * Gets whether table can overlap with other floating tables
   * @returns True if overlap is allowed, false if not, undefined if not set
   */
  getOverlap(): boolean | undefined {
    return this.formatting.overlap;
  }

  /**
   * Gets whether bidirectional (RTL) visual layout is enabled
   * @returns True if RTL layout, false if LTR, undefined if not set
   */
  getBidiVisual(): boolean | undefined {
    return this.formatting.bidiVisual;
  }

  /**
   * Gets the table grid column widths
   * @returns Array of column widths in twips, or undefined if not set
   */
  getTableGrid(): number[] | undefined {
    return this.formatting.tableGrid ? [...this.formatting.tableGrid] : undefined;
  }

  /**
   * Gets the table caption for accessibility
   * @returns Caption text or undefined if not set
   */
  getCaption(): string | undefined {
    return this.formatting.caption;
  }

  /**
   * Gets the table description for accessibility
   * @returns Description text or undefined if not set
   */
  getDescription(): string | undefined {
    return this.formatting.description;
  }

  /**
   * Gets the cell spacing type
   * @returns Cell spacing type or undefined if not set
   */
  getCellSpacingType(): TableWidthType | undefined {
    return this.formatting.cellSpacingType;
  }

  // ============================================================================
  // Checker Methods
  // ============================================================================

  /**
   * Checks if the table has any rows
   * @returns True if table has rows
   */
  hasRows(): boolean {
    return this.rows.length > 0;
  }

  /**
   * Checks if the table is a floating table (has positioning)
   * @returns True if table has positioning properties
   */
  isFloating(): boolean {
    return this.formatting.position !== undefined;
  }

  /**
   * Checks if the table has a style applied
   * @returns True if table has a style
   */
  hasStyle(): boolean {
    return this.formatting.style !== undefined && this.formatting.style !== '';
  }

  /**
   * Sets the StylesManager reference for conditional formatting resolution.
   * Propagates to all paragraphs in all cells.
   * @internal
   */
  _setStylesManager(manager: import('../formatting/StylesManager').StylesManager): void {
    this._stylesManager = manager;
    // Propagate to all paragraphs in all cells
    for (const row of this.rows) {
      for (const cell of row.getCells()) {
        for (const para of cell.getParagraphs()) {
          para._setStylesManager(manager);
        }
      }
    }
  }

  /**
   * Gets the StylesManager reference for conditional formatting resolution.
   * @internal
   */
  _getStylesManager(): import('../formatting/StylesManager').StylesManager | undefined {
    return this._stylesManager;
  }

  /**
   * Converts the table to WordprocessingML XML element
   * @returns XMLElement representing the table
   */
  toXML(): XMLElement {
    const tblPrChildren: XMLElement[] = [];

    // CT_TblPr element order per ECMA-376:
    // tblStyle, tblpPr, tblOverlap, bidiVisual, tblStyleRowBandSize, tblStyleColBandSize,
    // tblW, jc, tblCellSpacing, tblInd, tblBorders, shd, tblLayout, tblCellMar, tblLook,
    // tblCaption, tblDescription, tblPrChange

    // 1. tblStyle
    if (this.formatting.style) {
      tblPrChildren.push(XMLBuilder.wSelf('tblStyle', { 'w:val': this.formatting.style }));
    }

    // 2. tblpPr - table positioning properties (floating tables)
    if (this.formatting.position) {
      const pos = this.formatting.position;
      const posAttrs: Record<string, string | number> = {};

      if (pos.x !== undefined) posAttrs['w:tblpX'] = pos.x;
      if (pos.y !== undefined) posAttrs['w:tblpY'] = pos.y;
      if (pos.horizontalAnchor) posAttrs['w:horzAnchor'] = pos.horizontalAnchor;
      if (pos.verticalAnchor) posAttrs['w:vertAnchor'] = pos.verticalAnchor;
      if (pos.horizontalAlignment) posAttrs['w:tblpXSpec'] = pos.horizontalAlignment;
      if (pos.verticalAlignment) posAttrs['w:tblpYSpec'] = pos.verticalAlignment;
      if (pos.leftFromText !== undefined) posAttrs['w:leftFromText'] = pos.leftFromText;
      if (pos.rightFromText !== undefined) posAttrs['w:rightFromText'] = pos.rightFromText;
      if (pos.topFromText !== undefined) posAttrs['w:topFromText'] = pos.topFromText;
      if (pos.bottomFromText !== undefined) posAttrs['w:bottomFromText'] = pos.bottomFromText;

      if (Object.keys(posAttrs).length > 0) {
        tblPrChildren.push(XMLBuilder.wSelf('tblpPr', posAttrs));
      }
    }

    // 3. tblOverlap
    if (this.formatting.overlap !== undefined) {
      tblPrChildren.push(
        XMLBuilder.wSelf('tblOverlap', {
          'w:val': this.formatting.overlap ? 'overlap' : 'never',
        })
      );
    }

    // 4. bidiVisual — OnOffOnlyType (not CT_OnOff); w:val is restricted to "on"/"off"
    // by ST_OnOffOnly. Emit "off" for explicit false so the override-of-style
    // distinction between "inherited" (absent) and "explicitly off" survives.
    if (this.formatting.bidiVisual !== undefined) {
      if (this.formatting.bidiVisual) {
        tblPrChildren.push(XMLBuilder.wSelf('bidiVisual'));
      } else {
        tblPrChildren.push(XMLBuilder.wSelf('bidiVisual', { 'w:val': 'off' }));
      }
    }

    // 5-6. tblStyleRowBandSize / tblStyleColBandSize
    // Only valid within table style definitions (CT_TblPrBase in w:style),
    // not in direct tblPr per Transitional schema. Style.ts handles serialization.
    // Values are preserved in formatting for style use and tblPrChange serialization.

    // 7. tblW
    if (this.formatting.width !== undefined) {
      const widthType = this.formatting.widthType || 'dxa';
      tblPrChildren.push(
        XMLBuilder.wSelf('tblW', {
          'w:w': this.formatting.width,
          'w:type': widthType,
        })
      );
    }

    // 8. jc (alignment)
    if (this.formatting.alignment) {
      tblPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': this.formatting.alignment }));
    }

    // 9. tblCellSpacing
    if (this.formatting.cellSpacing !== undefined) {
      const cellSpacingType = this.formatting.cellSpacingType || 'dxa';
      tblPrChildren.push(
        XMLBuilder.wSelf('tblCellSpacing', {
          'w:w': this.formatting.cellSpacing,
          'w:type': cellSpacingType,
        })
      );
    }

    // 10. tblInd
    if (this.formatting.indent !== undefined) {
      tblPrChildren.push(
        XMLBuilder.wSelf('tblInd', {
          'w:w': this.formatting.indent,
          'w:type': this.formatting.indentType || 'dxa',
        })
      );
    }

    // 11. tblBorders
    if (this.formatting.borders) {
      const borderElements: XMLElement[] = [];
      const borders = this.formatting.borders;

      if (borders.top) {
        borderElements.push(XMLBuilder.createBorder('top', borders.top));
      }
      if (borders.left) {
        borderElements.push(XMLBuilder.createBorder('left', borders.left));
      }
      if (borders.bottom) {
        borderElements.push(XMLBuilder.createBorder('bottom', borders.bottom));
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

    // 12. shd (table shading/background)
    if (this.formatting.shading) {
      const shdAttrs = buildShadingAttributes(this.formatting.shading);
      if (Object.keys(shdAttrs).length > 0) {
        tblPrChildren.push(XMLBuilder.wSelf('shd', shdAttrs));
      }
    }

    // 13. tblLayout
    if (this.formatting.layout) {
      tblPrChildren.push(XMLBuilder.wSelf('tblLayout', { 'w:type': this.formatting.layout }));
    }

    // 14. tblCellMar
    if (this.formatting.cellMargins) {
      const marginElements: XMLElement[] = [];
      if (this.formatting.cellMargins.top !== undefined) {
        marginElements.push(
          XMLBuilder.wSelf('top', {
            'w:w': this.formatting.cellMargins.top,
            'w:type': 'dxa',
          })
        );
      }
      if (this.formatting.cellMargins.left !== undefined) {
        marginElements.push(
          XMLBuilder.wSelf('left', {
            'w:w': this.formatting.cellMargins.left,
            'w:type': 'dxa',
          })
        );
      }
      if (this.formatting.cellMargins.bottom !== undefined) {
        marginElements.push(
          XMLBuilder.wSelf('bottom', {
            'w:w': this.formatting.cellMargins.bottom,
            'w:type': 'dxa',
          })
        );
      }
      if (this.formatting.cellMargins.right !== undefined) {
        marginElements.push(
          XMLBuilder.wSelf('right', {
            'w:w': this.formatting.cellMargins.right,
            'w:type': 'dxa',
          })
        );
      }
      if (marginElements.length > 0) {
        tblPrChildren.push(XMLBuilder.w('tblCellMar', undefined, marginElements));
      }
    }

    // 15. tblLook
    if (this.formatting.tblLook) {
      tblPrChildren.push(
        XMLBuilder.wSelf('tblLook', buildTblLookAttributes(this.formatting.tblLook))
      );
    }

    // Add table caption (accessibility)
    if (this.formatting.caption) {
      tblPrChildren.push(XMLBuilder.wSelf('tblCaption', { 'w:val': this.formatting.caption }));
    }

    // Add table description (accessibility)
    if (this.formatting.description) {
      tblPrChildren.push(
        XMLBuilder.wSelf('tblDescription', {
          'w:val': this.formatting.description,
        })
      );
    }

    // Add table property change (w:tblPrChange) per ECMA-376 Part 1 §17.13.5.36
    // Must be last child of w:tblPr
    if (this.tblPrChange) {
      const changeAttrs: Record<string, string | number> = {
        'w:id': this.tblPrChange.id,
        'w:author': this.tblPrChange.author,
        'w:date': this.tblPrChange.date,
      };
      // tblPrChange previousProperties — same CT_TblPr order as main tblPr per ECMA-376 §17.4.58:
      // tblStyle, tblpPr, tblOverlap, bidiVisual, tblStyleRowBandSize, tblStyleColBandSize,
      // tblW, jc, tblCellSpacing, tblInd, tblBorders, shd, tblLayout, tblCellMar, tblLook,
      // tblCaption, tblDescription
      const prevTblPrChildren: XMLElement[] = [];
      const prev = this.tblPrChange.previousProperties;
      if (prev) {
        // 1. tblStyle
        if (prev.style) {
          prevTblPrChildren.push(XMLBuilder.wSelf('tblStyle', { 'w:val': prev.style }));
        }
        // 2. tblpPr (floating table position) — passthrough if present
        if (prev.position) {
          const posAttrs: Record<string, string | number> = {};
          const pos = prev.position;
          if (pos.x !== undefined) posAttrs['w:tblpX'] = pos.x;
          if (pos.y !== undefined) posAttrs['w:tblpY'] = pos.y;
          if (pos.horizontalAnchor) posAttrs['w:horzAnchor'] = pos.horizontalAnchor;
          if (pos.verticalAnchor) posAttrs['w:vertAnchor'] = pos.verticalAnchor;
          if (pos.leftFromText !== undefined) posAttrs['w:leftFromText'] = pos.leftFromText;
          if (pos.rightFromText !== undefined) posAttrs['w:rightFromText'] = pos.rightFromText;
          if (pos.topFromText !== undefined) posAttrs['w:topFromText'] = pos.topFromText;
          if (pos.bottomFromText !== undefined) posAttrs['w:bottomFromText'] = pos.bottomFromText;
          if (Object.keys(posAttrs).length > 0) {
            prevTblPrChildren.push(XMLBuilder.wSelf('tblpPr', posAttrs));
          }
        }
        // 3. tblOverlap
        if (prev.overlap) {
          prevTblPrChildren.push(XMLBuilder.wSelf('tblOverlap', { 'w:val': prev.overlap }));
        }
        // 4. bidiVisual — OnOffOnlyType; preserve explicit false inside tblPrChange
        if (prev.bidiVisual !== undefined) {
          if (prev.bidiVisual) {
            prevTblPrChildren.push(XMLBuilder.wSelf('bidiVisual'));
          } else {
            prevTblPrChildren.push(XMLBuilder.wSelf('bidiVisual', { 'w:val': 'off' }));
          }
        }
        // 5. tblStyleRowBandSize
        if (prev.tblStyleRowBandSize !== undefined) {
          prevTblPrChildren.push(
            XMLBuilder.wSelf('tblStyleRowBandSize', { 'w:val': prev.tblStyleRowBandSize })
          );
        }
        // 6. tblStyleColBandSize
        if (prev.tblStyleColBandSize !== undefined) {
          prevTblPrChildren.push(
            XMLBuilder.wSelf('tblStyleColBandSize', { 'w:val': prev.tblStyleColBandSize })
          );
        }
        // 7. tblW
        if (prev.width !== undefined) {
          prevTblPrChildren.push(
            XMLBuilder.wSelf('tblW', {
              'w:w': prev.width,
              'w:type': prev.widthType || 'dxa',
            })
          );
        }
        // 8. jc
        if (prev.alignment) {
          prevTblPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': prev.alignment }));
        }
        // 9. tblCellSpacing
        if (prev.cellSpacing !== undefined) {
          const csAttrs: Record<string, string | number> = {
            'w:w': prev.cellSpacing,
            'w:type': prev.cellSpacingType || 'dxa',
          };
          prevTblPrChildren.push(XMLBuilder.wSelf('tblCellSpacing', csAttrs));
        }
        // 10. tblInd
        if (prev.indent !== undefined) {
          prevTblPrChildren.push(
            XMLBuilder.wSelf('tblInd', { 'w:w': prev.indent, 'w:type': prev.indentType || 'dxa' })
          );
        }
        // 11. tblBorders — CT_Border §17.18.2 requires w:val; default "nil".
        // Full attribute set (themeColor/themeTint/themeShade/shadow/frame)
        // preserved so tblPrChange tracked history matches modern Word output.
        if (prev.borders) {
          const borderChildren: XMLElement[] = [];
          const bNames = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;
          for (const name of bNames) {
            const b = prev.borders[name];
            if (b) {
              const bAttrs: Record<string, string | number> = { 'w:val': b.style ?? 'nil' };
              if (b.size !== undefined) bAttrs['w:sz'] = b.size;
              if (b.color) bAttrs['w:color'] = b.color;
              const bx = b as {
                space?: number;
                themeColor?: string;
                themeTint?: string;
                themeShade?: string;
                shadow?: boolean;
                frame?: boolean;
              };
              if (bx.space !== undefined) bAttrs['w:space'] = bx.space;
              if (bx.themeColor) bAttrs['w:themeColor'] = bx.themeColor;
              if (bx.themeTint) bAttrs['w:themeTint'] = bx.themeTint;
              if (bx.themeShade) bAttrs['w:themeShade'] = bx.themeShade;
              if (bx.shadow !== undefined) bAttrs['w:shadow'] = bx.shadow ? '1' : '0';
              if (bx.frame !== undefined) bAttrs['w:frame'] = bx.frame ? '1' : '0';
              borderChildren.push(XMLBuilder.wSelf(name, bAttrs));
            }
          }
          if (borderChildren.length > 0) {
            prevTblPrChildren.push(XMLBuilder.w('tblBorders', undefined, borderChildren));
          }
        }
        // 12. shd
        if (prev.shading) {
          const shadingAttrs = buildShadingAttributes(prev.shading);
          if (Object.keys(shadingAttrs).length > 0) {
            prevTblPrChildren.push(XMLBuilder.wSelf('shd', shadingAttrs));
          }
        }
        // 13. tblLayout
        if (prev.layout) {
          prevTblPrChildren.push(XMLBuilder.wSelf('tblLayout', { 'w:type': prev.layout }));
        }
        // 14. tblCellMar
        if (prev.cellMargins) {
          const cmChildren: XMLElement[] = [];
          if (prev.cellMargins.top !== undefined) {
            cmChildren.push(
              XMLBuilder.wSelf('top', { 'w:w': prev.cellMargins.top, 'w:type': 'dxa' })
            );
          }
          if (prev.cellMargins.left !== undefined) {
            cmChildren.push(
              XMLBuilder.wSelf('left', { 'w:w': prev.cellMargins.left, 'w:type': 'dxa' })
            );
          }
          if (prev.cellMargins.bottom !== undefined) {
            cmChildren.push(
              XMLBuilder.wSelf('bottom', { 'w:w': prev.cellMargins.bottom, 'w:type': 'dxa' })
            );
          }
          if (prev.cellMargins.right !== undefined) {
            cmChildren.push(
              XMLBuilder.wSelf('right', { 'w:w': prev.cellMargins.right, 'w:type': 'dxa' })
            );
          }
          if (cmChildren.length > 0) {
            prevTblPrChildren.push(XMLBuilder.w('tblCellMar', undefined, cmChildren));
          }
        }
        // 15. tblLook
        if (prev.tblLook) {
          prevTblPrChildren.push(XMLBuilder.wSelf('tblLook', buildTblLookAttributes(prev.tblLook)));
        }
        // 16. tblCaption
        if (prev.caption) {
          prevTblPrChildren.push(XMLBuilder.wSelf('tblCaption', { 'w:val': prev.caption }));
        }
        // 17. tblDescription
        if (prev.description) {
          prevTblPrChildren.push(XMLBuilder.wSelf('tblDescription', { 'w:val': prev.description }));
        }
      }
      const prevTblPr = XMLBuilder.w('tblPr', undefined, prevTblPrChildren);
      tblPrChildren.push(XMLBuilder.w('tblPrChange', changeAttrs, [prevTblPr]));
    }

    // Build table element
    const tableChildren: XMLElement[] = [];

    // Add table properties
    tableChildren.push(XMLBuilder.w('tblPr', undefined, tblPrChildren));

    // Add table grid (column definitions)
    // Per ECMA-376 Part 1 §17.4.49, w:tblGrid MUST be present in w:tbl
    // Use custom tableGrid if specified, otherwise auto-generate
    // CRITICAL: Use getTotalGridSpan() instead of getCellCount() to account for
    // cells with gridSpan (column span). A row with 2 cells where one spans 4 columns
    // should generate 5 grid columns, not 2.
    const gridWidths = this.formatting.tableGrid;
    const maxColumns = gridWidths
      ? gridWidths.length
      : Math.max(...this.rows.map((row) => row.getTotalGridSpan()), 0);

    // Always generate tblGrid - use at least 1 column for empty tables
    const gridColumnCount = maxColumns > 0 ? maxColumns : 1;
    const tblGridChildren: XMLElement[] = [];

    for (let i = 0; i < gridColumnCount; i++) {
      if (gridWidths?.[i] !== undefined) {
        // Use specified grid width
        tblGridChildren.push(XMLBuilder.wSelf('gridCol', { 'w:w': gridWidths[i] }));
      } else {
        // Auto width (default to 2880 twips = 2 inches)
        tblGridChildren.push(XMLBuilder.wSelf('gridCol', { 'w:w': 2880 }));
      }
    }
    // Add tblGridChange if present (per ECMA-376 §17.13.5.35, must be last child of tblGrid)
    if (this.tblGridChange) {
      tblGridChildren.push(this.tblGridChange.toXML());
    }
    tableChildren.push(XMLBuilder.w('tblGrid', undefined, tblGridChildren));

    // Add rows
    for (const row of this.rows) {
      tableChildren.push(row.toXML());
    }

    return XMLBuilder.w('tbl', undefined, tableChildren);
  }

  /**
   * Removes a row from the table
   *
   * Deletes the row at the specified index and shifts subsequent rows up.
   *
   * @param index - The row index to remove (0-based)
   * @returns True if the row was removed, false if index was invalid
   *
   * @example
   * ```typescript
   * table.removeRow(0);  // Remove first row
   * table.removeRow(2);  // Remove third row
   * ```
   */
  removeRow(index: number): boolean {
    if (index >= 0 && index < this.rows.length) {
      // Per ECMA-376 §17.4.38, a table must contain at least one row
      if (this.rows.length <= 1 && !this.trackingContext?.isEnabled()) {
        return false;
      }
      // When tracking enabled, mark cells with cellDel and wrap content in w:del
      if (this.trackingContext?.isEnabled()) {
        const author = this.trackingContext.getAuthor();
        const row = this.rows[index]!;
        for (const cell of row.getCells()) {
          const cellDelRevision = Revision.createTableCellDelete(author, []);
          cell.setCellRevision(cellDelRevision);
          // Wrap paragraph runs in w:del so content appears as deleted
          for (const para of cell.getParagraphs()) {
            const runs = para.getRuns();
            if (runs.length > 0) {
              const deletion = Revision.createDeletion(author, runs);
              para.addRevision(deletion);
            }
          }
        }
        return true;
      }
      this.rows.splice(index, 1);
      return true;
    }
    return false;
  }

  /**
   * Internal: removes a row without triggering tracking-context side effects.
   * Used by the revision acceptor when resolving tracked row deletions —
   * accepting a pre-existing tracked deletion must NOT create a new
   * revision, so the usual `removeRow()` path (which wraps content in
   * `w:del`) is unsuitable. Bypasses the "at least one row" guard because
   * the caller is responsible for any empty-table cleanup.
   *
   * @internal
   */
  _removeRowAtIndex(index: number): boolean {
    if (index < 0 || index >= this.rows.length) return false;
    this.rows.splice(index, 1);
    return true;
  }

  /**
   * Inserts a row at a specific position
   *
   * Inserts a new row at the specified index. If no row is provided,
   * creates an empty row with columns matching the table's column count.
   *
   * @param index - Position to insert at (0-based, clamped to valid range)
   * @param row - Optional TableRow to insert (creates new row if not provided)
   * @returns The inserted TableRow instance
   *
   * @example
   * ```typescript
   * // Insert empty row at beginning
   * const row = table.insertRow(0);
   *
   * // Insert custom row in the middle
   * const customRow = new TableRow(3);
   * table.insertRow(2, customRow);
   * ```
   */
  insertRow(index: number, row?: TableRow): TableRow {
    // Clamp index to valid range
    if (index < 0) index = 0;
    if (index > this.rows.length) index = this.rows.length;

    // Create new row if not provided, matching the grid column count
    // Use getTotalGridSpan() instead of getColumnCount() to account for merged cells
    if (!row) {
      const gridColumns = this.formatting.tableGrid
        ? this.formatting.tableGrid.length
        : Math.max(...this.rows.map((r) => r.getTotalGridSpan()), 1);
      row = new TableRow(gridColumns);
    }

    // Insert the row
    this.rows.splice(index, 0, row);

    // When tracking enabled, mark every cell in the new row with cellIns
    if (this.trackingContext?.isEnabled()) {
      const author = this.trackingContext.getAuthor();
      for (const cell of row.getCells()) {
        const revision = Revision.createTableCellInsert(author, []);
        cell.setCellRevision(revision);
      }
    }

    return row;
  }

  /**
   * Adds a column to all rows in the table
   *
   * Inserts a new cell in each row at the specified position.
   * If no index is provided, adds the column at the end.
   *
   * @param index - Optional position to insert the column (0-based, defaults to end)
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * table.addColumn();     // Add column at end
   * table.addColumn(0);    // Insert column at beginning
   * table.addColumn(2);    // Insert column at position 2
   * ```
   */
  addColumn(index?: number): this {
    const isTracking = this.trackingContext?.isEnabled();
    const author = isTracking ? this.trackingContext!.getAuthor() : '';

    for (const row of this.rows) {
      const cell = new TableCell();

      // Mark cell as inserted when tracking
      if (isTracking) {
        const revision = Revision.createTableCellInsert(author, []);
        cell.setCellRevision(revision);
      }

      const cells = row.getCells();

      if (index === undefined || index >= cells.length) {
        // Add to end
        row.addCell(cell);
      } else {
        // Insert at specific position
        const idx = Math.max(0, index);
        row.insertCellAt(idx, cell);
      }
    }
    return this;
  }

  /**
   * Removes a column from all rows in the table
   *
   * Deletes the cell at the specified column index in every row.
   *
   * @param index - The column index to remove (0-based)
   * @returns True if the column was removed from at least one row, false if index was invalid
   *
   * @example
   * ```typescript
   * table.removeColumn(0);  // Remove first column
   * table.removeColumn(2);  // Remove third column
   * ```
   */
  removeColumn(index: number): boolean {
    if (index < 0 || this.rows.length === 0) {
      return false;
    }

    // When tracking enabled, mark cells with cellDel instead of removing
    if (this.trackingContext?.isEnabled()) {
      const author = this.trackingContext.getAuthor();
      let marked = false;
      for (const row of this.rows) {
        const cells = row.getCells();
        if (index < cells.length) {
          const revision = Revision.createTableCellDelete(author, []);
          cells[index]!.setCellRevision(revision);
          marked = true;
        }
      }
      return marked;
    }

    let removed = false;
    for (const row of this.rows) {
      const cells = row.getCells();
      if (index < cells.length) {
        row.removeCellAt(index);
        removed = true;
      }
    }

    return removed;
  }

  /**
   * Gets the maximum column count across all rows
   *
   * Returns the highest number of cells in any row. Useful since
   * rows may have different numbers of cells.
   *
   * @returns Maximum number of columns in the table
   *
   * @example
   * ```typescript
   * console.log(`Table has up to ${table.getColumnCount()} columns`);
   * ```
   */
  getColumnCount(): number {
    if (this.rows.length === 0) {
      return 0;
    }
    return Math.max(...this.rows.map((row) => row.getCellCount()));
  }

  /**
   * Gets all cells at a given column index across all rows
   *
   * Returns the cell at position `colIndex` from each row. Rows that have
   * fewer cells than the requested index are skipped (their slot is not
   * included in the result). This uses simple cell-index addressing, not
   * grid-span-aware column mapping.
   *
   * @param colIndex - Column index (0-based)
   * @returns Array of cells at that column position
   *
   * @example
   * ```typescript
   * // Sum values in the third column
   * const cells = table.getColumnCells(2);
   * const total = cells.reduce((sum, cell) => sum + Number(cell.getText()), 0);
   * ```
   */
  getColumnCells(colIndex: number): TableCell[] {
    const cells: TableCell[] = [];
    for (const row of this.rows) {
      const cell = row.getCell(colIndex);
      if (cell) cells.push(cell);
    }
    return cells;
  }

  /**
   * Gets the text content of all cells in a column as a string array
   *
   * Like `getColumnCells()` but returns just the text values, skipping
   * the cell object layer. Rows with fewer cells than the column index
   * are skipped.
   *
   * @param colIndex - Column index (0-based)
   * @returns Array of cell text values
   *
   * @example
   * ```typescript
   * const names = table.getColumnTexts(0); // ['Name', 'Alice', 'Bob']
   * const total = table.getColumnTexts(2)
   *   .slice(1)
   *   .reduce((sum, v) => sum + Number(v), 0);
   * ```
   */
  getColumnTexts(colIndex: number): string[] {
    return this.getColumnCells(colIndex).map((cell) => cell.getText());
  }

  /**
   * Transforms the text content of every cell in a column
   *
   * Calls the transform function for each cell in the column with the
   * current text and row index. The return value replaces the cell's content.
   * Useful for formatting values, applying calculations, or normalizing data.
   *
   * @param colIndex - Column index (0-based)
   * @param transform - Function that receives cell text and row index, returns new text
   * @returns This table for chaining
   *
   * @example
   * ```typescript
   * // Uppercase all values in column 0
   * table.mapColumn(0, (text) => text.toUpperCase());
   *
   * // Format numbers in column 2 (skip header row)
   * table.mapColumn(2, (text, rowIndex) =>
   *   rowIndex === 0 ? text : `$${Number(text).toFixed(2)}`
   * );
   *
   * // Add prefix to a column
   * table.mapColumn(1, (text) => `ID-${text}`);
   * ```
   */
  mapColumn(colIndex: number, transform: (text: string, rowIndex: number) => string): this {
    for (let r = 0; r < this.rows.length; r++) {
      const cell = this.rows[r]!.getCell(colIndex);
      if (!cell) continue;

      const currentText = cell.getText();
      const newText = transform(currentText, r);

      if (newText !== currentText) {
        // Clear existing paragraphs and set new text
        const paragraphs = cell.getParagraphs();
        if (paragraphs.length > 0) {
          // Preserve first paragraph's formatting, replace text
          const firstPara = paragraphs[0]!;
          firstPara.setText(newText);
          // Remove extra paragraphs if cell had multiple
          for (let p = paragraphs.length - 1; p >= 1; p--) {
            cell.removeParagraph(p);
          }
        } else {
          cell.createParagraph(newText);
        }
      }
    }

    return this;
  }

  /**
   * Iterates over every cell in the table, providing row and column indices
   *
   * Calls the callback for each cell with its row index, column index,
   * and the cell itself. Useful for bulk operations like formatting or
   * data extraction without manual nested loops.
   *
   * @param callback - Function called for each cell. Return `false` to stop iteration early.
   *
   * @example
   * ```typescript
   * // Apply shading to every other row
   * table.forEachCell((rowIndex, colIndex, cell) => {
   *   if (rowIndex % 2 === 1) {
   *     cell.setShading({ fill: 'F2F2F2', pattern: 'clear' });
   *   }
   * });
   *
   * // Find a cell by content
   * let found: TableCell | undefined;
   * table.forEachCell((row, col, cell) => {
   *   if (cell.getText().includes('Total')) {
   *     found = cell;
   *     return false; // stop early
   *   }
   * });
   * ```
   */
  forEachCell(
    callback: (rowIndex: number, colIndex: number, cell: TableCell) => void | false
  ): void {
    for (let r = 0; r < this.rows.length; r++) {
      const cells = this.rows[r]!.getCells();
      for (let c = 0; c < cells.length; c++) {
        const result = callback(r, c, cells[c]!);
        if (result === false) return;
      }
    }
  }

  /**
   * Finds the first cell matching a predicate, with its coordinates
   *
   * Iterates all cells in row-major order and returns the first one for
   * which the predicate returns true, along with its row and column index.
   *
   * @param predicate - Function that tests each cell
   * @returns Object with row, col, and cell, or undefined if not found
   *
   * @example
   * ```typescript
   * const result = table.findCell((cell) => cell.getText().includes('Total'));
   * if (result) {
   *   console.log(`Found at row ${result.row}, col ${result.col}`);
   * }
   * ```
   */
  findCell(
    predicate: (cell: TableCell, rowIndex: number, colIndex: number) => boolean
  ): { row: number; col: number; cell: TableCell } | undefined {
    for (let r = 0; r < this.rows.length; r++) {
      const cells = this.rows[r]!.getCells();
      for (let c = 0; c < cells.length; c++) {
        if (predicate(cells[c]!, r, c)) {
          return { row: r, col: c, cell: cells[c]! };
        }
      }
    }
    return undefined;
  }

  /**
   * Returns indices of rows matching a predicate
   *
   * Tests each row by passing all its cells to the predicate function.
   * Returns an array of row indices where the predicate returned true.
   *
   * @param predicate - Function that tests a row's cells
   * @returns Array of matching row indices (0-based)
   *
   * @example
   * ```typescript
   * // Find rows where the first cell contains "Total"
   * const totals = table.filterRows((cells) =>
   *   cells[0]?.getText().includes('Total') ?? false
   * );
   *
   * // Find rows where all cells are empty
   * const empty = table.filterRows((cells) =>
   *   cells.every((c) => c.getText().trim() === '')
   * );
   * ```
   */
  filterRows(predicate: (cells: TableCell[], rowIndex: number) => boolean): number[] {
    const indices: number[] = [];
    for (let r = 0; r < this.rows.length; r++) {
      if (predicate(this.rows[r]!.getCells(), r)) {
        indices.push(r);
      }
    }
    return indices;
  }

  /**
   * Removes rows where all cells are empty (no text content)
   *
   * A row is considered empty if every cell's trimmed text is the empty string.
   * Respects the ECMA-376 constraint that at least one row must remain — if all
   * rows are empty, the first row is kept.
   *
   * @returns Number of rows removed
   *
   * @example
   * ```typescript
   * const table = Table.fromArray([
   *   ['Name', 'Age'],
   *   ['', ''],
   *   ['Alice', '30'],
   *   ['', ''],
   * ]);
   * table.removeEmptyRows(); // Returns 2, table now has 2 rows
   * ```
   */
  removeEmptyRows(): number {
    const emptyIndices = this.filterRows((cells) => cells.every((c) => c.getText().trim() === ''));

    // Keep at least one row (ECMA-376 requires >= 1 row)
    const toRemove =
      emptyIndices.length === this.rows.length
        ? emptyIndices.slice(1) // keep first row
        : emptyIndices;

    // Remove in reverse order to preserve indices
    for (let i = toRemove.length - 1; i >= 0; i--) {
      this.removeRow(toRemove[i]!);
    }

    return toRemove.length;
  }

  /**
   * Removes columns where all cells are empty (no text content)
   *
   * A column is considered empty if every cell at that column index has
   * trimmed text equal to the empty string.
   *
   * @returns Number of columns removed
   *
   * @example
   * ```typescript
   * const table = Table.fromArray([
   *   ['Name', '', 'Age'],
   *   ['Alice', '', '30'],
   * ]);
   * table.removeEmptyColumns(); // Returns 1, middle column removed
   * ```
   */
  removeEmptyColumns(): number {
    const colCount = this.getColumnCount();
    let removed = 0;

    // Check columns in reverse order to preserve indices
    for (let c = colCount - 1; c >= 0; c--) {
      const columnCells = this.getColumnCells(c);
      const isEmpty = columnCells.every((cell) => cell.getText().trim() === '');

      if (isEmpty) {
        this.removeColumn(c);
        removed++;
      }
    }

    return removed;
  }

  /**
   * Sets specific widths for table columns
   *
   * Defines the width of each column. Use null for auto-width columns.
   *
   * @param widths - Array of column widths in twips (null = auto width)
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * // First column 2", second column 3", third auto
   * table.setColumnWidths([2880, 4320, null]);
   * ```
   */
  setColumnWidths(widths: (number | null)[]): this {
    // Use tableGrid (the property that toXML() actually reads)
    // Convert null values to undefined for auto-width handling in toXML()
    this.formatting.tableGrid = widths.map((w) => (w === null ? undefined : w)) as number[];
    return this;
  }

  /**
   * Creates a new Table instance
   *
   * Factory method for creating a table with specified dimensions and formatting.
   *
   * @param rows - Number of rows to create
   * @param columns - Number of columns per row
   * @param formatting - Optional table formatting properties
   * @returns New Table instance
   *
   * @example
   * ```typescript
   * const table = Table.create(3, 4);  // 3 rows × 4 columns
   * const styledTable = Table.create(2, 3, {
   *   alignment: 'center',
   *   layout: 'auto'
   * });
   * ```
   */
  static create(rows?: number, columns?: number, formatting?: TableFormatting): Table {
    return new Table(rows, columns, formatting);
  }

  /**
   * Creates a table from a 2D string array
   *
   * Each inner array becomes a row; each string becomes a cell with a single
   * paragraph. The table column count is determined by the longest row.
   * Shorter rows are padded with empty cells to keep the grid rectangular.
   *
   * @param data - 2D array of cell text values
   * @param formatting - Optional table formatting
   * @returns New Table populated with the data
   *
   * @example
   * ```typescript
   * const table = Table.fromArray([
   *   ['Name', 'Age', 'City'],
   *   ['Alice', '30', 'New York'],
   *   ['Bob', '25', 'London'],
   * ]);
   * ```
   *
   * @example
   * ```typescript
   * // With formatting
   * const table = Table.fromArray(
   *   [['Header 1', 'Header 2'], ['Data 1', 'Data 2']],
   *   { alignment: 'center', layout: 'fixed' }
   * );
   * ```
   */
  static fromArray(data: string[][], formatting?: TableFormatting): Table {
    if (data.length === 0) return new Table(0, 0, formatting);

    const maxCols = Math.max(...data.map((row) => row.length));
    const table = new Table(0, 0, formatting);

    for (const rowData of data) {
      const row = new TableRow();
      for (let c = 0; c < maxCols; c++) {
        const text = c < rowData.length ? rowData[c] : undefined;
        row.createCell(text || undefined);
      }
      table.addRow(row);
    }

    return table;
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
    // Validate bounds
    if (startRow < 0 || endRow >= this.rows.length || startCol < 0 || endCol < 0) {
      return this;
    }

    // Validate that end position is >= start position
    if (endRow < startRow || endCol < startCol) {
      return this;
    }

    const cell = this.getCell(startRow, startCol);
    if (!cell) {
      return this;
    }

    // Check for merge conflicts - cells already part of another merge region
    // This prevents undefined behavior when merging overlapping regions
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const checkCell = this.getCell(row, col);
        if (!checkCell) continue;

        // Check if cell is already part of a vertical merge
        const vMerge = checkCell.getVerticalMerge();
        if (vMerge) {
          // Allow if this is the start cell and it's a 'restart' (we'll overwrite it)
          if (row === startRow && col === startCol && vMerge === 'restart') {
            continue;
          }
          // Cell is part of an existing vertical merge - conflict
          throw new Error(
            `Cannot merge cells: Cell at row ${row}, column ${col} is already part of a vertical merge region. ` +
              `Use splitCell() or clear existing merges before creating new merge regions.`
          );
        }

        // Check if cell has a column span > 1 (already part of horizontal merge)
        const colSpan = checkCell.getColumnSpan();
        if (colSpan > 1) {
          // Allow if this is the start cell (we'll overwrite its span)
          if (row === startRow && col === startCol) {
            continue;
          }
          // Cell has existing column span - conflict
          throw new Error(
            `Cannot merge cells: Cell at row ${row}, column ${col} already has a column span of ${colSpan}. ` +
              `Use splitCell() or clear existing merges before creating new merge regions.`
          );
        }
      }
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

    // When tracking, mark absorbed cells with cellMerge
    if (this.trackingContext?.isEnabled()) {
      const author = this.trackingContext.getAuthor();
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          // Skip the start cell
          if (row === startRow && col === startCol) continue;
          const absorbedCell = this.getCell(row, col);
          if (absorbedCell) {
            const revision = Revision.createTableCellMerge(author, []);
            absorbedCell.setCellRevision(revision);
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
      cell.setColumnSpan(1); // Reset to single cell
      cell.setVerticalMerge(undefined); // Clear vertical merge
    }
    return this;
  }

  /**
   * Checks if a merge region would conflict with existing merges
   * @param startRow - Starting row index (0-based)
   * @param startCol - Starting column index (0-based)
   * @param endRow - Ending row index (0-based, inclusive)
   * @param endCol - Ending column index (0-based, inclusive)
   * @returns Object with valid flag and list of conflicts found
   * @example
   * ```typescript
   * const result = table.canMergeCells(0, 0, 2, 2);
   * if (!result.valid) {
   *   console.log('Conflicts:', result.conflicts);
   * }
   * ```
   */
  canMergeCells(
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): { valid: boolean; conflicts: string[] } {
    const conflicts: string[] = [];

    // Validate bounds
    if (
      startRow < 0 ||
      endRow >= this.rows.length ||
      startCol < 0 ||
      endCol < 0 ||
      endRow < startRow ||
      endCol < startCol
    ) {
      conflicts.push('Invalid cell range specified');
      return { valid: false, conflicts };
    }

    // Check each cell in the proposed merge region
    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const checkCell = this.getCell(row, col);
        if (!checkCell) {
          conflicts.push(`Cell at row ${row}, column ${col} does not exist`);
          continue;
        }

        // Check if cell is already part of a vertical merge
        const vMerge = checkCell.getVerticalMerge();
        if (vMerge) {
          // Allow if this is the start cell and it's a 'restart'
          if (!(row === startRow && col === startCol && vMerge === 'restart')) {
            conflicts.push(
              `Cell at row ${row}, column ${col} is part of a vertical merge (${vMerge})`
            );
          }
        }

        // Check if cell has a column span > 1
        const colSpan = checkCell.getColumnSpan();
        if (colSpan > 1) {
          // Allow if this is the start cell
          if (!(row === startRow && col === startCol)) {
            conflicts.push(`Cell at row ${row}, column ${col} has column span of ${colSpan}`);
          }
        }
      }
    }

    return { valid: conflicts.length === 0, conflicts };
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
   * Sets the width of a specific column
   *
   * Defines the width for a single column without affecting others.
   *
   * @param columnIndex - The column index (0-based)
   * @param width - Width in twips (1/20th of a point)
   * @returns This table instance for method chaining
   *
   * @example
   * ```typescript
   * table.setColumnWidth(0, 2880);  // First column = 2 inches
   * table.setColumnWidth(1, 1440);  // Second column = 1 inch
   * ```
   */
  setColumnWidth(columnIndex: number, width: number): this {
    // Use tableGrid (the property that toXML() actually reads)
    const tableGrid = this.formatting.tableGrid || [];
    tableGrid[columnIndex] = width;
    this.formatting.tableGrid = tableGrid;
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
    // Use grid column count (accounting for merged cells) instead of cell count
    const columnCount = this.formatting.tableGrid
      ? this.formatting.tableGrid.length
      : Math.max(...this.rows.map((r) => r.getTotalGridSpan()), 1);

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

    // When tracking enabled, mark each row's cells with cellDel + w:del
    if (this.trackingContext?.isEnabled()) {
      const author = this.trackingContext.getAuthor();
      for (let i = startIndex; i < startIndex + actualCount; i++) {
        const row = this.rows[i]!;
        for (const cell of row.getCells()) {
          const cellDelRevision = Revision.createTableCellDelete(author, []);
          cell.setCellRevision(cellDelRevision);
          for (const para of cell.getParagraphs()) {
            const runs = para.getRuns();
            if (runs.length > 0) {
              const deletion = Revision.createDeletion(author, runs);
              para.addRevision(deletion);
            }
          }
        }
      }
      return true;
    }

    this.rows.splice(startIndex, actualCount);
    return actualCount > 0;
  }

  /**
   * Sorts table rows by the content of a specified column
   *
   * Sorts all rows based on the text content of cells in the given column.
   * By default, excludes the first row (header row) from sorting.
   *
   * @param columnIndex - Column to sort by (0-based)
   * @param options - Sort options
   * @param options.ascending - Sort ascending (default: true)
   * @param options.numeric - Treat values as numbers (default: false, string sort)
   * @param options.skipHeaderRow - Skip first row from sorting (default: true)
   * @returns This table for chaining
   *
   * @example
   * ```typescript
   * // Sort by first column alphabetically
   * table.sortRows(0);
   *
   * // Sort by third column numerically, descending
   * table.sortRows(2, { numeric: true, ascending: false });
   *
   * // Sort all rows including header
   * table.sortRows(0, { skipHeaderRow: false });
   * ```
   */
  sortRows(
    columnIndex: number,
    options?: { ascending?: boolean; numeric?: boolean; skipHeaderRow?: boolean }
  ): this {
    const ascending = options?.ascending ?? true;
    const numeric = options?.numeric ?? false;
    const skipHeaderRow = options?.skipHeaderRow ?? true;

    if (this.rows.length <= 1) {
      return this; // Nothing to sort
    }

    // Determine which rows to sort
    const headerRow = skipHeaderRow ? this.rows.shift() : null;
    const rowsToSort = [...this.rows];

    // Sort the rows
    rowsToSort.sort((a, b) => {
      const cellA = a.getCell(columnIndex);
      const cellB = b.getCell(columnIndex);

      const textA = cellA?.getText().trim() || '';
      const textB = cellB?.getText().trim() || '';

      let comparison: number;
      if (numeric) {
        const numA = parseFloat(textA) || 0;
        const numB = parseFloat(textB) || 0;
        comparison = numA - numB;
      } else {
        comparison = textA.localeCompare(textB);
      }

      return ascending ? comparison : -comparison;
    });

    // Reconstruct rows array
    this.rows = headerRow ? [headerRow, ...rowsToSort] : rowsToSort;

    return this;
  }

  /**
   * Extracts table content as a 2D string array
   *
   * Returns an array of rows, where each row is an array of cell text values.
   * Multi-paragraph cells join their text with newlines. Rows with fewer cells
   * than the widest row are NOT padded — the inner arrays may have different lengths.
   *
   * @returns 2D array of cell text content
   *
   * @example
   * ```typescript
   * const table = new Table(2, 3);
   * table.getCell(0, 0)!.createParagraph('Name');
   * table.getCell(0, 1)!.createParagraph('Age');
   * table.getCell(1, 0)!.createParagraph('Alice');
   * table.getCell(1, 1)!.createParagraph('30');
   *
   * const data = table.toArray();
   * // [['Name', 'Age', ''], ['Alice', '30', '']]
   * ```
   */
  toArray(): string[][] {
    return this.rows.map((row) => row.getCells().map((cell) => cell.getText()));
  }

  /**
   * Converts table content to a plain text representation
   *
   * Renders the table as tab-separated values with rows on separate lines.
   * Useful for logging, debugging, or simple text export.
   *
   * @param columnSeparator - Separator between columns (default: tab)
   * @param rowSeparator - Separator between rows (default: newline)
   * @returns Plain text representation of the table
   *
   * @example
   * ```typescript
   * console.log(table.toPlainText());
   * // "Name\tAge\nAlice\t30"
   *
   * // CSV-style output
   * console.log(table.toPlainText(',', '\n'));
   * // "Name,Age\nAlice,30"
   * ```
   */
  toPlainText(columnSeparator = '\t', rowSeparator = '\n'): string {
    return this.rows
      .map((row) =>
        row
          .getCells()
          .map((cell) => cell.getText())
          .join(columnSeparator)
      )
      .join(rowSeparator);
  }

  /**
   * Exports table content as a CSV string
   *
   * Produces RFC 4180-compliant CSV: fields containing commas, quotes,
   * or newlines are quoted, and embedded quotes are escaped by doubling.
   *
   * @param delimiter - Field delimiter (default: ',')
   * @returns CSV string
   *
   * @example
   * ```typescript
   * const csv = table.toCSV();
   * // "Name,Age,City\nAlice,30,New York\nBob,25,London"
   *
   * // TSV variant
   * const tsv = table.toCSV('\t');
   * ```
   */
  toCSV(delimiter = ','): string {
    return this.rows
      .map((row) =>
        row
          .getCells()
          .map((cell) => {
            const text = cell.getText();
            // Quote field if it contains delimiter, quotes, or newlines
            if (text.includes(delimiter) || text.includes('"') || text.includes('\n')) {
              return '"' + text.replace(/"/g, '""') + '"';
            }
            return text;
          })
          .join(delimiter)
      )
      .join('\n');
  }

  /**
   * Creates a table from a CSV string
   *
   * Parses RFC 4180-compliant CSV: handles quoted fields, escaped quotes
   * (doubled `""`), and newlines within quoted fields.
   *
   * @param csv - CSV string to parse
   * @param delimiter - Field delimiter (default: ',')
   * @param formatting - Optional table formatting
   * @returns New Table populated with the parsed data
   *
   * @example
   * ```typescript
   * const table = Table.fromCSV('Name,Age\nAlice,30\nBob,25');
   *
   * // TSV input
   * const table2 = Table.fromCSV('Name\tAge\nAlice\t30', '\t');
   *
   * // Handles quoted fields
   * const table3 = Table.fromCSV('"City, State",Pop\n"New York, NY",8336817');
   * ```
   */
  static fromCSV(csv: string, delimiter = ',', formatting?: TableFormatting): Table {
    const rows: string[][] = [];
    let currentRow: string[] = [];
    let currentField = '';
    let inQuotes = false;
    let i = 0;

    while (i < csv.length) {
      const ch = csv[i]!;

      if (inQuotes) {
        if (ch === '"') {
          // Check for escaped quote ""
          if (i + 1 < csv.length && csv[i + 1] === '"') {
            currentField += '"';
            i += 2;
          } else {
            // End of quoted field
            inQuotes = false;
            i++;
          }
        } else {
          currentField += ch;
          i++;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
          i++;
        } else if (ch === delimiter) {
          currentRow.push(currentField);
          currentField = '';
          i++;
        } else if (ch === '\n') {
          currentRow.push(currentField);
          currentField = '';
          rows.push(currentRow);
          currentRow = [];
          i++;
        } else if (ch === '\r') {
          // Handle \r\n
          if (i + 1 < csv.length && csv[i + 1] === '\n') {
            i++;
          }
          currentRow.push(currentField);
          currentField = '';
          rows.push(currentRow);
          currentRow = [];
          i++;
        } else {
          currentField += ch;
          i++;
        }
      }
    }

    // Push last field and row
    if (currentField || currentRow.length > 0) {
      currentRow.push(currentField);
      rows.push(currentRow);
    }

    if (rows.length === 0) return new Table(0, 0, formatting);
    return Table.fromArray(rows, formatting);
  }

  /**
   * Creates a new table with rows and columns swapped
   *
   * Returns a new Table where the original rows become columns and columns
   * become rows. Cell text content is preserved; cell formatting from the
   * source is deep-cloned to the transposed position. The original table
   * is not modified.
   *
   * For non-rectangular tables (rows with different cell counts), the
   * result is padded to the longest row's length with empty cells.
   *
   * @returns New Table with transposed data
   *
   * @example
   * ```typescript
   * const table = Table.fromArray([
   *   ['Name', 'Alice', 'Bob'],
   *   ['Age',  '30',    '25'],
   * ]);
   * const transposed = table.transpose();
   * transposed.toArray();
   * // [['Name', 'Age'], ['Alice', '30'], ['Bob', '25']]
   * ```
   */
  transpose(): Table {
    const srcRows = this.rows.length;
    if (srcRows === 0) return new Table(0, 0);

    const srcCols = this.getColumnCount();
    if (srcCols === 0) return new Table(0, 0);

    const transposed = new Table(0, 0);

    for (let c = 0; c < srcCols; c++) {
      const newRow = new TableRow();
      for (let r = 0; r < srcRows; r++) {
        const srcCell = this.getCell(r, c);
        if (srcCell) {
          newRow.addCell(srcCell.clone());
        } else {
          newRow.createCell();
        }
      }
      transposed.addRow(newRow);
    }

    return transposed;
  }

  /**
   * Creates a deep clone of this table
   *
   * Creates a new Table with copies of all rows, cells, content, and formatting.
   * The clone is completely independent of the original.
   *
   * @returns New Table instance with the same structure and content
   *
   * @example
   * ```typescript
   * const original = new Table(2, 3);
   * original.getCell(0, 0)?.addParagraph(new Paragraph().addText('Data'));
   *
   * const copy = original.clone();
   * copy.getCell(0, 0)?.addParagraph(new Paragraph().addText(' - Modified'));
   * // Original table unchanged
   * ```
   */
  clone(): Table {
    const clonedFormatting: TableFormatting = deepClone(this.formatting);
    const clonedTable = new Table(0, 0, clonedFormatting);

    for (const row of this.rows) {
      clonedTable.addRow(row.clone());
    }

    return clonedTable;
  }

  /**
   * Duplicates a row at the given index, inserting copies after it
   *
   * Creates deep clones of the specified row (including all cell content,
   * formatting, and merged cell spans) and inserts them immediately after
   * the source row. Useful for template-based document generation where
   * a styled row needs to be repeated with different data.
   *
   * @param rowIndex - Index of the row to duplicate (0-based)
   * @param count - Number of copies to insert (default 1)
   * @returns Array of the newly inserted rows
   * @throws RangeError if rowIndex is out of bounds
   *
   * @example
   * ```typescript
   * // Create a table with a header and one data row
   * const table = new Table(2, 3);
   * table.getCell(0, 0)?.createParagraph('Name');
   * table.getCell(1, 0)?.createParagraph('Alice');
   *
   * // Duplicate the data row twice for more entries
   * const [row2, row3] = table.duplicateRow(1, 2);
   * row2.getCell(0)?.getParagraphs()[0]?.setText('Bob');
   * row3.getCell(0)?.getParagraphs()[0]?.setText('Charlie');
   * ```
   */
  duplicateRow(rowIndex: number, count = 1): TableRow[] {
    if (rowIndex < 0 || rowIndex >= this.rows.length) {
      throw new RangeError(`Row index ${rowIndex} is out of bounds (0-${this.rows.length - 1})`);
    }
    if (count < 1) return [];

    const sourceRow = this.rows[rowIndex]!;
    const inserted: TableRow[] = [];

    for (let i = 0; i < count; i++) {
      const clonedRow = sourceRow.clone();
      this.rows.splice(rowIndex + 1 + i, 0, clonedRow);
      clonedRow._setParentTable(this);
      inserted.push(clonedRow);
    }

    return inserted;
  }
}
