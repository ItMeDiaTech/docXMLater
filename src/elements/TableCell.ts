/**
 * TableCell - Represents a cell in a table
 */

import { Paragraph } from './Paragraph';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';

/**
 * Cell border style
 */
export type BorderStyle = 'none' | 'single' | 'double' | 'dashed' | 'dotted' | 'thick';

/**
 * Cell border definition
 */
export interface CellBorder {
  style?: BorderStyle;
  size?: number; // Size in eighths of a point
  color?: string; // Hex color without #
}

/**
 * Cell borders
 */
export interface CellBorders {
  top?: CellBorder;
  bottom?: CellBorder;
  left?: CellBorder;
  right?: CellBorder;
}

/**
 * Cell shading/background
 */
export interface CellShading {
  fill?: string; // Background color in hex
  color?: string; // Foreground/pattern color in hex
}

/**
 * Vertical alignment in cell
 */
export type CellVerticalAlignment = 'top' | 'center' | 'bottom';

/**
 * Cell formatting options
 */
export interface CellFormatting {
  width?: number; // Width in twips
  borders?: CellBorders;
  shading?: CellShading;
  verticalAlignment?: CellVerticalAlignment;
  columnSpan?: number; // Number of columns to span
  rowSpan?: number; // Number of rows to span (gridSpan)
}

/**
 * Represents a table cell
 */
export class TableCell {
  private paragraphs: Paragraph[] = [];
  private formatting: CellFormatting;

  /**
   * Creates a new TableCell
   * @param formatting - Cell formatting options
   */
  constructor(formatting: CellFormatting = {}) {
    this.formatting = formatting;
  }

  /**
   * Adds a paragraph to the cell
   * @param paragraph - Paragraph to add
   * @returns This cell for chaining
   */
  addParagraph(paragraph: Paragraph): this {
    this.paragraphs.push(paragraph);
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
   * Gets the text content of all paragraphs
   * @returns Combined text
   */
  getText(): string {
    return this.paragraphs.map(p => p.getText()).join('\n');
  }

  /**
   * Sets cell width
   * @param twips - Width in twips
   * @returns This cell for chaining
   */
  setWidth(twips: number): this {
    this.formatting.width = twips;
    return this;
  }

  /**
   * Sets cell borders
   * @param borders - Border definitions
   * @returns This cell for chaining
   */
  setBorders(borders: CellBorders): this {
    this.formatting.borders = borders;
    return this;
  }

  /**
   * Sets cell shading/background
   * @param shading - Shading definition
   * @returns This cell for chaining
   */
  setShading(shading: CellShading): this {
    this.formatting.shading = shading;
    return this;
  }

  /**
   * Sets vertical alignment
   * @param alignment - Vertical alignment
   * @returns This cell for chaining
   */
  setVerticalAlignment(alignment: CellVerticalAlignment): this {
    this.formatting.verticalAlignment = alignment;
    return this;
  }

  /**
   * Sets column span (merge cells horizontally)
   * @param span - Number of columns to span
   * @returns This cell for chaining
   */
  setColumnSpan(span: number): this {
    this.formatting.columnSpan = span;
    return this;
  }

  /**
   * Gets the cell formatting
   * @returns Cell formatting
   */
  getFormatting(): CellFormatting {
    return { ...this.formatting };
  }

  /**
   * Converts the cell to WordprocessingML XML element
   * @returns XMLElement representing the cell
   */
  toXML(): XMLElement {
    const tcPrChildren: XMLElement[] = [];

    // Add cell width
    if (this.formatting.width !== undefined) {
      tcPrChildren.push(
        XMLBuilder.wSelf('tcW', {
          'w:w': this.formatting.width,
          'w:type': 'dxa',
        })
      );
    }

    // Add cell borders
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

      if (borderElements.length > 0) {
        tcPrChildren.push(XMLBuilder.w('tcBorders', undefined, borderElements));
      }
    }

    // Add shading
    if (this.formatting.shading) {
      const shadingAttrs: Record<string, string> = {
        'w:val': 'clear',
      };

      if (this.formatting.shading.fill) {
        shadingAttrs['w:fill'] = this.formatting.shading.fill;
      }
      if (this.formatting.shading.color) {
        shadingAttrs['w:color'] = this.formatting.shading.color;
      }

      tcPrChildren.push(XMLBuilder.wSelf('shd', shadingAttrs));
    }

    // Add vertical alignment
    if (this.formatting.verticalAlignment) {
      tcPrChildren.push(
        XMLBuilder.wSelf('vAlign', { 'w:val': this.formatting.verticalAlignment })
      );
    }

    // Add column span (gridSpan)
    if (this.formatting.columnSpan && this.formatting.columnSpan > 1) {
      tcPrChildren.push(
        XMLBuilder.wSelf('gridSpan', { 'w:val': this.formatting.columnSpan })
      );
    }

    // Build cell element
    const cellChildren: XMLElement[] = [];

    // Add cell properties if there are any
    if (tcPrChildren.length > 0) {
      cellChildren.push(XMLBuilder.w('tcPr', undefined, tcPrChildren));
    }

    // Add paragraphs (at least one required)
    if (this.paragraphs.length > 0) {
      for (const para of this.paragraphs) {
        cellChildren.push(para.toXML());
      }
    } else {
      // Empty cell needs at least one empty paragraph
      cellChildren.push(new Paragraph().toXML());
    }

    return XMLBuilder.w('tc', undefined, cellChildren);
  }

  /**
   * Creates a border element
   */
  private createBorderElement(side: string, border: CellBorder): XMLElement {
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
   * Creates a new TableCell
   * @param formatting - Cell formatting
   * @returns New TableCell instance
   */
  static create(formatting?: CellFormatting): TableCell {
    return new TableCell(formatting);
  }
}
