/**
 * Run - Represents a run of text with uniform formatting
 * A run is the smallest unit of text formatting in a Word document
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';

/**
 * Text formatting options for a run
 */
export interface RunFormatting {
  /** Bold text */
  bold?: boolean;
  /** Italic text */
  italic?: boolean;
  /** Underline text */
  underline?: boolean | 'single' | 'double' | 'thick' | 'dotted' | 'dash';
  /** Strikethrough text */
  strike?: boolean;
  /** Double strikethrough */
  dstrike?: boolean;
  /** Subscript */
  subscript?: boolean;
  /** Superscript */
  superscript?: boolean;
  /** Font name */
  font?: string;
  /** Font size in points (half-points for Word) */
  size?: number;
  /** Text color in hex format (without #) */
  color?: string;
  /** Highlight color */
  highlight?: 'yellow' | 'green' | 'cyan' | 'magenta' | 'blue' | 'red' | 'darkBlue' | 'darkCyan' | 'darkGreen' | 'darkMagenta' | 'darkRed' | 'darkYellow' | 'darkGray' | 'lightGray' | 'black' | 'white';
  /** Small caps */
  smallCaps?: boolean;
  /** All caps */
  allCaps?: boolean;
}

/**
 * Represents a run of formatted text
 */
export class Run {
  private text: string;
  private formatting: RunFormatting;

  /**
   * Creates a new Run
   * @param text - The text content
   * @param formatting - Formatting options
   */
  constructor(text: string, formatting: RunFormatting = {}) {
    this.text = text;
    this.formatting = formatting;
  }

  /**
   * Gets the text content
   */
  getText(): string {
    return this.text;
  }

  /**
   * Sets the text content
   * @param text - New text content
   */
  setText(text: string): void {
    this.text = text;
  }

  /**
   * Gets the formatting
   */
  getFormatting(): RunFormatting {
    return { ...this.formatting };
  }

  /**
   * Sets bold formatting
   * @param bold - Whether text is bold
   */
  setBold(bold: boolean = true): this {
    this.formatting.bold = bold;
    return this;
  }

  /**
   * Sets italic formatting
   * @param italic - Whether text is italic
   */
  setItalic(italic: boolean = true): this {
    this.formatting.italic = italic;
    return this;
  }

  /**
   * Sets underline formatting
   * @param underline - Underline style or boolean
   */
  setUnderline(underline: RunFormatting['underline'] = true): this {
    this.formatting.underline = underline;
    return this;
  }

  /**
   * Sets strikethrough formatting
   * @param strike - Whether text has strikethrough
   */
  setStrike(strike: boolean = true): this {
    this.formatting.strike = strike;
    return this;
  }

  /**
   * Sets subscript formatting
   * @param subscript - Whether text is subscript
   */
  setSubscript(subscript: boolean = true): this {
    this.formatting.subscript = subscript;
    if (subscript) {
      this.formatting.superscript = false;
    }
    return this;
  }

  /**
   * Sets superscript formatting
   * @param superscript - Whether text is superscript
   */
  setSuperscript(superscript: boolean = true): this {
    this.formatting.superscript = superscript;
    if (superscript) {
      this.formatting.subscript = false;
    }
    return this;
  }

  /**
   * Sets font
   * @param font - Font name
   * @param size - Font size in points (optional)
   */
  setFont(font: string, size?: number): this {
    this.formatting.font = font;
    if (size !== undefined) {
      this.formatting.size = size;
    }
    return this;
  }

  /**
   * Sets font size
   * @param size - Font size in points
   */
  setSize(size: number): this {
    this.formatting.size = size;
    return this;
  }

  /**
   * Sets text color
   * @param color - Color in hex format (with or without #)
   */
  setColor(color: string): this {
    this.formatting.color = color.startsWith('#') ? color.substring(1) : color;
    return this;
  }

  /**
   * Sets highlight color
   * @param highlight - Highlight color
   */
  setHighlight(highlight: RunFormatting['highlight']): this {
    this.formatting.highlight = highlight;
    return this;
  }

  /**
   * Sets small caps
   * @param smallCaps - Whether text is in small caps
   */
  setSmallCaps(smallCaps: boolean = true): this {
    this.formatting.smallCaps = smallCaps;
    return this;
  }

  /**
   * Sets all caps
   * @param allCaps - Whether text is in all caps
   */
  setAllCaps(allCaps: boolean = true): this {
    this.formatting.allCaps = allCaps;
    return this;
  }

  /**
   * Converts the run to WordprocessingML XML element
   * @returns XMLElement representing the run
   */
  toXML(): XMLElement {
    const rPrChildren: XMLElement[] = [];

    // Add formatting elements
    if (this.formatting.bold) {
      rPrChildren.push(XMLBuilder.wSelf('b'));
    }
    if (this.formatting.italic) {
      rPrChildren.push(XMLBuilder.wSelf('i'));
    }
    if (this.formatting.underline) {
      const underlineValue = typeof this.formatting.underline === 'string'
        ? this.formatting.underline
        : 'single';
      rPrChildren.push(XMLBuilder.wSelf('u', { 'w:val': underlineValue }));
    }
    if (this.formatting.strike) {
      rPrChildren.push(XMLBuilder.wSelf('strike'));
    }
    if (this.formatting.dstrike) {
      rPrChildren.push(XMLBuilder.wSelf('dstrike'));
    }
    if (this.formatting.subscript) {
      rPrChildren.push(XMLBuilder.wSelf('vertAlign', { 'w:val': 'subscript' }));
    }
    if (this.formatting.superscript) {
      rPrChildren.push(XMLBuilder.wSelf('vertAlign', { 'w:val': 'superscript' }));
    }
    if (this.formatting.font) {
      rPrChildren.push(XMLBuilder.wSelf('rFonts', {
        'w:ascii': this.formatting.font,
        'w:hAnsi': this.formatting.font,
        'w:cs': this.formatting.font,
      }));
    }
    if (this.formatting.size !== undefined) {
      // Word uses half-points (size * 2)
      const halfPoints = this.formatting.size * 2;
      rPrChildren.push(XMLBuilder.wSelf('sz', { 'w:val': halfPoints }));
      rPrChildren.push(XMLBuilder.wSelf('szCs', { 'w:val': halfPoints }));
    }
    if (this.formatting.color) {
      rPrChildren.push(XMLBuilder.wSelf('color', { 'w:val': this.formatting.color }));
    }
    if (this.formatting.highlight) {
      rPrChildren.push(XMLBuilder.wSelf('highlight', { 'w:val': this.formatting.highlight }));
    }
    if (this.formatting.smallCaps) {
      rPrChildren.push(XMLBuilder.wSelf('smallCaps'));
    }
    if (this.formatting.allCaps) {
      rPrChildren.push(XMLBuilder.wSelf('caps'));
    }

    // Build the run element
    const runChildren: XMLElement[] = [];

    // Add run properties if there are any
    if (rPrChildren.length > 0) {
      runChildren.push(XMLBuilder.w('rPr', undefined, rPrChildren));
    }

    // Add text element
    runChildren.push(XMLBuilder.w('t', {
      'xml:space': 'preserve',
    }, [this.text]));

    return XMLBuilder.w('r', undefined, runChildren);
  }

  /**
   * Creates a new Run with the specified text and formatting
   * @param text - Text content
   * @param formatting - Formatting options
   * @returns New Run instance
   */
  static create(text: string, formatting?: RunFormatting): Run {
    return new Run(text, formatting);
  }
}
