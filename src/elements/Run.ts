/**
 * Run - Represents a run of text with uniform formatting
 * A run is the smallest unit of text formatting in a Word document
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { validateRunText } from '../utils/validation';

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
  /**
   * Automatically clean XML-like patterns from text content.
   * When true (default), removes XML tags like <w:t> from text to prevent display issues.
   * Set to false to disable auto-cleaning (useful for debugging).
   * Default: true (auto-clean enabled by default for defensive data handling)
   */
  cleanXmlFromText?: boolean;
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
    // Default to auto-cleaning XML patterns unless explicitly disabled
    const shouldClean = formatting.cleanXmlFromText !== false;

    // Validate text for XML patterns
    const validation = validateRunText(text, {
      context: 'Run constructor',
      autoClean: shouldClean,
      warnToConsole: false,  // Silent by default - team expects dirty data
    });

    // Use cleaned text if available and cleaning was requested
    this.text = validation.cleanedText || text;

    // Remove cleanXmlFromText from formatting as it's not a display property
    const { cleanXmlFromText, ...displayFormatting } = formatting;
    this.formatting = displayFormatting;
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
    // Respect original cleanXmlFromText setting (Issue #9 fix)
    // This ensures consistent behavior with constructor
    const shouldClean = this.formatting.cleanXmlFromText !== false;

    // Validate text for XML patterns
    const validation = validateRunText(text, {
      context: 'Run.setText',
      autoClean: shouldClean,
      warnToConsole: false,  // Silent by default
    });

    // Use cleaned text if available and cleaning was requested
    this.text = validation.cleanedText || text;
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
   * Sets text color with normalization to uppercase hex
   * @param color - Color in hex format (with or without #)
   * @throws Error if color format is invalid
   */
  setColor(color: string): this {
    this.formatting.color = this.normalizeColor(color);
    return this;
  }

  /**
   * Normalizes color to uppercase 6-character hex format per Microsoft convention
   * @param color - Color in hex format (with or without #)
   * @returns Normalized uppercase hex color
   * @throws Error if color format is invalid
   */
  private normalizeColor(color: string): string {
    // Remove # if present
    const hex = color.replace(/^#/, '');

    // Validate format: must be 3 or 6 character hex
    if (!/^[0-9A-Fa-f]{3}$|^[0-9A-Fa-f]{6}$/.test(hex)) {
      throw new Error(
        `Invalid color format: "${color}". Expected 3 or 6-character hex ` +
        `(e.g., "FF0000", "#FF0000", "F00", or "#F00")`
      );
    }

    // Expand 3-char to 6-char format and normalize to uppercase
    if (hex.length === 3) {
      return (hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2)).toUpperCase();
    }

    return hex.toUpperCase();
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
   *
   * **ECMA-376 Compliance:** Properties are generated in the order specified by
   * ECMA-376 Part 1 §17.3.2.28 to ensure strict OpenXML conformance.
   *
   * Per spec, the order is:
   * 1. rFonts (font family)
   * 2. b (bold)
   * 3. i (italic)
   * 4. caps/smallCaps (capitalization)
   * 5. strike/dstrike (strikethrough)
   * 6. u (underline)
   * 7. sz/szCs (font size)
   * 8. color (text color)
   * 9. highlight (highlight color)
   * 10. vertAlign (subscript/superscript)
   *
   * @returns XMLElement representing the run
   */
  toXML(): XMLElement {
    // Validate text content before serialization
    if (this.text === undefined || this.text === null) {
      console.warn(
        'DocXML Warning: Run has undefined/null text content - using empty string. ' +
        'This may indicate a bug in your code.'
      );
      this.text = '';
    }

    const rPrChildren: XMLElement[] = [];

    // 1. Font family (must be first per ECMA-376 §17.3.2.28)
    if (this.formatting.font) {
      rPrChildren.push(XMLBuilder.wSelf('rFonts', {
        'w:ascii': this.formatting.font,
        'w:hAnsi': this.formatting.font,
        'w:cs': this.formatting.font,
      }));
    }

    // 2. Bold
    if (this.formatting.bold) {
      rPrChildren.push(XMLBuilder.wSelf('b'));
    }

    // 3. Italic
    if (this.formatting.italic) {
      rPrChildren.push(XMLBuilder.wSelf('i'));
    }

    // 4. Capitalization (caps/smallCaps)
    if (this.formatting.allCaps) {
      rPrChildren.push(XMLBuilder.wSelf('caps'));
    }
    if (this.formatting.smallCaps) {
      rPrChildren.push(XMLBuilder.wSelf('smallCaps'));
    }

    // 5. Strikethrough
    if (this.formatting.strike) {
      rPrChildren.push(XMLBuilder.wSelf('strike'));
    }
    if (this.formatting.dstrike) {
      rPrChildren.push(XMLBuilder.wSelf('dstrike'));
    }

    // 6. Underline
    if (this.formatting.underline) {
      const underlineValue = typeof this.formatting.underline === 'string'
        ? this.formatting.underline
        : 'single';
      rPrChildren.push(XMLBuilder.wSelf('u', { 'w:val': underlineValue }));
    }

    // 7. Font size
    if (this.formatting.size !== undefined) {
      // Word uses half-points (size * 2)
      const halfPoints = this.formatting.size * 2;
      rPrChildren.push(XMLBuilder.wSelf('sz', { 'w:val': halfPoints }));
      rPrChildren.push(XMLBuilder.wSelf('szCs', { 'w:val': halfPoints }));
    }

    // 8. Text color
    if (this.formatting.color) {
      rPrChildren.push(XMLBuilder.wSelf('color', { 'w:val': this.formatting.color }));
    }

    // 9. Highlight color
    if (this.formatting.highlight) {
      rPrChildren.push(XMLBuilder.wSelf('highlight', { 'w:val': this.formatting.highlight }));
    }

    // 10. Vertical alignment (subscript/superscript) - must be last
    if (this.formatting.subscript) {
      rPrChildren.push(XMLBuilder.wSelf('vertAlign', { 'w:val': 'subscript' }));
    }
    if (this.formatting.superscript) {
      rPrChildren.push(XMLBuilder.wSelf('vertAlign', { 'w:val': 'superscript' }));
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
   * Checks if the run has non-empty text content
   * @returns True if the run has text with length > 0
   */
  hasText(): boolean {
    return this.text !== undefined &&
           this.text !== null &&
           this.text.length > 0;
  }

  /**
   * Checks if the run has any formatting applied
   * @returns True if any formatting properties are set
   */
  hasFormatting(): boolean {
    return Object.keys(this.formatting).length > 0;
  }

  /**
   * Checks if the run is valid (has either text or formatting)
   * An empty run with no formatting is considered invalid
   * @returns True if the run has text or formatting
   */
  isValid(): boolean {
    return this.hasText() || this.hasFormatting();
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

  /**
   * Creates a deep clone of this run
   * @returns New Run instance with copied text and formatting
   * @example
   * ```typescript
   * const original = new Run('Hello', { bold: true });
   * const copy = original.clone();
   * copy.setText('World');  // Original unchanged
   * ```
   */
  clone(): Run {
    // Deep copy formatting to avoid shared references
    const clonedFormatting: RunFormatting = JSON.parse(JSON.stringify(this.formatting));
    return new Run(this.text, clonedFormatting);
  }

  /**
   * Inserts text at a specific position
   * @param index - Position to insert at (0-based)
   * @param text - Text to insert
   * @returns This run for chaining
   * @example
   * ```typescript
   * const run = new Run('Hello World');
   * run.insertText(6, 'Beautiful ');  // "Hello Beautiful World"
   * ```
   */
  insertText(index: number, text: string): this {
    if (index < 0) index = 0;
    if (index > this.text.length) index = this.text.length;

    this.text = this.text.slice(0, index) + text + this.text.slice(index);
    return this;
  }

  /**
   * Appends text to the end of the run
   * @param text - Text to append
   * @returns This run for chaining
   * @example
   * ```typescript
   * const run = new Run('Hello');
   * run.appendText(' World');  // "Hello World"
   * ```
   */
  appendText(text: string): this {
    this.text += text;
    return this;
  }

  /**
   * Replaces text in a specific range
   * @param start - Start position (0-based, inclusive)
   * @param end - End position (0-based, exclusive)
   * @param text - Replacement text
   * @returns This run for chaining
   * @example
   * ```typescript
   * const run = new Run('Hello World');
   * run.replaceText(0, 5, 'Hi');  // "Hi World"
   * ```
   */
  replaceText(start: number, end: number, text: string): this {
    if (start < 0) start = 0;
    if (end > this.text.length) end = this.text.length;
    if (start > end) [start, end] = [end, start];  // Swap if reversed

    this.text = this.text.slice(0, start) + text + this.text.slice(end);
    return this;
  }
}
