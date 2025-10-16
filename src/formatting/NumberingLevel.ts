/**
 * NumberingLevel - Defines formatting for a single level in a list
 *
 * A numbering level specifies how a particular list level (0-8) should be formatted,
 * including the numbering format (bullet, decimal, roman, etc.), text template,
 * alignment, and indentation.
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';

/**
 * Numbering format types supported by Word
 */
export type NumberFormat =
  | 'bullet'        // Bullet character
  | 'decimal'       // 1, 2, 3, ...
  | 'lowerRoman'    // i, ii, iii, ...
  | 'upperRoman'    // I, II, III, ...
  | 'lowerLetter'   // a, b, c, ...
  | 'upperLetter'   // A, B, C, ...
  | 'ordinal'       // 1st, 2nd, 3rd, ...
  | 'cardinalText'  // One, Two, Three, ...
  | 'ordinalText'   // First, Second, Third, ...
  | 'hex'           // 0x01, 0x02, ...
  | 'chicago'       // *, †, ‡, §, ...
  | 'decimal zero'; // 01, 02, 03, ...

/**
 * Alignment for the numbering text
 */
export type NumberAlignment = 'left' | 'center' | 'right' | 'start' | 'end';

/**
 * Properties for creating a numbering level
 */
export interface NumberingLevelProperties {
  /** The level index (0-8, where 0 is the outermost level) */
  level: number;

  /** The numbering format */
  format: NumberFormat;

  /** The text template (e.g., "%1." for decimal, "•" for bullet) */
  text: string;

  /** Alignment of the numbering text */
  alignment?: NumberAlignment;

  /** Starting value (for numeric formats, default: 1) */
  start?: number;

  /** Left indentation in twips */
  leftIndent?: number;

  /** Hanging indentation in twips (for the text after the number) */
  hangingIndent?: number;

  /** Font family for the numbering character (useful for bullets) */
  font?: string;

  /** Font size in half-points (e.g., 22 = 11pt) */
  fontSize?: number;

  /** Whether to show text after the number (default: true) */
  isLegalNumberingStyle?: boolean;

  /** Suffix after the number (tab, space, or nothing) */
  suffix?: 'tab' | 'space' | 'nothing';
}

/**
 * Represents a single level in a numbering definition
 */
export class NumberingLevel {
  private properties: Required<NumberingLevelProperties>;

  /**
   * Creates a new numbering level
   * @param properties The level properties
   */
  constructor(properties: NumberingLevelProperties) {
    // Set defaults
    this.properties = {
      level: properties.level,
      format: properties.format,
      text: properties.text,
      alignment: properties.alignment || 'left',
      start: properties.start !== undefined ? properties.start : 1,
      leftIndent: properties.leftIndent !== undefined ? properties.leftIndent : 720 * (properties.level + 1),
      hangingIndent: properties.hangingIndent !== undefined ? properties.hangingIndent : 360,
      font: properties.font || (properties.format === 'bullet' ? 'Symbol' : 'Calibri'),
      fontSize: properties.fontSize || 22, // 11pt default
      isLegalNumberingStyle: properties.isLegalNumberingStyle !== undefined ? properties.isLegalNumberingStyle : false,
      suffix: properties.suffix || 'tab',
    };

    this.validate();
  }

  /**
   * Validates the level properties
   */
  private validate(): void {
    if (this.properties.level < 0 || this.properties.level > 8) {
      throw new Error(`Level must be between 0 and 8, got ${this.properties.level}`);
    }

    if (this.properties.leftIndent < 0) {
      throw new Error('Left indent must be non-negative');
    }

    if (this.properties.hangingIndent < 0) {
      throw new Error('Hanging indent must be non-negative');
    }

    if (this.properties.start < 0) {
      throw new Error('Start value must be non-negative');
    }
  }

  /**
   * Gets the level index
   */
  getLevel(): number {
    return this.properties.level;
  }

  /**
   * Gets the numbering format
   */
  getFormat(): NumberFormat {
    return this.properties.format;
  }

  /**
   * Gets the level properties
   */
  getProperties(): Required<NumberingLevelProperties> {
    return { ...this.properties };
  }

  /**
   * Sets the left indentation
   * @param twips Indentation in twips
   */
  setLeftIndent(twips: number): this {
    if (twips < 0) {
      throw new Error('Left indent must be non-negative');
    }
    this.properties.leftIndent = twips;
    return this;
  }

  /**
   * Sets the hanging indentation
   * @param twips Indentation in twips
   */
  setHangingIndent(twips: number): this {
    if (twips < 0) {
      throw new Error('Hanging indent must be non-negative');
    }
    this.properties.hangingIndent = twips;
    return this;
  }

  /**
   * Sets the font for the numbering character
   * @param font Font family name
   */
  setFont(font: string): this {
    this.properties.font = font;
    return this;
  }

  /**
   * Sets the alignment
   * @param alignment Alignment type
   */
  setAlignment(alignment: NumberAlignment): this {
    this.properties.alignment = alignment;
    return this;
  }

  /**
   * Generates the WordprocessingML XML for this level
   */
  toXML(): XMLElement {
    const children: XMLElement[] = [];

    // Start value
    children.push(
      XMLBuilder.wSelf('start', { 'w:val': this.properties.start.toString() })
    );

    // Number format
    children.push(
      XMLBuilder.wSelf('numFmt', { 'w:val': this.properties.format })
    );

    // Level text (e.g., "%1." or "•")
    children.push(
      XMLBuilder.wSelf('lvlText', { 'w:val': this.properties.text })
    );

    // Alignment
    children.push(
      XMLBuilder.wSelf('lvlJc', { 'w:val': this.properties.alignment })
    );

    // Paragraph properties (indentation)
    const ind = XMLBuilder.wSelf('ind', {
      'w:left': this.properties.leftIndent.toString(),
      'w:hanging': this.properties.hangingIndent.toString()
    });
    const pPr = XMLBuilder.w('pPr', undefined, [ind]);
    children.push(pPr);

    // Run properties (font)
    const rPrChildren: XMLElement[] = [];

    // Font
    rPrChildren.push(
      XMLBuilder.wSelf('rFonts', {
        'w:ascii': this.properties.font,
        'w:hAnsi': this.properties.font,
        'w:cs': this.properties.font,
        'w:hint': 'default'
      })
    );

    // Font size
    rPrChildren.push(
      XMLBuilder.wSelf('sz', { 'w:val': this.properties.fontSize.toString() })
    );
    rPrChildren.push(
      XMLBuilder.wSelf('szCs', { 'w:val': this.properties.fontSize.toString() })
    );

    const rPr = XMLBuilder.w('rPr', undefined, rPrChildren);
    children.push(rPr);

    // Suffix (what comes after the number)
    if (this.properties.suffix) {
      children.push(
        XMLBuilder.wSelf('suff', { 'w:val': this.properties.suffix })
      );
    }

    // Legal numbering style
    if (this.properties.isLegalNumberingStyle) {
      children.push(XMLBuilder.wSelf('isLgl'));
    }

    return XMLBuilder.w('lvl', { 'w:ilvl': this.properties.level.toString() }, children);
  }

  /**
   * Creates a bullet list level
   * @param level The level index (0-8)
   * @param bullet The bullet character (default: '•')
   */
  static createBulletLevel(level: number, bullet: string = '•'): NumberingLevel {
    return new NumberingLevel({
      level,
      format: 'bullet',
      text: bullet,
      alignment: 'left',
      font: 'Symbol',
      leftIndent: 720 * (level + 1),
      hangingIndent: 360,
    });
  }

  /**
   * Creates a decimal list level (1, 2, 3, ...)
   * @param level The level index (0-8)
   * @param template The text template (default: '%1.')
   */
  static createDecimalLevel(level: number, template: string = `%${level + 1}.`): NumberingLevel {
    return new NumberingLevel({
      level,
      format: 'decimal',
      text: template,
      alignment: 'left',
      leftIndent: 720 * (level + 1),
      hangingIndent: 360,
    });
  }

  /**
   * Creates a lower roman list level (i, ii, iii, ...)
   * @param level The level index (0-8)
   * @param template The text template (default: '%1.')
   */
  static createLowerRomanLevel(level: number, template: string = `%${level + 1}.`): NumberingLevel {
    return new NumberingLevel({
      level,
      format: 'lowerRoman',
      text: template,
      alignment: 'left',
      leftIndent: 720 * (level + 1),
      hangingIndent: 360,
    });
  }

  /**
   * Creates an upper roman list level (I, II, III, ...)
   * @param level The level index (0-8)
   * @param template The text template (default: '%1.')
   */
  static createUpperRomanLevel(level: number, template: string = `%${level + 1}.`): NumberingLevel {
    return new NumberingLevel({
      level,
      format: 'upperRoman',
      text: template,
      alignment: 'left',
      leftIndent: 720 * (level + 1),
      hangingIndent: 360,
    });
  }

  /**
   * Creates a lower letter list level (a, b, c, ...)
   * @param level The level index (0-8)
   * @param template The text template (default: '%1.')
   */
  static createLowerLetterLevel(level: number, template: string = `%${level + 1}.`): NumberingLevel {
    return new NumberingLevel({
      level,
      format: 'lowerLetter',
      text: template,
      alignment: 'left',
      leftIndent: 720 * (level + 1),
      hangingIndent: 360,
    });
  }

  /**
   * Creates an upper letter list level (A, B, C, ...)
   * @param level The level index (0-8)
   * @param template The text template (default: '%1.')
   */
  static createUpperLetterLevel(level: number, template: string = `%${level + 1}.`): NumberingLevel {
    return new NumberingLevel({
      level,
      format: 'upperLetter',
      text: template,
      alignment: 'left',
      leftIndent: 720 * (level + 1),
      hangingIndent: 360,
    });
  }

  /**
   * Factory method for creating a numbering level
   * @param properties The level properties
   */
  static create(properties: NumberingLevelProperties): NumberingLevel {
    return new NumberingLevel(properties);
  }
}
