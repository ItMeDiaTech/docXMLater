/**
 * Style - Represents a style definition in a Word document
 * Supports paragraph, character, table, and numbering styles
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { ParagraphFormatting } from '../elements/Paragraph';
import { RunFormatting } from '../elements/Run';

/**
 * Style type
 */
export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

/**
 * Style properties
 */
export interface StyleProperties {
  /** Unique style identifier */
  styleId: string;
  /** Display name */
  name: string;
  /** Style type */
  type: StyleType;
  /** Parent style ID for inheritance */
  basedOn?: string;
  /** Next style ID (auto-next paragraph style) */
  next?: string;
  /** Whether this is a default style */
  isDefault?: boolean;
  /** Whether this is a custom style */
  customStyle?: boolean;
  /** Paragraph formatting (for paragraph and table styles) */
  paragraphFormatting?: ParagraphFormatting;
  /** Run formatting (for character and paragraph styles) */
  runFormatting?: RunFormatting;
}

/**
 * Represents a style definition
 */
export class Style {
  private properties: StyleProperties;

  /**
   * Creates a new Style
   * @param properties - Style properties
   */
  constructor(properties: StyleProperties) {
    this.properties = { ...properties };
  }

  /**
   * Gets the style ID
   * @returns Style ID
   */
  getStyleId(): string {
    return this.properties.styleId;
  }

  /**
   * Gets the style name
   * @returns Style name
   */
  getName(): string {
    return this.properties.name;
  }

  /**
   * Gets the style type
   * @returns Style type
   */
  getType(): StyleType {
    return this.properties.type;
  }

  /**
   * Gets all style properties
   * @returns Style properties
   */
  getProperties(): StyleProperties {
    return { ...this.properties };
  }

  /**
   * Sets the base style
   * @param styleId - Parent style ID
   * @returns This style for chaining
   */
  setBasedOn(styleId: string): this {
    this.properties.basedOn = styleId;
    return this;
  }

  /**
   * Sets the next style
   * @param styleId - Next style ID
   * @returns This style for chaining
   */
  setNext(styleId: string): this {
    this.properties.next = styleId;
    return this;
  }

  /**
   * Sets paragraph formatting
   * @param formatting - Paragraph formatting options
   * @returns This style for chaining
   */
  setParagraphFormatting(formatting: ParagraphFormatting): this {
    this.properties.paragraphFormatting = { ...formatting };
    return this;
  }

  /**
   * Sets run formatting
   * @param formatting - Run formatting options
   * @returns This style for chaining
   */
  setRunFormatting(formatting: RunFormatting): this {
    this.properties.runFormatting = { ...formatting };
    return this;
  }

  /**
   * Validates that this style definition is valid
   *
   * Checks:
   * - Required fields (styleId, name, type)
   * - Valid type value
   * - No circular references (basedOn != styleId)
   * - Valid formatting values
   *
   * @returns True if the style is valid, false otherwise
   */
  isValid(): boolean {
    try {
      // Required fields
      if (!this.properties.styleId || !this.properties.name || !this.properties.type) {
        return false;
      }

      // Valid type
      const validTypes: StyleType[] = ['paragraph', 'character', 'table', 'numbering'];
      if (!validTypes.includes(this.properties.type)) {
        return false;
      }

      // No circular reference
      if (this.properties.basedOn === this.properties.styleId) {
        return false;
      }

      // Check paragraph formatting if present
      if (this.properties.paragraphFormatting) {
        const pf = this.properties.paragraphFormatting;

        // Check alignment
        if (pf.alignment) {
          const validAlignments = ['left', 'center', 'right', 'justify', 'both', 'distribute'];
          if (!validAlignments.includes(pf.alignment)) {
            return false;
          }
        }

        // Check spacing values
        if (pf.spacing) {
          const spacing = pf.spacing;
          if (spacing.before !== undefined && spacing.before < 0) return false;
          if (spacing.after !== undefined && spacing.after < 0) return false;
          if (spacing.line !== undefined && spacing.line < 0) return false;
          if (spacing.lineRule && !['auto', 'exact', 'atLeast'].includes(spacing.lineRule)) {
            return false;
          }
        }

        // Check indentation values
        if (pf.indentation) {
          const ind = pf.indentation;
          // Indentation values can be negative for hanging indent
          if (ind.left !== undefined && ind.left < -100000) return false;
          if (ind.right !== undefined && ind.right < -100000) return false;
        }
      }

      // Check run formatting if present
      if (this.properties.runFormatting) {
        const rf = this.properties.runFormatting;

        // Check font size
        if (rf.size !== undefined && (rf.size <= 0 || rf.size > 1638)) {
          return false; // Max font size in Word is 1638
        }

        // Check color format (should be 6 hex characters)
        if (rf.color && !/^[0-9A-Fa-f]{6}$/.test(rf.color)) {
          return false;
        }

        // Check highlight color
        if (rf.highlight) {
          const validHighlights = [
            'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray',
            'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green',
            'lightGray', 'magenta', 'none', 'red', 'white', 'yellow'
          ];
          if (!validHighlights.includes(rf.highlight)) {
            return false;
          }
        }
      }

      return true;
    } catch {
      return false;
    }
  }

  /**
   * Converts the style to WordprocessingML XML element
   * @returns XMLElement representing the style
   */
  toXML(): XMLElement {
    const styleAttrs: Record<string, string> = {
      'w:type': this.properties.type,
      'w:styleId': this.properties.styleId,
    };

    if (this.properties.isDefault) {
      styleAttrs['w:default'] = '1';
    }

    if (this.properties.customStyle) {
      styleAttrs['w:customStyle'] = '1';
    }

    const styleChildren: XMLElement[] = [];

    // Add style name
    styleChildren.push(XMLBuilder.wSelf('name', { 'w:val': this.properties.name }));

    // Add basedOn
    if (this.properties.basedOn) {
      styleChildren.push(XMLBuilder.wSelf('basedOn', { 'w:val': this.properties.basedOn }));
    }

    // Add next
    if (this.properties.next) {
      styleChildren.push(XMLBuilder.wSelf('next', { 'w:val': this.properties.next }));
    }

    // Add qFormat for built-in styles (makes them appear in quick styles)
    if (!this.properties.customStyle) {
      styleChildren.push(XMLBuilder.wSelf('qFormat'));
    }

    // Add paragraph properties
    if (this.properties.paragraphFormatting) {
      const pPr = this.generateParagraphProperties(this.properties.paragraphFormatting);
      if (pPr.children && pPr.children.length > 0) {
        styleChildren.push(pPr);
      }
    }

    // Add run properties
    if (this.properties.runFormatting) {
      const rPr = this.generateRunProperties(this.properties.runFormatting);
      if (rPr.children && rPr.children.length > 0) {
        styleChildren.push(rPr);
      }
    }

    return XMLBuilder.w('style', styleAttrs, styleChildren);
  }

  /**
   * Generates paragraph properties XML
   */
  private generateParagraphProperties(formatting: ParagraphFormatting): XMLElement {
    const pPrChildren: XMLElement[] = [];

    // Add alignment
    if (formatting.alignment) {
      // Map 'justify' to 'both' per ECMA-376 (Word uses 'both' for justified text)
      const alignmentValue = formatting.alignment === 'justify' ? 'both' : formatting.alignment;
      pPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': alignmentValue }));
    }

    // Add indentation
    if (formatting.indentation) {
      const ind = formatting.indentation;
      const attributes: Record<string, number> = {};
      if (ind.left !== undefined) attributes['w:left'] = ind.left;
      if (ind.right !== undefined) attributes['w:right'] = ind.right;
      if (ind.firstLine !== undefined) attributes['w:firstLine'] = ind.firstLine;
      if (ind.hanging !== undefined) attributes['w:hanging'] = ind.hanging;
      if (Object.keys(attributes).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('ind', attributes));
      }
    }

    // Add spacing
    if (formatting.spacing) {
      const spc = formatting.spacing;
      const attributes: Record<string, number | string> = {};
      if (spc.before !== undefined) attributes['w:before'] = spc.before;
      if (spc.after !== undefined) attributes['w:after'] = spc.after;
      if (spc.line !== undefined) attributes['w:line'] = spc.line;
      if (spc.lineRule) attributes['w:lineRule'] = spc.lineRule;
      if (Object.keys(attributes).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('spacing', attributes));
      }
    }

    // Add other properties
    if (formatting.keepNext) {
      pPrChildren.push(XMLBuilder.wSelf('keepNext'));
    }
    if (formatting.keepLines) {
      pPrChildren.push(XMLBuilder.wSelf('keepLines'));
    }
    if (formatting.pageBreakBefore) {
      pPrChildren.push(XMLBuilder.wSelf('pageBreakBefore'));
    }

    return XMLBuilder.w('pPr', undefined, pPrChildren);
  }

  /**
   * Generates run properties XML
   */
  private generateRunProperties(formatting: RunFormatting): XMLElement {
    const rPrChildren: XMLElement[] = [];

    // Add formatting elements
    if (formatting.bold) {
      rPrChildren.push(XMLBuilder.wSelf('b'));
    }
    if (formatting.italic) {
      rPrChildren.push(XMLBuilder.wSelf('i'));
    }
    if (formatting.underline) {
      const underlineValue = typeof formatting.underline === 'string'
        ? formatting.underline
        : 'single';
      rPrChildren.push(XMLBuilder.wSelf('u', { 'w:val': underlineValue }));
    }
    if (formatting.strike) {
      rPrChildren.push(XMLBuilder.wSelf('strike'));
    }
    if (formatting.dstrike) {
      rPrChildren.push(XMLBuilder.wSelf('dstrike'));
    }
    if (formatting.subscript) {
      rPrChildren.push(XMLBuilder.wSelf('vertAlign', { 'w:val': 'subscript' }));
    }
    if (formatting.superscript) {
      rPrChildren.push(XMLBuilder.wSelf('vertAlign', { 'w:val': 'superscript' }));
    }
    if (formatting.font) {
      rPrChildren.push(XMLBuilder.wSelf('rFonts', {
        'w:ascii': formatting.font,
        'w:hAnsi': formatting.font,
        'w:cs': formatting.font,
      }));
    }
    if (formatting.size !== undefined) {
      // Word uses half-points (size * 2)
      const halfPoints = formatting.size * 2;
      rPrChildren.push(XMLBuilder.wSelf('sz', { 'w:val': halfPoints }));
      rPrChildren.push(XMLBuilder.wSelf('szCs', { 'w:val': halfPoints }));
    }
    if (formatting.color) {
      rPrChildren.push(XMLBuilder.wSelf('color', { 'w:val': formatting.color }));
    }
    if (formatting.highlight) {
      rPrChildren.push(XMLBuilder.wSelf('highlight', { 'w:val': formatting.highlight }));
    }
    if (formatting.smallCaps) {
      rPrChildren.push(XMLBuilder.wSelf('smallCaps'));
    }
    if (formatting.allCaps) {
      rPrChildren.push(XMLBuilder.wSelf('caps'));
    }

    return XMLBuilder.w('rPr', undefined, rPrChildren);
  }

  /**
   * Creates a new Style
   * @param properties - Style properties
   * @returns New Style instance
   */
  static create(properties: StyleProperties): Style {
    return new Style(properties);
  }

  /**
   * Creates the Normal style (default paragraph style)
   * @returns Normal style
   */
  static createNormalStyle(): Style {
    return new Style({
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      next: 'Normal',
      paragraphFormatting: {
        spacing: {
          after: 200,
          line: 276,
          lineRule: 'auto',
        },
      },
      runFormatting: {
        font: 'Calibri',
        size: 11,
      },
    });
  }

  /**
   * Creates a Heading style
   * @param level - Heading level (1-9)
   * @returns Heading style
   */
  static createHeadingStyle(level: number): Style {
    if (level < 1 || level > 9) {
      throw new Error('Heading level must be between 1 and 9');
    }

    const sizes = [16, 13, 12, 11, 11, 11, 11, 11, 11]; // Font sizes for Heading1-9
    const bold = level <= 4; // Headings 1-4 are bold

    return new Style({
      styleId: `Heading${level}`,
      name: `Heading ${level}`,
      type: 'paragraph',
      basedOn: 'Normal',
      next: 'Normal',
      paragraphFormatting: {
        spacing: {
          before: level === 1 ? 240 : 120,
          after: 120,
        },
        keepNext: true,
        keepLines: true,
      },
      runFormatting: {
        font: 'Calibri Light',
        size: sizes[level - 1],
        bold: bold,
        color: level === 1 ? '2E74B5' : '1F4D78',
      },
    });
  }

  /**
   * Creates the Title style
   * @returns Title style
   */
  static createTitleStyle(): Style {
    return new Style({
      styleId: 'Title',
      name: 'Title',
      type: 'paragraph',
      basedOn: 'Normal',
      next: 'Normal',
      paragraphFormatting: {
        spacing: {
          after: 120,
        },
      },
      runFormatting: {
        font: 'Calibri Light',
        size: 28,
        color: '2E74B5',
      },
    });
  }

  /**
   * Creates the Subtitle style
   * @returns Subtitle style
   */
  static createSubtitleStyle(): Style {
    return new Style({
      styleId: 'Subtitle',
      name: 'Subtitle',
      type: 'paragraph',
      basedOn: 'Normal',
      next: 'Normal',
      paragraphFormatting: {
        spacing: {
          after: 120,
        },
      },
      runFormatting: {
        font: 'Calibri Light',
        size: 14,
        color: '595959',
        italic: true,
      },
    });
  }

  /**
   * Creates a List Paragraph style (for lists)
   * @returns List Paragraph style
   */
  static createListParagraphStyle(): Style {
    return new Style({
      styleId: 'ListParagraph',
      name: 'List Paragraph',
      type: 'paragraph',
      basedOn: 'Normal',
      next: 'ListParagraph',
      paragraphFormatting: {
        indentation: {
          left: 720, // 0.5 inch
        },
      },
    });
  }

  /**
   * Creates a TOC Heading style (for table of contents titles)
   * @returns TOC Heading style
   */
  static createTOCHeadingStyle(): Style {
    return new Style({
      styleId: 'TOCHeading',
      name: 'TOC Heading',
      type: 'paragraph',
      basedOn: 'Heading1',
      next: 'Normal',
      runFormatting: {
        bold: true,
        font: 'Calibri',
        size: 14,
        color: '000000', // Black (different from Heading1's blue)
      },
      paragraphFormatting: {
        spacing: {
          before: 480, // Larger spacing before TOC
          after: 240,
        },
      },
    });
  }

  /**
   * Creates a deep clone of this style
   * @returns New Style instance with copied properties
   * @example
   * ```typescript
   * const original = Style.createHeadingStyle(1);
   * const copy = original.clone();
   * copy.setRunFormatting({ color: 'FF0000' });  // Doesn't affect original
   * ```
   */
  clone(): Style {
    // Deep copy all properties
    const clonedProps: StyleProperties = JSON.parse(JSON.stringify(this.properties));
    return new Style(clonedProps);
  }

  /**
   * Merges properties from another style into this one
   * @param otherStyle - Style to merge from
   * @returns This style for chaining
   * @example
   * ```typescript
   * const base = Style.createNormalStyle();
   * const override = Style.create({
   *   styleId: 'Override',
   *   name: 'Override',
   *   type: 'paragraph',
   *   runFormatting: { bold: true, color: 'FF0000' }
   * });
   * base.mergeWith(override);  // base now has bold red text
   * ```
   */
  mergeWith(otherStyle: Style): this {
    const otherProps = otherStyle.getProperties();

    // Merge paragraph formatting
    if (otherProps.paragraphFormatting) {
      if (!this.properties.paragraphFormatting) {
        this.properties.paragraphFormatting = {};
      }

      // Merge top-level paragraph properties
      if (otherProps.paragraphFormatting.alignment) {
        this.properties.paragraphFormatting.alignment = otherProps.paragraphFormatting.alignment;
      }
      if (otherProps.paragraphFormatting.keepNext !== undefined) {
        this.properties.paragraphFormatting.keepNext = otherProps.paragraphFormatting.keepNext;
      }
      if (otherProps.paragraphFormatting.keepLines !== undefined) {
        this.properties.paragraphFormatting.keepLines = otherProps.paragraphFormatting.keepLines;
      }
      if (otherProps.paragraphFormatting.pageBreakBefore !== undefined) {
        this.properties.paragraphFormatting.pageBreakBefore = otherProps.paragraphFormatting.pageBreakBefore;
      }

      // Merge indentation
      if (otherProps.paragraphFormatting.indentation) {
        if (!this.properties.paragraphFormatting.indentation) {
          this.properties.paragraphFormatting.indentation = {};
        }
        Object.assign(this.properties.paragraphFormatting.indentation, otherProps.paragraphFormatting.indentation);
      }

      // Merge spacing
      if (otherProps.paragraphFormatting.spacing) {
        if (!this.properties.paragraphFormatting.spacing) {
          this.properties.paragraphFormatting.spacing = {};
        }
        Object.assign(this.properties.paragraphFormatting.spacing, otherProps.paragraphFormatting.spacing);
      }
    }

    // Merge run formatting
    if (otherProps.runFormatting) {
      if (!this.properties.runFormatting) {
        this.properties.runFormatting = {};
      }
      Object.assign(this.properties.runFormatting, otherProps.runFormatting);
    }

    // Merge other properties (but don't override styleId)
    if (otherProps.name) this.properties.name = otherProps.name;
    if (otherProps.basedOn) this.properties.basedOn = otherProps.basedOn;
    if (otherProps.next) this.properties.next = otherProps.next;

    return this;
  }
}
