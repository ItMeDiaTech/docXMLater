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
      pPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': formatting.alignment }));
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
}
