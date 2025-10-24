/**
 * StylesManager - Manages the collection of styles in a document
 * Handles style registration, retrieval, and styles.xml generation
 */

import { Style, StyleType } from './Style';
import { XMLBuilder } from '../xml/XMLBuilder';
import { XMLParser } from '../xml/XMLParser';

/**
 * Result of XML validation
 */
export interface ValidationResult {
  /** Whether the XML is valid */
  isValid: boolean;
  /** Validation errors if any */
  errors: string[];
  /** Validation warnings if any */
  warnings: string[];
  /** Number of styles found */
  styleCount: number;
  /** List of style IDs found */
  styleIds: string[];
}

/**
 * Manages document styles
 */
export class StylesManager {
  private styles: Map<string, Style> = new Map();
  private includeBuiltInStyles: boolean;

  /**
   * Registry of built-in style factory functions
   * Maps style ID to factory function for lazy loading
   */
  private static readonly BUILT_IN_STYLE_FACTORIES = new Map<string, () => Style>([
    ['Normal', () => Style.createNormalStyle()],
    ['Heading1', () => Style.createHeadingStyle(1)],
    ['Heading2', () => Style.createHeadingStyle(2)],
    ['Heading3', () => Style.createHeadingStyle(3)],
    ['Heading4', () => Style.createHeadingStyle(4)],
    ['Heading5', () => Style.createHeadingStyle(5)],
    ['Heading6', () => Style.createHeadingStyle(6)],
    ['Heading7', () => Style.createHeadingStyle(7)],
    ['Heading8', () => Style.createHeadingStyle(8)],
    ['Heading9', () => Style.createHeadingStyle(9)],
    ['Title', () => Style.createTitleStyle()],
    ['Subtitle', () => Style.createSubtitleStyle()],
    ['ListParagraph', () => Style.createListParagraphStyle()],
    ['TOCHeading', () => Style.createTOCHeadingStyle()],
    ['TableNormal', () => Style.createTableNormalStyle()],
    ['TableGrid', () => Style.createTableGridStyle()],
  ]);

  /**
   * Creates a new StylesManager
   * @param includeBuiltInStyles - Whether to include built-in styles (default: true)
   */
  constructor(includeBuiltInStyles: boolean = true) {
    this.includeBuiltInStyles = includeBuiltInStyles;

    // Always load Normal style if built-in styles are enabled
    // Normal is required and referenced by most other styles
    if (includeBuiltInStyles) {
      this.ensureStyleLoaded('Normal');
    }
  }

  /**
   * Ensures a built-in style is loaded (lazy loading)
   * @param styleId - Style ID to load
   */
  private ensureStyleLoaded(styleId: string): void {
    // Already loaded?
    if (this.styles.has(styleId)) {
      return;
    }

    // Built-in styles disabled?
    if (!this.includeBuiltInStyles) {
      return;
    }

    // Is this a built-in style?
    const factory = StylesManager.BUILT_IN_STYLE_FACTORIES.get(styleId);
    if (factory) {
      this.styles.set(styleId, factory());
    }
  }

  /**
   * Adds a style to the collection
   * @param style - Style to add
   * @returns This manager for chaining
   */
  addStyle(style: Style): this {
    this.styles.set(style.getStyleId(), style);
    return this;
  }

  /**
   * Gets a style by ID
   * Lazy-loads built-in styles on first access
   * @param styleId - Style ID to retrieve
   * @returns The style, or undefined if not found
   */
  getStyle(styleId: string): Style | undefined {
    // Ensure built-in style is loaded if applicable
    this.ensureStyleLoaded(styleId);
    return this.styles.get(styleId);
  }

  /**
   * Checks if a style exists or can be loaded
   * @param styleId - Style ID to check
   * @returns True if the style exists or is a built-in style
   */
  hasStyle(styleId: string): boolean {
    // Check if already loaded
    if (this.styles.has(styleId)) {
      return true;
    }

    // Check if it's a built-in style that can be loaded
    return (
      this.includeBuiltInStyles &&
      StylesManager.BUILT_IN_STYLE_FACTORIES.has(styleId)
    );
  }

  /**
   * Removes a style from the collection
   * @param styleId - Style ID to remove
   * @returns True if the style was removed
   */
  removeStyle(styleId: string): boolean {
    return this.styles.delete(styleId);
  }

  /**
   * Gets all styles
   * @returns Array of all styles
   */
  getAllStyles(): Style[] {
    return Array.from(this.styles.values());
  }

  /**
   * Gets styles by type
   * @param type - Style type to filter by
   * @returns Array of styles matching the type
   */
  getStylesByType(type: StyleType): Style[] {
    return this.getAllStyles().filter(style => style.getType() === type);
  }

  /**
   * Gets quick styles (styles that appear in the style gallery)
   * A style appears in the gallery when qFormat=true AND semiHidden=false
   * @returns Array of quick styles
   */
  getQuickStyles(): Style[] {
    return this.getAllStyles().filter(style => {
      const props = style.getProperties();
      const isQuick = props.qFormat === true || (!props.customStyle && props.qFormat !== false);
      const isVisible = !props.semiHidden;
      return isQuick && isVisible;
    });
  }

  /**
   * Gets visible styles (not semi-hidden)
   * @returns Array of visible styles
   */
  getVisibleStyles(): Style[] {
    return this.getAllStyles().filter(style => {
      const props = style.getProperties();
      return !props.semiHidden;
    });
  }

  /**
   * Gets styles sorted by UI priority
   * Lower priority values appear first (higher importance)
   * Styles without priority appear last
   * @returns Array of styles sorted by priority
   */
  getStylesByPriority(): Style[] {
    return this.getAllStyles().sort((a, b) => {
      const propsA = a.getProperties();
      const propsB = b.getProperties();

      const priorityA = propsA.uiPriority ?? 999;
      const priorityB = propsB.uiPriority ?? 999;

      return priorityA - priorityB;
    });
  }

  /**
   * Gets the linked style for a given style
   * @param styleId - Style ID to find the linked style for
   * @returns The linked style, or undefined if not found
   */
  getLinkedStyle(styleId: string): Style | undefined {
    const style = this.getStyle(styleId);
    if (!style) {
      return undefined;
    }

    const props = style.getProperties();
    if (!props.link) {
      return undefined;
    }

    return this.getStyle(props.link);
  }

  /**
   * Gets all table styles (Phase 5.1)
   * @returns Array of table styles
   */
  getTableStyles(): Style[] {
    return this.getAllStyles().filter(style => style.getType() === 'table');
  }

  /**
   * Creates and adds a table style (Phase 5.1)
   * @param styleId - Style ID
   * @param name - Style name
   * @param basedOn - Base style ID (optional)
   * @returns The created table style
   */
  createTableStyle(styleId: string, name: string, basedOn?: string): Style {
    const style = Style.create({
      styleId,
      name,
      type: 'table',
      basedOn,
      customStyle: true,
    });
    this.addStyle(style);
    return style;
  }

  /**
   * Gets the number of styles
   * @returns Number of styles
   */
  getStyleCount(): number {
    return this.styles.size;
  }

  /**
   * Clears all styles
   * @returns This manager for chaining
   */
  clear(): this {
    this.styles.clear();
    return this;
  }

  /**
   * Gets all available built-in style IDs
   * @returns Array of built-in style IDs
   */
  static getBuiltInStyleIds(): string[] {
    return Array.from(StylesManager.BUILT_IN_STYLE_FACTORIES.keys());
  }

  /**
   * Checks if a style ID is a built-in style
   * @param styleId - Style ID to check
   * @returns True if the style is a built-in style
   */
  static isBuiltInStyle(styleId: string): boolean {
    return StylesManager.BUILT_IN_STYLE_FACTORIES.has(styleId);
  }

  /**
   * Gets statistics about loaded vs available styles
   * @returns Object with style statistics
   */
  getStats(): {
    loadedStyles: number;
    availableBuiltInStyles: number;
    customStyles: number;
  } {
    const loadedStyles = this.styles.size;
    const customStyles = this.getAllStyles().filter(s => s.getProperties().customStyle).length;

    return {
      loadedStyles,
      availableBuiltInStyles: this.includeBuiltInStyles
        ? StylesManager.BUILT_IN_STYLE_FACTORIES.size
        : 0,
      customStyles,
    };
  }

  /**
   * Creates a new paragraph style
   * @param styleId - Unique style ID
   * @param name - Display name
   * @param basedOn - Parent style ID (optional)
   * @returns The created style
   */
  createParagraphStyle(styleId: string, name: string, basedOn?: string): Style {
    const style = Style.create({
      styleId,
      name,
      type: 'paragraph',
      basedOn,
      customStyle: true,
    });
    this.addStyle(style);
    return style;
  }

  /**
   * Creates a new character style
   * @param styleId - Unique style ID
   * @param name - Display name
   * @param basedOn - Parent style ID (optional)
   * @returns The created style
   */
  createCharacterStyle(styleId: string, name: string, basedOn?: string): Style {
    const style = Style.create({
      styleId,
      name,
      type: 'character',
      basedOn,
      customStyle: true,
    });
    this.addStyle(style);
    return style;
  }

  /**
   * Generates the complete styles.xml file
   * @returns XML string for word/styles.xml
   */
  generateStylesXml(): string {
    const builder = new XMLBuilder();

    // Create styles element with namespace
    const stylesChildren = [];

    // Add document defaults
    stylesChildren.push(this.generateDocDefaults());

    // Add all styles
    for (const style of this.getAllStyles()) {
      stylesChildren.push(style.toXML());
    }

    builder.element('w:styles', {
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }, stylesChildren);

    return builder.build(true);
  }

  /**
   * Generates document defaults
   */
  private generateDocDefaults() {
    const rPrDefaultChildren = [
      XMLBuilder.wSelf('rFonts', {
        'w:ascii': 'Calibri',
        'w:hAnsi': 'Calibri',
        'w:eastAsia': 'Calibri',
        'w:cs': 'Calibri',
      }),
      XMLBuilder.wSelf('sz', { 'w:val': '22' }), // 11pt
      XMLBuilder.wSelf('szCs', { 'w:val': '22' }),
      XMLBuilder.wSelf('lang', {
        'w:val': 'en-US',
        'w:eastAsia': 'en-US',
        'w:bidi': 'ar-SA',
      }),
    ];

    const pPrDefaultChildren = [
      XMLBuilder.wSelf('spacing', {
        'w:after': '200',
        'w:line': '276',
        'w:lineRule': 'auto',
      }),
    ];

    return XMLBuilder.w('docDefaults', undefined, [
      XMLBuilder.w('rPrDefault', undefined, [
        XMLBuilder.w('rPr', undefined, rPrDefaultChildren),
      ]),
      XMLBuilder.w('pPrDefault', undefined, [
        XMLBuilder.w('pPr', undefined, pPrDefaultChildren),
      ]),
    ]);
  }

  /**
   * Creates a new StylesManager with built-in styles
   * @returns New StylesManager instance
   */
  static create(): StylesManager {
    return new StylesManager(true);
  }

  /**
   * Creates an empty StylesManager (no built-in styles)
   * @returns New empty StylesManager instance
   */
  static createEmpty(): StylesManager {
    return new StylesManager(false);
  }

  /**
   * Validates styles.xml content for structure and correctness
   *
   * This performs string-based validation to avoid XML parsing corruption.
   * It checks for:
   * - Well-formed XML structure
   * - Required w:styles root element
   * - Valid style definitions
   * - No duplicate style IDs
   * - Required attributes
   *
   * @param xml - The raw styles.xml content to validate
   * @returns ValidationResult with details about validity
   */
  static validate(xml: string): ValidationResult {
    const result: ValidationResult = {
      isValid: true,
      errors: [],
      warnings: [],
      styleCount: 0,
      styleIds: []
    };

    // Check for empty or null
    if (!xml || xml.trim().length === 0) {
      result.isValid = false;
      result.errors.push('Styles XML is empty or null');
      return result;
    }

    // Check for common corruption patterns FIRST (before parsing)
    // This catches double-encoding issues that would break the parser
    if (xml.includes('&lt;w:') || xml.includes('&gt;')) {
      result.isValid = false;
      result.errors.push('XML contains escaped tags - possible double-encoding corruption');
      return result;
    }

    // Skip complex XML structure validation - focus on w:styles specific validation
    // Checking balanced tags with regex is unreliable and can give false positives

    // Use XMLParser to extract root element
    const stylesContent = XMLParser.extractBetweenTags(xml, '<w:styles', '</w:styles>');
    if (!stylesContent) {
      result.isValid = false;
      result.errors.push('Missing required <w:styles> root element');
      return result;
    }

    // Check for namespace declaration
    if (!xml.includes('xmlns:w=')) {
      result.warnings.push('Missing WordprocessingML namespace declaration');
    }

    // Use XMLParser to extract all style elements
    const styleElements = XMLParser.extractElements(stylesContent, 'w:style');
    result.styleCount = styleElements.length;

    // Check if any styles found
    if (styleElements.length === 0) {
      result.warnings.push('No styles found in document');
      return result;
    }

    // Check for styles without attributes (invalid)
    const styleWithoutAttrs = styleElements.filter(el => {
      // Check if element has any attributes
      const openTagEnd = el.indexOf('>');
      const openTag = el.substring(0, openTagEnd);
      return !openTag.includes('w:type') || !openTag.includes('w:styleId');
    });

    if (styleWithoutAttrs.length > 0) {
      result.isValid = false;
      result.errors.push('Style found without any attributes - w:type and w:styleId are required');
    }

    // Process each style element
    const foundStyleIds = new Set<string>();

    for (const styleElement of styleElements) {
      // Extract styleId using XMLParser
      const styleId = XMLParser.extractAttribute(styleElement, 'w:styleId');
      if (styleId) {
        // Check for duplicates
        if (foundStyleIds.has(styleId)) {
          result.isValid = false;
          result.errors.push(`Duplicate style ID found: "${styleId}"`);
        } else {
          foundStyleIds.add(styleId);
          result.styleIds.push(styleId);
        }
      } else {
        result.isValid = false;
        result.errors.push('Style found without required w:styleId attribute');
      }

      // Extract and validate type using XMLParser
      const type = XMLParser.extractAttribute(styleElement, 'w:type');
      if (type) {
        if (!['paragraph', 'character', 'table', 'numbering'].includes(type)) {
          result.warnings.push(`Invalid style type: "${type}"`);
        }
      } else {
        result.isValid = false;
        result.errors.push('Style found without required w:type attribute');
      }

      // Check for circular references - extract basedOn value
      const basedOnElement = XMLParser.extractElements(styleElement, 'w:basedOn')[0];
      if (basedOnElement && styleId) {
        const basedOn = XMLParser.extractAttribute(basedOnElement, 'w:val');
        if (basedOn && styleId === basedOn) {
          result.isValid = false;
          result.errors.push(`Circular reference detected: style "${styleId}" based on itself`);
        }
      }
    }

    // Check for required Normal style
    if (!foundStyleIds.has('Normal')) {
      result.warnings.push('Missing "Normal" style - document may not render correctly');
    }

    // Check for BOM or invalid characters
    if (xml.charCodeAt(0) === 0xFEFF) {
      result.warnings.push('XML contains BOM (Byte Order Mark) - may cause parsing issues');
    }

    // Summary
    if (result.styleCount === 0) {
      result.warnings.push('No styles found in document');
    }

    return result;
  }
}
