/**
 * XMLBuilder - Utility for building XML content
 * Provides a simple fluent API for generating WordprocessingML XML
 */

/**
 * Represents an XML element with attributes and children
 */
export interface XMLElement {
  name: string;
  attributes?: Record<string, string | number | boolean | undefined>;
  children?: (XMLElement | string)[];
  selfClosing?: boolean;
}

/**
 * XML Builder for creating WordprocessingML XML
 */
export class XMLBuilder {
  private elements: (XMLElement | string)[] = [];

  /**
   * Adds an element to the builder
   * @param name - Element name (with namespace prefix if needed)
   * @param attributes - Element attributes
   * @param children - Child elements or text
   * @returns This builder for chaining
   */
  element(
    name: string,
    attributes?: Record<string, string | number | boolean | undefined>,
    children?: (XMLElement | string)[]
  ): XMLBuilder {
    this.elements.push({
      name,
      attributes,
      children,
    });
    return this;
  }

  /**
   * Adds a self-closing element
   * @param name - Element name
   * @param attributes - Element attributes
   * @returns This builder for chaining
   */
  selfClosingElement(
    name: string,
    attributes?: Record<string, string | number | boolean | undefined>
  ): XMLBuilder {
    this.elements.push({
      name,
      attributes,
      selfClosing: true,
    });
    return this;
  }

  /**
   * Adds text content
   * @param text - Text to add
   * @returns This builder for chaining
   */
  text(text: string): XMLBuilder {
    this.elements.push(text);
    return this;
  }

  /**
   * Builds the XML string
   * @param includeDeclaration - Whether to include XML declaration
   * @returns Generated XML string
   */
  build(includeDeclaration = false): string {
    let xml = '';

    if (includeDeclaration) {
      xml += '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    }

    xml += this.elementsToString(this.elements);
    return xml;
  }

  /**
   * Converts elements to XML string
   */
  private elementsToString(elements: (XMLElement | string)[]): string {
    let xml = '';

    for (const element of elements) {
      if (typeof element === 'string') {
        xml += this.escapeXml(element);
      } else {
        xml += this.elementToString(element);
      }
    }

    return xml;
  }

  /**
   * Converts a single element to XML string
   */
  private elementToString(element: XMLElement): string {
    let xml = `<${element.name}`;

    // Add attributes
    if (element.attributes) {
      for (const [key, value] of Object.entries(element.attributes)) {
        if (value !== undefined && value !== null && value !== false) {
          // Handle boolean attributes
          const attrValue = value === true ? key : String(value);
          xml += ` ${key}="${this.escapeXml(attrValue)}"`;
        }
      }
    }

    // Self-closing element
    if (element.selfClosing) {
      xml += '/>';
      return xml;
    }

    xml += '>';

    // Add children
    if (element.children && element.children.length > 0) {
      xml += this.elementsToString(element.children);
    }

    xml += `</${element.name}>`;
    return xml;
  }

  /**
   * Escapes special XML characters (uses appropriate method based on context)
   */
  private escapeXml(text: string): string {
    // For attributes, escape quotes; for text content, don't
    // This method is used in both contexts, so use full escaping
    return XMLBuilder.escapeXmlAttribute(text);
  }

  /**
   * Escapes XML text content (element text nodes)
   * Only escapes: & < >
   * @param text Text to escape
   * @returns Escaped text safe for XML content
   */
  static escapeXmlText(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  /**
   * Escapes XML attribute values
   * Escapes: & < > " '
   * @param value Attribute value to escape
   * @returns Escaped value safe for XML attributes
   */
  static escapeXmlAttribute(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /**
   * Unescapes XML entities back to original characters
   * @param text Text with XML entities
   * @returns Unescaped text
   */
  static unescapeXml(text: string): string {
    return text
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&'); // Must be last to avoid double-unescaping
  }

  /**
   * Creates a WordprocessingML namespace attribute object
   */
  static createNamespaces(): Record<string, string> {
    return {
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
      'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'xmlns:pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    };
  }

  /**
   * Helper method to create a WordprocessingML element
   * @param name - Element name (without 'w:' prefix)
   * @param attributes - Element attributes
   * @param children - Child elements
   * @returns XMLElement
   */
  static w(
    name: string,
    attributes?: Record<string, string | number | boolean | undefined>,
    children?: (XMLElement | string)[]
  ): XMLElement {
    return {
      name: `w:${name}`,
      attributes,
      children,
    };
  }

  /**
   * Helper method to create a self-closing WordprocessingML element
   * @param name - Element name (without 'w:' prefix)
   * @param attributes - Element attributes
   * @returns XMLElement
   */
  static wSelf(
    name: string,
    attributes?: Record<string, string | number | boolean | undefined>
  ): XMLElement {
    return {
      name: `w:${name}`,
      attributes,
      selfClosing: true,
    };
  }

  /**
   * Creates a complete WordprocessingML document structure
   * @param bodyContent - Content for the document body
   * @returns XML string for word/document.xml
   */
  static createDocument(bodyContent: XMLElement[]): string {
    const builder = new XMLBuilder();

    builder.element('w:document', XMLBuilder.createNamespaces(), [
      XMLBuilder.w('body', undefined, bodyContent),
    ]);

    return builder.build(true);
  }
}
