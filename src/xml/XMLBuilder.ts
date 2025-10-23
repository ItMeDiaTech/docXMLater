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
   * @throws {Error} If attempting to create self-closing w:t element (not allowed per ECMA-376)
   */
  selfClosingElement(
    name: string,
    attributes?: Record<string, string | number | boolean | undefined>
  ): XMLBuilder {
    // Validation: Text elements (<w:t>) cannot be self-closing per ECMA-376
    // Self-closing <w:t/> elements cause Word to fail opening the document
    if (name === 'w:t' || name === 't') {
      throw new Error(
        'Text elements (<w:t>) cannot be self-closing per ECMA-376. ' +
        'Use element() with empty text content instead: XMLBuilder.w("t", attrs, [""])'
      );
    }

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
    let xml = "";

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
    let xml = "";

    for (const element of elements) {
      if (typeof element === "string") {
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
          // Use escapeXmlAttribute for attribute values (Issue #8)
          xml += ` ${key}="${XMLBuilder.escapeXmlAttribute(attrValue)}"`;
        }
      }
    }

    // Self-closing element validation
    if (element.selfClosing) {
      // CRITICAL: Certain elements must NEVER be self-closing in Word XML per ECMA-376
      // Self-closing these elements causes Word to not parse correctly or lose content
      const CANNOT_SELF_CLOSE = [
        "w:t",
        "w:r",
        "w:p",
        "w:tbl",
        "w:tr",
        "w:tc",
        "w:body",
        "w:document",
        "w:hyperlink",
        "w:sdt",
        "w:sdtContent",
        "w:sdtPr",
        "w:pPr",
        "w:rPr",
        "w:sectPr",
        "w:bookmarkStart",
        "w:bookmarkEnd",
      ];

      if (CANNOT_SELF_CLOSE.includes(element.name)) {
        // Instead of throwing, force open/close tags for safety
        xml += "></" + element.name + ">";
        return xml;
      }
      xml += "/>";
      return xml;
    }

    xml += ">";

    // Add children
    if (element.children && element.children.length > 0) {
      xml += this.elementsToString(element.children);
    }

    xml += `</${element.name}>`;
    return xml;
  }

  /**
   * Escapes special XML characters for text content
   * (Issue #8 fix: Use escapeXmlText for element text, escapeXmlAttribute called directly for attrs)
   */
  private escapeXml(text: string): string {
    // This method is now only used for text content in elementsToString()
    // Attributes call escapeXmlAttribute() directly in elementToString()
    // Text content should NOT escape quotes (only & < >)
    return XMLBuilder.escapeXmlText(text);
  }

  /**
   * Escapes XML text content (element text nodes)
   * Only escapes: & < >
   * @param text Text to escape
   * @returns Escaped text safe for XML content
   */
  static escapeXmlText(text: string): string {
    return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  /**
   * Escapes XML attribute values
   * Escapes: & < > " '
   * @param value Attribute value to escape
   * @returns Escaped value safe for XML attributes
   */
  static escapeXmlAttribute(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  /**
   * Unescapes XML entities back to original characters
   * @param text Text with XML entities
   * @returns Unescaped text
   */
  static unescapeXml(text: string): string {
    return text
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&"); // Must be last to avoid double-unescaping
  }

  /**
   * Sanitizes and escapes XML content for safe inclusion in XML documents
   * Removes control characters, null bytes, and escapes special XML characters
   * Use this for user-provided content that may contain unsafe characters
   *
   * @param text Text to sanitize and escape
   * @returns Sanitized text safe for XML content
   *
   * **Issue #11 fix:** Prevents malformed XML from CDATA markers, control chars, etc.
   */
  static sanitizeXmlContent(text: string): string {
    return (
      text
        // Remove control characters (except tab, newline, carriage return which are allowed in XML)
        .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "")
        // Escape CDATA end marker to prevent CDATA injection
        .replace(/\]\]>/g, "]]&gt;")
        // Standard XML escaping (& must be first to avoid double-escaping)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
    );
  }

  /**
   * Creates a WordprocessingML namespace attribute object
   */
  static createNamespaces(): Record<string, string> {
    return {
      "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "xmlns:r":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "xmlns:wp":
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
      "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
      "xmlns:wpc":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
      "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "xmlns:o": "urn:schemas-microsoft-com:office:office",
      "xmlns:v": "urn:schemas-microsoft-com:vml",
      "xmlns:wp14":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
      "xmlns:w10": "urn:schemas-microsoft-com:office:word",
      "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
      "xmlns:wpg":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
      "xmlns:wpi":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
      "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml",
      "xmlns:wps":
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
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
  static createDocument(
    bodyContent: XMLElement[],
    namespaces: Record<string, string> = {}
  ): string {
    const builder = new XMLBuilder();

    const allNamespaces = { ...XMLBuilder.createNamespaces(), ...namespaces };

    builder.element("w:document", allNamespaces, [
      XMLBuilder.w("body", undefined, bodyContent),
    ]);

    return builder.build(true);
  }

  /**
   * Builds an XML string from a JavaScript object.
   * This is the reverse of XMLParser.parseToObject
   */
  static buildObject(obj: any, rootName: string): string {
    const builder = new XMLBuilder();
    const element = XMLBuilder.objectToElement(obj, rootName);
    if (element) {
      if (typeof element === "string") {
        builder.text(element);
      } else {
        builder.elements.push(element);
      }
    }
    return builder.build();
  }

  /**
   * Converts a JavaScript object to an XMLElement.
   * @private
   */
  private static objectToElement(
    obj: any,
    name: string
  ): XMLElement | string | null {
    if (obj === null || obj === undefined) {
      return null;
    }

    if (typeof obj !== "object" || obj === null) {
      return String(obj);
    }

    const attributes: Record<string, any> = {};
    const children: (XMLElement | string)[] = [];

    if (obj["#text"] && Object.keys(obj).length === 1) {
      return String(obj["#text"]);
    }

    for (const key in obj) {
      if (key.startsWith("@_")) {
        attributes[key.substring(2)] = obj[key];
      } else if (key === "#text") {
        children.push(String(obj[key]));
      } else {
        const childObj = obj[key];
        if (Array.isArray(childObj)) {
          childObj.forEach((item) => {
            const childElement = XMLBuilder.objectToElement(item, key);
            if (childElement) {
              children.push(childElement);
            }
          });
        } else {
          const childElement = XMLBuilder.objectToElement(childObj, key);
          if (childElement) {
            children.push(childElement);
          }
        }
      }
    }

    const element: XMLElement = {
      name,
      attributes,
      children: children.length > 0 ? children : undefined,
    };

    if (!element.children || element.children.length === 0) {
      const CANNOT_SELF_CLOSE = ["w:t", "w:p", "w:r", "w:document", "w:body"];
      if (!CANNOT_SELF_CLOSE.includes(name)) {
        element.selfClosing = true;
      }
    }

    return element;
  }
}
