/**
 * XMLParser - Simple position-based XML parser
 * Avoids regex backtracking issues that can cause ReDoS attacks
 * Completes the DocXML framework (XMLBuilder + XMLParser)
 */

/**
 * Simple XML parser using position-based parsing instead of regex
 * Prevents catastrophic backtracking (ReDoS attacks) by avoiding nested regex patterns
 */
export class XMLParser {
  /**
   * Extracts the body content from a Word document XML
   * @param docXml - The complete document.xml content
   * @returns The body content, or empty string if not found
   */
  static extractBody(docXml: string): string {
    const startTag = '<w:body';
    const endTag = '</w:body>';

    const startIdx = docXml.indexOf(startTag);
    if (startIdx === -1) return '';

    // Find the closing > of opening tag
    const openEnd = docXml.indexOf('>', startIdx);
    if (openEnd === -1) return '';

    // Find matching closing tag
    const endIdx = docXml.indexOf(endTag, openEnd);
    if (endIdx === -1) return '';

    return docXml.substring(openEnd + 1, endIdx);
  }

  /**
   * Extracts all elements of a given type using position-based parsing
   * Handles nested tags correctly by tracking depth
   * @param xml - XML content to parse
   * @param tagName - Tag name to extract (e.g., 'w:p', 'w:r')
   * @returns Array of XML strings for each element
   */
  static extractElements(xml: string, tagName: string): string[] {
    const elements: string[] = [];
    const openTag = `<${tagName}`;
    const closeTag = `</${tagName}>`;
    const selfClosingEnd = '/>';

    let pos = 0;
    while (pos < xml.length) {
      const startIdx = xml.indexOf(openTag, pos);
      if (startIdx === -1) break;

      // Find the end of opening tag
      const openEnd = xml.indexOf('>', startIdx);
      if (openEnd === -1) break;

      // Check if self-closing
      if (xml.substring(openEnd - 1, openEnd + 1) === selfClosingEnd) {
        elements.push(xml.substring(startIdx, openEnd + 1));
        pos = openEnd + 1;
        continue;
      }

      // Find matching closing tag (handle nesting)
      let depth = 1;
      let searchPos = openEnd + 1;

      while (depth > 0 && searchPos < xml.length) {
        const nextOpen = xml.indexOf(openTag, searchPos);
        const nextClose = xml.indexOf(closeTag, searchPos);

        if (nextClose === -1) break;

        if (nextOpen !== -1 && nextOpen < nextClose) {
          depth++;
          searchPos = nextOpen + openTag.length;
        } else {
          depth--;
          if (depth === 0) {
            elements.push(xml.substring(startIdx, nextClose + closeTag.length));
            pos = nextClose + closeTag.length;
          } else {
            searchPos = nextClose + closeTag.length;
          }
        }
      }

      if (depth > 0) {
        // Unclosed tag - skip it
        pos = startIdx + openTag.length;
      }
    }

    return elements;
  }

  /**
   * Extracts attribute value from an XML string
   * @param xml - XML content
   * @param attributeName - Attribute name (e.g., 'w:val')
   * @returns Attribute value or undefined
   */
  static extractAttribute(xml: string, attributeName: string): string | undefined {
    // Use simple indexOf for bounded string search (safe)
    const attrPattern = `${attributeName}="`;
    const startIdx = xml.indexOf(attrPattern);
    if (startIdx === -1) return undefined;

    const valueStart = startIdx + attrPattern.length;
    const valueEnd = xml.indexOf('"', valueStart);
    if (valueEnd === -1) return undefined;

    return xml.substring(valueStart, valueEnd);
  }

  /**
   * Checks if an XML string contains a self-closing tag
   * @param xml - XML content
   * @param tagName - Tag name to check
   * @returns True if the tag exists as self-closing
   */
  static hasSelfClosingTag(xml: string, tagName: string): boolean {
    return xml.includes(`<${tagName}/>`) || xml.includes(`<${tagName} `);
  }

  /**
   * Extracts text content from within tags
   * Finds all <w:t>...</w:t> tags and extracts their text
   * @param xml - XML content
   * @returns Combined text content
   */
  static extractText(xml: string): string {
    const texts: string[] = [];
    const openTag = '<w:t';
    const closeTag = '</w:t>';

    let pos = 0;
    while (pos < xml.length) {
      const startIdx = xml.indexOf(openTag, pos);
      if (startIdx === -1) break;

      // Find the end of opening tag
      const openEnd = xml.indexOf('>', startIdx);
      if (openEnd === -1) break;

      // Find closing tag
      const closeIdx = xml.indexOf(closeTag, openEnd);
      if (closeIdx === -1) break;

      // Extract text between tags
      const text = xml.substring(openEnd + 1, closeIdx);
      texts.push(text);

      pos = closeIdx + closeTag.length;
    }

    return texts.join('');
  }

  /**
   * Validates input size to prevent excessive memory usage
   * @param xml - XML content
   * @param maxSize - Maximum size in bytes (default: 10MB)
   * @throws Error if XML exceeds max size
   */
  static validateSize(xml: string, maxSize: number = 10 * 1024 * 1024): void {
    if (xml.length > maxSize) {
      throw new Error(
        `XML content too large for parsing (${(xml.length / 1024 / 1024).toFixed(1)}MB). ` +
        `Maximum allowed: ${(maxSize / 1024 / 1024).toFixed(0)}MB`
      );
    }
  }

  /**
   * Extracts content between two specific tags
   * More efficient than regex for large documents
   * @param xml - XML content
   * @param startTag - Opening tag (e.g., '<w:pPr')
   * @param endTag - Closing tag (e.g., '</w:pPr>')
   * @returns Content between tags, or undefined if not found
   */
  static extractBetweenTags(xml: string, startTag: string, endTag: string): string | undefined {
    const startIdx = xml.indexOf(startTag);
    if (startIdx === -1) return undefined;

    // Find the end of the opening tag
    const openEnd = xml.indexOf('>', startIdx);
    if (openEnd === -1) return undefined;

    // Find the closing tag
    const endIdx = xml.indexOf(endTag, openEnd);
    if (endIdx === -1) return undefined;

    return xml.substring(openEnd + 1, endIdx);
  }
}
