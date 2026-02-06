/**
 * PreservedElement - Represents XML elements preserved as raw XML for round-trip fidelity
 *
 * Used for elements that don't need full object model support but must be
 * preserved during round-trip (load → save) to avoid data loss.
 *
 * Examples: w:proofErr, w:permStart, w:permEnd, w:altChunk, w:ruby
 */

import { XMLElement } from '../xml/XMLBuilder';

/**
 * Element context where it can appear
 */
export type PreservedElementContext = 'inline' | 'block';

/**
 * Represents an XML element preserved as raw XML
 */
export class PreservedElement {
  private rawXml: string;
  private elementType: string;
  private context: PreservedElementContext;

  constructor(rawXml: string, elementType: string, context: PreservedElementContext = 'inline') {
    this.rawXml = rawXml;
    this.elementType = elementType;
    this.context = context;
  }

  toXML(): XMLElement {
    return {
      name: '__rawXml',
      rawXml: this.rawXml,
    } as XMLElement;
  }

  getRawXml(): string {
    return this.rawXml;
  }

  getElementType(): string {
    return this.elementType;
  }

  getContext(): PreservedElementContext {
    return this.context;
  }

  getType(): string {
    return 'preservedElement';
  }
}
