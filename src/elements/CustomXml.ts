/**
 * CustomXml - Represents w:customXml block and inline elements
 *
 * Per ECMA-376 Part 1 §17.5.1.6, custom XML elements carry a URI and element
 * name, wrapping block-level or inline content.
 *
 * Stored as raw XML for round-trip fidelity.
 */

import { XMLElement } from '../xml/XMLBuilder';

/**
 * Block-level custom XML (w:customXml wrapping block content)
 */
export class CustomXmlBlock {
  private rawXml: string;

  constructor(rawXml: string) {
    this.rawXml = rawXml;
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

  getType(): string {
    return 'customXmlBlock';
  }
}
