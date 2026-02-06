/**
 * AlternateContent - Represents mc:AlternateContent block element
 *
 * Per ECMA-376 Part 3 §11.4, mc:AlternateContent provides markup compatibility.
 * Word uses this for newer features (Word 2010+ shapes as Choice, VML as Fallback).
 *
 * Stored as raw XML for round-trip fidelity since the internal structure varies widely.
 */

import { XMLElement } from '../xml/XMLBuilder';

export class AlternateContent {
  private rawXml: string;

  constructor(rawXml: string) {
    this.rawXml = rawXml;
  }

  /**
   * Returns the raw XML for round-trip serialization
   */
  toXML(): XMLElement {
    return {
      name: '__rawXml',
      rawXml: this.rawXml,
    } as XMLElement;
  }

  /**
   * Gets the raw XML string
   */
  getRawXml(): string {
    return this.rawXml;
  }

  /**
   * Returns the element type identifier
   */
  getType(): string {
    return 'alternateContent';
  }
}
