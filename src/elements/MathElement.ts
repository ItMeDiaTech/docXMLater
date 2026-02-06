/**
 * MathElement - Represents math elements (m:oMathPara and m:oMath)
 *
 * Per ECMA-376 Part 1 §22.1, math paragraphs (m:oMathPara) are block-level
 * and inline math expressions (m:oMath) can appear within paragraphs.
 *
 * Stored as raw XML for round-trip fidelity since math content has complex
 * internal structure (fractions, radicals, matrices, etc.).
 */

import { XMLElement } from '../xml/XMLBuilder';

/**
 * Block-level math paragraph (m:oMathPara)
 * Can appear directly in the document body alongside paragraphs and tables.
 */
export class MathParagraph {
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
    return 'mathParagraph';
  }
}

/**
 * Inline math expression (m:oMath)
 * Can appear within paragraphs alongside runs.
 */
export class MathExpression {
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
    return 'mathExpression';
  }
}
