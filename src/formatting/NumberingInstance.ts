/**
 * NumberingInstance - Links paragraphs to abstract numbering definitions
 *
 * A numbering instance references an abstract numbering definition and provides
 * the actual numId that paragraphs use. Multiple instances can reference the same
 * abstract numbering, creating separate list sequences.
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';

/**
 * Properties for creating a numbering instance
 */
export interface NumberingInstanceProperties {
  /** Unique numbering instance ID (numId) */
  numId: number;

  /** Reference to the abstract numbering definition */
  abstractNumId: number;
}

/**
 * Represents a numbering instance
 *
 * Numbering instances link paragraphs to abstract numbering definitions.
 * Each instance creates a separate numbering sequence in the document.
 */
export class NumberingInstance {
  private numId: number;
  private abstractNumId: number;

  /**
   * Creates a new numbering instance
   * @param numIdOrProps The numbering instance ID or properties object
   * @param abstractNumId The abstract numbering ID (if first param is a number)
   */
  constructor(numIdOrProps: number | NumberingInstanceProperties, abstractNumId?: number) {
    if (typeof numIdOrProps === 'number') {
      // Support simple constructor: new NumberingInstance(numId, abstractNumId)
      this.numId = numIdOrProps;
      this.abstractNumId = abstractNumId ?? 0;
    } else {
      // Support object constructor: new NumberingInstance({ numId, abstractNumId })
      this.numId = numIdOrProps.numId;
      this.abstractNumId = numIdOrProps.abstractNumId;
    }

    this.validate();
  }

  /**
   * Validates the numbering instance
   */
  private validate(): void {
    if (this.numId < 0) {
      throw new Error('Numbering instance ID must be non-negative');
    }

    if (this.abstractNumId < 0) {
      throw new Error('Abstract numbering ID must be non-negative');
    }
  }

  /**
   * Gets the numbering instance ID
   */
  getNumId(): number {
    return this.numId;
  }

  /**
   * Gets the abstract numbering ID
   */
  getAbstractNumId(): number {
    return this.abstractNumId;
  }

  /**
   * Alias for getNumId for backward compatibility
   */
  getId(): number {
    return this.numId;
  }

  /**
   * Gets level overrides
   */
  getLevelOverrides(): Map<number, number> {
    // Placeholder for level overrides (not yet implemented)
    return new Map();
  }

  /**
   * Sets level override for a specific level
   * @param _level The level index
   * @param _startValue The starting value for this level
   */
  setLevelOverride(_level: number, _startValue: number): this {
    // Placeholder for level overrides (not yet implemented in current version)
    return this;
  }

  /**
   * Generates the WordprocessingML XML for this numbering instance
   */
  toXML(): XMLElement {
    const children: XMLElement[] = [];

    // Reference to abstract numbering
    children.push(
      XMLBuilder.wSelf('abstractNumId', { 'w:val': this.abstractNumId.toString() })
    );

    return XMLBuilder.w('num', { 'w:numId': this.numId.toString() }, children);
  }

  /**
   * Factory method for creating a numbering instance
   * @param properties The instance properties
   */
  static create(properties: NumberingInstanceProperties): NumberingInstance {
    return new NumberingInstance(properties);
  }
}
