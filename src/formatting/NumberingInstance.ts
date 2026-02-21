/**
 * NumberingInstance - Links paragraphs to abstract numbering definitions
 *
 * A numbering instance references an abstract numbering definition and provides
 * the actual numId that paragraphs use. Multiple instances can reference the same
 * abstract numbering, creating separate list sequences.
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { NumberingLevel } from './NumberingLevel';

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
  private levelOverrides = new Map<number, number>();
  private fullLevelOverrides = new Map<number, NumberingLevel>();

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
   * Sets the abstract numbering ID this instance references
   * @param abstractNumId The new abstract numbering ID
   */
  setAbstractNumId(abstractNumId: number): this {
    if (abstractNumId < 0) {
      throw new Error('Abstract numbering ID must be non-negative');
    }
    this.abstractNumId = abstractNumId;
    return this;
  }

  /**
   * Alias for getNumId for backward compatibility
   */
  getId(): number {
    return this.numId;
  }

  /**
   * Gets level overrides
   * Returns a map of level indices to their override starting values
   */
  getLevelOverrides(): Map<number, number> {
    return new Map(this.levelOverrides);
  }

  /**
   * Sets level override for a specific level
   * Overrides the starting value for a particular numbering level
   *
   * @param level The level index (0-based)
   * @param startValue The starting value for this level
   * @returns This instance for method chaining
   */
  setLevelOverride(level: number, startValue: number): this {
    if (level < 0) {
      throw new Error('Level index must be non-negative');
    }
    if (startValue < 0) {
      throw new Error('Start value must be non-negative');
    }

    this.levelOverrides.set(level, startValue);
    return this;
  }

  /**
   * Clears a level override for a specific level
   *
   * @param level The level index to clear
   * @returns This instance for method chaining
   */
  clearLevelOverride(level: number): this {
    this.levelOverrides.delete(level);
    return this;
  }

  /**
   * Gets the override value for a specific level, if set
   *
   * @param level The level index
   * @returns The override starting value, or undefined if not set
   */
  getLevelOverride(level: number): number | undefined {
    return this.levelOverrides.get(level);
  }

  /**
   * Sets a full level definition override for a specific level
   * This replaces the entire level definition from the abstract numbering
   * (ECMA-376 §17.9.8 - w:lvlOverride with full w:lvl child)
   *
   * @param level The level index (0-based)
   * @param levelDef The full NumberingLevel definition to use as override
   */
  setFullLevelOverride(level: number, levelDef: NumberingLevel): this {
    if (level < 0) {
      throw new Error('Level index must be non-negative');
    }
    this.fullLevelOverrides.set(level, levelDef);
    return this;
  }

  /**
   * Gets a full level definition override for a specific level
   */
  getFullLevelOverride(level: number): NumberingLevel | undefined {
    return this.fullLevelOverrides.get(level);
  }

  /**
   * Gets all full level definition overrides
   */
  getFullLevelOverrides(): Map<number, NumberingLevel> {
    return new Map(this.fullLevelOverrides);
  }

  /**
   * Clears a full level override
   */
  clearFullLevelOverride(level: number): this {
    this.fullLevelOverrides.delete(level);
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

    // Add level overrides if any are set
    for (const [level, startValue] of this.levelOverrides) {
      // Skip levels that have a full level override (they take precedence)
      if (this.fullLevelOverrides.has(level)) continue;
      children.push({
        name: 'w:lvlOverride',
        attributes: { 'w:ilvl': level.toString() },
        children: [
          XMLBuilder.wSelf('startOverride', { 'w:val': startValue.toString() })
        ]
      });
    }

    // Add full level overrides
    for (const [level, levelDef] of this.fullLevelOverrides) {
      const overrideChildren: XMLElement[] = [];
      // Include startOverride if also set for this level
      if (this.levelOverrides.has(level)) {
        overrideChildren.push(
          XMLBuilder.wSelf('startOverride', { 'w:val': this.levelOverrides.get(level)!.toString() })
        );
      }
      overrideChildren.push(levelDef.toXML());
      children.push({
        name: 'w:lvlOverride',
        attributes: { 'w:ilvl': level.toString() },
        children: overrideChildren,
      });
    }

    return XMLBuilder.w('num', { 'w:numId': this.numId.toString() }, children);
  }

  /**
   * Factory method for creating a numbering instance
   * @param propertiesOrNumId The instance properties object, or numId (number)
   * @param abstractNumId The abstract numbering ID (if first param is a number)
   */
  static create(
    propertiesOrNumId: NumberingInstanceProperties | number,
    abstractNumId?: number
  ): NumberingInstance {
    if (typeof propertiesOrNumId === 'number') {
      return new NumberingInstance(propertiesOrNumId, abstractNumId);
    }
    return new NumberingInstance(propertiesOrNumId);
  }

  /**
   * Creates a NumberingInstance from XML element
   * @param xml The XML string of the <w:num> element
   * @returns NumberingInstance instance
   */
  static fromXML(xml: string): NumberingInstance {
    // Extract numId (required)
    const numIdMatch = /<w:num[^>]*w:numId="([^"]+)"/.exec(xml);
    if (!numIdMatch?.[1]) {
      throw new Error('Missing required w:numId attribute');
    }
    const numId = parseInt(numIdMatch[1], 10);

    // Extract abstractNumId (required)
    const abstractNumIdMatch = /<w:abstractNumId[^>]*w:val="([^"]+)"/.exec(xml);
    if (!abstractNumIdMatch?.[1]) {
      throw new Error('Missing required w:abstractNumId element');
    }
    const abstractNumId = parseInt(abstractNumIdMatch[1], 10);

    const instance = new NumberingInstance({
      numId,
      abstractNumId,
    });

    // Parse level overrides (w:lvlOverride)
    const lvlOverrideRegex = /<w:lvlOverride[^>]*w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvlOverride>/g;
    let match: RegExpExecArray | null;
    while ((match = lvlOverrideRegex.exec(xml)) !== null) {
      const levelStr = match[1]!;
      const content = match[2]!;
      const level = parseInt(levelStr, 10);

      // Check for startOverride
      const startOverrideMatch = /<w:startOverride[^>]*w:val="([^"]+)"/.exec(content);
      if (startOverrideMatch?.[1]) {
        instance.setLevelOverride(level, parseInt(startOverrideMatch[1], 10));
      }

      // Check for full w:lvl element
      const lvlMatch = /<w:lvl[^>]*>[\s\S]*?<\/w:lvl>/.exec(content);
      if (lvlMatch) {
        try {
          const levelDef = NumberingLevel.fromXML(lvlMatch[0]);
          instance.setFullLevelOverride(level, levelDef);
        } catch {
          // Skip invalid level definitions
        }
      }
    }

    return instance;
  }
}
