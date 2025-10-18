/**
 * AbstractNumbering - Defines a multi-level numbering scheme
 *
 * An abstract numbering definition is a template that defines up to 9 levels of
 * list formatting. It's referenced by numbering instances which link it to actual
 * paragraphs in the document.
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { NumberingLevel } from './NumberingLevel';

/**
 * Properties for creating an abstract numbering definition
 */
export interface AbstractNumberingProperties {
  /** Unique identifier for this abstract numbering */
  abstractNumId: number;

  /** Optional name for the numbering scheme */
  name?: string;

  /** The numbering levels (up to 9 levels, 0-8) */
  levels?: NumberingLevel[];

  /** Optional multiLevel type (0 = single level, 1 = multilevel) */
  multiLevelType?: number;
}

/**
 * Represents an abstract numbering definition
 *
 * Abstract numbering defines the template for a multi-level list. Each instance
 * of a list in the document references an abstract numbering definition.
 */
export class AbstractNumbering {
  private abstractNumId: number;
  private name?: string;
  private levels: Map<number, NumberingLevel>;
  private multiLevelType: number;

  /**
   * Creates a new abstract numbering definition
   * @param properties The abstract numbering properties
   */
  constructor(properties: AbstractNumberingProperties) {
    this.abstractNumId = properties.abstractNumId;
    this.name = properties.name;
    this.levels = new Map();
    this.multiLevelType = properties.multiLevelType !== undefined ? properties.multiLevelType : 1;

    if (properties.levels) {
      properties.levels.forEach(level => {
        this.addLevel(level);
      });
    }

    this.validate();
  }

  /**
   * Validates the abstract numbering
   */
  private validate(): void {
    if (this.abstractNumId < 0) {
      throw new Error('Abstract numbering ID must be non-negative');
    }

    if (this.levels.size > 9) {
      throw new Error('Cannot have more than 9 levels (0-8)');
    }
  }

  /**
   * Gets the abstract numbering ID
   */
  getAbstractNumId(): number {
    return this.abstractNumId;
  }

  /**
   * Alias for getAbstractNumId for backward compatibility
   */
  getId(): number {
    return this.abstractNumId;
  }

  /**
   * Gets the name
   */
  getName(): string | undefined {
    return this.name;
  }

  /**
   * Sets the name
   * @param name The numbering scheme name
   */
  setName(name: string): this {
    this.name = name;
    return this;
  }

  /**
   * Gets the multi-level type
   */
  getMultiLevelType(): string {
    return this.multiLevelType === 1 ? 'multilevel' : 'singleLevel';
  }

  /**
   * Sets the multi-level type
   * @param type The multi-level type ('multilevel' or 'singleLevel')
   */
  setMultiLevelType(type: 'multilevel' | 'singleLevel' | 'hybridMultilevel'): this {
    if (type === 'multilevel') {
      this.multiLevelType = 1;
    } else if (type === 'hybridMultilevel') {
      this.multiLevelType = 2;
    } else {
      this.multiLevelType = 0;
    }
    return this;
  }

  /**
   * Adds a numbering level
   * @param level The numbering level to add
   */
  addLevel(level: NumberingLevel): this {
    const levelIndex = level.getLevel();

    if (levelIndex < 0 || levelIndex > 8) {
      throw new Error(`Level must be between 0 and 8, got ${levelIndex}`);
    }

    this.levels.set(levelIndex, level);
    return this;
  }

  /**
   * Gets a numbering level by index
   * @param levelIndex The level index (0-8)
   */
  getLevel(levelIndex: number): NumberingLevel | undefined {
    return this.levels.get(levelIndex);
  }

  /**
   * Gets all levels
   */
  getAllLevels(): NumberingLevel[] {
    return Array.from(this.levels.values()).sort((a, b) => a.getLevel() - b.getLevel());
  }

  /**
   * Alias for getAllLevels for backward compatibility
   */
  getLevels(): NumberingLevel[] {
    return this.getAllLevels();
  }

  /**
   * Gets the number of levels defined
   */
  getLevelCount(): number {
    return this.levels.size;
  }

  /**
   * Checks if a level exists
   * @param levelIndex The level index (0-8)
   */
  hasLevel(levelIndex: number): boolean {
    return this.levels.has(levelIndex);
  }

  /**
   * Removes a level
   * @param levelIndex The level index (0-8)
   */
  removeLevel(levelIndex: number): boolean {
    return this.levels.delete(levelIndex);
  }

  /**
   * Generates the WordprocessingML XML for this abstract numbering
   */
  toXML(): XMLElement {
    const children: XMLElement[] = [];

    // Add name if present
    if (this.name) {
      children.push(
        XMLBuilder.wSelf('name', { 'w:val': this.name })
      );
    }

    // Add multiLevelType
    children.push(
      XMLBuilder.wSelf('multiLevelType', {
        'w:val': this.multiLevelType === 1 ? 'multilevel' : 'singleLevel'
      })
    );

    // Add all levels in order
    const sortedLevels = this.getAllLevels();
    sortedLevels.forEach(level => {
      children.push(level.toXML());
    });

    // If no levels defined, add a default level 0
    if (sortedLevels.length === 0) {
      children.push(NumberingLevel.createDecimalLevel(0).toXML());
    }

    return XMLBuilder.w('abstractNum', { 'w:abstractNumId': this.abstractNumId.toString() }, children);
  }

  /**
   * Creates a bullet list abstract numbering with specified levels
   * @param abstractNumId The abstract numbering ID
   * @param levels Number of levels (default: 3)
   * @param bullets Array of bullet characters (default: ['•', '○', '▪'])
   */
  static createBulletList(
    abstractNumId: number,
    levels: number = 3,
    bullets: string[] = ['•', '○', '▪']
  ): AbstractNumbering {
    const abstractNum = new AbstractNumbering({
      abstractNumId,
      name: 'Bullet List',
      multiLevelType: 1,
    });

    for (let i = 0; i < levels && i < 9; i++) {
      const bullet = bullets[i % bullets.length] || '•';
      abstractNum.addLevel(NumberingLevel.createBulletLevel(i, bullet));
    }

    return abstractNum;
  }

  /**
   * Creates a numbered list abstract numbering with specified levels
   * @param abstractNumId The abstract numbering ID
   * @param levels Number of levels (default: 3)
   * @param formats Array of formats for each level
   */
  static createNumberedList(
    abstractNumId: number,
    levels: number = 3,
    formats: Array<'decimal' | 'lowerLetter' | 'lowerRoman'> = ['decimal', 'lowerLetter', 'lowerRoman']
  ): AbstractNumbering {
    const abstractNum = new AbstractNumbering({
      abstractNumId,
      name: 'Numbered List',
      multiLevelType: 1,
    });

    for (let i = 0; i < levels && i < 9; i++) {
      const format = formats[i % formats.length] || 'decimal';
      const template = `%${i + 1}.`;

      let level: NumberingLevel;
      switch (format) {
        case 'lowerLetter':
          level = NumberingLevel.createLowerLetterLevel(i, template);
          break;
        case 'lowerRoman':
          level = NumberingLevel.createLowerRomanLevel(i, template);
          break;
        case 'decimal':
        default:
          level = NumberingLevel.createDecimalLevel(i, template);
          break;
      }

      abstractNum.addLevel(level);
    }

    return abstractNum;
  }

  /**
   * Creates a multi-level list with mixed formats
   * @param abstractNumId The abstract numbering ID
   */
  static createMultiLevelList(abstractNumId: number): AbstractNumbering {
    const abstractNum = new AbstractNumbering({
      abstractNumId,
      name: 'Multi-Level List',
      multiLevelType: 1,
    });

    // Level 0: 1, 2, 3, ...
    abstractNum.addLevel(NumberingLevel.createDecimalLevel(0, '%1.'));

    // Level 1: a, b, c, ...
    abstractNum.addLevel(NumberingLevel.createLowerLetterLevel(1, '%2.'));

    // Level 2: i, ii, iii, ...
    abstractNum.addLevel(NumberingLevel.createLowerRomanLevel(2, '%3.'));

    // Level 3: 1, 2, 3, ... (with more indent)
    abstractNum.addLevel(NumberingLevel.createDecimalLevel(3, '%4.'));

    return abstractNum;
  }

  /**
   * Creates an outline list abstract numbering
   * @param abstractNumId The abstract numbering ID
   */
  static createOutlineList(abstractNumId: number): AbstractNumbering {
    const abstractNum = new AbstractNumbering({
      abstractNumId,
      name: 'Outline List',
      multiLevelType: 1,
    });

    // Level 0: I, II, III, ...
    abstractNum.addLevel(NumberingLevel.createUpperRomanLevel(0, '%1.'));

    // Level 1: A, B, C, ...
    abstractNum.addLevel(NumberingLevel.createUpperLetterLevel(1, '%2.'));

    // Level 2: 1, 2, 3, ...
    abstractNum.addLevel(NumberingLevel.createDecimalLevel(2, '%3.'));

    // Level 3: a, b, c, ...
    abstractNum.addLevel(NumberingLevel.createLowerLetterLevel(3, '%4.'));

    // Level 4: i, ii, iii, ...
    abstractNum.addLevel(NumberingLevel.createLowerRomanLevel(4, '%5.'));

    // Level 5: A, B, C, ... (repeating)
    abstractNum.addLevel(NumberingLevel.createUpperLetterLevel(5, '%6.'));

    // Level 6: 1, 2, 3, ... (repeating)
    abstractNum.addLevel(NumberingLevel.createDecimalLevel(6, '%7.'));

    // Level 7: a, b, c, ... (repeating)
    abstractNum.addLevel(NumberingLevel.createLowerLetterLevel(7, '%8.'));

    // Level 8: i, ii, iii, ... (repeating)
    abstractNum.addLevel(NumberingLevel.createLowerRomanLevel(8, '%9.'));

    return abstractNum;
  }

  /**
   * Factory method for creating an abstract numbering definition
   * @param properties The abstract numbering properties
   */
  static create(properties: AbstractNumberingProperties): AbstractNumbering {
    return new AbstractNumbering(properties);
  }
}
