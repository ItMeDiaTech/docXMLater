/**
 * NumberingManager - Manages numbering definitions and generates numbering.xml
 *
 * The NumberingManager is responsible for managing all abstract numbering definitions
 * and numbering instances in a document, and generating the numbering.xml file.
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { AbstractNumbering } from './AbstractNumbering';
import { NumberingInstance } from './NumberingInstance';
import { NumberingLevel } from './NumberingLevel';
import { defaultLogger } from '../utils/logger';

/**
 * Options for numbering consolidation
 */
export interface NumberingConsolidationOptions {
  /** AbstractNumIds to exclude from consolidation (e.g., HLP/row-number lists) */
  protectedAbstractNumIds?: Set<number>;
}

/**
 * Result of numbering consolidation
 */
export interface NumberingConsolidationResult {
  abstractNumsRemoved: number;
  instancesRemapped: number;
  groupsConsolidated: number;
}

/**
 * Manages numbering definitions and instances for a document
 */
export class NumberingManager {
  private abstractNumberings: Map<number, AbstractNumbering>;
  private instances: Map<number, NumberingInstance>;
  private nextAbstractNumId: number;
  private nextNumId: number;

  // Track if numbering has been modified (for XML preservation)
  private _modified = false;

  // Track which specific definitions have been modified (for selective merging)
  private _modifiedAbstractNumIds = new Set<number>();
  private _modifiedNumIds = new Set<number>();

  // Track which definitions have been removed (for removal from original XML during merge)
  private _removedAbstractNumIds = new Set<number>();
  private _removedNumIds = new Set<number>();

  /**
   * Creates a new numbering manager
   * @param initializeDefaults Whether to initialize with default numbering definitions
   */
  constructor(initializeDefaults = false) {
    this.abstractNumberings = new Map();
    this.instances = new Map();
    this.nextAbstractNumId = 0;
    this.nextNumId = 1;

    if (initializeDefaults) {
      this.initializeDefaultNumberings();
    }
  }

  /**
   * Initializes default numbering definitions (bullet and numbered lists)
   */
  private initializeDefaultNumberings(): void {
    // Create default bullet list
    const bulletAbstract = AbstractNumbering.createBulletList(this.nextAbstractNumId++);
    this.addAbstractNumbering(bulletAbstract);

    // Create default numbered list
    const numberedAbstract = AbstractNumbering.createNumberedList(this.nextAbstractNumId++);
    this.addAbstractNumbering(numberedAbstract);
  }

  /**
   * Adds an abstract numbering definition
   * @param abstractNumbering The abstract numbering to add
   */
  addAbstractNumbering(abstractNumbering: AbstractNumbering): this {
    const id = abstractNumbering.getAbstractNumId();
    this.abstractNumberings.set(id, abstractNumbering);

    // Update next ID if necessary
    if (id >= this.nextAbstractNumId) {
      this.nextAbstractNumId = id + 1;
    }

    this._modified = true;
    this._modifiedAbstractNumIds.add(id);
    return this;
  }

  /**
   * Gets an abstract numbering by ID
   * @param abstractNumId The abstract numbering ID
   */
  getAbstractNumbering(abstractNumId: number): AbstractNumbering | undefined {
    return this.abstractNumberings.get(abstractNumId);
  }

  /**
   * Gets all abstract numberings
   */
  getAllAbstractNumberings(): AbstractNumbering[] {
    return Array.from(this.abstractNumberings.values()).sort(
      (a, b) => a.getAbstractNumId() - b.getAbstractNumId()
    );
  }

  /**
   * Checks if an abstract numbering exists
   * @param abstractNumId The abstract numbering ID
   */
  hasAbstractNumbering(abstractNumId: number): boolean {
    return this.abstractNumberings.has(abstractNumId);
  }

  /**
   * Removes an abstract numbering
   * @param abstractNumId The abstract numbering ID
   */
  removeAbstractNumbering(abstractNumId: number): boolean {
    // Also remove all instances referencing this abstract numbering
    const instancesToRemove: number[] = [];
    this.instances.forEach((instance, numId) => {
      if (instance.getAbstractNumId() === abstractNumId) {
        instancesToRemove.push(numId);
      }
    });

    instancesToRemove.forEach(numId => {
      this.instances.delete(numId);
      this._removedNumIds.add(numId);
    });

    const deleted = this.abstractNumberings.delete(abstractNumId);
    if (deleted) {
      this._modified = true;
      this._removedAbstractNumIds.add(abstractNumId);
    }
    return deleted;
  }

  /**
   * Adds a numbering instance
   * @param instance The numbering instance to add
   */
  addInstance(instance: NumberingInstance): this {
    const numId = instance.getNumId();
    const abstractNumId = instance.getAbstractNumId();

    // Verify that the abstract numbering exists
    if (!this.hasAbstractNumbering(abstractNumId)) {
      throw new Error(`Abstract numbering ${abstractNumId} does not exist`);
    }

    this.instances.set(numId, instance);
    this._modified = true;
    this._modifiedNumIds.add(numId);

    // Update next ID if necessary
    if (numId >= this.nextNumId) {
      this.nextNumId = numId + 1;
    }

    return this;
  }

  /**
   * Alias for addInstance for backward compatibility
   * @param instance The numbering instance to add
   */
  addNumberingInstance(instance: NumberingInstance): this {
    return this.addInstance(instance);
  }

  /**
   * Gets a numbering instance by ID
   * @param numId The numbering instance ID
   */
  getInstance(numId: number): NumberingInstance | undefined {
    return this.instances.get(numId);
  }

  /**
   * Alias for getInstance for backward compatibility
   * @param numId The numbering instance ID
   */
  getNumberingInstance(numId: number): NumberingInstance | undefined {
    return this.getInstance(numId);
  }

  /**
   * Gets all numbering instances
   */
  getAllInstances(): NumberingInstance[] {
    return Array.from(this.instances.values()).sort(
      (a, b) => a.getNumId() - b.getNumId()
    );
  }

  /**
   * Checks if numbering has been modified since loading
   * Used for XML preservation optimization
   * @returns True if numbering was added or modified
   */
  isModified(): boolean {
    return this._modified;
  }

  /**
   * Marks an abstract numbering as modified for XML regeneration.
   * Use when modifying NumberingLevel properties directly (setLeftIndent, etc.)
   * which don't automatically trigger the modified flag.
   * @param abstractNumId The abstract numbering ID to mark as modified
   */
  markAbstractNumberingModified(abstractNumId: number): void {
    if (this.abstractNumberings.has(abstractNumId)) {
      this._modified = true;
      this._modifiedAbstractNumIds.add(abstractNumId);
    }
  }

  /**
   * Resets the modified flag
   * Called after parsing to indicate that loaded numbering doesn't count as modifications
   */
  resetModified(): void {
    this._modified = false;
    this._modifiedAbstractNumIds.clear();
    this._modifiedNumIds.clear();
    this._removedAbstractNumIds.clear();
    this._removedNumIds.clear();
  }

  /**
   * Gets the IDs of abstract numberings that have been modified since loading
   * Used for selective merging with original numbering.xml
   */
  getModifiedAbstractNumIds(): Set<number> {
    return new Set(this._modifiedAbstractNumIds);
  }

  /**
   * Gets the IDs of numbering instances that have been modified since loading
   * Used for selective merging with original numbering.xml
   */
  getModifiedNumIds(): Set<number> {
    return new Set(this._modifiedNumIds);
  }

  /**
   * Gets the IDs of abstract numberings that have been removed since loading
   * Used for removal from original numbering.xml during merge
   */
  getRemovedAbstractNumIds(): Set<number> {
    return new Set(this._removedAbstractNumIds);
  }

  /**
   * Gets the IDs of numbering instances that have been removed since loading
   * Used for removal from original numbering.xml during merge
   */
  getRemovedNumIds(): Set<number> {
    return new Set(this._removedNumIds);
  }

  /**
   * Alias for getAllInstances for backward compatibility
   */
  getAllNumberingInstances(): NumberingInstance[] {
    return this.getAllInstances();
  }

  /**
   * Checks if a numbering instance exists
   * @param numId The numbering instance ID
   */
  hasInstance(numId: number): boolean {
    return this.instances.has(numId);
  }

  /**
   * Removes a numbering instance
   * @param numId The numbering instance ID
   */
  removeInstance(numId: number): boolean {
    const deleted = this.instances.delete(numId);
    if (deleted) {
      this._modified = true;
      this._removedNumIds.add(numId);
    }
    return deleted;
  }

  /**
   * Creates a new bullet list and returns its numId
   * @param levels Number of levels (default: 9)
   * @param bullets Array of bullet characters
   */
  createBulletList(levels = 9, bullets?: string[]): number {
    // Create abstract numbering
    const abstractNumId = this.nextAbstractNumId++;
    // Only pass bullets if it's defined, so defaults are used otherwise
    const abstractNumbering = bullets
      ? AbstractNumbering.createBulletList(abstractNumId, levels, bullets)
      : AbstractNumbering.createBulletList(abstractNumId, levels);
    this.addAbstractNumbering(abstractNumbering);

    // Create instance
    const numId = this.nextNumId++;
    const instance = NumberingInstance.create({ numId, abstractNumId });
    this.addInstance(instance);

    return numId;
  }

  /**
   * Creates a new numbered list and returns its numId
   * @param levels Number of levels (default: 9)
   * @param formats Array of formats for each level
   */
  createNumberedList(
    levels = 9,
    formats?: ('decimal' | 'lowerLetter' | 'lowerRoman')[]
  ): number {
    // Create abstract numbering
    const abstractNumId = this.nextAbstractNumId++;
    // Only pass formats if it's defined, so defaults are used otherwise
    const abstractNumbering = formats
      ? AbstractNumbering.createNumberedList(abstractNumId, levels, formats)
      : AbstractNumbering.createNumberedList(abstractNumId, levels);
    this.addAbstractNumbering(abstractNumbering);

    // Create instance
    const numId = this.nextNumId++;
    const instance = NumberingInstance.create({ numId, abstractNumId });
    this.addInstance(instance);

    return numId;
  }

  /**
   * Creates a new multi-level list and returns its numId
   */
  createMultiLevelList(): number {
    // Create abstract numbering
    const abstractNumId = this.nextAbstractNumId++;
    const abstractNumbering = AbstractNumbering.createMultiLevelList(abstractNumId);
    this.addAbstractNumbering(abstractNumbering);

    // Create instance
    const numId = this.nextNumId++;
    const instance = NumberingInstance.create({ numId, abstractNumId });
    this.addInstance(instance);

    return numId;
  }

  /**
   * Creates a new outline list and returns its numId
   */
  createOutlineList(): number {
    // Create abstract numbering
    const abstractNumId = this.nextAbstractNumId++;
    const abstractNumbering = AbstractNumbering.createOutlineList(abstractNumId);
    this.addAbstractNumbering(abstractNumbering);

    // Create instance
    const numId = this.nextNumId++;
    const instance = NumberingInstance.create({ numId, abstractNumId });
    this.addInstance(instance);

    return numId;
  }

  /**
   * Creates a custom list with specified levels and returns its numId
   * @param levels Array of numbering levels
   * @param name Optional name for the list
   */
  createCustomList(levels: NumberingLevel[], name?: string): number {
    // Create abstract numbering
    const abstractNumId = this.nextAbstractNumId++;
    const abstractNumbering = AbstractNumbering.create({
      abstractNumId,
      name,
      levels,
    });
    this.addAbstractNumbering(abstractNumbering);

    // Create instance
    const numId = this.nextNumId++;
    const instance = NumberingInstance.create({ numId, abstractNumId });
    this.addInstance(instance);

    return numId;
  }

  /**
   * Creates a new instance of an existing abstract numbering
   * @param abstractNumId The abstract numbering ID to create an instance of
   */
  createInstance(abstractNumId: number): number {
    if (!this.hasAbstractNumbering(abstractNumId)) {
      throw new Error(`Abstract numbering ${abstractNumId} does not exist`);
    }

    const numId = this.nextNumId++;
    const instance = NumberingInstance.create({ numId, abstractNumId });
    this.addInstance(instance);

    return numId;
  }

  /**
   * Gets the framework's standard indentation for a list level
   *
   * The framework uses a consistent indentation scheme:
   * - leftIndent: 720 + (level * 360) twips
   * - hangingIndent: 360 twips
   *
   * Examples:
   * - Level 0: 720 twips (0.5 inch) left, 360 twips hanging
   * - Level 1: 1080 twips (0.75 inch) left, 360 twips hanging
   * - Level 2: 1440 twips (1.0 inch) left, 360 twips hanging
   *
   * @param level The level (0-8)
   * @returns Object with leftIndent and hangingIndent in twips
   * @example
   * ```typescript
   * const indent = manager.getStandardIndentation(0);
   * // Returns: { leftIndent: 720, hangingIndent: 360 }
   * ```
   */
  getStandardIndentation(level: number): { leftIndent: number; hangingIndent: number } {
    if (level < 0 || level > 8) {
      throw new Error(`Invalid level ${level}. Level must be between 0 and 8.`);
    }

    return {
      leftIndent: 720 + (level * 360),
      hangingIndent: 360,
    };
  }

  /**
   * Sets custom indentation for a specific level in a numbering definition
   *
   * This updates the indentation for a specific level across ALL paragraphs
   * that use this numId and level combination.
   *
   * @param numId The numbering instance ID
   * @param level The level to modify (0-8)
   * @param leftIndent Left indentation in twips
   * @param hangingIndent Hanging indentation in twips (optional, defaults to 360)
   * @returns true if successful, false if numId doesn't exist
   * @example
   * ```typescript
   * // Set level 0 to 0.5 inch left, 0.25 inch hanging
   * manager.setListIndentation(1, 0, 720, 360);
   * ```
   */
  setListIndentation(
    numId: number,
    level: number,
    leftIndent: number,
    hangingIndent = 360
  ): boolean {
    // Validate level
    if (level < 0 || level > 8) {
      throw new Error(`Invalid level ${level}. Level must be between 0 and 8.`);
    }

    // Validate indents (clamp negatives to 0)
    leftIndent = Math.max(0, leftIndent);
    hangingIndent = Math.max(0, hangingIndent);

    // Get the numbering instance
    const instance = this.getInstance(numId);
    if (!instance) {
      defaultLogger.warn(`Numbering instance ${numId} does not exist`);
      return false;
    }

    // Get the abstract numbering
    const abstractNum = this.getAbstractNumbering(instance.getAbstractNumId());
    if (!abstractNum) {
      defaultLogger.warn(`Abstract numbering for instance ${numId} does not exist`);
      return false;
    }

    // Get the level from the abstract numbering
    const numLevel = abstractNum.getLevel(level);
    if (!numLevel) {
      defaultLogger.warn(`Level ${level} does not exist in abstract numbering`);
      return false;
    }

    // Update the level's indentation
    numLevel.setLeftIndent(leftIndent);
    numLevel.setHangingIndent(hangingIndent);

    return true;
  }

  /**
   * Resets all levels in a numbering definition to standard indentation
   *
   * This applies the framework's standard indentation formula to all levels:
   * - leftIndent: 720 + (level * 360) twips
   * - hangingIndent: 360 twips
   *
   * @param numId The numbering instance ID
   * @returns true if successful, false if numId doesn't exist
   * @example
   * ```typescript
   * // Reset list 1 to standard indentation
   * manager.normalizeListIndentation(1);
   * ```
   */
  normalizeListIndentation(numId: number): boolean {
    // Get the numbering instance
    const instance = this.getInstance(numId);
    if (!instance) {
      defaultLogger.warn(`Numbering instance ${numId} does not exist`);
      return false;
    }

    // Get the abstract numbering
    const abstractNum = this.getAbstractNumbering(instance.getAbstractNumId());
    if (!abstractNum) {
      defaultLogger.warn(`Abstract numbering for instance ${numId} does not exist`);
      return false;
    }

    // Get all levels
    const levels = abstractNum.getAllLevels();

    // Apply standard indentation to each level
    for (const level of levels) {
      const standardIndent = this.getStandardIndentation(level.getLevel());
      level.setLeftIndent(standardIndent.leftIndent);
      level.setHangingIndent(standardIndent.hangingIndent);
    }

    return true;
  }

  /**
   * Normalizes indentation for all lists in the document
   *
   * Applies standard indentation to every numbering instance:
   * - leftIndent: 720 + (level * 360) twips
   * - hangingIndent: 360 twips
   *
   * This ensures consistent spacing across all lists in the document.
   *
   * @returns Number of numbering instances updated
   * @example
   * ```typescript
   * const count = manager.normalizeAllListIndentation();
   * console.log(`Normalized ${count} lists`);
   * ```
   */
  normalizeAllListIndentation(): number {
    let count = 0;

    // Iterate over all instances
    for (const instance of this.getAllInstances()) {
      const success = this.normalizeListIndentation(instance.getNumId());
      if (success) {
        count++;
      }
    }

    return count;
  }

  /**
   * Gets the total number of abstract numberings
   */
  getAbstractNumberingCount(): number {
    return this.abstractNumberings.size;
  }

  /**
   * Gets the total number of numbering instances
   */
  getInstanceCount(): number {
    return this.instances.size;
  }

  /**
   * Clears all numbering definitions and instances
   */
  clear(): this {
    this.abstractNumberings.clear();
    this.instances.clear();
    this.nextAbstractNumId = 0;
    this.nextNumId = 1;
    return this;
  }

  /**
   * Removes unused numbering instances and abstract numberings
   *
   * This method cleans up numbering definitions that are no longer referenced
   * by any paragraphs in the document. It removes:
   * 1. Instances not in the usedNumIds set
   * 2. Abstract numberings not referenced by any remaining instance
   *
   * @param usedNumIds Set of numIds currently used by paragraphs
   * @returns Object with counts of removed instances and abstract numberings
   */
  cleanupUnusedNumbering(usedNumIds: Set<number>): { instancesRemoved: number; abstractsRemoved: number } {
    let instancesRemoved = 0;
    let abstractsRemoved = 0;

    // Step 1: Remove unused instances
    const instancesToRemove: number[] = [];
    this.instances.forEach((_instance, numId) => {
      if (!usedNumIds.has(numId)) {
        instancesToRemove.push(numId);
      }
    });

    for (const numId of instancesToRemove) {
      this.instances.delete(numId);
      this._modified = true;
      this._removedNumIds.add(numId);
      instancesRemoved++;
    }

    // Step 2: Find abstract numberings still referenced by remaining instances
    const referencedAbstractNumIds = new Set<number>();
    this.instances.forEach(instance => {
      referencedAbstractNumIds.add(instance.getAbstractNumId());
    });

    // Step 3: Remove unreferenced abstract numberings
    const abstractsToRemove: number[] = [];
    this.abstractNumberings.forEach((_abstractNum, abstractNumId) => {
      if (!referencedAbstractNumIds.has(abstractNumId)) {
        abstractsToRemove.push(abstractNumId);
      }
    });

    for (const abstractNumId of abstractsToRemove) {
      this.abstractNumberings.delete(abstractNumId);
      this._modified = true;
      this._removedAbstractNumIds.add(abstractNumId);
      abstractsRemoved++;
    }

    return { instancesRemoved, abstractsRemoved };
  }

  /**
   * Consolidates duplicate abstract numbering definitions
   *
   * Groups abstractNums by a deterministic fingerprint of their level properties
   * (format, text, font, fontSize, color, indentation, alignment, etc.).
   * For each group with >1 member, picks the lowest abstractNumId as canonical,
   * remaps all instances pointing to non-canonical IDs, and removes duplicates.
   *
   * This is safe because multiple num instances can reference the same abstractNum —
   * each instance maintains its own independent counter via level overrides.
   *
   * @param options Optional configuration (e.g., protected IDs to skip)
   * @returns Summary of what was consolidated
   */
  consolidateNumbering(options?: NumberingConsolidationOptions): NumberingConsolidationResult {
    const protectedIds = options?.protectedAbstractNumIds ?? new Set<number>();

    // 1. Compute fingerprint for each non-protected abstractNum
    const fingerprintGroups = new Map<string, number[]>();
    for (const abstractNum of this.abstractNumberings.values()) {
      const id = abstractNum.getAbstractNumId();
      if (protectedIds.has(id)) continue;

      const fingerprint = this._fingerprintAbstractNum(abstractNum);
      const group = fingerprintGroups.get(fingerprint);
      if (group) {
        group.push(id);
      } else {
        fingerprintGroups.set(fingerprint, [id]);
      }
    }

    // 2. For each group with >1 member, consolidate
    let abstractNumsRemoved = 0;
    let instancesRemapped = 0;
    let groupsConsolidated = 0;

    for (const [, ids] of fingerprintGroups) {
      if (ids.length <= 1) continue;

      // Sort to pick lowest ID as canonical
      ids.sort((a, b) => a - b);
      const canonicalId = ids[0]!;
      const duplicateIds = new Set(ids.slice(1));

      groupsConsolidated++;

      // Remap instances pointing to duplicate abstractNums
      for (const instance of this.instances.values()) {
        if (duplicateIds.has(instance.getAbstractNumId())) {
          instance.setAbstractNumId(canonicalId);
          this._modifiedNumIds.add(instance.getNumId());
          instancesRemapped++;
        }
      }

      // Remove duplicate abstractNums
      for (const dupId of duplicateIds) {
        this.abstractNumberings.delete(dupId);
        this._removedAbstractNumIds.add(dupId);
        abstractNumsRemoved++;
      }
    }

    if (abstractNumsRemoved > 0) {
      this._modified = true;
    }

    return { abstractNumsRemoved, instancesRemapped, groupsConsolidated };
  }

  /**
   * Computes a deterministic fingerprint for an abstract numbering definition
   * based on its multiLevelType and all level properties. Name and abstractNumId
   * are excluded since they don't affect rendering.
   */
  private _fingerprintAbstractNum(abstractNum: AbstractNumbering): string {
    const parts: string[] = [
      abstractNum.getNumStyleLink() ?? '',
      abstractNum.getStyleLink() ?? '',
      abstractNum.getMultiLevelType(),
    ];

    for (const level of abstractNum.getAllLevels()) {
      const props = level.getProperties();
      parts.push(
        `${props.level}|${props.format}|${props.text}|${props.font}|${props.fontSize}|` +
        `${props.color}|${props.leftIndent}|${props.hangingIndent}|${props.alignment}|` +
        `${props.start}|${props.bold}|${props.italic}|${props.underline ?? ''}|` +
        `${props.suffix}|${props.isLegalNumberingStyle}|${props.lvlRestart ?? ''}`
      );
    }

    return parts.join('::');
  }

  /**
   * Generates the complete numbering.xml content
   */
  generateNumberingXml(): string {
    const builder = new XMLBuilder();

    const children: XMLElement[] = [];

    // Add all abstract numberings
    const abstractNumberings = this.getAllAbstractNumberings();
    abstractNumberings.forEach(abstractNum => {
      children.push(abstractNum.toXML());
    });

    // Add all numbering instances
    const instances = this.getAllInstances();
    instances.forEach(instance => {
      children.push(instance.toXML());
    });

    const numbering = XMLBuilder.w('numbering', {
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }, children);

    builder.element(numbering.name, numbering.attributes, numbering.children);

    // Generate XML with declaration
    return builder.build(true);
  }

  /**
   * Generates the numbering.xml as XMLElement (for API compatibility)
   */
  toXML(): XMLElement {
    const children: XMLElement[] = [];

    // Add all abstract numberings
    const abstractNumberings = this.getAllAbstractNumberings();
    abstractNumberings.forEach(abstractNum => {
      children.push(abstractNum.toXML());
    });

    // Add all numbering instances
    const instances = this.getAllInstances();
    instances.forEach(instance => {
      children.push(instance.toXML());
    });

    return XMLBuilder.w('numbering', {
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }, children);
  }

  /**
   * Creates a numbering manager with default numbering definitions
   */
  static create(): NumberingManager {
    return new NumberingManager(false);
  }

  /**
   * Creates an empty numbering manager
   */
  static createEmpty(): NumberingManager {
    return new NumberingManager(false);
  }
}
