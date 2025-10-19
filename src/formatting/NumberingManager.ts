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

/**
 * Manages numbering definitions and instances for a document
 */
export class NumberingManager {
  private abstractNumberings: Map<number, AbstractNumbering>;
  private instances: Map<number, NumberingInstance>;
  private nextAbstractNumId: number;
  private nextNumId: number;

  /**
   * Creates a new numbering manager
   * @param initializeDefaults Whether to initialize with default numbering definitions
   */
  constructor(initializeDefaults: boolean = false) {
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

    instancesToRemove.forEach(numId => this.instances.delete(numId));

    return this.abstractNumberings.delete(abstractNumId);
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
    return this.instances.delete(numId);
  }

  /**
   * Creates a new bullet list and returns its numId
   * @param levels Number of levels (default: 9)
   * @param bullets Array of bullet characters
   */
  createBulletList(levels: number = 9, bullets?: string[]): number {
    // Create abstract numbering
    const abstractNumId = this.nextAbstractNumId++;
    const abstractNumbering = AbstractNumbering.createBulletList(abstractNumId, levels, bullets);
    this.addAbstractNumbering(abstractNumbering);

    // Create instance
    const numId = this.nextNumId++;
    const instance = NumberingInstance.create({ numId, abstractNumId });
    this.addInstance(instance);

    return numId;
  }

  /**
   * Creates a new numbered list and returns its numId
   * @param levels Number of levels (default: 3)
   * @param formats Array of formats for each level
   */
  createNumberedList(
    levels: number = 3,
    formats?: Array<'decimal' | 'lowerLetter' | 'lowerRoman'>
  ): number {
    // Create abstract numbering
    const abstractNumId = this.nextAbstractNumId++;
    const abstractNumbering = AbstractNumbering.createNumberedList(abstractNumId, levels, formats);
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
