/**
 * RelationshipManager - Manages collections of relationships
 *
 * Handles relationship creation, tracking, and XML generation for various
 * document parts (document.xml, header.xml, footer.xml, etc.)
 */

import { Relationship, RelationshipType } from './Relationship';

/**
 * Manages relationships for a document or document part
 */
export class RelationshipManager {
  private relationships: Map<string, Relationship>;
  private nextId: number;

  /**
   * Creates a new relationship manager
   */
  constructor() {
    this.relationships = new Map();
    this.nextId = 1;
  }

  /**
   * Adds a relationship
   * @param relationship The relationship to add
   * @returns The relationship that was added
   */
  addRelationship(relationship: Relationship): Relationship {
    this.relationships.set(relationship.getId(), relationship);

    // Update next ID if necessary
    const idMatch = relationship.getId().match(/^rId(\d+)$/);
    if (idMatch && idMatch[1]) {
      const idNum = parseInt(idMatch[1], 10);
      if (idNum >= this.nextId) {
        this.nextId = idNum + 1;
      }
    }

    return relationship;
  }

  /**
   * Gets a relationship by ID
   * @param id The relationship ID
   */
  getRelationship(id: string): Relationship | undefined {
    return this.relationships.get(id);
  }

  /**
   * Gets all relationships
   */
  getAllRelationships(): Relationship[] {
    return Array.from(this.relationships.values());
  }

  /**
   * Gets relationships of a specific type
   * @param type The relationship type
   */
  getRelationshipsByType(type: string | RelationshipType): Relationship[] {
    return this.getAllRelationships().filter(rel => rel.getType() === type);
  }

  /**
   * Checks if a relationship exists
   * @param id The relationship ID
   */
  hasRelationship(id: string): boolean {
    return this.relationships.has(id);
  }

  /**
   * Removes a relationship
   * @param id The relationship ID
   * @returns True if removed, false if not found
   */
  removeRelationship(id: string): boolean {
    return this.relationships.delete(id);
  }

  /**
   * Gets the number of relationships
   */
  getCount(): number {
    return this.relationships.size;
  }

  /**
   * Clears all relationships
   */
  clear(): this {
    this.relationships.clear();
    this.nextId = 1;
    return this;
  }

  /**
   * Generates a new unique relationship ID
   * @returns New relationship ID (e.g., 'rId1', 'rId2')
   */
  generateId(): string {
    return `rId${this.nextId++}`;
  }

  /**
   * Adds a styles relationship
   * @returns The created relationship
   */
  addStyles(): Relationship {
    const id = this.generateId();
    return this.addRelationship(Relationship.createStyles(id));
  }

  /**
   * Adds a numbering relationship
   * @returns The created relationship
   */
  addNumbering(): Relationship {
    const id = this.generateId();
    return this.addRelationship(Relationship.createNumbering(id));
  }

  /**
   * Adds an image relationship
   * @param target Image path relative to the part (e.g., 'media/image1.png')
   * @returns The created relationship
   */
  addImage(target: string): Relationship {
    const id = this.generateId();
    return this.addRelationship(Relationship.createImage(id, target));
  }

  /**
   * Adds a header relationship
   * @param target Header file path (e.g., 'header1.xml')
   * @returns The created relationship
   */
  addHeader(target: string): Relationship {
    const id = this.generateId();
    return this.addRelationship(Relationship.createHeader(id, target));
  }

  /**
   * Adds a footer relationship
   * @param target Footer file path (e.g., 'footer1.xml')
   * @returns The created relationship
   */
  addFooter(target: string): Relationship {
    const id = this.generateId();
    return this.addRelationship(Relationship.createFooter(id, target));
  }

  /**
   * Adds a hyperlink relationship
   * @param url The hyperlink URL
   * @returns The created relationship
   */
  addHyperlink(url: string): Relationship {
    const id = this.generateId();
    return this.addRelationship(Relationship.createHyperlink(id, url));
  }

  /**
   * Updates the target URL of an existing hyperlink relationship
   *
   * This method modifies an existing relationship's target in-place, maintaining
   * the same relationship ID. This is crucial for proper OpenXML compliance
   * per ECMA-376 §17.16.22, as it prevents orphaned relationships.
   *
   * @param relationshipId The ID of the relationship to update
   * @param newUrl The new URL to set
   * @returns True if updated, false if relationship not found
   */
  updateHyperlinkTarget(relationshipId: string, newUrl: string): boolean {
    const relationship = this.getRelationship(relationshipId);
    if (!relationship) {
      return false;
    }

    // Verify this is a hyperlink relationship
    if (relationship.getType() !== RelationshipType.HYPERLINK) {
      throw new Error(
        `Relationship ${relationshipId} is not a hyperlink relationship. ` +
        `Type is ${relationship.getType()}, expected ${RelationshipType.HYPERLINK}`
      );
    }

    // Update the target URL
    relationship.setTarget(newUrl);
    return true;
  }

  /**
   * Finds a hyperlink relationship by its target URL
   *
   * @param targetUrl The URL to search for
   * @returns The matching relationship, or undefined if not found
   */
  findHyperlinkByTarget(targetUrl: string): Relationship | undefined {
    return this.getAllRelationships().find(
      rel => rel.getType() === RelationshipType.HYPERLINK && rel.getTarget() === targetUrl
    );
  }

  /**
   * Gets or creates a hyperlink relationship for the given URL
   *
   * This method ensures we don't create duplicate relationships for the same URL.
   * If a relationship already exists for the URL, it returns the existing one.
   * Otherwise, it creates a new relationship.
   *
   * @param url The hyperlink URL
   * @returns The existing or newly created relationship
   */
  getOrCreateHyperlink(url: string): Relationship {
    // Check if relationship already exists for this URL
    const existing = this.findHyperlinkByTarget(url);
    if (existing) {
      return existing;
    }

    // Create new relationship
    return this.addHyperlink(url);
  }

  /**
   * Removes orphaned hyperlink relationships
   *
   * This method removes hyperlink relationships that are no longer referenced
   * by any hyperlink in the document. Call this after updating URLs to clean
   * up any orphaned relationships.
   *
   * @param referencedIds Set of relationship IDs that are still in use
   * @returns Number of relationships removed
   */
  removeOrphanedHyperlinks(referencedIds: Set<string>): number {
    let removed = 0;
    const toRemove: string[] = [];

    // Find orphaned relationships
    for (const rel of this.getAllRelationships()) {
      if (rel.getType() === RelationshipType.HYPERLINK && !referencedIds.has(rel.getId())) {
        toRemove.push(rel.getId());
      }
    }

    // Remove orphaned relationships
    for (const id of toRemove) {
      if (this.removeRelationship(id)) {
        removed++;
      }
    }

    return removed;
  }

  /**
   * Adds a comments relationship
   * @returns The created relationship
   */
  addComments(): Relationship {
    const id = this.generateId();
    return this.addRelationship(Relationship.createComments(id));
  }

  /**
   * Generates the relationships XML file content
   * @returns Complete XML string for .rels file
   */
  generateXml(): string {
    const relationships = this.getAllRelationships();

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';

    for (const rel of relationships) {
      xml += rel.toXML() + '\n';
    }

    xml += '</Relationships>';

    return xml;
  }

  /**
   * Creates a new relationship manager with common document relationships
   * @returns RelationshipManager with styles and numbering relationships
   */
  static createForDocument(): RelationshipManager {
    const manager = new RelationshipManager();
    manager.addStyles();
    manager.addNumbering();
    return manager;
  }

  /**
   * Creates an empty relationship manager
   * @returns Empty RelationshipManager
   */
  static create(): RelationshipManager {
    return new RelationshipManager();
  }

  /**
   * Parses relationships from XML string and creates a populated manager
   * @param xml The relationships XML content (.rels file)
   * @returns RelationshipManager with parsed relationships
   */
  static fromXml(xml: string): RelationshipManager {
    const manager = new RelationshipManager();

    // Prevent ReDoS: validate input size (typical .rels files are < 10KB)
    if (xml.length > 100000) {
      throw new Error('Relationships XML file too large (>100KB). Possible malicious input or corrupted file.');
    }

    // Simple XML parsing using regex (sufficient for .rels files)
    // Match all Relationship elements
    // Use non-backtracking pattern with lazy quantifier
    const relationshipPattern = /<Relationship\s+(.*?)\/>/g;
    let match;
    let iterations = 0;
    const MAX_ITERATIONS = 1000; // Prevent infinite loops

    while ((match = relationshipPattern.exec(xml)) !== null) {
      iterations++;
      if (iterations > MAX_ITERATIONS) {
        throw new Error('Too many relationships in XML file (>1000). Possible malicious input.');
      }

      const attributesString = match[1];

      // Skip if no attributes string found
      if (!attributesString) {
        continue;
      }

      // Extract attributes
      const id = this.extractAttribute(attributesString, 'Id');
      const type = this.extractAttribute(attributesString, 'Type');
      const target = this.extractAttribute(attributesString, 'Target');
      const targetMode = this.extractAttribute(attributesString, 'TargetMode');

      // Only create relationship if all required attributes present
      if (id !== undefined && type !== undefined && target !== undefined) {
        // Validate targetMode before type assertion
        const validatedTargetMode =
          targetMode === 'Internal' || targetMode === 'External' || targetMode === undefined
            ? (targetMode as 'Internal' | 'External' | undefined)
            : undefined;

        // Create and add relationship
        const relationship = Relationship.create({
          id,
          type,
          target,
          targetMode: validatedTargetMode || 'Internal',
        });

        manager.addRelationship(relationship);
      }
    }

    return manager;
  }

  /**
   * Extracts an attribute value from an XML attributes string
   * @param attributesString The attributes string
   * @param attributeName The attribute name to extract
   * @returns The attribute value or undefined if not found
   */
  private static extractAttribute(attributesString: string, attributeName: string): string | undefined {
    // Prevent ReDoS: validate input size first (relationships attributes are typically < 1KB)
    if (attributesString.length > 10000) {
      return undefined;
    }

    // Escape attribute name to prevent regex injection
    const escapedName = attributeName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    // Match: AttributeName="value" or AttributeName='value'
    // Use lazy quantifier (.*?) instead of greedy [^"']+ to prevent backtracking
    const pattern = new RegExp(`${escapedName}=["'](.*?)["']`, 'i');
    const match = attributesString.match(pattern);
    return match ? match[1] : undefined;
  }
}
