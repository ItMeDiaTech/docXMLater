/**
 * Revision - Represents tracked changes in a Word document
 *
 * Track changes allow tracking of insertions, deletions, and modifications
 * to document content, showing who made changes and when.
 */

import { Run } from './Run';
import { XMLElement } from '../xml/XMLBuilder';

/**
 * Revision type - All OpenXML WordprocessingML revision types
 */
export type RevisionType =
  // Content changes
  | 'insert'              // w:ins - Inserted content
  | 'delete'              // w:del - Deleted content
  // Property changes
  | 'runPropertiesChange' // w:rPrChange - Run formatting change (bold, italic, font, etc.)
  | 'paragraphPropertiesChange' // w:pPrChange - Paragraph formatting change
  | 'tablePropertiesChange'     // w:tblPrChange - Table formatting change
  | 'tableRowPropertiesChange'  // w:trPrChange - Table row properties change
  | 'tableCellPropertiesChange' // w:tcPrChange - Table cell properties change
  | 'sectionPropertiesChange'   // w:sectPrChange - Section properties change
  // Move operations
  | 'moveFrom'            // w:moveFrom - Content moved from this location
  | 'moveTo'              // w:moveTo - Content moved to this location
  // Table operations
  | 'tableCellInsert'     // w:cellIns - Table cell inserted
  | 'tableCellDelete'     // w:cellDel - Table cell deleted
  | 'tableCellMerge'      // w:cellMerge - Table cells merged
  // Numbering
  | 'numberingChange';    // w:numberingChange - List numbering changed

/**
 * Revision properties
 */
export interface RevisionProperties {
  /** Unique revision ID (assigned by RevisionManager) */
  id?: number;
  /** Author who made the change */
  author: string;
  /** Date when the change was made */
  date?: Date;
  /** Type of revision */
  type: RevisionType;
  /** Content affected by the revision */
  content: Run | Run[];
  /** Previous properties (for property change revisions) */
  previousProperties?: Record<string, any>;
  /** New properties (for property change revisions) */
  newProperties?: Record<string, any>;
  /** Move ID (for moveFrom/moveTo operations) */
  moveId?: string;
  /** Destination location (for moveFrom) or source location (for moveTo) */
  moveLocation?: string;
}

/**
 * Represents a tracked change (revision) in a document
 */
export class Revision {
  private id: number;
  private author: string;
  private date: Date;
  private type: RevisionType;
  private runs: Run[];
  private previousProperties?: Record<string, any>;
  private newProperties?: Record<string, any>;
  private moveId?: string;
  private moveLocation?: string;

  /**
   * Creates a new Revision
   * @param properties - Revision properties
   */
  constructor(properties: RevisionProperties) {
    this.id = properties.id ?? 0; // Will be assigned by RevisionManager
    this.author = properties.author;
    this.date = properties.date || new Date();
    this.type = properties.type;
    this.runs = Array.isArray(properties.content) ? properties.content : [properties.content];
    this.previousProperties = properties.previousProperties;
    this.newProperties = properties.newProperties;
    this.moveId = properties.moveId;
    this.moveLocation = properties.moveLocation;
  }

  /**
   * Gets the revision ID
   */
  getId(): number {
    return this.id;
  }

  /**
   * Sets the revision ID (used by RevisionManager)
   * @internal
   */
  setId(id: number): void {
    this.id = id;
  }

  /**
   * Gets the author
   */
  getAuthor(): string {
    return this.author;
  }

  /**
   * Sets the author
   */
  setAuthor(author: string): this {
    this.author = author;
    return this;
  }

  /**
   * Gets the revision date
   */
  getDate(): Date {
    return this.date;
  }

  /**
   * Sets the revision date
   */
  setDate(date: Date): this {
    this.date = date;
    return this;
  }

  /**
   * Gets the revision type
   */
  getType(): RevisionType {
    return this.type;
  }

  /**
   * Gets the runs affected by this revision
   */
  getRuns(): Run[] {
    return [...this.runs];
  }

  /**
   * Adds a run to this revision
   */
  addRun(run: Run): this {
    this.runs.push(run);
    return this;
  }

  /**
   * Gets the previous properties (for property change revisions)
   */
  getPreviousProperties(): Record<string, any> | undefined {
    return this.previousProperties;
  }

  /**
   * Gets the new properties (for property change revisions)
   */
  getNewProperties(): Record<string, any> | undefined {
    return this.newProperties;
  }

  /**
   * Gets the move ID (for moveFrom/moveTo operations)
   */
  getMoveId(): string | undefined {
    return this.moveId;
  }

  /**
   * Gets the move location
   */
  getMoveLocation(): string | undefined {
    return this.moveLocation;
  }

  /**
   * Formats a date to ISO 8601 format for XML
   * Per ECMA-376, revision dates must be in ISO 8601 format (e.g., "2024-01-01T12:00:00Z")
   * @param date - Date to format
   * @returns ISO 8601 formatted date string
   */
  private formatDate(date: Date): string {
    return date.toISOString();
  }

  /**
   * Gets the XML element name for this revision type
   * Maps internal revision types to OOXML WordprocessingML element names per ECMA-376
   *
   * Mappings:
   * - insert → w:ins (inserted content)
   * - delete → w:del (deleted content)
   * - runPropertiesChange → w:rPrChange (run formatting change)
   * - paragraphPropertiesChange → w:pPrChange (paragraph formatting change)
   * - tablePropertiesChange → w:tblPrChange (table formatting change)
   * - tableRowPropertiesChange → w:trPrChange (table row properties change)
   * - tableCellPropertiesChange → w:tcPrChange (table cell properties change)
   * - sectionPropertiesChange → w:sectPrChange (section properties change)
   * - moveFrom → w:moveFrom (source location of moved content)
   * - moveTo → w:moveTo (destination location of moved content)
   * - tableCellInsert → w:cellIns (inserted table cell)
   * - tableCellDelete → w:cellDel (deleted table cell)
   * - tableCellMerge → w:cellMerge (merged table cells)
   * - numberingChange → w:numberingChange (list numbering changed)
   *
   * @returns OOXML element name (e.g., "w:ins", "w:del")
   */
  private getElementName(): string {
    switch (this.type) {
      case 'insert':
        return 'w:ins';
      case 'delete':
        return 'w:del';
      case 'runPropertiesChange':
        return 'w:rPrChange';
      case 'paragraphPropertiesChange':
        return 'w:pPrChange';
      case 'tablePropertiesChange':
        return 'w:tblPrChange';
      case 'tableRowPropertiesChange':
        return 'w:trPrChange';
      case 'tableCellPropertiesChange':
        return 'w:tcPrChange';
      case 'sectionPropertiesChange':
        return 'w:sectPrChange';
      case 'moveFrom':
        return 'w:moveFrom';
      case 'moveTo':
        return 'w:moveTo';
      case 'tableCellInsert':
        return 'w:cellIns';
      case 'tableCellDelete':
        return 'w:cellDel';
      case 'tableCellMerge':
        return 'w:cellMerge';
      case 'numberingChange':
        return 'w:numberingChange';
      default:
        return 'w:ins'; // Fallback to insert
    }
  }

  /**
   * Generates XML for this revision per OOXML WordprocessingML specification (ECMA-376)
   *
   * **XML Structure:**
   *
   * Content revisions (w:ins, w:del, w:moveFrom, w:moveTo):
   * ```xml
   * <w:ins w:id="0" w:author="Author Name" w:date="2024-01-01T12:00:00Z">
   *   <w:r>
   *     <w:t>Inserted text</w:t>
   *   </w:r>
   * </w:ins>
   * ```
   *
   * Deletion revisions use w:delText instead of w:t:
   * ```xml
   * <w:del w:id="1" w:author="Author Name" w:date="2024-01-01T12:00:00Z">
   *   <w:r>
   *     <w:delText>Deleted text</w:delText>
   *   </w:r>
   * </w:del>
   * ```
   *
   * Property change revisions (w:rPrChange, w:pPrChange, etc.):
   * ```xml
   * <w:rPrChange w:id="2" w:author="Author Name" w:date="2024-01-01T12:00:00Z">
   *   <w:rPr>
   *     <w:b/>  <!-- Previous bold setting -->
   *     <w:sz w:val="24"/>  <!-- Previous font size -->
   *   </w:rPr>
   * </w:rPrChange>
   * ```
   *
   * **Required Attributes (per ECMA-376):**
   * - w:id: Unique revision identifier (ST_DecimalNumber) - REQUIRED
   * - w:author: Author who made the change (ST_String) - REQUIRED
   * - w:date: When the change was made (ST_DateTime, ISO 8601) - OPTIONAL
   *
   * **Move Operations:**
   * For moveFrom/moveTo, an additional w:moveId attribute links the source and destination:
   * ```xml
   * <w:moveFrom w:id="3" w:author="Author" w:date="..." w:moveId="move-1">...</w:moveFrom>
   * <w:moveTo w:id="4" w:author="Author" w:date="..." w:moveId="move-1">...</w:moveTo>
   * ```
   *
   * **Content vs Property Changes:**
   * - Content revisions (insert/delete/move): Contain w:r elements with text runs
   * - Property revisions (rPrChange/pPrChange): Contain previous property elements (w:rPr, w:pPr)
   *
   * @returns XMLElement representing the revision in OOXML format
   * @see ECMA-376 Part 1 §17.13.5 (Revision Identifiers for Paragraph Content)
   * @see ECMA-376 Part 1 §17.13.5.15 (Inserted Paragraph)
   * @see ECMA-376 Part 1 §17.13.5.14 (Deleted Paragraph)
   */
  toXML(): XMLElement {
    const attributes: Record<string, string> = {
      'w:id': this.id.toString(),
      'w:author': this.author,
      'w:date': this.formatDate(this.date),
    };

    // Add move-specific attributes
    if ((this.type === 'moveFrom' || this.type === 'moveTo') && this.moveId) {
      attributes['w:moveId'] = this.moveId;
    }

    const elementName = this.getElementName();
    const children: XMLElement[] = [];

    // Handle different revision types
    if (this.isPropertyChangeType()) {
      // Property change revisions contain the previous properties
      if (this.previousProperties) {
        children.push(this.createPropertiesElement());
      }
    }

    // Add runs to the revision
    for (const run of this.runs) {
      if (this.type === 'delete' || this.type === 'moveFrom') {
        // For deletions and moveFrom, we need to modify the run XML to use w:delText instead of w:t
        const runXml = this.createDeletedRunXml(run);
        children.push(runXml);
      } else {
        // For other types, use normal run XML
        children.push(run.toXML());
      }
    }

    return {
      name: elementName,
      attributes,
      children,
    };
  }

  /**
   * Checks if this is a property change revision type
   *
   * Property change revisions track formatting changes, not content changes.
   * They contain previous property elements (w:rPr, w:pPr, etc.) instead of text runs.
   *
   * **Property Change Types:**
   * - runPropertiesChange: Run formatting (bold, italic, font, color, etc.)
   * - paragraphPropertiesChange: Paragraph formatting (alignment, spacing, indentation, etc.)
   * - tablePropertiesChange: Table formatting
   * - tableRowPropertiesChange: Table row properties
   * - tableCellPropertiesChange: Table cell properties
   * - sectionPropertiesChange: Section properties (page size, margins, etc.)
   * - numberingChange: List numbering properties
   *
   * **Content Change Types (NOT property changes):**
   * - insert: Added text
   * - delete: Removed text
   * - moveFrom: Moved text source
   * - moveTo: Moved text destination
   * - tableCellInsert: Added table cell
   * - tableCellDelete: Removed table cell
   * - tableCellMerge: Merged table cells
   *
   * @returns true if this revision tracks a property/formatting change, false otherwise
   */
  private isPropertyChangeType(): boolean {
    return [
      'runPropertiesChange',
      'paragraphPropertiesChange',
      'tablePropertiesChange',
      'tableRowPropertiesChange',
      'tableCellPropertiesChange',
      'sectionPropertiesChange',
      'numberingChange',
    ].includes(this.type);
  }

  /**
   * Creates XML element for previous properties in property change revisions
   *
   * **Purpose:**
   * Property change revisions (w:rPrChange, w:pPrChange, etc.) must contain a child element
   * with the PREVIOUS state of the properties before the change. This allows Word to show
   * what changed and enables accepting/rejecting the change.
   *
   * **Structure:**
   * ```xml
   * <w:rPrChange w:id="0" w:author="Author" w:date="...">
   *   <w:rPr>
   *     <!-- Previous run properties -->
   *     <w:b/>  <!-- Was bold -->
   *     <w:sz w:val="24"/>  <!-- Was 12pt (24 half-points) -->
   *   </w:rPr>
   * </w:rPrChange>
   * ```
   *
   * **Property Element Mapping:**
   * - runPropertiesChange → w:rPr (run properties)
   * - paragraphPropertiesChange → w:pPr (paragraph properties)
   * - tablePropertiesChange → w:tblPr (table properties)
   * - tableRowPropertiesChange → w:trPr (table row properties)
   * - tableCellPropertiesChange → w:tcPr (table cell properties)
   * - sectionPropertiesChange → w:sectPr (section properties)
   * - numberingChange → w:numPr (numbering properties)
   *
   * **Implementation:**
   * This method converts the previousProperties object into OOXML elements.
   * - Boolean properties (e.g., bold) → <w:b/>
   * - Value properties (e.g., font size) → <w:sz w:val="24"/>
   *
   * @returns XMLElement containing previous properties (w:rPr, w:pPr, etc.)
   * @see ECMA-376 Part 1 §17.13.5.31 (Run Properties Change)
   * @see ECMA-376 Part 1 §17.13.5.29 (Paragraph Properties Change)
   */
  private createPropertiesElement(): XMLElement {
    // The property element name depends on the revision type
    let propElementName = 'w:rPr';

    switch (this.type) {
      case 'runPropertiesChange':
        propElementName = 'w:rPr';
        break;
      case 'paragraphPropertiesChange':
        propElementName = 'w:pPr';
        break;
      case 'tablePropertiesChange':
        propElementName = 'w:tblPr';
        break;
      case 'tableRowPropertiesChange':
        propElementName = 'w:trPr';
        break;
      case 'tableCellPropertiesChange':
        propElementName = 'w:tcPr';
        break;
      case 'sectionPropertiesChange':
        propElementName = 'w:sectPr';
        break;
      case 'numberingChange':
        propElementName = 'w:numPr';
        break;
    }

    // Build property children from previousProperties
    const propChildren: XMLElement[] = [];
    if (this.previousProperties) {
      for (const [key, value] of Object.entries(this.previousProperties)) {
        if (typeof value === 'boolean' && value) {
          // Boolean properties (e.g., bold, italic)
          propChildren.push({ name: `w:${key}`, attributes: {}, children: [] });
        } else if (typeof value === 'string' || typeof value === 'number') {
          // Value properties (e.g., font size, color)
          propChildren.push({
            name: `w:${key}`,
            attributes: { 'w:val': value.toString() },
            children: [],
          });
        }
      }
    }

    return {
      name: propElementName,
      attributes: {},
      children: propChildren,
    };
  }

  /**
   * Creates XML for a deleted run (uses w:delText instead of w:t)
   *
   * **OOXML Requirement:**
   * Per ECMA-376, deleted text must use w:delText element instead of w:t element.
   * This is required for proper rendering in Microsoft Word's Track Changes mode.
   *
   * **Transformation:**
   * ```xml
   * <!-- Normal run (NOT in deletion) -->
   * <w:r>
   *   <w:rPr><w:b/></w:rPr>
   *   <w:t>Text</w:t>
   * </w:r>
   *
   * <!-- Deleted run (inside w:del) -->
   * <w:r>
   *   <w:rPr><w:b/></w:rPr>
   *   <w:delText>Text</w:delText>
   * </w:r>
   * ```
   *
   * **Why This Matters:**
   * - w:delText tells Word to render with strikethrough in Track Changes mode
   * - w:t would render as normal text even inside w:del element
   * - Word will reject documents with w:t inside deletions as malformed
   *
   * **Implementation:**
   * This method gets the run's normal XML and replaces all w:t elements with w:delText
   * while preserving all other properties (formatting, spacing attributes, etc.)
   *
   * @param run - Run containing deleted text
   * @returns XMLElement with w:delText instead of w:t
   * @see ECMA-376 Part 1 §17.13.5.14 (Deleted Run Content)
   * @see ECMA-376 Part 1 §22.1.2.27 (w:delText element)
   */
  private createDeletedRunXml(run: Run): XMLElement {
    // Get the regular run XML
    const runXml = run.toXML();

    // We need to replace w:t elements with w:delText
    if (runXml.children) {
      const modifiedChildren = runXml.children.map(child => {
        if (typeof child === 'object' && child.name === 'w:t') {
          // Replace w:t with w:delText
          return {
            ...child,
            name: 'w:delText',
          };
        }
        return child;
      });

      return {
        ...runXml,
        children: modifiedChildren,
      };
    }

    return runXml;
  }

  /**
   * Creates an insertion revision
   * @param author - Author who made the insertion
   * @param content - Inserted content (Run or array of Runs)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createInsertion(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'insert',
      content,
      date,
    });
  }

  /**
   * Creates a deletion revision
   * @param author - Author who made the deletion
   * @param content - Deleted content (Run or array of Runs)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createDeletion(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'delete',
      content,
      date,
    });
  }

  /**
   * Creates a revision from text
   * Convenience method that creates a Run from the text
   * @param type - Revision type
   * @param author - Author who made the change
   * @param text - Text content
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static fromText(
    type: RevisionType,
    author: string,
    text: string,
    date?: Date
  ): Revision {
    const run = new Run(text);
    return new Revision({
      author,
      type,
      content: run,
      date,
    });
  }

  /**
   * Creates a run properties change revision
   * @param author - Author who made the change
   * @param content - Content with changed formatting
   * @param previousProperties - Previous run properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createRunPropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'runPropertiesChange',
      content,
      previousProperties,
      date,
    });
  }

  /**
   * Creates a paragraph properties change revision
   * @param author - Author who made the change
   * @param content - Paragraph content
   * @param previousProperties - Previous paragraph properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createParagraphPropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'paragraphPropertiesChange',
      content,
      previousProperties,
      date,
    });
  }

  /**
   * Creates a table properties change revision
   * @param author - Author who made the change
   * @param content - Table content
   * @param previousProperties - Previous table properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTablePropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'tablePropertiesChange',
      content,
      previousProperties,
      date,
    });
  }

  /**
   * Creates a moveFrom revision (source of moved content)
   * @param author - Author who moved the content
   * @param content - Content that was moved
   * @param moveId - Unique move operation ID (links moveFrom and moveTo)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createMoveFrom(
    author: string,
    content: Run | Run[],
    moveId: string,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'moveFrom',
      content,
      moveId,
      date,
    });
  }

  /**
   * Creates a moveTo revision (destination of moved content)
   * @param author - Author who moved the content
   * @param content - Content that was moved
   * @param moveId - Unique move operation ID (links moveFrom and moveTo)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createMoveTo(
    author: string,
    content: Run | Run[],
    moveId: string,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'moveTo',
      content,
      moveId,
      date,
    });
  }

  /**
   * Creates a table cell insertion revision
   * @param author - Author who inserted the cell
   * @param content - Cell content
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTableCellInsert(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'tableCellInsert',
      content,
      date,
    });
  }

  /**
   * Creates a table cell deletion revision
   * @param author - Author who deleted the cell
   * @param content - Cell content
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTableCellDelete(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'tableCellDelete',
      content,
      date,
    });
  }

  /**
   * Creates a table cell merge revision
   * @param author - Author who merged cells
   * @param content - Merged cell content
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTableCellMerge(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'tableCellMerge',
      content,
      date,
    });
  }

  /**
   * Creates a numbering change revision
   * @param author - Author who changed the numbering
   * @param content - Content with changed numbering
   * @param previousProperties - Previous numbering properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createNumberingChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'numberingChange',
      content,
      previousProperties,
      date,
    });
  }
}
