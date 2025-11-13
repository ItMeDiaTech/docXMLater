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
  | 'tableExceptionPropertiesChange' // w:tblPrExChange - Table exception properties change
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
  private isFieldInstruction: boolean = false;

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
   * Marks this revision as a field instruction deletion
   * When true, uses w:delInstrText instead of w:delText
   */
  setAsFieldInstruction(): this {
    this.isFieldInstruction = true;
    return this;
  }

  /**
   * Checks if this is a field instruction deletion
   */
  isFieldInstructionDeletion(): boolean {
    return this.isFieldInstruction;
  }

  /**
   * Formats a date to ISO 8601 format for XML
   */
  private formatDate(date: Date): string {
    return date.toISOString();
  }

  /**
   * Gets the XML element name for this revision type
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
      case 'tableExceptionPropertiesChange':
        return 'w:tblPrExChange';
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
   * Generates XML for this revision
   * @returns XMLElement representing the revision
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
   */
  private isPropertyChangeType(): boolean {
    return [
      'runPropertiesChange',
      'paragraphPropertiesChange',
      'tablePropertiesChange',
      'tableExceptionPropertiesChange',
      'tableRowPropertiesChange',
      'tableCellPropertiesChange',
      'sectionPropertiesChange',
      'numberingChange',
    ].includes(this.type);
  }

  /**
   * Creates XML element for previous properties
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
      case 'tableExceptionPropertiesChange':
        propElementName = 'w:tblPrEx';
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
   * Creates XML for a deleted run (uses w:delText or w:delInstrText instead of w:t)
   */
  private createDeletedRunXml(run: Run): XMLElement {
    // Get the regular run XML
    const runXml = run.toXML();

    // Determine which element to use for deleted text
    // w:delInstrText for field instructions, w:delText for regular text
    const deletedTextElement = this.isFieldInstruction ? 'w:delInstrText' : 'w:delText';

    // We need to replace w:t elements with w:delText or w:delInstrText
    if (runXml.children) {
      const modifiedChildren = runXml.children.map(child => {
        if (typeof child === 'object' && child.name === 'w:t') {
          // Replace w:t with appropriate deleted text element
          return {
            ...child,
            name: deletedTextElement,
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
   * Creates a field instruction deletion revision
   * Uses w:delInstrText instead of w:delText for field codes
   * @param author - Author who made the deletion
   * @param content - Deleted field instruction content (Run or array of Runs)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createFieldInstructionDeletion(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    const revision = new Revision({
      author,
      type: 'delete',
      content,
      date,
    });
    revision.setAsFieldInstruction();
    return revision;
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
   * Creates a table exception properties change revision
   * Tracks changes to table properties that override style defaults
   * @param author - Author who made the change
   * @param content - Table content
   * @param previousProperties - Previous table exception properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTableExceptionPropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'tableExceptionPropertiesChange',
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
