/**
 * Revision - Represents tracked changes in a Word document
 *
 * Track changes allow tracking of insertions, deletions, and modifications
 * to document content, showing who made changes and when.
 */

import { Run } from './Run';
import { XMLElement } from '../xml/XMLBuilder';

/**
 * Revision type
 */
export type RevisionType = 'insert' | 'delete';

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
   * Formats a date to ISO 8601 format for XML
   */
  private formatDate(date: Date): string {
    return date.toISOString();
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

    const elementName = this.type === 'insert' ? 'w:ins' : 'w:del';
    const children: XMLElement[] = [];

    // Add runs to the revision
    for (const run of this.runs) {
      if (this.type === 'delete') {
        // For deletions, we need to modify the run XML to use w:delText instead of w:t
        const runXml = this.createDeletedRunXml(run);
        children.push(runXml);
      } else {
        // For insertions, use normal run XML
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
   * Creates XML for a deleted run (uses w:delText instead of w:t)
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
}
