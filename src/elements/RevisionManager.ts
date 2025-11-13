/**
 * RevisionManager - Manages tracked changes (revisions) in a document
 *
 * Tracks all revisions, assigns unique IDs, and provides statistics.
 */

import { Revision, RevisionType } from './Revision';

/**
 * Manages document revisions (track changes)
 */
export class RevisionManager {
  private revisions: Revision[] = [];
  private nextId: number = 0;

  /**
   * Registers a revision with the manager
   * Assigns a unique ID
   * @param revision - Revision to register
   * @returns The registered revision (same instance)
   */
  register(revision: Revision): Revision {
    // Assign unique ID
    revision.setId(this.nextId++);

    // Store revision
    this.revisions.push(revision);

    return revision;
  }

  /**
   * Gets all revisions
   * @returns Array of all revisions
   */
  getAllRevisions(): Revision[] {
    return [...this.revisions];
  }

  /**
   * Gets revisions by type
   * @param type - Revision type to filter by
   * @returns Array of revisions of the specified type
   */
  getRevisionsByType(type: RevisionType): Revision[] {
    return this.revisions.filter(rev => rev.getType() === type);
  }

  /**
   * Gets revisions by author
   * @param author - Author name to filter by
   * @returns Array of revisions by the specified author
   */
  getRevisionsByAuthor(author: string): Revision[] {
    return this.revisions.filter(rev => rev.getAuthor() === author);
  }

  /**
   * Gets the number of revisions
   * @returns Number of revisions
   */
  getCount(): number {
    return this.revisions.length;
  }

  /**
   * Gets the number of insertions
   * @returns Number of insertion revisions
   */
  getInsertionCount(): number {
    return this.getRevisionsByType('insert').length;
  }

  /**
   * Gets the number of deletions
   * @returns Number of deletion revisions
   */
  getDeletionCount(): number {
    return this.getRevisionsByType('delete').length;
  }

  /**
   * Gets all unique authors who have made changes
   * @returns Array of unique author names
   */
  getAuthors(): string[] {
    const authorsSet = new Set<string>();
    for (const revision of this.revisions) {
      authorsSet.add(revision.getAuthor());
    }
    return Array.from(authorsSet);
  }

  /**
   * Clears all revisions
   */
  clear(): void {
    this.revisions = [];
    this.nextId = 0;
  }

  /**
   * Checks if there are no revisions
   * @returns True if there are no tracked changes
   */
  isEmpty(): boolean {
    return this.revisions.length === 0;
  }

  /**
   * Gets the most recent N revisions
   * @param count - Number of recent revisions to return
   * @returns Array of most recent revisions
   */
  getRecentRevisions(count: number): Revision[] {
    return [...this.revisions]
      .sort((a, b) => b.getDate().getTime() - a.getDate().getTime())
      .slice(0, count);
  }

  /**
   * Searches revisions by text content
   * @param searchText - Text to search for (case-insensitive)
   * @returns Array of revisions containing the search text
   */
  findRevisionsByText(searchText: string): Revision[] {
    const lowerSearch = searchText.toLowerCase();
    return this.revisions.filter(revision => {
      const text = revision.getRuns()
        .map(run => run.getText())
        .join('')
        .toLowerCase();
      return text.includes(lowerSearch);
    });
  }

  /**
   * Gets all insertions (added text)
   * @returns Array of insertion revisions
   */
  getAllInsertions(): Revision[] {
    return this.getRevisionsByType('insert');
  }

  /**
   * Gets all deletions (removed text)
   * @returns Array of deletion revisions
   */
  getAllDeletions(): Revision[] {
    return this.getRevisionsByType('delete');
  }

  /**
   * Gets all run properties changes (formatting changes)
   * @returns Array of run property change revisions
   */
  getAllRunPropertiesChanges(): Revision[] {
    return this.getRevisionsByType('runPropertiesChange');
  }

  /**
   * Gets all paragraph properties changes
   * @returns Array of paragraph property change revisions
   */
  getAllParagraphPropertiesChanges(): Revision[] {
    return this.getRevisionsByType('paragraphPropertiesChange');
  }

  /**
   * Gets all table properties changes
   * @returns Array of table property change revisions
   */
  getAllTablePropertiesChanges(): Revision[] {
    return this.getRevisionsByType('tablePropertiesChange');
  }

  /**
   * Gets all move operations (both moveFrom and moveTo)
   * @returns Array of move-related revisions
   */
  getAllMoves(): Revision[] {
    return this.revisions.filter(rev =>
      rev.getType() === 'moveFrom' || rev.getType() === 'moveTo'
    );
  }

  /**
   * Gets all moveFrom revisions (source of moves)
   * @returns Array of moveFrom revisions
   */
  getAllMoveFrom(): Revision[] {
    return this.getRevisionsByType('moveFrom');
  }

  /**
   * Gets all moveTo revisions (destination of moves)
   * @returns Array of moveTo revisions
   */
  getAllMoveTo(): Revision[] {
    return this.getRevisionsByType('moveTo');
  }

  /**
   * Gets all table cell changes (insert, delete, merge)
   * @returns Array of table cell change revisions
   */
  getAllTableCellChanges(): Revision[] {
    return this.revisions.filter(rev =>
      rev.getType() === 'tableCellInsert' ||
      rev.getType() === 'tableCellDelete' ||
      rev.getType() === 'tableCellMerge'
    );
  }

  /**
   * Gets all numbering changes
   * @returns Array of numbering change revisions
   */
  getAllNumberingChanges(): Revision[] {
    return this.getRevisionsByType('numberingChange');
  }

  /**
   * Gets all property change revisions (run, paragraph, table, etc.)
   * @returns Array of all property change revisions
   */
  getAllPropertyChanges(): Revision[] {
    return this.revisions.filter(rev =>
      rev.getType() === 'runPropertiesChange' ||
      rev.getType() === 'paragraphPropertiesChange' ||
      rev.getType() === 'tablePropertiesChange' ||
      rev.getType() === 'tableRowPropertiesChange' ||
      rev.getType() === 'tableCellPropertiesChange' ||
      rev.getType() === 'sectionPropertiesChange' ||
      rev.getType() === 'numberingChange'
    );
  }

  /**
   * Gets move pair by move ID
   * @param moveId - Move operation ID
   * @returns Object with moveFrom and moveTo revisions (if found)
   */
  getMovePair(moveId: string): { moveFrom?: Revision; moveTo?: Revision } {
    const moveFrom = this.revisions.find(
      rev => rev.getType() === 'moveFrom' && rev.getMoveId() === moveId
    );
    const moveTo = this.revisions.find(
      rev => rev.getType() === 'moveTo' && rev.getMoveId() === moveId
    );
    return { moveFrom, moveTo };
  }

  /**
   * Gets statistics about revisions
   * @returns Object with revision statistics
   */
  getStats(): {
    total: number;
    insertions: number;
    deletions: number;
    propertyChanges: number;
    moves: number;
    tableCellChanges: number;
    authors: string[];
    nextId: number;
  } {
    return {
      total: this.revisions.length,
      insertions: this.getInsertionCount(),
      deletions: this.getDeletionCount(),
      propertyChanges: this.getAllPropertyChanges().length,
      moves: this.getAllMoves().length,
      tableCellChanges: this.getAllTableCellChanges().length,
      authors: this.getAuthors(),
      nextId: this.nextId,
    };
  }

  /**
   * Checks if track changes is enabled (has any revisions)
   * @returns True if there are revisions
   */
  isTrackingChanges(): boolean {
    return this.revisions.length > 0;
  }

  /**
   * Gets the most recent revision
   * @returns The most recent revision, or undefined if no revisions
   */
  getLatestRevision(): Revision | undefined {
    if (this.revisions.length === 0) {
      return undefined;
    }
    return this.revisions[this.revisions.length - 1];
  }

  /**
   * Gets revisions within a date range
   * @param startDate - Start of date range
   * @param endDate - End of date range
   * @returns Array of revisions within the date range
   */
  getRevisionsByDateRange(startDate: Date, endDate: Date): Revision[] {
    return this.revisions.filter(rev => {
      const revDate = rev.getDate();
      return revDate >= startDate && revDate <= endDate;
    });
  }

  /**
   * Creates a new RevisionManager
   * @returns New RevisionManager instance
   */
  static create(): RevisionManager {
    return new RevisionManager();
  }
}
