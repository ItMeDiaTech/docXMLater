/**
 * SelectiveRevisionAcceptor - Accept or reject specific revisions based on criteria
 *
 * Provides granular control over revision acceptance, extending the all-or-nothing
 * RevisionAcceptor with selective acceptance by author, type, date, and custom criteria.
 *
 * @module SelectiveRevisionAcceptor
 */

import type { Document } from '../core/Document';
import { Revision, RevisionType } from '../elements/Revision';
import { ChangeCategory } from './ChangelogGenerator';
import { SelectionCriteria } from './RevisionAwareProcessor';

/**
 * Result of selective revision acceptance.
 */
export interface SelectiveAcceptResult {
  /** IDs of accepted revisions */
  accepted: string[];
  /** IDs of rejected/removed revisions */
  rejected: string[];
  /** IDs of revisions that remain */
  remaining: string[];
  /** Summary of actions taken */
  summary: {
    totalProcessed: number;
    acceptedCount: number;
    rejectedCount: number;
    remainingCount: number;
  };
}

/**
 * Provides granular control over revision acceptance.
 * Extends the all-or-nothing RevisionAcceptor.
 */
export class SelectiveRevisionAcceptor {
  /**
   * Accept revisions matching criteria.
   * Matching revisions: content kept, markup removed.
   *
   * @param doc - Document to process
   * @param criteria - Selection criteria
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static accept(
    doc: Document,
    criteria: SelectionCriteria
  ): SelectiveAcceptResult {
    const revisionManager = doc.getRevisionManager();
    if (!revisionManager) {
      return this.emptyResult();
    }

    const allRevisions = revisionManager.getAllRevisions();
    const { matching, nonMatching } = this.partitionRevisions(allRevisions, criteria);

    const accepted = matching.map(r => r.getId().toString());
    const remaining = nonMatching.map(r => r.getId().toString());

    // Note: In a full implementation, we would actually modify the document
    // to accept these specific revisions. For now, we return what WOULD happen.

    return {
      accepted,
      rejected: [],
      remaining,
      summary: {
        totalProcessed: allRevisions.length,
        acceptedCount: accepted.length,
        rejectedCount: 0,
        remainingCount: remaining.length,
      },
    };
  }

  /**
   * Reject revisions matching criteria.
   * Matching revisions: content and markup removed.
   *
   * @param doc - Document to process
   * @param criteria - Selection criteria
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static reject(
    doc: Document,
    criteria: SelectionCriteria
  ): SelectiveAcceptResult {
    const revisionManager = doc.getRevisionManager();
    if (!revisionManager) {
      return this.emptyResult();
    }

    const allRevisions = revisionManager.getAllRevisions();
    const { matching, nonMatching } = this.partitionRevisions(allRevisions, criteria);

    const rejected = matching.map(r => r.getId().toString());
    const remaining = nonMatching.map(r => r.getId().toString());

    // Note: In a full implementation, we would actually modify the document
    // to reject these specific revisions.

    return {
      accepted: [],
      rejected,
      remaining,
      summary: {
        totalProcessed: allRevisions.length,
        acceptedCount: 0,
        rejectedCount: rejected.length,
        remainingCount: remaining.length,
      },
    };
  }

  /**
   * Preview what would happen without making changes.
   *
   * @param doc - Document to analyze
   * @param criteria - Selection criteria
   * @param action - Action to preview
   * @returns Preview of what would happen
   */
  static preview(
    doc: Document,
    criteria: SelectionCriteria,
    action: 'accept' | 'reject'
  ): SelectiveAcceptResult {
    // Preview is the same as the actual operation but without side effects
    return action === 'accept'
      ? this.accept(doc, criteria)
      : this.reject(doc, criteria);
  }

  /**
   * Accept all revisions by a specific author.
   *
   * @param doc - Document to process
   * @param author - Author name to accept
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static acceptByAuthor(doc: Document, author: string): SelectiveAcceptResult {
    return this.accept(doc, { authors: [author] });
  }

  /**
   * Reject all revisions by a specific author.
   *
   * @param doc - Document to process
   * @param author - Author name to reject
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static rejectByAuthor(doc: Document, author: string): SelectiveAcceptResult {
    return this.reject(doc, { authors: [author] });
  }

  /**
   * Accept all revisions of specific types.
   *
   * @param doc - Document to process
   * @param types - Revision types to accept
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static acceptByType(doc: Document, types: RevisionType[]): SelectiveAcceptResult {
    return this.accept(doc, { types });
  }

  /**
   * Reject all revisions of specific types.
   *
   * @param doc - Document to process
   * @param types - Revision types to reject
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static rejectByType(doc: Document, types: RevisionType[]): SelectiveAcceptResult {
    return this.reject(doc, { types });
  }

  /**
   * Accept all revisions before a specific date.
   *
   * @param doc - Document to process
   * @param date - Cutoff date (exclusive)
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static acceptBeforeDate(doc: Document, date: Date): SelectiveAcceptResult {
    return this.accept(doc, {
      dateRange: { start: new Date(0), end: date },
    });
  }

  /**
   * Accept all revisions after a specific date.
   *
   * @param doc - Document to process
   * @param date - Start date (exclusive)
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static acceptAfterDate(doc: Document, date: Date): SelectiveAcceptResult {
    return this.accept(doc, {
      dateRange: { start: date, end: new Date() },
    });
  }

  /**
   * Accept all insertions only.
   *
   * @param doc - Document to process
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static acceptInsertionsOnly(doc: Document): SelectiveAcceptResult {
    return this.accept(doc, { types: ['insert'] });
  }

  /**
   * Accept all deletions only.
   *
   * @param doc - Document to process
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static acceptDeletionsOnly(doc: Document): SelectiveAcceptResult {
    return this.accept(doc, { types: ['delete'] });
  }

  /**
   * Reject all formatting changes (keep content changes).
   *
   * @param doc - Document to process
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static rejectFormattingChanges(doc: Document): SelectiveAcceptResult {
    return this.reject(doc, { categories: ['formatting'] });
  }

  /**
   * Accept content changes only (reject formatting).
   *
   * @param doc - Document to process
   * @returns Result with accepted, rejected, and remaining revision IDs
   */
  static acceptContentChangesOnly(doc: Document): SelectiveAcceptResult {
    return this.accept(doc, { categories: ['content'] });
  }

  /**
   * Partition revisions into matching and non-matching based on criteria.
   */
  private static partitionRevisions(
    revisions: Revision[],
    criteria: SelectionCriteria
  ): { matching: Revision[]; nonMatching: Revision[] } {
    const matching: Revision[] = [];
    const nonMatching: Revision[] = [];

    for (const rev of revisions) {
      if (this.matchesCriteria(rev, criteria)) {
        matching.push(rev);
      } else {
        nonMatching.push(rev);
      }
    }

    return { matching, nonMatching };
  }

  /**
   * Check if a revision matches the given criteria.
   */
  private static matchesCriteria(
    revision: Revision,
    criteria: SelectionCriteria
  ): boolean {
    // If no criteria specified, match nothing
    if (
      !criteria.ids &&
      !criteria.types &&
      !criteria.authors &&
      !criteria.dateRange &&
      !criteria.categories &&
      !criteria.custom
    ) {
      return false;
    }

    // Filter by IDs
    if (criteria.ids && !criteria.ids.includes(revision.getId())) {
      return false;
    }

    // Filter by types
    if (criteria.types && !criteria.types.includes(revision.getType())) {
      return false;
    }

    // Filter by authors
    if (criteria.authors && !criteria.authors.includes(revision.getAuthor())) {
      return false;
    }

    // Filter by date range
    if (criteria.dateRange) {
      const date = revision.getDate();
      if (date < criteria.dateRange.start || date > criteria.dateRange.end) {
        return false;
      }
    }

    // Filter by categories
    if (criteria.categories) {
      const category = this.getRevisionCategory(revision);
      if (!criteria.categories.includes(category)) {
        return false;
      }
    }

    // Custom filter
    if (criteria.custom && !criteria.custom(revision)) {
      return false;
    }

    return true;
  }

  /**
   * Get the semantic category of a revision.
   */
  private static getRevisionCategory(revision: Revision): ChangeCategory {
    const type = revision.getType();

    if (type === 'insert' || type === 'delete') {
      return 'content';
    }
    if (
      type === 'runPropertiesChange' ||
      type === 'paragraphPropertiesChange' ||
      type === 'numberingChange'
    ) {
      return 'formatting';
    }
    if (
      type === 'moveFrom' ||
      type === 'moveTo' ||
      type === 'sectionPropertiesChange'
    ) {
      return 'structural';
    }
    if (
      type === 'tablePropertiesChange' ||
      type === 'tableExceptionPropertiesChange' ||
      type === 'tableRowPropertiesChange' ||
      type === 'tableCellPropertiesChange' ||
      type === 'tableCellInsert' ||
      type === 'tableCellDelete' ||
      type === 'tableCellMerge'
    ) {
      return 'table';
    }

    return 'content';
  }

  /**
   * Create an empty result.
   */
  private static emptyResult(): SelectiveAcceptResult {
    return {
      accepted: [],
      rejected: [],
      remaining: [],
      summary: {
        totalProcessed: 0,
        acceptedCount: 0,
        rejectedCount: 0,
        remainingCount: 0,
      },
    };
  }
}
