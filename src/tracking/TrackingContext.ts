/**
 * TrackingContext - Interface for automatic change tracking
 *
 * When enabled via Document.enableTrackChanges(), element setter methods
 * use this context to automatically create Revision objects for changes.
 *
 * @module TrackingContext
 */

import type { Revision, RevisionType } from '../elements/Revision';
import type { RevisionManager } from '../elements/RevisionManager';
import type { Run } from '../elements/Run';
import type { Paragraph } from '../elements/Paragraph';
import type { Table } from '../elements/Table';
import type { TableRow } from '../elements/TableRow';
import type { TableCell } from '../elements/TableCell';
import type { Section } from '../elements/Section';

/** Union of element types that can be tracked */
export type TrackableElement =
  | Run
  | Paragraph
  | Table
  | TableRow
  | TableCell
  | Section
  | { constructor: { name: string } };

/**
 * Pending change entry before flushing to RevisionManager
 */
export interface PendingChange {
  /** Type of revision to create */
  type: RevisionType;
  /** Property that changed */
  property: string;
  /** Value before the change */
  previousValue: unknown;
  /** Value after the change */
  newValue: unknown;
  /** Element that was modified */
  element: TrackableElement;
  /** When the change occurred */
  timestamp: number;
  /** Count for consolidated changes */
  count?: number;
}

/**
 * Interface for tracking changes to document elements.
 *
 * Elements call tracking methods when their setters are invoked.
 * The context decides whether to create revisions based on enabled state.
 */
export interface TrackingContext {
  /**
   * Check if change tracking is currently enabled
   */
  isEnabled(): boolean;

  /**
   * Get the author name for new revisions
   */
  getAuthor(): string;

  /**
   * Get the RevisionManager for registering revisions
   */
  getRevisionManager(): RevisionManager;

  /**
   * Check if formatting changes should be tracked
   */
  isTrackFormattingEnabled(): boolean;

  /**
   * Track a Run property change (bold, italic, font, color, etc.)
   * @param run - The Run that was modified
   * @param property - Property name (e.g., 'bold', 'color')
   * @param oldValue - Value before the change
   * @param newValue - Value after the change
   */
  trackRunPropertyChange(run: Run, property: string, oldValue: unknown, newValue: unknown): void;

  /**
   * Track a Paragraph property change (alignment, spacing, etc.)
   * @param paragraph - The Paragraph that was modified
   * @param property - Property name (e.g., 'alignment', 'spaceBefore')
   * @param oldValue - Value before the change
   * @param newValue - Value after the change
   */
  trackParagraphPropertyChange(
    paragraph: TrackableElement,
    property: string,
    oldValue: unknown,
    newValue: unknown
  ): void;

  /**
   * Track a Hyperlink change (URL, anchor, text, formatting)
   * @param hyperlink - The Hyperlink that was modified
   * @param changeType - Type of change (e.g., 'url', 'text', 'formatting')
   * @param oldValue - Value before the change
   * @param newValue - Value after the change
   */
  trackHyperlinkChange(
    hyperlink: TrackableElement,
    changeType: string,
    oldValue: unknown,
    newValue: unknown
  ): void;

  /**
   * Track a Table element property change
   * @param element - The table/row/cell that was modified
   * @param property - Property name
   * @param oldValue - Value before the change
   * @param newValue - Value after the change
   */
  trackTableChange(
    element: TrackableElement,
    property: string,
    oldValue: unknown,
    newValue: unknown
  ): void;

  /**
   * Track a Section property change
   * @param section - The Section that was modified
   * @param property - Property name (e.g., 'pageSize', 'margins')
   * @param oldValue - Value before the change
   * @param newValue - Value after the change
   */
  trackSectionChange(
    section: TrackableElement,
    property: string,
    oldValue: unknown,
    newValue: unknown
  ): void;

  /**
   * Track text insertion
   * @param element - Element containing the insertion
   * @param text - Text that was inserted
   */
  trackInsertion(element: TrackableElement, text: string): void;

  /**
   * Track text deletion
   * @param element - Element containing the deletion
   * @param text - Text that was deleted
   */
  trackDeletion(element: TrackableElement, text: string): void;

  /**
   * Flush all pending changes and create Revision objects.
   * This is called automatically before document save.
   * @returns Array of created revisions
   */
  flushPendingChanges(): Revision[];

  /**
   * Create an insertion revision (factory to avoid circular dependency in Run)
   * @param content - Run or array of content for the insertion
   * @param date - Optional date for the revision
   */
  createInsertion(content: Run, date?: Date): Revision;

  /**
   * Create a deletion revision (factory to avoid circular dependency in Run)
   * @param content - Run or array of content for the deletion
   * @param date - Optional date for the revision
   */
  createDeletion(content: Run, date?: Date): Revision;
}
