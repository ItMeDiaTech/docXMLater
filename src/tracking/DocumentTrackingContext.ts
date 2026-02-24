/**
 * DocumentTrackingContext - Implementation of automatic change tracking
 *
 * Manages pending changes and creates Revision objects when flushed.
 * Supports consolidation of similar changes within a time window.
 *
 * @module DocumentTrackingContext
 */

import { Revision, RevisionType } from '../elements/Revision';
import { RevisionManager } from '../elements/RevisionManager';
import { Run, type RunFormatting } from '../elements/Run';
import { Paragraph } from '../elements/Paragraph';
import { Table } from '../elements/Table';
import { TableRow } from '../elements/TableRow';
import { TableCell } from '../elements/TableCell';
import { Section } from '../elements/Section';
import type { TrackingContext, PendingChange, TrackableElement } from './TrackingContext';
import { formatDateForXml } from '../utils/dateFormatting';

/**
 * Enable options for tracking context
 */
export interface TrackingEnableOptions {
  /** Author name for new revisions (default: 'DocHub') */
  author?: string;
  /** Whether to track formatting changes (default: true) */
  trackFormatting?: boolean;
}

/**
 * Implementation of TrackingContext for Document
 */
export class DocumentTrackingContext implements TrackingContext {
  private enabled = false;
  private trackFormatting = true;
  private author = 'DocHub';
  private revisionManager: RevisionManager;

  /** Counter for assigning stable element IDs */
  private elementIdCounter = 0;
  /** Stable element identity map (WeakMap so elements can be GC'd) */
  private elementIdMap = new WeakMap<object, number>();

  /** Pending changes waiting to be flushed */
  private pendingChanges = new Map<string, PendingChange>();

  /** Properties considered "formatting" (vs structural) */
  private static readonly FORMATTING_PROPERTIES = new Set([
    'bold',
    'italic',
    'underline',
    'strike',
    'dstrike',
    'subscript',
    'superscript',
    'font',
    'size',
    'color',
    'highlight',
    'smallCaps',
    'allCaps',
    'characterSpacing',
    'scaling',
    'position',
    'emphasis',
    'shadow',
    'emboss',
    'imprint',
    'outline',
    'vanish',
  ]);

  /**
   * Creates a new DocumentTrackingContext
   * @param revisionManager - RevisionManager to register revisions with
   */
  constructor(revisionManager: RevisionManager) {
    this.revisionManager = revisionManager;
  }

  /**
   * Enable change tracking
   * @param options - Enable options
   */
  enable(options?: TrackingEnableOptions): void {
    this.enabled = true;
    if (options?.author) {
      this.author = options.author;
    }
    if (options?.trackFormatting !== undefined) {
      this.trackFormatting = options.trackFormatting;
    }
  }

  /**
   * Disable change tracking and flush pending changes
   */
  disable(): void {
    this.flushPendingChanges();
    this.enabled = false;
  }

  /**
   * Set the author for new revisions
   * Flushes any pending changes before switching to prevent mixed authorship
   * @param author - Author name
   */
  setAuthor(author: string): void {
    // Flush pending changes before switching authors to prevent mixed authorship
    if (this.enabled && this.pendingChanges.size > 0) {
      this.flushPendingChanges();
    }
    this.author = author;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // TrackingContext Interface Implementation
  // ═══════════════════════════════════════════════════════════════════════════

  isEnabled(): boolean {
    return this.enabled;
  }

  getAuthor(): string {
    return this.author;
  }

  getRevisionManager(): RevisionManager {
    return this.revisionManager;
  }

  isTrackFormattingEnabled(): boolean {
    return this.trackFormatting;
  }

  trackRunPropertyChange(run: Run, property: string, oldValue: unknown, newValue: unknown): void {
    if (!this.enabled) return;
    if (this.valuesEqual(oldValue, newValue)) return;

    // Skip formatting changes if not tracking them
    if (!this.trackFormatting && DocumentTrackingContext.FORMATTING_PROPERTIES.has(property)) {
      return;
    }

    // Create consolidation key with element identity
    const key = `runProp:${property}:${this.stringifyValue(newValue)}@${this.getElementId(run)}`;

    this.addPendingChange(key, {
      type: 'runPropertiesChange',
      property,
      previousValue: oldValue,
      newValue,
      element: run,
      timestamp: Date.now(),
    });
  }

  trackParagraphPropertyChange(
    paragraph: TrackableElement,
    property: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    if (!this.enabled) return;
    if (this.valuesEqual(oldValue, newValue)) return;

    const key = `paraProp:${property}:${this.stringifyValue(newValue)}@${this.getElementId(paragraph)}`;

    this.addPendingChange(key, {
      type: 'paragraphPropertiesChange',
      property,
      previousValue: oldValue,
      newValue,
      element: paragraph,
      timestamp: Date.now(),
    });
  }

  trackHyperlinkChange(
    hyperlink: TrackableElement,
    changeType: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    if (!this.enabled) return;
    if (this.valuesEqual(oldValue, newValue)) return;

    // Hyperlink changes use dedicated type for proper categorization
    const key = `hyperlink:${changeType}:${this.stringifyValue(newValue)}@${this.getElementId(hyperlink)}`;

    this.addPendingChange(key, {
      type: 'hyperlinkChange',
      property: changeType,
      previousValue: oldValue,
      newValue,
      element: hyperlink,
      timestamp: Date.now(),
    });
  }

  trackTableChange(
    element: TrackableElement,
    property: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    if (!this.enabled) return;
    if (this.valuesEqual(oldValue, newValue)) return;

    // Determine revision type based on element type using instanceof (minification-safe)
    let type: RevisionType = 'tablePropertiesChange';
    let elementType = 'Table';

    if (element instanceof TableRow) {
      type = 'tableRowPropertiesChange';
      elementType = 'TableRow';
    } else if (element instanceof TableCell) {
      type = 'tableCellPropertiesChange';
      elementType = 'TableCell';
    }

    const key = `table:${elementType}:${property}:${this.stringifyValue(newValue)}@${this.getElementId(element)}`;

    this.addPendingChange(key, {
      type,
      property,
      previousValue: oldValue,
      newValue,
      element,
      timestamp: Date.now(),
    });
  }

  trackSectionChange(
    section: TrackableElement,
    property: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    if (!this.enabled) return;
    if (this.valuesEqual(oldValue, newValue)) return;

    const key = `section:${property}:${this.stringifyValue(newValue)}@${this.getElementId(section)}`;

    this.addPendingChange(key, {
      type: 'sectionPropertiesChange',
      property,
      previousValue: oldValue,
      newValue,
      element: section,
      timestamp: Date.now(),
    });
  }

  trackInsertion(element: TrackableElement, text: string): void {
    if (!this.enabled) return;
    if (!text) return;

    const key = `insert:${Date.now()}:${text.substring(0, 20)}`;

    this.addPendingChange(key, {
      type: 'insert',
      property: 'text',
      previousValue: undefined,
      newValue: text,
      element,
      timestamp: Date.now(),
    });
  }

  trackDeletion(element: TrackableElement, text: string): void {
    if (!this.enabled) return;
    if (!text) return;

    const key = `delete:${Date.now()}:${text.substring(0, 20)}`;

    this.addPendingChange(key, {
      type: 'delete',
      property: 'text',
      previousValue: text,
      newValue: undefined,
      element,
      timestamp: Date.now(),
    });
  }

  flushPendingChanges(): Revision[] {
    const revisions: Revision[] = [];

    // Group pending changes by element for consolidation
    const paragraphChanges = new Map<Paragraph, PendingChange[]>();
    const tableChanges = new Map<Table, PendingChange[]>();
    const rowChanges = new Map<TableRow, PendingChange[]>();
    const cellChanges = new Map<TableCell, PendingChange[]>();
    const sectionChanges = new Map<Section, PendingChange[]>();
    const runChanges = new Map<Run, PendingChange[]>();

    for (const [, pending] of this.pendingChanges) {
      const revision = this.createRevision(pending);
      this.revisionManager.register(revision);
      revisions.push(revision);

      // Collect changes by element type for *PrChange application
      if (pending.type === 'paragraphPropertiesChange' && pending.element instanceof Paragraph) {
        const changes = paragraphChanges.get(pending.element) || [];
        changes.push(pending);
        paragraphChanges.set(pending.element, changes);
      } else if (pending.type === 'tablePropertiesChange' && pending.element instanceof Table) {
        const changes = tableChanges.get(pending.element) || [];
        changes.push(pending);
        tableChanges.set(pending.element, changes);
      } else if (
        pending.type === 'tableRowPropertiesChange' &&
        pending.element instanceof TableRow
      ) {
        const changes = rowChanges.get(pending.element) || [];
        changes.push(pending);
        rowChanges.set(pending.element, changes);
      } else if (
        pending.type === 'tableCellPropertiesChange' &&
        pending.element instanceof TableCell
      ) {
        const changes = cellChanges.get(pending.element) || [];
        changes.push(pending);
        cellChanges.set(pending.element, changes);
      } else if (pending.type === 'sectionPropertiesChange' && pending.element instanceof Section) {
        const changes = sectionChanges.get(pending.element) || [];
        changes.push(pending);
        sectionChanges.set(pending.element, changes);
      } else if (pending.type === 'runPropertiesChange' && pending.element instanceof Run) {
        const changes = runChanges.get(pending.element) || [];
        changes.push(pending);
        runChanges.set(pending.element, changes);
      }
    }

    // Apply pPrChange to each paragraph that has property changes
    for (const [paragraph, changes] of paragraphChanges) {
      this.applyParagraphPrChange(paragraph, changes);
    }

    // Apply tblPrChange to each Table
    // Per ECMA-376 §17.13.5.36, tblPrChange must contain FULL previous tblPr,
    // not just the delta of changed properties.
    for (const [table, changes] of tableChanges) {
      // Build full snapshot: start from current formatting, roll back changed properties
      const currentFormatting = table.getFormatting();
      const fullPrevProps: Record<string, unknown> = {};

      for (const [key, value] of Object.entries(currentFormatting)) {
        if (value !== undefined) {
          fullPrevProps[key] = value;
        }
      }

      // Roll back changed properties to their previous values
      let latestTimestamp = 0;
      for (const change of changes) {
        if (change.previousValue !== undefined) {
          fullPrevProps[change.property] = change.previousValue;
        } else {
          delete fullPrevProps[change.property];
        }
        if (change.timestamp > latestTimestamp) {
          latestTimestamp = change.timestamp;
        }
      }

      const date = formatDateForXml(new Date(latestTimestamp));

      const existing = table.getTblPrChange();
      if (existing) {
        // Merge: existing previous state takes precedence (it's the ORIGINAL baseline)
        const merged = { ...fullPrevProps, ...(existing.previousProperties || {}) };
        table.setTblPrChange({ ...existing, previousProperties: merged });
      } else {
        table.setTblPrChange({
          author: this.author,
          date,
          id: String(this.revisionManager.consumeNextId()),
          previousProperties: fullPrevProps,
        });
      }
    }

    // Apply trPrChange to each TableRow
    for (const [row, changes] of rowChanges) {
      this.applyElementPrChange(changes, (prevProps, getNextId, date) => {
        const existing = row.getTrPrChange();
        if (existing) {
          const merged = { ...(existing.previousProperties || {}), ...prevProps };
          row.setTrPrChange({ ...existing, previousProperties: merged });
        } else {
          row.setTrPrChange({
            author: this.author,
            date,
            id: String(getNextId()),
            previousProperties: prevProps,
          });
        }
      });
    }

    // Apply tcPrChange to each TableCell
    for (const [cell, changes] of cellChanges) {
      this.applyElementPrChange(changes, (prevProps, getNextId, date) => {
        const existing = cell.getTcPrChange();
        if (existing) {
          const merged = { ...(existing.previousProperties || {}), ...prevProps };
          cell.setTcPrChange({ ...existing, previousProperties: merged });
        } else {
          cell.setTcPrChange({
            author: this.author,
            date,
            id: String(getNextId()),
            previousProperties: prevProps,
          });
        }
      });
    }

    // Apply sectPrChange to each Section
    for (const [section, changes] of sectionChanges) {
      this.applyElementPrChange(changes, (prevProps, getNextId, date) => {
        const existing = section.getSectPrChange();
        if (existing) {
          const merged = { ...(existing.previousProperties || {}), ...prevProps };
          section.setSectPrChange({ ...existing, previousProperties: merged });
        } else {
          section.setSectPrChange({
            author: this.author,
            date,
            id: String(getNextId()),
            previousProperties: prevProps,
          });
        }
      });
    }

    // Apply rPrChange to each Run that has property changes
    for (const [run, changes] of runChanges) {
      this.applyRunPrChange(run, changes);
    }

    this.pendingChanges.clear();
    return revisions;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // Private Methods
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Apply pPrChange to a paragraph (extracted from flushPendingChanges for readability)
   */
  private applyParagraphPrChange(paragraph: Paragraph, changes: PendingChange[]): void {
    const newPreviousProperties: Record<string, unknown> = {};
    let latestTimestamp = 0;

    for (const change of changes) {
      // Record previous value even if undefined (property wasn't set before)
      newPreviousProperties[change.property] = change.previousValue;
      if (change.timestamp > latestTimestamp) {
        latestTimestamp = change.timestamp;
      }
    }

    const existingChange = paragraph.formatting.pPrChange;

    if (existingChange) {
      const mergedPreviousProperties: Record<string, unknown> = {
        ...(existingChange.previousProperties || {}),
      };
      for (const [prop, value] of Object.entries(newPreviousProperties)) {
        mergedPreviousProperties[prop] = value;
      }
      paragraph.formatting.pPrChange = {
        author: existingChange.author,
        date: existingChange.date,
        id: existingChange.id,
        previousProperties: mergedPreviousProperties,
      };
    } else {
      const revisionId = this.revisionManager.consumeNextId();
      paragraph.formatting.pPrChange = {
        author: this.author,
        date: formatDateForXml(new Date(latestTimestamp)),
        id: String(revisionId),
        previousProperties: newPreviousProperties,
      };
    }
  }

  /**
   * Apply rPrChange to a run (mirrors applyParagraphPrChange for runs)
   */
  private applyRunPrChange(run: Run, changes: PendingChange[]): void {
    const newPreviousProperties: Partial<RunFormatting> = {};
    let latestTimestamp = 0;

    for (const change of changes) {
      (newPreviousProperties as Record<string, unknown>)[change.property] = change.previousValue;
      if (change.timestamp > latestTimestamp) {
        latestTimestamp = change.timestamp;
      }
    }

    const existingChange = run.getPropertyChangeRevision();

    if (existingChange) {
      // Merge with existing rPrChange — preserve original author/date,
      // add new previous properties
      const mergedPreviousProperties: Partial<RunFormatting> = {
        ...(existingChange.previousProperties || {}),
        ...newPreviousProperties,
      };
      run.setPropertyChangeRevision({
        ...existingChange,
        previousProperties: mergedPreviousProperties,
      });
    } else {
      const revisionId = this.revisionManager.consumeNextId();
      run.setPropertyChangeRevision({
        id: revisionId,
        author: this.author,
        date: new Date(latestTimestamp),
        previousProperties: newPreviousProperties,
      });
    }
  }

  /**
   * Generic helper to apply *PrChange to table/row/cell/section elements
   */
  private applyElementPrChange(
    changes: PendingChange[],
    applier: (prevProps: Record<string, unknown>, getNextId: () => number, date: string) => void
  ): void {
    const prevProps: Record<string, unknown> = {};
    let latestTimestamp = 0;

    for (const change of changes) {
      // Record previous value even if undefined (property wasn't set before)
      prevProps[change.property] = change.previousValue;
      if (change.timestamp > latestTimestamp) {
        latestTimestamp = change.timestamp;
      }
    }

    const date = formatDateForXml(new Date(latestTimestamp));
    applier(prevProps, () => this.revisionManager.consumeNextId(), date);
  }

  /**
   * Add a pending change, consolidating with existing if same key
   */
  private addPendingChange(key: string, change: PendingChange): void {
    const existing = this.pendingChanges.get(key);
    if (existing) {
      existing.count = (existing.count || 1) + 1;
      // Keep the first previousValue for consolidated changes
    } else {
      this.pendingChanges.set(key, { ...change, count: 1 });
    }
  }

  /**
   * Create a Revision from a pending change
   */
  private createRevision(pending: PendingChange): Revision {
    const previousProps: Record<string, any> = {};
    const newProps: Record<string, any> = {};

    if (pending.previousValue !== undefined) {
      previousProps[pending.property] = pending.previousValue;
    }
    if (pending.newValue !== undefined) {
      newProps[pending.property] = pending.newValue;
    }

    // Get or create a Run for the revision content
    const run = this.getRunFromElement(pending.element);

    return new Revision({
      author: this.author,
      type: pending.type,
      content: run,
      previousProperties: Object.keys(previousProps).length > 0 ? previousProps : undefined,
      newProperties: Object.keys(newProps).length > 0 ? newProps : undefined,
      date: new Date(pending.timestamp),
    });
  }

  /**
   * Get or create a Run from an element for revision content
   */
  private getRunFromElement(element: TrackableElement): Run {
    if (element instanceof Run) {
      return element;
    }

    // Use instanceof for type-safe element identification (minification-safe)
    if (element instanceof Table) return new Run('Table');
    if (element instanceof TableRow) return new Run('TableRow');
    if (element instanceof TableCell) return new Run('TableCell');
    if (element instanceof Section) return new Run('Section');
    if (element instanceof Paragraph) return new Run('Paragraph');

    // Fallback for other elements (e.g., Hyperlink)
    const hasGetText =
      'getText' in element && typeof (element as { getText?: () => string }).getText === 'function';
    const text = hasGetText
      ? (element as { getText: () => string }).getText()
      : element?.constructor?.name || 'Unknown element';
    return new Run(typeof text === 'string' ? text : String(text));
  }

  /**
   * Get a stable unique ID for an element (used in consolidation keys)
   */
  private getElementId(element: TrackableElement): number {
    let id = this.elementIdMap.get(element as object);
    if (id === undefined) {
      id = this.elementIdCounter++;
      this.elementIdMap.set(element as object, id);
    }
    return id;
  }

  /**
   * Deep equality check for tracking values (handles objects, primitives, null/undefined)
   */
  private valuesEqual(a: unknown, b: unknown): boolean {
    if (a === b) return true;
    if (a == null || b == null) return false;
    if (typeof a !== 'object' || typeof b !== 'object') return false;
    return JSON.stringify(a) === JSON.stringify(b);
  }

  /**
   * Stringify a value for use in consolidation key
   */
  private stringifyValue(value: unknown): string {
    if (value === undefined) return 'undefined';
    if (value === null) return 'null';
    if (typeof value === 'object') {
      return JSON.stringify(value);
    }
    return String(value);
  }

  /**
   * Create an insertion revision (factory to avoid circular dependency in Run)
   */
  createInsertion(content: Run, date?: Date): Revision {
    return Revision.createInsertion(this.author, content, date);
  }

  /**
   * Create a deletion revision (factory to avoid circular dependency in Run)
   */
  createDeletion(content: Run, date?: Date): Revision {
    return Revision.createDeletion(this.author, content, date);
  }

  /**
   * Get count of pending changes
   */
  getPendingCount(): number {
    return this.pendingChanges.size;
  }

  /**
   * Check if there are pending changes
   */
  hasPendingChanges(): boolean {
    return this.pendingChanges.size > 0;
  }

  /**
   * Clear all pending changes without creating revisions
   */
  clearPendingChanges(): void {
    this.pendingChanges.clear();
  }
}
