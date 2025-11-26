/**
 * ChangelogGenerator - Generates structured changelog from Word tracked changes
 *
 * Converts Word revisions (w:ins, w:del, property changes) into structured
 * changelog data with support for consolidation, categorization, and
 * multiple output formats.
 *
 * Follows ECMA-376 revision semantics.
 *
 * @module ChangelogGenerator
 */

import type { Document } from '../core/Document';
import { Revision, RevisionType } from '../elements/Revision';

/**
 * Semantic category for grouping changes.
 */
export type ChangeCategory =
  | 'content'      // Text insertions, deletions
  | 'formatting'   // Run/paragraph property changes
  | 'structural'   // Moves, section changes
  | 'table';       // Table structure changes

/**
 * Location of a change within the document.
 */
export interface ChangeLocation {
  /** Section index (0-based) */
  sectionIndex?: number;
  /** Paragraph index within body (0-based) */
  paragraphIndex: number;
  /** Run index within paragraph (0-based) */
  runIndex?: number;
  /** Nearest heading for context */
  nearestHeading?: string;
  /** Character offset within paragraph */
  characterOffset?: number;
}

/**
 * Represents a single change entry in the changelog.
 * Follows ECMA-376 revision semantics.
 */
export interface ChangeEntry {
  /** Unique identifier (matches revision ID) */
  id: string;

  /** ECMA-376 revision type */
  revisionType: RevisionType;

  /** Semantic category for grouping */
  category: ChangeCategory;

  /** Human-readable description */
  description: string;

  /** Author who made the change */
  author: string;

  /** ISO 8601 timestamp */
  date: Date;

  /** Location in document */
  location: ChangeLocation;

  /** Content details */
  content: {
    /** Text before change (for deletions/modifications) */
    before?: string;
    /** Text after change (for insertions/modifications) */
    after?: string;
    /** Affected text (for property changes) */
    affectedText?: string;
  };

  /** Property change details (for formatting changes) */
  propertyChange?: {
    property: string;
    oldValue?: string;
    newValue?: string;
  };
}

/**
 * Options for changelog generation.
 */
export interface ChangelogOptions {
  /** Include formatting/property changes (default: true) */
  includeFormattingChanges?: boolean;
  /** Consolidate similar changes (default: false) */
  consolidate?: boolean;
  /** Maximum context length for descriptions (default: 50) */
  maxContextLength?: number;
  /** Filter by authors */
  filterAuthors?: string[];
  /** Filter by date range */
  filterDateRange?: { start: Date; end: Date };
  /** Filter by categories */
  filterCategories?: ChangeCategory[];
}

/**
 * Consolidated change grouping similar changes together.
 */
export interface ConsolidatedChange {
  /** Description of the consolidated change */
  description: string;
  /** Number of individual changes */
  count: number;
  /** Category of changes */
  category: ChangeCategory;
  /** Common attributes shared by all changes */
  commonAttributes: {
    author?: string;
    revisionType?: RevisionType;
    propertyChanged?: string;
    newValue?: string;
  };
  /** Individual change IDs */
  changeIds: string[];
}

/**
 * Summary statistics for changelog entries.
 */
export interface ChangelogSummary {
  /** Total number of changes */
  total: number;
  /** Breakdown by category */
  byCategory: Record<ChangeCategory, number>;
  /** Breakdown by revision type */
  byType: Record<string, number>;
  /** Breakdown by author */
  byAuthor: Record<string, number>;
  /** Date range of changes */
  dateRange: { earliest: Date; latest: Date } | null;
}

/**
 * Generates changelog from Word tracked changes.
 * Follows ECMA-376 revision semantics.
 */
export class ChangelogGenerator {
  /**
   * Generate changelog entries from a document.
   * Document must be loaded with { revisionHandling: 'preserve' }.
   *
   * @param doc - Document to extract revisions from
   * @param options - Changelog generation options
   * @returns Array of changelog entries
   */
  static fromDocument(doc: Document, options?: ChangelogOptions): ChangeEntry[] {
    const revisionManager = doc.getRevisionManager();
    if (!revisionManager) {
      return [];
    }

    const revisions = revisionManager.getAllRevisions();
    return this.fromRevisions(revisions, options, doc);
  }

  /**
   * Generate changelog entries from specific revisions.
   *
   * @param revisions - Array of revisions to convert
   * @param options - Changelog generation options
   * @param doc - Optional document for context (paragraph indices, headings)
   * @returns Array of changelog entries
   */
  static fromRevisions(
    revisions: Revision[],
    options?: ChangelogOptions,
    doc?: Document
  ): ChangeEntry[] {
    const opts = {
      includeFormattingChanges: true,
      consolidate: false,
      maxContextLength: 50,
      ...options,
    };

    const entries: ChangeEntry[] = [];

    for (let i = 0; i < revisions.length; i++) {
      const revision = revisions[i];
      if (!revision) continue;

      const category = this.categorize(revision);

      // Filter by category
      if (opts.filterCategories && !opts.filterCategories.includes(category)) {
        continue;
      }

      // Filter out formatting changes if requested
      if (!opts.includeFormattingChanges && category === 'formatting') {
        continue;
      }

      // Filter by author
      if (opts.filterAuthors && !opts.filterAuthors.includes(revision.getAuthor())) {
        continue;
      }

      // Filter by date range
      if (opts.filterDateRange) {
        const revDate = revision.getDate();
        if (revDate < opts.filterDateRange.start || revDate > opts.filterDateRange.end) {
          continue;
        }
      }

      const entry = this.revisionToEntry(revision, i, opts.maxContextLength);
      entries.push(entry);
    }

    return entries;
  }

  /**
   * Convert a single revision to a changelog entry.
   *
   * @param revision - Revision to convert
   * @param index - Index for paragraph location (default location)
   * @param maxContextLength - Maximum length for text context
   * @returns Changelog entry
   */
  private static revisionToEntry(
    revision: Revision,
    index: number,
    maxContextLength: number
  ): ChangeEntry {
    const type = revision.getType();
    const category = this.categorize(revision);
    const runs = revision.getRuns();

    // Extract text content from runs
    const text = runs.map(r => r.getText()).join('');
    const truncatedText = text.length > maxContextLength
      ? text.substring(0, maxContextLength) + '...'
      : text;

    // Build content object based on revision type
    const content: ChangeEntry['content'] = {};
    if (type === 'insert' || type === 'moveTo') {
      content.after = truncatedText;
    } else if (type === 'delete' || type === 'moveFrom') {
      content.before = truncatedText;
    } else if (this.isPropertyChangeType(type)) {
      content.affectedText = truncatedText;
    }

    // Build property change details if applicable
    let propertyChange: ChangeEntry['propertyChange'] | undefined;
    const prevProps = revision.getPreviousProperties();
    const newProps = revision.getNewProperties();
    if (prevProps || newProps) {
      // Get the first property that changed
      const allKeys = new Set([
        ...Object.keys(prevProps || {}),
        ...Object.keys(newProps || {}),
      ]);
      const firstKey = Array.from(allKeys)[0];
      if (firstKey) {
        propertyChange = {
          property: firstKey,
          oldValue: prevProps?.[firstKey]?.toString(),
          newValue: newProps?.[firstKey]?.toString(),
        };
      }
    }

    return {
      id: revision.getId().toString(),
      revisionType: type,
      category,
      description: this.describeRevision(revision, maxContextLength),
      author: revision.getAuthor(),
      date: revision.getDate(),
      location: {
        paragraphIndex: index, // Default; would need document context for accurate location
      },
      content,
      propertyChange,
    };
  }

  /**
   * Get summary statistics for changelog entries.
   *
   * @param entries - Array of changelog entries
   * @returns Summary statistics
   */
  static getSummary(entries: ChangeEntry[]): ChangelogSummary {
    const byCategory: Record<ChangeCategory, number> = {
      content: 0,
      formatting: 0,
      structural: 0,
      table: 0,
    };
    const byType: Record<string, number> = {};
    const byAuthor: Record<string, number> = {};
    let earliest: Date | null = null;
    let latest: Date | null = null;

    for (const entry of entries) {
      // Count by category
      byCategory[entry.category]++;

      // Count by type
      byType[entry.revisionType] = (byType[entry.revisionType] || 0) + 1;

      // Count by author
      byAuthor[entry.author] = (byAuthor[entry.author] || 0) + 1;

      // Track date range
      if (!earliest || entry.date < earliest) {
        earliest = entry.date;
      }
      if (!latest || entry.date > latest) {
        latest = entry.date;
      }
    }

    return {
      total: entries.length,
      byCategory,
      byType,
      byAuthor,
      dateRange: earliest && latest ? { earliest, latest } : null,
    };
  }

  /**
   * Consolidate similar changes into groups.
   * Groups changes that share: same type, same property, same new value.
   *
   * @param entries - Array of changelog entries
   * @returns Array of consolidated changes
   */
  static consolidate(entries: ChangeEntry[]): ConsolidatedChange[] {
    const groups = new Map<string, ChangeEntry[]>();

    for (const entry of entries) {
      // Create grouping key
      let key = `${entry.revisionType}_${entry.category}`;

      // For property changes, include the property name and new value
      if (entry.propertyChange) {
        key += `_${entry.propertyChange.property}_${entry.propertyChange.newValue || ''}`;
      }

      // For content changes by same author, group them
      key += `_${entry.author}`;

      if (!groups.has(key)) {
        groups.set(key, []);
      }
      groups.get(key)!.push(entry);
    }

    const consolidated: ConsolidatedChange[] = [];

    for (const [_, groupEntries] of groups) {
      const first = groupEntries[0];
      if (!first) continue;

      let description: string;

      if (groupEntries.length === 1) {
        description = first.description;
      } else {
        // Generate consolidated description
        description = this.generateConsolidatedDescription(groupEntries);
      }

      consolidated.push({
        description,
        count: groupEntries.length,
        category: first.category,
        commonAttributes: {
          author: this.allSame(groupEntries.map(e => e.author)) ? first.author : undefined,
          revisionType: first.revisionType,
          propertyChanged: first.propertyChange?.property,
          newValue: first.propertyChange?.newValue,
        },
        changeIds: groupEntries.map(e => e.id),
      });
    }

    // Sort by count descending
    consolidated.sort((a, b) => b.count - a.count);

    return consolidated;
  }

  /**
   * Generate a consolidated description for a group of similar changes.
   */
  private static generateConsolidatedDescription(entries: ChangeEntry[]): string {
    const first = entries[0];
    if (!first) {
      return `${entries.length} changes`;
    }

    const count = entries.length;

    switch (first.revisionType) {
      case 'insert':
        return `Inserted text in ${count} locations`;
      case 'delete':
        return `Deleted text from ${count} locations`;
      case 'moveFrom':
      case 'moveTo':
        return `Moved content (${count} operations)`;
      case 'runPropertiesChange':
        if (first.propertyChange) {
          return `Changed ${first.propertyChange.property} to "${first.propertyChange.newValue}" (${count} times)`;
        }
        return `Changed run formatting (${count} times)`;
      case 'paragraphPropertiesChange':
        if (first.propertyChange) {
          return `Changed paragraph ${first.propertyChange.property} (${count} times)`;
        }
        return `Changed paragraph formatting (${count} times)`;
      case 'tablePropertiesChange':
      case 'tableRowPropertiesChange':
      case 'tableCellPropertiesChange':
        return `Changed table formatting (${count} times)`;
      case 'tableCellInsert':
        return `Inserted ${count} table cells`;
      case 'tableCellDelete':
        return `Deleted ${count} table cells`;
      case 'tableCellMerge':
        return `Merged table cells (${count} operations)`;
      case 'numberingChange':
        return `Changed list numbering (${count} times)`;
      case 'sectionPropertiesChange':
        return `Changed section properties (${count} times)`;
      default:
        return `${count} changes of type ${first.revisionType}`;
    }
  }

  /**
   * Check if all values in an array are the same.
   */
  private static allSame<T>(arr: T[]): boolean {
    if (arr.length === 0) return true;
    const first = arr[0];
    return arr.every(v => v === first);
  }

  /**
   * Categorize a revision into a semantic category.
   *
   * @param revision - Revision to categorize
   * @returns Semantic category
   */
  static categorize(revision: Revision): ChangeCategory {
    const type = revision.getType();

    switch (type) {
      // Content changes
      case 'insert':
      case 'delete':
        return 'content';

      // Structural changes
      case 'moveFrom':
      case 'moveTo':
      case 'sectionPropertiesChange':
        return 'structural';

      // Formatting changes
      case 'runPropertiesChange':
      case 'paragraphPropertiesChange':
      case 'numberingChange':
        return 'formatting';

      // Table changes
      case 'tablePropertiesChange':
      case 'tableExceptionPropertiesChange':
      case 'tableRowPropertiesChange':
      case 'tableCellPropertiesChange':
      case 'tableCellInsert':
      case 'tableCellDelete':
      case 'tableCellMerge':
        return 'table';

      default:
        return 'content';
    }
  }

  /**
   * Check if a revision type is a property change type.
   */
  private static isPropertyChangeType(type: RevisionType): boolean {
    return [
      'runPropertiesChange',
      'paragraphPropertiesChange',
      'tablePropertiesChange',
      'tableExceptionPropertiesChange',
      'tableRowPropertiesChange',
      'tableCellPropertiesChange',
      'sectionPropertiesChange',
      'numberingChange',
    ].includes(type);
  }

  /**
   * Generate human-readable description for a revision.
   *
   * @param revision - Revision to describe
   * @param maxLength - Maximum length for text excerpts
   * @returns Human-readable description
   */
  static describeRevision(revision: Revision, maxLength: number = 50): string {
    const type = revision.getType();
    const author = revision.getAuthor();
    const runs = revision.getRuns();
    const text = runs.map(r => r.getText()).join('');
    const excerpt = text.length > maxLength
      ? `"${text.substring(0, maxLength)}..."`
      : text ? `"${text}"` : '';

    switch (type) {
      case 'insert':
        return excerpt ? `Inserted ${excerpt}` : 'Inserted content';
      case 'delete':
        return excerpt ? `Deleted ${excerpt}` : 'Deleted content';
      case 'moveFrom':
        return excerpt ? `Moved ${excerpt} from here` : 'Moved content from here';
      case 'moveTo':
        return excerpt ? `Moved ${excerpt} to here` : 'Moved content to here';
      case 'runPropertiesChange':
        return this.describePropertyChange(revision, 'run formatting');
      case 'paragraphPropertiesChange':
        return this.describePropertyChange(revision, 'paragraph formatting');
      case 'tablePropertiesChange':
        return 'Changed table properties';
      case 'tableExceptionPropertiesChange':
        return 'Changed table exception properties';
      case 'tableRowPropertiesChange':
        return 'Changed table row properties';
      case 'tableCellPropertiesChange':
        return 'Changed table cell properties';
      case 'sectionPropertiesChange':
        return 'Changed section properties';
      case 'tableCellInsert':
        return 'Inserted table cell';
      case 'tableCellDelete':
        return 'Deleted table cell';
      case 'tableCellMerge':
        return 'Merged table cells';
      case 'numberingChange':
        return 'Changed list numbering';
      default:
        return `Changed (${type})`;
    }
  }

  /**
   * Generate description for a property change revision.
   */
  private static describePropertyChange(revision: Revision, context: string): string {
    const prevProps = revision.getPreviousProperties();
    const newProps = revision.getNewProperties();

    if (!prevProps && !newProps) {
      return `Changed ${context}`;
    }

    // Get meaningful property names
    const propNames: string[] = [];
    const allKeys = new Set([
      ...Object.keys(prevProps || {}),
      ...Object.keys(newProps || {}),
    ]);

    for (const key of allKeys) {
      const oldVal = prevProps?.[key];
      const newVal = newProps?.[key];
      if (oldVal !== newVal) {
        propNames.push(this.friendlyPropertyName(key));
      }
    }

    if (propNames.length === 0) {
      return `Changed ${context}`;
    }

    if (propNames.length === 1) {
      return `Changed ${propNames[0]}`;
    }

    if (propNames.length <= 3) {
      return `Changed ${propNames.join(', ')}`;
    }

    return `Changed ${propNames.slice(0, 2).join(', ')} and ${propNames.length - 2} more`;
  }

  /**
   * Convert property key to friendly name.
   */
  private static friendlyPropertyName(key: string): string {
    const friendlyNames: Record<string, string> = {
      b: 'bold',
      i: 'italic',
      u: 'underline',
      strike: 'strikethrough',
      sz: 'font size',
      color: 'text color',
      highlight: 'highlight',
      rFonts: 'font',
      jc: 'alignment',
      ind: 'indentation',
      spacing: 'spacing',
      pStyle: 'paragraph style',
      rStyle: 'character style',
      numPr: 'list numbering',
    };

    return friendlyNames[key] || key;
  }

  /**
   * Export changelog to Markdown format.
   *
   * @param entries - Array of changelog entries
   * @param options - Export options
   * @returns Markdown string
   */
  static toMarkdown(
    entries: ChangeEntry[],
    options?: { includeMetadata?: boolean }
  ): string {
    const opts = { includeMetadata: true, ...options };
    const lines: string[] = [];

    lines.push('# Document Changes');
    lines.push('');

    if (opts.includeMetadata) {
      const summary = this.getSummary(entries);
      lines.push(`**Total Changes:** ${summary.total}`);
      lines.push('');

      if (summary.dateRange) {
        lines.push(`**Date Range:** ${summary.dateRange.earliest.toLocaleDateString()} - ${summary.dateRange.latest.toLocaleDateString()}`);
        lines.push('');
      }

      const authors = Object.keys(summary.byAuthor);
      if (authors.length > 0) {
        lines.push(`**Authors:** ${authors.join(', ')}`);
        lines.push('');
      }

      lines.push('---');
      lines.push('');
    }

    // Group by category
    const byCategory = new Map<ChangeCategory, ChangeEntry[]>();
    for (const entry of entries) {
      if (!byCategory.has(entry.category)) {
        byCategory.set(entry.category, []);
      }
      byCategory.get(entry.category)!.push(entry);
    }

    const categoryTitles: Record<ChangeCategory, string> = {
      content: 'Content Changes',
      formatting: 'Formatting Changes',
      structural: 'Structural Changes',
      table: 'Table Changes',
    };

    for (const [category, categoryEntries] of byCategory) {
      if (categoryEntries.length === 0) continue;

      lines.push(`## ${categoryTitles[category]}`);
      lines.push('');

      for (const entry of categoryEntries) {
        const date = entry.date.toLocaleDateString();
        lines.push(`- ${entry.description} *(${entry.author}, ${date})*`);

        if (entry.content.before) {
          lines.push(`  - Removed: "${entry.content.before}"`);
        }
        if (entry.content.after) {
          lines.push(`  - Added: "${entry.content.after}"`);
        }
      }

      lines.push('');
    }

    return lines.join('\n');
  }

  /**
   * Export changelog to plain text format.
   *
   * @param entries - Array of changelog entries
   * @returns Plain text string
   */
  static toPlainText(entries: ChangeEntry[]): string {
    const lines: string[] = [];

    lines.push('DOCUMENT CHANGES');
    lines.push('================');
    lines.push('');

    for (const entry of entries) {
      const date = entry.date.toLocaleDateString();
      lines.push(`[${date}] ${entry.author}: ${entry.description}`);

      if (entry.content.before) {
        lines.push(`  - ${entry.content.before}`);
      }
      if (entry.content.after) {
        lines.push(`  + ${entry.content.after}`);
      }
    }

    return lines.join('\n');
  }

  /**
   * Export changelog to JSON (for programmatic consumption).
   *
   * @param entries - Array of changelog entries
   * @returns JSON string
   */
  static toJSON(entries: ChangeEntry[]): string {
    return JSON.stringify({
      generated: new Date().toISOString(),
      summary: this.getSummary(entries),
      entries,
    }, null, 2);
  }
}
