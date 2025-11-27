/**
 * RevisionContent - Defines valid content types for tracked changes (revisions)
 *
 * Per ECMA-376, w:ins and w:del elements can contain:
 * - w:r (runs) - Text with formatting
 * - w:hyperlink - Hyperlinks with nested runs
 *
 * This module provides the type definitions and type guards for revision content.
 */

import type { Run } from './Run';
import type { Hyperlink } from './Hyperlink';

/**
 * Content types valid within a revision (tracked change)
 *
 * Per ECMA-376 Part 1 section 17.13.5, revision elements can contain:
 * - Run elements (w:r) - the most common case
 * - Hyperlink elements (w:hyperlink) - for tracked hyperlink changes
 */
export type RevisionContent = Run | Hyperlink;

/**
 * Type guard to check if content is a Run
 * @param content - The content to check
 * @returns true if content is a Run instance
 */
export function isRunContent(content: RevisionContent): content is Run {
  // Use constructor name check to avoid circular imports
  return content?.constructor?.name === 'Run';
}

/**
 * Type guard to check if content is a Hyperlink
 * @param content - The content to check
 * @returns true if content is a Hyperlink instance
 */
export function isHyperlinkContent(content: RevisionContent): content is Hyperlink {
  // Use constructor name check to avoid circular imports
  return content?.constructor?.name === 'Hyperlink';
}
