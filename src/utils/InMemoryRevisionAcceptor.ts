/**
 * InMemoryRevisionAcceptor - Accept revisions by transforming the in-memory object model
 *
 * This approach follows the industry standard (OpenXML PowerTools, Aspose.Words):
 * - Transforms Revision objects in paragraph.content[] to their "accepted" state
 * - For insertions: Unwrap - extract child Runs/Hyperlinks into parent paragraph
 * - For deletions: Remove - delete the revision and its content from the model
 * - For property changes: Remove the change metadata, keep the current formatting
 *
 * Unlike the raw XML approach (acceptRevisions.ts), this allows subsequent modifications
 * to the in-memory model to be correctly serialized on save().
 *
 * @see https://github.com/OfficeDev/Open-Xml-PowerTools - RevisionAccepter.cs
 * @see https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/ee836138(v=office.12)
 */

import type { Document } from '../core/Document';
import type { Paragraph, ParagraphContent } from '../elements/Paragraph';
import { Revision, RevisionType } from '../elements/Revision';
import type { Run } from '../elements/Run';
import type { Hyperlink } from '../elements/Hyperlink';
import { isRunContent, isHyperlinkContent } from '../elements/RevisionContent';
import { getGlobalLogger, createScopedLogger, ILogger } from './logger';

/**
 * Get scoped logger for this module
 */
function getLogger(): ILogger {
  return createScopedLogger(getGlobalLogger(), 'InMemoryRevisionAcceptor');
}

/**
 * Options for accepting revisions
 */
export interface AcceptRevisionsOptions {
  /** Accept insertion revisions (w:ins) - default: true */
  acceptInsertions?: boolean;
  /** Accept deletion revisions (w:del) - default: true */
  acceptDeletions?: boolean;
  /** Accept move operations (w:moveFrom, w:moveTo) - default: true */
  acceptMoves?: boolean;
  /** Accept property change revisions (rPrChange, pPrChange, etc.) - default: true */
  acceptPropertyChanges?: boolean;
}

/**
 * Result of accepting revisions
 */
export interface AcceptRevisionsResult {
  /** Number of insertions accepted */
  insertionsAccepted: number;
  /** Number of deletions accepted */
  deletionsAccepted: number;
  /** Number of move operations accepted */
  movesAccepted: number;
  /** Number of property changes accepted */
  propertyChangesAccepted: number;
  /** Total revisions processed */
  totalAccepted: number;
}

/**
 * Revision types that represent content changes (contain actual text/runs)
 */
const CONTENT_REVISION_TYPES: RevisionType[] = [
  'insert',
  'delete',
  'moveFrom',
  'moveTo',
];

/**
 * Revision types that represent property/formatting changes
 */
const PROPERTY_REVISION_TYPES: RevisionType[] = [
  'runPropertiesChange',
  'paragraphPropertiesChange',
  'tablePropertiesChange',
  'tableExceptionPropertiesChange',
  'tableRowPropertiesChange',
  'tableCellPropertiesChange',
  'sectionPropertiesChange',
  'numberingChange',
];

/**
 * Accept all revisions in the document by transforming the in-memory model.
 *
 * This is the industry-standard approach used by OpenXML PowerTools, Aspose.Words,
 * and other production DOCX libraries. It allows subsequent modifications to the
 * document to work correctly.
 *
 * @param doc - Document to process
 * @param options - Options for which revision types to accept
 * @returns Result with counts of accepted revisions
 */
export function acceptRevisionsInMemory(
  doc: Document,
  options: AcceptRevisionsOptions = {}
): AcceptRevisionsResult {
  const logger = getLogger();
  const opts: Required<AcceptRevisionsOptions> = {
    acceptInsertions: options.acceptInsertions ?? true,
    acceptDeletions: options.acceptDeletions ?? true,
    acceptMoves: options.acceptMoves ?? true,
    acceptPropertyChanges: options.acceptPropertyChanges ?? true,
  };

  const result: AcceptRevisionsResult = {
    insertionsAccepted: 0,
    deletionsAccepted: 0,
    movesAccepted: 0,
    propertyChangesAccepted: 0,
    totalAccepted: 0,
  };

  logger.info('Accepting revisions in-memory', { options: opts });

  // Process all paragraphs in the document body
  const paragraphs = doc.getAllParagraphs();
  for (const paragraph of paragraphs) {
    const paragraphResult = acceptRevisionsInParagraph(paragraph, opts);
    result.insertionsAccepted += paragraphResult.insertionsAccepted;
    result.deletionsAccepted += paragraphResult.deletionsAccepted;
    result.movesAccepted += paragraphResult.movesAccepted;
    result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
  }

  // Process paragraphs in tables
  const tables = doc.getTables();
  for (const table of tables) {
    for (const row of table.getRows()) {
      for (const cell of row.getCells()) {
        for (const paragraph of cell.getParagraphs()) {
          const paragraphResult = acceptRevisionsInParagraph(paragraph, opts);
          result.insertionsAccepted += paragraphResult.insertionsAccepted;
          result.deletionsAccepted += paragraphResult.deletionsAccepted;
          result.movesAccepted += paragraphResult.movesAccepted;
          result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
        }
      }
    }
  }

  // Process paragraphs in headers
  const headerFooterManager = doc.getHeaderFooterManager();
  if (headerFooterManager) {
    const headers = headerFooterManager.getAllHeaders();
    for (const headerEntry of headers) {
      const elements = headerEntry.header.getElements();
      for (const element of elements) {
        // Element can be Paragraph or Table
        if ('getContent' in element && typeof element.getContent === 'function') {
          // It's a Paragraph
          const paragraphResult = acceptRevisionsInParagraph(element as Paragraph, opts);
          result.insertionsAccepted += paragraphResult.insertionsAccepted;
          result.deletionsAccepted += paragraphResult.deletionsAccepted;
          result.movesAccepted += paragraphResult.movesAccepted;
          result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
        } else if ('getRows' in element && typeof element.getRows === 'function') {
          // It's a Table - process its cells
          for (const row of (element as any).getRows()) {
            for (const cell of row.getCells()) {
              for (const paragraph of cell.getParagraphs()) {
                const paragraphResult = acceptRevisionsInParagraph(paragraph, opts);
                result.insertionsAccepted += paragraphResult.insertionsAccepted;
                result.deletionsAccepted += paragraphResult.deletionsAccepted;
                result.movesAccepted += paragraphResult.movesAccepted;
                result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
              }
            }
          }
        }
      }
    }

    // Process paragraphs in footers
    const footers = headerFooterManager.getAllFooters();
    for (const footerEntry of footers) {
      const elements = footerEntry.footer.getElements();
      for (const element of elements) {
        // Element can be Paragraph or Table
        if ('getContent' in element && typeof element.getContent === 'function') {
          // It's a Paragraph
          const paragraphResult = acceptRevisionsInParagraph(element as Paragraph, opts);
          result.insertionsAccepted += paragraphResult.insertionsAccepted;
          result.deletionsAccepted += paragraphResult.deletionsAccepted;
          result.movesAccepted += paragraphResult.movesAccepted;
          result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
        } else if ('getRows' in element && typeof element.getRows === 'function') {
          // It's a Table - process its cells
          for (const row of (element as any).getRows()) {
            for (const cell of row.getCells()) {
              for (const paragraph of cell.getParagraphs()) {
                const paragraphResult = acceptRevisionsInParagraph(paragraph, opts);
                result.insertionsAccepted += paragraphResult.insertionsAccepted;
                result.deletionsAccepted += paragraphResult.deletionsAccepted;
                result.movesAccepted += paragraphResult.movesAccepted;
                result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
              }
            }
          }
        }
      }
    }
  }

  // Clear revision manager
  const revisionManager = doc.getRevisionManager();
  if (revisionManager) {
    revisionManager.clear();
  }

  // Disable track changes setting
  doc.disableTrackChanges();

  result.totalAccepted =
    result.insertionsAccepted +
    result.deletionsAccepted +
    result.movesAccepted +
    result.propertyChangesAccepted;

  logger.info('Revisions accepted in-memory', {
    insertions: result.insertionsAccepted,
    deletions: result.deletionsAccepted,
    moves: result.movesAccepted,
    propertyChanges: result.propertyChangesAccepted,
    total: result.totalAccepted,
  });

  return result;
}

/**
 * Accept revisions in a single paragraph by transforming its content array.
 *
 * The transformation follows these rules:
 * - Insertions (w:ins): Unwrap - extract child content into parent position
 * - Deletions (w:del): Remove - delete revision and its content
 * - MoveFrom (w:moveFrom): Remove - content exists at moveTo destination
 * - MoveTo (w:moveTo): Unwrap - keep content, remove wrapper
 * - Property changes: Remove from model (current formatting is kept)
 *
 * @param paragraph - Paragraph to process
 * @param options - Options for which revision types to accept
 * @returns Result with counts of accepted revisions
 */
function acceptRevisionsInParagraph(
  paragraph: Paragraph,
  options: Required<AcceptRevisionsOptions>
): AcceptRevisionsResult {
  const result: AcceptRevisionsResult = {
    insertionsAccepted: 0,
    deletionsAccepted: 0,
    movesAccepted: 0,
    propertyChangesAccepted: 0,
    totalAccepted: 0,
  };

  const content = paragraph.getContent();
  const newContent: ParagraphContent[] = [];

  for (const item of content) {
    if (item instanceof Revision) {
      const revisionType = item.getType();

      // Handle insertion revisions (w:ins)
      if (revisionType === 'insert' && options.acceptInsertions) {
        // Unwrap: Extract child content into parent position
        const childContent = item.getContent();
        for (const child of childContent) {
          if (isRunContent(child)) {
            newContent.push(child as Run);
          } else if (isHyperlinkContent(child)) {
            newContent.push(child as Hyperlink);
          }
        }
        result.insertionsAccepted++;
        continue;
      }

      // Handle deletion revisions (w:del)
      if (revisionType === 'delete' && options.acceptDeletions) {
        // Remove: Don't add to newContent - content is deleted
        result.deletionsAccepted++;
        continue;
      }

      // Handle moveFrom revisions (source of moved content)
      if (revisionType === 'moveFrom' && options.acceptMoves) {
        // Remove: Content exists at moveTo destination
        result.movesAccepted++;
        continue;
      }

      // Handle moveTo revisions (destination of moved content)
      if (revisionType === 'moveTo' && options.acceptMoves) {
        // Unwrap: Keep content, remove wrapper
        const childContent = item.getContent();
        for (const child of childContent) {
          if (isRunContent(child)) {
            newContent.push(child as Run);
          } else if (isHyperlinkContent(child)) {
            newContent.push(child as Hyperlink);
          }
        }
        result.movesAccepted++;
        continue;
      }

      // Handle property change revisions
      if (PROPERTY_REVISION_TYPES.includes(revisionType) && options.acceptPropertyChanges) {
        // For property changes, the revision is metadata attached to runs
        // The current formatting (newProperties) is already applied to the run
        // We just need to remove the change tracking metadata
        // The content inside should be preserved
        const childContent = item.getContent();
        for (const child of childContent) {
          if (isRunContent(child)) {
            newContent.push(child as Run);
          } else if (isHyperlinkContent(child)) {
            newContent.push(child as Hyperlink);
          }
        }
        result.propertyChangesAccepted++;
        continue;
      }

      // If we reach here, this revision type is not being accepted
      // Keep it in the content
      newContent.push(item);
    } else {
      // Non-revision content - keep as-is
      newContent.push(item);
    }
  }

  // Replace paragraph content with the transformed content
  paragraph.setContent(newContent);

  // Clear paragraph property change tracking (pPrChange) if accepting property changes
  // This removes the w:pPrChange element from the paragraph's formatting
  if (options.acceptPropertyChanges) {
    const formatting = paragraph.getFormatting();
    if (formatting.pPrChange) {
      paragraph.clearParagraphPropertiesChange();
      result.propertyChangesAccepted++;
    }
  }

  return result;
}

/**
 * Check if a paragraph has any revisions
 */
export function paragraphHasRevisions(paragraph: Paragraph): boolean {
  const content = paragraph.getContent();
  return content.some((item) => item instanceof Revision);
}

/**
 * Get all revisions from a paragraph
 */
export function getRevisionsFromParagraph(paragraph: Paragraph): Revision[] {
  const content = paragraph.getContent();
  return content.filter((item): item is Revision => item instanceof Revision);
}

/**
 * Count revisions by type in a document
 */
export function countRevisionsByType(doc: Document): Map<RevisionType, number> {
  const counts = new Map<RevisionType, number>();

  const paragraphs = doc.getAllParagraphs();
  for (const paragraph of paragraphs) {
    const revisions = getRevisionsFromParagraph(paragraph);
    for (const revision of revisions) {
      const type = revision.getType();
      counts.set(type, (counts.get(type) || 0) + 1);
    }
  }

  return counts;
}
