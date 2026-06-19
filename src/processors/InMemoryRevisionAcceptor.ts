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

import type { Document } from '../core/Document.js';
import { Paragraph } from '../elements/Paragraph.js';
import type { ParagraphContent } from '../elements/Paragraph.js';
import { Revision, RevisionType } from '../elements/Revision.js';
import {
  isRunContent,
  isHyperlinkContent,
  isImageRunContent,
} from '../elements/RevisionContent.js';
import { ComplexField } from '../elements/Field.js';
import { RangeMarker, RangeMarkerType } from '../elements/RangeMarker.js';
import { PreservedElement } from '../elements/PreservedElement.js';
import { Table } from '../elements/Table.js';
import { getGlobalLogger, createScopedLogger, ILogger } from '../utils/logger.js';

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
  /** Remove empty tables after revision acceptance - default: true */
  cleanupEmptyTables?: boolean;
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
  /** Number of empty tables removed during cleanup */
  emptyTablesRemoved: number;
}

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
 * Range-marker boundary types dropped when the matching accept option is set.
 * These are paired boundaries for the move / ins / del revisions handled
 * elsewhere in the loop; once the revision they delimit is accepted, the
 * markers reference a change that no longer exists (ECMA-376 §17.13.5.21-28)
 * and Word still reports the document as containing tracked changes. The
 * raw-XML acceptor (acceptRevisions.ts) and stripRevisionsFromXml remove the
 * same set; this keeps the in-memory paragraph path in parity.
 */
const MOVE_RANGE_MARKER_TYPES = new Set<RangeMarkerType>([
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'customXmlMoveFromRangeStart',
  'customXmlMoveFromRangeEnd',
  'customXmlMoveToRangeStart',
  'customXmlMoveToRangeEnd',
]);
const INS_RANGE_MARKER_TYPES = new Set<RangeMarkerType>([
  'customXmlInsRangeStart',
  'customXmlInsRangeEnd',
]);
const DEL_RANGE_MARKER_TYPES = new Set<RangeMarkerType>([
  'customXmlDelRangeStart',
  'customXmlDelRangeEnd',
]);

/**
 * Strip revision markup from raw XML string.
 * Used for nested tables stored as raw XML that cannot be processed via the in-memory model.
 *
 * Follows the same rules as the main revision acceptor:
 * - Insertions: Keep content, remove wrapper tags
 * - Deletions: Remove entirely (content and tags)
 * - MoveFrom: Remove entirely (source of move)
 * - MoveTo: Keep content, remove wrapper
 * - Property changes: Remove change tracking elements
 * - Range markers: Remove boundary markers
 *
 * @param xml - Raw XML string containing revision markup
 * @returns Cleaned XML with revisions accepted
 */
export function stripRevisionsFromXml(xml: string): string {
  let result = xml;

  // Step 1: Remove range markers (must be done first)
  const rangePatterns = [
    /<w:moveFromRangeStart[^>]*(?:\/>|>.*?<\/w:moveFromRangeStart>)/gs,
    /<w:moveFromRangeEnd[^>]*(?:\/>|>.*?<\/w:moveFromRangeEnd>)/gs,
    /<w:moveToRangeStart[^>]*(?:\/>|>.*?<\/w:moveToRangeStart>)/gs,
    /<w:moveToRangeEnd[^>]*(?:\/>|>.*?<\/w:moveToRangeEnd>)/gs,
    /<w:customXmlInsRangeStart[^>]*(?:\/>|>.*?<\/w:customXmlInsRangeStart>)/gs,
    /<w:customXmlInsRangeEnd[^>]*(?:\/>|>.*?<\/w:customXmlInsRangeEnd>)/gs,
    /<w:customXmlDelRangeStart[^>]*(?:\/>|>.*?<\/w:customXmlDelRangeStart>)/gs,
    /<w:customXmlDelRangeEnd[^>]*(?:\/>|>.*?<\/w:customXmlDelRangeEnd>)/gs,
  ];
  for (const pattern of rangePatterns) {
    result = result.replace(pattern, '');
  }

  // Step 2: Remove property change elements
  const propChangePatterns = [
    /<w:rPrChange[^>]*>[\s\S]*?<\/w:rPrChange>/g,
    /<w:pPrChange[^>]*>[\s\S]*?<\/w:pPrChange>/g,
    /<w:tblPrChange[^>]*>[\s\S]*?<\/w:tblPrChange>/g,
    /<w:tblPrExChange[^>]*>[\s\S]*?<\/w:tblPrExChange>/g,
    /<w:tcPrChange[^>]*>[\s\S]*?<\/w:tcPrChange>/g,
    /<w:trPrChange[^>]*>[\s\S]*?<\/w:trPrChange>/g,
    /<w:sectPrChange[^>]*>[\s\S]*?<\/w:sectPrChange>/g,
    /<w:tblGridChange[^>]*>[\s\S]*?<\/w:tblGridChange>/g,
    /<w:numberingChange[^>]*>[\s\S]*?<\/w:numberingChange>/g,
  ];
  for (const pattern of propChangePatterns) {
    result = result.replace(pattern, '');
  }

  // Step 3: Remove deletions entirely (including content)
  // Self-closing markers (paragraph-mark deletions in w:pPr/w:rPr per
  // ECMA-376 §17.13.5.15) must be stripped BEFORE the block pattern runs:
  // its [^>]* also matches the trailing '/' of `<w:del .../>`, which would
  // turn the marker into an opening tag and swallow everything up to the
  // next </w:del>, leaving unbalanced XML.
  result = result.replace(/<w:del\b[^>]*\/>/g, '');
  // Iterate until no more deletions (handles nested cases)
  let prevLen = 0;
  while (result.length !== prevLen) {
    prevLen = result.length;
    result = result.replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, '');
  }

  // Step 4: Remove moveFrom entirely (source of moved content)
  // Same self-closing-first ordering as Step 3.
  result = result.replace(/<w:moveFrom\b[^>]*\/>/g, '');
  prevLen = 0;
  while (result.length !== prevLen) {
    prevLen = result.length;
    result = result.replace(/<w:moveFrom\b[^>]*>[\s\S]*?<\/w:moveFrom>/g, '');
  }

  // Step 5: Unwrap moveTo (keep content, remove wrapper)
  result = result.replace(/<\/w:moveTo>/g, '');
  result = result.replace(/<w:moveTo\b[^>]*>/g, '');

  // Step 6: Unwrap insertions (keep content, remove wrapper)
  result = result.replace(/<\/w:ins>/g, '');
  result = result.replace(/<w:ins\b[^>]*>/g, '');

  // Step 7: Clean up orphaned tags
  result = result.replace(/<w:ins\b[^>]*\/>/g, '');
  result = result.replace(/<w:del\b[^>]*\/>/g, '');
  result = result.replace(/<w:moveFrom\b[^>]*\/>/g, '');
  result = result.replace(/<w:moveTo\b[^>]*\/>/g, '');

  return result;
}

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
    cleanupEmptyTables: options.cleanupEmptyTables ?? true,
  };

  const result: AcceptRevisionsResult = {
    insertionsAccepted: 0,
    deletionsAccepted: 0,
    movesAccepted: 0,
    propertyChangesAccepted: 0,
    totalAccepted: 0,
    emptyTablesRemoved: 0,
  };

  logger.info('Accepting revisions in-memory', { options: opts });

  // Paragraphs whose tracked paragraph-mark deletion (w:del in w:pPr/w:rPr)
  // is accepted — they must be merged into their following sibling after the
  // walk completes (ECMA-376 §17.13.5.15).
  const acceptedMarkDeletions: Paragraph[] = [];

  // Validate move pairs before accepting if moves are being accepted
  // Orphaned moves could result in content loss (moveFrom without moveTo = content deleted)
  if (opts.acceptMoves) {
    const revisionManager = doc.getRevisionManager();
    if (revisionManager) {
      const movePairValidation = revisionManager.validateMovePairs();
      if (!movePairValidation.valid) {
        if (movePairValidation.orphanedMoveFrom.length > 0) {
          logger.warn(
            'Orphaned moveFrom revisions detected - accepting these will DELETE content permanently ' +
              '(content was moved from here but no moveTo destination exists)',
            { orphanedMoveIds: movePairValidation.orphanedMoveFrom }
          );
        }
        if (movePairValidation.orphanedMoveTo.length > 0) {
          logger.warn(
            'Orphaned moveTo revisions detected - content may be duplicated ' +
              '(content moved to here but no moveFrom source exists)',
            { orphanedMoveIds: movePairValidation.orphanedMoveTo }
          );
        }
      }
    }
  }

  // Process all paragraphs in the document body
  const paragraphs = doc.getAllParagraphs();
  for (const paragraph of paragraphs) {
    const paragraphResult = acceptRevisionsInParagraph(paragraph, opts, acceptedMarkDeletions);
    result.insertionsAccepted += paragraphResult.insertionsAccepted;
    result.deletionsAccepted += paragraphResult.deletionsAccepted;
    result.movesAccepted += paragraphResult.movesAccepted;
    result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
  }

  // Process paragraphs in tables and clear table/row/cell property changes
  const tables = doc.getTables();
  for (const table of tables) {
    // Clear tblPrChange on table
    if (opts.acceptPropertyChanges && table.getTblPrChange()) {
      table.clearTblPrChange();
      result.propertyChangesAccepted++;
    }

    // Row-level tracked ins / del (CT_TrPr: w:ins / w:del per ECMA-376
    // §17.13.5.19 / §17.13.5.14) — these mark the ENTIRE row as a tracked
    // insertion or deletion. Previously the acceptor only cleared the
    // tcPrChange / trPrChange / cellRevision metadata and silently ignored
    // row-level markers, so deleted rows persisted and inserted rows kept
    // their pending state. Collect deletion indices up front so we can
    // splice them out after iteration (reverse order to keep indices stable).
    const rowsToRemove: number[] = [];
    const allRows = table.getRows();
    for (let rowIdx = 0; rowIdx < allRows.length; rowIdx++) {
      const row = allRows[rowIdx]!;
      const fmt = row.getFormatting();
      if (fmt.rowInsertion && opts.acceptInsertions) {
        row.setRowInsertion(undefined);
        result.insertionsAccepted++;
      }
      if (fmt.rowDeletion && opts.acceptDeletions) {
        rowsToRemove.push(rowIdx);
        result.deletionsAccepted++;
      }
    }
    for (let i = rowsToRemove.length - 1; i >= 0; i--) {
      table._removeRowAtIndex(rowsToRemove[i]!);
    }

    for (const row of table.getRows()) {
      // Clear trPrChange on row
      if (opts.acceptPropertyChanges && row.getTrPrChange()) {
        row.clearTrPrChange();
        result.propertyChangesAccepted++;
      }

      for (const cell of row.getCells()) {
        // Clear tcPrChange on cell
        if (opts.acceptPropertyChanges && cell.getTcPrChange()) {
          cell.clearTcPrChange();
          result.propertyChangesAccepted++;
        }

        // Route cellIns / cellDel / cellMerge by their semantic revision type
        // (per ECMA-376 §17.13.5.4-6). Previously all three were lumped under
        // `acceptPropertyChanges`, which meant `{ acceptInsertions: true }`
        // alone never cleared a cellIns marker, and `{ acceptPropertyChanges:
        // true }` alone cleared everything including insertions. Correct
        // mapping:
        //   - tableCellInsert (w:cellIns)   → acceptInsertions
        //   - tableCellDelete (w:cellDel)   → acceptDeletions
        //   - tableCellMerge  (w:cellMerge) → acceptPropertyChanges
        const cellRev = cell.getCellRevision();
        if (cellRev) {
          const revType = cellRev.getType();
          if (revType === 'tableCellInsert' && opts.acceptInsertions) {
            cell.clearCellRevision();
            result.insertionsAccepted++;
          } else if (revType === 'tableCellDelete' && opts.acceptDeletions) {
            cell.clearCellRevision();
            result.deletionsAccepted++;
          } else if (revType === 'tableCellMerge' && opts.acceptPropertyChanges) {
            cell.clearCellRevision();
            result.propertyChangesAccepted++;
          }
        }

        // Process paragraphs in the cell
        for (const paragraph of cell.getParagraphs()) {
          const paragraphResult = acceptRevisionsInParagraph(
            paragraph,
            opts,
            acceptedMarkDeletions
          );
          result.insertionsAccepted += paragraphResult.insertionsAccepted;
          result.deletionsAccepted += paragraphResult.deletionsAccepted;
          result.movesAccepted += paragraphResult.movesAccepted;
          result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
        }

        // Process raw nested content (nested tables stored as XML)
        // These cannot be processed via the in-memory model, so we use XML-based stripping
        if (cell.hasRawNestedContent()) {
          const rawContent = cell.getRawNestedContent();
          for (let i = 0; i < rawContent.length; i++) {
            const item = rawContent[i];
            if (item) {
              const cleanedXml = stripRevisionsFromXml(item.xml);
              if (cleanedXml !== item.xml) {
                cell.updateRawNestedContent(i, cleanedXml);
                // Count revisions stripped from nested content
                // We can't distinguish types in raw XML, so count as property changes
                result.propertyChangesAccepted++;
                logger.debug('Stripped revisions from nested content', {
                  type: item.type,
                  position: item.position,
                  originalLength: item.xml.length,
                  cleanedLength: cleanedXml.length,
                });
              }
            }
          }
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
        // Element can be Paragraph or Table - use instanceof for type safety
        if (element instanceof Paragraph) {
          const paragraphResult = acceptRevisionsInParagraph(element, opts, acceptedMarkDeletions);
          result.insertionsAccepted += paragraphResult.insertionsAccepted;
          result.deletionsAccepted += paragraphResult.deletionsAccepted;
          result.movesAccepted += paragraphResult.movesAccepted;
          result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
        } else if (element instanceof Table) {
          // It's a Table - process its cells
          for (const row of element.getRows()) {
            for (const cell of row.getCells()) {
              for (const paragraph of cell.getParagraphs()) {
                const paragraphResult = acceptRevisionsInParagraph(
                  paragraph,
                  opts,
                  acceptedMarkDeletions
                );
                result.insertionsAccepted += paragraphResult.insertionsAccepted;
                result.deletionsAccepted += paragraphResult.deletionsAccepted;
                result.movesAccepted += paragraphResult.movesAccepted;
                result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
              }
              // Process raw nested content in header tables
              if (cell.hasRawNestedContent()) {
                const rawContent = cell.getRawNestedContent();
                for (let i = 0; i < rawContent.length; i++) {
                  const item = rawContent[i];
                  if (item) {
                    const cleanedXml = stripRevisionsFromXml(item.xml);
                    if (cleanedXml !== item.xml) {
                      cell.updateRawNestedContent(i, cleanedXml);
                      result.propertyChangesAccepted++;
                    }
                  }
                }
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
        // Element can be Paragraph or Table - use instanceof for type safety
        if (element instanceof Paragraph) {
          const paragraphResult = acceptRevisionsInParagraph(element, opts, acceptedMarkDeletions);
          result.insertionsAccepted += paragraphResult.insertionsAccepted;
          result.deletionsAccepted += paragraphResult.deletionsAccepted;
          result.movesAccepted += paragraphResult.movesAccepted;
          result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
        } else if (element instanceof Table) {
          // It's a Table - process its cells
          for (const row of element.getRows()) {
            for (const cell of row.getCells()) {
              for (const paragraph of cell.getParagraphs()) {
                const paragraphResult = acceptRevisionsInParagraph(
                  paragraph,
                  opts,
                  acceptedMarkDeletions
                );
                result.insertionsAccepted += paragraphResult.insertionsAccepted;
                result.deletionsAccepted += paragraphResult.deletionsAccepted;
                result.movesAccepted += paragraphResult.movesAccepted;
                result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
              }
              // Process raw nested content in footer tables
              if (cell.hasRawNestedContent()) {
                const rawContent = cell.getRawNestedContent();
                for (let i = 0; i < rawContent.length; i++) {
                  const item = rawContent[i];
                  if (item) {
                    const cleanedXml = stripRevisionsFromXml(item.xml);
                    if (cleanedXml !== item.xml) {
                      cell.updateRawNestedContent(i, cleanedXml);
                      result.propertyChangesAccepted++;
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  // Walk footnote and endnote paragraphs. Per ECMA-376 §17.11.4 /
  // §17.11.15, notes can hold any block-level content including
  // tracked changes; programmatically-added Revision objects on note
  // paragraphs (e.g.
  // `footnote.getParagraphs()[0].addContent(new Revision(...))`)
  // were never visited by the in-memory acceptor and stayed in the
  // model after `acceptAllRevisions()`. The raw-XML acceptor
  // (iter 135) already handles existing-document note revisions on
  // load, so this loop's job is the programmatic-API path.
  const footnoteManager = doc.getFootnoteManager?.();
  if (footnoteManager) {
    for (const fn of footnoteManager.getAllFootnotes()) {
      for (const paragraph of fn.getParagraphs()) {
        const paragraphResult = acceptRevisionsInParagraph(paragraph, opts, acceptedMarkDeletions);
        result.insertionsAccepted += paragraphResult.insertionsAccepted;
        result.deletionsAccepted += paragraphResult.deletionsAccepted;
        result.movesAccepted += paragraphResult.movesAccepted;
        result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
      }
    }
  }
  const endnoteManager = doc.getEndnoteManager?.();
  if (endnoteManager) {
    for (const en of endnoteManager.getAllEndnotes()) {
      for (const paragraph of en.getParagraphs()) {
        const paragraphResult = acceptRevisionsInParagraph(paragraph, opts, acceptedMarkDeletions);
        result.insertionsAccepted += paragraphResult.insertionsAccepted;
        result.deletionsAccepted += paragraphResult.deletionsAccepted;
        result.movesAccepted += paragraphResult.movesAccepted;
        result.propertyChangesAccepted += paragraphResult.propertyChangesAccepted;
      }
    }
  }

  // Clear sectPrChange on document section
  if (opts.acceptPropertyChanges) {
    const section = doc.getSection();
    if (section?.getSectPrChange()) {
      section.clearSectPrChange();
      result.propertyChangesAccepted++;
    }
  }

  // Clear revision manager
  const revisionManager = doc.getRevisionManager();
  if (revisionManager) {
    revisionManager.clear();
  }

  // Disable track changes setting
  doc.disableTrackChanges();

  // Combine paragraphs whose paragraph-mark deletion was accepted with their
  // following sibling (ECMA-376 §17.13.5.15). Runs after tracking is disabled
  // so the removals are plain splices rather than new tracked deletions.
  if (acceptedMarkDeletions.length > 0) {
    mergeParagraphsWithDeletedMarks(doc, new Set(acceptedMarkDeletions));
  }

  // Cleanup empty tables if enabled
  // This removes tables that have no visible content after revision acceptance
  // (e.g., tables where all content was deleted via tracked changes)
  if (opts.cleanupEmptyTables) {
    result.emptyTablesRemoved = cleanupEmptyTables(doc, logger);
  }

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
    emptyTablesRemoved: result.emptyTablesRemoved,
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
  options: Required<AcceptRevisionsOptions>,
  acceptedMarkDeletions?: Paragraph[]
): AcceptRevisionsResult {
  const result: AcceptRevisionsResult = {
    insertionsAccepted: 0,
    deletionsAccepted: 0,
    movesAccepted: 0,
    propertyChangesAccepted: 0,
    totalAccepted: 0,
    emptyTablesRemoved: 0,
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
          // Check ImageRun FIRST since ImageRun extends Run
          if (isImageRunContent(child)) {
            newContent.push(child);
          } else if (isRunContent(child)) {
            newContent.push(child);
          } else if (isHyperlinkContent(child)) {
            newContent.push(child);
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
          // Check ImageRun FIRST since ImageRun extends Run
          if (isImageRunContent(child)) {
            newContent.push(child);
          } else if (isRunContent(child)) {
            newContent.push(child);
          } else if (isHyperlinkContent(child)) {
            newContent.push(child);
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
          // Check ImageRun FIRST since ImageRun extends Run
          if (isImageRunContent(child)) {
            newContent.push(child);
          } else if (isRunContent(child)) {
            newContent.push(child);
          } else if (isHyperlinkContent(child)) {
            newContent.push(child);
          }
        }
        result.propertyChangesAccepted++;
        continue;
      }

      // If we reach here, this revision type is not being accepted
      // Keep it in the content
      newContent.push(item);
    } else if (item instanceof ComplexField && item.hasResultRevisions()) {
      // Accept revisions nested inside ComplexField result sections
      const fieldRevisions = item.getResultRevisions();
      const revisionsToKeep: Revision[] = [];

      for (const rev of fieldRevisions) {
        const revType = rev.getType();
        if (revType === 'insert' && options.acceptInsertions) {
          result.insertionsAccepted++;
        } else if (revType === 'delete' && options.acceptDeletions) {
          result.deletionsAccepted++;
        } else if (revType === 'moveTo' && options.acceptMoves) {
          result.movesAccepted++;
        } else if (revType === 'moveFrom' && options.acceptMoves) {
          result.movesAccepted++;
        } else {
          revisionsToKeep.push(rev);
        }
      }

      if (revisionsToKeep.length < fieldRevisions.length) {
        // Some revisions were accepted — use getAcceptedResultText() for correct interleaved ordering
        item.setResult(item.getAcceptedResultText());
        if (revisionsToKeep.length > 0) {
          item.setResultRevisions(revisionsToKeep);
        }
      }
      newContent.push(item);
    } else if (item instanceof RangeMarker) {
      // Boundary markers for a move / ins / del span. Drop them when the
      // revision they delimit is being accepted so no orphaned range markers
      // remain (the move/ins/del revisions themselves are counted above; the
      // markers carry no content, so they are not counted again).
      const markerType = item.getType();
      if (
        (options.acceptMoves && MOVE_RANGE_MARKER_TYPES.has(markerType)) ||
        (options.acceptInsertions && INS_RANGE_MARKER_TYPES.has(markerType)) ||
        (options.acceptDeletions && DEL_RANGE_MARKER_TYPES.has(markerType))
      ) {
        continue;
      }
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
    // Paragraph-mark rPrChange (CT_ParaRPrChange, §17.3.1.30) — same
    // rationale as pPrChange: acceptance clears the "previous" snapshot,
    // current formatting (paragraphMarkRunProperties) stays intact.
    if (formatting.paragraphMarkRunPropertiesChange) {
      paragraph.formatting.paragraphMarkRunPropertiesChange = undefined;
      result.propertyChangesAccepted++;
    }
  }

  // Clear paragraph mark deletion tracking if accepting deletions
  // This removes the w:del element from w:pPr/w:rPr. The paragraph is also
  // recorded for the post-walk merge pass: per ECMA-376 §17.13.5.15 a deleted
  // paragraph mark means this paragraph's contents combine with the FOLLOWING
  // paragraph, so clearing the marker alone would leave a leftover paragraph
  // Word's own Accept All removes.
  if (options.acceptDeletions) {
    const formatting = paragraph.getFormatting();
    if (formatting.paragraphMarkDeletion) {
      paragraph.clearParagraphMarkDeletion();
      result.deletionsAccepted++;
      acceptedMarkDeletions?.push(paragraph);
    }
  }

  // Clear paragraph mark insertion tracking if accepting insertions
  // This removes the w:ins element from w:pPr/w:rPr
  if (options.acceptInsertions) {
    const formatting = paragraph.getFormatting();
    if (formatting.paragraphMarkInsertion) {
      paragraph.clearParagraphMarkInsertion();
      result.insertionsAccepted++;
    }
  }

  return result;
}

/**
 * Merge paragraphs whose tracked paragraph-mark deletion was accepted into
 * their immediately following sibling paragraph.
 *
 * Per ECMA-376 §17.13.5.15, a w:del inside w:pPr/w:rPr marks the paragraph
 * mark (¶) itself as deleted: accepting it means the paragraph's contents are
 * no longer delimited by that mark and combine with the FOLLOWING paragraph.
 * Clearing the marker alone leaves a leftover paragraph (blank line) that
 * Word's own Accept All removes.
 *
 * Merging only happens between siblings of the same container (document body
 * or table cell). A paragraph with no following sibling paragraph — last in
 * its container, or followed by a table/SDT/raw nested content — keeps its
 * content in place, matching Word, which cannot delete a container's final
 * paragraph mark.
 */
export function mergeParagraphsWithDeletedMarks(doc: Document, marked: Set<Paragraph>): void {
  if (marked.size === 0) {
    return;
  }

  // Body-level siblings
  mergeMarkedSequence(doc.getBodyElements(), marked, (para) => {
    doc.removeElement(para);
  });

  // Table-cell siblings (same tables the acceptance walk visits)
  for (const table of doc.getTables()) {
    for (const row of table.getRows()) {
      for (const cell of row.getCells()) {
        // Raw nested content (nested tables/SDTs) at position p renders
        // between paragraph p-1 and paragraph p, so it interrupts sibling
        // adjacency the same way a body-level table does. Capture the
        // boundaries before any removal: removeParagraph keeps positions in
        // sync with live indices, which match this snapshot only up front.
        const blockedBoundaries = new Set(cell.getRawNestedContent().map((item) => item.position));
        mergeMarkedSequence(
          cell.getParagraphs(),
          marked,
          (para) => {
            const index = cell.getParagraphs().indexOf(para);
            if (index !== -1) {
              cell.removeParagraph(index);
            }
          },
          (index) => blockedBoundaries.has(index)
        );
      }
    }
  }
}

/**
 * Walk one sibling sequence (body elements or cell paragraphs) and fold each
 * run of marked paragraphs into the first unmarked paragraph that follows it.
 * Consecutive marked paragraphs chain: their contents concatenate, in
 * document order, ahead of the surviving paragraph's own content. The
 * surviving paragraph keeps its own pPr — the deleted marks took their
 * properties with them.
 */
function mergeMarkedSequence(
  sequence: readonly unknown[],
  marked: Set<Paragraph>,
  removeParagraph: (paragraph: Paragraph) => void,
  blockedBefore?: (index: number) => boolean
): void {
  let chain: Paragraph[] = [];

  for (let i = 0; i < sequence.length; i++) {
    const element = sequence[i];

    if (!(element instanceof Paragraph)) {
      // A non-paragraph block (table, SDT, TOC) is not a valid merge target
      chain = [];
      continue;
    }

    if (blockedBefore?.(i)) {
      chain = [];
    }

    if (marked.has(element)) {
      chain.push(element);
      continue;
    }

    if (chain.length > 0) {
      const inherited: ParagraphContent[] = [];
      for (const para of chain) {
        inherited.push(...para.getContent());
      }
      element.setContent([...inherited, ...element.getContent()]);
      for (const para of chain) {
        // Empty before removal so a tracking-bound removal path can never
        // duplicate content that now lives in the surviving paragraph
        para.setContent([]);
        removeParagraph(para);
      }
      chain = [];
    }
  }
  // A trailing chain has no following sibling: the final paragraph mark of a
  // container cannot be deleted, so those paragraphs stay as-is.
}

/**
 * Check if a paragraph has any revisions
 */
export function paragraphHasRevisions(paragraph: Paragraph): boolean {
  const content = paragraph.getContent();
  if (content.some((item) => item instanceof Revision)) {
    return true;
  }
  // Also check paragraph mark revision markers in w:pPr/w:rPr
  const formatting = paragraph.getFormatting();
  if (formatting.paragraphMarkDeletion || formatting.paragraphMarkInsertion) {
    return true;
  }
  return false;
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

/**
 * Inline preserved-element types that are pure range/proofing markers.
 * They carry no displayable content, so a table whose cells hold only these
 * (e.g. leftover proofErr markers after all text was deleted via tracked
 * changes) must still be eligible for cleanup.
 */
const MARKER_PRESERVED_ELEMENT_TYPES = new Set([
  'w:proofErr',
  'w:permStart',
  'w:permEnd',
  'w:commentRangeStart',
  'w:commentRangeEnd',
]);

/**
 * Check whether a paragraph carries visible content for empty-table cleanup.
 *
 * Paragraph.getText() only surfaces Run/Hyperlink text, so images, shapes,
 * text boxes, fields, pending revisions, and preserved raw XML (loaded-doc
 * hyperlinks, inline SDTs, math) would read as "empty" and the table holding
 * them would be silently deleted. Mirrors TableCell.isParaBlank's duck-typed
 * element detection.
 */
function paragraphHasVisibleContent(para: Paragraph): boolean {
  if (para.getText().trim().length > 0) {
    return true;
  }

  for (const item of para.getContent()) {
    if (!item) continue;

    // Revisions left behind by selective acceptance still render their
    // content (e.g. strikethrough deletions) — removing the table would
    // discard tracked changes the caller chose to keep.
    if (item instanceof Revision) {
      return true;
    }

    // Marker-only preserved elements are not content; everything else
    // preserved as raw XML (w:r, w:hyperlink, m:oMath, w:ruby, nested
    // revision wrappers) is.
    if (item instanceof PreservedElement) {
      if (!MARKER_PRESERVED_ELEMENT_TYPES.has(item.getElementType())) {
        return true;
      }
      continue;
    }

    // Duck-typed checks matching TableCell.isParaBlank: ImageRun
    // (getImageElement), Shape (getShapeType), TextBox (getTextContent),
    // Hyperlink (getUrl), Field/ComplexField (getInstruction).
    const candidate = item as unknown as Record<string, unknown>;
    if (
      typeof candidate.getImageElement === 'function' ||
      typeof candidate.getShapeType === 'function' ||
      typeof candidate.getTextContent === 'function' ||
      typeof candidate.getUrl === 'function' ||
      typeof candidate.getInstruction === 'function'
    ) {
      return true;
    }
  }

  return false;
}

/**
 * Remove tables that have no visible content after revision acceptance.
 *
 * A table is considered empty only if ALL cells in ALL rows have no visible
 * content: no text, images, shapes, text boxes, fields, hyperlinks, preserved
 * raw XML, or raw nested content (nested tables/SDTs). This handles cases
 * where all table content was deleted via tracked changes - the deletion
 * markers are stripped but the empty table structure remains.
 *
 * @param doc - Document to clean up
 * @param logger - Logger instance for debug output
 * @returns Number of empty tables removed
 */
function cleanupEmptyTables(doc: Document, logger: ILogger): number {
  const tables = doc.getTables();
  let removedCount = 0;
  const tablesToRemove: number[] = [];

  for (let tableIndex = 0; tableIndex < tables.length; tableIndex++) {
    const table = tables[tableIndex];
    if (!table) continue;

    let hasContent = false;
    const rows = table.getRows();

    for (const row of rows) {
      const cells = row.getCells();
      for (const cell of cells) {
        // Nested tables/SDTs live in raw XML passthrough, never in
        // paragraphs — a cell whose real content is a nested table would
        // otherwise read as empty and be deleted.
        if (cell.hasRawNestedContent()) {
          hasContent = true;
          break;
        }
        const paragraphs = cell.getParagraphs();
        for (const para of paragraphs) {
          if (paragraphHasVisibleContent(para)) {
            hasContent = true;
            break;
          }
        }
        if (hasContent) break;
      }
      if (hasContent) break;
    }

    if (!hasContent) {
      // Mark this table for removal (store index)
      tablesToRemove.push(tableIndex);
      logger.debug('Found empty table for removal', { tableIndex });
    }
  }

  // Remove tables in reverse order to preserve indices
  for (let i = tablesToRemove.length - 1; i >= 0; i--) {
    const tableIndex = tablesToRemove[i];
    if (tableIndex !== undefined && doc.removeTable(tableIndex)) {
      removedCount++;
      logger.debug('Removed empty table', { tableIndex });
    }
  }

  if (removedCount > 0) {
    logger.info('Empty table cleanup complete', { tablesRemoved: removedCount });
  }

  return removedCount;
}
