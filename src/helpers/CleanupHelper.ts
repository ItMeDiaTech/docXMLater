/**
 * CleanupHelper - Comprehensive document cleanup utilities
 *
 * Provides methods to clean up common issues in DOCX documents, including:
 * - Unlocking and removing SDTs
 * - Clearing preserve flags
 * - Defragmenting hyperlinks
 * - Cleaning unused elements
 * - Removing customXML
 * - Unlocking fields and frames
 * - Sanitizing tables
 *
 * Usage:
 * const cleanup = new CleanupHelper(doc);
 * cleanup.all(); // Run all cleanups
 */

import type { Document } from '../core/Document.js';
import { Field, ComplexField } from '../elements/Field.js';
import { Hyperlink } from '../elements/Hyperlink.js';
import { Paragraph } from '../elements/Paragraph.js';
import { Table } from '../elements/Table.js';
import { StructuredDocumentTag } from '../elements/StructuredDocumentTag.js';
import { StylesManager } from '../formatting/StylesManager.js';

export interface CleanupOptions {
  /** Unlock all SDTs to enable editing */
  unlockSDTs?: boolean;
  /** Remove all SDTs (unwrap content) */
  removeSDTs?: boolean;
  /** Clear paragraph preserve flags */
  clearPreserveFlags?: boolean;
  /** Merge fragmented hyperlinks */
  defragmentHyperlinks?: boolean;
  /** Reset hyperlink formatting to standard */
  resetHyperlinkFormatting?: boolean;
  /** Remove unused numbering definitions */
  cleanupNumbering?: boolean;
  /** Remove unused styles */
  cleanupStyles?: boolean;
  /** Remove orphaned relationships */
  cleanupRelationships?: boolean;
  /** Remove customXML elements */
  removeCustomXML?: boolean;
  /** Unlock field locks (enable field updates) */
  unlockFields?: boolean;
  /** Remove frame/text box locks */
  unlockFrames?: boolean;
  /** Sanitize table property exceptions (tblPrEx) */
  sanitizeTables?: boolean;
  /** Format internal anchor hyperlinks with standard styling (Verdana 12pt blue underlined) */
  formatInternalHyperlinks?: boolean;
  /** Format ALL hyperlinks (internal, external, and HYPERLINK fields) with standard styling (Verdana 12pt #0000FF underlined) */
  formatAllHyperlinks?: boolean;
}

export interface CleanupReport {
  sdtsUnlocked: number;
  sdtsRemoved: number;
  preserveFlagsCleared: number;
  hyperlinksDefragmented: number;
  numberingRemoved: number;
  stylesRemoved: number;
  relationshipsRemoved: number;
  customXMLRemoved: number;
  fieldsUnlocked: number;
  framesUnlocked: number;
  tablesProcessed: number;
  internalHyperlinksFormatted: number;
  allHyperlinksFormatted: number;
  warnings: string[];
}

export class CleanupHelper {
  private doc: Document;

  constructor(doc: Document) {
    this.doc = doc;
  }

  /**
   * Run all cleanup operations with default settings
   * @returns Cleanup report
   */
  all(): CleanupReport {
    return this.run({
      unlockSDTs: true,
      removeSDTs: true,
      clearPreserveFlags: true,
      defragmentHyperlinks: true,
      resetHyperlinkFormatting: true,
      cleanupNumbering: true,
      cleanupStyles: true,
      cleanupRelationships: true,
      removeCustomXML: true,
      unlockFields: true,
      unlockFrames: true,
      sanitizeTables: true,
      formatAllHyperlinks: true,
    });
  }

  /**
   * Run selective cleanup operations
   * @param options Cleanup options
   * @returns Cleanup report
   */
  run(options: CleanupOptions): CleanupReport {
    const report: CleanupReport = {
      sdtsUnlocked: 0,
      sdtsRemoved: 0,
      preserveFlagsCleared: 0,
      hyperlinksDefragmented: 0,
      numberingRemoved: 0,
      stylesRemoved: 0,
      relationshipsRemoved: 0,
      customXMLRemoved: 0,
      fieldsUnlocked: 0,
      framesUnlocked: 0,
      tablesProcessed: 0,
      internalHyperlinksFormatted: 0,
      allHyperlinksFormatted: 0,
      warnings: [],
    };

    if (options.unlockSDTs) {
      report.sdtsUnlocked = this.unlockSDTs();
    }

    if (options.removeSDTs) {
      report.sdtsRemoved = this.removeSDTs();
    }

    if (options.clearPreserveFlags) {
      report.preserveFlagsCleared = this.clearPreserveFlags();
    }

    if (options.defragmentHyperlinks) {
      report.hyperlinksDefragmented = this.defragmentHyperlinks(
        options.resetHyperlinkFormatting ?? false
      );
    }

    if (options.cleanupNumbering) {
      report.numberingRemoved = this.cleanupNumbering();
    }

    if (options.cleanupStyles) {
      report.stylesRemoved = this.cleanupStyles();
    }

    if (options.cleanupRelationships) {
      report.relationshipsRemoved = this.cleanupRelationships();
    }

    if (options.removeCustomXML) {
      report.customXMLRemoved = this.removeCustomXML();
    }

    if (options.unlockFields) {
      report.fieldsUnlocked = this.unlockFields();
    }

    if (options.unlockFrames) {
      report.framesUnlocked = this.unlockFrames();
    }

    if (options.sanitizeTables) {
      report.tablesProcessed = this.sanitizeTables();
    }

    if (options.formatInternalHyperlinks) {
      report.internalHyperlinksFormatted = this.formatInternalHyperlinks();
    }

    if (options.formatAllHyperlinks) {
      report.allHyperlinksFormatted = this.formatAllHyperlinks();
    }

    return report;
  }

  private unlockSDTs(): number {
    let count = 0;
    const bodyElements = this.doc.getBodyElements();

    for (const element of bodyElements) {
      if (element instanceof StructuredDocumentTag && element.isLocked()) {
        element.unlock();
        count++;
      }
    }

    // Also unlock in tables
    for (const table of this.doc.getAllTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            // SDTs can wrap paragraphs in cells
            const content = para.getContent();
            for (const item of content) {
              if (item instanceof StructuredDocumentTag && item.isLocked()) {
                item.unlock();
                count++;
              }
            }
          }
        }
      }
    }

    return count;
  }

  private removeSDTs(): number {
    // Unwrap SDT wrappers, preserving their content
    const bodyElements = this.doc.getBodyElements();
    type BodyElement = Paragraph | Table | StructuredDocumentTag;
    const unwrapped: BodyElement[] = [];
    let sdtCount = 0;

    const unwrapSDT = (sdt: StructuredDocumentTag, target: BodyElement[]) => {
      sdtCount++;
      for (const item of sdt.getContent()) {
        if (item instanceof Paragraph || item instanceof Table) {
          target.push(item);
        } else if (item instanceof StructuredDocumentTag) {
          unwrapSDT(item, target);
        }
      }
    };

    for (const element of bodyElements) {
      if (element instanceof StructuredDocumentTag) {
        unwrapSDT(element, unwrapped);
      } else {
        unwrapped.push(element as BodyElement);
      }
    }

    this.doc.setBodyElements(unwrapped);
    return sdtCount;
  }

  private clearPreserveFlags(): number {
    let cleared = 0;
    for (const para of this.doc.getAllParagraphs()) {
      if (para.isPreserved()) {
        para.setPreserved(false);
        cleared++;
      }
    }
    return cleared;
  }

  private defragmentHyperlinks(resetFormatting: boolean): number {
    return this.doc.defragmentHyperlinks({ resetFormatting, cleanupRelationships: true });
  }

  private cleanupNumbering(): number {
    const before = this.doc.getNumberingManager().getAllInstances().length;
    this.doc.cleanupUnusedNumbering();
    const after = this.doc.getNumberingManager().getAllInstances().length;
    return before - after;
  }

  private cleanupStyles(): number {
    const stylesManager = this.doc.getStylesManager();
    const usedStyles = new Set<string>();

    const collectParagraph = (para: Paragraph): void => {
      const paraStyle = para.getFormatting().style;
      if (paraStyle) usedStyles.add(paraStyle);
      for (const run of para.getRuns()) {
        const runStyle = run.getFormatting().characterStyle;
        if (runStyle) usedStyles.add(runStyle);
      }
    };

    // Body paragraphs (getAllParagraphs walks table cells and nested SDTs)
    for (const para of this.doc.getAllParagraphs()) {
      collectParagraph(para);
    }

    // Table styles (w:tblStyle) — tables reference styles without any
    // paragraph carrying the ID, so a paragraph-only scan misses them
    for (const table of this.doc.getAllTables()) {
      const tableStyle = table.getFormatting().style;
      if (tableStyle) usedStyles.add(tableStyle);
    }

    // Headers and footers live outside the body walk
    const headerFooterManager = this.doc.getHeaderFooterManager();
    const headerFooterElements = [
      ...headerFooterManager.getAllHeaders().flatMap((entry) => entry.header.getElements()),
      ...headerFooterManager.getAllFooters().flatMap((entry) => entry.footer.getElements()),
    ];
    for (const element of headerFooterElements) {
      if (element instanceof Paragraph) {
        collectParagraph(element);
      } else if (element instanceof Table) {
        const tableStyle = element.getFormatting().style;
        if (tableStyle) usedStyles.add(tableStyle);
        for (const row of element.getRows()) {
          for (const cell of row.getCells()) {
            for (const para of cell.getParagraphs()) {
              collectParagraph(para);
            }
          }
        }
      }
    }

    // Footnotes and endnotes also live outside the body walk
    for (const footnote of this.doc.getFootnoteManager().getAllFootnotes()) {
      for (const para of footnote.getParagraphs()) {
        collectParagraph(para);
      }
    }
    for (const endnote of this.doc.getEndnoteManager().getAllEndnotes()) {
      for (const para of endnote.getParagraphs()) {
        collectParagraph(para);
      }
    }

    // Numbering levels can bind a paragraph style (w:pStyle, ECMA-376 §17.9.23)
    for (const abstractNum of this.doc.getNumberingManager().getAllAbstractNumberings()) {
      for (const level of abstractNum.getAllLevels()) {
        const pStyle = level.getParagraphStyle();
        if (pStyle) usedStyles.add(pStyle);
      }
    }

    // Expand to basedOn/link/next ancestors transitively — removing a chain
    // member leaves dangling references that break style resolution in Word
    const allStyles = stylesManager.getAllStyles();
    const stylesById = new Map(allStyles.map((style) => [style.getStyleId(), style]));
    const pending = [...usedStyles];
    while (pending.length > 0) {
      const style = stylesById.get(pending.pop()!);
      if (!style) continue;
      const props = style.getProperties();
      for (const ref of [props.basedOn, props.link, props.next]) {
        if (ref && !usedStyles.has(ref)) {
          usedStyles.add(ref);
          pending.push(ref);
        }
      }
    }

    // Remove unused styles — but never built-ins or part defaults, which
    // apply without being referenced (mirrors StylesManager.cleanupUnusedStyles)
    let removed = 0;
    for (const style of allStyles) {
      const styleId = style.getStyleId();
      if (usedStyles.has(styleId)) continue;
      if (StylesManager.isBuiltInStyle(styleId)) continue;
      if (style.getProperties().isDefault) continue;
      if (stylesManager.removeStyle(styleId)) {
        removed++;
      }
    }

    return removed;
  }

  private cleanupRelationships(): number {
    // Use comprehensive scanning that includes raw nested content (nested tables),
    // headers/footers, footnotes, and endnotes — not just in-memory hyperlinks
    const referencedIds = this.doc.collectAllReferencedHyperlinkIds();

    // Remove orphaned hyperlink relationships
    return this.doc.getRelationshipManager().removeOrphanedHyperlinks(referencedIds);
  }

  private removeCustomXML(): number {
    const zipHandler = this.doc.getZipHandler();
    let removed = 0;

    // Remove customXML files
    const files = zipHandler.getFilePaths();
    for (const file of files) {
      if (file.startsWith('customXml/') || file.startsWith('customXML/')) {
        zipHandler.removeFile(file);
        removed++;
      }
    }

    // Remove customXML relationships
    const relManager = this.doc.getRelationshipManager();
    const customRels = relManager.getRelationshipsByType(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
    );
    for (const rel of customRels) {
      relManager.removeRelationship(rel.getId());
      removed++;
    }

    // Remove custom.xml if present (docProps/custom.xml)
    if (zipHandler.hasFile('docProps/custom.xml')) {
      zipHandler.removeFile('docProps/custom.xml');
      removed++;
    }

    return removed;
  }

  // Lock state must be cleared on the in-memory model, not the ZIP copy of
  // word/document.xml: prepareSave() regenerates document.xml from the model,
  // so raw-XML edits here would be overwritten and the locks re-emitted.
  private unlockFields(): number {
    let count = 0;

    for (const para of this.doc.getAllParagraphs()) {
      // w:fldChar w:fldLock on complex-field runs (ECMA-376 §17.16.18).
      // getRuns() includes runs inside revisions and hyperlinks; the
      // returned content elements are live references, so clearing the
      // flag mutates the run itself.
      for (const run of para.getRuns()) {
        for (const item of run.getContent()) {
          if (item.type === 'fieldChar' && item.fieldCharLocked === true) {
            item.fieldCharLocked = undefined;
            count++;
          }
        }
      }

      // w:fldLock on simple fields (w:fldSimple, ECMA-376 §17.16.16).
      // Field keeps the flag private with no mutator, so clear it
      // structurally — the regenerated fldSimple then omits the attribute.
      for (const item of para.getContent()) {
        if (item instanceof Field) {
          const lockable = item as unknown as { fldLock?: boolean };
          if (lockable.fldLock === true) {
            lockable.fldLock = undefined;
            count++;
          }
        }
      }
    }

    return count;
  }

  private unlockFrames(): number {
    let count = 0;

    for (const para of this.doc.getAllParagraphs()) {
      const framePr = para.getFormatting().framePr;
      if (framePr?.anchorLock === true) {
        para.setFrameProperties({ ...framePr, anchorLock: undefined });
        count++;
      }
    }

    return count;
  }

  private sanitizeTables(): number {
    const tables = this.doc.getAllTables();
    let processed = 0;
    for (const table of tables) {
      for (const row of table.getRows()) {
        const exceptions = row.getTablePropertyExceptions();
        if (exceptions && Object.keys(exceptions).length > 0) {
          row.setTablePropertyExceptions(undefined as any);
        }
      }
      processed++;
    }
    return processed;
  }

  private formatInternalHyperlinks(): number {
    let count = 0;
    const formatting = {
      font: 'Verdana',
      size: 12,
      color: '0000FF',
      underline: 'single' as const,
    };

    // Process body paragraphs
    for (const paragraph of this.doc.getAllParagraphs()) {
      for (const item of paragraph.getContent()) {
        if (item instanceof Hyperlink && item.isInternal()) {
          item.setFormatting(formatting, { replace: true });
          count++;
        }
      }
    }

    // Process table paragraphs
    for (const table of this.doc.getAllTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            for (const item of para.getContent()) {
              if (item instanceof Hyperlink && item.isInternal()) {
                item.setFormatting(formatting, { replace: true });
                count++;
              }
            }
          }
        }
      }
    }

    return count;
  }

  /**
   * Formats ALL hyperlinks in the document with standard styling
   * This includes:
   * - Internal w:hyperlink elements (bookmarks)
   * - External w:hyperlink elements (URLs)
   * - HYPERLINK fields (both simple w:fldSimple and complex fields)
   *
   * Standard formatting: Verdana 12pt, #0000FF blue, single underline
   * @returns Number of hyperlinks formatted
   */
  private formatAllHyperlinks(): number {
    let count = 0;
    const formatting = {
      font: 'Verdana',
      size: 12,
      color: '0000FF',
      underline: 'single' as const,
    };

    // Helper to process paragraph content
    const processParagraph = (paragraph: Paragraph): void => {
      for (const item of paragraph.getContent()) {
        // Process all Hyperlink instances (both internal AND external)
        if (item instanceof Hyperlink) {
          item.setFormatting(formatting, { replace: true });
          count++;
        }
        // Process simple HYPERLINK fields
        if (item instanceof Field && item.isHyperlinkField()) {
          item.setFormatting(formatting);
          count++;
        }
        // Process complex HYPERLINK fields
        if (item instanceof ComplexField && item.isHyperlinkField()) {
          item.setResultFormatting(formatting);
          count++;
        }
      }
    };

    // Process body paragraphs
    for (const paragraph of this.doc.getAllParagraphs()) {
      processParagraph(paragraph);
    }

    // Process table paragraphs
    for (const table of this.doc.getAllTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            processParagraph(para);
          }
        }
      }
    }

    return count;
  }

  /**
   * Preset: Google Docs cleanup
   */
  static googleDocsPreset(): CleanupOptions {
    return {
      unlockSDTs: true,
      removeSDTs: true,
      defragmentHyperlinks: true,
      resetHyperlinkFormatting: true,
      cleanupRelationships: true,
      removeCustomXML: true,
      sanitizeTables: true,
    };
  }

  /**
   * Preset: Full cleanup
   */
  static fullCleanupPreset(): CleanupOptions {
    return {
      unlockSDTs: true,
      removeSDTs: true,
      clearPreserveFlags: true,
      defragmentHyperlinks: true,
      resetHyperlinkFormatting: true,
      cleanupNumbering: true,
      cleanupStyles: true,
      cleanupRelationships: true,
      removeCustomXML: true,
      unlockFields: true,
      unlockFrames: true,
      sanitizeTables: true,
      formatAllHyperlinks: true,
    };
  }

  /**
   * Preset: Minimal cleanup
   */
  static minimalPreset(): CleanupOptions {
    return {
      cleanupRelationships: true,
      removeCustomXML: true,
    };
  }
}
