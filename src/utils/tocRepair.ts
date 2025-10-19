/**
 * TOC Repair Utility - Fix orphaned TOC links and generate proper Table of Contents
 *
 * Uses Word's native TOC field that references Header 2 styles directly.
 * Word automatically generates bookmarks and hyperlinks when the field is updated.
 */

import type { Document } from '../core/Document';
import { Paragraph } from '../elements/Paragraph';
import { Field } from '../elements/Field';
import { Hyperlink } from '../elements/Hyperlink';

/**
 * Result of TOC repair operation
 */
export interface TOCRepairResult {
  /** Whether TOC was successfully generated */
  tocGenerated: boolean;
  /** Number of Header 2 headings found */
  header2Count: number;
  /** Number of "Top of the Document" links added */
  topLinksAdded: number;
}

/**
 * Information about a Header 2 paragraph
 */
interface Header2Info {
  /** The paragraph containing Header 2 */
  paragraph: Paragraph;
  /** Global paragraph index in document */
  paragraphIndex: number;
  /** Text content of the header */
  text: string;
}

/**
 * Repairs TOC by generating Header 2 only TOC and adding navigation links
 *
 * @param doc - Document to repair
 * @returns Repair result with statistics
 *
 * @example
 * ```typescript
 * const doc = await Document.load('document.docx');
 * const result = repairTOC(doc);
 * console.log(`Added ${result.topLinksAdded} navigation links`);
 * await doc.save('repaired.docx');
 * ```
 */
export function repairTOC(doc: Document): TOCRepairResult {
  // Step 1: Find Header 1 (title) position
  const titleIndex = findHeader1(doc);

  // Step 2: Find all Header 2 in 1x1 tables
  const header2s = findHeader2InTables(doc);

  if (header2s.length === 0) {
    return {
      tocGenerated: false,
      header2Count: 0,
      topLinksAdded: 0,
    };
  }

  // Step 3: Generate TOC field after title
  generateTOCField(doc, titleIndex + 1);

  // Step 4: Add "Top of the Document" links
  const topLinksAdded = addTopLinks(doc, header2s);

  return {
    tocGenerated: true,
    header2Count: header2s.length,
    topLinksAdded,
  };
}

/**
 * Find Header 1 (Title) in first 10 paragraphs
 */
function findHeader1(doc: Document): number {
  const paragraphs = doc.getParagraphs();

  for (let i = 0; i < Math.min(10, paragraphs.length); i++) {
    const para = paragraphs[i];
    if (!para) continue;

    const styleName = para.getStyle();

    if (
      styleName === 'Heading1' ||
      styleName === 'Header1' ||
      styleName === 'Heading 1' ||
      styleName === 'Title' ||
      styleName === 'title'
    ) {
      return i;
    }
  }

  return 0;
}

/**
 * Find all Header 2 paragraphs in 1x1 tables
 *
 * Searches for:
 * - Tables with exactly 1 row and 1 cell
 * - Paragraphs styled as Heading2/Header2/Heading 2
 * - With non-empty text content
 */
function findHeader2InTables(doc: Document): Header2Info[] {
  const header2s: Header2Info[] = [];
  const paragraphs = doc.getParagraphs();
  const tables = doc.getTables();

  for (const table of tables) {
    const rows = table.getRows();

    // Only check 1x1 tables
    if (rows.length !== 1) continue;

    const firstRow = rows[0];
    if (!firstRow) continue;

    const cells = firstRow.getCells();
    if (cells.length !== 1) continue;

    const firstCell = cells[0];
    if (!firstCell) continue;

    // Check paragraphs in the single cell
    const cellParas = firstCell.getParagraphs();

    for (const para of cellParas) {
      const styleName = para.getStyle();

      if (
        styleName === 'Heading2' ||
        styleName === 'Header2' ||
        styleName === 'Heading 2'
      ) {
        const text = para.getText().trim();
        if (!text) continue;

        const paragraphIndex = paragraphs.indexOf(para);

        header2s.push({
          paragraph: para,
          paragraphIndex,
          text,
        });
      }
    }
  }

  return header2s;
}

/**
 * Generate TOC field using Word's native TOC functionality
 *
 * Creates field with switches:
 * - \o "2-2" = Only outline level 2 (Header 2)
 * - \h = Hyperlinks enabled
 * - \z = Hide tab leader (no dots)
 * - \t = Use only specified styles
 *
 * Word auto-generates the TOC content when the field is updated.
 */
function generateTOCField(doc: Document, position: number): void {
  // Clear any existing TOC entries
  clearExistingTOC(doc, position);

  // Create TOC field
  const tocPara = new Paragraph();
  const field = new Field({
    type: 'TOC' as any,
    instruction: 'TOC \\o "2-2" \\h \\z \\t',
  });
  tocPara.addField(field);

  // Insert at position
  doc.insertParagraphAt(position, tocPara);

  // Add blank line after TOC
  const blankPara = new Paragraph();
  blankPara.addText(' ');
  doc.insertParagraphAt(position + 1, blankPara);
}

/**
 * Clear existing TOC section (orphaned hyperlinks)
 *
 * Removes up to 30 paragraphs after the position that contain hyperlinks
 * with anchors (likely old TOC entries).
 */
function clearExistingTOC(doc: Document, startIndex: number): void {
  const paragraphs = doc.getParagraphs();
  let removed = 0;

  for (let i = startIndex; i < Math.min(startIndex + 30, paragraphs.length); i++) {
    const para = paragraphs[i];
    if (para === undefined) continue;

    const content = para.getContent();
    let hasAnchorLink = false;

    for (const item of content) {
      if (item instanceof Hyperlink) {
        const anchor = (item as any).getAnchor?.();
        if (anchor && anchor !== '_top') {
          hasAnchorLink = true;
          break;
        }
      }
    }

    if (hasAnchorLink) {
      doc.removeParagraph(i - removed);
      removed++;
    } else {
      // Stop at first non-TOC paragraph
      break;
    }
  }
}

/**
 * Add "Top of the Document" links throughout document
 *
 * Adds links:
 * - Before each Header 2 (except first)
 * - Before proprietary notice at end of document
 *
 * Note: Uses "_top" built-in Word bookmark (no need to create it)
 */
function addTopLinks(doc: Document, header2s: Header2Info[]): number {
  let linksAdded = 0;

  // Add before each Header 2 (except first)
  for (let i = 1; i < header2s.length; i++) {
    const insertIndex = header2s[i]!.paragraphIndex;

    if (!hasTopLink(doc, insertIndex - 1)) {
      const topLink = createTopLink();
      doc.insertParagraphAt(insertIndex, topLink);
      linksAdded++;

      // Update indices for remaining headers
      for (let j = i; j < header2s.length; j++) {
        header2s[j]!.paragraphIndex++;
      }
    }
  }

  // Add before proprietary notice
  const noticeIndex = findProprietaryNotice(doc);
  if (noticeIndex > 0 && !hasTopLink(doc, noticeIndex - 1)) {
    const topLink = createTopLink();
    doc.insertParagraphAt(noticeIndex, topLink);
    linksAdded++;
  }

  return linksAdded;
}

/**
 * Create "Top of the Document" hyperlink paragraph
 *
 * Links to "_top" built-in Word bookmark.
 */
function createTopLink(): Paragraph {
  const para = new Paragraph();
  para.setAlignment('right');

  // Set spacing (in twips: 1 point = 20 twips, so 3 points = 60 twips)
  if (typeof (para as any).setSpacingBefore === 'function') {
    (para as any).setSpacingBefore(60);
  }
  if (typeof (para as any).setSpacingAfter === 'function') {
    (para as any).setSpacingAfter(0);
  }

  const hyperlink = new Hyperlink({
    anchor: '_top', // Built-in Word bookmark for document top
    text: 'Top of the Document', // Note: includes "the"
    formatting: {
      font: 'Verdana',
      size: 12,
      underline: 'single',
      color: '0000FF',
    },
  });

  para.addHyperlink(hyperlink);

  return para;
}

/**
 * Check if paragraph contains a "Top of the Document" link
 */
function hasTopLink(doc: Document, index: number): boolean {
  const paragraphs = doc.getParagraphs();
  if (index < 0 || index >= paragraphs.length) return false;

  const para = paragraphs[index];
  if (!para) return false;

  const content = para.getContent();

  for (const item of content) {
    if (item instanceof Hyperlink) {
      const text = (item as Hyperlink).getText().trim().toLowerCase();
      // Check for both "Top of the Document" and "Top of Document"
      if (
        text === 'top of the document' ||
        text === 'top of document'
      ) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Find proprietary notice paragraph (usually near end of document)
 */
function findProprietaryNotice(doc: Document): number {
  const paragraphs = doc.getParagraphs();
  const searchStart = Math.max(0, paragraphs.length - 20);

  for (let i = searchStart; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;

    const text = para.getText().toLowerCase();
    if (text.includes('proprietary') || text.includes('confidential')) {
      return i;
    }
  }

  return paragraphs.length;
}
