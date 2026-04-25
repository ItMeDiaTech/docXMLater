/**
 * DocumentParser - Handles parsing of DOCX files
 * Extracts content from ZIP archives and converts XML to structured data
 */

import { AlternateContent } from '../elements/AlternateContent';
import { Bookmark } from '../elements/Bookmark';
import { Endnote, EndnoteType } from '../elements/Endnote';
import { Footnote, FootnoteType } from '../elements/Footnote';
import { BookmarkManager } from '../elements/BookmarkManager';
import { Comment } from '../elements/Comment';
import { CustomXmlBlock } from '../elements/CustomXml';
import { PreservedElement } from '../elements/PreservedElement';
import { MathParagraph } from '../elements/MathElement';
import { ComplexField, Field } from '../elements/Field';
import { isHyperlinkInstruction, parseHyperlinkInstruction } from '../elements/FieldHelpers';
import { Footer } from '../elements/Footer';
import { Header } from '../elements/Header';
import { Hyperlink } from '../elements/Hyperlink';
import { ImageManager } from '../elements/ImageManager';
import { ImageRun } from '../elements/ImageRun';
import { Paragraph, ParagraphFormatting, ParagraphContent } from '../elements/Paragraph';
import { Revision } from '../elements/Revision';
import {
  BreakType,
  FormFieldCheckBox,
  FormFieldData,
  FormFieldDropDownList,
  FormFieldTextInput,
  Run,
  RunContent,
  RunFormatting,
} from '../elements/Run';
import { Section, SectionProperties, SectionType } from '../elements/Section';
import { StructuredDocumentTag } from '../elements/StructuredDocumentTag';
import { Table, TableBorder } from '../elements/Table';
import { TableCell } from '../elements/TableCell';
import { TableOfContents } from '../elements/TableOfContents';
import { TableOfContentsElement } from '../elements/TableOfContentsElement';
import { TableGridChange } from '../elements/TableGridChange';
import { TableRow } from '../elements/TableRow';
import { AbstractNumbering } from '../formatting/AbstractNumbering';
import { NumberingInstance } from '../formatting/NumberingInstance';
import { Style, StyleProperties, StyleType } from '../formatting/Style';
import { logParagraphContent, logParsing, logTextDirection } from '../utils/diagnostics';
import { getGlobalLogger, createScopedLogger, ILogger, defaultLogger } from '../utils/logger';
import {
  safeParseInt,
  isExplicitlySet,
  parseOoxmlBoolean,
  parseOnOffAttribute,
} from '../utils/parsingHelpers';
import { halfPointsToPoints } from '../utils/units';
import type { ShadingConfig } from '../elements/CommonTypes';

// Create scoped logger for DocumentParser operations
function getLogger(): ILogger {
  return createScopedLogger(getGlobalLogger(), 'DocumentParser');
}
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { XMLParser } from '../xml/XMLParser';
import { ZipHandler } from '../zip/ZipHandler';
import { DOCX_PATHS } from '../zip/types';
import type { DocumentProperties } from '../types/document-types';
import { BodyElement } from './DocumentContent';
import { RelationshipManager } from './RelationshipManager';

/**
 * Parse error tracking
 */
export interface ParseError {
  element: string;
  error: Error;
}

/**
 * DocumentParser handles all document parsing logic
 */
export class DocumentParser {
  private parseErrors: ParseError[] = [];
  private strictParsing: boolean;
  private bookmarkManager: BookmarkManager | null = null;
  /**
   * Current part name being parsed (e.g., 'header1.xml', 'footer1.xml').
   * Used to create composite keys for image relationship IDs to distinguish
   * between images in different parts (headers/footers have their own .rels files).
   */
  private currentPartName: string | undefined = undefined;

  constructor(strictParsing = false) {
    this.strictParsing = strictParsing;
  }

  /**
   * Gets accumulated parse errors/warnings
   */
  getParseErrors(): ParseError[] {
    return [...this.parseErrors];
  }

  /**
   * Clears accumulated parse errors
   */
  clearParseErrors(): void {
    this.parseErrors = [];
  }

  /**
   * Parses the document XML and extracts content
   * @param zipHandler - ZIP handler containing the document
   * @param relationshipManager - Relationship manager to populate
   * @param imageManager - Image manager to register parsed images
   * @returns Parsed body elements, properties, and updated relationship manager
   */
  async parseDocument(
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager,
    imageManager: ImageManager,
    bookmarkManager?: BookmarkManager
  ): Promise<{
    bodyElements: BodyElement[];
    properties: DocumentProperties;
    relationshipManager: RelationshipManager;
    styles: Style[];
    abstractNumberings: AbstractNumbering[];
    numberingInstances: NumberingInstance[];
    section: Section | null;
    namespaces: Record<string, string>;
    documentBackground?: {
      color?: string;
      themeColor?: string;
      themeTint?: string;
      themeShade?: string;
    };
  }> {
    const logger = getLogger();
    logger.info('Parsing document');

    // Store bookmarkManager for use in parsing methods
    this.bookmarkManager = bookmarkManager || null;

    // Verify the document exists
    const docXml = zipHandler.getFileAsString(DOCX_PATHS.DOCUMENT);
    if (!docXml) {
      logger.error('Invalid document: word/document.xml not found');
      throw new Error('Invalid document: word/document.xml not found');
    }

    logger.info('Parsing document.xml', { xmlSize: docXml.length });

    // Parse existing relationships to avoid ID collisions
    const parsedRelationshipManager = this.parseRelationships(zipHandler, relationshipManager);

    // Parse document properties
    const properties = this.parseProperties(zipHandler);

    // Parse body elements (paragraphs and tables)
    // Now includes image parsing support
    const bodyElements = await this.parseBodyElements(
      docXml,
      parsedRelationshipManager,
      zipHandler,
      imageManager
    );

    // Parse styles from styles.xml
    const styles = this.parseStyles(zipHandler);
    logger.info('Parsed styles', { styleCount: styles.length });

    // Parse numbering from numbering.xml
    const numbering = this.parseNumbering(zipHandler);
    logger.info('Parsed numbering', {
      abstractCount: numbering.abstractNumberings.length,
      instanceCount: numbering.numberingInstances.length,
    });

    // Parse section properties from document.xml
    const section = this.parseSectionProperties(docXml);

    // Parse document background (w:background) per ECMA-376 Part 1 §17.2.1
    const documentBackground = this.parseDocumentBackground(docXml);

    // Parse and preserve namespaces from the root <w:document> tag
    const namespaces = this.parseNamespaces(docXml);

    // Count element types for logging
    let paragraphCount = 0;
    let tableCount = 0;
    for (const elem of bodyElements) {
      if (elem instanceof Paragraph) paragraphCount++;
      else if (elem instanceof Table) tableCount++;
    }
    logger.info('Document parsed', {
      paragraphs: paragraphCount,
      tables: tableCount,
      totalElements: bodyElements.length,
    });

    // Log any parse errors
    if (this.parseErrors.length > 0) {
      logger.warn('Parse errors encountered', { errorCount: this.parseErrors.length });
    }

    return {
      bodyElements,
      properties,
      relationshipManager: parsedRelationshipManager,
      styles,
      abstractNumberings: numbering.abstractNumberings,
      numberingInstances: numbering.numberingInstances,
      section,
      namespaces,
      documentBackground,
    };
  }

  /**
   * Parses body elements from document XML
   * Extracts paragraphs and tables with their formatting
   * Uses XMLParser for safe position-based parsing (prevents ReDoS)
   *
   * CRITICAL: Preserves document order by parsing elements sequentially
   * instead of by type. This prevents massive content loss and corruption.
   */
  private async parseBodyElements(
    docXml: string,
    relationshipManager: RelationshipManager,
    zipHandler: ZipHandler,
    imageManager: ImageManager
  ): Promise<BodyElement[]> {
    const bodyElements: BodyElement[] = [];

    // Extract the body content using safe position-based parsing
    const bodyContent = XMLParser.extractBody(docXml);
    if (!bodyContent) {
      return bodyElements;
    }

    let pos = 0;
    let pendingBookmarkStarts: Bookmark[] = [];
    while (pos < bodyContent.length) {
      const nextP = this.findNextTopLevelTag(bodyContent, 'w:p', pos);
      const nextTbl = this.findNextTopLevelTag(bodyContent, 'w:tbl', pos);
      const nextSdt = this.findNextTopLevelTag(bodyContent, 'w:sdt', pos);
      const nextAC = this.findNextTopLevelTag(bodyContent, 'mc:AlternateContent', pos);
      const nextMath = this.findNextTopLevelTag(bodyContent, 'm:oMathPara', pos);
      const nextCXml = this.findNextTopLevelTag(bodyContent, 'w:customXml', pos);
      const nextAltChunk = this.findNextTopLevelTag(bodyContent, 'w:altChunk', pos);

      const candidates = [];
      if (nextP !== -1) candidates.push({ type: 'p', pos: nextP });
      if (nextTbl !== -1) candidates.push({ type: 'tbl', pos: nextTbl });
      if (nextSdt !== -1) candidates.push({ type: 'sdt', pos: nextSdt });
      if (nextAC !== -1) candidates.push({ type: 'alternateContent', pos: nextAC });
      if (nextMath !== -1) candidates.push({ type: 'mathParagraph', pos: nextMath });
      if (nextCXml !== -1) candidates.push({ type: 'customXml', pos: nextCXml });
      if (nextAltChunk !== -1) candidates.push({ type: 'altChunk', pos: nextAltChunk });

      if (candidates.length === 0) break;

      candidates.sort((a, b) => a.pos - b.pos);
      const next = candidates[0];

      if (next) {
        // Check for body-level bookmarkEnd elements BEFORE this element
        // These appear between the previous element and this one
        if (bodyElements.length > 0 && next.pos > pos) {
          const bookmarkEnds = this.extractBodyLevelBookmarkEnds(bodyContent, pos, next.pos);
          if (bookmarkEnds.length > 0) {
            // Attach to the previous element
            const prevElement = bodyElements[bodyElements.length - 1];
            if (prevElement instanceof Paragraph) {
              for (const bookmark of bookmarkEnds) {
                prevElement.addBookmarkEnd(bookmark);
              }
            } else if (prevElement instanceof Table) {
              // For tables, attach to the last paragraph in the last cell
              const lastPara = prevElement.getLastParagraph();
              if (lastPara) {
                for (const bookmark of bookmarkEnds) {
                  lastPara.addBookmarkEnd(bookmark);
                }
              }
            }
          }
        }

        // Check for body-level bookmarkStart elements BEFORE this element
        // These appear between the previous element and this one
        // Unlike bookmarkEnds (attached to previous), bookmarkStarts are
        // collected as pending and attached to the NEXT parsed element
        if (next.pos > pos) {
          const bookmarkStarts = this.extractBodyLevelBookmarkStarts(bodyContent, pos, next.pos);
          if (bookmarkStarts.length > 0) {
            pendingBookmarkStarts.push(...bookmarkStarts);
          }
        }

        // Check if this element is inside a w:del block (deleted via Track Changes)
        // If so, skip the entire w:del block including all its content
        if (this.isPositionInsideDel(bodyContent, next.pos)) {
          const delEndPos = this.findDelEndPosition(bodyContent, next.pos);
          if (delEndPos > 0) {
            pos = delEndPos;
          } else {
            pos = next.pos + 1;
          }
          continue;
        }

        if (next.type === 'p') {
          const elementXml = this.extractSingleElement(bodyContent, 'w:p', next.pos);

          if (elementXml) {
            // Parse paragraph with order preservation
            // Use the new method that preserves the exact order of runs and hyperlinks
            const paragraph = await this.parseParagraphWithOrder(
              elementXml,
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (paragraph) bodyElements.push(paragraph);
            pos = next.pos + elementXml.length;
          } else {
            pos = next.pos + 1;
          }
        } else if (next.type === 'tbl') {
          const elementXml = this.extractSingleElement(bodyContent, 'w:tbl', next.pos);
          if (elementXml) {
            // Parse XML to object, then extract the table content
            // IMPORTANT: trimValues: false preserves whitespace from xml:space="preserve" attributes
            const parsed = XMLParser.parseToObject(elementXml, {
              trimValues: false,
            });
            const table = await this.parseTableFromObject(
              parsed['w:tbl'],
              relationshipManager,
              zipHandler,
              imageManager,
              elementXml
            );
            if (table) bodyElements.push(table);
            pos = next.pos + elementXml.length;
          } else {
            pos = next.pos + 1;
          }
        } else if (next.type === 'sdt') {
          const elementXml = this.extractSingleElement(bodyContent, 'w:sdt', next.pos);
          if (elementXml) {
            // Parse XML to object, then extract the SDT content
            // IMPORTANT: trimValues: false preserves whitespace from xml:space="preserve" attributes
            const parsed = XMLParser.parseToObject(elementXml, {
              trimValues: false,
            });
            const sdt = await this.parseSDTFromObject(
              parsed['w:sdt'],
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (sdt) bodyElements.push(sdt);
            pos = next.pos + elementXml.length;
          } else {
            pos = next.pos + 1;
          }
        } else if (next.type === 'alternateContent') {
          // mc:AlternateContent - preserve as raw XML for round-trip fidelity
          const elementXml = this.extractSingleElement(
            bodyContent,
            'mc:AlternateContent',
            next.pos
          );
          if (elementXml) {
            bodyElements.push(new AlternateContent(elementXml));
            pos = next.pos + elementXml.length;
          } else {
            pos = next.pos + 1;
          }
        } else if (next.type === 'mathParagraph') {
          // m:oMathPara - preserve as raw XML for round-trip fidelity
          const elementXml = this.extractSingleElement(bodyContent, 'm:oMathPara', next.pos);
          if (elementXml) {
            bodyElements.push(new MathParagraph(elementXml));
            pos = next.pos + elementXml.length;
          } else {
            pos = next.pos + 1;
          }
        } else if (next.type === 'customXml') {
          // w:customXml - preserve as raw XML for round-trip fidelity
          const elementXml = this.extractSingleElement(bodyContent, 'w:customXml', next.pos);
          if (elementXml) {
            bodyElements.push(new CustomXmlBlock(elementXml));
            pos = next.pos + elementXml.length;
          } else {
            pos = next.pos + 1;
          }
        } else if (next.type === 'altChunk') {
          // w:altChunk - preserve as raw XML for round-trip fidelity
          const elementXml = this.extractSingleElement(bodyContent, 'w:altChunk', next.pos);
          if (elementXml) {
            bodyElements.push(new PreservedElement(elementXml, 'altChunk', 'block'));
            pos = next.pos + elementXml.length;
          } else {
            pos = next.pos + 1;
          }
        }

        // Attach any pending body-level bookmarkStarts to the just-parsed element
        if (pendingBookmarkStarts.length > 0 && bodyElements.length > 0) {
          const currentElement = bodyElements[bodyElements.length - 1];
          if (currentElement instanceof Paragraph) {
            for (const bookmark of pendingBookmarkStarts) {
              currentElement.addBookmarkStart(bookmark);
            }
          } else if (currentElement instanceof Table) {
            // For tables, attach to the first paragraph of the first cell
            const firstPara = currentElement.getFirstParagraph();
            if (firstPara) {
              for (const bookmark of pendingBookmarkStarts) {
                firstPara.addBookmarkStart(bookmark);
              }
            }
          }
          pendingBookmarkStarts = [];
        }
      }
    }

    // Check for any trailing body-level bookmarkEnd elements after the last element
    if (bodyElements.length > 0 && pos < bodyContent.length) {
      const trailingBookmarkEnds = this.extractBodyLevelBookmarkEnds(bodyContent, pos, -1);
      if (trailingBookmarkEnds.length > 0) {
        const lastElement = bodyElements[bodyElements.length - 1];
        if (lastElement instanceof Paragraph) {
          for (const bookmark of trailingBookmarkEnds) {
            lastElement.addBookmarkEnd(bookmark);
          }
        } else if (lastElement instanceof Table) {
          const lastPara = lastElement.getLastParagraph();
          if (lastPara) {
            for (const bookmark of trailingBookmarkEnds) {
              lastPara.addBookmarkEnd(bookmark);
            }
          }
        }
      }
    }

    // Check for any trailing body-level bookmarkStart elements after the last element
    if (bodyElements.length > 0 && pos < bodyContent.length) {
      const trailingBookmarkStarts = this.extractBodyLevelBookmarkStarts(bodyContent, pos, -1);
      if (trailingBookmarkStarts.length > 0) {
        const lastElement = bodyElements[bodyElements.length - 1];
        if (lastElement instanceof Paragraph) {
          for (const bookmark of trailingBookmarkStarts) {
            lastElement.addBookmarkStart(bookmark);
          }
        } else if (lastElement instanceof Table) {
          const lastPara = lastElement.getLastParagraph();
          if (lastPara) {
            for (const bookmark of trailingBookmarkStarts) {
              lastPara.addBookmarkStart(bookmark);
            }
          }
        }
      }
    }

    // Assemble multi-paragraph complex fields (e.g., TOC fields spanning multiple paragraphs)
    this.assembleMultiParagraphFields(bodyElements);

    // Validate that we didn't load an empty/corrupted document
    this.validateLoadedContent(bodyElements);

    return bodyElements;
  }

  /**
   * Extracts body-level bookmarkEnd elements between two positions in the content.
   * These are bookmarkEnd elements that appear outside of paragraphs/tables.
   * @param content - The body content XML
   * @param startPos - Start position to search from
   * @param endPos - End position to search to (or end of content if -1)
   * @returns Array of Bookmark objects for the bookmarkEnd elements
   */
  private extractBodyLevelBookmarkEnds(
    content: string,
    startPos: number,
    endPos: number
  ): Bookmark[] {
    const searchContent = endPos === -1 ? content.slice(startPos) : content.slice(startPos, endPos);

    return this.extractBookmarkEndsFromContent(searchContent);
  }

  /**
   * Extracts bookmarkEnd elements from any XML content string.
   * Uses position-based parsing via XMLParser (ECMA-376 compliant, ReDoS-safe).
   * @param content - The XML content to search
   * @returns Array of Bookmark objects for the bookmarkEnd elements
   */
  private extractBookmarkEndsFromContent(content: string): Bookmark[] {
    const bookmarks: Bookmark[] = [];

    // Use position-based XMLParser instead of regex for ECMA-376 compliance
    // This handles any attribute order and is safe from ReDoS attacks
    const bookmarkEndXmls = XMLParser.extractElements(content, 'w:bookmarkEnd');

    for (const bookmarkEndXml of bookmarkEndXmls) {
      const idAttr = XMLParser.extractAttribute(bookmarkEndXml, 'w:id');
      if (idAttr) {
        const id = parseInt(idAttr, 10);
        if (!isNaN(id)) {
          // CT_MarkupRange §17.13.5 — preserve w:displacedByCustomXml.
          const displacedAttr = XMLParser.extractAttribute(
            bookmarkEndXml,
            'w:displacedByCustomXml'
          );
          const displacedByCustomXml =
            displacedAttr === 'next' || displacedAttr === 'prev' ? displacedAttr : undefined;
          const bookmark = new Bookmark({
            name: `_end_${id}`,
            id: id,
            skipNormalization: true,
            displacedByCustomXml,
          });
          bookmarks.push(bookmark);
        }
      }
    }

    return bookmarks;
  }

  /**
   * Extracts body-level bookmarkStart elements between two positions in the content.
   * These are bookmarkStart elements that appear outside of paragraphs/tables.
   * @param content - The body content XML
   * @param startPos - Start position to search from
   * @param endPos - End position to search to (or end of content if -1)
   * @returns Array of Bookmark objects for the bookmarkStart elements
   */
  private extractBodyLevelBookmarkStarts(
    content: string,
    startPos: number,
    endPos: number
  ): Bookmark[] {
    const searchContent = endPos === -1 ? content.slice(startPos) : content.slice(startPos, endPos);

    return this.extractBookmarkStartsFromContent(searchContent);
  }

  /**
   * Extracts bookmarkStart elements from any XML content string.
   * Uses position-based parsing via XMLParser (ECMA-376 compliant, ReDoS-safe).
   * Reuses the existing parseBookmarkStart() method for consistent parsing.
   * @param content - The XML content to search
   * @returns Array of Bookmark objects for the bookmarkStart elements
   */
  private extractBookmarkStartsFromContent(content: string): Bookmark[] {
    const bookmarks: Bookmark[] = [];

    const bookmarkStartXmls = XMLParser.extractElements(content, 'w:bookmarkStart');

    for (const bookmarkStartXml of bookmarkStartXmls) {
      const bookmark = this.parseBookmarkStart(bookmarkStartXml);
      if (bookmark) {
        bookmarks.push(bookmark);
      }
    }

    return bookmarks;
  }

  /**
   * Finds the next occurrence of a tag in the content
   * Returns the position of the opening '<' or -1 if not found
   */
  private findNextTag(content: string, tagName: string, startPos: number): number {
    const tag = `<${tagName}`;
    let pos = content.indexOf(tag, startPos);

    while (pos !== -1) {
      // Verify this is the exact tag (not a prefix match like <w:p matching <w:pPr>)
      // The character after the tag name must be either '>', '/' or whitespace
      const charAfterTag = content[pos + tag.length];
      if (
        charAfterTag &&
        charAfterTag !== '>' &&
        charAfterTag !== '/' &&
        charAfterTag !== ' ' &&
        charAfterTag !== '\t' &&
        charAfterTag !== '\n' &&
        charAfterTag !== '\r'
      ) {
        // This is a prefix match (e.g., <w:pPr> when looking for <w:p>), skip it
        pos = content.indexOf(tag, pos + tag.length);
        continue;
      }
      return pos;
    }
    return -1;
  }

  /**
   * Finds the next TOP-LEVEL occurrence of a tag (not nested inside tables)
   * This prevents paragraphs inside table cells from being extracted as body paragraphs
   * Returns the position of the opening '<' or -1 if not found
   */
  private findNextTopLevelTag(content: string, tagName: string, startPos: number): number {
    let pos = startPos;

    while (pos < content.length) {
      // Find the next occurrence of the tag
      const tagPos = this.findNextTag(content, tagName, pos);
      if (tagPos === -1) {
        return -1; // No more tags found
      }

      // Check if this tag is nested inside a table
      // Look backwards from tagPos to see if we're inside an unclosed <w:tbl>
      const isInsideTable = this.isPositionInsideTable(content, tagPos);

      if (!isInsideTable) {
        // This is a top-level tag
        return tagPos;
      }

      // This tag is inside a table, skip past it and continue searching
      pos = tagPos + 1;
    }

    return -1;
  }

  /**
   * Counts unclosed opening tags of a specific element type in content.
   * Uses position-based parsing instead of regex for ECMA-376 compliance.
   * @param content - XML content to search
   * @param tagName - Element name (e.g., "w:tbl", "w:p", "w:sdt")
   * @returns Number of unclosed opening tags (opens - closes)
   */
  private countUnclosedTags(content: string, tagName: string): number {
    let opens = 0;
    let closes = 0;
    let pos = 0;

    // Count opening tags (handles self-closing correctly)
    while (pos < content.length) {
      const openPos = content.indexOf(`<${tagName}`, pos);
      if (openPos === -1) break;

      // Check the character after the tag name to ensure exact match
      const afterTagName = openPos + tagName.length + 1;
      if (afterTagName < content.length) {
        const charAfter = content.charAt(afterTagName);
        // Valid tag terminators: space, >, /, tab, newline
        if (
          charAfter === ' ' ||
          charAfter === '>' ||
          charAfter === '/' ||
          charAfter === '\t' ||
          charAfter === '\n' ||
          charAfter === '\r'
        ) {
          // Find the end of this tag
          const tagEnd = content.indexOf('>', openPos);
          if (tagEnd !== -1) {
            // Check if self-closing (ends with />)
            const isSelfClosing = content.charAt(tagEnd - 1) === '/';
            if (!isSelfClosing) {
              opens++;
            }
          }
        }
      }
      pos = openPos + 1;
    }

    // Count closing tags
    pos = 0;
    const closeTag = `</${tagName}>`;
    while (pos < content.length) {
      const closePos = content.indexOf(closeTag, pos);
      if (closePos === -1) break;
      closes++;
      pos = closePos + closeTag.length;
    }

    return opens - closes;
  }

  /**
   * Checks if a position in the content is inside a table element
   * Returns true if there's an unclosed <w:tbl> before this position
   */
  private isPositionInsideTable(content: string, position: number): boolean {
    const beforeContent = content.substring(0, position);
    return this.countUnclosedTags(beforeContent, 'w:tbl') > 0;
  }

  /**
   * Checks if a position in the content is inside a BODY-LEVEL w:del element
   * This is used to skip body-level elements that were deleted via Track Changes.
   *
   * IMPORTANT: This only detects body-level deletions (direct children of w:body).
   * Paragraph-level deletions (w:del inside w:p wrapping w:r runs) are NOT detected
   * because those are handled separately during run parsing.
   *
   * @param content - The body content XML
   * @param position - Position to check
   * @returns true if the position is inside a body-level <w:del> element
   */
  private isPositionInsideDel(content: string, position: number): boolean {
    // Look backward to find the most recent <w:del opening tag
    const beforeContent = content.substring(0, position);
    const lastDelOpen = beforeContent.lastIndexOf('<w:del');

    if (lastDelOpen === -1) return false;

    // Check if there's a </w:del> between the <w:del and our position
    const betweenContent = content.substring(lastDelOpen, position);
    if (betweenContent.includes('</w:del>')) return false;

    // Check if this <w:del is a self-closing tag (no content to skip)
    // Self-closing format: <w:del ... />
    const tagEnd = content.indexOf('>', lastDelOpen);
    if (tagEnd !== -1 && content.charAt(tagEnd - 1) === '/') return false;

    // Now check if this <w:del is BODY-LEVEL (not inside a w:p or w:tbl)
    // If the <w:del is inside a paragraph or table, it's NOT body-level
    const contentBeforeDel = content.substring(0, lastDelOpen);

    // Check if inside a paragraph using position-based parsing
    if (this.countUnclosedTags(contentBeforeDel, 'w:p') > 0) return false;

    // Check if inside a table
    if (this.countUnclosedTags(contentBeforeDel, 'w:tbl') > 0) return false;

    // Check if inside an SDT (structured document tag)
    if (this.countUnclosedTags(contentBeforeDel, 'w:sdt') > 0) return false;

    return true; // This is a body-level <w:del>
  }

  /**
   * Finds the end position of the w:del element that contains the given position
   * Returns the position after the closing </w:del> tag
   * @param content - The body content XML
   * @param startPos - Position inside the w:del element
   * @returns Position after the </w:del> tag, or -1 if not found
   */
  private findDelEndPosition(content: string, startPos: number): number {
    // Find the closing </w:del> tag after startPos
    // We need to handle nested w:del elements correctly
    let depth = 1;
    let pos = startPos;

    while (pos < content.length && depth > 0) {
      const nextOpen = content.indexOf('<w:del', pos);
      const nextClose = content.indexOf('</w:del>', pos);

      if (nextClose === -1) {
        // No closing tag found - malformed XML
        return -1;
      }

      if (nextOpen !== -1 && nextOpen < nextClose) {
        // Found another opening tag before the close
        depth++;
        pos = nextOpen + 6; // Move past "<w:del"
      } else {
        // Found a closing tag
        depth--;
        if (depth === 0) {
          // This is our closing tag
          return nextClose + '</w:del>'.length;
        }
        pos = nextClose + '</w:del>'.length;
      }
    }

    return -1;
  }

  /**
   * Extracts a single element from the content starting at the given position
   * Returns the complete element XML including opening and closing tags
   *
   * FIX (v1.3.1): Uses XMLParser.extractElements to ensure consistent extraction
   * behavior and prevent loss of self-closing elements like <w:tab/>
   * This fixes the TOC tab preservation bug where tabs were lost during extraction.
   */
  private extractSingleElement(content: string, tagName: string, startPos: number): string {
    // Extract the substring starting from the position
    const remainingContent = content.substring(startPos);

    // Use XMLParser.extractElements to get all elements of this type
    // This ensures we use the same proven extraction logic throughout
    const elements = XMLParser.extractElements(remainingContent, tagName);

    // Return the first element (which starts at position 0 in remainingContent)
    // This is the element at the specified startPos in the original content
    const extracted = elements.length > 0 ? elements[0]! : '';

    return extracted;
  }

  /**
   * Validates loaded content to detect corrupted or empty documents
   * Adds warnings if the document appears to have lost text content
   */
  private validateLoadedContent(bodyElements: BodyElement[]): void {
    const paragraphs = bodyElements.filter((el): el is Paragraph => el instanceof Paragraph);

    if (paragraphs.length === 0) {
      return; // Empty document is valid
    }

    // Count total runs and empty runs
    let totalRuns = 0;
    let emptyRuns = 0;
    let runsWithText = 0;

    for (const para of paragraphs) {
      const runs = para.getRuns();
      totalRuns += runs.length;

      for (const run of runs) {
        const text = run.getText();
        if (text.length === 0) {
          emptyRuns++;
        } else {
          runsWithText++;
        }
      }
    }

    // If more than 90% of runs are empty, warn about potential corruption
    if (totalRuns > 0) {
      const emptyPercentage = (emptyRuns / totalRuns) * 100;

      if (emptyPercentage > 90 && emptyRuns > 10) {
        const warning = new Error(
          `WARNING: Document appears to be corrupted or empty. ` +
            `${emptyRuns} out of ${totalRuns} runs (${emptyPercentage.toFixed(
              1
            )}%) have no text content. ` +
            `This may indicate:\n` +
            `  - The document was already corrupted before loading\n` +
            `  - Text content was stripped by another application\n` +
            `  - Encoding issues during document creation\n` +
            `Original document structure is preserved, but text may be lost.`
        );
        this.parseErrors.push({
          element: 'document-validation',
          error: warning,
        });

        // Always warn to console, even in non-strict mode
        defaultLogger.warn(`\nDocXML Load Warning:\n${warning.message}\n`);
      } else if (emptyPercentage > 50 && emptyRuns > 5) {
        const warning = new Error(
          `Document has ${emptyRuns} out of ${totalRuns} runs (${emptyPercentage.toFixed(
            1
          )}%) with no text. ` + `This is higher than normal and may indicate partial data loss.`
        );
        this.parseErrors.push({
          element: 'document-validation',
          error: warning,
        });
        defaultLogger.warn(`\nDocXML Load Warning:\n${warning.message}\n`);
      }
    }
  }

  /**
   * Extracts paragraph children (runs and hyperlinks) in document order
   * This ensures we preserve the original ordering of text and hyperlinks
   */

  private async parseParagraphWithOrder(
    paraXml: string,
    relationshipManager: RelationshipManager,
    zipHandler?: ZipHandler,
    imageManager?: ImageManager
  ): Promise<Paragraph | null> {
    try {
      const paragraph = new Paragraph();

      // Parse the paragraph object with order preservation
      // IMPORTANT: trimValues: false preserves whitespace from xml:space="preserve" attributes
      const paraObj = XMLParser.parseToObject(paraXml, { trimValues: false });
      const pElement = paraObj['w:p'] as any;
      if (!pElement) {
        return null;
      }

      // Parse paragraph properties
      const pPr = pElement['w:pPr'];
      this.parseParagraphPropertiesFromObject(pPr, paragraph);

      // If the paragraph has inline sectPr, extract the raw XML for round-trip fidelity.
      // The parsed object from parseParagraphPropertiesFromObject is a complex structure that
      // cannot be serialized correctly by XMLBuilder.wSelf (which expects flat attributes).
      // Using raw XML passthrough prevents corruption from malformed sectPr serialization.
      if (paragraph.formatting.sectPr) {
        const pPrElements = XMLParser.extractElements(paraXml, 'w:pPr');
        const pPrXml = pPrElements[0];
        if (pPrXml) {
          const sectPrElements = XMLParser.extractElements(pPrXml, 'w:sectPr');
          const rawSectPr = sectPrElements[0];
          if (rawSectPr) {
            // Store as raw XML string (including the <w:sectPr>...</w:sectPr> wrapper)
            paragraph.setSectionProperties(rawSectPr);
          }
        }
      }

      // Parse w14:paraId and w14:textId (Word 2010+ paragraph identifiers
      // per MC-DOCX §2.6.19, ST_LongHexNumber — 8-char hex string). These
      // are XML *attributes* on w:p, so XMLParser stores them under the
      // @_-prefixed keys. The previous lookup (`pElement['w14:paraId']`)
      // accessed an element-shaped key that never exists, silently
      // dropping both IDs on every load → save cycle. XMLParser's
      // numeric coercion of purely-digit hex strings (e.g. "00000001" →
      // 1) means we normalise back to the zero-padded 8-char form so
      // validators accept the output.
      const normaliseHexId = (raw: unknown): string | undefined => {
        if (raw === undefined || raw === null) return undefined;
        const asStr = typeof raw === 'number' ? raw.toString(16) : String(raw);
        return asStr.toUpperCase().padStart(8, '0');
      };
      const paraId = normaliseHexId(pElement['@_w14:paraId']);
      if (paraId) {
        paragraph.formatting.paraId = paraId;
      }
      const textId = normaliseHexId(pElement['@_w14:textId']);
      if (textId) {
        paragraph.formatting.textId = textId;
      }

      // CRITICAL FIX: Preserve document order of paragraph children (runs, hyperlinks, fields)
      // When XMLParser.parseToObject groups multiple runs/hyperlinks, it creates arrays
      // We need to reconstruct the original sequence by scanning the raw XML
      // This prevents punctuation from being reordered (e.g., "Heading." becoming ".Heading")
      await this.parseOrderedParagraphChildren(
        paraXml,
        pElement,
        paragraph,
        relationshipManager,
        zipHandler,
        imageManager
      );

      // NOTE: Complex field assembly (both single and multi-paragraph) is now handled
      // at the parseBodyElements level via assembleMultiParagraphFields(), which calls
      // assembleComplexFields() for single-paragraph fields after processing multi-paragraph ones.

      // Diagnostic logging for paragraph
      const runs = paragraph.getRuns();
      const runData = runs.map((run) => ({
        text: run.getText(),
        rtl: run.isRTL(),
      }));
      const bidi = paragraph.getFormatting().bidi;

      logParagraphContent('parsing', -1, runData, bidi);

      if (bidi) {
        logTextDirection(`Paragraph has BiDi enabled`);
      }

      // Merge consecutive hyperlinks with the same URL (handles Google Docs fragmentation)
      this.mergeConsecutiveHyperlinks(paragraph);

      return paragraph;
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'paragraph', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse paragraph: ${err.message}`);
      }

      return null;
    }
  }

  /**
   * CRITICAL FIX: Parses paragraph children in document order
   * This preserves the original sequence of runs, hyperlinks, and fields
   * Fixes issues where punctuation appears at the beginning instead of end
   * Example: ".Aetna" becomes "Aetna." by respecting original XML order
   *
   * @param paraXml - Raw paragraph XML
   * @param pElement - Parsed paragraph object
   * @param paragraph - Paragraph instance to populate
   * @param relationshipManager - Relationship manager
   * @param zipHandler - ZIP handler for resources
   * @param imageManager - Image manager for images
   */
  private async parseOrderedParagraphChildren(
    paraXml: string,
    pElement: any,
    paragraph: Paragraph,
    relationshipManager: RelationshipManager,
    zipHandler?: ZipHandler,
    imageManager?: ImageManager
  ): Promise<void> {
    // Extract the body portion (between pPr and closing tag) to scan for child elements
    // Find where w:pPr ends
    const pPrEnd = paraXml.indexOf('</w:pPr>');
    const contentStart = pPrEnd !== -1 ? pPrEnd + 8 : paraXml.indexOf('>') + 1;
    const contentEnd = paraXml.lastIndexOf('</w:p>');

    if (contentEnd <= contentStart) {
      return; // Empty paragraph
    }

    const paraContent = paraXml.substring(contentStart, contentEnd);

    // Track children by scanning XML for opening tags
    interface ChildMarker {
      type:
        | 'w:r'
        | 'w:hyperlink'
        | 'w:fldSimple'
        | 'w:ins'
        | 'w:del'
        | 'w:moveFrom'
        | 'w:moveTo'
        | 'w:bookmarkStart'
        | 'w:bookmarkEnd'
        | 'w:proofErr'
        | 'w:permStart'
        | 'w:permEnd'
        | 'm:oMath'
        | 'w:ruby'
        | 'w:commentRangeStart'
        | 'w:commentRangeEnd';
      pos: number;
      index: number;
    }

    const children: ChildMarker[] = [];
    let runIndex = 0;
    let hyperlinkIndex = 0;
    let fieldIndex = 0;
    let insIndex = 0;
    let delIndex = 0;
    let moveFromIndex = 0;
    let moveToIndex = 0;
    let bookmarkStartIndex = 0;
    let bookmarkEndIndex = 0;
    let proofErrIndex = 0;
    let permStartIndex = 0;
    let permEndIndex = 0;
    let oMathIndex = 0;
    let rubyIndex = 0;
    let commentRangeStartIndex = 0;
    let commentRangeEndIndex = 0;

    // Helper to find closing tag position for a given tag name starting from position
    const findClosingTagEnd = (content: string, tagName: string, startPos: number): number => {
      const closingTag = `</${tagName}>`;
      const closingPos = content.indexOf(closingTag, startPos);
      if (closingPos === -1) return startPos; // Fallback if not found
      return closingPos + closingTag.length;
    };

    // Helper to check if tag is self-closing
    const isSelfClosing = (tagContent: string): boolean => {
      return tagContent.endsWith('/');
    };

    // Scan for all first-level child elements in document order
    let searchPos = 0;
    while (searchPos < paraContent.length) {
      // Find the next opening tag
      const tagStart = paraContent.indexOf('<', searchPos);
      if (tagStart === -1) break;

      // Extract tag name
      const tagEnd = paraContent.indexOf('>', tagStart);
      if (tagEnd === -1) break;

      const tagContent = paraContent.substring(tagStart + 1, tagEnd);
      const tagName = tagContent.split(/[\s\/>]/)[0];
      const selfClosing = isSelfClosing(tagContent);

      if (tagName === 'w:r') {
        children.push({ type: 'w:r', pos: tagStart, index: runIndex++ });
        // Skip past closing tag to avoid counting nested elements
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, 'w:r', tagEnd);
      } else if (tagName === 'w:hyperlink') {
        children.push({
          type: 'w:hyperlink',
          pos: tagStart,
          index: hyperlinkIndex++,
        });
        // Skip past closing tag - hyperlinks contain nested runs we shouldn't count
        searchPos = selfClosing
          ? tagEnd + 1
          : findClosingTagEnd(paraContent, 'w:hyperlink', tagEnd);
      } else if (tagName === 'w:fldSimple') {
        children.push({
          type: 'w:fldSimple',
          pos: tagStart,
          index: fieldIndex++,
        });
        // Skip past closing tag
        searchPos = selfClosing
          ? tagEnd + 1
          : findClosingTagEnd(paraContent, 'w:fldSimple', tagEnd);
      } else if (tagName === 'w:ins') {
        children.push({
          type: 'w:ins',
          pos: tagStart,
          index: insIndex++,
        });
        // Skip past closing tag - ins contains nested runs
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, 'w:ins', tagEnd);
      } else if (tagName === 'w:del') {
        children.push({
          type: 'w:del',
          pos: tagStart,
          index: delIndex++,
        });
        // Skip past closing tag - del contains nested runs
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, 'w:del', tagEnd);
      } else if (tagName === 'w:moveFrom') {
        children.push({
          type: 'w:moveFrom',
          pos: tagStart,
          index: moveFromIndex++,
        });
        // Skip past closing tag
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, 'w:moveFrom', tagEnd);
      } else if (tagName === 'w:moveTo') {
        children.push({
          type: 'w:moveTo',
          pos: tagStart,
          index: moveToIndex++,
        });
        // Skip past closing tag
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, 'w:moveTo', tagEnd);
      } else if (tagName === 'w:bookmarkStart') {
        // Bookmark start markers - always self-closing
        children.push({
          type: 'w:bookmarkStart',
          pos: tagStart,
          index: bookmarkStartIndex++,
        });
        searchPos = tagEnd + 1;
      } else if (tagName === 'w:bookmarkEnd') {
        // Bookmark end markers - always self-closing
        children.push({
          type: 'w:bookmarkEnd',
          pos: tagStart,
          index: bookmarkEndIndex++,
        });
        searchPos = tagEnd + 1;
      } else if (tagName === 'w:commentRangeStart') {
        // Comment range start markers - always self-closing
        children.push({
          type: 'w:commentRangeStart',
          pos: tagStart,
          index: commentRangeStartIndex++,
        });
        searchPos = tagEnd + 1;
      } else if (tagName === 'w:commentRangeEnd') {
        // Comment range end markers - always self-closing
        children.push({
          type: 'w:commentRangeEnd',
          pos: tagStart,
          index: commentRangeEndIndex++,
        });
        searchPos = tagEnd + 1;
      } else if (tagName === 'w:proofErr') {
        // Proofing error markers - always self-closing
        children.push({
          type: 'w:proofErr',
          pos: tagStart,
          index: proofErrIndex++,
        });
        searchPos = tagEnd + 1;
      } else if (tagName === 'w:permStart') {
        // Permission range start - always self-closing
        children.push({
          type: 'w:permStart',
          pos: tagStart,
          index: permStartIndex++,
        });
        searchPos = tagEnd + 1;
      } else if (tagName === 'w:permEnd') {
        // Permission range end - always self-closing
        children.push({
          type: 'w:permEnd',
          pos: tagStart,
          index: permEndIndex++,
        });
        searchPos = tagEnd + 1;
      } else if (tagName === 'm:oMath') {
        // Inline math expression - preserve as raw XML
        children.push({
          type: 'm:oMath',
          pos: tagStart,
          index: oMathIndex++,
        });
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, 'm:oMath', tagEnd);
      } else if (tagName === 'w:ruby') {
        // Ruby text (phonetic guides) - preserve as raw XML
        children.push({
          type: 'w:ruby',
          pos: tagStart,
          index: rubyIndex++,
        });
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, 'w:ruby', tagEnd);
      } else {
        searchPos = tagEnd + 1;
      }
    }

    // Helper to extract raw run XML from paraContent using position
    const extractRunXmlAtPosition = (pos: number): string | null => {
      // Find the end of this run element
      const closeTag = '</w:r>';
      let depth = 1;
      let searchPos = paraContent.indexOf('>', pos) + 1;
      while (depth > 0 && searchPos < paraContent.length) {
        const nextOpen = paraContent.indexOf('<w:r', searchPos);
        const nextClose = paraContent.indexOf(closeTag, searchPos);
        if (nextClose === -1) break;
        if (nextOpen !== -1 && nextOpen < nextClose) {
          // Check if it's actually <w:r> not <w:rPr> etc
          const charAfter = paraContent[nextOpen + 4];
          if (charAfter === '>' || charAfter === ' ' || charAfter === '/') {
            depth++;
          }
          searchPos = nextOpen + 4;
        } else {
          depth--;
          if (depth === 0) {
            return paraContent.substring(pos, nextClose + closeTag.length);
          }
          searchPos = nextClose + closeTag.length;
        }
      }
      return null;
    };

    // Helper to extract element XML at a known position, depth-aware
    const extractElementXmlAtPosition = (pos: number, tagName: string): string => {
      const openTag = `<${tagName}`;
      const closeTag = `</${tagName}>`;
      const openEnd = paraContent.indexOf('>', pos);
      if (openEnd === -1) return '';
      // Self-closing
      if (paraContent[openEnd - 1] === '/') {
        return paraContent.substring(pos, openEnd + 1);
      }
      let depth = 1;
      let searchFrom = openEnd + 1;
      while (depth > 0 && searchFrom < paraContent.length) {
        const nextOpen = paraContent.indexOf(openTag, searchFrom);
        const nextClose = paraContent.indexOf(closeTag, searchFrom);
        if (nextClose === -1) break;
        if (nextOpen !== -1 && nextOpen < nextClose) {
          const charAfter = paraContent[nextOpen + openTag.length];
          if (
            charAfter === '>' ||
            charAfter === ' ' ||
            charAfter === '/' ||
            charAfter === '\t' ||
            charAfter === '\n' ||
            charAfter === '\r'
          ) {
            depth++;
          }
          searchFrom = nextOpen + openTag.length;
        } else {
          depth--;
          if (depth === 0) {
            return paraContent.substring(pos, nextClose + closeTag.length);
          }
          searchFrom = nextClose + closeTag.length;
        }
      }
      return '';
    };

    // Now process children in the order they were found
    for (const child of children) {
      if (child.type === 'w:r') {
        const runs = pElement['w:r'];
        const runArray = Array.isArray(runs) ? runs : runs ? [runs] : [];
        if (child.index < runArray.length) {
          const runObj = runArray[child.index];
          if (runObj['w:commentReference']) {
            // Comment reference run — preserve as raw XML so the
            // w:commentReference element inside survives round-trip.
            // Without this, Word can't link comment ranges to comments.xml.
            const runXml = extractRunXmlAtPosition(child.pos);
            if (runXml) {
              paragraph.addContent(new PreservedElement(runXml, 'w:r', 'inline'));
            }
          } else if (runObj['w:drawing']) {
            if (zipHandler && imageManager) {
              const imageRun = await this.parseDrawingFromObject(
                runObj['w:drawing'],
                zipHandler,
                relationshipManager,
                imageManager
              );
              if (imageRun) {
                paragraph.addRun(imageRun);
              }
            }
          } else if (runObj['w:pict']) {
            // VML graphics - preserve as raw XML for passthrough
            // Extract the raw run XML using position to get the exact w:pict content
            const runXml = extractRunXmlAtPosition(child.pos);
            if (runXml) {
              const pictXmls = XMLParser.extractElements(runXml, 'w:pict');
              if (pictXmls.length > 0 && pictXmls[0]) {
                const run = Run.createFromContent([{ type: 'vml', rawXml: pictXmls[0] }]);
                // Apply any run properties (formatting) to the VML run
                this.parseRunPropertiesFromObject(runObj['w:rPr'], run);
                paragraph.addRun(run);
              }
            }
          } else if (runObj['w:object']) {
            // Embedded OLE object - preserve as raw XML for round-trip fidelity
            // Per ECMA-376 Part 1 §17.3.3.19
            const runXml = extractRunXmlAtPosition(child.pos);
            if (runXml) {
              const objectXmls = XMLParser.extractElements(runXml, 'w:object');
              if (objectXmls.length > 0 && objectXmls[0]) {
                const run = Run.createFromContent([
                  { type: 'embeddedObject', rawXml: objectXmls[0] },
                ]);
                this.parseRunPropertiesFromObject(runObj['w:rPr'], run);
                paragraph.addRun(run);
              }
            }
          } else {
            const run = this.parseRunFromObject(runObj);
            if (run) {
              paragraph.addRun(run);
            }
          }
        }
      } else if (child.type === 'w:hyperlink') {
        const hyperlinks = pElement['w:hyperlink'];
        const hyperlinkArray = Array.isArray(hyperlinks)
          ? hyperlinks
          : hyperlinks
            ? [hyperlinks]
            : [];
        if (child.index < hyperlinkArray.length) {
          const hyperlinkObj = hyperlinkArray[child.index];

          // Hyperlinks containing tracked changes (w:del/w:ins inside w:hyperlink)
          // cannot survive parseHyperlinkFromObject round-trip — preserve as raw XML
          const hasRevisionChildren =
            hyperlinkObj['w:del'] ||
            hyperlinkObj['w:ins'] ||
            hyperlinkObj['w:moveFrom'] ||
            hyperlinkObj['w:moveTo'];
          if (hasRevisionChildren) {
            // Flatten revisions to make hyperlink editable (setUrl/setText).
            // Trades revision fidelity inside the hyperlink for editability.
            const flattenedObj = { ...hyperlinkObj };
            const allRuns: any[] = [];

            // Keep existing direct runs
            if (flattenedObj['w:r']) {
              const directRuns = Array.isArray(flattenedObj['w:r'])
                ? flattenedObj['w:r']
                : [flattenedObj['w:r']];
              allRuns.push(...directRuns);
            }

            // Unwrap w:ins runs (inserted content — keep)
            if (flattenedObj['w:ins']) {
              const insArr = Array.isArray(flattenedObj['w:ins'])
                ? flattenedObj['w:ins']
                : [flattenedObj['w:ins']];
              for (const ins of insArr) {
                if (ins['w:r']) {
                  const insRuns = Array.isArray(ins['w:r']) ? ins['w:r'] : [ins['w:r']];
                  allRuns.push(...insRuns);
                }
              }
            }

            // Unwrap w:moveTo runs (move destination — keep)
            if (flattenedObj['w:moveTo']) {
              const moveToArr = Array.isArray(flattenedObj['w:moveTo'])
                ? flattenedObj['w:moveTo']
                : [flattenedObj['w:moveTo']];
              for (const mt of moveToArr) {
                if (mt['w:r']) {
                  const mtRuns = Array.isArray(mt['w:r']) ? mt['w:r'] : [mt['w:r']];
                  allRuns.push(...mtRuns);
                }
              }
            }

            // Drop w:del and w:moveFrom (deleted/moved-away content)
            flattenedObj['w:r'] = allRuns.length > 0 ? allRuns : undefined;
            delete flattenedObj['w:del'];
            delete flattenedObj['w:ins'];
            delete flattenedObj['w:moveFrom'];
            delete flattenedObj['w:moveTo'];

            const result = this.parseHyperlinkFromObject(flattenedObj, relationshipManager);
            if (result.hyperlink) {
              paragraph.addHyperlink(result.hyperlink);
            }
            for (const bookmark of result.bookmarkStarts) {
              paragraph.addBookmarkStart(bookmark);
            }
            for (const bookmark of result.bookmarkEnds) {
              paragraph.addBookmarkEnd(bookmark);
            }
          } else {
            const result = this.parseHyperlinkFromObject(hyperlinkObj, relationshipManager);
            if (result.hyperlink) {
              paragraph.addHyperlink(result.hyperlink);
            }
            // Add any bookmarks found inside the hyperlink
            for (const bookmark of result.bookmarkStarts) {
              paragraph.addBookmarkStart(bookmark);
            }
            for (const bookmark of result.bookmarkEnds) {
              paragraph.addBookmarkEnd(bookmark);
            }
          }
        }
      } else if (child.type === 'w:fldSimple') {
        const fields = pElement['w:fldSimple'];
        const fieldArray = Array.isArray(fields) ? fields : fields ? [fields] : [];
        if (child.index < fieldArray.length) {
          const field = this.parseSimpleFieldFromObject(fieldArray[child.index]);
          if (field) {
            paragraph.addField(field);
          }
        }
      } else if (
        child.type === 'w:ins' ||
        child.type === 'w:del' ||
        child.type === 'w:moveFrom' ||
        child.type === 'w:moveTo'
      ) {
        const revisionXml = extractElementXmlAtPosition(child.pos, child.type);
        if (revisionXml) {
          // Detect nested revision elements (e.g., w:del inside w:ins)
          const innerContent = revisionXml.substring(revisionXml.indexOf('>') + 1);
          const hasNestedRevision = /<w:(del|moveFrom|moveTo|ins)\s/.test(innerContent);
          if (hasNestedRevision) {
            // Preserve entire nested structure as raw XML for round-trip fidelity
            paragraph.addContent(new PreservedElement(revisionXml, child.type, 'inline'));
          } else {
            const revResult = await this.parseRevisionFromXml(
              revisionXml,
              child.type,
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revResult.revision) {
              paragraph.addRevision(revResult.revision);
            }
            for (const bookmark of revResult.bookmarkStarts) {
              paragraph.addBookmarkStart(bookmark);
            }
            for (const bookmark of revResult.bookmarkEnds) {
              paragraph.addBookmarkEnd(bookmark);
            }
          }
        }
      } else if (child.type === 'w:bookmarkStart') {
        const endPos = paraContent.indexOf('>', child.pos);
        if (endPos !== -1) {
          const bookmarkXml = paraContent.substring(child.pos, endPos + 1);
          const bookmark = this.parseBookmarkStart(bookmarkXml);
          if (bookmark) {
            paragraph.addBookmarkStart(bookmark);
          }
        }
      } else if (child.type === 'w:bookmarkEnd') {
        const endPos = paraContent.indexOf('>', child.pos);
        if (endPos !== -1) {
          const bookmarkXml = paraContent.substring(child.pos, endPos + 1);
          const bookmark = this.parseBookmarkEnd(bookmarkXml);
          if (bookmark) {
            paragraph.addBookmarkEnd(bookmark);
          }
        }
      } else if (child.type === 'w:commentRangeStart' || child.type === 'w:commentRangeEnd') {
        // Preserve comment range markers as raw XML for round-trip fidelity.
        // Without these, comments.xml references become orphaned and Word
        // flags the document as corrupt ("unreadable content").
        const selfCloseEnd = paraContent.indexOf('>', child.pos);
        if (selfCloseEnd !== -1) {
          const rawXml = paraContent.substring(child.pos, selfCloseEnd + 1);
          paragraph.addContent(new PreservedElement(rawXml, child.type, 'inline'));
        }
      } else if (
        child.type === 'w:proofErr' ||
        child.type === 'w:permStart' ||
        child.type === 'w:permEnd'
      ) {
        // Preserve proofing errors and permission ranges as raw XML
        // These are self-closing elements: extract from pos to end of tag
        const selfCloseEnd = paraContent.indexOf('>', child.pos);
        if (selfCloseEnd !== -1) {
          const rawXml = paraContent.substring(child.pos, selfCloseEnd + 1);
          paragraph.addContent(new PreservedElement(rawXml, child.type, 'inline'));
        }
      } else if (child.type === 'm:oMath' || child.type === 'w:ruby') {
        // Preserve inline math and ruby elements as raw XML for round-trip fidelity
        const elementXml = this.extractSingleElement(paraContent, child.type, child.pos);
        if (elementXml) {
          paragraph.addContent(new PreservedElement(elementXml, child.type, 'inline'));
        }
      }
    }
  }

  /**
   * Parses a revision element from raw XML
   * Extracts revision metadata (id, author, date), contained runs, and any bookmarks
   * @param revisionXml - Raw XML of the revision element
   * @param tagName - Tag name (w:ins, w:del, w:moveFrom, w:moveTo)
   * @param relationshipManager - Relationship manager for image relationships
   * @param zipHandler - ZIP handler for loading image data
   * @param imageManager - Image manager for registering images
   * @returns Object with parsed Revision and any bookmarks found inside, or null
   */
  private async parseRevisionFromXml(
    revisionXml: string,
    tagName: string,
    relationshipManager: RelationshipManager,
    zipHandler?: ZipHandler,
    imageManager?: ImageManager
  ): Promise<{
    revision: Revision | null;
    bookmarkStarts: Bookmark[];
    bookmarkEnds: Bookmark[];
  }> {
    const result: {
      revision: Revision | null;
      bookmarkStarts: Bookmark[];
      bookmarkEnds: Bookmark[];
    } = { revision: null, bookmarkStarts: [], bookmarkEnds: [] };
    try {
      // Map XML tag to RevisionType
      let revisionType: import('../elements/Revision').RevisionType;
      switch (tagName) {
        case 'w:ins':
          revisionType = 'insert';
          break;
        case 'w:del':
          revisionType = 'delete';
          break;
        case 'w:moveFrom':
          revisionType = 'moveFrom';
          break;
        case 'w:moveTo':
          revisionType = 'moveTo';
          break;
        default:
          return result;
      }

      // Extract attributes
      const idAttr = XMLParser.extractAttribute(revisionXml, 'w:id');
      const author = XMLParser.extractAttribute(revisionXml, 'w:author');
      const dateAttr = XMLParser.extractAttribute(revisionXml, 'w:date');
      const moveId = XMLParser.extractAttribute(revisionXml, 'w:moveId');

      if (!idAttr || !author) {
        return result; // Required attributes missing
      }

      const id = parseInt(idAttr, 10);
      const date = dateAttr ? new Date(dateAttr) : new Date();

      // Extract content from revision element (runs and hyperlinks)
      // IMPORTANT: Extract hyperlinks FIRST, then extract runs from content
      // that is NOT inside hyperlinks to avoid duplicate content
      const hyperlinkXmls = XMLParser.extractElements(revisionXml, 'w:hyperlink');

      // Create a version of the XML with hyperlinks removed to extract standalone runs
      // Use split().join() instead of replace() to remove ALL occurrences of identical hyperlinks
      // (replace() only removes the first match, causing duplicate content)
      let xmlWithoutHyperlinks = revisionXml;
      for (const hyperlinkXml of hyperlinkXmls) {
        xmlWithoutHyperlinks = xmlWithoutHyperlinks.split(hyperlinkXml).join('');
      }

      // Extract runs from the XML without hyperlinks (these are standalone runs)
      const runXmls = XMLParser.extractElements(xmlWithoutHyperlinks, 'w:r');

      // Use RevisionContent to hold both Run and Hyperlink objects
      const content: import('../elements/RevisionContent').RevisionContent[] = [];

      // Parse standalone runs (not inside hyperlinks)
      for (const runXml of runXmls) {
        // Parse the run object
        const runObj = XMLParser.parseToObject(runXml, { trimValues: false });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const runElement = runObj['w:r'] as any;

        // Check if this run contains a drawing (image)
        if (runElement?.['w:drawing']) {
          if (zipHandler && imageManager) {
            const imageRun = await this.parseDrawingFromObject(
              runElement['w:drawing'],
              zipHandler,
              relationshipManager,
              imageManager
            );
            if (imageRun) {
              imageRun.setRawRunXml(runXml);
              content.push(imageRun);
            }
          }
        } else {
          // Parse as regular text run
          const run = this.parseRunFromObject(runElement);
          if (run) {
            content.push(run);
          }
        }
      }

      // Parse hyperlinks inside revision (for tracked hyperlink changes)
      for (const hyperlinkXml of hyperlinkXmls) {
        const hyperlinkObj = XMLParser.parseToObject(hyperlinkXml, { trimValues: false });
        const hyperlinkResult = this.parseHyperlinkFromObject(
          hyperlinkObj['w:hyperlink'],
          relationshipManager
        );
        if (hyperlinkResult.hyperlink) {
          content.push(hyperlinkResult.hyperlink);
        }
        // Collect bookmarks from hyperlinks inside revisions
        result.bookmarkStarts.push(...hyperlinkResult.bookmarkStarts);
        result.bookmarkEnds.push(...hyperlinkResult.bookmarkEnds);
      }

      // Extract bookmarks directly inside the revision (not nested in hyperlinks)
      const bookmarkStartXmls = XMLParser.extractElements(xmlWithoutHyperlinks, 'w:bookmarkStart');
      const bookmarkEndXmls = XMLParser.extractElements(xmlWithoutHyperlinks, 'w:bookmarkEnd');

      for (const bookmarkXml of bookmarkStartXmls) {
        const bookmark = this.parseBookmarkStart(bookmarkXml);
        if (bookmark) {
          result.bookmarkStarts.push(bookmark);
        }
      }

      for (const bookmarkXml of bookmarkEndXmls) {
        const bookmark = this.parseBookmarkEnd(bookmarkXml);
        if (bookmark) {
          result.bookmarkEnds.push(bookmark);
        }
      }

      if (content.length === 0) {
        // Log debug info for empty revisions (may indicate malformed XML)
        defaultLogger.debug('[DocumentParser] Empty revision content skipped', {
          tagName,
          id: idAttr,
          author,
        });
        // Still return any bookmarks found even if revision has no content
        return result;
      }

      // Create Revision instance
      const revision = new Revision({
        id,
        author,
        date,
        type: revisionType,
        content,
        moveId,
      });

      result.revision = revision;
      return result;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse revision:',
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return result;
    }
  }

  /**
   * Parses comments from word/comments.xml
   * @param commentsXml - Raw XML content of comments.xml
   * @returns Array of parsed Comment objects
   */
  parseCommentsXml(commentsXml: string): Comment[] {
    const comments: Comment[] = [];

    // Extract all w:comment elements
    const commentXmls = XMLParser.extractElements(commentsXml, 'w:comment');

    for (const commentXml of commentXmls) {
      const comment = this.parseCommentFromXml(commentXml);
      if (comment) {
        comments.push(comment);
      }
    }

    return comments;
  }

  /**
   * Parses a single comment element from XML
   * @param commentXml - XML string for one w:comment element
   * @returns Parsed Comment or null
   */
  private parseCommentFromXml(commentXml: string): Comment | null {
    try {
      // Extract attributes
      const idAttr = XMLParser.extractAttribute(commentXml, 'w:id');
      const author = XMLParser.extractAttribute(commentXml, 'w:author') || 'Unknown';
      const dateAttr = XMLParser.extractAttribute(commentXml, 'w:date');
      const initials = XMLParser.extractAttribute(commentXml, 'w:initials');
      const parentIdAttr =
        XMLParser.extractAttribute(commentXml, 'w15:parentId') ||
        XMLParser.extractAttribute(commentXml, 'w:parentId');
      const doneAttr =
        XMLParser.extractAttribute(commentXml, 'w15:done') ||
        XMLParser.extractAttribute(commentXml, 'w:done');

      if (!idAttr) {
        return null; // ID is required
      }

      const id = parseInt(idAttr, 10);
      const date = dateAttr ? new Date(dateAttr) : new Date();
      const parentId = parentIdAttr ? parseInt(parentIdAttr, 10) : undefined;
      // Per ECMA-376 §17.17.4, w:done is ST_OnOff — accept 1/0/true/false/on/off
      const done = parseOnOffAttribute(doneAttr);

      // Parse content (runs from paragraphs within the comment)
      const runs: Run[] = [];
      const runXmls = XMLParser.extractElements(commentXml, 'w:r');

      for (const runXml of runXmls) {
        const runObj = XMLParser.parseToObject(runXml, { trimValues: false });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const run = this.parseRunFromObject(runObj['w:r'] as any);
        if (run) {
          runs.push(run);
        }
      }

      // Create comment with parsed data
      const comment = new Comment({
        id,
        author,
        initials: initials || undefined,
        date,
        content: runs.length > 0 ? runs : '',
        parentId,
        done,
      });

      return comment;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse comment:',
        error instanceof Error ? { message: error.message } : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Parses footnotes.xml into Footnote array
   */
  parseFootnotesXml(footnotesXml: string): Footnote[] {
    const footnotes: Footnote[] = [];
    const footnoteXmls = XMLParser.extractElements(footnotesXml, 'w:footnote');

    for (const footnoteXml of footnoteXmls) {
      const footnote = this.parseFootnoteFromXml(footnoteXml);
      if (footnote) {
        footnotes.push(footnote);
      }
    }

    return footnotes;
  }

  private parseFootnoteFromXml(footnoteXml: string): Footnote | null {
    try {
      const idAttr = XMLParser.extractAttribute(footnoteXml, 'w:id');
      const typeAttr = XMLParser.extractAttribute(footnoteXml, 'w:type');

      if (idAttr === undefined) {
        return null;
      }

      const id = parseInt(idAttr, 10);

      let type: FootnoteType | undefined;
      if (typeAttr === 'separator') {
        type = FootnoteType.Separator;
      } else if (typeAttr === 'continuationSeparator') {
        type = FootnoteType.ContinuationSeparator;
      } else if (typeAttr === 'continuationNotice') {
        type = FootnoteType.ContinuationNotice;
      }

      const footnote = new Footnote({ id, type });

      const paraXmls = XMLParser.extractElements(footnoteXml, 'w:p');
      for (const paraXml of paraXmls) {
        const para = this.parseNoteParaFromXml(paraXml);
        if (para) {
          footnote.addParagraph(para);
        }
      }

      return footnote;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse footnote:',
        error instanceof Error ? { message: error.message } : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Parses endnotes.xml into Endnote array
   */
  parseEndnotesXml(endnotesXml: string): Endnote[] {
    const endnotes: Endnote[] = [];
    const endnoteXmls = XMLParser.extractElements(endnotesXml, 'w:endnote');

    for (const endnoteXml of endnoteXmls) {
      const endnote = this.parseEndnoteFromXml(endnoteXml);
      if (endnote) {
        endnotes.push(endnote);
      }
    }

    return endnotes;
  }

  private parseEndnoteFromXml(endnoteXml: string): Endnote | null {
    try {
      const idAttr = XMLParser.extractAttribute(endnoteXml, 'w:id');
      const typeAttr = XMLParser.extractAttribute(endnoteXml, 'w:type');

      if (idAttr === undefined) {
        return null;
      }

      const id = parseInt(idAttr, 10);

      let type: EndnoteType | undefined;
      if (typeAttr === 'separator') {
        type = EndnoteType.Separator;
      } else if (typeAttr === 'continuationSeparator') {
        type = EndnoteType.ContinuationSeparator;
      } else if (typeAttr === 'continuationNotice') {
        type = EndnoteType.ContinuationNotice;
      }

      const endnote = new Endnote({ id, type });

      const paraXmls = XMLParser.extractElements(endnoteXml, 'w:p');
      for (const paraXml of paraXmls) {
        const para = this.parseNoteParaFromXml(paraXml);
        if (para) {
          endnote.addParagraph(para);
        }
      }

      return endnote;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse endnote:',
        error instanceof Error ? { message: error.message } : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Parses a paragraph from a footnote or endnote element.
   * Uses run-level parsing (like comments) to avoid async dependency on relationship manager.
   */
  private parseNoteParaFromXml(paraXml: string): Paragraph | null {
    try {
      const para = new Paragraph();
      const runXmls = XMLParser.extractElements(paraXml, 'w:r');
      for (const runXml of runXmls) {
        const runObj = XMLParser.parseToObject(runXml, { trimValues: false });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const run = this.parseRunFromObject(runObj['w:r'] as any);
        if (run) {
          para.addRun(run);
        }
      }
      return para;
    } catch {
      return null;
    }
  }

  /**
   * Parses a bookmark start element from XML
   * @param bookmarkXml - Raw XML of the w:bookmarkStart element
   * @returns Bookmark instance or null
   */
  private parseBookmarkStart(bookmarkXml: string): Bookmark | null {
    try {
      const idAttr = XMLParser.extractAttribute(bookmarkXml, 'w:id');
      const nameAttr = XMLParser.extractAttribute(bookmarkXml, 'w:name');

      if (!idAttr || !nameAttr) {
        return null; // Required attributes missing
      }

      const id = parseInt(idAttr, 10);

      // Create bookmark with skipNormalization to preserve original name exactly
      // (Word allows special characters like = and . in bookmark names)
      // Parse optional column range for table bookmarks (ECMA-376 §17.16.5)
      const colFirstAttr = XMLParser.extractAttribute(bookmarkXml, 'w:colFirst');
      const colLastAttr = XMLParser.extractAttribute(bookmarkXml, 'w:colLast');
      // Parse optional w:displacedByCustomXml per CT_MarkupRange (§17.13.5).
      // Without this the attribute was dropped on load, so any Word document
      // with custom-XML-displaced bookmarks lost the disambiguator even
      // though the model now supports round-tripping it.
      const displacedAttr = XMLParser.extractAttribute(bookmarkXml, 'w:displacedByCustomXml');
      const displacedByCustomXml =
        displacedAttr === 'next' || displacedAttr === 'prev' ? displacedAttr : undefined;
      const bookmark = new Bookmark({
        name: nameAttr,
        id: id,
        skipNormalization: true,
        colFirst: colFirstAttr ? parseInt(colFirstAttr, 10) : undefined,
        colLast: colLastAttr ? parseInt(colLastAttr, 10) : undefined,
        displacedByCustomXml,
      });

      // Register with BookmarkManager to enable hasBookmark() checks
      // This prevents duplicate bookmarks when Template_UI adds bookmarks
      if (this.bookmarkManager) {
        try {
          this.bookmarkManager.registerExisting(bookmark);
        } catch (e) {
          // Bookmark might already be registered (duplicate in source doc)
          // Just log debug, don't fail - the bookmark is still valid for output
          defaultLogger.debug('[DocumentParser] Bookmark already registered:', {
            name: nameAttr,
            id: id,
          });
        }
      }

      return bookmark;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse bookmark start:',
        error instanceof Error ? { message: error.message } : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Parses a bookmark end element from XML
   * @param bookmarkXml - Raw XML of the w:bookmarkEnd element
   * @returns Bookmark instance or null
   */
  private parseBookmarkEnd(bookmarkXml: string): Bookmark | null {
    try {
      const idAttr = XMLParser.extractAttribute(bookmarkXml, 'w:id');

      if (!idAttr) {
        return null; // Required attribute missing
      }

      const id = parseInt(idAttr, 10);

      // CT_MarkupRange (§17.13.5) also permits w:displacedByCustomXml on
      // the end marker. Previously dropped on load, so a Word document
      // whose bookmark-end was displaced across a custom-XML node lost
      // the disambiguator even though the Bookmark model already emits
      // it from toEndXML().
      const displacedAttr = XMLParser.extractAttribute(bookmarkXml, 'w:displacedByCustomXml');
      const displacedByCustomXml =
        displacedAttr === 'next' || displacedAttr === 'prev' ? displacedAttr : undefined;

      // Create a placeholder bookmark for the end marker
      // The name doesn't matter for bookmarkEnd as it only uses the ID
      const bookmark = new Bookmark({
        name: `_end_${id}`,
        id: id,
        skipNormalization: true,
        displacedByCustomXml,
      });

      return bookmark;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse bookmark end:',
        error instanceof Error ? { message: error.message } : { error: String(error) }
      );
      return null;
    }
  }

  private async parseParagraphFromObject(
    paraObj: any,
    relationshipManager: RelationshipManager,
    zipHandler?: ZipHandler,
    imageManager?: ImageManager
  ): Promise<Paragraph | null> {
    try {
      const paragraph = new Paragraph();

      // Parse w14:paraId and w14:textId attributes from paragraph element
      // (Word 2010+, ST_LongHexNumber 8-char hex). XMLParser keys
      // attributes under the @_ prefix and may numeric-coerce purely-
      // digit hex strings like "00000001" to the number 1 — normalise
      // back to 8-char uppercase hex so the output passes strict
      // validation. The prior code used the un-prefixed element-shaped
      // keys and always saw `undefined`.
      const normaliseHexId = (raw: unknown): string | undefined => {
        if (raw === undefined || raw === null) return undefined;
        const asStr = typeof raw === 'number' ? raw.toString(16) : String(raw);
        return asStr.toUpperCase().padStart(8, '0');
      };
      const paraId = normaliseHexId(paraObj['@_w14:paraId']);
      if (paraId) {
        paragraph.formatting.paraId = paraId;
      }
      const textId = normaliseHexId(paraObj['@_w14:textId']);
      if (textId) {
        paragraph.formatting.textId = textId;
      }

      // Parse paragraph properties
      this.parseParagraphPropertiesFromObject(paraObj['w:pPr'], paragraph);

      // Check if we have ordered children metadata from the enhanced parser
      const orderedChildren = paraObj._orderedChildren as
        | { type: string; index: number }[]
        | undefined;

      if (orderedChildren && orderedChildren.length > 0) {
        // Use the preserved order from the parser
        for (const childInfo of orderedChildren) {
          const elementType = childInfo.type;
          const elementIndex = childInfo.index;

          if (elementType === 'w:r') {
            const runs = paraObj['w:r'];
            const runArray = Array.isArray(runs) ? runs : runs ? [runs] : [];
            if (elementIndex < runArray.length) {
              const child = runArray[elementIndex];
              if (child['w:drawing']) {
                if (zipHandler && imageManager) {
                  const imageRun = await this.parseDrawingFromObject(
                    child['w:drawing'],
                    zipHandler,
                    relationshipManager,
                    imageManager
                  );
                  if (imageRun) {
                    paragraph.addRun(imageRun);
                  }
                }
              } else {
                const run = this.parseRunFromObject(child);
                if (run) {
                  paragraph.addRun(run);
                }
              }
            }
          } else if (elementType === 'w:hyperlink') {
            const hyperlinks = paraObj['w:hyperlink'];
            const hyperlinkArray = Array.isArray(hyperlinks)
              ? hyperlinks
              : hyperlinks
                ? [hyperlinks]
                : [];
            if (elementIndex < hyperlinkArray.length) {
              const result = this.parseHyperlinkFromObject(
                hyperlinkArray[elementIndex],
                relationshipManager
              );
              if (result.hyperlink) {
                paragraph.addHyperlink(result.hyperlink);
              }
              // Add any bookmarks found inside the hyperlink
              for (const bookmark of result.bookmarkStarts) {
                paragraph.addBookmarkStart(bookmark);
              }
              for (const bookmark of result.bookmarkEnds) {
                paragraph.addBookmarkEnd(bookmark);
              }
            }
          } else if (elementType === 'w:fldSimple') {
            const fields = paraObj['w:fldSimple'];
            const fieldArray = Array.isArray(fields) ? fields : fields ? [fields] : [];
            if (elementIndex < fieldArray.length) {
              const field = this.parseSimpleFieldFromObject(fieldArray[elementIndex]);
              if (field) {
                paragraph.addField(field);
              }
            }
          }
        }
      } else {
        // Fallback to sequential processing if no order metadata
        // Handle runs (w:r)
        const runs = paraObj['w:r'];
        const runChildren = Array.isArray(runs) ? runs : runs ? [runs] : [];

        for (const child of runChildren) {
          if (child['w:drawing']) {
            if (zipHandler && imageManager) {
              // Parse as image run
              const imageRun = await this.parseDrawingFromObject(
                child['w:drawing'],
                zipHandler,
                relationshipManager,
                imageManager
              );
              if (imageRun) {
                paragraph.addRun(imageRun);
              }
            }
          } else {
            // Parse as normal text run
            const run = this.parseRunFromObject(child);
            if (run) {
              paragraph.addRun(run);
            }
          }
        }

        // Handle hyperlinks (w:hyperlink)
        const hyperlinks = paraObj['w:hyperlink'];
        const hyperlinkChildren = Array.isArray(hyperlinks)
          ? hyperlinks
          : hyperlinks
            ? [hyperlinks]
            : [];

        for (const hyperlinkObj of hyperlinkChildren) {
          const result = this.parseHyperlinkFromObject(hyperlinkObj, relationshipManager);
          if (result.hyperlink) {
            paragraph.addHyperlink(result.hyperlink);
          }
          // Add any bookmarks found inside the hyperlink
          for (const bookmark of result.bookmarkStarts) {
            paragraph.addBookmarkStart(bookmark);
          }
          for (const bookmark of result.bookmarkEnds) {
            paragraph.addBookmarkEnd(bookmark);
          }
        }

        // Handle simple fields (w:fldSimple)
        const fields = paraObj['w:fldSimple'];
        const fieldChildren = Array.isArray(fields) ? fields : fields ? [fields] : [];

        for (const fieldObj of fieldChildren) {
          const field = this.parseSimpleFieldFromObject(fieldObj);
          if (field) {
            paragraph.addField(field);
          }
        }
      }

      // Merge consecutive hyperlinks with the same URL (handles Google Docs fragmentation)
      this.mergeConsecutiveHyperlinks(paragraph);

      return paragraph;
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'paragraph', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse paragraph: ${err.message}`);
      }

      // In lenient mode, log warning and continue
      return null;
    }
  }

  private parseParagraphPropertiesFromObject(pPrObj: any, paragraph: Paragraph): void {
    if (!pPrObj) return;

    // Paragraph mark run properties (w:rPr within w:pPr) per ECMA-376 Part 1 §17.3.1.29
    // This controls formatting of the paragraph mark (¶ symbol) itself
    if (pPrObj['w:rPr']) {
      const rPrObj = pPrObj['w:rPr'];
      // Create a temporary Run to use the existing parseRunPropertiesFromObject method
      const tempRun = new Run('');
      this.parseRunPropertiesFromObject(rPrObj, tempRun);
      // Extract the formatting and set it as paragraph mark properties
      paragraph.setParagraphMarkFormatting(tempRun.getFormatting());

      // Transfer w:rPrChange (CT_ParaRPrChange, §17.3.1.30) from the
      // temp run onto the paragraph's formatting. Without this the
      // paragraph-mark rPrChange is silently dropped because
      // `tempRun.getFormatting()` exposes RunFormatting fields only —
      // `propertyChangeRevision` is a separate field on Run that was
      // previously discarded along with the temp run.
      const rPrChangeRev = tempRun.getPropertyChangeRevision();
      if (rPrChangeRev) {
        paragraph.formatting.paragraphMarkRunPropertiesChange = {
          id: rPrChangeRev.id,
          author: rPrChangeRev.author,
          date: rPrChangeRev.date,
          previousProperties: rPrChangeRev.previousProperties,
        };
      }

      // Parse paragraph mark deletion tracking (w:del in w:pPr/w:rPr)
      // Per ECMA-376 Part 1 §17.13.5.14 - indicates the paragraph mark was deleted
      if (rPrObj['w:del']) {
        // Handle array case (malformed XML with duplicate w:del elements)
        const rawDel = rPrObj['w:del'];
        const delObj = Array.isArray(rawDel) ? rawDel[0] : rawDel;
        if (delObj && typeof delObj === 'object') {
          const id = parseInt(delObj['@_w:id'] || '0', 10) || 0;
          const author = String(delObj['@_w:author'] ?? '');
          const dateStr = delObj['@_w:date'];
          const parsedDate = dateStr ? new Date(String(dateStr)) : new Date();
          const date = isNaN(parsedDate.getTime()) ? new Date() : parsedDate;
          paragraph.markParagraphMarkAsDeleted(id, author, date);
        }
      }

      // Parse paragraph mark insertion tracking (w:ins in w:pPr/w:rPr)
      // Per ECMA-376 Part 1 §17.13.5.18 - indicates the paragraph mark was inserted
      if (rPrObj['w:ins']) {
        // Handle array case (malformed XML with duplicate w:ins elements)
        const rawIns = rPrObj['w:ins'];
        const insObj = Array.isArray(rawIns) ? rawIns[0] : rawIns;
        if (insObj && typeof insObj === 'object') {
          const id = parseInt(insObj['@_w:id'] || '0', 10) || 0;
          const author = String(insObj['@_w:author'] ?? '');
          const dateStr = insObj['@_w:date'];
          const parsedDate = dateStr ? new Date(String(dateStr)) : new Date();
          const date = isNaN(parsedDate.getTime()) ? new Date() : parsedDate;
          paragraph.markParagraphMarkAsInserted(id, author, date);
        }
      }
    }

    // Alignment
    // XMLParser adds @_ prefix to attributes
    if (pPrObj['w:jc']?.['@_w:val']) {
      paragraph.setAlignment(pPrObj['w:jc']['@_w:val']);
    }

    // Style (w:pStyle per ECMA-376 §17.3.1.27 — `w:val` is ST_String
    // referencing a style ID). Cast via String(...) so purely-numeric
    // style IDs that XMLParser's `parseAttributeValue: true` coerces to
    // JS numbers (e.g., a custom styleId of "1") survive as strings,
    // matching the `style?: string` field contract on
    // ParagraphFormatting.
    if (pPrObj['w:pStyle']?.['@_w:val'] !== undefined) {
      paragraph.setStyle(String(pPrObj['w:pStyle']['@_w:val']));
    }

    // Indentation
    // Note: XMLParser converts numeric strings to numbers, so "0" becomes 0 (falsy)
    // Must use !== undefined check instead of truthy check to handle left="0"
    if (pPrObj['w:ind']) {
      const ind = pPrObj['w:ind'];
      // Use isExplicitlySet and safeParseInt for robust zero-value handling
      // Per ECMA-376 §17.3.1.15: w:start/w:end are bidi-aware alternatives to w:left/w:right
      const leftVal = ind['@_w:start'] ?? ind['@_w:left'];
      const rightVal = ind['@_w:end'] ?? ind['@_w:right'];
      if (isExplicitlySet(leftVal)) paragraph.setLeftIndent(safeParseInt(leftVal));
      if (isExplicitlySet(rightVal)) paragraph.setRightIndent(safeParseInt(rightVal));
      if (isExplicitlySet(ind['@_w:firstLine']))
        paragraph.setFirstLineIndent(safeParseInt(ind['@_w:firstLine']));
      // Parse hanging indent per ECMA-376 Part 1 §17.3.1.17
      if (isExplicitlySet(ind['@_w:hanging']))
        paragraph.setHangingIndent(safeParseInt(ind['@_w:hanging']));

      // CJK character-unit indentation attributes per ECMA-376 §17.3.1.12.
      // start/endChars are bidi-aware alternatives to left/rightChars; collapse
      // them onto the leftChars/rightChars fields the same way the twips parser
      // collapses w:start → left. Values are ST_DecimalNumber (hundredths of a
      // character unit), and 0 is a legitimate value — use isExplicitlySet so
      // number-0 from XMLParser.parseAttributeValue is preserved.
      if (!paragraph.formatting.indentation) paragraph.formatting.indentation = {};
      const leftCharsVal = ind['@_w:startChars'] ?? ind['@_w:leftChars'];
      const rightCharsVal = ind['@_w:endChars'] ?? ind['@_w:rightChars'];
      if (isExplicitlySet(leftCharsVal))
        paragraph.formatting.indentation.leftChars = safeParseInt(leftCharsVal);
      if (isExplicitlySet(rightCharsVal))
        paragraph.formatting.indentation.rightChars = safeParseInt(rightCharsVal);
      if (isExplicitlySet(ind['@_w:firstLineChars']))
        paragraph.formatting.indentation.firstLineChars = safeParseInt(ind['@_w:firstLineChars']);
      if (isExplicitlySet(ind['@_w:hangingChars']))
        paragraph.formatting.indentation.hangingChars = safeParseInt(ind['@_w:hangingChars']);
    }

    // Spacing (ECMA-376 §17.3.1.33 — 8 attributes)
    if (pPrObj['w:spacing']) {
      const spacing = pPrObj['w:spacing'];
      // Use isExplicitlySet to properly handle 0 values (0 spacing is valid)
      if (isExplicitlySet(spacing['@_w:before']))
        paragraph.setSpaceBefore(safeParseInt(spacing['@_w:before']));
      if (isExplicitlySet(spacing['@_w:after']))
        paragraph.setSpaceAfter(safeParseInt(spacing['@_w:after']));
      if (isExplicitlySet(spacing['@_w:line'])) {
        paragraph.setLineSpacing(safeParseInt(spacing['@_w:line']), spacing['@_w:lineRule']);
      }
      // Parse extended spacing attributes — write directly to paragraph.formatting
      // (getFormatting() returns a shallow copy, so we must access the internal object)
      if (!paragraph.formatting.spacing) paragraph.formatting.spacing = {};
      if (isExplicitlySet(spacing['@_w:beforeLines']))
        paragraph.formatting.spacing.beforeLines = safeParseInt(spacing['@_w:beforeLines']);
      if (isExplicitlySet(spacing['@_w:afterLines']))
        paragraph.formatting.spacing.afterLines = safeParseInt(spacing['@_w:afterLines']);
      // ST_OnOff per ECMA-376 §17.17.4 — accept 1/0/true/false/on/off
      const beforeAuto = spacing['@_w:beforeAutospacing'];
      if (beforeAuto !== undefined)
        paragraph.formatting.spacing.beforeAutospacing = parseOnOffAttribute(beforeAuto);
      const afterAuto = spacing['@_w:afterAutospacing'];
      if (afterAuto !== undefined)
        paragraph.formatting.spacing.afterAutospacing = parseOnOffAttribute(afterAuto);
    }

    // Keep properties — preserve explicit val="0" to override style inheritance
    // Parse pageBreakBefore FIRST, then keep properties (triggers automatic conflict resolution)
    if (pPrObj['w:pageBreakBefore'] !== undefined) {
      paragraph.formatting.pageBreakBefore = parseOoxmlBoolean(pPrObj['w:pageBreakBefore']);
    }
    if (pPrObj['w:keepNext'] !== undefined) {
      paragraph.setKeepNext(parseOoxmlBoolean(pPrObj['w:keepNext']));
    }
    if (pPrObj['w:keepLines'] !== undefined) {
      paragraph.setKeepLines(parseOoxmlBoolean(pPrObj['w:keepLines']));
    }

    // Contextual spacing
    if (pPrObj['w:contextualSpacing'] !== undefined) {
      paragraph.setContextualSpacing(parseOoxmlBoolean(pPrObj['w:contextualSpacing']));
    }

    // Numbering
    // Note: When track changes are present (w:pPrChange), XMLParser merges the
    // current numPr and the old numPr from pPrChange into an array. The first
    // element is always the current value per _orderedChildren ordering.
    if (pPrObj['w:numPr']) {
      const numPrRaw = pPrObj['w:numPr'];
      // Handle both single object and array (when pPrChange is present)
      const numPr = Array.isArray(numPrRaw) ? numPrRaw[0] : numPrRaw;
      const numId = numPr?.['w:numId']?.['@_w:val'];
      const ilvl = numPr?.['w:ilvl']?.['@_w:val'] || '0';
      if (numId !== undefined && numId !== null) {
        const parsedNumId = parseInt(numId, 10);
        const parsedIlvl = parseInt(ilvl, 10);
        if (!isNaN(parsedNumId) && !isNaN(parsedIlvl)) {
          if (parsedNumId === 0) {
            paragraph.formatting.numberingSuppressed = true;
          } else {
            paragraph.setNumbering(parsedNumId, parsedIlvl);
          }
        }
      }
    }

    // Borders per ECMA-376 Part 1 §17.3.1.24
    if (pPrObj['w:pBdr']) {
      const pBdr = pPrObj['w:pBdr'];
      const borders: any = {};

      // Helper function to parse border definition.
      // Covers the full CT_Border attribute set per ECMA-376 §17.18.2:
      // w:val, w:sz, w:color, w:space, w:themeColor, w:themeTint,
      // w:themeShade, w:shadow, w:frame. The last two are ST_OnOff —
      // route through parseOnOffAttribute so "off"/"false"/"0"/"on"
      // all resolve correctly even after XMLParser numeric coercion.
      const parseBorder = (borderObj: any): any => {
        if (!borderObj) return undefined;
        const border: any = {};
        if (borderObj['@_w:val']) border.style = borderObj['@_w:val'];
        if (borderObj['@_w:sz'] !== undefined) border.size = safeParseInt(borderObj['@_w:sz']);
        if (borderObj['@_w:color']) border.color = borderObj['@_w:color'];
        if (borderObj['@_w:space'] !== undefined)
          border.space = safeParseInt(borderObj['@_w:space']);
        if (borderObj['@_w:themeColor']) border.themeColor = String(borderObj['@_w:themeColor']);
        if (borderObj['@_w:themeTint']) border.themeTint = String(borderObj['@_w:themeTint']);
        if (borderObj['@_w:themeShade']) border.themeShade = String(borderObj['@_w:themeShade']);
        if (borderObj['@_w:shadow'] !== undefined) {
          border.shadow = parseOnOffAttribute(String(borderObj['@_w:shadow']), true);
        }
        if (borderObj['@_w:frame'] !== undefined) {
          border.frame = parseOnOffAttribute(String(borderObj['@_w:frame']), true);
        }
        return Object.keys(border).length > 0 ? border : undefined;
      };

      // Parse each border side
      if (pBdr['w:top']) borders.top = parseBorder(pBdr['w:top']);
      if (pBdr['w:bottom']) borders.bottom = parseBorder(pBdr['w:bottom']);
      if (pBdr['w:left']) borders.left = parseBorder(pBdr['w:left']);
      if (pBdr['w:right']) borders.right = parseBorder(pBdr['w:right']);
      if (pBdr['w:between']) borders.between = parseBorder(pBdr['w:between']);
      if (pBdr['w:bar']) borders.bar = parseBorder(pBdr['w:bar']);

      if (Object.keys(borders).length > 0) {
        paragraph.setBorder(borders);
      }
    }

    // Shading per ECMA-376 Part 1 §17.3.1.32
    if (pPrObj['w:shd']) {
      const shading = this.parseShadingFromObj(pPrObj['w:shd']);
      if (shading) {
        paragraph.setShading(shading);
      }
    }

    // Tab stops per ECMA-376 Part 1 §17.3.1.38
    if (pPrObj['w:tabs']) {
      const tabsObj = pPrObj['w:tabs'];
      const tabs: any[] = [];

      // Handle both single tab and array of tabs
      const tabElements = Array.isArray(tabsObj['w:tab'])
        ? tabsObj['w:tab']
        : tabsObj['w:tab']
          ? [tabsObj['w:tab']]
          : [];

      for (const tabObj of tabElements) {
        const tab: any = {};
        // w:pos is REQUIRED per §17.3.1.38 and is ST_SignedTwipsMeasure — 0 and
        // negative values are both valid. Use `!== undefined` so that XMLParser's
        // parseAttributeValue coercion of "0" to number 0 doesn't silently drop
        // tabs at the left margin (the previous `if (tabObj['@_w:pos'])` truthy
        // check turned pos=0 into an invisible tab-loss bug).
        if (tabObj['@_w:pos'] !== undefined) {
          const parsed = parseInt(String(tabObj['@_w:pos']), 10);
          if (!isNaN(parsed)) tab.position = parsed;
        }
        if (tabObj['@_w:val']) tab.val = tabObj['@_w:val'];
        if (tabObj['@_w:leader']) tab.leader = tabObj['@_w:leader'];

        if (tab.position !== undefined) {
          tabs.push(tab);
        }
      }

      if (tabs.length > 0) {
        paragraph.setTabs(tabs);
      }
    }

    // Widow control per ECMA-376 Part 1 §17.3.1.40
    if (pPrObj['w:widowControl'] !== undefined) {
      // Delegate to parseOoxmlBoolean so every ST_OnOff literal — including
      // "off" / "on" — resolves correctly. The previous bespoke check missed
      // "off", silently flipping explicit-off to explicit-on.
      paragraph.setWidowControl(parseOoxmlBoolean(pPrObj['w:widowControl']));
    }

    // Outline level per ECMA-376 Part 1 §17.3.1.19
    if (pPrObj['w:outlineLvl'] !== undefined && pPrObj['w:outlineLvl']['@_w:val'] !== undefined) {
      const level = parseInt(pPrObj['w:outlineLvl']['@_w:val'], 10);
      if (!isNaN(level) && level >= 0 && level <= 9) {
        paragraph.setOutlineLevel(level);
      }
    }

    // Suppress line numbers per ECMA-376 Part 1 §17.3.1.34
    if (pPrObj['w:suppressLineNumbers'] !== undefined) {
      paragraph.setSuppressLineNumbers(parseOoxmlBoolean(pPrObj['w:suppressLineNumbers']));
    }

    // Bidirectional layout per ECMA-376 Part 1 §17.3.1.6 — delegate to
    // parseOoxmlBoolean so "off"/"on" literals resolve correctly (the
    // previous bespoke check missed them).
    if (pPrObj['w:bidi'] !== undefined) {
      paragraph.setBidi(parseOoxmlBoolean(pPrObj['w:bidi']));
    }

    // Text direction per ECMA-376 Part 1 §17.3.1.36
    if (pPrObj['w:textDirection']?.['@_w:val']) {
      paragraph.setTextDirection(pPrObj['w:textDirection']['@_w:val']);
    }

    // Text vertical alignment per ECMA-376 Part 1 §17.3.1.35
    if (pPrObj['w:textAlignment']?.['@_w:val']) {
      paragraph.setTextAlignment(pPrObj['w:textAlignment']['@_w:val']);
    }

    // Mirror indents per ECMA-376 Part 1 §17.3.1.18
    if (pPrObj['w:mirrorIndents'] !== undefined) {
      paragraph.setMirrorIndents(parseOoxmlBoolean(pPrObj['w:mirrorIndents']));
    }

    // Auto-adjust right indent per ECMA-376 Part 1 §17.3.1.1 — delegate to
    // parseOoxmlBoolean so "off"/"on" literals resolve correctly.
    if (pPrObj['w:adjustRightInd'] !== undefined) {
      paragraph.setAdjustRightInd(parseOoxmlBoolean(pPrObj['w:adjustRightInd']));
    }

    // Text frame properties per ECMA-376 Part 1 §17.3.1.11
    if (pPrObj['w:framePr']) {
      const framePr = pPrObj['w:framePr'];
      const frameProps: any = {};
      // Use isExplicitlySet for numeric values to properly handle 0
      if (isExplicitlySet(framePr['@_w:w'])) frameProps.w = safeParseInt(framePr['@_w:w']);
      if (isExplicitlySet(framePr['@_w:h'])) frameProps.h = safeParseInt(framePr['@_w:h']);
      if (framePr['@_w:hRule']) frameProps.hRule = framePr['@_w:hRule'];
      if (isExplicitlySet(framePr['@_w:x'])) frameProps.x = safeParseInt(framePr['@_w:x']);
      if (isExplicitlySet(framePr['@_w:y'])) frameProps.y = safeParseInt(framePr['@_w:y']);
      if (framePr['@_w:xAlign']) frameProps.xAlign = framePr['@_w:xAlign'];
      if (framePr['@_w:yAlign']) frameProps.yAlign = framePr['@_w:yAlign'];
      if (framePr['@_w:hAnchor']) frameProps.hAnchor = framePr['@_w:hAnchor'];
      if (framePr['@_w:vAnchor']) frameProps.vAnchor = framePr['@_w:vAnchor'];
      if (isExplicitlySet(framePr['@_w:hSpace']))
        frameProps.hSpace = safeParseInt(framePr['@_w:hSpace']);
      if (isExplicitlySet(framePr['@_w:vSpace']))
        frameProps.vSpace = safeParseInt(framePr['@_w:vSpace']);
      if (framePr['@_w:wrap']) frameProps.wrap = framePr['@_w:wrap'];
      if (framePr['@_w:dropCap']) frameProps.dropCap = framePr['@_w:dropCap'];
      if (isExplicitlySet(framePr['@_w:lines']))
        frameProps.lines = safeParseInt(framePr['@_w:lines']);
      if (isExplicitlySet(framePr['@_w:anchorLock'])) {
        // Use parseOoxmlBoolean-style check for attribute value
        const val = framePr['@_w:anchorLock'];
        frameProps.anchorLock =
          val === '1' || val === 1 || val === 'true' || val === true || val === 'on';
      }
      if (Object.keys(frameProps).length > 0) {
        paragraph.setFrameProperties(frameProps);
      }
    }

    // Suppress automatic hyphenation per ECMA-376 Part 1 §17.3.1.33
    if (pPrObj['w:suppressAutoHyphens'] !== undefined) {
      paragraph.setSuppressAutoHyphens(parseOoxmlBoolean(pPrObj['w:suppressAutoHyphens']));
    }

    // CJK paragraph properties per ECMA-376 Part 1
    if (pPrObj['w:kinsoku']) {
      paragraph.setKinsoku(parseOoxmlBoolean(pPrObj['w:kinsoku']));
    }
    if (pPrObj['w:wordWrap']) {
      paragraph.setWordWrap(parseOoxmlBoolean(pPrObj['w:wordWrap']));
    }
    if (pPrObj['w:overflowPunct']) {
      paragraph.setOverflowPunct(parseOoxmlBoolean(pPrObj['w:overflowPunct']));
    }
    if (pPrObj['w:topLinePunct']) {
      paragraph.setTopLinePunct(parseOoxmlBoolean(pPrObj['w:topLinePunct']));
    }
    if (pPrObj['w:autoSpaceDE']) {
      paragraph.setAutoSpaceDE(parseOoxmlBoolean(pPrObj['w:autoSpaceDE']));
    }
    if (pPrObj['w:autoSpaceDN']) {
      paragraph.setAutoSpaceDN(parseOoxmlBoolean(pPrObj['w:autoSpaceDN']));
    }

    // Suppress text frame overlap per ECMA-376 Part 1 §17.3.1.34
    if (pPrObj['w:suppressOverlap'] !== undefined) {
      paragraph.setSuppressOverlap(parseOoxmlBoolean(pPrObj['w:suppressOverlap']));
    }

    // Textbox tight wrap per ECMA-376 Part 1 §17.3.1.37
    if (pPrObj['w:textboxTightWrap']) {
      const wrapVal = pPrObj['w:textboxTightWrap']?.['@_w:val'];
      if (wrapVal) {
        paragraph.setTextboxTightWrap(wrapVal);
      }
    }

    // HTML div ID per ECMA-376 Part 1 §17.3.1.10 (CT_DivId). `w:val` is
    // ST_DecimalNumber — 0 is a valid ID referencing the first div in
    // web settings. XMLParser coerces `"0"` to the number 0, and the
    // previous `if (divIdVal)` truthy check silently dropped it, breaking
    // the paragraph's link to div index 0 on every round-trip.
    if (pPrObj['w:divId']) {
      const divIdVal = pPrObj['w:divId']?.['@_w:val'];
      if (isExplicitlySet(divIdVal)) {
        const parsed = safeParseInt(divIdVal);
        if (!isNaN(parsed)) paragraph.setDivId(parsed);
      }
    }

    // Conditional table style formatting per ECMA-376 Part 1 §17.3.1.8
    if (pPrObj['w:cnfStyle']) {
      const cnfStyleVal = pPrObj['w:cnfStyle']?.['@_w:val'];
      if (cnfStyleVal !== undefined) {
        // Ensure it's a string and pad to 12 characters (standard bitmask length)
        // XML parser may convert to number, removing leading zeros
        const bitmask = String(cnfStyleVal).padStart(12, '0');
        paragraph.setConditionalFormatting(bitmask);
      }
    }

    // Paragraph property change tracking per ECMA-376 Part 1 §17.3.1.27.
    // CT_TrackChange attributes — `w:id` (ST_DecimalNumber, required),
    // `w:author` (ST_String, required), `w:date` (ST_DateTime, optional).
    // XMLParser coerces `w:id="0"` to the number 0; the previous
    // `if (changeObj['@_w:id'])` truthy gate silently dropped id=0,
    // producing `<w:pPrChange w:author="…" w:date="…"/>` on emission —
    // missing the required `w:id` and failing strict validation. The
    // sibling `trPrChange` / `tblPrChange` / `tcPrChange` / `sectPrChange`
    // parsers already use `|| '0'` or `!== undefined` for the same reason.
    if (pPrObj['w:pPrChange']) {
      const changeObj = pPrObj['w:pPrChange'];
      const change: any = {};
      if (changeObj['@_w:author'] !== undefined) {
        change.author = String(changeObj['@_w:author']);
      }
      if (changeObj['@_w:date'] !== undefined) {
        change.date = String(changeObj['@_w:date']);
      }
      if (changeObj['@_w:id'] !== undefined) {
        change.id = String(changeObj['@_w:id']);
      }

      // Parse child w:pPr for previousProperties to preserve tracked change history
      if (changeObj['w:pPr']) {
        const prevPPr = changeObj['w:pPr'];
        const previousProperties: any = {};

        // Parse previous style
        if (prevPPr['w:pStyle']?.['@_w:val']) {
          previousProperties.style = String(prevPPr['w:pStyle']['@_w:val']);
        }

        // Parse previous numbering
        if (prevPPr['w:numPr']) {
          const numPr = prevPPr['w:numPr'];
          previousProperties.numbering = {};
          if (numPr['w:ilvl']?.['@_w:val'] !== undefined) {
            const parsedLevel = parseInt(numPr['w:ilvl']['@_w:val'], 10);
            if (!isNaN(parsedLevel)) {
              previousProperties.numbering.level = parsedLevel;
            }
          }
          if (numPr['w:numId']?.['@_w:val'] !== undefined) {
            const parsedNumId = parseInt(numPr['w:numId']['@_w:val'], 10);
            if (!isNaN(parsedNumId)) {
              previousProperties.numbering.numId = parsedNumId;
            }
          }
        }

        // Parse previous indentation
        // Per ECMA-376 §17.3.1.15: w:start/w:end are bidi-aware alternatives to w:left/w:right.
        // Also parse the six CJK character-unit variants (ST_DecimalNumber) per §17.3.1.12;
        // these round-trip alongside the twips so Word's rendering of the tracked "previous"
        // state stays locale-accurate for CJK-authored documents. Matches the iteration-21
        // fix on the main-path parser.
        if (prevPPr['w:ind']) {
          const ind = prevPPr['w:ind'];
          previousProperties.indentation = {};
          const leftVal = ind['@_w:start'] ?? ind['@_w:left'];
          const rightVal = ind['@_w:end'] ?? ind['@_w:right'];
          if (leftVal !== undefined) previousProperties.indentation.left = parseInt(leftVal, 10);
          if (rightVal !== undefined) previousProperties.indentation.right = parseInt(rightVal, 10);
          if (ind['@_w:firstLine'] !== undefined)
            previousProperties.indentation.firstLine = parseInt(ind['@_w:firstLine'], 10);
          if (ind['@_w:hanging'] !== undefined)
            previousProperties.indentation.hanging = parseInt(ind['@_w:hanging'], 10);
          // CJK character-unit variants. startChars/endChars collapse onto
          // leftChars/rightChars (same pattern as the twips variants).
          const leftCharsVal = ind['@_w:startChars'] ?? ind['@_w:leftChars'];
          const rightCharsVal = ind['@_w:endChars'] ?? ind['@_w:rightChars'];
          if (leftCharsVal !== undefined)
            previousProperties.indentation.leftChars = parseInt(leftCharsVal, 10);
          if (rightCharsVal !== undefined)
            previousProperties.indentation.rightChars = parseInt(rightCharsVal, 10);
          if (ind['@_w:firstLineChars'] !== undefined)
            previousProperties.indentation.firstLineChars = parseInt(ind['@_w:firstLineChars'], 10);
          if (ind['@_w:hangingChars'] !== undefined)
            previousProperties.indentation.hangingChars = parseInt(ind['@_w:hangingChars'], 10);
        }

        // Parse previous alignment
        if (prevPPr['w:jc']?.['@_w:val']) {
          previousProperties.alignment = String(prevPPr['w:jc']['@_w:val']);
        }

        // Parse previous spacing (all 8 CT_Spacing attributes per ECMA-376 §17.3.1.33)
        if (prevPPr['w:spacing']) {
          const spacing = prevPPr['w:spacing'];
          previousProperties.spacing = {};
          if (spacing['@_w:before'] !== undefined)
            previousProperties.spacing.before = parseInt(spacing['@_w:before'], 10);
          if (spacing['@_w:after'] !== undefined)
            previousProperties.spacing.after = parseInt(spacing['@_w:after'], 10);
          if (spacing['@_w:line'] !== undefined)
            previousProperties.spacing.line = parseInt(spacing['@_w:line'], 10);
          if (spacing['@_w:lineRule'])
            previousProperties.spacing.lineRule = String(spacing['@_w:lineRule']);
          if (spacing['@_w:beforeLines'] !== undefined)
            previousProperties.spacing.beforeLines = parseInt(spacing['@_w:beforeLines'], 10);
          if (spacing['@_w:afterLines'] !== undefined)
            previousProperties.spacing.afterLines = parseInt(spacing['@_w:afterLines'], 10);
          // ST_OnOff per ECMA-376 §17.17.4 — accept 1/0/true/false/on/off
          const beforeAuto = spacing['@_w:beforeAutospacing'];
          if (beforeAuto !== undefined)
            previousProperties.spacing.beforeAutospacing = parseOnOffAttribute(beforeAuto);
          const afterAuto = spacing['@_w:afterAutospacing'];
          if (afterAuto !== undefined)
            previousProperties.spacing.afterAutospacing = parseOnOffAttribute(afterAuto);
        }

        // CT_OnOff properties per ECMA-376 §17.17.4 — accept "1"/"0"/"true"/"false"/"on"/"off"
        // plus the number forms produced by fast-xml-parser's parseAttributeValue. Using
        // parseOoxmlBoolean() keeps pPrChange round-trips consistent with the main pPr parser;
        // the previous `!== '0'` pattern silently flipped "false", "off", and the numeric 0.
        if (prevPPr['w:keepNext']) {
          previousProperties.keepNext = parseOoxmlBoolean(prevPPr['w:keepNext']);
        }
        if (prevPPr['w:keepLines']) {
          previousProperties.keepLines = parseOoxmlBoolean(prevPPr['w:keepLines']);
        }
        if (prevPPr['w:pageBreakBefore']) {
          previousProperties.pageBreakBefore = parseOoxmlBoolean(prevPPr['w:pageBreakBefore']);
        }

        // === Extended paragraph property parsing per ECMA-376 Part 1 §17.3.1 ===

        // Parse widowControl (w:widowControl) - orphan/widow control
        if (prevPPr['w:widowControl']) {
          previousProperties.widowControl = parseOoxmlBoolean(prevPPr['w:widowControl']);
        }

        // Parse suppressAutoHyphens (w:suppressAutoHyphens)
        if (prevPPr['w:suppressAutoHyphens']) {
          previousProperties.suppressAutoHyphens = parseOoxmlBoolean(
            prevPPr['w:suppressAutoHyphens']
          );
        }

        // Parse contextualSpacing (w:contextualSpacing)
        if (prevPPr['w:contextualSpacing']) {
          previousProperties.contextualSpacing = parseOoxmlBoolean(prevPPr['w:contextualSpacing']);
        }

        // Parse mirrorIndents (w:mirrorIndents)
        if (prevPPr['w:mirrorIndents']) {
          previousProperties.mirrorIndents = parseOoxmlBoolean(prevPPr['w:mirrorIndents']);
        }

        // Parse outlineLevel (w:outlineLvl @w:val)
        if (prevPPr['w:outlineLvl']?.['@_w:val'] !== undefined) {
          previousProperties.outlineLevel = parseInt(prevPPr['w:outlineLvl']['@_w:val'], 10);
        }

        // Parse previous text frame properties (w:framePr) per ECMA-376
        // Part 1 §17.3.1.11 CT_FramePr. The pPrChange emitter already
        // rebuilds every framePr attribute (see Paragraph.ts §3634), but
        // the parser never read them — so a tracked change to any
        // frame property (drop-cap, text-box positioning, wrap mode,
        // anchor lock…) silently lost the previous state on round-trip.
        if (prevPPr['w:framePr']) {
          const framePr = prevPPr['w:framePr'];
          const frameProps: any = {};
          if (isExplicitlySet(framePr['@_w:w'])) frameProps.w = safeParseInt(framePr['@_w:w']);
          if (isExplicitlySet(framePr['@_w:h'])) frameProps.h = safeParseInt(framePr['@_w:h']);
          if (framePr['@_w:hRule']) frameProps.hRule = String(framePr['@_w:hRule']);
          if (isExplicitlySet(framePr['@_w:x'])) frameProps.x = safeParseInt(framePr['@_w:x']);
          if (isExplicitlySet(framePr['@_w:y'])) frameProps.y = safeParseInt(framePr['@_w:y']);
          if (framePr['@_w:xAlign']) frameProps.xAlign = String(framePr['@_w:xAlign']);
          if (framePr['@_w:yAlign']) frameProps.yAlign = String(framePr['@_w:yAlign']);
          if (framePr['@_w:hAnchor']) frameProps.hAnchor = String(framePr['@_w:hAnchor']);
          if (framePr['@_w:vAnchor']) frameProps.vAnchor = String(framePr['@_w:vAnchor']);
          if (isExplicitlySet(framePr['@_w:hSpace'])) {
            frameProps.hSpace = safeParseInt(framePr['@_w:hSpace']);
          }
          if (isExplicitlySet(framePr['@_w:vSpace'])) {
            frameProps.vSpace = safeParseInt(framePr['@_w:vSpace']);
          }
          if (framePr['@_w:wrap']) frameProps.wrap = String(framePr['@_w:wrap']);
          if (framePr['@_w:dropCap']) frameProps.dropCap = String(framePr['@_w:dropCap']);
          if (isExplicitlySet(framePr['@_w:lines'])) {
            frameProps.lines = safeParseInt(framePr['@_w:lines']);
          }
          if (isExplicitlySet(framePr['@_w:anchorLock'])) {
            frameProps.anchorLock = parseOnOffAttribute(String(framePr['@_w:anchorLock']), true);
          }
          if (Object.keys(frameProps).length > 0) {
            previousProperties.framePr = frameProps;
          }
        }

        // Parse bidi (w:bidi) - right-to-left paragraph
        if (prevPPr['w:bidi']) {
          previousProperties.bidi = parseOoxmlBoolean(prevPPr['w:bidi']);
        }

        // Parse suppressLineNumbers (w:suppressLineNumbers)
        if (prevPPr['w:suppressLineNumbers']) {
          previousProperties.suppressLineNumbers = parseOoxmlBoolean(
            prevPPr['w:suppressLineNumbers']
          );
        }

        // Parse adjustRightInd (w:adjustRightInd)
        if (prevPPr['w:adjustRightInd']) {
          previousProperties.adjustRightInd = parseOoxmlBoolean(prevPPr['w:adjustRightInd']);
        }

        // Parse snapToGrid (w:snapToGrid)
        if (prevPPr['w:snapToGrid']) {
          previousProperties.snapToGrid = parseOoxmlBoolean(prevPPr['w:snapToGrid']);
        }

        // Parse wordWrap (w:wordWrap)
        if (prevPPr['w:wordWrap']) {
          previousProperties.wordWrap = parseOoxmlBoolean(prevPPr['w:wordWrap']);
        }

        // Parse autoSpaceDE (w:autoSpaceDE) - East Asian/numeric spacing
        if (prevPPr['w:autoSpaceDE']) {
          previousProperties.autoSpaceDE = parseOoxmlBoolean(prevPPr['w:autoSpaceDE']);
        }

        // Parse autoSpaceDN (w:autoSpaceDN) - East Asian/Western spacing
        if (prevPPr['w:autoSpaceDN']) {
          previousProperties.autoSpaceDN = parseOoxmlBoolean(prevPPr['w:autoSpaceDN']);
        }

        // Parse kinsoku / overflowPunct / topLinePunct / suppressOverlap —
        // CJK typography CT_OnOff flags. The Paragraph pPrChange generator
        // already emits these in the previous-properties block, but the
        // parser was missing the read side, so tracked paragraph-property
        // revisions that recorded any of these four flags were silently
        // dropped on load → save. Uses `parseOoxmlBoolean` to honour every
        // ST_OnOff literal (bare, 1/0, true/false, on/off).
        if (prevPPr['w:kinsoku']) {
          (previousProperties as { kinsoku?: boolean }).kinsoku = parseOoxmlBoolean(
            prevPPr['w:kinsoku']
          );
        }
        if (prevPPr['w:overflowPunct']) {
          (previousProperties as { overflowPunct?: boolean }).overflowPunct = parseOoxmlBoolean(
            prevPPr['w:overflowPunct']
          );
        }
        if (prevPPr['w:topLinePunct']) {
          (previousProperties as { topLinePunct?: boolean }).topLinePunct = parseOoxmlBoolean(
            prevPPr['w:topLinePunct']
          );
        }
        if (prevPPr['w:suppressOverlap']) {
          (previousProperties as { suppressOverlap?: boolean }).suppressOverlap = parseOoxmlBoolean(
            prevPPr['w:suppressOverlap']
          );
        }

        // Parse textDirection (w:textDirection @w:val)
        if (prevPPr['w:textDirection']?.['@_w:val']) {
          previousProperties.textDirection = String(prevPPr['w:textDirection']['@_w:val']);
        }

        // Parse textAlignment (w:textAlignment @w:val) per ECMA-376 Part 1 §17.3.1.39
        if (prevPPr['w:textAlignment']?.['@_w:val']) {
          previousProperties.textAlignment = String(prevPPr['w:textAlignment']['@_w:val']);
        }

        // Parse previous divId (w:divId) per ECMA-376 §17.3.1.10 —
        // ST_DecimalNumber referencing a web-settings div. Zero is a
        // legal ID (first div). XMLParser coerces `"0"` to number 0, so
        // gate via `isExplicitlySet` to preserve divId=0 on tracked
        // previous state. The pPrChange emitter (Paragraph.ts §3915)
        // re-emits prev.divId via `!== undefined`.
        if (prevPPr['w:divId']?.['@_w:val'] !== undefined) {
          const rawDivId = prevPPr['w:divId']['@_w:val'];
          const parsedDivId = safeParseInt(rawDivId);
          if (!isNaN(parsedDivId)) {
            previousProperties.divId = parsedDivId;
          }
        }

        // Parse previous cnfStyle (w:cnfStyle) per ECMA-376 §17.3.1.8 —
        // 12-character bitmask identifying which conditional-formatting
        // flags from the parent table style apply. XMLParser coerces
        // purely-numeric hex strings, but the custom parseValue keeps
        // 7+-digit strings as-is (so 12-char bitmasks survive); use
        // String + padStart to defensively normalise any shorter form.
        if (prevPPr['w:cnfStyle']?.['@_w:val'] !== undefined) {
          previousProperties.cnfStyle = String(prevPPr['w:cnfStyle']['@_w:val']).padStart(12, '0');
        }

        // Parse paragraph borders (w:pBdr) per ECMA-376 Part 1 §17.3.1.24
        // Previous versions stored the attribute values under the wrong
        // field names (`val`/`sz` instead of `style`/`size`) — the
        // paragraph emitter reads `style`/`size`, so every tracked
        // previous border collapsed to `<w:top w:val="nil"/>` on
        // round-trip. The CT_Border attribute coverage here now matches
        // the main parser (§17.18.2): all nine attrs, with shadow/frame
        // routed through parseOnOffAttribute so ST_OnOff literals
        // ("on"/"off"/"true"/"false") resolve correctly.
        if (prevPPr['w:pBdr']) {
          const pBdr = prevPPr['w:pBdr'];
          previousProperties.borders = {};

          const parseBorder = (borderObj: any) => {
            if (!borderObj) return undefined;
            const border: any = {};
            if (borderObj['@_w:val']) border.style = borderObj['@_w:val'];
            if (borderObj['@_w:sz'] !== undefined) border.size = safeParseInt(borderObj['@_w:sz']);
            if (borderObj['@_w:space'] !== undefined) {
              border.space = safeParseInt(borderObj['@_w:space']);
            }
            if (borderObj['@_w:color']) border.color = borderObj['@_w:color'];
            if (borderObj['@_w:themeColor']) {
              border.themeColor = String(borderObj['@_w:themeColor']);
            }
            if (borderObj['@_w:themeTint']) {
              border.themeTint = String(borderObj['@_w:themeTint']);
            }
            if (borderObj['@_w:themeShade']) {
              border.themeShade = String(borderObj['@_w:themeShade']);
            }
            if (borderObj['@_w:shadow'] !== undefined) {
              border.shadow = parseOnOffAttribute(String(borderObj['@_w:shadow']), true);
            }
            if (borderObj['@_w:frame'] !== undefined) {
              border.frame = parseOnOffAttribute(String(borderObj['@_w:frame']), true);
            }
            return Object.keys(border).length > 0 ? border : undefined;
          };

          if (pBdr['w:top']) previousProperties.borders.top = parseBorder(pBdr['w:top']);
          if (pBdr['w:bottom']) previousProperties.borders.bottom = parseBorder(pBdr['w:bottom']);
          if (pBdr['w:left']) previousProperties.borders.left = parseBorder(pBdr['w:left']);
          if (pBdr['w:right']) previousProperties.borders.right = parseBorder(pBdr['w:right']);
          if (pBdr['w:between'])
            previousProperties.borders.between = parseBorder(pBdr['w:between']);
          if (pBdr['w:bar']) previousProperties.borders.bar = parseBorder(pBdr['w:bar']);

          // Clean up empty borders object
          if (Object.keys(previousProperties.borders).length === 0) {
            delete previousProperties.borders;
          }
        }

        // Parse paragraph shading (w:shd) per ECMA-376 Part 1 §17.3.1.32
        if (prevPPr['w:shd']) {
          const shading = this.parseShadingFromObj(prevPPr['w:shd']);
          if (shading) {
            previousProperties.shading = shading;
          }
        }

        // Parse tab stops (w:tabs) per ECMA-376 Part 1 §17.3.1.38
        if (prevPPr['w:tabs']) {
          const tabsObj = prevPPr['w:tabs'];
          const tabArray = tabsObj['w:tab'];
          if (tabArray) {
            const tabs = Array.isArray(tabArray) ? tabArray : [tabArray];
            previousProperties.tabs = tabs.map((tab: any) => ({
              val: tab['@_w:val'],
              position: tab['@_w:pos'] !== undefined ? parseInt(tab['@_w:pos'], 10) : undefined,
              leader: tab['@_w:leader'],
            }));
          }
        }

        if (Object.keys(previousProperties).length > 0) {
          change.previousProperties = previousProperties;
        }
      }

      if (Object.keys(change).length > 0) {
        paragraph.setParagraphPropertiesChange(change);
      }
    }

    // Section properties per ECMA-376 Part 1 §17.3.1.30
    // Note: sectPr is set as raw XML string by callers that have access to the raw paragraph XML.
    // The parsed object is stored as a marker; callers override with raw XML for round-trip fidelity.
    if (pPrObj['w:sectPr']) {
      paragraph.setSectionProperties(pPrObj['w:sectPr']);
    }
  }

  /**
   * NEW: Assemble complex fields from run tokens
   * Groups begin→instr→sep→result→end sequences into ComplexField objects
   * Also handles Revision objects (w:ins, w:del) that appear in the result section of complex fields
   *
   * @param paragraph The paragraph containing runs to process
   */
  private assembleComplexFields(paragraph: Paragraph): void {
    // Skip if this paragraph is part of a multi-paragraph field (e.g., TOC)
    // These paragraphs have already been processed and their field structure
    // should be preserved - re-processing would corrupt the field runs
    if (paragraph._isPartOfMultiParagraphField) {
      defaultLogger.debug(
        '[DocumentParser] Skipping assembleComplexFields for paragraph marked as part of multi-paragraph field'
      );
      return;
    }

    const content = paragraph.getContent();
    const groupedContent: any[] = [];
    // Invariant: only Run instances — non-Run items filtered at entry (line ~3109)
    let fieldRuns: Run[] = [];
    let fieldRevisions: Revision[] = []; // Track revisions inside field result section
    let instructionRevisions: Revision[] = []; // Track revisions in instruction area
    let orderedResultItems: Array<Run | Revision> = []; // Track interleaved order of result runs + revisions
    let fieldState: 'begin' | 'instruction' | 'separate' | 'result' | 'end' | null = null;
    let nestingDepth = 0;
    let hasNestedFields = false;
    let fieldStartIndex = -1;

    for (let i = 0; i < content.length; i++) {
      const item = content[i];

      if (item instanceof Run) {
        const runContent = item.getContent();
        const hasFieldContent = runContent.some(
          (c: any) => c.type === 'fieldChar' || c.type === 'instructionText'
        );

        if (hasFieldContent) {
          // This run is part of a field
          fieldRuns.push(item);
          const fieldChar = runContent.find((c: any) => c.type === 'fieldChar');

          if (fieldChar) {
            switch (fieldChar.fieldCharType) {
              case 'begin':
                nestingDepth++;
                if (nestingDepth === 1) {
                  fieldState = 'begin';
                  fieldStartIndex = i;
                } else {
                  hasNestedFields = true;
                }
                break;
              case 'separate':
                if (nestingDepth <= 1) {
                  fieldState = 'separate';
                }
                break;
              case 'end':
                nestingDepth--;
                if (nestingDepth > 0) {
                  // Still inside a nested field - continue collecting
                  break;
                }
                fieldState = 'end';

                // Complete field assembly
                if (fieldRuns.length > 0) {
                  // NESTED FIELD: preserve raw structure to maintain field balance
                  // Nested fields (e.g. INCLUDEPICTURE multiplication from email-pasted images)
                  // have properly balanced begin/separate/end markers across nesting levels.
                  // Flattening them into a single ComplexField would collapse all instrTexts
                  // into one instruction and orphan the inner end markers, corrupting the document.
                  if (hasNestedFields) {
                    for (let j = fieldStartIndex; j <= i; j++) {
                      groupedContent.push(content[j]);
                    }
                    defaultLogger.debug(
                      `Preserved raw structure for nested field spanning ${i - fieldStartIndex + 1} content items`
                    );
                    fieldRuns = [];
                    instructionRevisions = [];
                    fieldRevisions = [];
                    orderedResultItems = [];
                    fieldState = null;
                    hasNestedFields = false;
                    fieldStartIndex = -1;
                    nestingDepth = 0;
                    break;
                  }
                  // If there are revisions in the instruction area, we MUST preserve raw structure
                  // The instruction text is INSIDE the revisions, so we can't extract it
                  // This happens with field code hyperlinks inside tracked changes (w:ins/w:del)
                  if (instructionRevisions.length > 0) {
                    // Output field structure in original order: runs interleaved with revisions
                    // We need to reconstruct the original order based on field state transitions
                    let hasSep = false;
                    for (const run of fieldRuns) {
                      if (!((run as unknown) instanceof Run)) {
                        defaultLogger.warn(
                          `Non-Run item in fieldRuns: ${(run as any)?.constructor?.name}`
                        );
                        continue;
                      }
                      const runContent = run.getContent();
                      const fieldCharToken = runContent.find((c: any) => c.type === 'fieldChar');

                      // Output begin run
                      if (fieldCharToken?.fieldCharType === 'begin') {
                        groupedContent.push(run);
                        // Instruction revisions come after begin
                        for (const rev of instructionRevisions) {
                          groupedContent.push(rev);
                        }
                      } else if (fieldCharToken?.fieldCharType === 'separate') {
                        // Separator
                        groupedContent.push(run);
                        hasSep = true;
                        // Result revisions come after separator
                        for (const rev of fieldRevisions) {
                          groupedContent.push(rev);
                        }
                      } else if (fieldCharToken?.fieldCharType === 'end') {
                        // End marker - if no separator was found, output revisions first
                        if (!hasSep) {
                          for (const rev of fieldRevisions) {
                            groupedContent.push(rev);
                          }
                        }
                        groupedContent.push(run);
                      } else {
                        // Other runs (instruction text, result text outside revisions)
                        groupedContent.push(run);
                      }
                    }

                    defaultLogger.debug(
                      `Preserved raw structure for field with ${instructionRevisions.length} instruction revisions`
                    );
                    fieldRuns = [];
                    instructionRevisions = [];
                    fieldRevisions = [];
                    orderedResultItems = [];
                    fieldState = null;
                    break;
                  }

                  // Extract instruction to determine field type
                  let instruction = '';
                  let resultText = '';
                  let resultFormatting: RunFormatting | undefined;
                  let hasSeparate = false;

                  for (const run of fieldRuns) {
                    if (!((run as unknown) instanceof Run)) {
                      defaultLogger.warn(
                        `Non-Run item in fieldRuns: ${(run as any)?.constructor?.name}`
                      );
                      continue;
                    }
                    const runContent = run.getContent();
                    const instrText = runContent.find((c: any) => c.type === 'instructionText');
                    if (instrText) {
                      instruction += instrText.value || '';
                    }
                    const fieldCharToken = runContent.find((c: any) => c.type === 'fieldChar');
                    if (fieldCharToken?.fieldCharType === 'separate') {
                      hasSeparate = true;
                    }
                    const textContent = runContent.find((c: any) => c.type === 'text');
                    if (textContent && hasSeparate) {
                      resultText += textContent.value || '';
                      resultFormatting = run.getFormatting();
                    }
                  }

                  instruction = instruction.trim();

                  // Check if this is a HYPERLINK field code - convert to Hyperlink element
                  if (isHyperlinkInstruction(instruction)) {
                    const parsed = parseHyperlinkInstruction(instruction);
                    if (parsed && resultText) {
                      // HYPERLINK fields with tracked changes in the result region cannot be
                      // converted to Hyperlink objects — the revisions would become orphaned
                      // standalone paragraph content. Keep as ComplexField instead.
                      if (fieldRevisions.length > 0) {
                        const complexField = this.createComplexFieldFromRuns(fieldRuns);
                        if (complexField) {
                          // Build ordered resultContent for correct round-trip serialization
                          if (orderedResultItems.length > 0) {
                            const orderedContent: XMLElement[] = [];
                            for (const resultItem of orderedResultItems) {
                              if (resultItem instanceof Run) {
                                orderedContent.push(resultItem.toXML());
                              } else if (resultItem instanceof Revision) {
                                const xml = resultItem.toXML();
                                if (xml) orderedContent.push(xml);
                              }
                            }
                            // Compute accepted text (non-revision runs + insertion text)
                            let acceptedText = '';
                            for (const resultItem of orderedResultItems) {
                              if (resultItem instanceof Run) {
                                acceptedText += resultItem.getText() || '';
                              } else if (
                                resultItem instanceof Revision &&
                                resultItem.getType() === 'insert'
                              ) {
                                for (const child of resultItem.getRuns()) {
                                  acceptedText += child.getText() || '';
                                }
                              }
                            }
                            // Set accepted text as the result (for getResult() API)
                            complexField.setResult(acceptedText || complexField.getResult() || '');
                            // Add ordered content for serialization (overrides the text-based path)
                            for (const xmlEl of orderedContent) {
                              complexField.addResultContent(xmlEl);
                            }
                          }
                          // Attach revisions for tracking/acceptance
                          for (const rev of fieldRevisions) {
                            rev.setFieldContext({
                              field: complexField,
                              instruction: complexField.getInstruction(),
                              position: 'result',
                            });
                          }
                          complexField.setResultRevisions(fieldRevisions);
                          groupedContent.push(complexField);
                        } else {
                          fieldRuns.forEach((run) => groupedContent.push(run));
                          for (const revision of fieldRevisions) {
                            groupedContent.push(revision);
                          }
                        }
                        fieldRuns = [];
                        instructionRevisions = [];
                        fieldRevisions = [];
                        orderedResultItems = [];
                        fieldState = null;
                        break;
                      }

                      // Normal path (no revisions) — convert to Hyperlink object
                      let hyperlink: Hyperlink;
                      if (parsed.url && parsed.url.trim() !== '') {
                        // External hyperlink (or combined external + anchor)
                        hyperlink = Hyperlink.createExternal(parsed.fullUrl, resultText, {
                          ...resultFormatting,
                          color: resultFormatting?.color || '0000FF',
                          underline: resultFormatting?.underline || 'single',
                        });
                      } else if (parsed.anchor) {
                        // Internal hyperlink (anchor only)
                        hyperlink = Hyperlink.createInternal(parsed.anchor, resultText, {
                          ...resultFormatting,
                          color: resultFormatting?.color || '0000FF',
                          underline: resultFormatting?.underline || 'single',
                        });
                      } else {
                        // Fallback to ComplexField if we can't determine link type
                        const complexField = this.createComplexFieldFromRuns(fieldRuns);
                        if (complexField) {
                          // Attach revisions to ComplexField for proper serialization
                          if (fieldRevisions.length > 0) {
                            // Update field context on each revision with field reference and instruction
                            for (const rev of fieldRevisions) {
                              rev.setFieldContext({
                                field: complexField,
                                instruction: complexField.getInstruction(),
                                position: 'result',
                              });
                            }
                            complexField.setResultRevisions(fieldRevisions);
                          }
                          groupedContent.push(complexField);
                        } else {
                          fieldRuns.forEach((run) => groupedContent.push(run));
                          // Only push revisions separately if ComplexField creation failed
                          for (const revision of fieldRevisions) {
                            groupedContent.push(revision);
                          }
                        }
                        fieldRuns = [];
                        instructionRevisions = [];
                        fieldRevisions = [];
                        orderedResultItems = [];
                        fieldState = null;
                        break;
                      }

                      // Set tooltip if present
                      if (parsed.tooltip) {
                        hyperlink.setTooltip(parsed.tooltip);
                      }

                      groupedContent.push(hyperlink);
                      for (const revision of fieldRevisions) {
                        groupedContent.push(revision);
                      }
                      defaultLogger.debug(
                        `Converted single-paragraph HYPERLINK field to Hyperlink element`
                      );
                      fieldRuns = [];
                      instructionRevisions = [];
                      fieldRevisions = [];
                      orderedResultItems = [];
                      fieldState = null;
                      break;
                    }
                  }

                  // Non-HYPERLINK field: create ComplexField as usual
                  const complexField = this.createComplexFieldFromRuns(fieldRuns);
                  if (complexField) {
                    // Attach any revisions that were inside the field result section
                    // These are tracked changes within the field's display content
                    // They MUST be serialized before the end marker per ECMA-376
                    if (fieldRevisions.length > 0) {
                      // Update field context on each revision with field reference and instruction
                      for (const rev of fieldRevisions) {
                        rev.setFieldContext({
                          field: complexField,
                          instruction: complexField.getInstruction(),
                          position: 'result',
                        });
                      }
                      complexField.setResultRevisions(fieldRevisions);
                    }
                    groupedContent.push(complexField);
                  } else {
                    // If assembly failed, add individual runs
                    fieldRuns.forEach((run) => groupedContent.push(run));
                    // Still add the revisions
                    for (const revision of fieldRevisions) {
                      groupedContent.push(revision);
                    }
                  }
                  fieldRuns = [];
                  instructionRevisions = [];
                  fieldRevisions = [];
                  orderedResultItems = [];
                  fieldState = null;
                }
                break;
            }
          } else {
            // Instruction text run
            if (fieldState === 'begin' || fieldState === 'instruction') {
              fieldState = 'instruction';
            }
          }
        } else {
          // Regular run - check if we're inside a field result section
          if (nestingDepth > 0) {
            // Inside a field - collect all runs to preserve raw structure
            fieldRuns.push(item);
            // Track result runs for interleaved ordering (only outer field level)
            if (nestingDepth === 1 && (fieldState === 'separate' || fieldState === 'result')) {
              orderedResultItems.push(item);
            }
          } else if (fieldState === 'separate' || fieldState === 'result') {
            // This run is part of the field result - collect it
            fieldRuns.push(item);
            orderedResultItems.push(item);
            fieldState = 'result';
          } else if (fieldState === 'begin' || fieldState === 'instruction') {
            // We're in the middle of parsing field instruction
            // Empty runs between begin and separate should be preserved as part of the field
            // (Word sometimes inserts empty formatting runs in field code sections)
            fieldRuns.push(item);
          } else if (fieldRuns.length > 0) {
            // Incomplete field - add as individual runs
            fieldRuns.forEach((run) => groupedContent.push(run));
            for (const revision of instructionRevisions) {
              groupedContent.push(revision);
            }
            for (const revision of fieldRevisions) {
              groupedContent.push(revision);
            }
            fieldRuns = [];
            instructionRevisions = [];
            fieldRevisions = [];
            orderedResultItems = [];
            fieldState = null;
            groupedContent.push(item);
          } else {
            groupedContent.push(item);
          }
        }
      } else if (item instanceof Revision) {
        // Handle Revision objects (w:ins, w:del) that may appear inside field sections
        if (nestingDepth > 1) {
          // Inside a nested field (depth 2+) - collect revision to preserve raw structure
          fieldRevisions.push(item);
        } else if (fieldState === 'separate' || fieldState === 'result') {
          // This revision is inside the field result section - track it
          // Set preliminary field context (field reference will be set when ComplexField is created)
          item.setFieldContext({ position: 'result' });
          fieldRevisions.push(item);
          orderedResultItems.push(item);
          fieldState = 'result';
          defaultLogger.debug(
            `Found revision inside complex field result: type=${item.getType()}, id=${item.getId()}`
          );
        } else if (fieldState === 'begin' || fieldState === 'instruction') {
          // Revision in instruction area - track separately and continue field assembly
          // This happens with field code hyperlinks inside tracked changes
          // The instruction text is INSIDE the revision, so we can't extract it into ComplexField
          // Set preliminary field context (field reference will be set when ComplexField is created)
          item.setFieldContext({ position: 'instruction' });
          instructionRevisions.push(item);
          fieldState = 'instruction';
          defaultLogger.debug(
            `Found revision inside complex field instruction: type=${item.getType()}, id=${item.getId()}`
          );
        } else if (fieldRuns.length > 0) {
          // Revision appears in unexpected location during field assembly
          // Add incomplete field runs and the revision
          fieldRuns.forEach((run) => groupedContent.push(run));
          for (const revision of instructionRevisions) {
            groupedContent.push(revision);
          }
          for (const revision of fieldRevisions) {
            groupedContent.push(revision);
          }
          fieldRuns = [];
          instructionRevisions = [];
          fieldRevisions = [];
          orderedResultItems = [];
          fieldState = null;
          groupedContent.push(item);
        } else {
          // Normal revision outside of any field — check if it contains a complete
          // field code sequence (Scenario A: entire HYPERLINK fldChar sequence inside w:ins).
          // parseRevisionFromXml() creates individual Runs from the fldChar/instrText/result
          // elements, but assembleComplexFields() only sees the Revision wrapper.
          // Promote complete field code sequences to ComplexField at the paragraph level.
          const promoted = this.tryPromoteRevisionFieldCode(item);
          if (promoted) {
            groupedContent.push(promoted);
          } else {
            groupedContent.push(item);
          }
        }
      } else {
        // Non-run content (hyperlinks, images, etc.)
        if (nestingDepth > 0) {
          // Non-Run items (e.g., w:proofErr) can't be processed as field runs.
          // Drop them — Word regenerates these markers on open.
          defaultLogger.debug(
            `Dropping non-Run item inside field (depth=${nestingDepth}): ${(item as any)?.getElementType?.() || (item as any)?.constructor?.name}`
          );
          continue;
        } else if (fieldRuns.length > 0) {
          // Incomplete field - add as individual runs
          fieldRuns.forEach((run) => groupedContent.push(run));
          for (const revision of instructionRevisions) {
            groupedContent.push(revision);
          }
          for (const revision of fieldRevisions) {
            groupedContent.push(revision);
          }
          fieldRuns = [];
          instructionRevisions = [];
          fieldRevisions = [];
          orderedResultItems = [];
          fieldState = null;
        }
        groupedContent.push(item);
      }
    }

    // Handle any remaining incomplete field
    if (fieldRuns.length > 0) {
      fieldRuns.forEach((run) => groupedContent.push(run));
      for (const revision of instructionRevisions) {
        groupedContent.push(revision);
      }
      for (const revision of fieldRevisions) {
        groupedContent.push(revision);
      }
    }

    // Replace paragraph content with grouped content using setContent
    paragraph.setContent(groupedContent as ParagraphContent[]);

    defaultLogger.debug(
      `Assembled ${groupedContent.length - content.length} complex fields in paragraph`
    );
  }

  /**
   * Assembles complex fields that span multiple paragraphs (e.g., TOC fields)
   *
   * In Word documents, complex fields like TOC can have their begin/separate/end
   * markers distributed across multiple paragraphs. This method:
   * 1. Scans all paragraphs for field markers
   * 2. When a field spans multiple paragraphs, creates a ComplexField marked as multiParagraph
   * 3. Places the ComplexField in the first paragraph and removes field tokens from subsequent paragraphs
   *
   * For single-paragraph fields, delegates to assembleComplexFields for efficiency.
   *
   * @param bodyElements All parsed body elements (paragraphs, tables, SDTs)
   */
  private assembleMultiParagraphFields(bodyElements: BodyElement[]): void {
    // Extract all paragraphs (including those nested in tables and SDTs)
    const allParagraphs = this.collectAllParagraphs(bodyElements);

    if (allParagraphs.length === 0) {
      return;
    }

    // Track field state across paragraphs
    interface FieldTracker {
      startParagraphIndex: number;
      startRunIndex: number;
      fieldRuns: { paragraphIndex: number; runIndex: number; run: Run }[];
      hasBegin: boolean;
      hasSeparate: boolean;
      hasEnd: boolean;
    }

    let currentField: FieldTracker | null = null;
    const completedFields: FieldTracker[] = [];

    // First pass: identify multi-paragraph fields
    for (let pIdx = 0; pIdx < allParagraphs.length; pIdx++) {
      const paragraph = allParagraphs[pIdx]!;
      const content = paragraph.getContent();

      for (let rIdx = 0; rIdx < content.length; rIdx++) {
        const item = content[rIdx];

        if (!(item instanceof Run)) {
          continue;
        }

        const runContent = item.getContent();
        const fieldChar = runContent.find((c: any) => c.type === 'fieldChar');

        if (fieldChar) {
          switch (fieldChar.fieldCharType) {
            case 'begin':
              // Start tracking a new field
              currentField = {
                startParagraphIndex: pIdx,
                startRunIndex: rIdx,
                fieldRuns: [{ paragraphIndex: pIdx, runIndex: rIdx, run: item }],
                hasBegin: true,
                hasSeparate: false,
                hasEnd: false,
              };
              break;

            case 'separate':
              if (currentField) {
                currentField.hasSeparate = true;
                currentField.fieldRuns.push({ paragraphIndex: pIdx, runIndex: rIdx, run: item });
              }
              break;

            case 'end':
              if (currentField) {
                currentField.hasEnd = true;
                currentField.fieldRuns.push({ paragraphIndex: pIdx, runIndex: rIdx, run: item });

                // Check if this field spans multiple paragraphs
                const startPIdx = currentField.startParagraphIndex;
                const endPIdx = pIdx;

                if (endPIdx > startPIdx) {
                  // Multi-paragraph field - add to completed list for processing
                  completedFields.push(currentField);
                }
                // Single-paragraph fields will be handled by assembleComplexFields
                currentField = null;
              }
              break;
          }
        } else {
          // Check for instruction text or result text (part of field)
          const hasFieldContent = runContent.some((c: any) => c.type === 'instructionText');
          if (hasFieldContent && currentField) {
            currentField.fieldRuns.push({ paragraphIndex: pIdx, runIndex: rIdx, run: item });
          } else if (currentField && runContent.some((c: any) => c.type === 'text')) {
            // Text between separate and end is result text
            if (currentField.hasSeparate && !currentField.hasEnd) {
              currentField.fieldRuns.push({ paragraphIndex: pIdx, runIndex: rIdx, run: item });
            }
          }
        }
      }
    }

    // Process completed multi-paragraph fields
    for (const fieldTracker of completedFields) {
      this.processMultiParagraphField(fieldTracker, allParagraphs);
    }

    // Now handle single-paragraph fields in remaining paragraphs
    for (const paragraph of allParagraphs) {
      this.assembleComplexFields(paragraph);
    }

    const multiFieldCount = completedFields.length;
    if (multiFieldCount > 0) {
      defaultLogger.info(`Assembled ${multiFieldCount} multi-paragraph complex field(s)`);
    }
  }

  /**
   * Collects all paragraphs from body elements, including those nested in tables and SDTs
   */
  private collectAllParagraphs(bodyElements: BodyElement[]): Paragraph[] {
    const paragraphs: Paragraph[] = [];

    for (const element of bodyElements) {
      if (element instanceof Paragraph) {
        paragraphs.push(element);
      } else if (element instanceof Table) {
        // Extract paragraphs from table cells
        for (const row of element.getRows()) {
          for (const cell of row.getCells()) {
            paragraphs.push(...cell.getParagraphs());
          }
        }
      } else if (element instanceof StructuredDocumentTag) {
        // SDT content may contain paragraphs
        const sdtContent = element.getContent();
        for (const item of sdtContent) {
          if (item instanceof Paragraph) {
            paragraphs.push(item);
          }
        }
      }
    }

    return paragraphs;
  }

  /**
   * Processes a multi-paragraph field by:
   * 1. Creating a ComplexField from the collected runs
   * 2. Placing it in the first paragraph
   * 3. Removing field-related runs from subsequent paragraphs
   */
  private processMultiParagraphField(
    fieldTracker: {
      startParagraphIndex: number;
      fieldRuns: { paragraphIndex: number; runIndex: number; run: Run }[];
    },
    allParagraphs: Paragraph[]
  ): void {
    const runs = fieldTracker.fieldRuns.map((fr) => fr.run);

    // Extract field instruction and result text to determine field type
    let instruction = '';
    let resultText = '';
    let resultFormatting: RunFormatting | undefined;
    let hasSeparate = false;
    let resultParagraphIndex: number | undefined;

    for (let i = 0; i < fieldTracker.fieldRuns.length; i++) {
      const fr = fieldTracker.fieldRuns[i]!;
      const runContent = fr.run.getContent();

      // Check for fieldChar tokens
      const fieldCharToken = runContent.find((c: any) => c.type === 'fieldChar');
      if (fieldCharToken) {
        if (fieldCharToken.fieldCharType === 'separate') {
          hasSeparate = true;
        }
      }

      // Check for instruction text
      const instrText = runContent.find((c: any) => c.type === 'instructionText');
      if (instrText) {
        instruction += instrText.value || '';
      }

      // Check for result text (between separate and end)
      const textContent = runContent.find((c: any) => c.type === 'text');
      if (textContent && hasSeparate) {
        resultText += textContent.value || '';
        resultFormatting = fr.run.getFormatting();
        resultParagraphIndex = fr.paragraphIndex;
      }
    }

    instruction = instruction.trim();

    // Check if this is a HYPERLINK field code - convert to Hyperlink element
    if (isHyperlinkInstruction(instruction)) {
      const parsed = parseHyperlinkInstruction(instruction);
      if (parsed && resultText) {
        // Determine target paragraph: where result text resides (not first paragraph)
        const targetParagraphIndex = resultParagraphIndex ?? fieldTracker.startParagraphIndex;
        const targetParagraph = allParagraphs[targetParagraphIndex];

        if (targetParagraph) {
          // Create Hyperlink element with proper styling
          // If there's a URL, it's external; if only anchor, it's internal
          let hyperlink: Hyperlink;
          if (parsed.url && parsed.url.trim() !== '') {
            // External hyperlink (or combined external + anchor)
            hyperlink = Hyperlink.createExternal(parsed.fullUrl, resultText, {
              ...resultFormatting,
              // Ensure hyperlink styling if not already present
              color: resultFormatting?.color || '0000FF',
              underline: resultFormatting?.underline || 'single',
            });
          } else if (parsed.anchor) {
            // Internal hyperlink (anchor only)
            hyperlink = Hyperlink.createInternal(parsed.anchor, resultText, {
              ...resultFormatting,
              color: resultFormatting?.color || '0000FF',
              underline: resultFormatting?.underline || 'single',
            });
          } else {
            // Fallback: create as ComplexField if we can't determine the link type
            defaultLogger.debug(
              'HYPERLINK field missing URL and anchor, falling back to ComplexField'
            );
            this.processMultiParagraphFieldAsComplexField(fieldTracker, allParagraphs, runs);
            return;
          }

          // Set tooltip if present
          if (parsed.tooltip) {
            hyperlink.setTooltip(parsed.tooltip);
          }

          // Group field runs by paragraph index
          const runsByParagraph = new Map<number, Set<number>>();
          for (const fr of fieldTracker.fieldRuns) {
            if (!runsByParagraph.has(fr.paragraphIndex)) {
              runsByParagraph.set(fr.paragraphIndex, new Set());
            }
            runsByParagraph.get(fr.paragraphIndex)!.add(fr.runIndex);
          }

          // Process each affected paragraph
          const affectedParagraphIndices = Array.from(runsByParagraph.keys()).sort((a, b) => a - b);

          for (const pIdx of affectedParagraphIndices) {
            const paragraph = allParagraphs[pIdx]!;
            const runIndicesToRemove = runsByParagraph.get(pIdx)!;
            const content = paragraph.getContent();

            if (pIdx === targetParagraphIndex) {
              // Target paragraph: replace field runs with Hyperlink
              const newContent: ParagraphContent[] = [];
              let hyperlinkInserted = false;

              for (let rIdx = 0; rIdx < content.length; rIdx++) {
                if (runIndicesToRemove.has(rIdx)) {
                  // Insert Hyperlink at position of first field run in this paragraph
                  if (!hyperlinkInserted) {
                    newContent.push(hyperlink);
                    hyperlinkInserted = true;
                  }
                  // Skip this run (it's part of the field)
                } else {
                  newContent.push(content[rIdx]!);
                }
              }

              paragraph.setContent(newContent);
            } else {
              // Other paragraphs: remove field runs entirely
              const newContent = content.filter(
                (_: ParagraphContent, rIdx: number) => !runIndicesToRemove.has(rIdx)
              );
              paragraph.setContent(newContent);
            }
          }

          defaultLogger.debug(
            `Converted multi-paragraph HYPERLINK field to Hyperlink element in paragraph ${targetParagraphIndex}`
          );
          return;
        }
      }
    }

    // For non-HYPERLINK fields (or if HYPERLINK conversion failed), use standard ComplexField handling
    this.processMultiParagraphFieldAsComplexField(fieldTracker, allParagraphs, runs);
  }

  /**
   * Process a multi-paragraph field as a ComplexField (standard behavior for non-HYPERLINK fields)
   * For fields like TOC that span multiple paragraphs with content in between (begin → entries → end),
   * we preserve the original runs instead of creating a ComplexField.
   */
  private processMultiParagraphFieldAsComplexField(
    fieldTracker: {
      startParagraphIndex: number;
      fieldRuns: { paragraphIndex: number; runIndex: number; run: Run }[];
    },
    allParagraphs: Paragraph[],
    runs: Run[]
  ): void {
    // Mark all affected paragraphs so assembleComplexFields() skips them
    // This is critical for TOC and other multi-paragraph fields to preserve
    // the original field structure (fldChar begin/separate/end + instrText)
    const affectedParaIndices = new Set<number>();
    for (const fr of fieldTracker.fieldRuns) {
      affectedParaIndices.add(fr.paragraphIndex);
    }

    // Get first and last paragraph indices
    const sortedIndices = Array.from(affectedParaIndices).sort((a, b) => a - b);
    const firstParaIdx = sortedIndices[0]!;
    const lastParaIdx = sortedIndices[sortedIndices.length - 1]!;

    // For fields spanning more than 2 consecutive paragraphs (like TOC),
    // preserve the original runs instead of creating a ComplexField.
    // This is because the field content (e.g., TOC entries) is in the intermediate
    // paragraphs between begin/separate and end markers.
    if (lastParaIdx - firstParaIdx > 1) {
      // Mark ALL paragraphs in the range (including intermediate ones with field content)
      // For TOC: para 2 has begin/sep, paras 3-5 have TOC entries, para 6 has end
      for (let pIdx = firstParaIdx; pIdx <= lastParaIdx; pIdx++) {
        const paragraph = allParagraphs[pIdx];
        if (paragraph) {
          paragraph._isPartOfMultiParagraphField = true;
        }
      }
      defaultLogger.debug(
        `Preserving original runs for multi-paragraph field spanning paragraphs ${firstParaIdx} to ${lastParaIdx} (gap > 1)`
      );
      return;
    }

    // For fields spanning exactly 2 consecutive paragraphs, proceed with ComplexField
    for (const pIdx of affectedParaIndices) {
      const paragraph = allParagraphs[pIdx];
      if (paragraph) {
        paragraph._isPartOfMultiParagraphField = true;
      }
    }

    // Create the ComplexField
    const complexField = this.createComplexFieldFromRuns(runs);
    if (!complexField) {
      defaultLogger.debug('Failed to create ComplexField from multi-paragraph runs');
      return;
    }

    // Mark as multi-paragraph
    complexField.setMultiParagraph(true);

    // Group field runs by paragraph index
    const runsByParagraph = new Map<number, Set<number>>();
    for (const fr of fieldTracker.fieldRuns) {
      if (!runsByParagraph.has(fr.paragraphIndex)) {
        runsByParagraph.set(fr.paragraphIndex, new Set());
      }
      runsByParagraph.get(fr.paragraphIndex)!.add(fr.runIndex);
    }

    // Process each affected paragraph
    const affectedParagraphIndices = Array.from(runsByParagraph.keys()).sort((a, b) => a - b);

    for (let i = 0; i < affectedParagraphIndices.length; i++) {
      const pIdx = affectedParagraphIndices[i]!;
      const paragraph = allParagraphs[pIdx]!;
      const runIndicesToRemove = runsByParagraph.get(pIdx)!;
      const content = paragraph.getContent();

      if (i === 0) {
        // First paragraph: replace field runs with ComplexField
        const newContent: ParagraphContent[] = [];
        let fieldInserted = false;

        for (let rIdx = 0; rIdx < content.length; rIdx++) {
          if (runIndicesToRemove.has(rIdx)) {
            // Insert ComplexField at position of first field run
            if (!fieldInserted) {
              newContent.push(complexField);
              fieldInserted = true;
            }
            // Skip this run (it's part of the field)
          } else {
            newContent.push(content[rIdx]!);
          }
        }

        paragraph.setContent(newContent);
      } else {
        // Subsequent paragraphs: remove field runs entirely
        const newContent = content.filter(
          (_: ParagraphContent, rIdx: number) => !runIndicesToRemove.has(rIdx)
        );
        paragraph.setContent(newContent);
      }
    }

    defaultLogger.debug(
      `Processed multi-paragraph field spanning paragraphs ${fieldTracker.startParagraphIndex} to ${affectedParagraphIndices[affectedParagraphIndices.length - 1]}`
    );
  }

  /**
   * Checks if a Revision contains a complete field code sequence (fldChar begin → instrText →
   * fldChar separate → result text → fldChar end) and promotes it to a ComplexField.
   *
   * This handles the case where an entire HYPERLINK field code is inside a tracked change
   * (w:ins), producing Runs with fldChar tokens inside the Revision. Since assembleComplexFields()
   * only sees the Revision wrapper (not the Runs inside), the field code is never assembled.
   *
   * @returns ComplexField if promotion succeeded, null otherwise
   */
  private tryPromoteRevisionFieldCode(revision: Revision): ComplexField | null {
    const revType = revision.getType();
    // Do not promote fields inside tracked change revisions.
    // Keeping them as Revision objects preserves round-trip fidelity:
    // - For w:del/w:moveFrom: createDeletedRunXml() converts w:t→w:delText, w:instrText→w:delInstrText
    // - For w:ins/w:moveTo: normal Run.toXML() inside the revision wrapper
    if (
      revType === 'delete' ||
      revType === 'insert' ||
      revType === 'moveFrom' ||
      revType === 'moveTo'
    ) {
      return null;
    }

    // Non-content revision types (property changes, table cell ops) never contain field sequences
    return null;
  }

  /**
   * NEW: Create ComplexField from sequence of field runs
   * Extracts instruction from instrText runs and result from text runs
   *
   * @param fieldRuns Array of runs containing field tokens
   * @returns ComplexField or null if invalid sequence
   */
  private createComplexFieldFromRuns(fieldRuns: Run[]): ComplexField | null {
    if (fieldRuns.length < 2) {
      // LENIENT: Just log debug message, don't prevent document loading
      defaultLogger.debug(
        'Skipping ComplexField assembly: insufficient runs (minimum 2: begin and instr)'
      );
      return null;
    }

    let instruction = '';
    let resultText = '';
    let instructionFormatting: RunFormatting | undefined;
    let resultFormatting: RunFormatting | undefined;
    let hasBegin = false;
    let hasEnd = false;
    let hasSeparate = false;
    let formFieldData: any = undefined;

    for (const run of fieldRuns) {
      if (!((run as unknown) instanceof Run)) {
        defaultLogger.warn(`Non-Run item in fieldRuns: ${(run as any)?.constructor?.name}`);
        continue;
      }
      const runContent = run.getContent();

      // Check for fieldChar tokens
      const fieldCharToken = runContent.find((c: any) => c.type === 'fieldChar');
      if (fieldCharToken) {
        switch (fieldCharToken.fieldCharType) {
          case 'begin':
            hasBegin = true;
            // Capture formatting from begin run
            instructionFormatting = run.getFormatting();
            // Capture form field data from begin field char
            if (fieldCharToken.formFieldData) {
              formFieldData = fieldCharToken.formFieldData;
            }
            break;
          case 'separate':
            hasSeparate = true;
            break;
          case 'end':
            hasEnd = true;
            break;
        }
      }

      // Check for instruction text
      const instrText = runContent.find((c: any) => c.type === 'instructionText');
      if (instrText) {
        instruction += instrText.value || '';
      }

      // Check for result text (between separate and end)
      const textContent = runContent.find((c: any) => c.type === 'text');
      if (textContent && hasSeparate) {
        resultText += textContent.value || '';
        resultFormatting = run.getFormatting();
      }
    }

    // Second pass: collect non-text result content (e.g., ImageRuns with drawings)
    // These runs are in the result section (past separator) but don't contribute text.
    // Their XML representation must be stored as resultContent so it survives round-trip.
    const resultContentElements: XMLElement[] = [];
    let pastSeparator = false;
    for (const run of fieldRuns) {
      if (!((run as unknown) instanceof Run)) {
        defaultLogger.warn(`Non-Run item in fieldRuns: ${(run as any)?.constructor?.name}`);
        continue;
      }
      const rc = run.getContent();
      const fc = rc.find((c: any) => c.type === 'fieldChar');
      if (fc?.fieldCharType === 'separate') {
        pastSeparator = true;
        continue;
      }
      if (fc?.fieldCharType === 'end') break;
      if (fc) continue; // skip begin runs
      if (rc.some((c: any) => c.type === 'instructionText')) continue; // skip instrText runs
      if (pastSeparator) {
        // Check if this is a non-text result run (e.g., ImageRun with drawing content)
        const hasNonEmptyText = rc.some(
          (c: any) => c.type === 'text' && c.value && c.value.length > 0
        );
        if (!hasNonEmptyText && run instanceof ImageRun) {
          resultContentElements.push(run.toXML());
        }
      }
    }

    // Validate field structure with detailed diagnostics
    if (!hasBegin) {
      const instrPreview = instruction ? instruction.substring(0, 50) : '<none>';
      defaultLogger.warn(`ComplexField missing 'begin' marker. Instruction: "${instrPreview}..."`);
      this.parseErrors.push({
        element: 'complex-field-structure',
        error: new Error('Missing field begin marker'),
      });
      return null;
    }

    if (!hasEnd) {
      const instrPreview = instruction ? instruction.substring(0, 50) : '<none>';
      defaultLogger.warn(`ComplexField missing 'end' marker. Instruction: "${instrPreview}..."`);
      this.parseErrors.push({
        element: 'complex-field-structure',
        error: new Error('Missing field end marker'),
      });
      return null;
    }

    if (!instruction.trim()) {
      defaultLogger.warn(`ComplexField has no instruction content`);
      this.parseErrors.push({
        element: 'complex-field-structure',
        error: new Error('Empty field instruction'),
      });
      return null;
    }

    // Trim and clean instruction
    instruction = instruction.trim();

    defaultLogger.debug(
      `Created ComplexField: ${instruction.substring(0, 50)}... (result: "${resultText}")`
    );

    const properties: any = {
      instruction,
      result: resultText,
      resultContent: resultContentElements.length > 0 ? resultContentElements : undefined,
      instructionFormatting,
      resultFormatting,
      multiParagraph: false, // Default - can be set later if needed
      hasResult: hasSeparate, // Track if field had a separator/result section per ECMA-376
      formFieldData,
    };

    return new ComplexField(properties);
  }

  private parseRunFromObject(runObj: any): Run | null {
    try {
      // Extract all run content elements (text, tabs, breaks, etc.)
      // Per ECMA-376 §17.3.3 EG_RunInnerContent, runs can contain multiple content types
      const content: RunContent[] = [];

      const toArray = <T>(value: T | T[] | undefined | null): T[] =>
        Array.isArray(value) ? value : value !== undefined && value !== null ? [value] : [];

      const extractTextValue = (node: any): string => {
        if (node === undefined || node === null) {
          return '';
        }
        if (typeof node === 'object') {
          return XMLBuilder.unescapeXml(node['#text'] || '');
        }
        return XMLBuilder.unescapeXml(String(node));
      };

      // Field-character attributes (w:dirty, w:fldLock, w:lock on w:fldChar) are
      // ST_OnOff per ECMA-376 §17.16.18. Delegate to parseOnOffAttribute so every
      // literal is honoured — the previous inline check missed "on" (silently
      // coerced to false) and was tighter than the spec requires.
      const parseBooleanAttr = (value: unknown): boolean | undefined => {
        if (value === undefined || value === null) {
          return undefined;
        }
        return parseOnOffAttribute(value);
      };

      // Parse w:ffData from a fldChar object (form field data per ECMA-376 §17.16.17)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      type XmlNode = Record<string, any>;

      const parseFormFieldData = (fldCharObj: XmlNode): FormFieldData | undefined => {
        const ffDataObj: XmlNode | undefined = fldCharObj['w:ffData'];
        if (!ffDataObj || typeof ffDataObj !== 'object') return undefined;
        const ffd: FormFieldData = {};
        // w:name
        if (ffDataObj['w:name']?.['@_w:val'] !== undefined) {
          ffd.name = String(ffDataObj['w:name']['@_w:val']);
        }
        // w:enabled — CT_OnOff per ECMA-376 §17.16.11; presence = true, w:val honours ST_OnOff
        if (ffDataObj['w:enabled'] !== undefined) {
          ffd.enabled = parseOoxmlBoolean(ffDataObj['w:enabled']);
        }
        // w:calcOnExit — CT_OnOff per ECMA-376 §17.16.4; presence = true, w:val honours ST_OnOff
        if (ffDataObj['w:calcOnExit'] !== undefined) {
          ffd.calcOnExit = parseOoxmlBoolean(ffDataObj['w:calcOnExit']);
        }
        // w:helpText
        if (ffDataObj['w:helpText']?.['@_w:val'] !== undefined) {
          ffd.helpText = String(ffDataObj['w:helpText']['@_w:val']);
        }
        // w:statusText
        if (ffDataObj['w:statusText']?.['@_w:val'] !== undefined) {
          ffd.statusText = String(ffDataObj['w:statusText']['@_w:val']);
        }
        // w:entryMacro
        if (ffDataObj['w:entryMacro']?.['@_w:val'] !== undefined) {
          ffd.entryMacro = String(ffDataObj['w:entryMacro']['@_w:val']);
        }
        // w:exitMacro
        if (ffDataObj['w:exitMacro']?.['@_w:val'] !== undefined) {
          ffd.exitMacro = String(ffDataObj['w:exitMacro']['@_w:val']);
        }
        // w:textInput
        if (ffDataObj['w:textInput'] !== undefined) {
          const ti: XmlNode = ffDataObj['w:textInput'];
          const textInput: FormFieldTextInput = { type: 'textInput' };
          if (ti['w:type']?.['@_w:val'] !== undefined)
            textInput.inputType = String(ti['w:type']['@_w:val']);
          if (ti['w:default']?.['@_w:val'] !== undefined)
            textInput.defaultValue = String(ti['w:default']['@_w:val']);
          if (ti['w:maxLength']?.['@_w:val'] !== undefined)
            textInput.maxLength = Number(ti['w:maxLength']['@_w:val']);
          if (ti['w:format']?.['@_w:val'] !== undefined)
            textInput.format = String(ti['w:format']['@_w:val']);
          ffd.fieldType = textInput;
        }
        // w:checkBox
        if (ffDataObj['w:checkBox'] !== undefined) {
          const cb: XmlNode = ffDataObj['w:checkBox'];
          const checkBox: FormFieldCheckBox = { type: 'checkBox' };
          // w:default / w:checked are CT_OnOff per ECMA-376 §17.16.18 —
          // honour every ST_OnOff literal ("true"/"false"/"1"/"0"/"on"/"off")
          // and treat a bare self-closing element as true.
          if (cb['w:default'] !== undefined) {
            checkBox.defaultChecked = parseOoxmlBoolean(cb['w:default']);
          }
          if (cb['w:checked'] !== undefined) {
            checkBox.checked = parseOoxmlBoolean(cb['w:checked']);
          }
          if (cb['w:size']?.['@_w:val'] !== undefined) {
            checkBox.size = Number(cb['w:size']['@_w:val']);
          } else if (cb['w:sizeAuto'] !== undefined) {
            checkBox.size = 'auto';
          }
          ffd.fieldType = checkBox;
        }
        // w:ddList
        if (ffDataObj['w:ddList'] !== undefined) {
          const dd: XmlNode = ffDataObj['w:ddList'];
          const ddList: FormFieldDropDownList = { type: 'dropDownList' };
          if (dd['w:result']?.['@_w:val'] !== undefined)
            ddList.result = Number(dd['w:result']['@_w:val']);
          if (dd['w:default']?.['@_w:val'] !== undefined)
            ddList.defaultResult = Number(dd['w:default']['@_w:val']);
          if (dd['w:listEntry'] !== undefined) {
            const entries = Array.isArray(dd['w:listEntry'])
              ? dd['w:listEntry']
              : [dd['w:listEntry']];
            ddList.listEntries = entries.map((e: XmlNode) => String(e?.['@_w:val'] ?? e ?? ''));
          }
          ffd.fieldType = ddList;
        }
        return Object.keys(ffd).length > 0 ? ffd : undefined;
      };

      // Use _orderedChildren to preserve element order (critical for TOC entries)
      // TOC entries have structure: text → tab → text (heading, tab, page number)
      if (runObj._orderedChildren) {
        // Process elements in their original order
        for (const child of runObj._orderedChildren) {
          const elementType = child.type;
          const elementIndex = child.index;

          switch (elementType) {
            case 'w:t': {
              const textElements = toArray(runObj['w:t']);
              // Bounds check with debug logging for malformed documents
              if (elementIndex >= textElements.length) {
                defaultLogger.debug('[DocumentParser] Invalid _orderedChildren index for w:t', {
                  index: elementIndex,
                  arrayLength: textElements.length,
                });
                break;
              }
              const te = textElements[elementIndex];
              if (te !== undefined && te !== null) {
                const text = extractTextValue(te);
                if (text) {
                  content.push({ type: 'text', value: text });
                }
              }
              break;
            }

            case 'w:instrText': {
              const instrElements = toArray(runObj['w:instrText']);
              // Bounds check with debug logging for malformed documents
              if (elementIndex >= instrElements.length) {
                defaultLogger.debug(
                  '[DocumentParser] Invalid _orderedChildren index for w:instrText',
                  { index: elementIndex, arrayLength: instrElements.length }
                );
                break;
              }
              const instr = instrElements[elementIndex];
              if (instr !== undefined && instr !== null) {
                const text = extractTextValue(instr);
                content.push({ type: 'instructionText', value: text });
              }
              break;
            }

            // Deleted text element (w:delText) - inside w:del tracked changes
            // Per ECMA-376 Part 1 §22.1.2.27, deleted text uses w:delText instead of w:t
            case 'w:delText': {
              const delTextElements = toArray(runObj['w:delText']);
              if (elementIndex >= delTextElements.length) {
                defaultLogger.debug(
                  '[DocumentParser] Invalid _orderedChildren index for w:delText',
                  { index: elementIndex, arrayLength: delTextElements.length }
                );
                break;
              }
              const te = delTextElements[elementIndex];
              if (te !== undefined && te !== null) {
                const text = extractTextValue(te);
                if (text) {
                  // Store as deleted text - same as regular text for content purposes
                  content.push({ type: 'text', value: text, isDeleted: true });
                }
              }
              break;
            }

            // Deleted instruction text (w:delInstrText) - inside w:del for field codes
            // Per ECMA-376 Part 1 §22.1.2.26, deleted field instructions use w:delInstrText
            case 'w:delInstrText': {
              const delInstrElements = toArray(runObj['w:delInstrText']);
              if (elementIndex >= delInstrElements.length) {
                defaultLogger.debug(
                  '[DocumentParser] Invalid _orderedChildren index for w:delInstrText',
                  { index: elementIndex, arrayLength: delInstrElements.length }
                );
                break;
              }
              const instr = delInstrElements[elementIndex];
              if (instr !== undefined && instr !== null) {
                const text = extractTextValue(instr);
                // Store as instruction text - field assembly will handle it
                content.push({ type: 'instructionText', value: text, isDeleted: true });
              }
              break;
            }

            case 'w:fldChar': {
              const fldChars = toArray(runObj['w:fldChar']);
              // Bounds check with debug logging for malformed documents
              if (elementIndex >= fldChars.length) {
                defaultLogger.debug(
                  '[DocumentParser] Invalid _orderedChildren index for w:fldChar',
                  { index: elementIndex, arrayLength: fldChars.length }
                );
                break;
              }
              const fldChar = fldChars[elementIndex];
              if (fldChar && typeof fldChar === 'object') {
                const charType = (fldChar['@_w:fldCharType'] || fldChar['@_fldCharType']) as
                  | 'begin'
                  | 'separate'
                  | 'end'
                  | undefined;
                if (charType) {
                  const fldContent: any = {
                    type: 'fieldChar',
                    fieldCharType: charType,
                    fieldCharDirty: parseBooleanAttr(fldChar['@_w:dirty']),
                    fieldCharLocked: parseBooleanAttr(
                      fldChar['@_w:fldLock'] ?? fldChar['@_w:lock']
                    ),
                  };
                  if (charType === 'begin') {
                    const ffData = parseFormFieldData(fldChar);
                    if (ffData) fldContent.formFieldData = ffData;
                  }
                  content.push(fldContent);
                }
              }
              break;
            }

            case 'w:tab':
              content.push({ type: 'tab' });
              break;

            case 'w:br': {
              const brElements = toArray(runObj['w:br']);
              const brElement = brElements[elementIndex] || brElements[0];
              const breakType = brElement?.['@_w:type'] as BreakType | undefined;
              const breakClear = brElement?.['@_w:clear'] as
                | 'none'
                | 'left'
                | 'right'
                | 'all'
                | undefined;
              content.push({ type: 'break', breakType, breakClear });
              break;
            }

            case 'w:cr':
              content.push({ type: 'carriageReturn' });
              break;

            case 'w:softHyphen':
              content.push({ type: 'softHyphen' });
              break;

            case 'w:noBreakHyphen':
              content.push({ type: 'noBreakHyphen' });
              break;

            // Simple marker elements per ECMA-376 Part 1 §17.3.3
            case 'w:lastRenderedPageBreak':
              content.push({ type: 'lastRenderedPageBreak' });
              break;

            case 'w:separator':
              content.push({ type: 'separator' });
              break;

            case 'w:continuationSeparator':
              content.push({ type: 'continuationSeparator' });
              break;

            case 'w:pgNum':
              content.push({ type: 'pageNumber' });
              break;

            case 'w:annotationRef':
              content.push({ type: 'annotationRef' });
              break;

            // Footnote reference (w:footnoteReference) per ECMA-376 Part 1 §17.11.13.
            // w:customMarkFollows is ST_OnOff — honour every literal via parseOnOffAttribute.
            case 'w:footnoteReference': {
              const fnRefElements = toArray(runObj['w:footnoteReference']);
              const fnRef = fnRefElements[elementIndex] || fnRefElements[0];
              const fnId = fnRef?.['@_w:id'];
              const fnCustomMark = fnRef?.['@_w:customMarkFollows'];
              content.push({
                type: 'footnoteReference',
                footnoteId: fnId !== undefined ? parseInt(fnId, 10) : undefined,
                customMarkFollows:
                  fnCustomMark !== undefined ? parseOnOffAttribute(fnCustomMark) : undefined,
              });
              break;
            }

            // Endnote reference (w:endnoteReference) per ECMA-376 Part 1 §17.11.2.
            // Same ST_OnOff treatment for w:customMarkFollows.
            case 'w:endnoteReference': {
              const enRefElements = toArray(runObj['w:endnoteReference']);
              const enRef = enRefElements[elementIndex] || enRefElements[0];
              const enId = enRef?.['@_w:id'];
              const enCustomMark = enRef?.['@_w:customMarkFollows'];
              content.push({
                type: 'endnoteReference',
                endnoteId: enId !== undefined ? parseInt(enId, 10) : undefined,
                customMarkFollows:
                  enCustomMark !== undefined ? parseOnOffAttribute(enCustomMark) : undefined,
              });
              break;
            }

            // Auto-numbered marks INSIDE a footnote/endnote body per
            // ECMA-376 §17.11.14 / §17.11.3. Empty self-closing elements.
            case 'w:footnoteRef':
              content.push({ type: 'footnoteRef' });
              break;
            case 'w:endnoteRef':
              content.push({ type: 'endnoteRef' });
              break;

            case 'w:dayShort':
              content.push({ type: 'dayShort' });
              break;

            case 'w:dayLong':
              content.push({ type: 'dayLong' });
              break;

            case 'w:monthShort':
              content.push({ type: 'monthShort' });
              break;

            case 'w:monthLong':
              content.push({ type: 'monthLong' });
              break;

            case 'w:yearShort':
              content.push({ type: 'yearShort' });
              break;

            case 'w:yearLong':
              content.push({ type: 'yearLong' });
              break;

            // Symbol character (w:sym) per ECMA-376 Part 1 §17.3.3.30
            case 'w:sym': {
              const symElements = toArray(runObj['w:sym']);
              const sym = symElements[elementIndex] || symElements[0];
              if (sym && typeof sym === 'object') {
                content.push({
                  type: 'symbol',
                  symbolFont: sym['@_w:font'],
                  symbolChar: sym['@_w:char'],
                });
              } else {
                content.push({ type: 'symbol' });
              }
              break;
            }

            // Absolute position tab (w:ptab) per ECMA-376 Part 1 §17.3.3.23
            case 'w:ptab': {
              const ptabElements = toArray(runObj['w:ptab']);
              const ptab = ptabElements[elementIndex] || ptabElements[0];
              if (ptab && typeof ptab === 'object') {
                content.push({
                  type: 'positionTab',
                  ptabAlignment: ptab['@_w:alignment'],
                  ptabRelativeTo: ptab['@_w:relativeTo'],
                  ptabLeader: ptab['@_w:leader'],
                });
              } else {
                content.push({ type: 'positionTab' });
              }
              break;
            }

            // Ignore formatting elements (w:rPr) - handled separately
            case 'w:rPr':
              break;
          }
        }
      } else {
        // Fallback: No _orderedChildren (older parser or simple run with single text)
        defaultLogger.warn(
          '[DocumentParser] _orderedChildren missing - using fallback element ordering which may affect tab/break positions'
        );
        // Extract text elements (can be array if multiple <w:t> in one run)
        const textElement = runObj['w:t'];
        if (textElement !== undefined && textElement !== null) {
          const textElements = toArray(textElement);

          for (const te of textElements) {
            const text = extractTextValue(te);
            if (text) {
              content.push({ type: 'text', value: text });
            }
          }
        }

        const instrTextElement = runObj['w:instrText'];
        if (instrTextElement !== undefined && instrTextElement !== null) {
          const instrElements = toArray(instrTextElement);
          for (const instr of instrElements) {
            const text = extractTextValue(instr);
            content.push({ type: 'instructionText', value: text });
          }
        }

        // Handle deleted text (w:delText) - inside w:del tracked changes
        // Per ECMA-376 Part 1 §22.1.2.27
        const delTextElement = runObj['w:delText'];
        if (delTextElement !== undefined && delTextElement !== null) {
          const delTextElements = toArray(delTextElement);
          for (const te of delTextElements) {
            const text = extractTextValue(te);
            if (text) {
              content.push({ type: 'text', value: text, isDeleted: true });
            }
          }
        }

        // Handle deleted instruction text (w:delInstrText) - inside w:del for field codes
        // Per ECMA-376 Part 1 §22.1.2.26
        const delInstrTextElement = runObj['w:delInstrText'];
        if (delInstrTextElement !== undefined && delInstrTextElement !== null) {
          const delInstrElements = toArray(delInstrTextElement);
          for (const instr of delInstrElements) {
            const text = extractTextValue(instr);
            content.push({ type: 'instructionText', value: text, isDeleted: true });
          }
        }

        const fldCharElement = runObj['w:fldChar'];
        if (fldCharElement !== undefined && fldCharElement !== null) {
          const fldChars = toArray(fldCharElement);
          for (const fldChar of fldChars) {
            if (fldChar && typeof fldChar === 'object') {
              const charType = (fldChar['@_w:fldCharType'] || fldChar['@_fldCharType']) as
                | 'begin'
                | 'separate'
                | 'end'
                | undefined;
              if (charType) {
                const fldContent: any = {
                  type: 'fieldChar',
                  fieldCharType: charType,
                  fieldCharDirty: parseBooleanAttr(fldChar['@_w:dirty']),
                  fieldCharLocked: parseBooleanAttr(fldChar['@_w:fldLock'] ?? fldChar['@_w:lock']),
                };
                if (charType === 'begin') {
                  const ffData = parseFormFieldData(fldChar);
                  if (ffData) fldContent.formFieldData = ffData;
                }
                content.push(fldContent);
              }
            }
          }
        }

        // Extract other elements (order doesn't matter in simple case)
        if (runObj['w:tab'] !== undefined) {
          content.push({ type: 'tab' });
        }

        if (runObj['w:br'] !== undefined) {
          const brElement = runObj['w:br'];
          const breakType = brElement?.['@_w:type'] as BreakType | undefined;
          const breakClear = brElement?.['@_w:clear'] as
            | 'none'
            | 'left'
            | 'right'
            | 'all'
            | undefined;
          content.push({ type: 'break', breakType, breakClear });
        }

        if (runObj['w:cr'] !== undefined) {
          content.push({ type: 'carriageReturn' });
        }

        if (runObj['w:softHyphen'] !== undefined) {
          content.push({ type: 'softHyphen' });
        }

        if (runObj['w:noBreakHyphen'] !== undefined) {
          content.push({ type: 'noBreakHyphen' });
        }

        // Simple marker elements (fallback path)
        if (runObj['w:lastRenderedPageBreak'] !== undefined) {
          content.push({ type: 'lastRenderedPageBreak' });
        }
        if (runObj['w:separator'] !== undefined) {
          content.push({ type: 'separator' });
        }
        if (runObj['w:continuationSeparator'] !== undefined) {
          content.push({ type: 'continuationSeparator' });
        }
        if (runObj['w:pgNum'] !== undefined) {
          content.push({ type: 'pageNumber' });
        }
        if (runObj['w:annotationRef'] !== undefined) {
          content.push({ type: 'annotationRef' });
        }
        // Footnote/endnote reference fallback. w:customMarkFollows is ST_OnOff
        // per ECMA-376 §17.11.13 / §17.11.2 — honour every literal.
        if (runObj['w:footnoteReference'] !== undefined) {
          const fnRefElements = toArray(runObj['w:footnoteReference']);
          for (const fnRef of fnRefElements) {
            const fnId = fnRef?.['@_w:id'];
            const fnCustomMark = fnRef?.['@_w:customMarkFollows'];
            content.push({
              type: 'footnoteReference',
              footnoteId: fnId !== undefined ? parseInt(fnId, 10) : undefined,
              customMarkFollows:
                fnCustomMark !== undefined ? parseOnOffAttribute(fnCustomMark) : undefined,
            });
          }
        }
        if (runObj['w:endnoteReference'] !== undefined) {
          const enRefElements = toArray(runObj['w:endnoteReference']);
          for (const enRef of enRefElements) {
            const enId = enRef?.['@_w:id'];
            const enCustomMark = enRef?.['@_w:customMarkFollows'];
            content.push({
              type: 'endnoteReference',
              endnoteId: enId !== undefined ? parseInt(enId, 10) : undefined,
              customMarkFollows:
                enCustomMark !== undefined ? parseOnOffAttribute(enCustomMark) : undefined,
            });
          }
        }
        // Auto-numbered marks INSIDE a footnote/endnote body — empty elements.
        if (runObj['w:footnoteRef'] !== undefined) {
          content.push({ type: 'footnoteRef' });
        }
        if (runObj['w:endnoteRef'] !== undefined) {
          content.push({ type: 'endnoteRef' });
        }
        if (runObj['w:dayShort'] !== undefined) {
          content.push({ type: 'dayShort' });
        }
        if (runObj['w:dayLong'] !== undefined) {
          content.push({ type: 'dayLong' });
        }
        if (runObj['w:monthShort'] !== undefined) {
          content.push({ type: 'monthShort' });
        }
        if (runObj['w:monthLong'] !== undefined) {
          content.push({ type: 'monthLong' });
        }
        if (runObj['w:yearShort'] !== undefined) {
          content.push({ type: 'yearShort' });
        }
        if (runObj['w:yearLong'] !== undefined) {
          content.push({ type: 'yearLong' });
        }

        // Symbol character (w:sym) - fallback path
        if (runObj['w:sym'] !== undefined) {
          const symElements = toArray(runObj['w:sym']);
          for (const sym of symElements) {
            if (sym && typeof sym === 'object') {
              content.push({
                type: 'symbol',
                symbolFont: sym['@_w:font'],
                symbolChar: sym['@_w:char'],
              });
            } else {
              content.push({ type: 'symbol' });
            }
          }
        }

        // Absolute position tab (w:ptab) - fallback path
        if (runObj['w:ptab'] !== undefined) {
          const ptabElements = toArray(runObj['w:ptab']);
          for (const ptab of ptabElements) {
            if (ptab && typeof ptab === 'object') {
              content.push({
                type: 'positionTab',
                ptabAlignment: ptab['@_w:alignment'],
                ptabRelativeTo: ptab['@_w:relativeTo'],
                ptabLeader: ptab['@_w:leader'],
              });
            } else {
              content.push({ type: 'positionTab' });
            }
          }
        }
      }

      // Create run from content elements
      const run = Run.createFromContent(content, { cleanXmlFromText: false });

      // Parse and apply run properties (formatting)
      this.parseRunPropertiesFromObject(runObj['w:rPr'], run);

      // Diagnostic logging
      const text = run.getText();
      const formatting = run.getFormatting();
      if (formatting.rtl) {
        logTextDirection(`Run with RTL: "${text}"`);
      }
      logParsing(`Parsed run: "${text}" (${content.length} content element(s))`, {
        rtl: formatting.rtl || false,
      });

      return run;
    } catch (error: unknown) {
      return null;
    }
  }

  private parseHyperlinkFromObject(
    hyperlinkObj: any,
    relationshipManager: RelationshipManager
  ): {
    hyperlink: Hyperlink | null;
    bookmarkStarts: Bookmark[];
    bookmarkEnds: Bookmark[];
  } {
    const result: {
      hyperlink: Hyperlink | null;
      bookmarkStarts: Bookmark[];
      bookmarkEnds: Bookmark[];
    } = { hyperlink: null, bookmarkStarts: [], bookmarkEnds: [] };

    try {
      // Extract bookmark elements inside the hyperlink
      // These need to be added to the containing paragraph
      if (hyperlinkObj['w:bookmarkStart']) {
        const bookmarkStarts = Array.isArray(hyperlinkObj['w:bookmarkStart'])
          ? hyperlinkObj['w:bookmarkStart']
          : [hyperlinkObj['w:bookmarkStart']];
        for (const bs of bookmarkStarts) {
          const id = bs['@_w:id'];
          // w:name is ST_String per §17.16.5 CT_Bookmark. XMLParser
          // coerces purely-numeric bookmark names ("12345") to JS
          // numbers; cast so Bookmark.name holds the declared string
          // type contract (parent parsers already do the same —
          // iter 125 toOptString helper).
          const rawName = bs['@_w:name'];
          const name =
            rawName === undefined || rawName === null || rawName === ''
              ? undefined
              : String(rawName);
          if (id !== undefined && name) {
            // CT_Bookmark per ECMA-376 §17.16.5: the object-form parser
            // must carry the same four "markup" attributes that the
            // XML-string bookmarkStart parser handles — colFirst/colLast
            // (table-column-scoped bookmarks) and displacedByCustomXml
            // (custom-XML boundary disambiguator). Previously dropped
            // whenever a hyperlink wrapped a bookmark, so inline
            // hyperlinks anchored to table-column bookmarks lost their
            // column range on round-trip.
            const rawColFirst = bs['@_w:colFirst'];
            const rawColLast = bs['@_w:colLast'];
            const rawDisplaced = bs['@_w:displacedByCustomXml'];
            const colFirst =
              rawColFirst === undefined ? undefined : parseInt(String(rawColFirst), 10);
            const colLast = rawColLast === undefined ? undefined : parseInt(String(rawColLast), 10);
            const displacedByCustomXml =
              rawDisplaced === 'next' || rawDisplaced === 'prev' ? rawDisplaced : undefined;
            const bookmark = new Bookmark({
              name: name,
              id: typeof id === 'number' ? id : parseInt(id, 10),
              skipNormalization: true,
              colFirst: Number.isNaN(colFirst as number) ? undefined : colFirst,
              colLast: Number.isNaN(colLast as number) ? undefined : colLast,
              displacedByCustomXml,
            });
            result.bookmarkStarts.push(bookmark);
            // Also register with BookmarkManager
            if (this.bookmarkManager) {
              try {
                this.bookmarkManager.registerExisting(bookmark);
              } catch (e) {
                getLogger().debug('Bookmark already registered', { id: bookmark.getId?.() });
              }
            }
          }
        }
      }

      if (hyperlinkObj['w:bookmarkEnd']) {
        const bookmarkEnds = Array.isArray(hyperlinkObj['w:bookmarkEnd'])
          ? hyperlinkObj['w:bookmarkEnd']
          : [hyperlinkObj['w:bookmarkEnd']];
        for (const be of bookmarkEnds) {
          const id = be['@_w:id'];
          if (id !== undefined) {
            // CT_MarkupRange per ECMA-376 §17.13.5 — preserve
            // w:displacedByCustomXml on bookmarkEnd when a custom-XML
            // boundary forced the marker to be displaced.
            const rawDisplaced = be['@_w:displacedByCustomXml'];
            const displacedByCustomXml =
              rawDisplaced === 'next' || rawDisplaced === 'prev' ? rawDisplaced : undefined;
            const bookmark = new Bookmark({
              name: `_end_${id}`,
              id: typeof id === 'number' ? id : parseInt(id, 10),
              skipNormalization: true,
              displacedByCustomXml,
            });
            result.bookmarkEnds.push(bookmark);
          }
        }
      }

      // Extract hyperlink attributes. Per ECMA-376 §17.16.22 CT_Hyperlink,
      // w:anchor / w:tooltip / w:tgtFrame / w:docLocation / r:id are all
      // ST_String. XMLParser's `parseAttributeValue: true` coerces
      // purely-numeric strings (e.g., a bookmark name like "12345") to
      // JS numbers — cast via String(...) so downstream `Hyperlink`
      // storage and string-method callers see the declared `string`
      // type contract.
      const toOptString = (v: unknown): string | undefined =>
        v === undefined || v === null ? undefined : String(v);
      const relationshipId = toOptString(hyperlinkObj['@_r:id']);
      const anchor = toOptString(hyperlinkObj['@_w:anchor']);
      const tooltip = toOptString(hyperlinkObj['@_w:tooltip']);
      const tgtFrame = toOptString(hyperlinkObj['@_w:tgtFrame']);
      // w:history is CT_OnOff per ECMA-376 §17.16.22 — honour every
      // ST_OnOff literal ("1"/"0"/"true"/"false"/"on"/"off") and every
      // XMLParser-coerced form (number 0/1, boolean). The Hyperlink
      // serializer accepts a string, so normalise to the canonical
      // "1"/"0" form. Without this, `w:history="0"` or `w:history="false"`
      // coerced to falsy values and the emitter's truthy check dropped
      // the attribute on round-trip.
      const rawHistory = hyperlinkObj['@_w:history'];
      const history =
        rawHistory === undefined ? undefined : parseOnOffAttribute(rawHistory) ? '1' : '0';
      const docLocation = toOptString(hyperlinkObj['@_w:docLocation']);

      // Parse runs inside the hyperlink
      const runs = hyperlinkObj['w:r'];
      const runChildren = Array.isArray(runs) ? runs : runs ? [runs] : [];

      // Parse ALL runs to handle multi-run hyperlinks (e.g., varied formatting within one hyperlink)
      // Google Docs often splits hyperlinks by formatting changes, creating multiple runs
      const parsedRuns: Run[] = [];
      let text = '';
      let formatting: RunFormatting = {};

      if (runChildren.length > 0) {
        // Parse all runs, not just the first one
        for (const runChild of runChildren) {
          const parsedRun = this.parseRunFromObject(runChild);
          if (parsedRun) {
            parsedRuns.push(parsedRun);
            text += parsedRun.getText();
          }
        }

        // Use the first run's formatting as the base formatting
        if (parsedRuns.length > 0 && parsedRuns[0]) {
          formatting = parsedRuns[0].getFormatting();
        }
      }

      // For TOC hyperlinks with tabs/breaks, use the first parsed run
      const parsedRun = parsedRuns.length > 0 ? parsedRuns[0] : null;

      // Resolve URL from relationship if external hyperlink
      let url: string | undefined;
      if (relationshipId) {
        const relationship = relationshipManager.getRelationship(relationshipId);
        if (relationship) {
          url = relationship.getTarget();
          // Decode URL-encoded characters (e.g., %23 -> #, %20 -> space)
          // This ensures proper handling of URLs with fragments and special characters
          if (url) {
            try {
              url = decodeURIComponent(url);
            } catch {
              // If decoding fails, use the original URL
              defaultLogger.debug(`[DocumentParser] Failed to decode URL: ${url}`);
            }
          }
        }
      }

      // Per ECMA-376 §17.16.22, a hyperlink can have BOTH r:id (external URL) and w:anchor
      // (bookmark) simultaneously — e.g., linking to a bookmark in an external document.
      // Preserve both attributes as-is; the serializer supports writing both.
      const finalAnchor = anchor;
      const finalRelationshipId = relationshipId;

      // Skip hyperlinks that have no destination (neither URL nor anchor nor relationship ID)
      // This can happen with malformed HYPERLINK field codes or corrupted documents
      // Note: If there's a relationshipId but the relationship is missing, we still keep the hyperlink
      // (it has text and a reference that might be resolved later or is just broken)
      if (!url && !finalAnchor && !finalRelationshipId) {
        defaultLogger.debug(
          `[DocumentParser] Skipping hyperlink with no URL, anchor, or relationship ID. Text: "${text}"`
        );
        return result;
      }

      // Skip self-closing external hyperlinks (no runs at all)
      // These are invisible hyperlinks that exist in the document structure but have no visible content
      // They should be removed rather than having URLs appear as text
      const isSelfClosingHyperlink = runChildren.length === 0;
      const isExternalLink = (url || finalRelationshipId) && !finalAnchor;

      if (isSelfClosingHyperlink && isExternalLink) {
        // Skip self-closing external hyperlink - it has no visible content
        defaultLogger.debug(
          `[DocumentParser] Skipping self-closing external hyperlink with no display text. URL: "${url}"`
        );
        return result;
      }

      // Create hyperlink with display text
      // NOTE: Do NOT use anchor (bookmark ID) as display text - it should only be used for navigation
      const displayText = text || url || '[Link]';

      // Warn if hyperlink has no display text (possible TOC corruption or malformed hyperlink)
      if (!text && finalAnchor) {
        defaultLogger.warn(
          `[DocumentParser] Hyperlink to anchor "${finalAnchor}" has no display text. ` +
            `Using placeholder "[Link]" to prevent bookmark ID from appearing as visible text. ` +
            `This may indicate a corrupted TOC or malformed hyperlink in the source document.`
        );
      }

      const hyperlink = new Hyperlink({
        url,
        anchor: finalAnchor,
        text: displayText,
        formatting,
        tooltip,
        relationshipId: finalRelationshipId,
        tgtFrame,
        history,
        docLocation,
      });

      // If we successfully parsed a run with tabs/breaks, use it instead of the default run
      // This preserves TOC structure (text → tab → text)
      if (parsedRun && parsedRun.getContent().length > 1) {
        hyperlink.setRun(parsedRun);
      }

      result.hyperlink = hyperlink;
      return result;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse hyperlink:',
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return result;
    }
  }

  /**
   * Merges TRULY consecutive hyperlinks with the same URL into a single hyperlink
   * This handles Google Docs-style hyperlinks that are split by formatting changes
   *
   * IMPORTANT: Only merges hyperlinks that are IMMEDIATELY adjacent (no content between them).
   * Hyperlinks with the same URL but with intervening text/runs are NOT merged.
   *
   * @param paragraph - Paragraph containing hyperlinks to merge
   * @param resetFormatting - Whether to reset hyperlinks to standard formatting
   * @private
   */
  private mergeConsecutiveHyperlinks(paragraph: Paragraph, resetFormatting = false): void {
    const content = paragraph.getContent();
    if (!content || content.length < 2) return;

    // Guard: Skip merge when tracking is enabled - it corrupts field structures
    // The clearContent() + addHyperlink() pattern creates new revisions
    // that end up at the wrong position in the content array, placing them
    // OUTSIDE field boundaries when field codes are present
    const trackingEnabled = (paragraph as any).trackingContext?.isEnabled();
    if (trackingEnabled) {
      const hasFieldContent = content.some((item) => {
        if (item instanceof Run) {
          const runContent = item.getContent();
          return runContent.some(
            (c: any) => c.type === 'fieldChar' || c.type === 'instructionText'
          );
        }
        // Also check inside Revisions for field content
        if (item instanceof Revision) {
          const revContent = item.getContent();
          for (const inner of revContent) {
            if (inner instanceof Run) {
              const innerRunContent = inner.getContent();
              if (
                innerRunContent.some(
                  (c: any) => c.type === 'fieldChar' || c.type === 'instructionText'
                )
              ) {
                return true;
              }
            }
          }
        }
        return false;
      });

      if (hasFieldContent) {
        defaultLogger.debug(
          'Skipping hyperlink merge: tracking enabled with field content present'
        );
        return;
      }
    }

    const mergedContent: any[] = [];
    let i = 0;
    let contentChanged = false;

    while (i < content.length) {
      const item = content[i];

      if (item instanceof Hyperlink) {
        // Look ahead to find CONSECUTIVE hyperlinks with the same URL/anchor
        const url = item.getUrl() || '';
        const anchor = item.getAnchor() || '';
        const consecutiveHyperlinks: Hyperlink[] = [item];
        let j = i + 1;

        // Only merge if IMMEDIATELY adjacent (next item is also a hyperlink with same URL)
        while (j < content.length) {
          const nextItem = content[j];
          if (
            nextItem instanceof Hyperlink &&
            (nextItem.getUrl() || '') === url &&
            (nextItem.getAnchor() || '') === anchor
          ) {
            consecutiveHyperlinks.push(nextItem);
            j++;
          } else {
            // Stop at first non-matching item (including non-hyperlink items)
            break;
          }
        }

        if (consecutiveHyperlinks.length > 1) {
          // Merge consecutive hyperlinks
          const mergedText = consecutiveHyperlinks.map((h) => h.getText()).join('');

          const mergedHyperlink = new Hyperlink({
            url: item.getUrl(),
            anchor: item.getAnchor(),
            text: mergedText,
            formatting: resetFormatting
              ? this.getStandardHyperlinkFormatting()
              : item.getFormatting(),
            tooltip: item.getTooltip(),
            relationshipId: item.getRelationshipId(),
          });
          mergedContent.push(mergedHyperlink);
          contentChanged = true;
        } else {
          // Single hyperlink (no consecutive ones to merge)
          if (resetFormatting) {
            const resetHyperlink = new Hyperlink({
              url: item.getUrl(),
              anchor: item.getAnchor(),
              text: item.getText(),
              formatting: this.getStandardHyperlinkFormatting(),
              tooltip: item.getTooltip(),
              relationshipId: item.getRelationshipId(),
            });
            mergedContent.push(resetHyperlink);
            contentChanged = true;
          } else {
            mergedContent.push(item);
          }
        }
        i = j;
      } else {
        // Not a hyperlink, keep as-is
        mergedContent.push(item);
        i++;
      }
    }

    // Update paragraph content only if we actually changed something
    if (contentChanged) {
      // Clear current content
      paragraph.clearContent();

      // Add merged content back
      for (const item of mergedContent) {
        if (item instanceof Hyperlink) {
          paragraph.addHyperlink(item);
        } else if (item instanceof Run) {
          paragraph.addRun(item);
        } else if (item instanceof Field) {
          paragraph.addField(item);
        } else {
          // Preserve all other content types: Revision, RangeMarker, Shape, TextBox, PreservedElement
          paragraph.addContent(item);
        }
      }
    }
  }

  /**
   * Get standard hyperlink formatting (blue, underline)
   * @private
   */
  private getStandardHyperlinkFormatting(): any {
    return {
      color: '0000FF', // Standard hyperlink blue
      underline: 'single',
    };
  }

  /**
   * Parses a simple field (w:fldSimple) from XML object
   * @param fieldObj - Field XML object
   * @returns Parsed Field or null if parsing fails
   * @private
   */
  private parseSimpleFieldFromObject(fieldObj: any): Field | null {
    try {
      // Extract field instruction from w:instr attribute
      const instruction = fieldObj['@_w:instr'];
      if (!instruction) {
        return null;
      }

      // Extract field type from instruction (first word)
      const typeMatch = String(instruction)
        .trim()
        .match(/^(\w+)/);
      const type = (typeMatch?.[1] || 'PAGE') as import('../elements/Field').FieldType;

      // CT_SimpleField (§17.16.16) carries two ST_OnOff attributes besides
      // the required w:instr — w:fldLock (update lock) and w:dirty
      // (cached-result staleness). Previously neither was parsed, so
      // Word's "update field" indicator and "lock field" flag were
      // silently cleared on every load → save round-trip.
      const fldLockRaw = fieldObj['@_w:fldLock'];
      const dirtyRaw = fieldObj['@_w:dirty'];
      const fldLock = fldLockRaw !== undefined ? parseOnOffAttribute(fldLockRaw) : undefined;
      const dirty = dirtyRaw !== undefined ? parseOnOffAttribute(dirtyRaw) : undefined;

      // Parse run formatting from w:rPr if present
      let formatting: RunFormatting | undefined;
      if (fieldObj['w:rPr']) {
        const tempRun = new Run('');
        this.parseRunPropertiesFromObject(fieldObj['w:rPr'], tempRun);
        formatting = tempRun.getFormatting();
      }

      // Create field with instruction
      const field = Field.create({
        type,
        instruction: String(instruction),
        formatting,
        fldLock,
        dirty,
      });

      return field;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse field:',
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
  }

  private parseRunPropertiesFromObject(rPrObj: any, run: Run): void {
    if (!rPrObj) return;

    // Parse character style reference (w:rStyle) per ECMA-376 Part 1
    // §17.3.2.36 — `w:val` is ST_String referencing a style ID. Cast
    // via String(...) so a purely-numeric style ID (e.g., "1") that
    // XMLParser coerces to the number 1 survives as the string "1",
    // matching the `characterStyle?: string` field contract on
    // RunFormatting.
    if (rPrObj['w:rStyle']) {
      const styleId = rPrObj['w:rStyle']['@_w:val'];
      if (styleId !== undefined && styleId !== null && styleId !== '') {
        run.setCharacterStyle(String(styleId));
      }
    }

    // Parse text border (w:bdr) per ECMA-376 Part 1 §17.3.2.5 — CT_Border
    // §17.18.2 attribute set: val / sz / space / color / themeColor /
    // themeTint / themeShade / shadow / frame. Previously only the first
    // four were read, so themed character borders lost their theme linkage
    // on round-trip. The emitter (Run.generateRunPropertiesXML) handles
    // all nine since iteration 79.
    if (rPrObj['w:bdr']) {
      const bdr = rPrObj['w:bdr'];
      const border: any = {};
      if (bdr['@_w:val']) border.style = bdr['@_w:val'];
      if (bdr['@_w:sz']) border.size = parseInt(bdr['@_w:sz'], 10);
      if (bdr['@_w:color']) border.color = bdr['@_w:color'];
      if (bdr['@_w:space']) border.space = parseInt(bdr['@_w:space'], 10);
      // Per ECMA-376 §17.18.82 CT_Border: themeTint / themeShade are
      // ST_UcharHexNumber (2-char hex). XMLParser coerces purely-digit
      // hex strings like "80" / "50" to JS numbers; cast via String(...)
      // so the declared `string` contract on the model holds for any
      // downstream code that calls string methods (.toUpperCase(), etc.).
      if (bdr['@_w:themeColor']) border.themeColor = String(bdr['@_w:themeColor']);
      if (bdr['@_w:themeTint']) border.themeTint = String(bdr['@_w:themeTint']);
      if (bdr['@_w:themeShade']) border.themeShade = String(bdr['@_w:themeShade']);
      if (bdr['@_w:shadow'] !== undefined) {
        border.shadow = parseOnOffAttribute(String(bdr['@_w:shadow']), true);
      }
      if (bdr['@_w:frame'] !== undefined) {
        border.frame = parseOnOffAttribute(String(bdr['@_w:frame']), true);
      }
      if (Object.keys(border).length > 0) {
        run.setBorder(border);
      }
    }

    // Parse character shading (w:shd) per ECMA-376 Part 1 §17.3.2.32
    if (rPrObj['w:shd']) {
      const shading = this.parseShadingFromObj(rPrObj['w:shd']);
      if (shading) {
        run.setShading(shading);
      }
    }

    // Parse emphasis marks (w:em) per ECMA-376 Part 1 §17.3.2.13
    if (rPrObj['w:em']) {
      const val = rPrObj['w:em']['@_w:val'];
      if (val) run.setEmphasis(val);
    }

    // CT_OnOff text effects — presence + w:val both matter. Use `!== undefined`
    // to detect presence, then parseOoxmlBoolean() for the value, so an explicit
    // `<w:outline w:val="0"/>` override of a style-inherited true is preserved
    // (not silently dropped into "inherit"). Applies to all OnOffType rPr flags
    // per ECMA-376 §17.3.2.
    if (rPrObj['w:outline'] !== undefined) run.setOutline(parseOoxmlBoolean(rPrObj['w:outline']));
    if (rPrObj['w:shadow'] !== undefined) run.setShadow(parseOoxmlBoolean(rPrObj['w:shadow']));
    if (rPrObj['w:emboss'] !== undefined) run.setEmboss(parseOoxmlBoolean(rPrObj['w:emboss']));
    if (rPrObj['w:imprint'] !== undefined) run.setImprint(parseOoxmlBoolean(rPrObj['w:imprint']));
    if (rPrObj['w:noProof'] !== undefined) run.setNoProof(parseOoxmlBoolean(rPrObj['w:noProof']));
    // snapToGrid: default when absent is true (§17.3.2.34), so explicit val="0" must be preserved
    if (rPrObj['w:snapToGrid'] !== undefined) {
      run.setSnapToGrid(parseOoxmlBoolean(rPrObj['w:snapToGrid']));
    }
    if (rPrObj['w:vanish'] !== undefined) run.setVanish(parseOoxmlBoolean(rPrObj['w:vanish']));
    if (rPrObj['w:specVanish'] !== undefined)
      run.setSpecVanish(parseOoxmlBoolean(rPrObj['w:specVanish']));

    // Boolean properties - use parseOoxmlBoolean helper
    // Per ECMA-376: <w:b/> or <w:b w:val="1"/> or <w:b w:val="true"/> means true
    // <w:b w:val="0"/> or <w:b w:val="false"/> means false (omit from document)

    // Parse RTL text (w:rtl) per ECMA-376 Part 1 §17.3.2.30
    if (rPrObj['w:rtl'] !== undefined) run.setRTL(parseOoxmlBoolean(rPrObj['w:rtl']));

    // b, bCs, i, iCs: preserve explicit val="0" to override style-inherited formatting
    if (rPrObj['w:b'] !== undefined) run.setBold(parseOoxmlBoolean(rPrObj['w:b']));
    if (rPrObj['w:bCs'] !== undefined) run.setComplexScriptBold(parseOoxmlBoolean(rPrObj['w:bCs']));
    if (rPrObj['w:i'] !== undefined) run.setItalic(parseOoxmlBoolean(rPrObj['w:i']));
    if (rPrObj['w:iCs'] !== undefined)
      run.setComplexScriptItalic(parseOoxmlBoolean(rPrObj['w:iCs']));
    // strike, dstrike, smallCaps, caps: preserve explicit val="0" to override style-inherited formatting
    if (rPrObj['w:strike'] !== undefined) run.setStrike(parseOoxmlBoolean(rPrObj['w:strike']));
    if (rPrObj['w:dstrike'] !== undefined) {
      (run as any).formatting.dstrike = parseOoxmlBoolean(rPrObj['w:dstrike']);
    }
    if (rPrObj['w:smallCaps'] !== undefined)
      run.setSmallCaps(parseOoxmlBoolean(rPrObj['w:smallCaps']));
    if (rPrObj['w:caps'] !== undefined) run.setAllCaps(parseOoxmlBoolean(rPrObj['w:caps']));

    // Parse complex script flag (w:cs) per ECMA-376 Part 1 §17.3.2.7 — CT_OnOff
    if (rPrObj['w:cs'] !== undefined) run.setComplexScript(parseOoxmlBoolean(rPrObj['w:cs']));

    // Parse web hidden (w:webHidden) per ECMA-376 Part 1 §17.3.2.44 — CT_OnOff
    if (rPrObj['w:webHidden'] !== undefined)
      run.setWebHidden(parseOoxmlBoolean(rPrObj['w:webHidden']));

    if (rPrObj['w:u']) {
      // XMLParser adds @_ prefix to attributes
      const uVal = rPrObj['w:u']['@_w:val'];
      run.setUnderline(uVal || true);
      // Parse underline color attributes per ECMA-376 Part 1 §17.3.2.40
      const uColor = rPrObj['w:u']['@_w:color'];
      if (uColor) run.setUnderlineColor(uColor);
      const uThemeColor = rPrObj['w:u']['@_w:themeColor'];
      if (uThemeColor)
        run.setUnderlineThemeColor(
          uThemeColor,
          rPrObj['w:u']['@_w:themeTint'] ? parseInt(rPrObj['w:u']['@_w:themeTint'], 16) : undefined,
          rPrObj['w:u']['@_w:themeShade']
            ? parseInt(rPrObj['w:u']['@_w:themeShade'], 16)
            : undefined
        );
    }

    // Parse character spacing (w:spacing) per ECMA-376 Part 1 §17.3.2.35.
    // ST_SignedTwipsMeasure — 0 and negative values are valid (default /
    // tighter spacing). XMLParser.parseAttributeValue coerces "0" to number 0,
    // which is falsy — so the previous `if (val)` truthy check silently dropped
    // explicit zero / baseline-reset formatting on every run that used it.
    // Matches the rPrChange parser below which already uses `!== undefined`.
    if (rPrObj['w:spacing']) {
      const val = rPrObj['w:spacing']['@_w:val'];
      if (val !== undefined) run.setCharacterSpacing(parseInt(String(val), 10));
    }

    // Parse horizontal scaling (w:w) per ECMA-376 Part 1 §17.3.2.43.
    // ST_TextScale — min 1 per schema, so value 0 is not spec-valid; keep
    // truthy check as a mild sanity guard against malformed sources.
    if (rPrObj['w:w']) {
      const val = rPrObj['w:w']['@_w:val'];
      if (val) run.setScaling(parseInt(String(val), 10));
    }

    // Parse vertical position (w:position) per ECMA-376 Part 1 §17.3.2.31.
    // ST_SignedHpsMeasure — 0 = baseline (default / explicit reset).
    if (rPrObj['w:position']) {
      const val = rPrObj['w:position']['@_w:val'];
      if (val !== undefined) run.setPosition(parseInt(String(val), 10));
    }

    // Parse kerning (w:kern) per ECMA-376 Part 1 §17.3.2.20.
    // ST_HpsMeasure — 0 means "kern at every size" (no minimum threshold).
    if (rPrObj['w:kern']) {
      const val = rPrObj['w:kern']['@_w:val'];
      if (val !== undefined) run.setKerning(parseInt(String(val), 10));
    }

    // Parse language (w:lang) per ECMA-376 Part 1 §17.3.2.20 (CT_Language)
    if (rPrObj['w:lang']) {
      const langObj = rPrObj['w:lang'];
      const val = langObj['@_w:val'];
      const eastAsia = langObj['@_w:eastAsia'];
      const bidi = langObj['@_w:bidi'];
      if (eastAsia || bidi) {
        run.setLanguage({
          val: val ? String(val) : undefined,
          eastAsia: eastAsia ? String(eastAsia) : undefined,
          bidi: bidi ? String(bidi) : undefined,
        });
      } else if (val) {
        run.setLanguage(String(val));
      }
    }

    // Parse East Asian layout (w:eastAsianLayout) per ECMA-376 Part 1
    // §17.3.2.10 CT_EastAsianLayout. `w:vert` / `w:vertCompress` /
    // `w:combine` are ST_OnOff attributes — route through
    // parseOnOffAttribute so every literal ("1"/"0"/"true"/"false"/
    // "on"/"off") resolves correctly. The previous truthy gate both
    // dropped explicit false (`w:vert="0"` → coerced 0 → undefined) AND
    // wrongly marked `w:vert="off"` as true (non-empty string is truthy
    // without parsing).
    if (rPrObj['w:eastAsianLayout']) {
      const layoutObj = rPrObj['w:eastAsianLayout'];
      const layout: any = {};
      if (layoutObj['@_w:id'] !== undefined) layout.id = Number(layoutObj['@_w:id']);
      if (layoutObj['@_w:vert'] !== undefined) {
        layout.vert = parseOnOffAttribute(String(layoutObj['@_w:vert']), true);
      }
      if (layoutObj['@_w:vertCompress'] !== undefined) {
        layout.vertCompress = parseOnOffAttribute(String(layoutObj['@_w:vertCompress']), true);
      }
      if (layoutObj['@_w:combine'] !== undefined) {
        layout.combine = parseOnOffAttribute(String(layoutObj['@_w:combine']), true);
      }
      if (layoutObj['@_w:combineBrackets'])
        layout.combineBrackets = layoutObj['@_w:combineBrackets'];

      if (Object.keys(layout).length > 0) {
        run.setEastAsianLayout(layout);
      }
    }

    // Parse fit text (w:fitText) per ECMA-376 Part 1 §17.3.2.15
    if (rPrObj['w:fitText']) {
      const val = rPrObj['w:fitText']['@_w:val'];
      if (val !== undefined) run.setFitText(Number(val));
    }

    // Parse text effect (w:effect) per ECMA-376 Part 1 §17.3.2.12
    if (rPrObj['w:effect']) {
      const val = rPrObj['w:effect']['@_w:val'];
      if (val) run.setEffect(val);
    }

    if (rPrObj['w:vertAlign']) {
      const val = rPrObj['w:vertAlign']['@_w:val'];
      if (val === 'subscript') run.setSubscript(true);
      else if (val === 'superscript') run.setSuperscript(true);
      else if (val === 'baseline') (run as any).formatting.vertAlignBaseline = true;
    }

    if (rPrObj['w:rFonts']) {
      const rFonts = rPrObj['w:rFonts'];
      // Per ECMA-376 §17.3.2.26 CT_Fonts, all four literal-font
      // attributes (ascii/hAnsi/eastAsia/cs) are ST_String. XMLParser
      // coerces purely-numeric font names ("2010", etc.) to JS
      // numbers; cast through String() so RunFormatting's
      // declared-string font fields keep their type contract.
      if (rFonts['@_w:ascii'] !== undefined) run.setFont(String(rFonts['@_w:ascii']));
      // Parse additional font variants per ECMA-376 Part 1 §17.3.2.26
      if (rFonts['@_w:hAnsi'] !== undefined) run.setFontHAnsi(String(rFonts['@_w:hAnsi']));
      if (rFonts['@_w:eastAsia'] !== undefined) run.setFontEastAsia(String(rFonts['@_w:eastAsia']));
      if (rFonts['@_w:cs'] !== undefined) run.setFontCs(String(rFonts['@_w:cs']));
      if (rFonts['@_w:hint']) run.setFontHint(String(rFonts['@_w:hint']));
      // Parse theme font references per ECMA-376 Part 1 §17.3.2.26
      if (rFonts['@_w:asciiTheme']) run.setFontAsciiTheme(String(rFonts['@_w:asciiTheme']));
      if (rFonts['@_w:hAnsiTheme']) run.setFontHAnsiTheme(String(rFonts['@_w:hAnsiTheme']));
      if (rFonts['@_w:eastAsiaTheme'])
        run.setFontEastAsiaTheme(String(rFonts['@_w:eastAsiaTheme']));
      if (rFonts['@_w:cstheme']) run.setFontCsTheme(String(rFonts['@_w:cstheme']));
    }

    if (rPrObj['w:sz']) {
      run.setSize(halfPointsToPoints(parseInt(rPrObj['w:sz']['@_w:val'], 10)));
    }

    // Parse complex script font size (w:szCs) per ECMA-376 Part 1 §17.3.2.40
    // This is separate from regular size to support RTL languages
    if (rPrObj['w:szCs']) {
      const szCsVal = halfPointsToPoints(parseInt(rPrObj['w:szCs']['@_w:val'], 10));
      // Only set sizeCs if it differs from size (to avoid redundant storage)
      // When sz is not present (undefined), always store szCs
      const sizeVal = rPrObj['w:sz']
        ? halfPointsToPoints(parseInt(rPrObj['w:sz']['@_w:val'], 10))
        : undefined;
      if (sizeVal === undefined || szCsVal !== sizeVal) {
        run.setSizeCs(szCsVal);
      }
    }

    if (rPrObj['w:color']) {
      const colorObj = rPrObj['w:color'];
      const colorVal = colorObj['@_w:val'];
      // Per ECMA-376 §17.18.6, w:val can be a hex color OR the special value "auto"
      // "auto" means use the automatic/window text color — must be preserved for round-trip
      if (colorVal) {
        if (colorVal === 'auto') {
          // Bypass normalizeColor() which rejects non-hex values
          (run as any).formatting.color = 'auto';
        } else {
          run.setColor(colorVal);
        }
      }
      // Parse theme color attributes per ECMA-376 Part 1 Section 17.3.2.6
      if (colorObj['@_w:themeColor']) {
        run.setThemeColor(colorObj['@_w:themeColor']);
      }
      if (colorObj['@_w:themeTint']) {
        // Theme tint is stored as hex string, convert to number
        run.setThemeTint(parseInt(colorObj['@_w:themeTint'], 16));
      }
      if (colorObj['@_w:themeShade']) {
        // Theme shade is stored as hex string, convert to number
        run.setThemeShade(parseInt(colorObj['@_w:themeShade'], 16));
      }
    }

    if (rPrObj['w:highlight']) {
      run.setHighlight(rPrObj['w:highlight']['@_w:val']);
    }

    // Collect w14: namespace elements from rPr for passthrough (Word 2010+ text effects)
    // w14:textOutline, w14:shadow, w14:reflection, w14:glow, w14:ligatures,
    // w14:numForm, w14:numSpacing, w14:cntxtAlts, w14:stylisticSets
    for (const key of Object.keys(rPrObj)) {
      if (key.startsWith('w14:')) {
        const rawXml = this.objectToXml({ [key]: rPrObj[key] });
        if (rawXml) {
          run.addRawW14Property(rawXml);
        }
      }
    }

    // Parse run property change tracking (w:rPrChange) per ECMA-376 Part 1 §17.13.5.30
    // This records what the run formatting was BEFORE a change was made
    if (rPrObj['w:rPrChange']) {
      const changeObj = rPrObj['w:rPrChange'];
      const propChange: import('../elements/PropertyChangeTypes').RunPropertyChange = {
        id: changeObj['@_w:id'] !== undefined ? parseInt(String(changeObj['@_w:id']), 10) : 0,
        author: changeObj['@_w:author'] ? String(changeObj['@_w:author']) : '',
        date: changeObj['@_w:date'] ? new Date(String(changeObj['@_w:date'])) : new Date(),
        previousProperties: {},
      };

      // Parse previous run properties from child w:rPr element
      if (changeObj['w:rPr']) {
        const prevRPr = changeObj['w:rPr'];
        const prevProps: Partial<import('../elements/Run').RunFormatting> = {};

        // Parse previous bold
        if (prevRPr['w:b']) {
          prevProps.bold = parseOoxmlBoolean(prevRPr['w:b']);
        }

        // Parse previous italic
        if (prevRPr['w:i']) {
          prevProps.italic = parseOoxmlBoolean(prevRPr['w:i']);
        }

        // Parse previous underline — CT_Underline per §17.3.2.40 has `val`
        // plus color / themeColor / themeTint / themeShade. Main rPr parser
        // reads all of them; rPrChange previously only read `val`, so
        // underline color metadata on tracked "previous" state was dropped.
        if (prevRPr['w:u']) {
          const uObj = prevRPr['w:u'];
          const uVal = uObj['@_w:val'];
          prevProps.underline = uVal || true;
          if (uObj['@_w:color']) prevProps.underlineColor = uObj['@_w:color'];
          if (uObj['@_w:themeColor']) prevProps.underlineThemeColor = uObj['@_w:themeColor'];
          if (uObj['@_w:themeTint'] !== undefined) {
            prevProps.underlineThemeTint = parseInt(String(uObj['@_w:themeTint']), 16);
          }
          if (uObj['@_w:themeShade'] !== undefined) {
            prevProps.underlineThemeShade = parseInt(String(uObj['@_w:themeShade']), 16);
          }
        }

        // Parse previous strikethrough
        if (prevRPr['w:strike']) {
          prevProps.strike = parseOoxmlBoolean(prevRPr['w:strike']);
        }

        // Parse previous font (all w:rFonts attributes per ECMA-376 Part 1 §17.3.2.26)
        // including theme font references (asciiTheme/hAnsiTheme/eastAsiaTheme/
        // cstheme). Previously only the literal-font attributes were read, so
        // rPrChange tracked history of theme-font changes lost the theme linkage
        // on round-trip — a paragraph whose "previous" font was a theme
        // reference (e.g. w:asciiTheme="minorHAnsi") silently dropped it.
        if (prevRPr['w:rFonts']) {
          const rFonts = prevRPr['w:rFonts'];
          // Mirror the main-path String() casts on rPrChange
          // previous-font reads — ECMA-376 §17.3.2.26 CT_Fonts declares
          // ascii/hAnsi/eastAsia/cs as ST_String, so purely-numeric
          // font names must survive round-trip as strings here too.
          if (rFonts['@_w:ascii'] !== undefined) prevProps.font = String(rFonts['@_w:ascii']);
          if (rFonts['@_w:hAnsi'] !== undefined) prevProps.fontHAnsi = String(rFonts['@_w:hAnsi']);
          if (rFonts['@_w:eastAsia'] !== undefined)
            prevProps.fontEastAsia = String(rFonts['@_w:eastAsia']);
          if (rFonts['@_w:cs'] !== undefined) prevProps.fontCs = String(rFonts['@_w:cs']);
          if (rFonts['@_w:hint']) prevProps.fontHint = String(rFonts['@_w:hint']);
          if (rFonts['@_w:asciiTheme']) prevProps.fontAsciiTheme = String(rFonts['@_w:asciiTheme']);
          if (rFonts['@_w:hAnsiTheme']) prevProps.fontHAnsiTheme = String(rFonts['@_w:hAnsiTheme']);
          if (rFonts['@_w:eastAsiaTheme'])
            prevProps.fontEastAsiaTheme = String(rFonts['@_w:eastAsiaTheme']);
          if (rFonts['@_w:cstheme']) prevProps.fontCsTheme = String(rFonts['@_w:cstheme']);
        }

        // Parse previous size (half-points to points)
        if (prevRPr['w:sz']) {
          prevProps.size = halfPointsToPoints(safeParseInt(prevRPr['w:sz']['@_w:val']));
        }

        // Parse previous complex script size (w:szCs) per ECMA-376 Part 1 §17.3.2.40
        if (prevRPr['w:szCs']) {
          const szCsVal = halfPointsToPoints(safeParseInt(prevRPr['w:szCs']['@_w:val']));
          // Only store if different from regular size
          if (!prevRPr['w:sz'] || szCsVal !== prevProps.size) {
            prevProps.sizeCs = szCsVal;
          }
        }

        // Parse previous color (per ECMA-376 Part 1 Section 17.3.2.6)
        if (prevRPr['w:color']) {
          const colorObj = prevRPr['w:color'];
          const colorVal = colorObj['@_w:val'];
          if (colorVal) {
            prevProps.color = colorVal;
          }
          // Parse theme color attributes
          if (colorObj['@_w:themeColor']) {
            prevProps.themeColor = colorObj['@_w:themeColor'];
          }
          if (colorObj['@_w:themeTint']) {
            prevProps.themeTint = parseInt(colorObj['@_w:themeTint'], 16);
          }
          if (colorObj['@_w:themeShade']) {
            prevProps.themeShade = parseInt(colorObj['@_w:themeShade'], 16);
          }
        }

        // Parse previous highlight
        if (prevRPr['w:highlight']) {
          prevProps.highlight = prevRPr['w:highlight']['@_w:val'];
        }

        // Parse previous subscript/superscript/baseline per ECMA-376 §17.18.96
        if (prevRPr['w:vertAlign']) {
          const val = prevRPr['w:vertAlign']['@_w:val'];
          if (val === 'subscript') prevProps.subscript = true;
          else if (val === 'superscript') prevProps.superscript = true;
          else if (val === 'baseline') prevProps.vertAlignBaseline = true;
        }

        // Parse previous smallCaps/allCaps
        if (prevRPr['w:smallCaps']) {
          prevProps.smallCaps = parseOoxmlBoolean(prevRPr['w:smallCaps']);
        }
        if (prevRPr['w:caps']) {
          prevProps.allCaps = parseOoxmlBoolean(prevRPr['w:caps']);
        }

        // === Extended run property parsing per ECMA-376 Part 1 §17.3.2 ===

        // Parse double strikethrough (w:dstrike)
        if (prevRPr['w:dstrike']) {
          prevProps.dstrike = parseOoxmlBoolean(prevRPr['w:dstrike']);
        }

        // Parse text effects (w:outline, w:shadow, w:emboss, w:imprint)
        if (prevRPr['w:outline']) {
          prevProps.outline = parseOoxmlBoolean(prevRPr['w:outline']);
        }
        if (prevRPr['w:shadow']) {
          prevProps.shadow = parseOoxmlBoolean(prevRPr['w:shadow']);
        }
        if (prevRPr['w:emboss']) {
          prevProps.emboss = parseOoxmlBoolean(prevRPr['w:emboss']);
        }
        if (prevRPr['w:imprint']) {
          prevProps.imprint = parseOoxmlBoolean(prevRPr['w:imprint']);
        }

        // Parse vanish/hidden text (w:vanish, w:specVanish)
        if (prevRPr['w:vanish']) {
          prevProps.vanish = parseOoxmlBoolean(prevRPr['w:vanish']);
        }
        if (prevRPr['w:specVanish']) {
          prevProps.specVanish = parseOoxmlBoolean(prevRPr['w:specVanish']);
        }

        // Parse web hidden (w:webHidden) per ECMA-376 Part 1 §17.3.2.44
        if (prevRPr['w:webHidden']) {
          prevProps.webHidden = parseOoxmlBoolean(prevRPr['w:webHidden']);
        }

        // Parse RTL and no-proofing (w:rtl, w:noProof)
        if (prevRPr['w:rtl']) {
          prevProps.rtl = parseOoxmlBoolean(prevRPr['w:rtl']);
        }
        if (prevRPr['w:noProof']) {
          prevProps.noProof = parseOoxmlBoolean(prevRPr['w:noProof']);
        }

        // Parse snap to grid (w:snapToGrid)
        if (prevRPr['w:snapToGrid']) {
          prevProps.snapToGrid = parseOoxmlBoolean(prevRPr['w:snapToGrid']);
        }

        // Parse complex script bold/italic (w:bCs, w:iCs)
        if (prevRPr['w:bCs']) {
          prevProps.complexScriptBold = parseOoxmlBoolean(prevRPr['w:bCs']);
        }
        if (prevRPr['w:iCs']) {
          prevProps.complexScriptItalic = parseOoxmlBoolean(prevRPr['w:iCs']);
        }

        // Parse complex script flag (w:cs) per ECMA-376 Part 1 §17.3.2.7
        if (prevRPr['w:cs']) {
          prevProps.complexScript = parseOoxmlBoolean(prevRPr['w:cs']);
        }

        // Parse character spacing (w:spacing @w:val in twips)
        if (prevRPr['w:spacing']) {
          const spacingVal = prevRPr['w:spacing']['@_w:val'];
          if (spacingVal !== undefined) {
            prevProps.characterSpacing = safeParseInt(spacingVal);
          }
        }

        // Parse horizontal scaling (w:w @w:val as percentage)
        if (prevRPr['w:w']) {
          const scaleVal = prevRPr['w:w']['@_w:val'];
          if (scaleVal !== undefined) {
            prevProps.scaling = safeParseInt(scaleVal);
          }
        }

        // Parse vertical position (w:position @w:val in half-points)
        if (prevRPr['w:position']) {
          const posVal = prevRPr['w:position']['@_w:val'];
          if (posVal !== undefined) {
            prevProps.position = safeParseInt(posVal);
          }
        }

        // Parse kerning threshold (w:kern @w:val in half-points)
        if (prevRPr['w:kern']) {
          const kernVal = prevRPr['w:kern']['@_w:val'];
          if (kernVal !== undefined) {
            prevProps.kerning = safeParseInt(kernVal);
          }
        }

        // Parse language (w:lang) per ECMA-376 CT_Language (w:val, w:eastAsia, w:bidi)
        if (prevRPr['w:lang']) {
          const langObj = prevRPr['w:lang'];
          const langVal = langObj['@_w:val'];
          const langEastAsia = langObj['@_w:eastAsia'];
          const langBidi = langObj['@_w:bidi'];
          if (langEastAsia || langBidi) {
            prevProps.language = {
              val: langVal ? String(langVal) : undefined,
              eastAsia: langEastAsia ? String(langEastAsia) : undefined,
              bidi: langBidi ? String(langBidi) : undefined,
            };
          } else if (langVal) {
            prevProps.language = String(langVal);
          }
        }

        // Parse character style reference (w:rStyle @w:val)
        if (prevRPr['w:rStyle']) {
          const styleVal = prevRPr['w:rStyle']['@_w:val'];
          if (styleVal) {
            prevProps.characterStyle = String(styleVal);
          }
        }

        // Parse text effect/animation (w:effect @w:val)
        if (prevRPr['w:effect']) {
          const effectVal = prevRPr['w:effect']['@_w:val'];
          if (effectVal) {
            prevProps.effect = effectVal as RunFormatting['effect'];
          }
        }

        // Parse fit text width (w:fitText @w:val in twips)
        if (prevRPr['w:fitText']) {
          const fitVal = prevRPr['w:fitText']['@_w:val'];
          if (fitVal !== undefined) {
            prevProps.fitText = safeParseInt(fitVal);
          }
        }

        // Parse emphasis mark (w:em @w:val)
        if (prevRPr['w:em']) {
          const emVal = prevRPr['w:em']['@_w:val'];
          if (emVal) {
            prevProps.emphasis = emVal as RunFormatting['emphasis'];
          }
        }

        // Parse text border (w:bdr) per ECMA-376 Part 1 §17.3.2.5 — full
        // CT_Border attribute set for rPrChange previous-properties fidelity.
        if (prevRPr['w:bdr']) {
          const bdrObj = prevRPr['w:bdr'];
          const tb: import('../elements/Run').TextBorder = {
            style: bdrObj['@_w:val'] as import('../elements/Run').TextBorderStyle,
            size: bdrObj['@_w:sz'] !== undefined ? safeParseInt(bdrObj['@_w:sz']) : undefined,
            space:
              bdrObj['@_w:space'] !== undefined ? safeParseInt(bdrObj['@_w:space']) : undefined,
            color: bdrObj['@_w:color'],
          };
          // String(...) cast: XMLParser coerces "80"/"50" hex to numbers
          // — preserve the declared string contract on the model.
          if (bdrObj['@_w:themeColor']) {
            tb.themeColor = String(
              bdrObj['@_w:themeColor']
            ) as import('../elements/Run').ThemeColorValue;
          }
          if (bdrObj['@_w:themeTint']) tb.themeTint = String(bdrObj['@_w:themeTint']);
          if (bdrObj['@_w:themeShade']) tb.themeShade = String(bdrObj['@_w:themeShade']);
          if (bdrObj['@_w:shadow'] !== undefined) {
            tb.shadow = parseOnOffAttribute(String(bdrObj['@_w:shadow']), true);
          }
          if (bdrObj['@_w:frame'] !== undefined) {
            tb.frame = parseOnOffAttribute(String(bdrObj['@_w:frame']), true);
          }
          prevProps.border = tb;
        }

        // Parse character shading (w:shd) per ECMA-376 Part 1 §17.3.2.32
        if (prevRPr['w:shd']) {
          const shading = this.parseShadingFromObj(prevRPr['w:shd']);
          if (shading) {
            prevProps.shading = shading;
          }
        }

        // Parse East Asian layout (w:eastAsianLayout) per ECMA-376 Part 1
        // §17.3.2.10 CT_EastAsianLayout. Parity-fix with the main rPr
        // parser: route the three ST_OnOff attributes through
        // parseOnOffAttribute so every literal — including "0"
        // (explicit-false override) and "off" — resolves correctly. The
        // previous truthy gate both dropped explicit-false (XMLParser
        // coerces "0" to number 0 → falsy → undefined) and wrongly
        // coerced "off" to true.
        if (prevRPr['w:eastAsianLayout']) {
          const eaObj = prevRPr['w:eastAsianLayout'];
          prevProps.eastAsianLayout = {
            id: eaObj['@_w:id'] !== undefined ? safeParseInt(eaObj['@_w:id']) : undefined,
            combine:
              eaObj['@_w:combine'] !== undefined
                ? parseOnOffAttribute(String(eaObj['@_w:combine']), true)
                : undefined,
            combineBrackets: eaObj['@_w:combineBrackets'],
            vert:
              eaObj['@_w:vert'] !== undefined
                ? parseOnOffAttribute(String(eaObj['@_w:vert']), true)
                : undefined,
            vertCompress:
              eaObj['@_w:vertCompress'] !== undefined
                ? parseOnOffAttribute(String(eaObj['@_w:vertCompress']), true)
                : undefined,
          };
        }

        // Collect w14: namespace elements from the previous rPr for
        // passthrough (Word 2010+ text effects: w14:textOutline,
        // w14:shadow, w14:reflection, w14:glow, w14:ligatures,
        // w14:numForm, w14:numSpacing, w14:cntxtAlts, w14:stylisticSets).
        // The main rPr parser already collects these and the rPrChange
        // emitter (via generateRunPropertiesXML line 3130) re-emits
        // prevProps.rawW14Properties, but the rPrChange parser never
        // captured them — so tracked changes to any w14 text effect
        // silently lost the previous state on load → save.
        const prevRawW14: string[] = [];
        for (const key of Object.keys(prevRPr)) {
          if (key.startsWith('w14:')) {
            const rawXml = this.objectToXml({ [key]: prevRPr[key] });
            if (rawXml) prevRawW14.push(rawXml);
          }
        }
        if (prevRawW14.length > 0) {
          (prevProps as { rawW14Properties?: string[] }).rawW14Properties = prevRawW14;
        }

        propChange.previousProperties = prevProps;
      }

      run.setPropertyChangeRevision(propChange);
    }
  }

  private async parseDrawingFromObject(
    drawingObj: any,
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager,
    imageManager: ImageManager
  ): Promise<ImageRun | null> {
    const logger = getLogger();
    try {
      // Drawing can contain either wp:inline (inline image) or wp:anchor (floating image)
      const inlineObj = drawingObj['wp:inline'];
      const anchorObj = drawingObj['wp:anchor'];
      const imageObj = inlineObj || anchorObj;

      if (!imageObj) {
        logger.debug('Drawing found but no wp:inline or wp:anchor element');
        return null;
      }

      const isFloating = !!anchorObj;

      // Extract dimensions from wp:extent
      const extentObj = imageObj['wp:extent'];
      let width = 0;
      let height = 0;
      if (extentObj) {
        width = parseInt(extentObj['@_cx'] || '0', 10);
        height = parseInt(extentObj['@_cy'] || '0', 10);
      }

      // Extract effect extent
      const effectExtentObj = imageObj['wp:effectExtent'];
      let effectExtent = undefined;
      if (effectExtentObj) {
        effectExtent = {
          left: parseInt(effectExtentObj['@_l'] || '0', 10),
          top: parseInt(effectExtentObj['@_t'] || '0', 10),
          right: parseInt(effectExtentObj['@_r'] || '0', 10),
          bottom: parseInt(effectExtentObj['@_b'] || '0', 10),
        };
      }

      // --- Group A: Parse inline dist attributes ---
      let inlineDistT = 0;
      let inlineDistB = 0;
      let inlineDistL = 0;
      let inlineDistR = 0;
      if (inlineObj) {
        const toInt = (v: any) =>
          v !== undefined && v !== null ? parseInt(String(v), 10) || 0 : 0;
        inlineDistT = toInt(inlineObj['@_distT']);
        inlineDistB = toInt(inlineObj['@_distB']);
        inlineDistL = toInt(inlineObj['@_distL']);
        inlineDistR = toInt(inlineObj['@_distR']);
      }

      // Extract name, description, title, hidden, and ID from wp:docPr
      const docPrObj = imageObj['wp:docPr'];
      let name = 'image';
      let description = 'Image';
      let title: string | undefined = undefined;
      let docPrId = 1;
      let hidden = false;
      if (docPrObj) {
        // wp:docPr @name and @descr are xsd:string per ECMA-376
        // §20.4.2.5 CT_NonVisualDrawingProps. XMLParser coerces
        // purely-numeric values ("2010") to JS numbers; cast through
        // String() so Image.name / Image.description keep the declared
        // string contract (matches the @_title handling below).
        const rawName = docPrObj['@_name'];
        name = rawName !== undefined && rawName !== null ? String(rawName) : 'image';
        const rawDescr = docPrObj['@_descr'];
        description = rawDescr !== undefined && rawDescr !== null ? String(rawDescr) : 'Image';
        if (docPrObj['@_title']) {
          title = String(docPrObj['@_title']);
        }
        const idAttr = docPrObj['@_id'];
        if (idAttr) {
          docPrId = parseInt(String(idAttr), 10);
        }
        // Group A: hidden attribute
        if (
          docPrObj['@_hidden'] === '1' ||
          docPrObj['@_hidden'] === 1 ||
          docPrObj['@_hidden'] === true
        ) {
          hidden = true;
        }
      }

      // --- Group A: Parse noChangeAspect from wp:cNvGraphicFramePr ---
      let noChangeAspect = true; // default
      const cNvGfPrObj = imageObj['wp:cNvGraphicFramePr'];
      if (cNvGfPrObj) {
        const gfLocks = cNvGfPrObj['a:graphicFrameLocks'];
        if (gfLocks) {
          const val = gfLocks['@_noChangeAspect'];
          noChangeAspect = val === '1' || val === 1 || val === true;
        }
      }

      // Parse wrap settings (for floating images)
      let wrap = undefined;
      if (isFloating) {
        wrap = this.parseWrapSettings(anchorObj);
      }

      // Parse position (for floating images)
      let position = undefined;
      if (isFloating) {
        position = this.parseImagePosition(anchorObj);
      }

      // Parse anchor configuration (for floating images)
      let anchor = undefined;
      if (isFloating && anchorObj) {
        const toBool = (val: any) => val === '1' || val === 1 || val === true;
        const toOptInt = (val: any): number | undefined => {
          if (val === undefined || val === null) return undefined;
          const n = parseInt(String(val), 10);
          return isNaN(n) ? undefined : n;
        };

        anchor = {
          behindDoc: toBool(anchorObj['@_behindDoc']),
          locked: toBool(anchorObj['@_locked']),
          layoutInCell: toBool(anchorObj['@_layoutInCell']),
          allowOverlap: toBool(anchorObj['@_allowOverlap']),
          relativeHeight: parseInt(anchorObj['@_relativeHeight'] || '251658240', 10),
          simplePos: toBool(anchorObj['@_simplePos']),
          distT: toOptInt(anchorObj['@_distT']),
          distB: toOptInt(anchorObj['@_distB']),
          distL: toOptInt(anchorObj['@_distL']),
          distR: toOptInt(anchorObj['@_distR']),
        };

        // Also check anchor hidden attribute
        if (
          anchorObj['@_hidden'] === '1' ||
          anchorObj['@_hidden'] === 1 ||
          anchorObj['@_hidden'] === true
        ) {
          hidden = true;
        }
      }

      // Navigate through the graphic structure to find the relationship ID
      const graphicObj = imageObj['a:graphic'];
      if (!graphicObj) return null;
      const graphicDataObj = graphicObj['a:graphicData'];
      if (!graphicDataObj) return null;
      const picPicObj = graphicDataObj['pic:pic'];
      if (!picPicObj) return null;
      const blipFillObj = picPicObj['pic:blipFill'];
      if (!blipFillObj) return null;
      const blipObj = blipFillObj['a:blip'];
      if (!blipObj) return null;

      // Parse crop settings
      const crop = this.parseImageCrop(blipFillObj);

      // Parse effects (includes transparency via a:alphaModFix)
      const effects = this.parseImageEffects(blipObj);

      // --- Group A: Parse blip attributes ---
      const compressionState = blipObj['@_cstate'] || 'none';

      // --- Group A: Parse blipFill attributes ---
      const blipFillDpi =
        blipFillObj['@_dpi'] !== undefined ? parseInt(String(blipFillObj['@_dpi']), 10) : undefined;
      let blipFillRotWithShape: boolean | undefined;
      if (blipFillObj['@_rotWithShape'] !== undefined) {
        const rws = blipFillObj['@_rotWithShape'];
        blipFillRotWithShape = rws === '1' || rws === 1 || rws === true;
      }

      // --- Group A: Parse pic:nvPicPr (non-visual properties and locks) ---
      const nvPicPrObj = picPicObj['pic:nvPicPr'];
      let picNonVisualProps = { id: '0', name: '', descr: '' };
      let picLocks: Record<string, boolean> = { noChangeAspect: true, noChangeArrowheads: true };
      if (nvPicPrObj) {
        const cNvPr = nvPicPrObj['pic:cNvPr'];
        if (cNvPr) {
          picNonVisualProps = {
            id: String(cNvPr['@_id'] ?? '0'),
            name: String(cNvPr['@_name'] ?? ''),
            descr: String(cNvPr['@_descr'] ?? ''),
          };
        }
        const cNvPicPr = nvPicPrObj['pic:cNvPicPr'];
        if (cNvPicPr) {
          const locks = cNvPicPr['a:picLocks'];
          if (locks) {
            picLocks = {};
            const lockAttrs = [
              'noChangeAspect',
              'noChangeArrowheads',
              'noSelect',
              'noMove',
              'noResize',
              'noEditPoints',
              'noAdjustHandles',
              'noRot',
              'noChangeShapeType',
              'noCrop',
              'noGrp',
            ];
            for (const attr of lockAttrs) {
              const val = locks[`@_${attr}`];
              if (val === '1' || val === 1 || val === true) {
                picLocks[attr] = true;
              }
            }
          }
        }
      }

      // --- Group A & C: Parse pic:spPr (shape properties) ---
      const spPrObj = picPicObj['pic:spPr'];
      let border: any = undefined;
      let zeroWidthLnXml: string | null = null;
      let hasSpPrNoFill = false;
      let rotation = 0;
      let flipH = false;
      let flipV = false;
      let presetGeometry = 'rect';
      let bwMode = 'auto';
      if (spPrObj) {
        // Group A: bwMode
        if (spPrObj['@_bwMode']) {
          bwMode = String(spPrObj['@_bwMode']);
        }

        // Group A: presetGeometry
        const prstGeom = spPrObj['a:prstGeom'];
        if (prstGeom?.['@_prst']) {
          presetGeometry = String(prstGeom['@_prst']);
        }

        // Group C: Enhanced border parsing (a:ln)
        const lnObj = spPrObj['a:ln'];
        if (lnObj) {
          const widthEmu = parseInt(lnObj['@_w'] || '0', 10);
          if (widthEmu > 0) {
            border = { width: widthEmu / 12700, _fromParsed: true } as any;
            // Parse additional a:ln attributes
            if (lnObj['@_cap']) border.cap = String(lnObj['@_cap']);
            if (lnObj['@_cmpd']) border.compound = String(lnObj['@_cmpd']);
            if (lnObj['@_algn']) border.alignment = String(lnObj['@_algn']);
            // Parse fill
            const solidFill = lnObj['a:solidFill'];
            if (solidFill) {
              const srgbClr = solidFill['a:srgbClr'];
              const schemeClr = solidFill['a:schemeClr'];
              if (srgbClr) {
                border.fill = { type: 'srgbClr', value: String(srgbClr['@_val'] || '') };
                border.fill.modifiers = this.parseColorModifiers(srgbClr);
              } else if (schemeClr) {
                border.fill = { type: 'schemeClr', value: String(schemeClr['@_val'] || '') };
                border.fill.modifiers = this.parseColorModifiers(schemeClr);
              }
            } else {
              // Non-solid fill — store as raw XML passthrough
              const fillKeys = ['a:gradFill', 'a:pattFill', 'a:noFill'];
              for (const fk of fillKeys) {
                if (lnObj[fk]) {
                  border.rawFillXml = this.objectToXml({ [fk]: lnObj[fk] });
                  break;
                }
              }
            }
            // Parse dash pattern
            const prstDash = lnObj['a:prstDash'];
            if (prstDash?.['@_val']) border.dashPattern = String(prstDash['@_val']);
            // Parse join
            if (lnObj['a:round']) border.join = 'round';
            else if (lnObj['a:bevel']) border.join = 'bevel';
            else if (lnObj['a:miter']) {
              border.join = 'miter';
              if (lnObj['a:miter']['@_lim'])
                border.miterLimit = parseInt(String(lnObj['a:miter']['@_lim']), 10);
            }
            // Parse head/tail end
            if (lnObj['a:headEnd']) {
              border.headEnd = {};
              if (lnObj['a:headEnd']['@_type'])
                border.headEnd.type = String(lnObj['a:headEnd']['@_type']);
              if (lnObj['a:headEnd']['@_w'])
                border.headEnd.width = String(lnObj['a:headEnd']['@_w']);
              if (lnObj['a:headEnd']['@_len'])
                border.headEnd.length = String(lnObj['a:headEnd']['@_len']);
            }
            if (lnObj['a:tailEnd']) {
              border.tailEnd = {};
              if (lnObj['a:tailEnd']['@_type'])
                border.tailEnd.type = String(lnObj['a:tailEnd']['@_type']);
              if (lnObj['a:tailEnd']['@_w'])
                border.tailEnd.width = String(lnObj['a:tailEnd']['@_w']);
              if (lnObj['a:tailEnd']['@_len'])
                border.tailEnd.length = String(lnObj['a:tailEnd']['@_len']);
            }
          } else {
            // Zero-width or absent-width a:ln: preserve as raw XML (BUG 8 fix)
            zeroWidthLnXml = this.objectToXml({ 'a:ln': lnObj });
          }
        }

        // Detect spPr-level a:noFill (preserve independently from border)
        if (spPrObj['a:noFill']) {
          hasSpPrNoFill = true;
        }

        // Parse rotation and flip from a:xfrm
        const xfrmObj = spPrObj['a:xfrm'];
        if (xfrmObj?.['@_rot']) {
          rotation = parseInt(xfrmObj['@_rot'], 10) / 60000;
        }
        if (
          xfrmObj?.['@_flipH'] === '1' ||
          xfrmObj?.['@_flipH'] === 1 ||
          xfrmObj?.['@_flipH'] === true
        ) {
          flipH = true;
        }
        if (
          xfrmObj?.['@_flipV'] === '1' ||
          xfrmObj?.['@_flipV'] === 1 ||
          xfrmObj?.['@_flipV'] === true
        ) {
          flipV = true;
        }
      }

      // --- Linked vs embedded image (r:embed or r:link) ---
      let relationshipId = blipObj['@_r:embed'];
      let isLinked = false;
      if (!relationshipId) {
        relationshipId = blipObj['@_r:link'];
        if (relationshipId) {
          isLinked = true;
        } else {
          logger.debug('Drawing has no r:embed or r:link relationship ID');
          return null;
        }
      }

      // --- SVG detection: asvg:svgBlip in a:extLst ---
      let svgRelationshipId: string | undefined;
      const blipExtLst = blipObj['a:extLst'];
      if (blipExtLst) {
        const exts = Array.isArray(blipExtLst['a:ext'])
          ? blipExtLst['a:ext']
          : blipExtLst['a:ext']
            ? [blipExtLst['a:ext']]
            : [];
        for (const ext of exts) {
          const svgBlip = ext?.['asvg:svgBlip'];
          if (svgBlip?.['@_r:embed']) {
            svgRelationshipId = String(svgBlip['@_r:embed']);
          }
        }
      }

      // Get the image from the relationship
      const partContext = this.currentPartName ? ` (part: ${this.currentPartName})` : '';

      let imageData: Buffer;
      let imagePath: string;
      let originalFilename: string | undefined;

      if (isLinked) {
        // Linked image: no data in package, create empty buffer
        imageData = Buffer.alloc(0);
        imagePath = '';
        originalFilename = undefined;
      } else {
        const relationship = relationshipManager.getRelationship(relationshipId);
        if (!relationship) {
          logger.warn(`Image relationship not found: ${relationshipId}${partContext}`);
          return null;
        }
        const imageTarget = relationship.getTarget();
        if (!imageTarget) {
          logger.warn(`Image relationship has no target: ${relationshipId}${partContext}`);
          return null;
        }
        imagePath = `word/${imageTarget}`;
        const data = zipHandler.getFileAsBuffer(imagePath);
        if (!data) {
          logger.warn(`Image file not found in ZIP: ${imagePath}${partContext}`);
          return null;
        }
        imageData = data;
        originalFilename = imageTarget.split('/').pop();
      }

      logger.debug(
        `Parsing image: ${imagePath || '(linked)'}, relId: ${relationshipId}${partContext}`
      );

      // Create image from buffer with all properties
      const { Image: ImageClass } = await import('../elements/Image');
      const image = await ImageClass.create({
        source: imageData,
        width,
        height,
        name,
        description,
        title,
        effectExtent,
        wrap,
        position,
        anchor,
        crop,
        effects,
        border,
        rotation,
        flipH,
        flipV,
        presetGeometry,
        compressionState,
        bwMode,
        inlineDistT,
        inlineDistB,
        inlineDistL,
        inlineDistR,
        noChangeAspect,
        hidden,
        blipFillDpi,
        blipFillRotWithShape,
        picLocks,
        picNonVisualProps,
        isLinked,
        svgRelationshipId,
      });

      // Preserve zero-width a:ln as raw passthrough (BUG 8 fix)
      if (zeroWidthLnXml) {
        image._setRawPassthrough('zero-width-ln', zeroWidthLnXml);
      }

      // Preserve spPr-level a:noFill independently from border (Bug C fix)
      if (hasSpPrNoFill) {
        image._setRawPassthrough('spPr-noFill', '<a:noFill/>');
      }

      // --- Group B: Collect raw passthrough for unmodeled XML subtrees ---
      // Blip effects (children of a:blip that aren't modeled)
      const blipEffectsRaw = this.collectUnmodeledChildren(blipObj, [
        'a:lum',
        'a:grayscl',
        'a:alphaModFix',
        'a:extLst',
      ]);
      if (blipEffectsRaw) image._setRawPassthrough('blip-effects', blipEffectsRaw);

      // Blip extLst (must come last per schema)
      if (blipExtLst && !svgRelationshipId) {
        // Only pass through if we're not already handling SVG
        const extLstXml = this.objectToXml({ 'a:extLst': blipExtLst });
        if (extLstXml) image._setRawPassthrough('blip-extLst', extLstXml);
      } else if (blipExtLst && svgRelationshipId) {
        // Preserve the full extLst including SVG ref
        const extLstXml = this.objectToXml({ 'a:extLst': blipExtLst });
        if (extLstXml) image._setRawPassthrough('blip-extLst', extLstXml);
      }

      // BlipFill extras (a:tile instead of a:stretch)
      if (blipFillObj['a:tile']) {
        const tileXml = this.objectToXml({ 'a:tile': blipFillObj['a:tile'] });
        if (tileXml) image._setRawPassthrough('blipFill-extra', tileXml);
      }

      // Geometry passthrough (a:custGeom, or prstGeom with non-empty avLst)
      if (spPrObj) {
        if (spPrObj['a:custGeom']) {
          const geomXml = this.objectToXml({ 'a:custGeom': spPrObj['a:custGeom'] });
          if (geomXml) image._setRawPassthrough('geometry', geomXml);
        } else if (spPrObj['a:prstGeom']) {
          const prstGeom = spPrObj['a:prstGeom'];
          const avLst = prstGeom['a:avLst'];
          // Check if avLst has actual children (not just empty element)
          if (
            avLst &&
            typeof avLst === 'object' &&
            Object.keys(avLst).some((k) => !k.startsWith('@_') && k !== '_orderedChildren')
          ) {
            const geomXml = this.objectToXml({ 'a:prstGeom': prstGeom });
            if (geomXml) image._setRawPassthrough('geometry', geomXml);
          }
        }

        // spPr effects passthrough (effectLst, effectDag, scene3d, sp3d, extLst)
        const spPrEffects = this.collectUnmodeledChildren(spPrObj, [
          'a:xfrm',
          'a:prstGeom',
          'a:custGeom',
          'a:noFill',
          'a:solidFill',
          'a:ln',
        ]);
        if (spPrEffects) image._setRawPassthrough('spPr-effects', spPrEffects);
      }

      // Wrap polygon passthrough
      if (isFloating) {
        const wrapTypes = ['wp:wrapTight', 'wp:wrapThrough'];
        for (const wt of wrapTypes) {
          const wrapObj = anchorObj[wt];
          if (wrapObj?.['wp:wrapPolygon']) {
            const polyXml = this.objectToXml({ 'wp:wrapPolygon': wrapObj['wp:wrapPolygon'] });
            if (polyXml) image._setRawPassthrough('wrap-polygon', polyXml);
            break;
          }
        }

        // Anchor extras (wp14:sizeRelH, wp14:sizeRelV)
        const anchorExtras = this.collectUnmodeledChildren(anchorObj, [
          'wp:simplePos',
          'wp:positionH',
          'wp:positionV',
          'wp:extent',
          'wp:effectExtent',
          'wp:wrapSquare',
          'wp:wrapTight',
          'wp:wrapThrough',
          'wp:wrapTopAndBottom',
          'wp:wrapNone',
          'wp:docPr',
          'wp:cNvGraphicFramePr',
          'a:graphic',
        ]);
        if (anchorExtras) image._setRawPassthrough('anchor-extra', anchorExtras);
      }

      // DocPr extras (a:hlinkClick, a:hlinkHover, a:extLst)
      if (docPrObj) {
        const docPrExtras = this.collectUnmodeledChildren(docPrObj, []);
        if (docPrExtras) image._setRawPassthrough('docPr-extra', docPrExtras);
      }

      // cNvPr extras from pic:nvPicPr > pic:cNvPr
      if (nvPicPrObj?.['pic:cNvPr']) {
        const cNvPrExtras = this.collectUnmodeledChildren(nvPicPrObj['pic:cNvPr'], []);
        if (cNvPrExtras) image._setRawPassthrough('cNvPr-extra', cNvPrExtras);
      }

      // Register image
      if (!isLinked) {
        imageManager.registerImage(image, relationshipId, originalFilename, this.currentPartName);
      }
      image.setRelationshipId(relationshipId);
      image.setDocPrId(docPrId);

      logger.debug(
        `Image registered: ${originalFilename || '(linked)'}, relId: ${relationshipId}${partContext}`
      );

      return new ImageRun(image);
    } catch (error: unknown) {
      const partContext = this.currentPartName ? ` (part: ${this.currentPartName})` : '';
      logger.warn(
        `Failed to parse drawing${partContext}:`,
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Parses text wrap settings from anchor object
   * @private
   */
  private parseWrapSettings(anchorObj: any): any {
    // Check for different wrap types
    const wrapSquare = anchorObj['wp:wrapSquare'];
    const wrapTight = anchorObj['wp:wrapTight'];
    const wrapThrough = anchorObj['wp:wrapThrough'];
    const wrapTopBottom = anchorObj['wp:wrapTopAndBottom'];
    const wrapNone = anchorObj['wp:wrapNone'];

    const wrapObj = wrapSquare || wrapTight || wrapThrough || wrapTopBottom || wrapNone;
    if (!wrapObj) {
      return undefined;
    }

    // Determine wrap type
    let type: any = 'square';
    if (wrapTight) type = 'tight';
    else if (wrapThrough) type = 'through';
    else if (wrapTopBottom) type = 'topAndBottom';
    else if (wrapNone) type = 'none';

    return {
      type,
      side: wrapObj['@_wrapText'] || 'bothSides',
      // Distance attributes are on the wrap element, not the anchor
      distanceTop: wrapObj['@_distT'] ? parseInt(wrapObj['@_distT'], 10) : undefined,
      distanceBottom: wrapObj['@_distB'] ? parseInt(wrapObj['@_distB'], 10) : undefined,
      distanceLeft: wrapObj['@_distL'] ? parseInt(wrapObj['@_distL'], 10) : undefined,
      distanceRight: wrapObj['@_distR'] ? parseInt(wrapObj['@_distR'], 10) : undefined,
    };
  }

  /**
   * Parses image position from anchor object
   * @private
   */
  private parseImagePosition(anchorObj: any): any {
    const posH = anchorObj['wp:positionH'];
    const posV = anchorObj['wp:positionV'];

    if (!posH || !posV) {
      return undefined;
    }

    // Parse horizontal position
    const horizontal: any = {
      anchor: posH['@_relativeFrom'] || 'page',
    };

    if (posH['wp:posOffset']) {
      const offsetText = Array.isArray(posH['wp:posOffset'])
        ? posH['wp:posOffset'][0]
        : posH['wp:posOffset'];
      horizontal.offset = parseInt(
        typeof offsetText === 'string' ? offsetText : offsetText?.['#text'] || '0',
        10
      );
    } else if (posH['wp:align']) {
      const alignText = Array.isArray(posH['wp:align']) ? posH['wp:align'][0] : posH['wp:align'];
      horizontal.alignment = typeof alignText === 'string' ? alignText : alignText?.['#text'];
    }

    // Parse vertical position
    const vertical: any = {
      anchor: posV['@_relativeFrom'] || 'page',
    };

    if (posV['wp:posOffset']) {
      const offsetText = Array.isArray(posV['wp:posOffset'])
        ? posV['wp:posOffset'][0]
        : posV['wp:posOffset'];
      vertical.offset = parseInt(
        typeof offsetText === 'string' ? offsetText : offsetText?.['#text'] || '0',
        10
      );
    } else if (posV['wp:align']) {
      const alignText = Array.isArray(posV['wp:align']) ? posV['wp:align'][0] : posV['wp:align'];
      vertical.alignment = typeof alignText === 'string' ? alignText : alignText?.['#text'];
    }

    return { horizontal, vertical };
  }

  /**
   * Parses image crop from blip object
   * @private
   */
  private parseImageCrop(blipObj: any): any {
    const srcRect = blipObj['a:srcRect'];
    if (!srcRect) {
      return undefined;
    }

    return {
      left: parseInt(srcRect['@_l'] || '0', 10) / 1000, // Convert from per-mille to percentage
      top: parseInt(srcRect['@_t'] || '0', 10) / 1000,
      right: parseInt(srcRect['@_r'] || '0', 10) / 1000,
      bottom: parseInt(srcRect['@_b'] || '0', 10) / 1000,
    };
  }

  /**
   * Parses image effects from blip object
   * Per ECMA-376, effects are direct children of a:blip (not in a:extLst)
   * @private
   */
  private parseImageEffects(blipObj: any): any {
    const effects: any = {};

    // Look for lum (luminance) element for brightness/contrast
    // Per ECMA-376: a:lum is a direct child of a:blip with bright/contrast attributes
    const lum = blipObj['a:lum'];
    if (lum) {
      if (lum['@_bright']) {
        effects.brightness = parseInt(lum['@_bright'], 10) / 1000;
      }
      if (lum['@_contrast']) {
        effects.contrast = parseInt(lum['@_contrast'], 10) / 1000;
      }
    }

    // Check for grayscale (direct child of a:blip)
    if (blipObj['a:grayscl']) {
      effects.grayscale = true;
    }

    // Parse transparency via a:alphaModFix (ECMA-376 §20.1.8.4)
    const alphaModFix = blipObj['a:alphaModFix'];
    if (alphaModFix?.['@_amt'] !== undefined) {
      // amt is in 1/1000ths of percent (e.g., 50000 = 50% opacity = 50% transparency)
      const amt = parseInt(String(alphaModFix['@_amt']), 10);
      const transparency = Math.round(100 - amt / 1000);
      if (transparency > 0) {
        effects.transparency = transparency;
      }
    }

    return Object.keys(effects).length > 0 ? effects : undefined;
  }

  /**
   * Collects unmodeled children from a parsed XML object as raw XML strings.
   * Used for Group B passthrough — preserves XML subtrees that the framework
   * doesn't model, preventing data loss during round-trip.
   * @private
   */
  private collectUnmodeledChildren(parentObj: any, modeledKeys: string[]): string | null {
    if (!parentObj || typeof parentObj !== 'object') return null;
    const parts: string[] = [];
    for (const key of Object.keys(parentObj)) {
      if (key.startsWith('@_') || key === '#text' || key === '_orderedChildren') continue;
      if (modeledKeys.includes(key)) continue;
      const child = parentObj[key];
      if (Array.isArray(child)) {
        for (const item of child) {
          parts.push(this.objectToXml({ [key]: item }));
        }
      } else {
        parts.push(this.objectToXml({ [key]: child }));
      }
    }
    return parts.length > 0 ? parts.join('') : null;
  }

  /**
   * Parses color modifier children from a color element (a:srgbClr or a:schemeClr).
   * Returns array of { name, val } for modifiers like lumMod, lumOff, satMod, etc.
   * @private
   */
  private parseColorModifiers(colorObj: any): { name: string; val: string }[] | undefined {
    if (!colorObj || typeof colorObj !== 'object') return undefined;
    const modifiers: { name: string; val: string }[] = [];
    const modNames = ['a:lumMod', 'a:lumOff', 'a:satMod', 'a:shade', 'a:tint', 'a:alpha'];
    for (const mod of modNames) {
      const modObj = colorObj[mod];
      if (modObj?.['@_val'] !== undefined) {
        modifiers.push({ name: mod.replace('a:', ''), val: String(modObj['@_val']) });
      }
    }
    return modifiers.length > 0 ? modifiers : undefined;
  }

  /**
   * Parse a single border element from a parsed XML object.
   * Shared by table border and cell border parsing.
   */
  private parseBorderElement(borderObj: any): TableBorder | undefined {
    if (!borderObj) return undefined;
    // Extract the full CT_Border attribute set per ECMA-376 §17.18.2:
    // val (required) / sz / space / color / themeColor / themeTint /
    // themeShade / shadow / frame. Previously the last five were silently
    // dropped on load, so themed borders and shadow/frame flags were lost
    // on every round-trip.
    const border: TableBorder = {
      style: (borderObj['@_w:val'] || 'single') as TableBorder['style'],
    };
    if (borderObj['@_w:sz'] !== undefined) border.size = safeParseInt(borderObj['@_w:sz']);
    if (borderObj['@_w:space'] !== undefined) border.space = safeParseInt(borderObj['@_w:space']);
    if (borderObj['@_w:color']) border.color = String(borderObj['@_w:color']);
    // String(...) cast: themeTint / themeShade are ST_UcharHexNumber
    // (2-char hex) declared as `string` on the model. XMLParser coerces
    // purely-digit hex like "80"/"50" to numbers — cast to preserve
    // the type contract.
    if (borderObj['@_w:themeColor']) {
      (border as any).themeColor = String(borderObj['@_w:themeColor']);
    }
    if (borderObj['@_w:themeTint']) {
      (border as any).themeTint = String(borderObj['@_w:themeTint']);
    }
    if (borderObj['@_w:themeShade']) {
      (border as any).themeShade = String(borderObj['@_w:themeShade']);
    }
    if (borderObj['@_w:shadow'] !== undefined) {
      (border as any).shadow = parseOnOffAttribute(String(borderObj['@_w:shadow']), true);
    }
    if (borderObj['@_w:frame'] !== undefined) {
      (border as any).frame = parseOnOffAttribute(String(borderObj['@_w:frame']), true);
    }
    return border;
  }

  private async parseTableFromObject(
    tableObj: any,
    relationshipManager: RelationshipManager,
    zipHandler: ZipHandler,
    imageManager: ImageManager,
    rawTableXml?: string
  ): Promise<Table | null> {
    try {
      // Create empty table
      const table = new Table();

      // Parse table properties (w:tblPr)
      if (tableObj['w:tblPr']) {
        this.parseTablePropertiesFromObject(tableObj['w:tblPr'], table);
      }

      // Parse table grid (w:tblGrid) - column widths
      if (tableObj['w:tblGrid']?.['w:gridCol']) {
        const gridCols = tableObj['w:tblGrid']['w:gridCol'];
        const gridColArray = Array.isArray(gridCols) ? gridCols : [gridCols];
        const widths = gridColArray.map((col: any) => {
          const w = col['@_w:w'];
          // Use isExplicitlySet to properly handle 0 values
          return isExplicitlySet(w) ? safeParseInt(w, 2880) : 2880; // default to 2 inches
        });
        if (widths.length > 0) {
          table.setTableGrid(widths);
        }
      }

      // Parse table grid change (w:tblGridChange) per ECMA-376 §17.13.5.35
      if (tableObj['w:tblGrid']?.['w:tblGridChange']) {
        const changeObj = tableObj['w:tblGrid']['w:tblGridChange'];
        const prevGridCols = changeObj['w:tblGrid']?.['w:gridCol'];
        if (prevGridCols) {
          const prevArray = Array.isArray(prevGridCols) ? prevGridCols : [prevGridCols];
          const prevWidths = prevArray.map((col: any) => ({
            width: isExplicitlySet(col['@_w:w']) ? safeParseInt(col['@_w:w'], 2880) : 2880,
          }));
          const gridChange = TableGridChange.create(
            safeParseInt(changeObj['@_w:id'], 0),
            prevWidths,
            changeObj['@_w:author'] !== undefined ? String(changeObj['@_w:author']) : undefined,
            changeObj['@_w:date'] !== undefined
              ? new Date(String(changeObj['@_w:date']))
              : undefined
          );
          table.setTblGridChange(gridChange);
        }
      }

      // Parse table rows (w:tr)
      const rows = tableObj['w:tr'];
      const rowChildren = Array.isArray(rows) ? rows : rows ? [rows] : [];

      // Extract row XMLs from raw table XML if available
      let rowXmls: string[] = [];
      if (rawTableXml) {
        rowXmls = XMLParser.extractElements(rawTableXml, 'w:tr');
      }

      // Track row positions in raw XML for bookmarkEnd extraction between rows
      const rowPositions: { start: number; end: number }[] = [];
      if (rawTableXml) {
        let searchPos = 0;
        for (const rowXml of rowXmls) {
          const rowStart = rawTableXml.indexOf(rowXml, searchPos);
          if (rowStart !== -1) {
            rowPositions.push({
              start: rowStart,
              end: rowStart + rowXml.length,
            });
            searchPos = rowStart + rowXml.length;
          }
        }
      }

      for (let i = 0; i < rowChildren.length; i++) {
        const rowObj = rowChildren[i];
        const rawRowXml = i < rowXmls.length ? rowXmls[i] : undefined;

        const row = await this.parseTableRowFromObject(
          rowObj,
          relationshipManager,
          zipHandler,
          imageManager,
          rawRowXml
        );
        if (row) {
          table.addRow(row);

          // Check for bookmarkEnd elements AFTER this row (between rows)
          if (rawTableXml && i < rowPositions.length) {
            const currentRowEnd = rowPositions[i]?.end || 0;
            const nextRowStart =
              i + 1 < rowPositions.length ? rowPositions[i + 1]?.start : rawTableXml.length;

            if (nextRowStart && currentRowEnd < nextRowStart) {
              const betweenContent = rawTableXml.slice(currentRowEnd, nextRowStart);
              const bookmarkEnds = this.extractBookmarkEndsFromContent(betweenContent);

              if (bookmarkEnds.length > 0) {
                // Attach to last paragraph in last cell of this row
                const cells = row.getCells();
                const lastCell = cells[cells.length - 1];
                if (lastCell) {
                  const paras = lastCell.getParagraphs();
                  const lastPara = paras[paras.length - 1];
                  if (lastPara) {
                    for (const bookmark of bookmarkEnds) {
                      lastPara.addBookmarkEnd(bookmark);
                    }
                  }
                }
              }
            }
          }
        }
      }

      // Note: StylesManager is injected by Document.parseDocument() after styles are loaded

      return table;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse table:',
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Parses table properties from XML object and applies to table
   */
  private parseTablePropertiesFromObject(tblPrObj: any, table: Table): void {
    if (!tblPrObj) return;

    // Parse table style reference (w:tblStyle) per ECMA-376 §17.7.4.62.
    // w:val is ST_String — cast through String() so purely-numeric
    // custom style IDs ("2025", "1", …) don't leak as JS numbers
    // through XMLParser's parseAttributeValue coercion into the
    // string-typed `formatting.style` field.
    if (tblPrObj['w:tblStyle']) {
      const styleId = tblPrObj['w:tblStyle']['@_w:val'];
      if (styleId !== undefined && styleId !== null && styleId !== '') {
        table.setStyle(String(styleId));
      }
    }

    // Parse table look flags (w:tblLook) per ECMA-376 §17.4.57 — supports both
    // hex-string format (w:val="04A0") AND individual ST_OnOff attributes
    // (firstRow/lastRow/firstColumn/lastColumn/noHBand/noVBand).
    //
    // XMLParser.parseToObject runs with `parseAttributeValue: true` by default,
    // so `"1"` coerces to the number `1` and `"true"` to the boolean `true`.
    // The previous `=== '1'` strict-string comparison missed both coerced
    // forms, silently flipping every individually-set flag to OFF and
    // producing `tblLook="0000"` for every Word-authored document whose
    // tblLook used the expanded attribute syntax. Route each attribute
    // through `parseOoxmlBoolean` (attribute form) so string/number/boolean
    // representations all resolve correctly.
    if (tblPrObj['w:tblLook']) {
      const look = tblPrObj['w:tblLook'];
      if (look['@_w:val']) {
        // Hex string format
        table.setTblLook(look['@_w:val']);
      } else {
        // Individual attribute format — construct hex value.
        // Bits per §17.4.57: firstRow=0x0020, lastRow=0x0040, firstCol=0x0080,
        // lastCol=0x0100, noHBand=0x0200, noVBand=0x0400.
        const attrIsOn = (name: string): boolean => {
          const v = look[name];
          if (v === undefined) return false;
          // parseOoxmlBoolean accepts the value wrapped as `{'@_w:val': v}` —
          // handles string "1"/"0"/"true"/"false"/"on"/"off", number 1/0,
          // and boolean true/false uniformly.
          return parseOoxmlBoolean({ '@_w:val': v });
        };
        let value = 0;
        if (attrIsOn('@_w:firstRow')) value |= 0x0020;
        if (attrIsOn('@_w:lastRow')) value |= 0x0040;
        if (attrIsOn('@_w:firstColumn')) value |= 0x0080;
        if (attrIsOn('@_w:lastColumn')) value |= 0x0100;
        if (attrIsOn('@_w:noHBand')) value |= 0x0200;
        if (attrIsOn('@_w:noVBand')) value |= 0x0400;
        table.setTblLook(value.toString(16).toUpperCase().padStart(4, '0'));
      }
    }

    // Parse table positioning (tblpPr) - for floating tables.
    // Per ECMA-376 §17.4.52 CT_TblPPr, the six numeric attributes
    // (tblpX/tblpY/leftFromText/rightFromText/topFromText/bottomFromText)
    // are ST_SignedTwipsMeasure / ST_TwipsMeasure where 0 is a valid
    // value (e.g. float table anchored exactly at the anchor point).
    // XMLParser coerces "0" to the number 0 (falsy), so the previous
    // truthy gate silently dropped zero-offset positions. Table's
    // emitter uses `!== undefined`, so the asymmetry lost zeroes on
    // round-trip. Route each numeric read through isExplicitlySet +
    // safeParseInt.
    if (tblPrObj['w:tblpPr']) {
      const tblpPr = tblPrObj['w:tblpPr'];
      const position: any = {};

      if (isExplicitlySet(tblpPr['@_w:tblpX'])) position.x = safeParseInt(tblpPr['@_w:tblpX']);
      if (isExplicitlySet(tblpPr['@_w:tblpY'])) position.y = safeParseInt(tblpPr['@_w:tblpY']);
      if (tblpPr['@_w:horzAnchor']) position.horizontalAnchor = tblpPr['@_w:horzAnchor'];
      if (tblpPr['@_w:vertAnchor']) position.verticalAnchor = tblpPr['@_w:vertAnchor'];
      if (tblpPr['@_w:tblpXSpec']) position.horizontalAlignment = tblpPr['@_w:tblpXSpec'];
      if (tblpPr['@_w:tblpYSpec']) position.verticalAlignment = tblpPr['@_w:tblpYSpec'];
      if (isExplicitlySet(tblpPr['@_w:leftFromText'])) {
        position.leftFromText = safeParseInt(tblpPr['@_w:leftFromText']);
      }
      if (isExplicitlySet(tblpPr['@_w:rightFromText'])) {
        position.rightFromText = safeParseInt(tblpPr['@_w:rightFromText']);
      }
      if (isExplicitlySet(tblpPr['@_w:topFromText'])) {
        position.topFromText = safeParseInt(tblpPr['@_w:topFromText']);
      }
      if (isExplicitlySet(tblpPr['@_w:bottomFromText'])) {
        position.bottomFromText = safeParseInt(tblpPr['@_w:bottomFromText']);
      }

      if (Object.keys(position).length > 0) {
        table.setPosition(position);
      }
    }

    // Parse table overlap
    if (tblPrObj['w:tblOverlap']) {
      const val = tblPrObj['w:tblOverlap']['@_w:val'];
      table.setOverlap(val === 'overlap');
    }

    // Parse bidirectional visual layout — CT_OnOff, honour w:val per ECMA-376 §17.17.4
    if (tblPrObj['w:bidiVisual']) {
      table.setBidiVisual(parseOoxmlBoolean(tblPrObj['w:bidiVisual']));
    }

    // Parse table width — always set when w:tblW is present, including w:w="0" w:type="auto"
    // (auto-sized tables). Skipping w:w="0" would leave the constructor default (9360/dxa),
    // causing tblPrChange snapshots to capture wrong "previous" width values.
    if (tblPrObj['w:tblW']) {
      const width = safeParseInt(tblPrObj['w:tblW']['@_w:w'], 0);
      const widthType = tblPrObj['w:tblW']['@_w:type'] || 'dxa';
      table.setWidth(width);
      table.setWidthType(widthType);
    }

    // Parse table caption — ST_String per §17.4.62. Cast through
    // String() so a purely-numeric caption ("42") is preserved as a
    // string in `formatting.caption` rather than a JS number.
    if (tblPrObj['w:tblCaption']) {
      const caption = tblPrObj['w:tblCaption']['@_w:val'];
      if (caption !== undefined && caption !== null && caption !== '') {
        table.setCaption(String(caption));
      }
    }

    // Parse table description — ST_String per §17.4.63.
    if (tblPrObj['w:tblDescription']) {
      const description = tblPrObj['w:tblDescription']['@_w:val'];
      if (description !== undefined && description !== null && description !== '') {
        table.setDescription(String(description));
      }
    }

    // Parse table-level cell spacing (w:tblCellSpacing) per ECMA-376
    // §17.4.44 CT_TblCellSpacing. w:w is ST_MeasurementOrPercent; 0 is
    // a legal "explicit zero spacing" value (overrides any style-level
    // inherited tblCellSpacing). The emitter uses `!== undefined`, so
    // the previous `spacing > 0` gate created a parser/emitter
    // asymmetry: a tracked table-property change recording a *previous*
    // state of `<w:tblCellSpacing w:w="0" …/>` lost the override on
    // every round-trip.
    if (tblPrObj['w:tblCellSpacing']) {
      const rawW = tblPrObj['w:tblCellSpacing']['@_w:w'];
      if (isExplicitlySet(rawW)) {
        table.setCellSpacing(safeParseInt(rawW));
        const spacingType = tblPrObj['w:tblCellSpacing']['@_w:type'] || 'dxa';
        table.setCellSpacingType(spacingType);
      }
    }

    // Parse table layout (w:tblLayout) per ECMA-376 Part 1 §17.4.52
    if (tblPrObj['w:tblLayout']) {
      const layoutType = tblPrObj['w:tblLayout']['@_w:type'];
      if (layoutType) {
        table.setLayout(layoutType);
      }
    }

    // Parse table indentation (w:tblInd) per ECMA-376 Part 1 §17.4.43
    if (tblPrObj['w:tblInd']) {
      const indentVal = safeParseInt(tblPrObj['w:tblInd']['@_w:w'], 0);
      table.setIndent(indentVal);
      const indentType = tblPrObj['w:tblInd']['@_w:type'];
      if (indentType) {
        table.setIndentType(indentType as import('../elements/Table').TableWidthType);
      }
    }

    // Parse table cell margins (w:tblCellMar) per ECMA-376 Part 1 §17.4.42
    // Supports both legacy w:left/w:right and bidi-aware w:start/w:end (w:start takes precedence)
    if (tblPrObj['w:tblCellMar']) {
      const cellMar = tblPrObj['w:tblCellMar'];
      const margins: { top?: number; bottom?: number; left?: number; right?: number } = {};

      if (cellMar['w:top']) {
        const w = cellMar['w:top']['@_w:w'];
        if (w !== undefined) margins.top = parseInt(w, 10);
      }
      if (cellMar['w:bottom']) {
        const w = cellMar['w:bottom']['@_w:w'];
        if (w !== undefined) margins.bottom = parseInt(w, 10);
      }
      const leftSource = cellMar['w:start'] || cellMar['w:left'];
      if (leftSource) {
        const w = leftSource['@_w:w'];
        if (w !== undefined) margins.left = parseInt(w, 10);
      }
      const rightSource = cellMar['w:end'] || cellMar['w:right'];
      if (rightSource) {
        const w = rightSource['@_w:w'];
        if (w !== undefined) margins.right = parseInt(w, 10);
      }

      if (Object.keys(margins).length > 0) {
        table.setCellMargins(margins);
      }
    }

    // Parse table alignment (w:jc) - IMPORTANT for preserving table centering
    if (tblPrObj['w:jc']) {
      const alignment = tblPrObj['w:jc']['@_w:val'];
      if (alignment) {
        table.setAlignment(alignment as 'left' | 'center' | 'right');
      }
    }

    // Parse table style row band size (w:tblStyleRowBandSize) per ECMA-376 Part 1 §17.4.52
    if (tblPrObj['w:tblStyleRowBandSize']) {
      const val = parseInt(tblPrObj['w:tblStyleRowBandSize']['@_w:val'] || '0', 10);
      if (val > 0) {
        table.setStyleRowBandSize(val);
      }
    }

    // Parse table style column band size (w:tblStyleColBandSize) per ECMA-376 Part 1 §17.4.51
    if (tblPrObj['w:tblStyleColBandSize']) {
      const val = parseInt(tblPrObj['w:tblStyleColBandSize']['@_w:val'] || '0', 10);
      if (val > 0) {
        table.setStyleColBandSize(val);
      }
    }

    // Parse table shading (w:shd) per ECMA-376 Part 1 §17.4.56
    if (tblPrObj['w:shd']) {
      const shading = this.parseShadingFromObj(tblPrObj['w:shd']);
      if (shading) {
        table.setShading(shading);
      }
    }

    // Parse table borders (w:tblBorders) per ECMA-376 Part 1 §17.4.40.
    // left / right have bidi-aware aliases `w:start` / `w:end` (the
    // preferred spelling in modern Word-authored documents). Prefer
    // them when present, falling back to the legacy names — the
    // internal model stores under `left` / `right`, matching the
    // emitter. Without this fallback, any table whose side borders
    // were authored with the bidi-aware form silently lost those
    // borders on every round-trip (the emitter would replace them
    // with absent w:left/w:right, and the parser would never revive
    // the w:start/w:end it dropped).
    if (tblPrObj['w:tblBorders']) {
      const bordersObj = tblPrObj['w:tblBorders'];
      const borders: import('../elements/Table').TableBorders = {};

      if (bordersObj['w:top']) borders.top = this.parseBorderElement(bordersObj['w:top']);
      if (bordersObj['w:bottom']) borders.bottom = this.parseBorderElement(bordersObj['w:bottom']);
      const leftBorder = bordersObj['w:start'] ?? bordersObj['w:left'];
      if (leftBorder) borders.left = this.parseBorderElement(leftBorder);
      const rightBorder = bordersObj['w:end'] ?? bordersObj['w:right'];
      if (rightBorder) borders.right = this.parseBorderElement(rightBorder);
      if (bordersObj['w:insideH'])
        borders.insideH = this.parseBorderElement(bordersObj['w:insideH']);
      if (bordersObj['w:insideV'])
        borders.insideV = this.parseBorderElement(bordersObj['w:insideV']);

      if (Object.keys(borders).length > 0) {
        table.setBorders(borders);
      }
    }

    // Parse table property change (w:tblPrChange) per ECMA-376 Part 1 §17.13.5.36
    if (tblPrObj['w:tblPrChange']) {
      const changeObj = tblPrObj['w:tblPrChange'];
      table.setTblPrChange({
        id: String(changeObj['@_w:id'] || '0'),
        author: changeObj['@_w:author'] !== undefined ? String(changeObj['@_w:author']) : '',
        date: changeObj['@_w:date'] !== undefined ? String(changeObj['@_w:date']) : '',
        previousProperties: this.parseGenericPreviousProperties(changeObj['w:tblPr']),
      });
    }
  }

  private async parseTableRowFromObject(
    rowObj: any,
    relationshipManager: RelationshipManager,
    zipHandler: ZipHandler,
    imageManager: ImageManager,
    rawRowXml?: string
  ): Promise<TableRow | null> {
    try {
      // Create empty row
      const row = new TableRow();

      // Parse row properties (w:trPr) per ECMA-376 Part 1 §17.4.82
      const trPr = rowObj['w:trPr'];
      if (trPr) {
        this.parseTableRowPropertiesFromObject(trPr, row);
      }

      // Parse table property exceptions (w:tblPrEx) per ECMA-376 Part 1 §17.4.61
      const tblPrEx = rowObj['w:tblPrEx'];
      if (tblPrEx) {
        const exceptions = this.parseTablePropertyExceptionsFromObject(tblPrEx);
        if (exceptions) {
          row.setTablePropertyExceptions(exceptions);
        }
      }

      // Parse table cells (w:tc)
      const cells = rowObj['w:tc'];
      const cellChildren = Array.isArray(cells) ? cells : cells ? [cells] : [];

      // Extract cell XMLs from raw row XML if available
      let cellXmls: string[] = [];
      if (rawRowXml) {
        cellXmls = XMLParser.extractElements(rawRowXml, 'w:tc');
      }

      for (let i = 0; i < cellChildren.length; i++) {
        const cellObj = cellChildren[i];
        const rawCellXml = i < cellXmls.length ? cellXmls[i] : undefined;

        const cell = await this.parseTableCellFromObject(
          cellObj,
          relationshipManager,
          zipHandler,
          imageManager,
          rawCellXml
        );
        if (cell) {
          row.addCell(cell);
        }
      }

      return row;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse table row:',
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Parses table row properties from XML object and applies to row
   * Per ECMA-376 Part 1 §17.4.82 (w:trPr - Table Row Properties)
   */
  private parseTableRowPropertiesFromObject(trPrObj: any, row: TableRow): void {
    if (!trPrObj) return;

    // Parse row height (w:trHeight) per ECMA-376 Part 1 §17.4.81.
    // w:val is ST_TwipsMeasure; zero is a valid value and, combined
    // with w:hRule="exact", represents a hidden / collapsed row.
    // XMLParser coerces "0" to the number 0 (falsy), and the previous
    // `heightVal > 0` gate silently dropped explicit zero-height rows
    // even though the emitter (TableRow.ts §914) preserves them via
    // `!== undefined`. Route through isExplicitlySet so zero survives.
    // Per §17.18.33 (ST_HeightRule), when w:hRule is absent the
    // default is "auto".
    if (trPrObj['w:trHeight']) {
      const rawVal = trPrObj['w:trHeight']['@_w:val'];
      const heightRule = trPrObj['w:trHeight']['@_w:hRule'];
      if (isExplicitlySet(rawVal)) {
        const heightVal = safeParseInt(rawVal);
        row.setHeight(heightVal);
        if (heightRule) {
          row.setHeightRule(heightRule);
        } else {
          // When w:hRule is absent, clear the defaulted rule so the generator omits it,
          // preserving round-trip fidelity (absent = "auto" per ECMA-376 §17.18.33)
          row.setHeightRule(undefined);
        }
      }
    }

    // Parse table header row (w:tblHeader) per ECMA-376 Part 1 §17.4.49 — CT_OnOff
    if (trPrObj['w:tblHeader']) {
      row.setHeader(parseOoxmlBoolean(trPrObj['w:tblHeader']));
    }

    // Parse can't split (w:cantSplit) per ECMA-376 Part 1 §17.4.5 — CT_OnOff
    if (trPrObj['w:cantSplit']) {
      row.setCantSplit(parseOoxmlBoolean(trPrObj['w:cantSplit']));
    }

    // Parse row justification (w:jc) per ECMA-376 Part 1 §17.4.79
    if (trPrObj['w:jc']) {
      const val = trPrObj['w:jc']['@_w:val'];
      if (val) {
        row.setJustification(val);
      }
    }

    // Parse hidden (w:hidden) per ECMA-376 Part 1 §17.4.23 — CT_OnOff
    if (trPrObj['w:hidden']) {
      row.setHidden(parseOoxmlBoolean(trPrObj['w:hidden']));
    }

    // Parse grid before (w:gridBefore) per ECMA-376 Part 1 §17.4.15
    if (trPrObj['w:gridBefore']) {
      const val = parseInt(trPrObj['w:gridBefore']['@_w:val'] || '0', 10);
      if (val > 0) {
        row.setGridBefore(val);
      }
    }

    // Parse grid after (w:gridAfter) per ECMA-376 Part 1 §17.4.14
    if (trPrObj['w:gridAfter']) {
      const val = parseInt(trPrObj['w:gridAfter']['@_w:val'] || '0', 10);
      if (val > 0) {
        row.setGridAfter(val);
      }
    }

    // Parse width before (w:wBefore) per ECMA-376 Part 1 §17.4.83.
    // w:w is ST_TblWidth; 0 paired with w:type="auto" is the idiomatic
    // "no width" form, and explicit 0 in dxa twips can override an
    // inherited wBefore. Previous `w > 0` gate silently dropped both.
    if (trPrObj['w:wBefore']) {
      const rawW = trPrObj['w:wBefore']['@_w:w'];
      if (isExplicitlySet(rawW)) {
        const type = (trPrObj['w:wBefore']['@_w:type'] as string | undefined) || 'dxa';
        row.setWBefore(safeParseInt(rawW), type);
      }
    }

    // Parse width after (w:wAfter) per ECMA-376 Part 1 §17.4.82 — same
    // ST_TblWidth semantics as wBefore.
    if (trPrObj['w:wAfter']) {
      const rawW = trPrObj['w:wAfter']['@_w:w'];
      if (isExplicitlySet(rawW)) {
        const type = (trPrObj['w:wAfter']['@_w:type'] as string | undefined) || 'dxa';
        row.setWAfter(safeParseInt(rawW), type);
      }
    }

    // Parse row-level cell spacing (w:tblCellSpacing). Zero is a valid
    // override — "explicitly no extra spacing" on a row overriding a
    // non-zero table-level tblCellSpacing.
    if (trPrObj['w:tblCellSpacing']) {
      const rawW = trPrObj['w:tblCellSpacing']['@_w:w'];
      if (isExplicitlySet(rawW)) {
        const type = (trPrObj['w:tblCellSpacing']['@_w:type'] as string | undefined) || 'dxa';
        row.setRowCellSpacing(safeParseInt(rawW), type);
      }
    }

    // Parse conditional formatting (w:cnfStyle) per ECMA-376 Part 1 §17.3.1.8
    if (trPrObj['w:cnfStyle']) {
      const val = trPrObj['w:cnfStyle']['@_w:val'];
      if (val) {
        row.setCnfStyle(val);
      }
    }

    // Parse divId (w:divId) per ECMA-376 Part 1 §17.4.9. `w:val` is
    // ST_DecimalNumber; 0 is a valid reference to the first div in web
    // settings. The previous `val > 0` gate silently dropped it on load.
    if (trPrObj['w:divId']) {
      const rawVal = trPrObj['w:divId']['@_w:val'];
      if (isExplicitlySet(rawVal)) {
        const parsed = safeParseInt(rawVal);
        if (!isNaN(parsed)) row.setDivId(parsed);
      }
    }

    // Parse tracked row insertion / deletion (CT_TrackChange inside CT_TrPr)
    // per ECMA-376 Part 1 §17.13.5.19 (ins) / §17.13.5.14 (del). These mark
    // the entire row as a tracked revision; a previous version silently
    // dropped both markers on load → save because the parser skipped them.
    if (trPrObj['w:ins']) {
      const insObj = Array.isArray(trPrObj['w:ins']) ? trPrObj['w:ins'][0] : trPrObj['w:ins'];
      if (insObj && typeof insObj === 'object') {
        row.setRowInsertion({
          id: String(insObj['@_w:id'] ?? '0'),
          author: String(insObj['@_w:author'] ?? ''),
          date: String(insObj['@_w:date'] ?? ''),
        });
      }
    }
    if (trPrObj['w:del']) {
      const delObj = Array.isArray(trPrObj['w:del']) ? trPrObj['w:del'][0] : trPrObj['w:del'];
      if (delObj && typeof delObj === 'object') {
        row.setRowDeletion({
          id: String(delObj['@_w:id'] ?? '0'),
          author: String(delObj['@_w:author'] ?? ''),
          date: String(delObj['@_w:date'] ?? ''),
        });
      }
    }

    // Parse table row property change (w:trPrChange) per ECMA-376 Part 1 §17.13.5.38
    if (trPrObj['w:trPrChange']) {
      const changeObj = trPrObj['w:trPrChange'];
      row.setTrPrChange({
        id: String(changeObj['@_w:id'] || '0'),
        author: changeObj['@_w:author'] !== undefined ? String(changeObj['@_w:author']) : '',
        date: changeObj['@_w:date'] !== undefined ? String(changeObj['@_w:date']) : '',
        previousProperties: this.parseGenericPreviousProperties(changeObj['w:trPr']),
      });
    }
  }

  /**
   * Parses table property exceptions from row XML object
   * Per ECMA-376 Part 1 §17.4.61 (w:tblPrEx)
   * @private
   */
  private parseTablePropertyExceptionsFromObject(tblPrExObj: any): any {
    if (!tblPrExObj) return undefined;

    const exceptions: any = {};

    // Parse table width exception (w:tblW). The `val > 0` gate previously
    // dropped both w:w="0" (explicit zero-width override, valid when
    // paired with w:type="nil"/"auto") and negative overrides. Route
    // through isExplicitlySet + safeParseInt so zero and negative widths
    // round-trip.
    if (tblPrExObj['w:tblW']) {
      const rawW = tblPrExObj['w:tblW']['@_w:w'];
      if (isExplicitlySet(rawW)) {
        exceptions.width = safeParseInt(rawW);
      }
    }

    // Parse table justification exception (w:jc)
    if (tblPrExObj['w:jc']) {
      const val = tblPrExObj['w:jc']['@_w:val'];
      if (val) {
        exceptions.justification = val;
      }
    }

    // Parse cell spacing exception (w:tblCellSpacing). Zero-value
    // override is valid (= "explicit no cell spacing" on a row that
    // would otherwise inherit non-zero spacing from the table-level
    // tblCellSpacing).
    if (tblPrExObj['w:tblCellSpacing']) {
      const rawW = tblPrExObj['w:tblCellSpacing']['@_w:w'];
      if (isExplicitlySet(rawW)) {
        exceptions.cellSpacing = safeParseInt(rawW);
      }
    }

    // Parse table indentation exception (w:tblInd). Per ECMA-376
    // §17.4.62 CT_TblWidth, w:w is ST_MeasurementOrPercent — 0 is a
    // legal "reset" value and negative values indicate an outdent (table
    // hanging into the page margin). The previous `val > 0` check
    // silently dropped both.
    if (tblPrExObj['w:tblInd']) {
      const rawW = tblPrExObj['w:tblInd']['@_w:w'];
      if (isExplicitlySet(rawW)) {
        exceptions.indentation = safeParseInt(rawW);
      }
    }

    // Parse table borders exception (w:tblBorders)
    if (tblPrExObj['w:tblBorders']) {
      exceptions.borders = this.parseTableBordersFromObject(tblPrExObj['w:tblBorders']);
    }

    // Parse shading exception (w:shd)
    if (tblPrExObj['w:shd']) {
      const shading = this.parseShadingFromObj(tblPrExObj['w:shd']);
      if (shading) {
        exceptions.shading = shading;
      }
    }

    return Object.keys(exceptions).length > 0 ? exceptions : undefined;
  }

  /**
   * Parses table borders from XML object
   * @private
   */
  private parseTableBordersFromObject(bordersObj: any): any {
    if (!bordersObj) return undefined;

    const borders: any = {};
    const borderNames = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];

    for (const name of borderNames) {
      // Prefer bidi-aware `w:start`/`w:end` aliases over legacy `w:left`/
      // `w:right` per ECMA-376 §17.4.40 CT_TblBorders. Modern Word-
      // authored documents emit the bidi-aware form by default; the
      // internal model stores under the legacy keys to match the emitter.
      const aliasKey = name === 'left' ? 'w:start' : name === 'right' ? 'w:end' : undefined;
      const borderObj = (aliasKey && bordersObj[aliasKey]) || bordersObj[`w:${name}`];
      if (borderObj) {
        borders[name] = {};

        // Full CT_Border attribute set (§17.18.2) — previously only the four
        // basic attrs were read, so tblPrEx borders lost themed-color linkage
        // on every round-trip.
        if (borderObj['@_w:val']) borders[name].style = borderObj['@_w:val'];
        if (borderObj['@_w:sz']) borders[name].size = parseInt(borderObj['@_w:sz'], 10);
        if (borderObj['@_w:space']) borders[name].space = parseInt(borderObj['@_w:space'], 10);
        if (borderObj['@_w:color']) borders[name].color = String(borderObj['@_w:color']);
        // String(...) cast: themeTint / themeShade are ST_UcharHexNumber
        // (2-char hex). XMLParser coerces purely-digit hex to numbers —
        // cast so the string contract on the model is preserved.
        if (borderObj['@_w:themeColor']) {
          borders[name].themeColor = String(borderObj['@_w:themeColor']);
        }
        if (borderObj['@_w:themeTint']) {
          borders[name].themeTint = String(borderObj['@_w:themeTint']);
        }
        if (borderObj['@_w:themeShade']) {
          borders[name].themeShade = String(borderObj['@_w:themeShade']);
        }
        if (borderObj['@_w:shadow'] !== undefined) {
          borders[name].shadow = parseOnOffAttribute(String(borderObj['@_w:shadow']), true);
        }
        if (borderObj['@_w:frame'] !== undefined) {
          borders[name].frame = parseOnOffAttribute(String(borderObj['@_w:frame']), true);
        }
      }
    }

    return Object.keys(borders).length > 0 ? borders : undefined;
  }

  private async parseTableCellFromObject(
    cellObj: any,
    relationshipManager: RelationshipManager,
    zipHandler: ZipHandler,
    imageManager: ImageManager,
    rawCellXml?: string
  ): Promise<TableCell | null> {
    try {
      // Create empty cell
      const cell = new TableCell();

      // Parse cell properties (w:tcPr) per ECMA-376 Part 1 §17.4.42
      const tcPr = cellObj['w:tcPr'];
      if (tcPr) {
        // Parse cell width (w:tcW) with type per ECMA-376 Part 1 §17.4.72
        // CT_TblWidth — w:w is ST_MeasurementOrPercent, w:type is
        // ST_TblWidth. Zero is a legal explicit override:
        //   - `w:w="0" w:type="auto"` is the idiomatic "size to content"
        //     form (also the default when w:tcW is absent).
        //   - `w:w="0" w:type="dxa"` / `"pct"` / `"nil"` explicitly
        //     override an inherited non-zero width back to zero.
        // The emitter at TableCell.ts:1353 uses `!== undefined`, so the
        // previous `widthVal > 0 || widthType === 'auto'` gate created a
        // parser/emitter asymmetry — any cell with an explicit zero
        // override in a non-auto width type silently reinherited the
        // style-level width on every round-trip.
        if (tcPr['w:tcW']) {
          const rawW = tcPr['w:tcW']['@_w:w'];
          if (isExplicitlySet(rawW)) {
            const widthType = (tcPr['w:tcW']['@_w:type'] as string | undefined) || 'dxa';
            cell.setWidthType(
              safeParseInt(rawW),
              widthType as import('../elements/TableCell').CellWidthType
            );
          }
        }

        // Parse conditional style (w:cnfStyle) per ECMA-376 Part 1 §17.4.7
        if (tcPr['w:cnfStyle']) {
          const cnfStyle = tcPr['w:cnfStyle']['@_w:val'];
          if (cnfStyle) {
            cell.setConditionalStyle(cnfStyle);
          }
        }

        // Parse cell borders (w:tcBorders) per ECMA-376 Part 1 §17.4.66.
        // Supports both legacy LTR names (w:left / w:right) and bidi-
        // aware aliases (w:start / w:end). Prefer w:start / w:end when
        // present. Includes diagonal borders (w:tl2br / w:tr2bl) which
        // are cell-specific.
        if (tcPr['w:tcBorders']) {
          const bordersObj = tcPr['w:tcBorders'];
          const borders: any = {};

          if (bordersObj['w:top']) borders.top = this.parseBorderElement(bordersObj['w:top']);
          if (bordersObj['w:bottom'])
            borders.bottom = this.parseBorderElement(bordersObj['w:bottom']);
          const leftBorder = bordersObj['w:start'] ?? bordersObj['w:left'];
          if (leftBorder) borders.left = this.parseBorderElement(leftBorder);
          const rightBorder = bordersObj['w:end'] ?? bordersObj['w:right'];
          if (rightBorder) borders.right = this.parseBorderElement(rightBorder);
          if (bordersObj['w:tl2br']) borders.tl2br = this.parseBorderElement(bordersObj['w:tl2br']);
          if (bordersObj['w:tr2bl']) borders.tr2bl = this.parseBorderElement(bordersObj['w:tr2bl']);

          if (Object.keys(borders).length > 0) {
            cell.setBorders(borders);
          }
        }

        // Parse cell shading (w:shd)
        if (tcPr['w:shd']) {
          const shading = this.parseShadingFromObj(tcPr['w:shd']);
          if (shading) {
            cell.setShading(shading);
          }
        }

        // Parse cell margins (w:tcMar) per ECMA-376 Part 1 §17.4.43
        // Supports both legacy w:left/w:right and bidi-aware w:start/w:end (w:start takes precedence)
        if (tcPr['w:tcMar']) {
          const tcMar = tcPr['w:tcMar'];
          const margins: any = {};

          if (tcMar['w:top']) {
            margins.top = parseInt(tcMar['w:top']['@_w:w'] || '0', 10);
          }
          if (tcMar['w:bottom']) {
            margins.bottom = parseInt(tcMar['w:bottom']['@_w:w'] || '0', 10);
          }
          const leftSrc = tcMar['w:start'] || tcMar['w:left'];
          if (leftSrc) {
            margins.left = parseInt(leftSrc['@_w:w'] || '0', 10);
          }
          const rightSrc = tcMar['w:end'] || tcMar['w:right'];
          if (rightSrc) {
            margins.right = parseInt(rightSrc['@_w:w'] || '0', 10);
          }

          if (Object.keys(margins).length > 0) {
            cell.setMargins(margins);
          }
        }

        // Parse vertical alignment (w:vAlign) per ECMA-376 §17.4.83.
        // ST_VerticalJc has four values (§17.18.101): top, center, both,
        // bottom. The previous whitelist dropped "both" silently — the
        // style-level parser accepts it, so the asymmetry truncated cell
        // vertical alignment on cells using the "both" (justified)
        // vertical alignment on load.
        if (tcPr['w:vAlign']) {
          const valign = tcPr['w:vAlign']['@_w:val'];
          if (valign === 'top' || valign === 'center' || valign === 'both' || valign === 'bottom') {
            cell.setVerticalAlignment(valign);
          }
        }

        // Parse column span (w:gridSpan)
        if (tcPr['w:gridSpan']) {
          const span = parseInt(tcPr['w:gridSpan']['@_w:val'] || '1', 10);
          if (span > 1) {
            cell.setColumnSpan(span);
          }
        }

        // Parse text direction (w:textDirection) per ECMA-376 Part 1 §17.4.72
        if (tcPr['w:textDirection']) {
          const direction = tcPr['w:textDirection']['@_w:val'];
          if (direction) {
            cell.setTextDirection(direction);
          }
        }

        // Parse no wrap (w:noWrap) per ECMA-376 Part 1 §17.4.34 — CT_OnOff
        if (tcPr['w:noWrap']) {
          cell.setNoWrap(parseOoxmlBoolean(tcPr['w:noWrap']));
        }

        // Parse hide mark (w:hideMark) per ECMA-376 Part 1 §17.4.24 — CT_OnOff
        if (tcPr['w:hideMark']) {
          cell.setHideMark(parseOoxmlBoolean(tcPr['w:hideMark']));
        }

        // Parse headers (w:headers) per ECMA-376 Part 1 §17.4.26
        if (tcPr['w:headers']) {
          const headersVal = tcPr['w:headers']['@_w:val'];
          if (headersVal) {
            cell.setHeaders(headersVal);
          }
        }

        // Parse fit text (w:tcFitText) per ECMA-376 Part 1 §17.4.68 — CT_OnOff
        if (tcPr['w:tcFitText']) {
          cell.setFitText(parseOoxmlBoolean(tcPr['w:tcFitText']));
        }

        // Parse vertical merge (w:vMerge) per ECMA-376 Part 1 §17.4.85
        if (tcPr['w:vMerge']) {
          const vMergeVal = tcPr['w:vMerge']['@_w:val'];
          // Empty element or "continue" means continue, "restart" means restart
          if (vMergeVal === 'restart') {
            cell.setVerticalMerge('restart');
          } else {
            cell.setVerticalMerge('continue');
          }
        }

        // Parse legacy horizontal merge (w:hMerge) per ECMA-376 Part 1 §17.4.22
        if (tcPr['w:hMerge']) {
          const hMergeVal = tcPr['w:hMerge']['@_w:val'];
          if (hMergeVal === 'restart') {
            cell.setHorizontalMerge('restart');
          } else {
            cell.setHorizontalMerge('continue');
          }
        }

        // Parse table cell insertion marker (w:cellIns) per ECMA-376 Part 1 §17.13.5.5
        if (tcPr['w:cellIns']) {
          const cellIns = tcPr['w:cellIns'];
          const id = parseInt(String(cellIns['@_w:id'] ?? '0'), 10);
          const author =
            cellIns['@_w:author'] !== undefined ? String(cellIns['@_w:author']) : 'Unknown';
          const dateAttr = cellIns['@_w:date'];
          const date = dateAttr !== undefined ? new Date(String(dateAttr)) : new Date();

          const revision = new Revision({
            id,
            author,
            date,
            type: 'tableCellInsert',
            content: [],
          });
          cell.setCellRevision(revision);
        }

        // Parse table cell deletion marker (w:cellDel) per ECMA-376 Part 1 §17.13.5.6
        if (tcPr['w:cellDel']) {
          const cellDel = tcPr['w:cellDel'];
          const id = parseInt(String(cellDel['@_w:id'] ?? '0'), 10);
          const author =
            cellDel['@_w:author'] !== undefined ? String(cellDel['@_w:author']) : 'Unknown';
          const dateAttr = cellDel['@_w:date'];
          const date = dateAttr !== undefined ? new Date(String(dateAttr)) : new Date();

          const revision = new Revision({
            id,
            author,
            date,
            type: 'tableCellDelete',
            content: [],
          });
          cell.setCellRevision(revision);
        }

        // Parse table cell merge marker (w:cellMerge) per ECMA-376 Part 1 §17.13.5.4
        if (tcPr['w:cellMerge']) {
          const cellMerge = tcPr['w:cellMerge'];
          const id = parseInt(String(cellMerge['@_w:id'] ?? '0'), 10);
          const author =
            cellMerge['@_w:author'] !== undefined ? String(cellMerge['@_w:author']) : 'Unknown';
          const dateAttr = cellMerge['@_w:date'];
          const date = dateAttr !== undefined ? new Date(String(dateAttr)) : new Date();
          const vMergeAttr = cellMerge['@_w:vMerge'];
          const vMergeOrigAttr = cellMerge['@_w:vMergeOrig'];
          // ST_AnnotationVMerge uses "rest"/"cont" but API uses "restart"/"continue"
          const mergeRevMap: Record<string, string> = { rest: 'restart', cont: 'continue' };

          const revision = new Revision({
            id,
            author,
            date,
            type: 'tableCellMerge',
            content: [],
            previousProperties: {
              vMerge: (vMergeAttr && mergeRevMap[vMergeAttr]) || vMergeAttr,
              vMergeOrig: (vMergeOrigAttr && mergeRevMap[vMergeOrigAttr]) || vMergeOrigAttr,
            },
          });
          cell.setCellRevision(revision);
        }

        // Parse table cell property change (w:tcPrChange) per ECMA-376 Part 1 §17.13.5.37
        if (tcPr['w:tcPrChange']) {
          const changeObj = tcPr['w:tcPrChange'];
          cell.setTcPrChange({
            id: String(changeObj['@_w:id'] || '0'),
            author: changeObj['@_w:author'] !== undefined ? String(changeObj['@_w:author']) : '',
            date: changeObj['@_w:date'] !== undefined ? String(changeObj['@_w:date']) : '',
            previousProperties: this.parseGenericPreviousProperties(changeObj['w:tcPr']),
          });
        }
      }

      // Parse cell content - use order-preserving extraction if raw XML available
      // This is critical for preserving nested tables and SDTs
      if (rawCellXml) {
        // Extract all cell content in order (paragraphs, nested tables, SDTs, bookmarkEnds)
        const cellContent = this.extractCellContentInOrder(rawCellXml);
        let paragraphIndex = 0;
        let lastParagraph: Paragraph | null = null;

        for (const item of cellContent) {
          if (item.type === 'paragraph') {
            // Parse paragraph using raw XML to preserve revisions
            const paragraph = await this.parseParagraphWithOrder(
              item.xml,
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (paragraph) {
              cell.addParagraph(paragraph);
              lastParagraph = paragraph;
              paragraphIndex++;
            }
          } else if (item.type === 'table' || item.type === 'sdt') {
            // Store nested tables and SDTs as raw XML for passthrough
            // These are preserved exactly as-is to avoid any modifications
            cell.addRawNestedContent(paragraphIndex, item.xml, item.type);
          } else if (item.type === 'bookmarkEnd') {
            // BookmarkEnd between paragraphs in cell - attach to previous paragraph
            if (lastParagraph) {
              const bookmarkEnds = this.extractBookmarkEndsFromContent(item.xml);
              for (const bookmark of bookmarkEnds) {
                lastParagraph.addBookmarkEnd(bookmark);
              }
            }
          }
        }
      } else {
        // Fallback: Parse paragraphs from object (no raw XML available)
        const paragraphs = cellObj['w:p'];
        const paraChildren = Array.isArray(paragraphs)
          ? paragraphs
          : paragraphs
            ? [paragraphs]
            : [];

        for (const paraObj of paraChildren) {
          const paragraph = await this.parseParagraphFromObject(
            paraObj,
            relationshipManager,
            zipHandler,
            imageManager
          );
          if (paragraph) {
            cell.addParagraph(paragraph);
          }
        }
      }

      return cell;
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse table cell:',
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Extracts cell content elements in order, preserving sequence of paragraphs, nested tables, and SDTs
   * Used to properly handle nested tables in table cells
   * @param cellXml - Raw XML of the cell (w:tc element)
   * @returns Array of content items with type and XML
   */
  private extractCellContentInOrder(
    cellXml: string
  ): { type: 'paragraph' | 'table' | 'sdt' | 'bookmarkEnd'; xml: string }[] {
    const result: {
      type: 'paragraph' | 'table' | 'sdt' | 'bookmarkEnd';
      xml: string;
    }[] = [];

    // Find the start of cell content (after w:tcPr if present)
    let contentStart = 0;
    const tcPrStart = cellXml.indexOf('<w:tcPr');
    if (tcPrStart !== -1) {
      // Skip past the tcPr element
      const tcPrEnd = this.findClosingTag(cellXml, 'w:tcPr', tcPrStart);
      if (tcPrEnd !== -1) {
        contentStart = tcPrEnd;
      }
    }

    // Find end of cell (before closing </w:tc>)
    const contentEnd = cellXml.lastIndexOf('</w:tc>');
    if (contentEnd === -1) return result;

    const content = cellXml.substring(contentStart, contentEnd);

    // Scan for w:p, w:tbl, w:sdt, and w:bookmarkEnd elements at the current level
    let pos = 0;
    while (pos < content.length) {
      // Find next element start
      const pStart = content.indexOf('<w:p', pos);
      const tblStart = content.indexOf('<w:tbl', pos);
      const sdtStart = content.indexOf('<w:sdt', pos);
      const bookmarkEndStart = content.indexOf('<w:bookmarkEnd', pos);

      // Find which comes first
      let nextStart = -1;
      let nextType: 'paragraph' | 'table' | 'sdt' | 'bookmarkEnd' | null = null;
      let nextTag = '';

      if (
        pStart !== -1 &&
        this.isExactTag(content, pStart, 'w:p') &&
        (nextStart === -1 || pStart < nextStart)
      ) {
        nextStart = pStart;
        nextType = 'paragraph';
        nextTag = 'w:p';
      }
      if (
        tblStart !== -1 &&
        this.isExactTag(content, tblStart, 'w:tbl') &&
        (nextStart === -1 || tblStart < nextStart)
      ) {
        nextStart = tblStart;
        nextType = 'table';
        nextTag = 'w:tbl';
      }
      if (
        sdtStart !== -1 &&
        this.isExactTag(content, sdtStart, 'w:sdt') &&
        (nextStart === -1 || sdtStart < nextStart)
      ) {
        nextStart = sdtStart;
        nextType = 'sdt';
        nextTag = 'w:sdt';
      }
      if (bookmarkEndStart !== -1 && (nextStart === -1 || bookmarkEndStart < nextStart)) {
        nextStart = bookmarkEndStart;
        nextType = 'bookmarkEnd';
        nextTag = 'w:bookmarkEnd';
      }

      if (nextStart === -1 || nextType === null) break;

      // Extract the complete element
      if (nextType === 'bookmarkEnd') {
        // Self-closing tag - find the end of this element
        const elementEnd = content.indexOf('>', nextStart) + 1;
        if (elementEnd === 0) {
          pos = nextStart + 1;
          continue;
        }
        const elementXml = content.substring(nextStart, elementEnd);
        result.push({ type: nextType, xml: elementXml });
        pos = elementEnd;
      } else {
        const elementEnd = this.findClosingTag(content, nextTag, nextStart);
        if (elementEnd === -1) {
          pos = nextStart + 1;
          continue;
        }
        const elementXml = content.substring(nextStart, elementEnd);
        result.push({ type: nextType, xml: elementXml });
        pos = elementEnd;
      }
    }

    return result;
  }

  /**
   * Checks if a tag at the given position is an exact match (not a prefix like w:pPr for w:p)
   */
  private isExactTag(content: string, position: number, tagName: string): boolean {
    const afterTag = content[position + 1 + tagName.length];
    return (
      afterTag === '>' ||
      afterTag === '/' ||
      afterTag === ' ' ||
      afterTag === '\t' ||
      afterTag === '\n' ||
      afterTag === '\r'
    );
  }

  /**
   * Finds the closing tag for an element, handling nested elements of the same type
   */
  private findClosingTag(content: string, tagName: string, startPos: number): number {
    const openTag = `<${tagName}`;
    const closeTag = `</${tagName}>`;

    // Check for self-closing tag first
    const openTagEnd = content.indexOf('>', startPos);
    if (openTagEnd !== -1 && content[openTagEnd - 1] === '/') {
      return openTagEnd + 1;
    }

    let depth = 1;
    let pos = openTagEnd + 1;

    while (pos < content.length && depth > 0) {
      const nextOpen = content.indexOf(openTag, pos);
      const nextClose = content.indexOf(closeTag, pos);

      if (nextClose === -1) break;

      if (nextOpen !== -1 && nextOpen < nextClose) {
        // Check if it's an exact tag match
        if (this.isExactTag(content, nextOpen, tagName)) {
          depth++;
        }
        pos = nextOpen + openTag.length;
      } else {
        depth--;
        pos = nextClose + closeTag.length;
      }
    }

    return depth === 0 ? pos : -1;
  }

  private async parseSDTFromObject(
    sdtObj: any,
    relationshipManager: RelationshipManager,
    zipHandler: ZipHandler,
    imageManager: ImageManager
  ): Promise<StructuredDocumentTag | TableOfContentsElement | null> {
    try {
      if (!sdtObj) return null;

      const properties: any = {};

      // Parse SDT properties (sdtPr)
      const sdtPr = sdtObj['w:sdtPr'];
      if (sdtPr) {
        // Parse `<w:id w:val="…"/>` per ECMA-376 §17.5.2.18. `w:val` is
        // ST_DecimalNumber (xsd:integer) — 0 is legal. XMLParser coerces
        // `"0"` to the number `0`, so the previous truthy gate silently
        // dropped w:id=0 on every load → save cycle. The emitter uses
        // `!== undefined`, creating a parser/emitter asymmetry.
        const idElement = sdtPr['w:id'];
        if (isExplicitlySet(idElement?.['@_w:val'])) {
          const parsed = safeParseInt(idElement['@_w:val']);
          if (!isNaN(parsed)) properties.id = parsed;
        }

        // Parse `<w:tag w:val="…"/>` per ECMA-376 §17.5.2.34. `w:val`
        // is ST_String — any string is legal, including numeric-looking
        // strings like "123" that XMLParser coerces to the number 123.
        // Cast via `String(…)` so the tag round-trips as text rather
        // than leaking a JS number into a `tag?: string` field.
        const tagElement = sdtPr['w:tag'];
        if (tagElement?.['@_w:val'] !== undefined) {
          properties.tag = String(tagElement['@_w:val']);
        }

        // Parse lock — ST_Lock enum: "sdtLocked" / "contentLocked" /
        // "sdtContentLocked" / "unlocked". Always a non-numeric string,
        // so no XMLParser coercion concern; truthy check fine.
        const lockElement = sdtPr['w:lock'];
        if (lockElement?.['@_w:val']) {
          properties.lock = lockElement['@_w:val'];
        }

        // Parse alias — ST_String. Same numeric-coercion concern as
        // `w:tag`; cast via `String(…)`.
        const aliasElement = sdtPr['w:alias'];
        if (aliasElement?.['@_w:val'] !== undefined) {
          properties.alias = String(aliasElement['@_w:val']);
        }

        // Parse control type from various elements
        if (sdtPr['w:richText']) {
          properties.controlType = 'richText';
        } else if (sdtPr['w:text']) {
          properties.controlType = 'plainText';
          const textElement = sdtPr['w:text'];
          // w:multiLine is an OPTIONAL ST_OnOff attribute per ECMA-376
          // §17.5.2.33 CT_SdtText. Only record a value when the source
          // actually set it — otherwise leave the field undefined so
          // the emitter (which uses `!== undefined`) preserves the
          // "attribute absent" state on round-trip. Previously the
          // parser unconditionally stored `false` for any absent
          // attribute, then the emitter wrote `w:multiLine="0"` —
          // adding spec-noise that wasn't in the source.
          const rawMultiLine = textElement?.['@_w:multiLine'];
          properties.plainText = {
            multiLine:
              rawMultiLine === undefined ? undefined : parseOnOffAttribute(String(rawMultiLine)),
          };
        } else if (sdtPr['w:comboBox']) {
          properties.controlType = 'comboBox';
          const comboBoxElement = sdtPr['w:comboBox'];
          properties.comboBox = this.parseListItems(comboBoxElement);
        } else if (sdtPr['w:dropDownList']) {
          properties.controlType = 'dropDownList';
          const dropDownElement = sdtPr['w:dropDownList'];
          properties.dropDownList = this.parseListItems(dropDownElement);
        } else if (sdtPr['w:date']) {
          properties.controlType = 'datePicker';
          const dateElement = sdtPr['w:date'];
          // Date properties can be either attributes on w:date or nested elements
          properties.datePicker = {
            // Check attribute first, then nested element
            dateFormat:
              dateElement?.['@_w:dateFormat'] || dateElement?.['w:dateFormat']?.['@_w:val'],
            fullDate:
              dateElement?.['@_w:fullDate'] || dateElement?.['w:fullDate']?.['@_w:val']
                ? new Date(dateElement['@_w:fullDate'] || dateElement['w:fullDate']['@_w:val'])
                : undefined,
            lid: dateElement?.['@_w:lid'] || dateElement?.['w:lid']?.['@_w:val'],
            calendar: dateElement?.['@_w:calendar'] || dateElement?.['w:calendar']?.['@_w:val'],
          };
        } else if (sdtPr['w14:checkbox']) {
          properties.controlType = 'checkbox';
          const checkboxElement = sdtPr['w14:checkbox'];
          // <w14:checked> is CT_OnOff in the Word 2010+ extension namespace.
          // Honour every ST_OnOff literal ("1"/"0"/"true"/"false"/"on"/"off")
          // and treat a bare self-closing `<w14:checked/>` as true.
          properties.checkbox = {
            checked: parseOoxmlBoolean(checkboxElement?.['w14:checked'], '@_w14:val'),
            checkedState: String(checkboxElement?.['w14:checkedState']?.['@_w14:val'] ?? ''),
            uncheckedState: String(checkboxElement?.['w14:uncheckedState']?.['@_w14:val'] ?? ''),
          };
        } else if (sdtPr['w:picture']) {
          properties.controlType = 'picture';
        } else if (sdtPr['w:docPartObj']) {
          properties.controlType = 'buildingBlock';
          const docPartObj = sdtPr['w:docPartObj'];
          properties.buildingBlock = {
            gallery: docPartObj?.['w:docPartGallery']?.['@_w:val'],
            category: docPartObj?.['w:docPartCategory']?.['@_w:val'],
          };
        } else if (sdtPr['w:group']) {
          properties.controlType = 'group';
        } else if (sdtPr['w:citation']) {
          properties.controlType = 'citation';
        } else if (sdtPr['w:bibliography']) {
          properties.controlType = 'bibliography';
        } else if (sdtPr['w:equation']) {
          properties.controlType = 'equation';
        } else if (sdtPr['w:docPartList']) {
          properties.controlType = 'docPartList';
          const docPartList = sdtPr['w:docPartList'];
          properties.buildingBlock = {
            gallery: docPartList?.['w:docPartGallery']?.['@_w:val'],
            category: docPartList?.['w:docPartCategory']?.['@_w:val'],
            isList: true,
          };
        }

        // Parse placeholder (w:placeholder/w:docPart)
        const placeholderElement = sdtPr['w:placeholder'];
        if (placeholderElement) {
          const docPartVal = placeholderElement?.['w:docPart']?.['@_w:val'];
          if (docPartVal) {
            properties.placeholder = { docPart: docPartVal };
          }
        }

        // Parse data binding (w:dataBinding)
        const dataBindingElement = sdtPr['w:dataBinding'];
        if (dataBindingElement) {
          properties.dataBinding = {
            xpath: dataBindingElement['@_w:xpath'] || '',
            prefixMappings: dataBindingElement['@_w:prefixMappings'],
            storeItemId: dataBindingElement['@_w:storeItemID'],
          };
        }

        // Parse showing placeholder flag (w:showingPlcHdr) — CT_OnOff per ECMA-376 §17.5.2.40
        const showingPlcHdr = sdtPr['w:showingPlcHdr'];
        if (showingPlcHdr) {
          properties.showingPlcHdr = parseOoxmlBoolean(showingPlcHdr);
        }
      }

      // Parse SDT content (sdtContent)
      const content: any[] = [];
      const sdtContent = sdtObj['w:sdtContent'];
      if (sdtContent) {
        // Check for ordered children (preserves element order)
        const orderedChildren = sdtContent._orderedChildren as
          | { type: string; index: number }[]
          | undefined;

        if (orderedChildren && orderedChildren.length > 0) {
          // Process in original order
          for (const childInfo of orderedChildren) {
            const elementType = childInfo.type;
            const elementIndex = childInfo.index;

            if (elementType === 'w:p') {
              const paragraphs = sdtContent['w:p'];
              const paraArray = Array.isArray(paragraphs)
                ? paragraphs
                : paragraphs
                  ? [paragraphs]
                  : [];
              if (elementIndex < paraArray.length) {
                // Reconstruct XML for paragraph parsing
                const paraXml = this.objectToXml({
                  'w:p': paraArray[elementIndex],
                });
                const para = await this.parseParagraphWithOrder(
                  paraXml,
                  relationshipManager,
                  zipHandler,
                  imageManager
                );
                if (para) content.push(para);
              }
            } else if (elementType === 'w:tbl') {
              const tables = sdtContent['w:tbl'];
              const tableArray = Array.isArray(tables) ? tables : tables ? [tables] : [];
              if (elementIndex < tableArray.length) {
                const tableObj = tableArray[elementIndex];
                const table = await this.parseTableFromObject(
                  tableObj,
                  relationshipManager,
                  zipHandler,
                  imageManager
                );
                if (table) content.push(table);
              }
            } else if (elementType === 'w:sdt') {
              const sdts = sdtContent['w:sdt'];
              const sdtArray = Array.isArray(sdts) ? sdts : sdts ? [sdts] : [];
              if (elementIndex < sdtArray.length) {
                const nestedSdt = await this.parseSDTFromObject(
                  sdtArray[elementIndex],
                  relationshipManager,
                  zipHandler,
                  imageManager
                );
                if (nestedSdt) content.push(nestedSdt);
              }
            }
          }
        } else {
          // Fallback: process sequentially
          // Parse paragraphs
          const paragraphs = sdtContent['w:p'];
          const paraArray = Array.isArray(paragraphs) ? paragraphs : paragraphs ? [paragraphs] : [];
          for (const paraObj of paraArray) {
            const paraXml = this.objectToXml({ 'w:p': paraObj });
            const para = await this.parseParagraphWithOrder(
              paraXml,
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (para) content.push(para);
          }

          // Parse tables
          const tables = sdtContent['w:tbl'];
          const tableArray = Array.isArray(tables) ? tables : tables ? [tables] : [];
          for (const tableObj of tableArray) {
            const table = await this.parseTableFromObject(
              tableObj,
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (table) content.push(table);
          }

          // Parse nested SDTs
          const nestedSdts = sdtContent['w:sdt'];
          const sdtArray = Array.isArray(nestedSdts) ? nestedSdts : nestedSdts ? [nestedSdts] : [];
          for (const nestedSdtObj of sdtArray) {
            const nestedSdt = await this.parseSDTFromObject(
              nestedSdtObj,
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (nestedSdt) content.push(nestedSdt);
          }
        }
      }

      // Check if this is a Table of Contents SDT
      if (properties.buildingBlock?.gallery === 'Table of Contents') {
        // This is a TOC - create TableOfContentsElement instead
        const toc = this.parseTOCFromSDTContent(content, properties, sdtContent);
        if (toc) {
          return new TableOfContentsElement(toc);
        }
      }

      return new StructuredDocumentTag(properties, content);
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse SDT:',
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Extracts TOC field instruction from assembled ComplexField objects
   * This is the preferred method as it uses already-assembled fields
   * @private
   */
  private extractInstructionFromContent(content: any[]): string | undefined {
    defaultLogger.debug(
      `[TOC Parser] Searching ${content.length} content elements for TOC ComplexField`
    );

    for (const element of content) {
      if (element instanceof Paragraph) {
        const paragraphContent = element.getContent();

        for (const item of paragraphContent) {
          // Check for ComplexField (assembled field with begin/sep/end)
          if (item instanceof ComplexField) {
            const instruction = item.getInstruction();

            // Check if this is a TOC field
            if (instruction && instruction.trim().startsWith('TOC')) {
              defaultLogger.debug(
                `[TOC Parser] Found ComplexField with TOC instruction: "${instruction.substring(
                  0,
                  100
                )}..."`
              );
              return instruction.trim();
            }
          }
        }
      }
    }

    defaultLogger.debug(
      '[TOC Parser] No ComplexField with TOC instruction found in assembled content'
    );
    return undefined;
  }

  /**
   * Fallback: Extracts TOC instruction from raw XML when ComplexField not available
   * Handles instructions split across multiple runs by concatenating all runs between begin/separate markers
   * @private
   */
  private extractInstructionFromRawXML(sdtContent: any): string | undefined {
    const paragraphs = sdtContent['w:p'];
    const paraArray = Array.isArray(paragraphs) ? paragraphs : paragraphs ? [paragraphs] : [];

    defaultLogger.debug(
      `[TOC Parser] Fallback: Parsing raw XML from ${paraArray.length} paragraph(s)`
    );

    // Track field state across paragraphs (TOC fields can span multiple paragraphs)
    let inField = false;
    let instructionParts: string[] = [];
    let foundTOCInstruction: string | undefined;

    for (let pIdx = 0; pIdx < paraArray.length; pIdx++) {
      const paraObj = paraArray[pIdx];
      const runs = paraObj['w:r'];
      const runArray = Array.isArray(runs) ? runs : runs ? [runs] : [];

      defaultLogger.debug(`[TOC Parser] Paragraph ${pIdx + 1}: ${runArray.length} runs`);

      for (let rIdx = 0; rIdx < runArray.length; rIdx++) {
        const runObj = runArray[rIdx];

        // Check for field character(s) - can be single object or array
        const fldChar = runObj['w:fldChar'];
        if (fldChar) {
          // Handle both single object and array of objects
          const fldCharArray = Array.isArray(fldChar) ? fldChar : [fldChar];

          defaultLogger.debug(
            `[TOC Parser] Paragraph ${pIdx + 1}, Run ${rIdx + 1}: Found ${
              fldCharArray.length
            } fldChar element(s)`
          );

          // Process each fldChar in the array
          for (const fldCharObj of fldCharArray) {
            const charType = fldCharObj['@_w:fldCharType'] || fldCharObj['@_fldCharType'];

            if (!charType) {
              defaultLogger.debug(`[TOC Parser] Warning: fldChar without charType attribute`);
              continue;
            }

            defaultLogger.debug(
              `[TOC Parser] Paragraph ${pIdx + 1}, Run ${rIdx + 1}: fldChar type = "${charType}"`
            );

            if (charType === 'begin') {
              inField = true;
              instructionParts = [];
              defaultLogger.debug(
                '[TOC Parser] Found field begin marker, starting instruction collection'
              );
              continue;
            }

            // Check for field end/separate marker
            if (charType === 'end' || charType === 'separate') {
              // If we collected instruction parts, join and check if TOC
              const fullInstruction = instructionParts.join('').trim();
              defaultLogger.debug(
                `[TOC Parser] Field ${charType} marker found. Collected instruction: "${fullInstruction.substring(
                  0,
                  100
                )}..."`
              );

              if (fullInstruction.startsWith('TOC')) {
                foundTOCInstruction = fullInstruction;
                defaultLogger.debug(
                  `[TOC Parser] ✓ Extracted complete TOC instruction from ${instructionParts.length} part(s)`
                );
                // Don't return yet - field might have more content after separate
                // But save it in case we don't find end marker
              }

              // Only reset on "end", not "separate" (there's content between separate and end)
              if (charType === 'end') {
                inField = false;
                instructionParts = [];

                // If we found a TOC instruction, return it now
                if (foundTOCInstruction) {
                  return foundTOCInstruction;
                }
              }
              continue;
            }
          }
        }

        // Collect instruction text while inside field
        if (inField) {
          const instrText = runObj['w:instrText'];
          if (instrText) {
            // Extract text value and decode XML entities
            let text = '';
            if (typeof instrText === 'string') {
              text = instrText;
            } else if (instrText['#text']) {
              text = instrText['#text'];
            } else {
              // Sometimes the text is directly in the object
              text = String(instrText);
            }

            instructionParts.push(text);
            defaultLogger.debug(
              `[TOC Parser] Paragraph ${pIdx + 1}, Run ${
                rIdx + 1
              }: Collected instrText = "${text.substring(0, 50)}..."`
            );
          }
        }
      }
    }

    // If we found a TOC instruction but never hit the end marker, return it anyway
    if (foundTOCInstruction) {
      defaultLogger.debug(
        `[TOC Parser] Returning TOC instruction (no end marker found): "${foundTOCInstruction.substring(
          0,
          100
        )}..."`
      );
      return foundTOCInstruction;
    }

    defaultLogger.debug('[TOC Parser] No TOC instruction found in raw XML fallback');
    return undefined;
  }

  /**
   * Normalizes TOC field instructions by converting incomplete \t switches
   * that reference only standard heading styles to the equivalent \o switch format.
   *
   * This fixes documents where TOC was created with \t "Heading 2,2," instead of
   * the correct \o "1-3" format, which causes the TOC to only capture one heading level.
   *
   * @param instruction - The original TOC field instruction
   * @returns Normalized instruction with \o switch if applicable, otherwise original
   * @private
   */
  private normalizeTOCFieldInstruction(instruction: string): string {
    // Check if instruction uses \t switch
    if (!instruction.includes('\\t')) {
      return instruction; // Already using \o or no style switch
    }

    // Extract all \t switches
    const tSwitchPattern = /\\t\s+"([^"]+)"/g;
    const matches = [...instruction.matchAll(tSwitchPattern)];

    if (matches.length === 0) {
      return instruction; // No \t switches found
    }

    // Parse all styles from \t switches
    const styles: { styleName: string; level: number }[] = [];
    for (const match of matches) {
      const stylesStr = match[1];
      if (!stylesStr) continue;

      const parts = stylesStr.split(',').filter((p: string) => p.trim());
      for (let i = 0; i < parts.length; i += 2) {
        const styleName = parts[i];
        const levelStr = parts[i + 1];
        if (styleName && levelStr) {
          styles.push({
            styleName: styleName.trim(),
            level: parseInt(levelStr.trim(), 10),
          });
        }
      }
    }

    // Check if all styles are standard headings (Heading1, Heading2, etc.)
    const standardHeadingPattern = /^Heading\s*(\d+)$/i;
    const allStandardHeadings = styles.every((style) =>
      standardHeadingPattern.test(style.styleName)
    );

    if (!allStandardHeadings) {
      return instruction; // Contains custom styles, keep \t switches
    }

    // Extract heading numbers and find the range
    const headingLevels = styles
      .map((style) => {
        const match = standardHeadingPattern.exec(style.styleName);
        return match ? parseInt(match[1]!, 10) : 0;
      })
      .filter((level) => level > 0)
      .sort((a, b) => a - b);

    if (headingLevels.length === 0) {
      return instruction; // No valid heading levels found
    }

    const minLevel = headingLevels[0]!;
    const maxLevel = headingLevels[headingLevels.length - 1]!;

    // Check if we should normalize (e.g., only "Heading 2,2," should become "1-3")
    // We normalize if:
    // 1. Only one heading style is referenced, OR
    // 2. The heading levels don't form a complete sequence from 1
    const shouldNormalize =
      headingLevels.length === 1 || // Single heading style
      minLevel !== 1 || // Doesn't start at Heading 1
      headingLevels.length !== maxLevel - minLevel + 1; // Not a complete sequence

    if (!shouldNormalize) {
      return instruction; // Already a complete sequence, keep as-is
    }

    // FIX: Use actual min/max levels from the \t switch, not hardcoded "1-N"
    // If \t "Heading 2,2," was specified, output should be \o "2-2", not \o "1-3"

    // Remove all \t switches and replace with \o switch
    let normalized = instruction.replace(tSwitchPattern, '').trim();

    // Insert \o switch after "TOC" using the actual min/max levels
    normalized = normalized.replace(/^TOC\s*/, `TOC \\o "${minLevel}-${maxLevel}" `);

    defaultLogger.debug(
      `[TOC Parser] Normalized field instruction from "${instruction}" to "${normalized}"`
    );

    return normalized;
  }

  /**
   * Helper to parse TOC from SDT content
   * Now uses two-tier extraction: ComplexField objects first, then raw XML fallback
   * Automatically normalizes incomplete \t switches to \o format for standard headings
   */
  private parseTOCFromSDTContent(
    content: any[],
    _properties: any,
    sdtContent: any
  ): TableOfContents | null {
    try {
      let title: string | undefined;
      let fieldInstruction: string | undefined;

      // Extract title from parsed content
      for (const element of content) {
        if (element instanceof Paragraph) {
          const style = element.getStyle();
          if (style === 'TOCHeading') {
            // Extract title text
            const runs = element.getRuns();
            title = runs.map((r) => r.getText()).join('');
          }
        }
      }

      // NEW: Extract field instruction from assembled ComplexField objects first
      fieldInstruction = this.extractInstructionFromContent(content);

      // FALLBACK: Use raw XML parsing if no ComplexField found
      if (!fieldInstruction) {
        defaultLogger.debug(
          '[TOC Parser] ComplexField extraction failed, falling back to raw XML parsing'
        );
        fieldInstruction = this.extractInstructionFromRawXML(sdtContent);
      }

      if (!fieldInstruction) {
        defaultLogger.warn(
          '[DocumentParser] No TOC field instruction found in SDT content (tried both ComplexField and raw XML)'
        );
        return null;
      }

      defaultLogger.debug(`[TOC Parser] Successfully extracted instruction: "${fieldInstruction}"`);

      // Decode HTML entities before parsing switches
      // XML stores quotes as &quot; which need to be converted for regex matching
      const decodedInstruction = fieldInstruction
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&');

      defaultLogger.debug(`[TOC Parser] Decoded instruction: "${decodedInstruction}"`);

      // NORMALIZATION: Convert incomplete \t switches to \o format for standard headings
      const normalizedInstruction = this.normalizeTOCFieldInstruction(decodedInstruction);

      // Parse field switches from normalized instruction
      const tocOptions: any = {
        title,
        originalFieldInstruction: normalizedInstruction.trim(), // Store normalized version
      };

      // Check for \h (hyperlinks)
      if (normalizedInstruction.includes('\\h')) {
        tocOptions.useHyperlinks = true;
      }

      // Check for \n (omit page numbers)
      if (normalizedInstruction.includes('\\n')) {
        tocOptions.showPageNumbers = false;
      }

      // Check for \z (hide in web layout)
      if (normalizedInstruction.includes('\\z')) {
        tocOptions.hideInWebLayout = true;
      }

      // Check for \o "x-y" (outline levels) - supports quoted and unquoted formats
      const outlineMatch = /\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/.exec(
        normalizedInstruction
      );
      if (outlineMatch) {
        // Extract captured groups from whichever format matched
        const minLevel = outlineMatch[1] || outlineMatch[3] || outlineMatch[5];
        const maxLevel = outlineMatch[2] || outlineMatch[4] || outlineMatch[6];
        if (minLevel && maxLevel) {
          tocOptions.minLevel = parseInt(minLevel, 10);
          tocOptions.maxLevel = parseInt(maxLevel, 10);
        }
      }

      // Check for \t "styles..." (include styles)
      const stylesMatch = /\\t\s+"([^"]+)"/.exec(normalizedInstruction);
      if (stylesMatch?.[1]) {
        const stylesStr = stylesMatch[1];
        const styles: { styleName: string; level: number }[] = [];

        // Parse "StyleName,Level,StyleName2,Level2,..."
        const parts = stylesStr.split(',').filter((p) => p.trim());
        for (let i = 0; i < parts.length; i += 2) {
          const styleName = parts[i];
          const levelStr = parts[i + 1];
          if (styleName && levelStr) {
            styles.push({
              styleName: styleName.trim(),
              level: parseInt(levelStr.trim(), 10),
            });
          }
        }

        if (styles.length > 0) {
          tocOptions.includeStyles = styles;
        }
      }

      return new TableOfContents(tocOptions);
    } catch (error: unknown) {
      defaultLogger.warn(
        '[DocumentParser] Failed to parse TOC from SDT content:',
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Helper to parse list items for combo box / dropdown per ECMA-376
   * Part 1 §17.5.2.13 CT_SdtListItem. `w:value` is required; both
   * `w:displayText` and `w:value` are ST_String so any string
   * (including the empty string) is legal.
   *
   * The previous truthy gate dropped legitimate list items whenever:
   *   - `w:value="0"` / `w:value="123"` — XMLParser coerces numeric
   *     strings to numbers; `0` fails the truthy check entirely, and
   *     storing a raw number instead of a string breaks the `ListItem`
   *     `value: string` contract downstream.
   *   - `w:displayText=""` — empty displayText is legal (e.g. a
   *     separator / blank choice); the gate dropped it.
   * The fix:
   *   - Gate on presence (`!== undefined`), not truthiness.
   *   - Coerce both attributes to `String(…)` so numeric-coerced
   *     attribute values serialise back to their original textual form.
   *   - Default missing `w:displayText` to the stringified `w:value`
   *     (the idiomatic Word fallback when authors author list items
   *     with only a value attribute).
   */
  private parseListItems(element: any): any {
    const items: any[] = [];
    const listItems = element?.['w:listItem'];
    const itemArray = Array.isArray(listItems) ? listItems : listItems ? [listItems] : [];

    for (const item of itemArray) {
      const rawValue = item['@_w:value'];
      if (rawValue === undefined) continue; // w:value is required by the schema
      const value = String(rawValue);
      const rawDisplay = item['@_w:displayText'];
      const displayText = rawDisplay === undefined ? value : String(rawDisplay);
      items.push({ displayText, value });
    }

    const rawLast = element?.['@_w:lastValue'];
    return {
      items,
      lastValue: rawLast === undefined ? undefined : String(rawLast),
    };
  }

  /**
   * Helper to convert object back to XML string
   */
  private objectToXml(obj: any): string {
    // XML reconstruction that preserves element order using _orderedChildren
    // FIX (v1.3.1): Use _orderedChildren to maintain document order of elements
    // This fixes TOC tab preservation - tabs must be in correct position
    const buildXml = (o: any, name?: string): string => {
      if (typeof o === 'string') return o;
      if (typeof o !== 'object') return String(o);

      const keys = Object.keys(o);

      // FIX: If a name is provided, we're building a specific element (possibly self-closing)
      // Don't return empty string for empty objects with a name - they should become self-closing tags
      if (keys.length === 0 && !name) return '';

      const tagName = name || keys[0]!; // keys[0] is guaranteed to exist due to length check (or name is provided)
      const element = name ? o : o[tagName];

      let xml = `<${tagName}`;

      // Add attributes
      if (element && typeof element === 'object') {
        for (const key of Object.keys(element)) {
          if (key.startsWith('@_')) {
            const attrName = key.substring(2);
            xml += ` ${attrName}="${element[key]}"`;
          }
        }
      }

      // Check for children
      const hasChildren =
        element &&
        typeof element === 'object' &&
        Object.keys(element).some(
          (k) => !k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren'
        );

      if (!hasChildren && !element?.['#text']) {
        xml += '/>';
      } else {
        xml += '>';

        // Add text content
        if (element?.['#text']) {
          xml += element['#text'];
        }

        // Add child elements using _orderedChildren if available
        if (element && typeof element === 'object') {
          const orderedChildren = element._orderedChildren as
            | { type: string; index: number }[]
            | undefined;

          if (orderedChildren && orderedChildren.length > 0) {
            // Use _orderedChildren to preserve element order
            for (const childInfo of orderedChildren) {
              const childType = childInfo.type;
              const childIndex = childInfo.index;

              if (element[childType] !== undefined) {
                const children = element[childType];

                if (Array.isArray(children)) {
                  if (childIndex < children.length) {
                    const childXml = buildXml(children[childIndex], childType);
                    xml += childXml;
                  }
                } else {
                  // Single child element
                  if (childIndex === 0) {
                    const childXml = buildXml(children, childType);
                    xml += childXml;
                  }
                }
              }
            }
          } else {
            // Fallback: iterate through keys if no _orderedChildren
            for (const key of Object.keys(element)) {
              if (!key.startsWith('@_') && key !== '#text' && key !== '_orderedChildren') {
                const children = element[key];
                if (Array.isArray(children)) {
                  for (const child of children) {
                    xml += buildXml(child, key);
                  }
                } else {
                  xml += buildXml(children, key);
                }
              }
            }
          }
        }

        xml += `</${tagName}>`;
      }

      return xml;
    };

    return buildXml(obj);
  }

  /**
   * Parses existing relationships from word/_rels/document.xml.rels
   * This ensures new relationships don't collide with existing IDs
   * Returns the parsed RelationshipManager if found, otherwise returns the provided one
   */
  private parseRelationships(
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager
  ): RelationshipManager {
    const relsPath = 'word/_rels/document.xml.rels';
    const relsXml = zipHandler.getFileAsString(relsPath);

    if (relsXml) {
      // Parse and replace the relationship manager with populated one
      return RelationshipManager.fromXml(relsXml);
    }

    // No existing relationships - return the provided manager
    // The Document class will add default relationships
    return relationshipManager;
  }

  /**
   * Parses document properties from core.xml, app.xml, and custom.xml
   */
  private parseProperties(zipHandler: ZipHandler): DocumentProperties {
    const extractTag = (xml: string, tag: string): string | undefined => {
      const tagContent = XMLParser.extractBetweenTags(xml, `<${tag}`, `</${tag}>`);
      return tagContent ? XMLBuilder.unescapeXml(tagContent) : undefined;
    };

    const properties: DocumentProperties = {};

    // Parse core.xml (core properties)
    const coreXml = zipHandler.getFileAsString(DOCX_PATHS.CORE_PROPS);
    if (coreXml) {
      properties.title = extractTag(coreXml, 'dc:title');
      properties.subject = extractTag(coreXml, 'dc:subject');
      properties.creator = extractTag(coreXml, 'dc:creator');
      properties.keywords = extractTag(coreXml, 'cp:keywords');
      properties.description = extractTag(coreXml, 'dc:description');
      properties.lastModifiedBy = extractTag(coreXml, 'cp:lastModifiedBy');

      // Phase 5.5 - Extended core properties
      properties.category = extractTag(coreXml, 'cp:category');
      properties.contentStatus = extractTag(coreXml, 'cp:contentStatus');
      properties.language = extractTag(coreXml, 'dc:language');

      // Parse revision as number
      const revisionStr = extractTag(coreXml, 'cp:revision');
      if (revisionStr) {
        properties.revision = parseInt(revisionStr, 10);
      }

      // Parse dates
      const createdStr = extractTag(coreXml, 'dcterms:created');
      if (createdStr) {
        properties.created = new Date(createdStr);
      }

      const modifiedStr = extractTag(coreXml, 'dcterms:modified');
      if (modifiedStr) {
        properties.modified = new Date(modifiedStr);
      }
    }

    // Parse app.xml (extended properties)
    const appXml = zipHandler.getFileAsString(DOCX_PATHS.APP_PROPS);
    if (appXml) {
      properties.application = extractTag(appXml, 'Application');
      properties.appVersion = extractTag(appXml, 'AppVersion');
      properties.company = extractTag(appXml, 'Company');
      properties.manager = extractTag(appXml, 'Manager');

      // Also check for version field
      if (!properties.appVersion) {
        properties.version = extractTag(appXml, 'Version');
      }
    }

    // Parse custom.xml (custom properties)
    const customXml = zipHandler.getFileAsString('docProps/custom.xml');
    if (customXml) {
      properties.customProperties = this.parseCustomProperties(customXml);
    }

    return properties;
  }

  /**
   * Parses custom properties from custom.xml
   */
  private parseCustomProperties(xml: string): Record<string, string | number | boolean | Date> {
    const customProps: Record<string, string | number | boolean | Date> = {};

    // Extract all property elements
    const propertyElements = XMLParser.extractElements(xml, 'property');

    for (const propXml of propertyElements) {
      // Extract name attribute
      const nameMatch = /name="([^"]+)"/.exec(propXml);
      if (!nameMatch?.[1]) continue;
      const name = XMLBuilder.unescapeXml(nameMatch[1]);

      // Determine value type and extract value
      if (propXml.includes('<vt:lpwstr>')) {
        const value = XMLParser.extractBetweenTags(propXml, '<vt:lpwstr>', '</vt:lpwstr>');
        if (value !== undefined) {
          customProps[name] = XMLBuilder.unescapeXml(value);
        }
      } else if (propXml.includes('<vt:r8>')) {
        const value = XMLParser.extractBetweenTags(propXml, '<vt:r8>', '</vt:r8>');
        if (value !== undefined) {
          customProps[name] = parseFloat(value);
        }
      } else if (propXml.includes('<vt:bool>')) {
        const value = XMLParser.extractBetweenTags(propXml, '<vt:bool>', '</vt:bool>');
        if (value !== undefined) {
          customProps[name] = value === 'true';
        }
      } else if (propXml.includes('<vt:filetime>')) {
        const value = XMLParser.extractBetweenTags(propXml, '<vt:filetime>', '</vt:filetime>');
        if (value !== undefined) {
          customProps[name] = new Date(value);
        }
      }
    }

    return customProps;
  }

  /**
   * Parses styles from styles.xml
   * @param zipHandler - ZIP handler containing the document
   * @returns Array of parsed Style objects
   */
  private parseStyles(zipHandler: ZipHandler): Style[] {
    const styles: Style[] = [];
    const stylesXml = zipHandler.getFileAsString(DOCX_PATHS.STYLES);

    if (!stylesXml) {
      return styles; // No styles.xml file
    }

    try {
      // Extract all <w:style> elements using XMLParser
      const styleElements = XMLParser.extractElements(stylesXml, 'w:style');

      for (const styleXml of styleElements) {
        try {
          const style = this.parseStyle(styleXml);
          if (style) {
            styles.push(style);
          }
        } catch (error: unknown) {
          const err = error instanceof Error ? error : new Error(String(error));
          this.parseErrors.push({ element: 'style', error: err });

          if (this.strictParsing) {
            throw error;
          }
          // In lenient mode, continue parsing other styles
        }
      }
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'styles.xml', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse styles.xml: ${err.message}`);
      }
    }

    return styles;
  }

  /**
   * Parses numbering definitions from numbering.xml
   * Extracts abstract numbering definitions and numbering instances
   * @param zipHandler - ZIP handler containing the document
   * @returns Object containing abstractNumberings and numberingInstances arrays
   */
  private parseNumbering(zipHandler: ZipHandler): {
    abstractNumberings: AbstractNumbering[];
    numberingInstances: NumberingInstance[];
  } {
    const abstractNumberings: AbstractNumbering[] = [];
    const numberingInstances: NumberingInstance[] = [];

    const numberingXml = zipHandler.getFileAsString(DOCX_PATHS.NUMBERING);

    if (!numberingXml) {
      return { abstractNumberings, numberingInstances }; // No numbering.xml file
    }

    try {
      // Extract all <w:abstractNum> elements
      const abstractNumElements = XMLParser.extractElements(numberingXml, 'w:abstractNum');

      for (const abstractNumXml of abstractNumElements) {
        try {
          const abstractNum = AbstractNumbering.fromXML(abstractNumXml);
          abstractNumberings.push(abstractNum);
        } catch (error: unknown) {
          const err = error instanceof Error ? error : new Error(String(error));
          this.parseErrors.push({ element: 'abstractNum', error: err });

          if (this.strictParsing) {
            throw error;
          }
          // In lenient mode, continue parsing other abstract numberings
        }
      }

      // Extract all <w:num> elements (numbering instances)
      const numElements = XMLParser.extractElements(numberingXml, 'w:num');

      for (const numXml of numElements) {
        try {
          const instance = NumberingInstance.fromXML(numXml);
          numberingInstances.push(instance);
        } catch (error: unknown) {
          const err = error instanceof Error ? error : new Error(String(error));
          this.parseErrors.push({ element: 'num', error: err });

          if (this.strictParsing) {
            throw error;
          }
          // In lenient mode, continue parsing other instances
        }
      }
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'numbering.xml', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse numbering.xml: ${err.message}`);
      }
    }

    return { abstractNumberings, numberingInstances };
  }

  /**
   * Parses section properties from document XML
   * @param docXml - Document XML content
   * @returns Parsed Section object or null if not found
   */
  private parseSectionProperties(docXml: string): Section | null {
    try {
      // Extract the final <w:sectPr> from <w:body>
      const bodyElements = XMLParser.extractElements(docXml, 'w:body');
      if (bodyElements.length === 0) {
        return null;
      }
      const bodyContent = bodyElements[0];
      if (!bodyContent) {
        return null;
      }

      // Find the body-level sectPr (direct child of w:body), NOT inline sectPr inside w:pPr.
      // Per ECMA-376, the body-level sectPr appears as the last child of w:body,
      // after all paragraphs and tables. Inline sectPr elements inside w:pPr define
      // section breaks and must not be confused with the document-level section properties.
      // Search only in the content after the last block-level element to avoid picking up
      // inline sectPr from paragraph properties.
      const lastPClose = bodyContent.lastIndexOf('</w:p>');
      const lastTblClose = bodyContent.lastIndexOf('</w:tbl>');
      const lastSdtClose = bodyContent.lastIndexOf('</w:sdt>');
      const lastBlockEnd = Math.max(lastPClose, lastTblClose, lastSdtClose);

      let sectPr: string | undefined;
      if (lastBlockEnd !== -1) {
        // Search for sectPr only after the last block element
        const tailContent = bodyContent.substring(lastBlockEnd);
        const sectPrElements = XMLParser.extractElements(tailContent, 'w:sectPr');
        if (sectPrElements.length > 0) {
          sectPr = sectPrElements[0];
        }
      }
      if (!sectPr) {
        // Fallback: no block elements found, or sectPr not found after last block.
        // Search the entire body content (covers edge cases like empty documents).
        const sectPrElements = XMLParser.extractElements(bodyContent, 'w:sectPr');
        if (sectPrElements.length === 0) {
          return null;
        }
        sectPr = sectPrElements[sectPrElements.length - 1];
        if (!sectPr) {
          return null;
        }
      }

      const sectionProps: SectionProperties = {};

      // Parse page size
      const pgSzElements = XMLParser.extractElements(sectPr, 'w:pgSz');
      if (pgSzElements.length > 0) {
        const pgSz = pgSzElements[0];
        if (pgSz) {
          const width = XMLParser.extractAttribute(pgSz, 'w:w');
          const height = XMLParser.extractAttribute(pgSz, 'w:h');
          const orient = XMLParser.extractAttribute(pgSz, 'w:orient');
          const code = XMLParser.extractAttribute(pgSz, 'w:code');

          if (width && height) {
            sectionProps.pageSize = {
              width: parseInt(width, 10),
              height: parseInt(height, 10),
              orientation: orient === 'landscape' ? 'landscape' : 'portrait',
              code: code ? parseInt(code, 10) : undefined,
            };
          }
        }
      }

      // Parse margins
      const pgMarElements = XMLParser.extractElements(sectPr, 'w:pgMar');
      if (pgMarElements.length > 0) {
        const pgMar = pgMarElements[0];
        if (pgMar) {
          const top = XMLParser.extractAttribute(pgMar, 'w:top');
          const bottom = XMLParser.extractAttribute(pgMar, 'w:bottom');
          const left = XMLParser.extractAttribute(pgMar, 'w:left');
          const right = XMLParser.extractAttribute(pgMar, 'w:right');
          const header = XMLParser.extractAttribute(pgMar, 'w:header');
          const footer = XMLParser.extractAttribute(pgMar, 'w:footer');
          const gutter = XMLParser.extractAttribute(pgMar, 'w:gutter');

          if (top && bottom && left && right) {
            sectionProps.margins = {
              top: parseInt(top, 10),
              bottom: parseInt(bottom, 10),
              left: parseInt(left, 10),
              right: parseInt(right, 10),
              header: header ? parseInt(header, 10) : undefined,
              footer: footer ? parseInt(footer, 10) : undefined,
              gutter: gutter ? parseInt(gutter, 10) : undefined,
            };
          }
        }
      }

      // Parse page borders (w:pgBorders) per ECMA-376 Part 1 §17.6.10
      const pgBordersElements = XMLParser.extractElements(sectPr, 'w:pgBorders');
      if (pgBordersElements.length > 0) {
        const pgBordersXml = pgBordersElements[0];
        if (pgBordersXml) {
          const pageBorders: any = {};
          const offsetFrom = XMLParser.extractAttribute(pgBordersXml, 'w:offsetFrom');
          if (offsetFrom) pageBorders.offsetFrom = offsetFrom;
          const display = XMLParser.extractAttribute(pgBordersXml, 'w:display');
          if (display) pageBorders.display = display;
          const zOrder = XMLParser.extractAttribute(pgBordersXml, 'w:zOrder');
          if (zOrder) pageBorders.zOrder = zOrder;

          const parseBorder = (sideXml: string): any | undefined => {
            if (!sideXml) return undefined;
            const border: any = {};
            const val = XMLParser.extractAttribute(sideXml, 'w:val');
            if (val) border.style = val;
            const sz = XMLParser.extractAttribute(sideXml, 'w:sz');
            if (sz) border.size = parseInt(sz.toString(), 10);
            const color = XMLParser.extractAttribute(sideXml, 'w:color');
            if (color) border.color = color;
            const space = XMLParser.extractAttribute(sideXml, 'w:space');
            if (space) border.space = parseInt(space.toString(), 10);
            // w:shadow and w:frame are ST_OnOff per ECMA-376 §17.17.4.
            // Use `!== undefined` gating so explicit-false survives round-trip
            // (previous code only stored `true`, silently dropping `w:shadow="0"`).
            const shadow = XMLParser.extractAttribute(sideXml, 'w:shadow');
            if (shadow !== undefined) border.shadow = parseOnOffAttribute(shadow, true);
            const frame = XMLParser.extractAttribute(sideXml, 'w:frame');
            if (frame !== undefined) border.frame = parseOnOffAttribute(frame, true);
            const themeColor = XMLParser.extractAttribute(sideXml, 'w:themeColor');
            if (themeColor) border.themeColor = themeColor;
            // Theme tint / shade per §17.18.82 — CT_TopBorder/CT_BottomBorder extend
            // CT_Border so inherit the full themed-color attribute set.
            const themeTint = XMLParser.extractAttribute(sideXml, 'w:themeTint');
            if (themeTint) border.themeTint = themeTint;
            const themeShade = XMLParser.extractAttribute(sideXml, 'w:themeShade');
            if (themeShade) border.themeShade = themeShade;
            const artId = XMLParser.extractAttribute(sideXml, 'w:id');
            if (artId) border.artId = parseInt(artId.toString(), 10);
            return Object.keys(border).length > 0 ? border : undefined;
          };

          const sides = ['top', 'left', 'bottom', 'right'];
          for (const side of sides) {
            const sideElements = XMLParser.extractElements(pgBordersXml, `w:${side}`);
            if (sideElements.length > 0 && sideElements[0]) {
              const border = parseBorder(sideElements[0]);
              if (border) pageBorders[side] = border;
            }
          }

          if (Object.keys(pageBorders).length > 0) {
            sectionProps.pageBorders = pageBorders;
          }
        }
      }

      // Parse columns per ECMA-376 §17.6.4 CT_Columns. Every attribute
      // (num / sep / space / equalWidth) is optional with spec-defined
      // defaults — num defaults to 1, equalWidth to true, sep to false,
      // space to 720 twips. The previous `if (num)` gate silently dropped
      // every `<w:cols>` that relied on the default num=1 (e.g. a bare
      // `<w:cols w:sep="1" w:space="720"/>` specifying a single column
      // with a separator), which is the exact form Word emits when the
      // user toggles the column separator without changing column count.
      const colsElements = XMLParser.extractElements(sectPr, 'w:cols');
      if (colsElements.length > 0) {
        const cols = colsElements[0];
        if (cols) {
          const num = XMLParser.extractAttribute(cols, 'w:num');
          const space = XMLParser.extractAttribute(cols, 'w:space');
          const equalWidth = XMLParser.extractAttribute(cols, 'w:equalWidth');
          const sep = XMLParser.extractAttribute(cols, 'w:sep');

          // Extract individual column widths and per-column spacing (CT_Column: w:w, w:space)
          const colElements = XMLParser.extractElements(cols, 'w:col');
          const columnWidths: number[] = [];
          const columnSpaces: number[] = [];
          let hasColumnSpaces = false;
          for (const col of colElements) {
            const width = XMLParser.extractAttribute(col, 'w:w');
            if (width) {
              columnWidths.push(parseInt(width.toString(), 10));
            }
            const colSpace = XMLParser.extractAttribute(col, 'w:space');
            if (colSpace) {
              columnSpaces.push(parseInt(colSpace.toString(), 10));
              hasColumnSpaces = true;
            } else {
              columnSpaces.push(0);
            }
          }

          // Spec default for num is 1; fall back to column-count from
          // child `<w:col>` children when available (the expanded per-column
          // form), otherwise to the literal default.
          const count = num
            ? parseInt(num.toString(), 10)
            : columnWidths.length > 0
              ? columnWidths.length
              : 1;

          sectionProps.columns = {
            count,
            space: space ? parseInt(space.toString(), 10) : undefined,
            equalWidth: equalWidth ? parseOnOffAttribute(equalWidth) : undefined,
            separator: sep ? parseOnOffAttribute(sep) : undefined,
            columnWidths: columnWidths.length > 0 ? columnWidths : undefined,
            columnSpaces: hasColumnSpaces ? columnSpaces : undefined,
          };
        }
      }

      // Parse section type
      const typeElements = XMLParser.extractElements(sectPr, 'w:type');
      if (typeElements.length > 0) {
        const type = typeElements[0];
        if (type) {
          const typeVal = XMLParser.extractAttribute(type, 'w:val') as SectionType;
          if (typeVal) {
            sectionProps.type = typeVal;
          }
        }
      }

      // Parse page numbering
      const pgNumTypeElements = XMLParser.extractElements(sectPr, 'w:pgNumType');
      if (pgNumTypeElements.length > 0) {
        const pgNumType = pgNumTypeElements[0];
        if (pgNumType) {
          const start = XMLParser.extractAttribute(pgNumType, 'w:start');
          const fmt = XMLParser.extractAttribute(pgNumType, 'w:fmt')!;

          sectionProps.pageNumbering = {
            start: start ? parseInt(start, 10) : undefined,
            format: fmt,
          };

          // Parse chapter numbering (w:chapStyle, w:chapSep)
          const chapStyle = XMLParser.extractAttribute(pgNumType, 'w:chapStyle');
          if (chapStyle) {
            sectionProps.chapStyle = parseInt(chapStyle, 10);
          }
          const chapSep = XMLParser.extractAttribute(pgNumType, 'w:chapSep');
          if (chapSep) {
            sectionProps.chapSep = chapSep as any;
          }
        }
      }

      // Parse title page flag (w:titlePg) — CT_OnOff per ECMA-376 §17.6.23;
      // honour w:val so an explicit `w:val="0"` override of an inherited
      // true is not silently flipped to true.
      const titlePgEls = XMLParser.extractElements(sectPr, 'w:titlePg');
      if (titlePgEls.length > 0 && titlePgEls[0]) {
        const v = XMLParser.extractAttribute(titlePgEls[0], 'w:val');
        sectionProps.titlePage = parseOnOffAttribute(v, true);
      }

      // Parse header references
      const headerRefs = XMLParser.extractElements(sectPr, 'w:headerReference');
      if (headerRefs.length > 0) {
        sectionProps.headers = {};
        for (const headerRef of headerRefs) {
          const type = XMLParser.extractAttribute(headerRef, 'w:type');
          const rId = XMLParser.extractAttribute(headerRef, 'r:id');
          if (type && rId) {
            if (type === 'default') sectionProps.headers.default = rId;
            else if (type === 'first') sectionProps.headers.first = rId;
            else if (type === 'even') sectionProps.headers.even = rId;
          }
        }
      }

      // Parse footer references
      const footerRefs = XMLParser.extractElements(sectPr, 'w:footerReference');
      if (footerRefs.length > 0) {
        sectionProps.footers = {};
        for (const footerRef of footerRefs) {
          const type = XMLParser.extractAttribute(footerRef, 'w:type');
          const rId = XMLParser.extractAttribute(footerRef, 'r:id');
          if (type && rId) {
            if (type === 'default') sectionProps.footers.default = rId;
            else if (type === 'first') sectionProps.footers.first = rId;
            else if (type === 'even') sectionProps.footers.even = rId;
          }
        }
      }

      // Parse vertical alignment
      const vAlignElements = XMLParser.extractElements(sectPr, 'w:vAlign');
      if (vAlignElements.length > 0) {
        const vAlign = vAlignElements[0];
        if (vAlign) {
          const val = XMLParser.extractAttribute(vAlign, 'w:val');
          if (val) {
            sectionProps.verticalAlignment = val as 'top' | 'center' | 'bottom' | 'both';
          }
        }
      }

      // Parse paper source
      const paperSrcElements = XMLParser.extractElements(sectPr, 'w:paperSrc');
      if (paperSrcElements.length > 0) {
        const paperSrc = paperSrcElements[0];
        if (paperSrc) {
          const first = XMLParser.extractAttribute(paperSrc, 'w:first');
          const other = XMLParser.extractAttribute(paperSrc, 'w:other');

          if (first || other) {
            sectionProps.paperSource = {
              first: first ? parseInt(first.toString(), 10) : undefined,
              other: other ? parseInt(other.toString(), 10) : undefined,
            };
          }
        }
      }

      // Parse text direction
      const textDirElements = XMLParser.extractElements(sectPr, 'w:textDirection');
      if (textDirElements.length > 0) {
        const textDir = textDirElements[0];
        if (textDir) {
          const val = XMLParser.extractAttribute(textDir, 'w:val');
          if (val) {
            // Reverse-map OOXML value "lrTb" to API shorthand "ltr".
            // "tbRl" and "btLr" are used directly in both OOXML and API.
            const reverseTextDirMap: Record<string, string> = {
              lrTb: 'ltr',
            };
            sectionProps.textDirection = (reverseTextDirMap[val] || val) as
              | 'ltr'
              | 'rtl'
              | 'tbRl'
              | 'btLr';
          }
        }
      }

      // Parse bidi (w:bidi) — CT_OnOff per ECMA-376 §17.6.1 (RTL section)
      const bidiEls = XMLParser.extractElements(sectPr, 'w:bidi');
      if (bidiEls.length > 0 && bidiEls[0]) {
        const v = XMLParser.extractAttribute(bidiEls[0], 'w:val');
        sectionProps.bidi = parseOnOffAttribute(v, true);
      }

      // Parse RTL gutter (w:rtlGutter) — CT_OnOff per ECMA-376 §17.6.16
      const rtlGutterEls = XMLParser.extractElements(sectPr, 'w:rtlGutter');
      if (rtlGutterEls.length > 0 && rtlGutterEls[0]) {
        const v = XMLParser.extractAttribute(rtlGutterEls[0], 'w:val');
        sectionProps.rtlGutter = parseOnOffAttribute(v, true);
      }

      // Parse document grid (w:docGrid)
      const docGridElements = XMLParser.extractElements(sectPr, 'w:docGrid');
      if (docGridElements.length > 0) {
        const docGrid = docGridElements[0];
        if (docGrid) {
          const gridType = XMLParser.extractAttribute(docGrid, 'w:type');
          const linePitch = XMLParser.extractAttribute(docGrid, 'w:linePitch');
          const charSpace = XMLParser.extractAttribute(docGrid, 'w:charSpace');
          sectionProps.docGrid = {};
          if (gridType) sectionProps.docGrid.type = gridType as any;
          if (linePitch) sectionProps.docGrid.linePitch = parseInt(linePitch, 10);
          if (charSpace) sectionProps.docGrid.charSpace = parseInt(charSpace, 10);
        }
      }

      // Parse line numbering (w:lnNumType)
      const lnNumElements = XMLParser.extractElements(sectPr, 'w:lnNumType');
      if (lnNumElements.length > 0) {
        const lnNum = lnNumElements[0];
        if (lnNum) {
          const countBy = XMLParser.extractAttribute(lnNum, 'w:countBy');
          const start = XMLParser.extractAttribute(lnNum, 'w:start');
          const distance = XMLParser.extractAttribute(lnNum, 'w:distance');
          const restart = XMLParser.extractAttribute(lnNum, 'w:restart');
          sectionProps.lineNumbering = {};
          if (countBy) sectionProps.lineNumbering.countBy = parseInt(countBy, 10);
          if (start) sectionProps.lineNumbering.start = parseInt(start, 10);
          if (distance) sectionProps.lineNumbering.distance = parseInt(distance, 10);
          if (restart) sectionProps.lineNumbering.restart = restart as any;
        }
      }

      // Helper to extract a single attribute from a child element
      const extractChildAttr = (
        parentXml: string,
        childTag: string,
        attr: string
      ): string | undefined => {
        const els = XMLParser.extractElements(parentXml, childTag);
        if (els.length > 0 && els[0]) return XMLParser.extractAttribute(els[0], attr);
        return undefined;
      };

      // Parse footnote properties (w:footnotePr)
      const footnotePrElements = XMLParser.extractElements(sectPr, 'w:footnotePr');
      if (footnotePrElements.length > 0 && footnotePrElements[0]) {
        const fnPr = footnotePrElements[0];
        const props: any = {};
        const pos = extractChildAttr(fnPr, 'w:pos', 'w:val');
        if (pos) props.position = pos;
        const fmt = extractChildAttr(fnPr, 'w:numFmt', 'w:val');
        if (fmt) props.numberFormat = fmt;
        const startVal = extractChildAttr(fnPr, 'w:numStart', 'w:val');
        if (startVal) props.startNumber = parseInt(startVal, 10);
        const restart = extractChildAttr(fnPr, 'w:numRestart', 'w:val');
        if (restart) props.restart = restart;
        if (Object.keys(props).length > 0) sectionProps.footnotePr = props;
      }

      // Parse endnote properties (w:endnotePr)
      const endnotePrElements = XMLParser.extractElements(sectPr, 'w:endnotePr');
      if (endnotePrElements.length > 0 && endnotePrElements[0]) {
        const enPr = endnotePrElements[0];
        const props: any = {};
        const pos = extractChildAttr(enPr, 'w:pos', 'w:val');
        if (pos) props.position = pos;
        const fmt = extractChildAttr(enPr, 'w:numFmt', 'w:val');
        if (fmt) props.numberFormat = fmt;
        const startVal = extractChildAttr(enPr, 'w:numStart', 'w:val');
        if (startVal) props.startNumber = parseInt(startVal, 10);
        const restart = extractChildAttr(enPr, 'w:numRestart', 'w:val');
        if (restart) props.restart = restart;
        if (Object.keys(props).length > 0) sectionProps.endnotePr = props;
      }

      // Parse noEndnote (w:noEndnote) — CT_OnOff per ECMA-376 §17.11.14
      const noEndEls = XMLParser.extractElements(sectPr, 'w:noEndnote');
      if (noEndEls.length > 0 && noEndEls[0]) {
        const v = XMLParser.extractAttribute(noEndEls[0], 'w:val');
        sectionProps.noEndnote = parseOnOffAttribute(v, true);
      }

      // Parse form protection (w:formProt) — CT_OnOff per ECMA-376 §17.6.8
      const formProtEls = XMLParser.extractElements(sectPr, 'w:formProt');
      if (formProtEls.length > 0 && formProtEls[0]) {
        const v = XMLParser.extractAttribute(formProtEls[0], 'w:val');
        sectionProps.formProt = parseOnOffAttribute(v, true);
      }

      // Parse printer settings (w:printerSettings r:id)
      const printerElements = XMLParser.extractElements(sectPr, 'w:printerSettings');
      if (printerElements.length > 0) {
        const printer = printerElements[0];
        if (printer) {
          const rId = XMLParser.extractAttribute(printer, 'r:id');
          if (rId) sectionProps.printerSettingsId = rId;
        }
      }

      const section = new Section(sectionProps);

      // Parse section property change (w:sectPrChange) per ECMA-376 Part 1 §17.13.5.32
      const sectPrChangeElements = XMLParser.extractElements(sectPr, 'w:sectPrChange');
      if (sectPrChangeElements.length > 0 && sectPrChangeElements[0]) {
        const changeXml = sectPrChangeElements[0];
        const id = XMLParser.extractAttribute(changeXml, 'w:id') || '0';
        const author = XMLParser.extractAttribute(changeXml, 'w:author') || '';
        const date = XMLParser.extractAttribute(changeXml, 'w:date') || '';

        // Extract the previous w:sectPr child
        const prevSectPrElements = XMLParser.extractElements(changeXml, 'w:sectPr');
        const prevSectPrXml = prevSectPrElements.length > 0 ? prevSectPrElements[0] : undefined;
        const prevProps = prevSectPrXml ? this.parsePreviousSectionProperties(prevSectPrXml) : {};

        section.setSectPrChange({ id, author, date, previousProperties: prevProps });
      }

      return section;
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'sectPr', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse section properties: ${err.message}`);
      }

      return null;
    }
  }

  /**
   * Parses a single style element from XML
   * @param styleXml - XML string of a single w:style element
   * @returns Parsed Style object or null if invalid
   */
  private parseStyle(styleXml: string): Style | null {
    // Extract style attributes
    const typeAttr = XMLParser.extractAttribute(styleXml, 'w:type') as StyleType;
    const styleId = XMLParser.extractAttribute(styleXml, 'w:styleId') || '';
    const defaultAttr = XMLParser.extractAttribute(styleXml, 'w:default');
    const customStyleAttr = XMLParser.extractAttribute(styleXml, 'w:customStyle');

    if (!styleId || !typeAttr) {
      return null; // Invalid style, missing required attributes
    }

    // Extract style name (handles both self-closing and non-self-closing tags)
    // First try self-closing tag: <w:name w:val="List Paragraph"/>
    let name: string = styleId;
    const selfClosingNameAttrs = XMLParser.extractSelfClosingTag(styleXml, 'w:name');
    if (selfClosingNameAttrs) {
      // Got attributes like: ' w:val="List Paragraph"'
      const extractedName = XMLParser.extractAttribute(`<w:name${selfClosingNameAttrs}/>`, 'w:val');
      if (extractedName) {
        name = extractedName;
      }
    } else {
      // Try non-self-closing: <w:name w:val="..."></w:name>
      const nameElement = XMLParser.extractBetweenTags(styleXml, '<w:name', '</w:name>');
      if (nameElement) {
        const extractedName = XMLParser.extractAttribute(`<w:name${nameElement}`, 'w:val');
        if (extractedName) {
          name = extractedName;
        }
      }
    }

    // Extract basedOn
    const basedOnElement = XMLParser.extractBetweenTags(styleXml, '<w:basedOn', '</w:basedOn>');
    const basedOn = basedOnElement
      ? XMLParser.extractAttribute(`<w:basedOn${basedOnElement}`, 'w:val')
      : undefined;

    // Extract next
    const nextElement = XMLParser.extractBetweenTags(styleXml, '<w:next', '</w:next>');
    const next = nextElement
      ? XMLParser.extractAttribute(`<w:next${nextElement}`, 'w:val')
      : undefined;

    // Parse paragraph formatting (w:pPr)
    let paragraphFormatting: ParagraphFormatting | undefined;
    let styleNumPr: { numId?: number; ilvl?: number } | undefined;
    const pPrXml = XMLParser.extractBetweenTags(styleXml, '<w:pPr>', '</w:pPr>');
    if (pPrXml) {
      paragraphFormatting = this.parseParagraphFormattingFromXml(pPrXml);

      // Parse numPr (numbering properties) from style's paragraph properties
      // Styles like ListParagraph can inherit list formatting via numPr
      const numPrXml = XMLParser.extractBetweenTags(pPrXml, '<w:numPr>', '</w:numPr>');
      if (numPrXml) {
        styleNumPr = {};
        const numIdMatch = /<w:numId[^>]*w:val="(\d+)"/.exec(numPrXml);
        if (numIdMatch?.[1]) {
          styleNumPr.numId = parseInt(numIdMatch[1], 10);
        }
        const ilvlMatch = /<w:ilvl[^>]*w:val="(\d+)"/.exec(numPrXml);
        if (ilvlMatch?.[1]) {
          styleNumPr.ilvl = parseInt(ilvlMatch[1], 10);
        }
      }
    }

    // Parse run formatting (w:rPr)
    let runFormatting: RunFormatting | undefined;
    const rPrXml = XMLParser.extractBetweenTags(styleXml, '<w:rPr>', '</w:rPr>');
    if (rPrXml) {
      runFormatting = this.parseRunFormattingFromXml(rPrXml);
    }

    // Parse metadata CT_OnOff flags per ECMA-376 §17.7.4 (OnOffType bindings).
    // Each flag honours `w:val` so an explicit `<w:qFormat w:val="0"/>` override
    // of a based-on style's qFormat=true round-trips as `false`. The old code
    // detected presence via `styleXml.includes('<w:qFormat/>')` which ignored
    // w:val entirely and flipped any explicit-false to true.
    const parseStyleOnOffFlag = (tagName: string): boolean | undefined => {
      const els = XMLParser.extractElements(styleXml, tagName);
      if (els.length === 0 || !els[0]) return undefined;
      const v = XMLParser.extractAttribute(els[0], 'w:val');
      return parseOnOffAttribute(v, true);
    };

    const qFormat = parseStyleOnOffFlag('w:qFormat');
    const semiHidden = parseStyleOnOffFlag('w:semiHidden');
    const unhideWhenUsed = parseStyleOnOffFlag('w:unhideWhenUsed');
    const locked = parseStyleOnOffFlag('w:locked');
    const personal = parseStyleOnOffFlag('w:personal');
    const personalCompose = parseStyleOnOffFlag('w:personalCompose');
    const personalReply = parseStyleOnOffFlag('w:personalReply');
    const autoRedefine = parseStyleOnOffFlag('w:autoRedefine');
    // `<w:hidden>` (CT_Style §17.7.4, OnOffOnlyType) — completely hide the
    // style. Previously not modeled; now round-trips as `properties.hidden`.
    const hidden = parseStyleOnOffFlag('w:hidden');

    // `<w:rsid w:val="HEX"/>` (CT_Style §17.7.4, CT_LongHexNumber §17.18.50) —
    // revision-save ID stamp identifying the session in which this style
    // definition was last edited. Schema position: between `personalReply`
    // and `pPr`. Previously dropped entirely on parse, now preserved on
    // StyleProperties so round-trips stay faithful.
    let styleRsid: string | undefined;
    if (styleXml.includes('<w:rsid')) {
      const rsidTag = XMLParser.extractSelfClosingTag(styleXml, 'w:rsid');
      if (rsidTag) {
        const v = XMLParser.extractAttribute(`<w:rsid${rsidTag}`, 'w:val');
        if (v && v.length > 0) {
          styleRsid = v;
        }
      }
    }

    // uiPriority - Sort order
    let uiPriority: number | undefined;
    if (styleXml.includes('<w:uiPriority')) {
      const uiPriorityStart = styleXml.indexOf('<w:uiPriority');
      const uiPriorityEnd = styleXml.indexOf('/>', uiPriorityStart);
      if (uiPriorityEnd !== -1) {
        const uiPriorityTag = styleXml.substring(uiPriorityStart, uiPriorityEnd + 2);
        const valStr = XMLParser.extractAttribute(uiPriorityTag, 'w:val');
        if (valStr) {
          uiPriority = parseInt(valStr, 10);
        }
      }
    }

    // link - Linked character/paragraph style
    let link: string | undefined;
    if (styleXml.includes('<w:link')) {
      const linkStart = styleXml.indexOf('<w:link');
      const linkEnd = styleXml.indexOf('/>', linkStart);
      if (linkEnd !== -1) {
        const linkTag = styleXml.substring(linkStart, linkEnd + 2);
        link = XMLParser.extractAttribute(linkTag, 'w:val') || undefined;
      }
    }

    // aliases - Alternative names
    let aliases: string | undefined;
    if (styleXml.includes('<w:aliases')) {
      const aliasesStart = styleXml.indexOf('<w:aliases');
      const aliasesEnd = styleXml.indexOf('/>', aliasesStart);
      if (aliasesEnd !== -1) {
        const aliasesTag = styleXml.substring(aliasesStart, aliasesEnd + 2);
        aliases = XMLParser.extractAttribute(aliasesTag, 'w:val') || undefined;
      }
    }

    // Parse table style properties (Phase 5.1)
    let tableStyle: import('../formatting/Style').TableStyleProperties | undefined;
    if (typeAttr === 'table') {
      tableStyle = this.parseTableStyleProperties(styleXml);
    }

    // Create style properties
    const properties: StyleProperties = {
      styleId,
      name,
      type: typeAttr,
      basedOn,
      next,
      // w:default and w:customStyle are ST_OnOff per ECMA-376 §17.17.4
      isDefault: parseOnOffAttribute(defaultAttr),
      customStyle: parseOnOffAttribute(customStyleAttr),
      paragraphFormatting,
      numPr: styleNumPr,
      runFormatting,
      tableStyle,
      // Metadata CT_OnOff flags (ECMA-376 §17.7.4). parseStyleOnOffFlag returns
      // undefined when the element is absent, or the actual boolean (true/false)
      // when present — preserve both "explicit false" (override) and "absent"
      // (inherit) faithfully through the Style properties record.
      qFormat,
      semiHidden,
      hidden,
      unhideWhenUsed,
      locked,
      personal,
      personalCompose,
      personalReply,
      rsid: styleRsid,
      autoRedefine,
      uiPriority,
      link,
      aliases,
    };

    return Style.create(properties);
  }

  /**
   * Parses paragraph formatting from XML (w:pPr element content)
   * @param pPrXml - XML content inside w:pPr tags
   * @returns ParagraphFormatting object
   */
  private parseParagraphFormattingFromXml(pPrXml: string): ParagraphFormatting {
    const formatting: ParagraphFormatting = {};

    // Parse framePr (text frame properties) per ECMA-376 Part 1 §17.3.1.11 —
    // CT_FramePr is a CT_PPrBase child (#5, between pageBreakBefore and
    // widowControl). Each attribute is independently optional; numeric
    // attributes (w/h/x/y/hSpace/vSpace/lines) may legitimately be zero
    // so use explicit string presence rather than truthy checks.
    const framePrTag = XMLParser.extractSelfClosingTag(pPrXml, 'w:framePr');
    if (framePrTag) {
      const fpStr = `<w:framePr${framePrTag}`;
      const frameProps: NonNullable<ParagraphFormatting['framePr']> = {};
      const wAttr = XMLParser.extractAttribute(fpStr, 'w:w');
      if (wAttr !== undefined) frameProps.w = parseInt(wAttr, 10);
      const hAttr = XMLParser.extractAttribute(fpStr, 'w:h');
      if (hAttr !== undefined) frameProps.h = parseInt(hAttr, 10);
      const hRule = XMLParser.extractAttribute(fpStr, 'w:hRule');
      if (hRule === 'auto' || hRule === 'atLeast' || hRule === 'exact') {
        frameProps.hRule = hRule;
      }
      const xAttr = XMLParser.extractAttribute(fpStr, 'w:x');
      if (xAttr !== undefined) frameProps.x = parseInt(xAttr, 10);
      const yAttr = XMLParser.extractAttribute(fpStr, 'w:y');
      if (yAttr !== undefined) frameProps.y = parseInt(yAttr, 10);
      const xAlign = XMLParser.extractAttribute(fpStr, 'w:xAlign');
      if (
        xAlign === 'left' ||
        xAlign === 'center' ||
        xAlign === 'right' ||
        xAlign === 'inside' ||
        xAlign === 'outside'
      ) {
        frameProps.xAlign = xAlign;
      }
      const yAlign = XMLParser.extractAttribute(fpStr, 'w:yAlign');
      if (
        yAlign === 'top' ||
        yAlign === 'center' ||
        yAlign === 'bottom' ||
        yAlign === 'inline' ||
        yAlign === 'inside' ||
        yAlign === 'outside'
      ) {
        frameProps.yAlign = yAlign;
      }
      const hAnchor = XMLParser.extractAttribute(fpStr, 'w:hAnchor');
      if (hAnchor === 'page' || hAnchor === 'margin' || hAnchor === 'text') {
        frameProps.hAnchor = hAnchor;
      }
      const vAnchor = XMLParser.extractAttribute(fpStr, 'w:vAnchor');
      if (vAnchor === 'page' || vAnchor === 'margin' || vAnchor === 'text') {
        frameProps.vAnchor = vAnchor;
      }
      const hSpace = XMLParser.extractAttribute(fpStr, 'w:hSpace');
      if (hSpace !== undefined) frameProps.hSpace = parseInt(hSpace, 10);
      const vSpace = XMLParser.extractAttribute(fpStr, 'w:vSpace');
      if (vSpace !== undefined) frameProps.vSpace = parseInt(vSpace, 10);
      const wrap = XMLParser.extractAttribute(fpStr, 'w:wrap');
      if (
        wrap === 'around' ||
        wrap === 'auto' ||
        wrap === 'none' ||
        wrap === 'notBeside' ||
        wrap === 'through' ||
        wrap === 'tight'
      ) {
        frameProps.wrap = wrap;
      }
      const dropCap = XMLParser.extractAttribute(fpStr, 'w:dropCap');
      if (dropCap === 'none' || dropCap === 'drop' || dropCap === 'margin') {
        frameProps.dropCap = dropCap;
      }
      const lines = XMLParser.extractAttribute(fpStr, 'w:lines');
      if (lines !== undefined) frameProps.lines = parseInt(lines, 10);
      const anchorLock = XMLParser.extractAttribute(fpStr, 'w:anchorLock');
      if (anchorLock !== undefined) {
        frameProps.anchorLock = parseOnOffAttribute(anchorLock, true);
      }
      if (Object.keys(frameProps).length > 0) {
        formatting.framePr = frameProps;
      }
    }

    // Parse alignment (w:jc)
    const jcElement = XMLParser.extractSelfClosingTag(pPrXml, 'w:jc');
    if (jcElement) {
      const alignment = XMLParser.extractAttribute(`<w:jc${jcElement}`, 'w:val');
      if (alignment) {
        formatting.alignment = alignment as 'left' | 'center' | 'right' | 'justify';
      }
    }

    // Parse spacing (w:spacing) — all 8 CT_Spacing attributes per ECMA-376 §17.3.1.33
    const spacingElement = XMLParser.extractSelfClosingTag(pPrXml, 'w:spacing');
    if (spacingElement) {
      const spacingTag = `<w:spacing${spacingElement}`;
      const before = XMLParser.extractAttribute(spacingTag, 'w:before');
      const after = XMLParser.extractAttribute(spacingTag, 'w:after');
      const line = XMLParser.extractAttribute(spacingTag, 'w:line');
      const lineRule = XMLParser.extractAttribute(spacingTag, 'w:lineRule');
      const beforeLines = XMLParser.extractAttribute(spacingTag, 'w:beforeLines');
      const afterLines = XMLParser.extractAttribute(spacingTag, 'w:afterLines');
      const beforeAutosp = XMLParser.extractAttribute(spacingTag, 'w:beforeAutospacing');
      const afterAutosp = XMLParser.extractAttribute(spacingTag, 'w:afterAutospacing');

      // Validate lineRule
      let validatedLineRule: 'auto' | 'exact' | 'atLeast' | undefined;
      if (lineRule) {
        const validLineRules = ['auto', 'exact', 'atLeast'];
        if (validLineRules.includes(lineRule)) {
          validatedLineRule = lineRule as 'auto' | 'exact' | 'atLeast';
        }
      }

      formatting.spacing = {
        before: before ? parseInt(before, 10) : undefined,
        after: after ? parseInt(after, 10) : undefined,
        // If lineRule exists without line, use default 240 twips
        line: line ? parseInt(line, 10) : validatedLineRule ? 240 : undefined,
        lineRule: validatedLineRule,
        beforeLines: beforeLines ? parseInt(beforeLines, 10) : undefined,
        afterLines: afterLines ? parseInt(afterLines, 10) : undefined,
        // ST_OnOff per ECMA-376 §17.17.4 — accept 1/0/true/false/on/off
        beforeAutospacing: beforeAutosp ? parseOnOffAttribute(beforeAutosp) : undefined,
        afterAutospacing: afterAutosp ? parseOnOffAttribute(afterAutosp) : undefined,
      };
    }

    // Parse indentation (w:ind)
    // Per ECMA-376 §17.3.1.15: w:start/w:end are bidi-aware alternatives to
    // w:left/w:right. §17.3.1.12 also defines six CJK character-unit variants
    // (ST_DecimalNumber) — parse those alongside so styles authored in CJK
    // locales preserve their character-unit indent spec through round-trip.
    const indElement = XMLParser.extractSelfClosingTag(pPrXml, 'w:ind');
    if (indElement) {
      const indTag = `<w:ind${indElement}`;
      const start = XMLParser.extractAttribute(indTag, 'w:start');
      const left = XMLParser.extractAttribute(indTag, 'w:left');
      const end = XMLParser.extractAttribute(indTag, 'w:end');
      const right = XMLParser.extractAttribute(indTag, 'w:right');
      const firstLine = XMLParser.extractAttribute(indTag, 'w:firstLine');
      const hanging = XMLParser.extractAttribute(indTag, 'w:hanging');
      // CJK character-unit variants. startChars/endChars collapse to
      // leftChars/rightChars (same bidi-aware rule as the twips pair).
      const startChars = XMLParser.extractAttribute(indTag, 'w:startChars');
      const leftChars = XMLParser.extractAttribute(indTag, 'w:leftChars');
      const endChars = XMLParser.extractAttribute(indTag, 'w:endChars');
      const rightChars = XMLParser.extractAttribute(indTag, 'w:rightChars');
      const firstLineChars = XMLParser.extractAttribute(indTag, 'w:firstLineChars');
      const hangingChars = XMLParser.extractAttribute(indTag, 'w:hangingChars');

      const leftVal = start || left;
      const rightVal = end || right;
      const leftCharsVal = startChars || leftChars;
      const rightCharsVal = endChars || rightChars;

      formatting.indentation = {
        left: leftVal ? parseInt(leftVal, 10) : undefined,
        right: rightVal ? parseInt(rightVal, 10) : undefined,
        firstLine: firstLine ? parseInt(firstLine, 10) : undefined,
        hanging: hanging ? parseInt(hanging, 10) : undefined,
        leftChars: leftCharsVal ? parseInt(leftCharsVal, 10) : undefined,
        rightChars: rightCharsVal ? parseInt(rightCharsVal, 10) : undefined,
        firstLineChars: firstLineChars ? parseInt(firstLineChars, 10) : undefined,
        hangingChars: hangingChars ? parseInt(hangingChars, 10) : undefined,
      };
    }

    // Parse CT_OnOff boolean flags per ECMA-376 §17.17.4 / §17.3.1. The previous
    // substring-only detection (`pPrXml.includes('<w:keepNext/>') ||
    // pPrXml.includes('<w:keepNext ')`) hard-coded the flag to true whenever
    // the element appeared at all — silently flipping `<w:keepNext w:val="0"/>`
    // (explicit override) into an enabled flag. Read w:val when present and
    // honour every ST_OnOff literal (1/0/true/false/on/off).
    const parseStylePPrCtOnOff = (tagName: string): boolean | undefined => {
      // extractSelfClosingTag returns the ATTRIBUTE STRING (possibly empty)
      // when found, or `undefined` when absent. Earlier this helper checked
      // `=== null` by mistake — that let the "absent" case fall through and
      // construct a garbage tag that produced `true`, silently enabling the
      // flag on every style that didn't set it.
      const el = XMLParser.extractSelfClosingTag(pPrXml, tagName);
      if (el === undefined) return undefined;
      const val = XMLParser.extractAttribute(`<${tagName}${el}`, 'w:val');
      if (val === undefined) return true;
      return parseOnOffAttribute(val, true);
    };

    const keepNextVal = parseStylePPrCtOnOff('w:keepNext');
    if (keepNextVal !== undefined) formatting.keepNext = keepNextVal;

    const keepLinesVal = parseStylePPrCtOnOff('w:keepLines');
    if (keepLinesVal !== undefined) formatting.keepLines = keepLinesVal;

    const pageBreakBeforeVal = parseStylePPrCtOnOff('w:pageBreakBefore');
    if (pageBreakBeforeVal !== undefined) formatting.pageBreakBefore = pageBreakBeforeVal;

    // Contextual spacing per ECMA-376 Part 1 §17.3.1.9
    // "Don't add space between paragraphs of the same style"
    const contextualSpacingVal = parseStylePPrCtOnOff('w:contextualSpacing');
    if (contextualSpacingVal !== undefined) formatting.contextualSpacing = contextualSpacingVal;

    // Remaining CT_PPrBase CT_OnOff flags per ECMA-376 Part 1 §17.3.1.
    // The main paragraph parser handles all of these; the style-level parser
    // previously dropped them (substring matches existed only for the four
    // flags above). Any style using the explicit-false form to override a
    // based-on style's enabled flag was silently losing the override.
    const widowControlVal = parseStylePPrCtOnOff('w:widowControl');
    if (widowControlVal !== undefined) formatting.widowControl = widowControlVal;

    const suppressLineNumbersVal = parseStylePPrCtOnOff('w:suppressLineNumbers');
    if (suppressLineNumbersVal !== undefined)
      formatting.suppressLineNumbers = suppressLineNumbersVal;

    const bidiVal = parseStylePPrCtOnOff('w:bidi');
    if (bidiVal !== undefined) formatting.bidi = bidiVal;

    const mirrorIndentsVal = parseStylePPrCtOnOff('w:mirrorIndents');
    if (mirrorIndentsVal !== undefined) formatting.mirrorIndents = mirrorIndentsVal;

    const adjustRightIndVal = parseStylePPrCtOnOff('w:adjustRightInd');
    if (adjustRightIndVal !== undefined) formatting.adjustRightInd = adjustRightIndVal;

    const suppressAutoHyphensVal = parseStylePPrCtOnOff('w:suppressAutoHyphens');
    if (suppressAutoHyphensVal !== undefined)
      formatting.suppressAutoHyphens = suppressAutoHyphensVal;

    const kinsokuVal = parseStylePPrCtOnOff('w:kinsoku');
    if (kinsokuVal !== undefined) formatting.kinsoku = kinsokuVal;

    const wordWrapVal = parseStylePPrCtOnOff('w:wordWrap');
    if (wordWrapVal !== undefined) formatting.wordWrap = wordWrapVal;

    const overflowPunctVal = parseStylePPrCtOnOff('w:overflowPunct');
    if (overflowPunctVal !== undefined) formatting.overflowPunct = overflowPunctVal;

    const topLinePunctVal = parseStylePPrCtOnOff('w:topLinePunct');
    if (topLinePunctVal !== undefined) formatting.topLinePunct = topLinePunctVal;

    const autoSpaceDEVal = parseStylePPrCtOnOff('w:autoSpaceDE');
    if (autoSpaceDEVal !== undefined) formatting.autoSpaceDE = autoSpaceDEVal;

    const autoSpaceDNVal = parseStylePPrCtOnOff('w:autoSpaceDN');
    if (autoSpaceDNVal !== undefined) formatting.autoSpaceDN = autoSpaceDNVal;

    const suppressOverlapVal = parseStylePPrCtOnOff('w:suppressOverlap');
    if (suppressOverlapVal !== undefined) formatting.suppressOverlap = suppressOverlapVal;

    // Parse `w:val`-attribute string-enum children per CT_PPrBase.
    // Position #28 textDirection (ST_TextDirection), #29 textAlignment
    // (ST_TextAlignment), #30 textboxTightWrap (ST_TextboxTightWrapType).
    // The main paragraph parser handles these; the style-level parser
    // previously dropped them because the substring scan was never
    // extended past the iteration-25 CT_OnOff helper.
    const parseStylePPrValAttr = (tagName: string): string | undefined => {
      const el = XMLParser.extractSelfClosingTag(pPrXml, tagName);
      if (el === undefined) return undefined;
      const val = XMLParser.extractAttribute(`<${tagName}${el}`, 'w:val');
      return val === undefined ? undefined : String(val);
    };

    const textDirectionVal = parseStylePPrValAttr('w:textDirection');
    if (textDirectionVal !== undefined) {
      formatting.textDirection = textDirectionVal as ParagraphFormatting['textDirection'];
    }

    const textAlignmentVal = parseStylePPrValAttr('w:textAlignment');
    if (textAlignmentVal !== undefined) {
      formatting.textAlignment = textAlignmentVal as ParagraphFormatting['textAlignment'];
    }

    const textboxTightWrapVal = parseStylePPrValAttr('w:textboxTightWrap');
    if (textboxTightWrapVal !== undefined) {
      formatting.textboxTightWrap = textboxTightWrapVal as ParagraphFormatting['textboxTightWrap'];
    }

    // Parse outline level (w:outlineLvl) - used for TOC generation
    // Per ECMA-376 Part 1 §17.3.1.20: val is 0-8 (heading levels 1-9)
    const outlineLvlElement = XMLParser.extractSelfClosingTag(pPrXml, 'w:outlineLvl');
    if (outlineLvlElement) {
      const outlineVal = XMLParser.extractAttribute(`<w:outlineLvl${outlineLvlElement}`, 'w:val');
      if (outlineVal) {
        const level = parseInt(outlineVal, 10);
        if (!isNaN(level) && level >= 0 && level <= 8) {
          formatting.outlineLevel = level;
        }
      }
    }

    // Parse divId (CT_PPrBase #32, §17.3.1.10) — numeric HTML div
    // association. Previously dropped on the style parser; the main
    // paragraph parser reads it at the pPrObj level but the style pPr
    // parser used string-based extraction and skipped both divId and
    // cnfStyle below.
    const divIdVal = parseStylePPrValAttr('w:divId');
    if (divIdVal !== undefined) {
      const parsedDivId = parseInt(divIdVal, 10);
      if (!isNaN(parsedDivId)) formatting.divId = parsedDivId;
    }

    // Parse cnfStyle (CT_PPrBase #33, §17.3.1.8) — conditional formatting
    // bitmask string (12-char 0/1 sequence, e.g. "100000000100").
    const cnfStyleVal = parseStylePPrValAttr('w:cnfStyle');
    if (cnfStyleVal !== undefined) formatting.cnfStyle = cnfStyleVal;

    // Parse paragraph borders (w:pBdr) per ECMA-376 Part 1 §17.3.1.24.
    // Covers the full CT_Border attribute set (§17.18.2): val, sz, space,
    // color, themeColor, themeTint, themeShade, shadow, frame. The style
    // *emitter* already round-trips all nine, so any style-pBdr authored
    // by Word with themed or shadow/frame attributes was silently flattened
    // here before this fix. Shadow/frame route through parseOnOffAttribute
    // so ST_OnOff literals ("on"/"off"/"1"/"0"/"true"/"false") resolve.
    const pBdrXml = XMLParser.extractBetweenTags(pPrXml, '<w:pBdr>', '</w:pBdr>');
    if (pBdrXml) {
      const borders: any = {};
      const borderTypes = ['top', 'left', 'bottom', 'right', 'between', 'bar'];
      for (const type of borderTypes) {
        if (pBdrXml.includes(`<w:${type}`)) {
          const tag = XMLParser.extractSelfClosingTag(pBdrXml, `w:${type}`);
          if (tag) {
            const bTag = `<w:${type}${tag}`;
            const style = XMLParser.extractAttribute(bTag, 'w:val');
            const size = XMLParser.extractAttribute(bTag, 'w:sz');
            const space = XMLParser.extractAttribute(bTag, 'w:space');
            const color = XMLParser.extractAttribute(bTag, 'w:color');
            const themeColor = XMLParser.extractAttribute(bTag, 'w:themeColor');
            const themeTint = XMLParser.extractAttribute(bTag, 'w:themeTint');
            const themeShade = XMLParser.extractAttribute(bTag, 'w:themeShade');
            const shadow = XMLParser.extractAttribute(bTag, 'w:shadow');
            const frame = XMLParser.extractAttribute(bTag, 'w:frame');
            const border: any = {};
            if (style) border.style = style;
            if (size) border.size = parseInt(size, 10);
            if (space) border.space = parseInt(space, 10);
            if (color) border.color = color;
            if (themeColor) border.themeColor = themeColor;
            if (themeTint) border.themeTint = themeTint;
            if (themeShade) border.themeShade = themeShade;
            if (shadow !== undefined) border.shadow = parseOnOffAttribute(shadow, true);
            if (frame !== undefined) border.frame = parseOnOffAttribute(frame, true);
            if (Object.keys(border).length > 0) borders[type] = border;
          }
        }
      }
      if (Object.keys(borders).length > 0) formatting.borders = borders;
    }

    // Parse tab stops (w:tabs) per ECMA-376 Part 1 §17.3.1.38
    const tabsXml = XMLParser.extractBetweenTags(pPrXml, '<w:tabs>', '</w:tabs>');
    if (tabsXml) {
      const tabs: any[] = [];
      // Extract all w:tab elements
      const tabRegex = /<w:tab\s[^>]*\/>/g;
      let tabMatch;
      while ((tabMatch = tabRegex.exec(tabsXml)) !== null) {
        const tabTag = tabMatch[0];
        const pos = XMLParser.extractAttribute(tabTag, 'w:pos');
        const val = XMLParser.extractAttribute(tabTag, 'w:val');
        const leader = XMLParser.extractAttribute(tabTag, 'w:leader');
        if (pos) {
          const tab: any = { position: parseInt(pos, 10) };
          if (val) tab.val = val;
          if (leader) tab.leader = leader;
          tabs.push(tab);
        }
      }
      if (tabs.length > 0) formatting.tabs = tabs;
    }

    // Parse shading (w:shd) per ECMA-376 Part 1 §17.3.1.32
    const shading = this.parseShadingFromXml(pPrXml);
    if (shading) {
      formatting.shading = shading;
    }

    return formatting;
  }

  /**
   * Parses run formatting from XML (w:rPr element content)
   * @param rPrXml - XML content inside w:rPr tags
   * @returns RunFormatting object
   */
  private parseRunFormattingFromXml(rPrXml: string): RunFormatting {
    const formatting: RunFormatting = {};

    // CT_OnOff rPr children per ECMA-376 §17.3.2. Previously detected via
    // substring-include which hard-coded the flag to `true` whenever the
    // element appeared — silently flipping `<w:b w:val="0"/>` (explicit
    // override of a based-on style's bold) into an enabled flag, and
    // never setting the field to `false` for legitimate overrides.
    // Mirrors the pPr `parseStylePPrCtOnOff` helper introduced in
    // iteration 25 / 26.
    const parseStyleRPrCtOnOff = (tagName: string): boolean | undefined => {
      const el = XMLParser.extractSelfClosingTag(rPrXml, tagName);
      if (el === undefined) return undefined;
      const val = XMLParser.extractAttribute(`<${tagName}${el}`, 'w:val');
      if (val === undefined) return true;
      return parseOnOffAttribute(val, true);
    };

    const boldVal = parseStyleRPrCtOnOff('w:b');
    if (boldVal !== undefined) formatting.bold = boldVal;

    const italicVal = parseStyleRPrCtOnOff('w:i');
    if (italicVal !== undefined) formatting.italic = italicVal;

    const strikeVal = parseStyleRPrCtOnOff('w:strike');
    if (strikeVal !== undefined) formatting.strike = strikeVal;

    const smallCapsVal = parseStyleRPrCtOnOff('w:smallCaps');
    if (smallCapsVal !== undefined) formatting.smallCaps = smallCapsVal;

    const allCapsVal = parseStyleRPrCtOnOff('w:caps');
    if (allCapsVal !== undefined) formatting.allCaps = allCapsVal;

    // Extended CT_OnOff run children per ECMA-376 §17.3.2. The style-level
    // rPr parser previously dropped all of these silently, so character
    // styles setting dstrike, outline, shadow, emboss, imprint, rtl,
    // vanish, noProof, snapToGrid, specVanish, webHidden, or complex-script
    // variants (bCs / iCs / cs) lost their overrides on programmatic save.
    const boldCsVal = parseStyleRPrCtOnOff('w:bCs');
    if (boldCsVal !== undefined) formatting.complexScriptBold = boldCsVal;

    const italicCsVal = parseStyleRPrCtOnOff('w:iCs');
    if (italicCsVal !== undefined) formatting.complexScriptItalic = italicCsVal;

    const csVal = parseStyleRPrCtOnOff('w:cs');
    if (csVal !== undefined) formatting.complexScript = csVal;

    const dstrikeVal = parseStyleRPrCtOnOff('w:dstrike');
    if (dstrikeVal !== undefined) formatting.dstrike = dstrikeVal;

    const outlineVal = parseStyleRPrCtOnOff('w:outline');
    if (outlineVal !== undefined) formatting.outline = outlineVal;

    const shadowVal = parseStyleRPrCtOnOff('w:shadow');
    if (shadowVal !== undefined) formatting.shadow = shadowVal;

    const embossVal = parseStyleRPrCtOnOff('w:emboss');
    if (embossVal !== undefined) formatting.emboss = embossVal;

    const imprintVal = parseStyleRPrCtOnOff('w:imprint');
    if (imprintVal !== undefined) formatting.imprint = imprintVal;

    const rtlVal = parseStyleRPrCtOnOff('w:rtl');
    if (rtlVal !== undefined) formatting.rtl = rtlVal;

    const vanishVal = parseStyleRPrCtOnOff('w:vanish');
    if (vanishVal !== undefined) formatting.vanish = vanishVal;

    const noProofVal = parseStyleRPrCtOnOff('w:noProof');
    if (noProofVal !== undefined) formatting.noProof = noProofVal;

    const snapToGridVal = parseStyleRPrCtOnOff('w:snapToGrid');
    if (snapToGridVal !== undefined) formatting.snapToGrid = snapToGridVal;

    const specVanishVal = parseStyleRPrCtOnOff('w:specVanish');
    if (specVanishVal !== undefined) formatting.specVanish = specVanishVal;

    const webHiddenVal = parseStyleRPrCtOnOff('w:webHidden');
    if (webHiddenVal !== undefined) formatting.webHidden = webHiddenVal;

    // Parse underline — all attributes per ECMA-376 §17.3.2.40.
    // Whitelist covers the full ST_Underline enumeration (18 values);
    // unknown / out-of-spec values fall through to `underline = true`
    // (underline enabled with default style) to match the main parser.
    const uElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:u');
    if (uElement) {
      const uTag = `<w:u${uElement}`;
      const uVal = XMLParser.extractAttribute(uTag, 'w:val');
      const ST_UNDERLINE = new Set<string>([
        'single',
        'words',
        'double',
        'thick',
        'dotted',
        'dottedHeavy',
        'dash',
        'dashedHeavy',
        'dashLong',
        'dashLongHeavy',
        'dotDash',
        'dashDotHeavy',
        'dotDotDash',
        'dashDotDotHeavy',
        'wave',
        'wavyHeavy',
        'wavyDouble',
        'none',
      ]);
      if (uVal !== undefined && ST_UNDERLINE.has(String(uVal))) {
        formatting.underline = String(uVal) as RunFormatting['underline'];
      } else {
        formatting.underline = true;
      }
      const uColor = XMLParser.extractAttribute(uTag, 'w:color');
      if (uColor) formatting.underlineColor = uColor;
      const uThemeColor = XMLParser.extractAttribute(uTag, 'w:themeColor');
      if (uThemeColor) {
        formatting.underlineThemeColor = uThemeColor as import('../elements/Run').ThemeColorValue;
      }
      const uThemeTint = XMLParser.extractAttribute(uTag, 'w:themeTint');
      if (uThemeTint) formatting.underlineThemeTint = parseInt(uThemeTint, 16);
      const uThemeShade = XMLParser.extractAttribute(uTag, 'w:themeShade');
      if (uThemeShade) formatting.underlineThemeShade = parseInt(uThemeShade, 16);
    }

    // Parse subscript/superscript/baseline per ECMA-376 §17.18.96 (ST_VerticalAlignRun)
    const vertAlignElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:vertAlign');
    if (vertAlignElement) {
      const val = XMLParser.extractAttribute(`<w:vertAlign${vertAlignElement}`, 'w:val');
      if (val === 'subscript') {
        formatting.subscript = true;
      } else if (val === 'superscript') {
        formatting.superscript = true;
      } else if (val === 'baseline') {
        formatting.vertAlignBaseline = true;
      }
    }

    // Parse font (w:rFonts) — all attributes per ECMA-376 §17.3.2.26
    const rFontsElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:rFonts');
    if (rFontsElement) {
      const rFontsTag = `<w:rFonts${rFontsElement}`;
      const ascii = XMLParser.extractAttribute(rFontsTag, 'w:ascii');
      if (ascii) formatting.font = ascii;
      const hAnsi = XMLParser.extractAttribute(rFontsTag, 'w:hAnsi');
      if (hAnsi) formatting.fontHAnsi = hAnsi;
      const eastAsia = XMLParser.extractAttribute(rFontsTag, 'w:eastAsia');
      if (eastAsia) formatting.fontEastAsia = eastAsia;
      const cs = XMLParser.extractAttribute(rFontsTag, 'w:cs');
      if (cs) formatting.fontCs = cs;
      const hint = XMLParser.extractAttribute(rFontsTag, 'w:hint');
      if (hint) formatting.fontHint = hint;
      const asciiTheme = XMLParser.extractAttribute(rFontsTag, 'w:asciiTheme');
      if (asciiTheme) formatting.fontAsciiTheme = asciiTheme;
      const hAnsiTheme = XMLParser.extractAttribute(rFontsTag, 'w:hAnsiTheme');
      if (hAnsiTheme) formatting.fontHAnsiTheme = hAnsiTheme;
      const eastAsiaTheme = XMLParser.extractAttribute(rFontsTag, 'w:eastAsiaTheme');
      if (eastAsiaTheme) formatting.fontEastAsiaTheme = eastAsiaTheme;
      const cstheme = XMLParser.extractAttribute(rFontsTag, 'w:cstheme');
      if (cstheme) formatting.fontCsTheme = cstheme;
    }

    // Parse size (w:sz) - size is in half-points
    // Use extractSelfClosingTag to avoid matching w:szCs
    const szElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:sz');
    if (szElement) {
      const val = XMLParser.extractAttribute(`<w:sz${szElement}`, 'w:val');
      if (val) {
        formatting.size = halfPointsToPoints(parseInt(val, 10));
      }
    }

    // Parse complex script size (w:szCs) per ECMA-376 §17.3.2.40
    const szCsElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:szCs');
    if (szCsElement) {
      const val = XMLParser.extractAttribute(`<w:szCs${szCsElement}`, 'w:val');
      if (val) {
        const szCsVal = halfPointsToPoints(parseInt(val, 10));
        if (formatting.size === undefined || szCsVal !== formatting.size) {
          formatting.sizeCs = szCsVal;
        }
      }
    }

    // Parse color (w:color) — all attributes per ECMA-376 §17.3.2.6 / ST_HexColor
    // per §17.18.38. `w:val="auto"` is a valid ST_HexColorAuto sentinel that
    // tells Word to use the automatic/window text color; the previous parser
    // dropped it (only storing non-auto hex values), so a style-level rPr with
    // `<w:color w:val="auto"/>` silently lost that marker on round-trip and
    // the emitter defaulted to `"000000"` — changing the rendering of any
    // style that relied on the auto fallback. Preserve the literal "auto" so
    // it survives through emission. (Matches the object-format parser path
    // for direct-run rPr at parseRunFromObject line ~5210.)
    const colorElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:color');
    if (colorElement) {
      const colorTag = `<w:color${colorElement}`;
      const val = XMLParser.extractAttribute(colorTag, 'w:val');
      if (val) {
        formatting.color = val;
      }
      const themeColor = XMLParser.extractAttribute(colorTag, 'w:themeColor');
      if (themeColor) {
        formatting.themeColor = themeColor as import('../elements/Run').ThemeColorValue;
      }
      const themeTint = XMLParser.extractAttribute(colorTag, 'w:themeTint');
      if (themeTint) {
        formatting.themeTint = parseInt(themeTint, 16);
      }
      const themeShade = XMLParser.extractAttribute(colorTag, 'w:themeShade');
      if (themeShade) {
        formatting.themeShade = parseInt(themeShade, 16);
      }
    }

    // Parse highlight (w:highlight) - use extractSelfClosingTag
    const highlightElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:highlight');
    if (highlightElement) {
      const val = XMLParser.extractAttribute(`<w:highlight${highlightElement}`, 'w:val');
      if (val) {
        const validHighlights = [
          'yellow',
          'green',
          'cyan',
          'magenta',
          'blue',
          'red',
          'darkBlue',
          'darkCyan',
          'darkGreen',
          'darkMagenta',
          'darkRed',
          'darkYellow',
          'darkGray',
          'lightGray',
          'black',
          'white',
          'none',
        ];
        if (validHighlights.includes(val)) {
          formatting.highlight = val as
            | 'yellow'
            | 'green'
            | 'cyan'
            | 'magenta'
            | 'blue'
            | 'red'
            | 'darkBlue'
            | 'darkCyan'
            | 'darkGreen'
            | 'darkMagenta'
            | 'darkRed'
            | 'darkYellow'
            | 'darkGray'
            | 'lightGray'
            | 'black'
            | 'white'
            | 'none';
        }
      }
    }

    // Parse shading (w:shd) per ECMA-376 Part 1 §17.3.2.32
    const shading = this.parseShadingFromXml(rPrXml);
    if (shading) {
      formatting.shading = shading;
    }

    // Character spacing (w:spacing §17.3.2.35, ST_SignedTwipsMeasure) —
    // previously dropped on the style parser; 0 and negative values are
    // valid per spec.
    const spacingEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:spacing');
    if (spacingEl !== undefined) {
      const val = XMLParser.extractAttribute(`<w:spacing${spacingEl}`, 'w:val');
      if (val !== undefined) {
        const n = parseInt(String(val), 10);
        if (!isNaN(n)) formatting.characterSpacing = n;
      }
    }

    // Vertical position (w:position §17.3.2.31, ST_SignedHpsMeasure).
    // 0 = baseline (explicit reset); negative = lowered; positive = raised.
    const positionEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:position');
    if (positionEl !== undefined) {
      const val = XMLParser.extractAttribute(`<w:position${positionEl}`, 'w:val');
      if (val !== undefined) {
        const n = parseInt(String(val), 10);
        if (!isNaN(n)) formatting.position = n;
      }
    }

    // Kerning threshold (w:kern §17.3.2.20, ST_HpsMeasure). 0 = kern at
    // every size (no minimum font-size threshold).
    const kernEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:kern');
    if (kernEl !== undefined) {
      const val = XMLParser.extractAttribute(`<w:kern${kernEl}`, 'w:val');
      if (val !== undefined) {
        const n = parseInt(String(val), 10);
        if (!isNaN(n)) formatting.kerning = n;
      }
    }

    // Language (w:lang §17.3.2.20, CT_Language). Single val → plain string;
    // multi-script (eastAsia and/or bidi present) → LanguageConfig object.
    const langEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:lang');
    if (langEl !== undefined) {
      const langTag = `<w:lang${langEl}`;
      const val = XMLParser.extractAttribute(langTag, 'w:val');
      const eastAsia = XMLParser.extractAttribute(langTag, 'w:eastAsia');
      const bidi = XMLParser.extractAttribute(langTag, 'w:bidi');
      if (eastAsia || bidi) {
        formatting.language = {
          val: val ? String(val) : undefined,
          eastAsia: eastAsia ? String(eastAsia) : undefined,
          bidi: bidi ? String(bidi) : undefined,
        };
      } else if (val) {
        formatting.language = String(val);
      }
    }

    // Horizontal scaling (w:w §17.3.2.43, ST_TextScale — percentage,
    // min 1 per spec, so 0 is not valid and we keep a truthy check).
    const scaleEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:w');
    if (scaleEl !== undefined) {
      const val = XMLParser.extractAttribute(`<w:w${scaleEl}`, 'w:val');
      if (val) {
        const n = parseInt(String(val), 10);
        if (!isNaN(n)) formatting.scaling = n;
      }
    }

    // Emphasis mark (w:em §17.3.2.13, ST_Em — "dot"/"comma"/"circle"/
    // "underDot"/"none"). Commonly paired with East Asian typography.
    const emEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:em');
    if (emEl !== undefined) {
      const val = XMLParser.extractAttribute(`<w:em${emEl}`, 'w:val');
      if (val) {
        formatting.emphasis = String(val) as RunFormatting['emphasis'];
      }
    }

    // Animated text effect (w:effect §17.3.2.12, ST_TextEffect —
    // "blinkBackground"/"lights"/"antsBlack"/"antsRed"/"shimmer"/"sparkle"/
    // "none"). Legacy feature but still valid per schema.
    const effectEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:effect');
    if (effectEl !== undefined) {
      const val = XMLParser.extractAttribute(`<w:effect${effectEl}`, 'w:val');
      if (val) {
        formatting.effect = String(val) as RunFormatting['effect'];
      }
    }

    // Text border (w:bdr §17.3.2.5) — character/run border. Full CT_Border
    // attribute set (§17.18.2): val / sz / space / color / themeColor /
    // themeTint / themeShade / shadow / frame.
    const bdrEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:bdr');
    if (bdrEl !== undefined) {
      const bdrTag = `<w:bdr${bdrEl}`;
      const border: {
        style?: string;
        size?: number;
        color?: string;
        space?: number;
        themeColor?: string;
        themeTint?: string;
        themeShade?: string;
        shadow?: boolean;
        frame?: boolean;
      } = {};
      const val = XMLParser.extractAttribute(bdrTag, 'w:val');
      if (val) border.style = String(val);
      const sz = XMLParser.extractAttribute(bdrTag, 'w:sz');
      if (sz !== undefined) {
        const n = parseInt(String(sz), 10);
        if (!isNaN(n)) border.size = n;
      }
      const color = XMLParser.extractAttribute(bdrTag, 'w:color');
      if (color) border.color = String(color);
      const space = XMLParser.extractAttribute(bdrTag, 'w:space');
      if (space !== undefined) {
        const n = parseInt(String(space), 10);
        if (!isNaN(n)) border.space = n;
      }
      const themeColor = XMLParser.extractAttribute(bdrTag, 'w:themeColor');
      if (themeColor) border.themeColor = String(themeColor);
      const themeTint = XMLParser.extractAttribute(bdrTag, 'w:themeTint');
      if (themeTint) border.themeTint = String(themeTint);
      const themeShade = XMLParser.extractAttribute(bdrTag, 'w:themeShade');
      if (themeShade) border.themeShade = String(themeShade);
      const shadow = XMLParser.extractAttribute(bdrTag, 'w:shadow');
      if (shadow !== undefined) border.shadow = parseOnOffAttribute(shadow, true);
      const frame = XMLParser.extractAttribute(bdrTag, 'w:frame');
      if (frame !== undefined) border.frame = parseOnOffAttribute(frame, true);
      if (Object.keys(border).length > 0) {
        formatting.border = border as RunFormatting['border'];
      }
    }

    // Manual run width (w:fitText §17.3.2.15). Value is twips; 0 is
    // technically representable as "explicit zero" — use `!== undefined`.
    const fitTextEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:fitText');
    if (fitTextEl !== undefined) {
      const val = XMLParser.extractAttribute(`<w:fitText${fitTextEl}`, 'w:val');
      if (val !== undefined) {
        const n = parseInt(String(val), 10);
        if (!isNaN(n)) formatting.fitText = n;
      }
    }

    // East Asian layout (w:eastAsianLayout §17.3.2.10) — combined
    // characters / vertical text / compression attributes.
    const ealEl = XMLParser.extractSelfClosingTag(rPrXml, 'w:eastAsianLayout');
    if (ealEl !== undefined) {
      const ealTag = `<w:eastAsianLayout${ealEl}`;
      const layout: Partial<{
        id: number;
        vert: boolean;
        vertCompress: boolean;
        combine: boolean;
        combineBrackets: 'none' | 'round' | 'square' | 'angle' | 'curly';
      }> = {};
      const id = XMLParser.extractAttribute(ealTag, 'w:id');
      if (id !== undefined) {
        const n = Number(id);
        if (!isNaN(n)) layout.id = n;
      }
      const vert = XMLParser.extractAttribute(ealTag, 'w:vert');
      if (vert !== undefined && parseOnOffAttribute(vert, true)) layout.vert = true;
      const vertCompress = XMLParser.extractAttribute(ealTag, 'w:vertCompress');
      if (vertCompress !== undefined && parseOnOffAttribute(vertCompress, true))
        layout.vertCompress = true;
      const combine = XMLParser.extractAttribute(ealTag, 'w:combine');
      if (combine !== undefined && parseOnOffAttribute(combine, true)) layout.combine = true;
      const combineBrackets = XMLParser.extractAttribute(ealTag, 'w:combineBrackets');
      if (combineBrackets) {
        layout.combineBrackets = String(combineBrackets) as
          | 'none'
          | 'round'
          | 'square'
          | 'angle'
          | 'curly';
      }
      if (Object.keys(layout).length > 0) {
        formatting.eastAsianLayout = layout as RunFormatting['eastAsianLayout'];
      }
    }

    return formatting;
  }

  /**
   * Parses table style properties from style XML (Phase 5.1)
   * @param styleXml - XML string of a table style element
   * @returns TableStyleProperties object
   */
  private parseTableStyleProperties(
    styleXml: string
  ): import('../formatting/Style').TableStyleProperties {
    const tableStyle: import('../formatting/Style').TableStyleProperties = {};

    // Parse tblPr (table properties)
    const tblPrXml = XMLParser.extractBetweenTags(styleXml, '<w:tblPr>', '</w:tblPr>');
    if (tblPrXml) {
      tableStyle.table = this.parseTableFormattingFromXml(tblPrXml);

      // Row band size
      if (tblPrXml.includes('<w:tblStyleRowBandSize')) {
        const tag = XMLParser.extractSelfClosingTag(tblPrXml, 'w:tblStyleRowBandSize');
        if (tag) {
          const val = XMLParser.extractAttribute(`<w:tblStyleRowBandSize${tag}`, 'w:val');
          if (val) {
            tableStyle.rowBandSize = parseInt(val, 10);
          }
        }
      }

      // Column band size
      if (tblPrXml.includes('<w:tblStyleColBandSize')) {
        const tag = XMLParser.extractSelfClosingTag(tblPrXml, 'w:tblStyleColBandSize');
        if (tag) {
          const val = XMLParser.extractAttribute(`<w:tblStyleColBandSize${tag}`, 'w:val');
          if (val) {
            tableStyle.colBandSize = parseInt(val, 10);
          }
        }
      }
    }

    // Parse tcPr (cell properties)
    const tcPrXml = XMLParser.extractBetweenTags(styleXml, '<w:tcPr>', '</w:tcPr>');
    if (tcPrXml) {
      tableStyle.cell = this.parseTableCellFormattingFromXml(tcPrXml);
    }

    // Parse trPr (row properties)
    const trPrXml = XMLParser.extractBetweenTags(styleXml, '<w:trPr>', '</w:trPr>');
    if (trPrXml) {
      tableStyle.row = this.parseTableRowFormattingFromXml(trPrXml);
    }

    // Parse tblStylePr (conditional formatting)
    tableStyle.conditionalFormatting = this.parseConditionalFormattingFromXml(styleXml);

    return tableStyle;
  }

  /**
   * Parses table formatting from tblPr XML (Phase 5.1)
   */
  private parseTableFormattingFromXml(
    tblPrXml: string
  ): import('../formatting/Style').TableStyleFormatting {
    const formatting: import('../formatting/Style').TableStyleFormatting = {};

    // Parse indent (w:tblInd) — preserve w:type per ECMA-376 ST_TblWidth
    if (tblPrXml.includes('<w:tblInd')) {
      const tag = XMLParser.extractSelfClosingTag(tblPrXml, 'w:tblInd');
      if (tag) {
        const tblIndTag = `<w:tblInd${tag}`;
        const w = XMLParser.extractAttribute(tblIndTag, 'w:w');
        if (w) {
          formatting.indent = parseInt(w, 10);
        }
        const type = XMLParser.extractAttribute(tblIndTag, 'w:type');
        if (type) {
          formatting.indentType = type as import('../elements/Table').TableWidthType;
        }
      }
    }

    // Parse alignment — ST_JcTable has 5 values (start, end, center, left,
    // right) per ECMA-376 §17.18.45. The whitelist previously only accepted
    // the three legacy LTR-centric values, silently dropping `start` / `end`
    // (the bidi-aware defaults a modern authoring tool emits).
    if (tblPrXml.includes('<w:jc')) {
      const tag = XMLParser.extractSelfClosingTag(tblPrXml, 'w:jc');
      if (tag) {
        const val = XMLParser.extractAttribute(`<w:jc${tag}`, 'w:val');
        if (
          val === 'left' ||
          val === 'center' ||
          val === 'right' ||
          val === 'start' ||
          val === 'end'
        ) {
          formatting.alignment = val as import('../formatting/Style').TableAlignment;
        }
      }
    }

    // Parse cell spacing
    if (tblPrXml.includes('<w:tblCellSpacing')) {
      const tag = XMLParser.extractSelfClosingTag(tblPrXml, 'w:tblCellSpacing');
      if (tag) {
        const w = XMLParser.extractAttribute(`<w:tblCellSpacing${tag}`, 'w:w');
        if (w) {
          formatting.cellSpacing = parseInt(w, 10);
        }
      }
    }

    // Parse borders
    const bordersXml = XMLParser.extractBetweenTags(tblPrXml, '<w:tblBorders>', '</w:tblBorders>');
    if (bordersXml) {
      formatting.borders = this.parseBordersFromXml(bordersXml, false);
    }

    // Parse shading
    if (tblPrXml.includes('<w:shd')) {
      formatting.shading = this.parseShadingFromXml(tblPrXml);
    }

    // Parse cell margins
    const marginXml = XMLParser.extractBetweenTags(tblPrXml, '<w:tblCellMar>', '</w:tblCellMar>');
    if (marginXml) {
      formatting.cellMargins = this.parseCellMarginsFromXml(marginXml);
    }

    return formatting;
  }

  /**
   * Parses table cell formatting from tcPr XML (Phase 5.1)
   */
  private parseTableCellFormattingFromXml(
    tcPrXml: string
  ): import('../formatting/Style').TableCellStyleFormatting {
    const formatting: import('../formatting/Style').TableCellStyleFormatting = {};

    // Parse borders
    const bordersXml = XMLParser.extractBetweenTags(tcPrXml, '<w:tcBorders>', '</w:tcBorders>');
    if (bordersXml) {
      formatting.borders = this.parseBordersFromXml(
        bordersXml,
        true
      ) as import('../formatting/Style').CellBorders;
    }

    // Parse shading
    if (tcPrXml.includes('<w:shd')) {
      formatting.shading = this.parseShadingFromXml(tcPrXml);
    }

    // Parse margins
    const marginXml = XMLParser.extractBetweenTags(tcPrXml, '<w:tcMar>', '</w:tcMar>');
    if (marginXml) {
      formatting.margins = this.parseCellMarginsFromXml(marginXml);
    }

    // Parse vertical alignment — ST_VerticalJc has four values
    // (top / center / both / bottom) per ECMA-376 §17.18.101. Previously
    // the whitelist only accepted the first three, silently dropping
    // `<w:vAlign w:val="both"/>` on cell styles.
    if (tcPrXml.includes('<w:vAlign')) {
      const tag = XMLParser.extractSelfClosingTag(tcPrXml, 'w:vAlign');
      if (tag) {
        const val = XMLParser.extractAttribute(`<w:vAlign${tag}`, 'w:val');
        if (val === 'top' || val === 'center' || val === 'both' || val === 'bottom') {
          formatting.verticalAlignment = val;
        }
      }
    }

    return formatting;
  }

  /**
   * Parses table row formatting from trPr XML (Phase 5.1)
   */
  private parseTableRowFormattingFromXml(
    trPrXml: string
  ): import('../formatting/Style').TableRowStyleFormatting {
    const formatting: import('../formatting/Style').TableRowStyleFormatting = {};

    // Parse height
    if (trPrXml.includes('<w:trHeight')) {
      const tag = XMLParser.extractSelfClosingTag(trPrXml, 'w:trHeight');
      if (tag) {
        const val = XMLParser.extractAttribute(`<w:trHeight${tag}`, 'w:val');
        const hRule = XMLParser.extractAttribute(`<w:trHeight${tag}`, 'w:hRule');
        if (val) {
          formatting.height = parseInt(val, 10);
        }
        if (hRule === 'auto' || hRule === 'exact' || hRule === 'atLeast') {
          formatting.heightRule = hRule;
        }
      }
    }

    // Parse cantSplit / tblHeader — both OnOffOnlyType (§17.4.6, §17.4.50).
    // Previous substring-include detection hard-coded the flag to `true`
    // whenever the element appeared, silently flipping an explicit-off
    // override (e.g., a tblStylePr conditional un-splitting a header row)
    // into an enabled flag. Reuse parseOnOffAttribute so bare, "on", and
    // "off" all map correctly, and so absent stays undefined.
    const parseTrPrOnOffOnly = (tagName: string): boolean | undefined => {
      const el = XMLParser.extractSelfClosingTag(trPrXml, tagName);
      if (el === undefined) return undefined;
      const val = XMLParser.extractAttribute(`<${tagName}${el}`, 'w:val');
      if (val === undefined) return true;
      return parseOnOffAttribute(val, true);
    };

    const cantSplitVal = parseTrPrOnOffOnly('w:cantSplit');
    if (cantSplitVal !== undefined) formatting.cantSplit = cantSplitVal;

    const tblHeaderVal = parseTrPrOnOffOnly('w:tblHeader');
    if (tblHeaderVal !== undefined) formatting.isHeader = tblHeaderVal;

    return formatting;
  }

  /**
   * Parses conditional formatting from style XML (Phase 5.1)
   */
  private parseConditionalFormattingFromXml(
    styleXml: string
  ): import('../formatting/Style').ConditionalTableFormatting[] | undefined {
    const conditionalFormatting: import('../formatting/Style').ConditionalTableFormatting[] = [];

    // Find all tblStylePr elements
    let searchFrom = 0;
    while (true) {
      const startIdx = styleXml.indexOf('<w:tblStylePr', searchFrom);
      if (startIdx === -1) break;

      const endIdx = styleXml.indexOf('</w:tblStylePr>', startIdx);
      if (endIdx === -1) break;

      const tblStylePrXml = styleXml.substring(startIdx, endIdx + 15); // 15 = length of "</w:tblStylePr>"

      // Extract type attribute
      const typeAttr = XMLParser.extractAttribute(tblStylePrXml, 'w:type');
      if (typeAttr) {
        const conditional: import('../formatting/Style').ConditionalTableFormatting = {
          type: typeAttr as import('../formatting/Style').ConditionalFormattingType,
        };

        // Parse pPr
        const pPrXml = XMLParser.extractBetweenTags(tblStylePrXml, '<w:pPr>', '</w:pPr>');
        if (pPrXml) {
          conditional.paragraphFormatting = this.parseParagraphFormattingFromXml(pPrXml);
        }

        // Parse rPr
        const rPrXml = XMLParser.extractBetweenTags(tblStylePrXml, '<w:rPr>', '</w:rPr>');
        if (rPrXml) {
          conditional.runFormatting = this.parseRunFormattingFromXml(rPrXml);
        }

        // Parse tblPr
        const tblPrXml = XMLParser.extractBetweenTags(tblStylePrXml, '<w:tblPr>', '</w:tblPr>');
        if (tblPrXml) {
          conditional.tableFormatting = this.parseTableFormattingFromXml(tblPrXml);
        }

        // Parse tcPr
        const tcPrXml = XMLParser.extractBetweenTags(tblStylePrXml, '<w:tcPr>', '</w:tcPr>');
        if (tcPrXml) {
          conditional.cellFormatting = this.parseTableCellFormattingFromXml(tcPrXml);
        }

        // Parse trPr
        const trPrXml = XMLParser.extractBetweenTags(tblStylePrXml, '<w:trPr>', '</w:trPr>');
        if (trPrXml) {
          conditional.rowFormatting = this.parseTableRowFormattingFromXml(trPrXml);
        }

        conditionalFormatting.push(conditional);
      }

      searchFrom = endIdx + 15;
    }

    return conditionalFormatting.length > 0 ? conditionalFormatting : undefined;
  }

  /**
   * Parses borders from XML (Phase 5.1)
   * @param bordersXml - XML content from tblBorders or tcBorders
   * @param includeDiagonals - Whether to include diagonal borders (for cells)
   */
  private parseBordersFromXml(
    bordersXml: string,
    includeDiagonals: boolean
  ): import('../formatting/Style').TableBorders | import('../formatting/Style').CellBorders {
    const borders: any = {};

    // Local helper so both the main-side loop and the diagonal loop share
    // the full CT_Border attribute set (§17.18.2): val / sz / space / color
    // / themeColor / themeTint / themeShade / shadow / frame. Previously
    // this parser only extracted the four "basic" attrs, so themed borders
    // and shadow/frame flags on page/table/cell borders were silently
    // dropped on every load → save round-trip.
    const parseBorderAttrs = (
      type: string
    ): import('../formatting/Style').BorderProperties | null => {
      const tag = XMLParser.extractSelfClosingTag(bordersXml, `w:${type}`);
      if (!tag) return null;
      const ref = `<w:${type}${tag}`;
      const border: import('../formatting/Style').BorderProperties = {};
      const style = XMLParser.extractAttribute(ref, 'w:val');
      const size = XMLParser.extractAttribute(ref, 'w:sz');
      const space = XMLParser.extractAttribute(ref, 'w:space');
      const color = XMLParser.extractAttribute(ref, 'w:color');
      const themeColor = XMLParser.extractAttribute(ref, 'w:themeColor');
      const themeTint = XMLParser.extractAttribute(ref, 'w:themeTint');
      const themeShade = XMLParser.extractAttribute(ref, 'w:themeShade');
      const shadow = XMLParser.extractAttribute(ref, 'w:shadow');
      const frame = XMLParser.extractAttribute(ref, 'w:frame');
      if (style) border.style = style as any;
      if (size) border.size = parseInt(size, 10);
      if (space) border.space = parseInt(space, 10);
      if (color) border.color = color;
      if (themeColor) (border as any).themeColor = themeColor;
      if (themeTint) (border as any).themeTint = themeTint;
      if (themeShade) (border as any).themeShade = themeShade;
      if (shadow !== undefined) (border as any).shadow = parseOnOffAttribute(shadow, true);
      if (frame !== undefined) (border as any).frame = parseOnOffAttribute(frame, true);
      return Object.keys(border).length > 0 ? border : null;
    };

    // Per ECMA-376 §17.4.40 CT_TblBorders and §17.4.66 CT_TcBorders the
    // left / right borders have bidi-aware aliases `w:start` / `w:end`.
    // Modern authoring tools (Word 2013+, Google Docs) emit the bidi-
    // aware form by default — prefer those over the legacy `w:left` /
    // `w:right` so bidi-authored tables round-trip their side borders
    // (the internal model stores under the left/right keys, matching
    // the emitter).
    const borderTypes = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];
    for (const type of borderTypes) {
      // For left/right: prefer bidi-aware start/end alias if present.
      const alias = type === 'left' ? 'start' : type === 'right' ? 'end' : type;
      const tagNameToRead = bordersXml.includes(`<w:${alias}`) ? alias : type;
      if (bordersXml.includes(`<w:${tagNameToRead}`)) {
        const border = parseBorderAttrs(tagNameToRead);
        if (border) borders[type] = border;
      }
    }

    // Add diagonal borders for cells
    if (includeDiagonals) {
      const diagonalTypes = ['tl2br', 'tr2bl'];
      for (const type of diagonalTypes) {
        if (bordersXml.includes(`<w:${type}`)) {
          const border = parseBorderAttrs(type);
          if (border) borders[type] = border;
        }
      }
    }

    return borders;
  }

  /**
   * Parses shading from an object-based XML representation (parseToObject format).
   * Extracts all 9 ECMA-376 shading attributes including theme colors.
   */
  private parseShadingFromObj(shd: any): ShadingConfig | undefined {
    // Per ECMA-376 §17.3.1.32 CT_Shd, every string-typed attribute
    // (ST_UcharHexNumber tint/shade, ST_ThemeColor theme refs,
    // ST_HexColor fill/color, ST_Shd pattern) can be purely numeric in
    // hex form — "80", "00", "FF", "80000000", etc. XMLParser's
    // `parseAttributeValue: true` coerces purely-digit hex strings like
    // "80" to the JS number 80, violating the `string` type on
    // ShadingConfig. Previously stored values leaked downstream as
    // numbers (e.g. `.toUpperCase()` would throw); cast every attribute
    // through `String(...)` so the declared-type contract holds.
    const shading: ShadingConfig = {};
    if (shd['@_w:val']) shading.pattern = String(shd['@_w:val']) as ShadingConfig['pattern'];
    if (shd['@_w:fill']) shading.fill = String(shd['@_w:fill']);
    if (shd['@_w:color']) shading.color = String(shd['@_w:color']);
    if (shd['@_w:themeFill']) shading.themeFill = String(shd['@_w:themeFill']);
    if (shd['@_w:themeColor']) shading.themeColor = String(shd['@_w:themeColor']);
    if (shd['@_w:themeFillTint']) shading.themeFillTint = String(shd['@_w:themeFillTint']);
    if (shd['@_w:themeFillShade']) shading.themeFillShade = String(shd['@_w:themeFillShade']);
    if (shd['@_w:themeTint']) shading.themeTint = String(shd['@_w:themeTint']);
    if (shd['@_w:themeShade']) shading.themeShade = String(shd['@_w:themeShade']);
    return Object.keys(shading).length > 0 ? shading : undefined;
  }

  /**
   * Parses shading from XML (Phase 5.1)
   */
  private parseShadingFromXml(xml: string): ShadingConfig | undefined {
    const tag = XMLParser.extractSelfClosingTag(xml, 'w:shd');
    if (!tag) return undefined;

    const shading: ShadingConfig = {};
    const fullTag = `<w:shd${tag}`;
    const val = XMLParser.extractAttribute(fullTag, 'w:val');
    const color = XMLParser.extractAttribute(fullTag, 'w:color');
    const fill = XMLParser.extractAttribute(fullTag, 'w:fill');
    const themeFill = XMLParser.extractAttribute(fullTag, 'w:themeFill');
    const themeColor = XMLParser.extractAttribute(fullTag, 'w:themeColor');
    const themeFillTint = XMLParser.extractAttribute(fullTag, 'w:themeFillTint');
    const themeFillShade = XMLParser.extractAttribute(fullTag, 'w:themeFillShade');
    const themeTint = XMLParser.extractAttribute(fullTag, 'w:themeTint');
    const themeShade = XMLParser.extractAttribute(fullTag, 'w:themeShade');

    if (val) shading.pattern = val as ShadingConfig['pattern'];
    if (color) shading.color = color;
    if (fill) shading.fill = fill;
    if (themeFill) shading.themeFill = themeFill;
    if (themeColor) shading.themeColor = themeColor;
    if (themeFillTint) shading.themeFillTint = themeFillTint;
    if (themeFillShade) shading.themeFillShade = themeFillShade;
    if (themeTint) shading.themeTint = themeTint;
    if (themeShade) shading.themeShade = themeShade;

    return Object.keys(shading).length > 0 ? shading : undefined;
  }

  /**
   * Parses cell margins from XML (Phase 5.1)
   */
  private parseCellMarginsFromXml(
    marginXml: string
  ): import('../formatting/Style').CellMargins | undefined {
    const margins: import('../formatting/Style').CellMargins = {};

    // Parse top and bottom directly
    for (const type of ['top', 'bottom'] as const) {
      if (marginXml.includes(`<w:${type}`)) {
        const tag = XMLParser.extractSelfClosingTag(marginXml, `w:${type}`);
        if (tag) {
          const w = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:w');
          if (w) {
            margins[type] = parseInt(w, 10);
          }
        }
      }
    }

    // Parse left/right with bidi-aware w:start/w:end fallback (ECMA-376 §17.4.42/§17.4.43)
    // w:start takes precedence over w:left; w:end takes precedence over w:right
    const leftTag = marginXml.includes('<w:start')
      ? XMLParser.extractSelfClosingTag(marginXml, 'w:start')
      : XMLParser.extractSelfClosingTag(marginXml, 'w:left');
    if (leftTag) {
      const tagName = marginXml.includes('<w:start') ? 'w:start' : 'w:left';
      const w = XMLParser.extractAttribute(`<${tagName}${leftTag}`, 'w:w');
      if (w) {
        margins.left = parseInt(w, 10);
      }
    }

    const rightTag = marginXml.includes('<w:end')
      ? XMLParser.extractSelfClosingTag(marginXml, 'w:end')
      : XMLParser.extractSelfClosingTag(marginXml, 'w:right');
    if (rightTag) {
      const tagName = marginXml.includes('<w:end') ? 'w:end' : 'w:right';
      const w = XMLParser.extractAttribute(`<${tagName}${rightTag}`, 'w:w');
      if (w) {
        margins.right = parseInt(w, 10);
      }
    }

    return Object.keys(margins).length > 0 ? margins : undefined;
  }

  /**
   * Helper: Gets raw XML string from a document part
   * Utility function for parser to retrieve unparsed XML content
   * @param zipHandler - ZIP handler containing the document
   * @param partName - Part path (e.g., 'word/document.xml')
   * @returns Raw XML string or null if not found
   */
  static getRawXml(zipHandler: ZipHandler, partName: string): string | null {
    try {
      const file = zipHandler.getFile(partName);
      if (!file) {
        return null;
      }

      // If already a string, return as-is
      if (typeof file.content === 'string') {
        return file.content;
      }

      // If Buffer, decode as UTF-8
      if (Buffer.isBuffer(file.content)) {
        return file.content.toString('utf8');
      }

      return null;
    } catch (error: unknown) {
      return null;
    }
  }

  /**
   * Helper: Updates raw XML content in a document part
   * Utility function for parser to set unparsed XML content
   * @param zipHandler - ZIP handler containing the document
   * @param partName - Part path (e.g., 'word/document.xml')
   * @param xmlContent - Raw XML string to set
   * @returns True if successful, false otherwise
   */
  static setRawXml(zipHandler: ZipHandler, partName: string, xmlContent: string): boolean {
    try {
      if (typeof xmlContent !== 'string') {
        return false;
      }

      // Add or update the file in the ZIP handler
      // Convert string to UTF-8 Buffer for consistent encoding
      zipHandler.addFile(partName, Buffer.from(xmlContent, 'utf8'), {
        binary: true,
      });
      return true;
    } catch (error: unknown) {
      return false;
    }
  }

  /**
   * Helper: Gets relationships for a specific document part
   * Utility function for parser to access .rels files
   * @param zipHandler - ZIP handler containing the document
   * @param partName - Part name to get relationships for (e.g., 'word/document.xml')
   * @returns Array of relationships for that part, or empty array if none found
   */
  static getRelationships(
    zipHandler: ZipHandler,
    partName: string
  ): {
    id?: string;
    type?: string;
    target?: string;
    targetMode?: string;
  }[] {
    try {
      // Construct the .rels path from the part name
      // For 'word/document.xml' -> 'word/_rels/document.xml.rels'
      const lastSlash = partName.lastIndexOf('/');
      const relsPath =
        lastSlash === -1
          ? `_rels/${partName}.rels`
          : `${partName.substring(0, lastSlash)}/_rels/${partName.substring(lastSlash + 1)}.rels`;

      const relsContent = zipHandler.getFileAsString(relsPath);
      if (!relsContent) {
        return [];
      }

      interface ParsedRelationship {
        id?: string;
        type?: string;
        target?: string;
        targetMode?: string;
      }

      const relationships: ParsedRelationship[] = [];

      // Parse relationship XML
      const relPattern = /<Relationship\s+([^>]+)\/>/g;
      let match;

      while ((match = relPattern.exec(relsContent)) !== null) {
        const attrs = match[1];
        if (!attrs) continue;

        const rel: ParsedRelationship = {};

        // Extract attributes
        const idMatch = /Id="([^"]+)"/.exec(attrs);
        const typeMatch = /Type="([^"]+)"/.exec(attrs);
        const targetMatch = /Target="([^"]+)"/.exec(attrs);
        const modeMatch = /TargetMode="([^"]+)"/.exec(attrs);

        if (idMatch) rel.id = idMatch[1];
        if (typeMatch) rel.type = typeMatch[1];
        if (targetMatch) rel.target = targetMatch[1];
        if (modeMatch) rel.targetMode = modeMatch[1];

        relationships.push(rel);
      }

      return relationships;
    } catch (error: unknown) {
      return [];
    }
  }

  /**
   * Parses and extracts all namespace declarations from the root <w:document> tag
   *
   * @param docXml - The full XML content of word/document.xml
   * @returns A record of namespace prefixes to their URIs
   *
   * **Known Limitation - Default Namespace Not Extracted:**
   *
   * This method intentionally does NOT extract the default namespace (`xmlns="..."`).
   * Only prefixed namespaces (`xmlns:w`, `xmlns:r`, etc.) are extracted.
   *
   * **Rationale:**
   * 1. Word documents use prefixed namespaces exclusively (w:p, w:r, w:t)
   * 2. Adding a default namespace to <w:document> causes document corruption
   * 3. Microsoft Word and LibreOffice reject documents with default namespaces
   * 4. ECMA-376 examples show only prefixed namespace declarations
   *
   * **Impact:**
   * - If an input document has a default namespace, it will be lost during round-trip
   * - This is rare (valid OOXML documents don't use default namespaces)
   * - Documents using default namespace are likely non-standard or corrupted
   *
   * **Per ECMA-376 Part 1 Section 11.3.4:**
   * The document element should use the WordprocessingML namespace with prefix.
   *
   * @see ECMA-376-1:2016 Part 1, Section 11.3.4
   */
  private parseNamespaces(docXml: string): Record<string, string> {
    const namespaces: Record<string, string> = {};
    const docTagMatch = /<w:document([^>]+)>/.exec(docXml);

    if (docTagMatch?.[1]) {
      const attributes = docTagMatch[1];
      const nsPattern = /xmlns:([^=]+)="([^"]+)"/g;
      let match;

      while ((match = nsPattern.exec(attributes)) !== null) {
        if (match[1] && match[2]) {
          namespaces[`xmlns:${match[1]}`] = match[2];
        }
      }

      // Capture mc:Ignorable attribute (critical for w14/w15/etc. compatibility).
      // Without this, Word treats extended namespace attributes (w14:paraId etc.)
      // in raw XML passthrough zones as invalid content and reports corruption.
      const ignorableMatch = /mc:Ignorable="([^"]+)"/.exec(attributes);
      if (ignorableMatch?.[1]) {
        namespaces['mc:Ignorable'] = ignorableMatch[1];
      }
    }

    return namespaces;
  }

  /**
   * Parses document background (w:background) per ECMA-376 Part 1 §17.2.1
   * The w:background element appears as a child of w:document, before w:body
   */
  private parseDocumentBackground(
    docXml: string
  ): { color?: string; themeColor?: string; themeTint?: string; themeShade?: string } | undefined {
    const bgMatch = /<w:background([^>]*?)\/>/.exec(docXml);
    if (!bgMatch?.[1]) return undefined;

    const attrStr = bgMatch[1];
    const result: { color?: string; themeColor?: string; themeTint?: string; themeShade?: string } =
      {};

    const colorMatch = /w:color="([^"]+)"/.exec(attrStr);
    if (colorMatch?.[1]) result.color = colorMatch[1];

    const themeColorMatch = /w:themeColor="([^"]+)"/.exec(attrStr);
    if (themeColorMatch?.[1]) result.themeColor = themeColorMatch[1];

    const themeTintMatch = /w:themeTint="([^"]+)"/.exec(attrStr);
    if (themeTintMatch?.[1]) result.themeTint = themeTintMatch[1];

    const themeShadeMatch = /w:themeShade="([^"]+)"/.exec(attrStr);
    if (themeShadeMatch?.[1]) result.themeShade = themeShadeMatch[1];

    return Object.keys(result).length > 0 ? result : undefined;
  }

  /**
   * Parses headers and footers from a loaded document
   * Extracts header/footer XML files and creates Header/Footer objects
   * @param zipHandler - ZIP handler containing the document
   * @param section - Parsed section with header/footer references
   * @param relationshipManager - Relationship manager to lookup targets
   * @param imageManager - Image manager for parsing images in headers/footers
   * @returns Object with parsed headers and footers
   */
  async parseHeadersAndFooters(
    zipHandler: ZipHandler,
    section: Section | null,
    relationshipManager: RelationshipManager,
    imageManager: ImageManager
  ): Promise<{
    headers: {
      header: import('../elements/Header').Header;
      relationshipId: string;
      filename: string;
    }[];
    footers: {
      footer: import('../elements/Footer').Footer;
      relationshipId: string;
      filename: string;
    }[];
  }> {
    const headers: {
      header: import('../elements/Header').Header;
      relationshipId: string;
      filename: string;
    }[] = [];
    const footers: {
      footer: import('../elements/Footer').Footer;
      relationshipId: string;
      filename: string;
    }[] = [];

    if (!section) {
      return { headers, footers };
    }

    const sectionProps = section.getProperties();

    // Parse headers
    // Track already-parsed headers by rId to avoid creating duplicates
    // when multiple section property types (default, first, even) reference the same header file
    const parsedHeadersByRId = new Map<string, import('../elements/Header').Header>();

    if (sectionProps.headers) {
      for (const [type, rId] of Object.entries(sectionProps.headers)) {
        if (!rId) continue;

        // Get relationship to find header filename
        const rel = relationshipManager.getRelationship(rId);
        if (!rel) continue;

        const headerTarget = rel.getTarget();

        // Reuse already-parsed header if same rId was already processed
        const existingHeader = parsedHeadersByRId.get(rId);
        if (existingHeader) {
          headers.push({
            header: existingHeader,
            relationshipId: rId,
            filename: headerTarget,
          });
          continue;
        }

        const headerPath = `word/${headerTarget}`;
        const headerXml = zipHandler.getFileAsString(headerPath);
        if (!headerXml) continue;

        // Load header-specific relationships (for images in headers)
        // Header relationships are in word/_rels/header1.xml.rels for word/header1.xml
        const headerRelsPath = `word/_rels/${headerTarget}.rels`;
        const headerRelsXml = zipHandler.getFileAsString(headerRelsPath);
        const headerRelManager = headerRelsXml
          ? RelationshipManager.fromXml(headerRelsXml)
          : relationshipManager;

        // Set current part name for image registration (distinguishes header images from body images)
        this.currentPartName = headerTarget;
        try {
          // Create Header object using header-specific relationship manager
          const header = await this.parseHeader(
            headerXml,
            type as 'default' | 'first' | 'even',
            zipHandler,
            headerRelManager,
            imageManager
          );
          if (header) {
            parsedHeadersByRId.set(rId, header);
            headers.push({
              header,
              relationshipId: rId,
              filename: headerTarget,
            });
          }
        } finally {
          this.currentPartName = undefined;
        }
      }
    }

    // Parse footers
    // Track already-parsed footers by rId to avoid creating duplicates
    // when multiple section property types (default, first, even) reference the same footer file
    const parsedFootersByRId = new Map<string, import('../elements/Footer').Footer>();

    if (sectionProps.footers) {
      for (const [type, rId] of Object.entries(sectionProps.footers)) {
        if (!rId) continue;

        // Get relationship to find footer filename
        const rel = relationshipManager.getRelationship(rId);
        if (!rel) continue;

        const footerTarget = rel.getTarget();

        // Reuse already-parsed footer if same rId was already processed
        const existingFooter = parsedFootersByRId.get(rId);
        if (existingFooter) {
          footers.push({
            footer: existingFooter,
            relationshipId: rId,
            filename: footerTarget,
          });
          continue;
        }

        const footerPath = `word/${footerTarget}`;
        const footerXml = zipHandler.getFileAsString(footerPath);
        if (!footerXml) continue;

        // Load footer-specific relationships (for images in footers)
        // Footer relationships are in word/_rels/footer1.xml.rels for word/footer1.xml
        const footerRelsPath = `word/_rels/${footerTarget}.rels`;
        const footerRelsXml = zipHandler.getFileAsString(footerRelsPath);
        const footerRelManager = footerRelsXml
          ? RelationshipManager.fromXml(footerRelsXml)
          : relationshipManager;

        // Set current part name for image registration (distinguishes footer images from body images)
        this.currentPartName = footerTarget;
        try {
          // Create Footer object using footer-specific relationship manager
          const footer = await this.parseFooter(
            footerXml,
            type as 'default' | 'first' | 'even',
            zipHandler,
            footerRelManager,
            imageManager
          );
          if (footer) {
            parsedFootersByRId.set(rId, footer);
            footers.push({
              footer,
              relationshipId: rId,
              filename: footerTarget,
            });
          }
        } finally {
          this.currentPartName = undefined;
        }
      }
    }

    return { headers, footers };
  }

  /**
   * Parses a single header XML file
   * @param headerXml - Header XML content
   * @param type - Header type (default, first, even)
   * @param zipHandler - ZIP handler for accessing resources
   * @param relationshipManager - Relationship manager for resolving links
   * @param imageManager - Image manager for registering images
   * @returns Parsed Header object or null
   */
  private async parseHeader(
    headerXml: string,
    type: 'default' | 'first' | 'even',
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager,
    imageManager: ImageManager
  ): Promise<Header | null> {
    try {
      const header = new Header({ type });

      // Store raw XML for preservation when saving
      header.setRawXML(headerXml);

      // Extract w:hdr content
      const hdrContent = XMLParser.extractBetweenTags(headerXml, '<w:hdr', '</w:hdr>');
      if (!hdrContent) {
        return header; // Empty header
      }

      // Parse paragraphs and tables within header
      const elements = await this.parseBodyElements(
        `<w:body>${hdrContent}</w:body>`,
        relationshipManager,
        zipHandler,
        imageManager
      );

      for (const element of elements) {
        if (element instanceof Paragraph) {
          header.addParagraph(element);
        } else if (element instanceof Table) {
          header.addTable(element);
        }
      }

      return header;
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'header', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse header: ${err.message}`);
      }

      return null;
    }
  }

  /**
   * Parses a single footer XML file
   * @param footerXml - Footer XML content
   * @param type - Footer type (default, first, even)
   * @param zipHandler - ZIP handler for accessing resources
   * @param relationshipManager - Relationship manager for resolving links
   * @param imageManager - Image manager for registering images
   * @returns Parsed Footer object or null
   */
  private async parseFooter(
    footerXml: string,
    type: 'default' | 'first' | 'even',
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager,
    imageManager: ImageManager
  ): Promise<Footer | null> {
    try {
      const footer = new Footer({ type });

      // Store raw XML for preservation when saving
      footer.setRawXML(footerXml);

      // Extract w:ftr content
      const ftrContent = XMLParser.extractBetweenTags(footerXml, '<w:ftr', '</w:ftr>');
      if (!ftrContent) {
        return footer; // Empty footer
      }

      // Parse paragraphs and tables within footer
      const elements = await this.parseBodyElements(
        `<w:body>${ftrContent}</w:body>`,
        relationshipManager,
        zipHandler,
        imageManager
      );

      for (const element of elements) {
        if (element instanceof Paragraph) {
          footer.addParagraph(element);
        } else if (element instanceof Table) {
          footer.addTable(element);
        }
      }

      return footer;
    } catch (error: unknown) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'footer', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse footer: ${err.message}`);
      }

      return null;
    }
  }

  /**
   * Parse previous properties from a *PrChange child element (parsed object format).
   * Handles w:tblPr, w:trPr, and w:tcPr children by extracting known property keys
   * into the Record<string, any> format expected by toXML() serialization.
   * Per ECMA-376 Part 1 §17.13.5.36-38
   */
  private parseGenericPreviousProperties(propsObj: any): Record<string, any> {
    if (!propsObj) return {};
    const result: Record<string, any> = {};

    // Table-level properties (w:tblPr context)
    if (propsObj['w:tblStyle']) {
      // w:tblStyle w:val is ST_String (§17.7.4.62). XMLParser coerces
      // purely-numeric style IDs (e.g. "2025") to numbers; cast so the
      // declared `string` contract holds on tracked-change history.
      const v = propsObj['w:tblStyle']['@_w:val'];
      result.style = v !== undefined && v !== null ? String(v) : '';
    }
    // tblpPr (floating table position) — mirror main-path zero-value
    // preservation. The tblPrChange emitter re-emits position via
    // `!== undefined`, so dropping zero-valued tracked "previous"
    // positions here lost them silently on round-trip.
    if (propsObj['w:tblpPr']) {
      const tblpPr = propsObj['w:tblpPr'];
      const pos: any = {};
      if (isExplicitlySet(tblpPr['@_w:tblpX'])) pos.x = safeParseInt(tblpPr['@_w:tblpX']);
      if (isExplicitlySet(tblpPr['@_w:tblpY'])) pos.y = safeParseInt(tblpPr['@_w:tblpY']);
      if (tblpPr['@_w:horzAnchor']) pos.horizontalAnchor = tblpPr['@_w:horzAnchor'];
      if (tblpPr['@_w:vertAnchor']) pos.verticalAnchor = tblpPr['@_w:vertAnchor'];
      if (isExplicitlySet(tblpPr['@_w:leftFromText'])) {
        pos.leftFromText = safeParseInt(tblpPr['@_w:leftFromText']);
      }
      if (isExplicitlySet(tblpPr['@_w:rightFromText'])) {
        pos.rightFromText = safeParseInt(tblpPr['@_w:rightFromText']);
      }
      if (isExplicitlySet(tblpPr['@_w:topFromText'])) {
        pos.topFromText = safeParseInt(tblpPr['@_w:topFromText']);
      }
      if (isExplicitlySet(tblpPr['@_w:bottomFromText'])) {
        pos.bottomFromText = safeParseInt(tblpPr['@_w:bottomFromText']);
      }
      if (Object.keys(pos).length > 0) result.position = pos;
    }
    if (propsObj['w:tblOverlap']) {
      result.overlap = propsObj['w:tblOverlap']['@_w:val'];
    }
    if (propsObj['w:bidiVisual']) {
      result.bidiVisual = parseOoxmlBoolean(propsObj['w:bidiVisual']);
    }
    if (propsObj['w:tblStyleRowBandSize']) {
      result.tblStyleRowBandSize = parseInt(
        propsObj['w:tblStyleRowBandSize']['@_w:val'] || '1',
        10
      );
    }
    if (propsObj['w:tblStyleColBandSize']) {
      result.tblStyleColBandSize = parseInt(
        propsObj['w:tblStyleColBandSize']['@_w:val'] || '1',
        10
      );
    }
    if (propsObj['w:tblW']) {
      result.width = parseInt(propsObj['w:tblW']['@_w:w'] || '0', 10);
      result.widthType = propsObj['w:tblW']['@_w:type'] || 'dxa';
    }
    if (propsObj['w:tblLayout']) {
      result.layout = propsObj['w:tblLayout']['@_w:type'];
    }
    if (propsObj['w:tblInd']) {
      result.indent = parseInt(propsObj['w:tblInd']['@_w:w'] || '0', 10);
      const indType = propsObj['w:tblInd']['@_w:type'];
      if (indType) result.indentType = indType;
    }
    if (propsObj['w:tblCellSpacing']) {
      result.cellSpacing = parseInt(propsObj['w:tblCellSpacing']['@_w:w'] || '0', 10);
      const csType = propsObj['w:tblCellSpacing']['@_w:type'];
      if (csType) result.cellSpacingType = csType;
    }
    if (propsObj['w:tblCellMar']) {
      const cellMar = propsObj['w:tblCellMar'];
      const margins: any = {};
      if (cellMar['w:top']) margins.top = parseInt(cellMar['w:top']['@_w:w'] || '0', 10);
      if (cellMar['w:bottom']) margins.bottom = parseInt(cellMar['w:bottom']['@_w:w'] || '0', 10);
      const leftSrc = cellMar['w:start'] || cellMar['w:left'];
      if (leftSrc) margins.left = parseInt(leftSrc['@_w:w'] || '0', 10);
      const rightSrc = cellMar['w:end'] || cellMar['w:right'];
      if (rightSrc) margins.right = parseInt(rightSrc['@_w:w'] || '0', 10);
      if (Object.keys(margins).length > 0) result.cellMargins = margins;
    }
    if (propsObj['w:tblBorders']) {
      const borders: any = {};
      const bordersObj = propsObj['w:tblBorders'];
      for (const side of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']) {
        // Prefer bidi-aware w:start/w:end aliases over legacy w:left/
        // w:right (ECMA-376 §17.4.40 CT_TblBorders). Same pattern as
        // the main table borders parser — the bidi-aware form is the
        // preferred modern spelling.
        const aliasKey = side === 'left' ? 'w:start' : side === 'right' ? 'w:end' : undefined;
        const borderObj = (aliasKey && bordersObj[aliasKey]) || bordersObj[`w:${side}`];
        if (borderObj) {
          borders[side] = this.parseBorderElement(borderObj);
        }
      }
      if (Object.keys(borders).length > 0) result.borders = borders;
    }
    // tblLook per ECMA-376 §17.4.57 — supports both hex-string format
    // (w:val="04A0") AND the expanded individual-attribute form
    // (firstRow/lastRow/firstColumn/lastColumn/noHBand/noVBand).
    // Word often emits the expanded form (no w:val) inside *PrChange
    // previous-properties; the hex-only read silently collapsed every
    // flag to "0000" on round-trip.
    if (propsObj['w:tblLook']) {
      const look = propsObj['w:tblLook'];
      if (look['@_w:val']) {
        result.tblLook = String(look['@_w:val']);
      } else {
        const attrIsOn = (name: string): boolean => {
          const v = look[name];
          if (v === undefined) return false;
          return parseOoxmlBoolean({ '@_w:val': v });
        };
        let value = 0;
        if (attrIsOn('@_w:firstRow')) value |= 0x0020;
        if (attrIsOn('@_w:lastRow')) value |= 0x0040;
        if (attrIsOn('@_w:firstColumn')) value |= 0x0080;
        if (attrIsOn('@_w:lastColumn')) value |= 0x0100;
        if (attrIsOn('@_w:noHBand')) value |= 0x0200;
        if (attrIsOn('@_w:noVBand')) value |= 0x0400;
        result.tblLook = value.toString(16).toUpperCase().padStart(4, '0');
      }
    }
    if (propsObj['w:tblCaption']) {
      // w:tblCaption w:val is ST_String (§17.4.62). Cast through
      // String() so purely-numeric caption text round-trips as a
      // string inside the tracked-change previousProperties.
      const v = propsObj['w:tblCaption']['@_w:val'];
      result.caption = v !== undefined && v !== null ? String(v) : undefined;
    }
    if (propsObj['w:tblDescription']) {
      // w:tblDescription w:val is ST_String (§17.4.63).
      const v = propsObj['w:tblDescription']['@_w:val'];
      result.description = v !== undefined && v !== null ? String(v) : undefined;
    }

    // Row-level properties (w:trPr context) — all CT_TrPr elements
    if (propsObj['w:cnfStyle']) {
      result.cnfStyle = propsObj['w:cnfStyle']['@_w:val'];
    }
    if (propsObj['w:divId']) {
      result.divId = propsObj['w:divId']['@_w:val'];
    }
    if (propsObj['w:gridBefore']) {
      result.gridBefore = parseInt(propsObj['w:gridBefore']['@_w:val'] || '0', 10);
    }
    if (propsObj['w:gridAfter']) {
      result.gridAfter = parseInt(propsObj['w:gridAfter']['@_w:val'] || '0', 10);
    }
    if (propsObj['w:wBefore']) {
      result.wBefore = parseInt(propsObj['w:wBefore']['@_w:w'] || '0', 10);
      const wbType = propsObj['w:wBefore']['@_w:type'];
      if (wbType) result.wBeforeType = wbType;
    }
    if (propsObj['w:wAfter']) {
      result.wAfter = parseInt(propsObj['w:wAfter']['@_w:w'] || '0', 10);
      const waType = propsObj['w:wAfter']['@_w:type'];
      if (waType) result.wAfterType = waType;
    }
    if (propsObj['w:trHeight']) {
      result.height = parseInt(propsObj['w:trHeight']['@_w:val'] || '0', 10);
      const rule = propsObj['w:trHeight']['@_w:hRule'];
      if (rule) result.heightRule = rule;
    }
    // Row CT_OnOff — honour w:val per ECMA-376 §17.17.4 (ST_OnOff)
    if (propsObj['w:tblHeader']) {
      result.isHeader = parseOoxmlBoolean(propsObj['w:tblHeader']);
    }
    if (propsObj['w:cantSplit']) {
      result.cantSplit = parseOoxmlBoolean(propsObj['w:cantSplit']);
    }
    if (propsObj['w:hidden']) {
      result.hidden = parseOoxmlBoolean(propsObj['w:hidden']);
    }

    // Cell-level properties (w:tcPr context) — all CT_TcPr elements
    if (propsObj['w:tcW']) {
      result.width = parseInt(propsObj['w:tcW']['@_w:w'] || '0', 10);
      result.widthType = propsObj['w:tcW']['@_w:type'] || 'dxa';
    }
    if (propsObj['w:gridSpan']) {
      result.columnSpan = parseInt(propsObj['w:gridSpan']['@_w:val'] || '1', 10);
    }
    if (propsObj['w:hMerge']) {
      result.hMerge = propsObj['w:hMerge']['@_w:val'] || 'continue';
    }
    if (propsObj['w:vMerge']) {
      result.vMerge = propsObj['w:vMerge']['@_w:val'] || 'continue';
    }
    if (propsObj['w:tcBorders']) {
      const borders: any = {};
      const bordersObj = propsObj['w:tcBorders'];
      for (const side of ['top', 'bottom', 'left', 'right', 'tl2br', 'tr2bl']) {
        // Prefer bidi-aware w:start/w:end aliases for left/right
        // (ECMA-376 §17.4.66 CT_TcBorders). Diagonals (tl2br/tr2bl)
        // have no bidi aliases.
        const aliasKey = side === 'left' ? 'w:start' : side === 'right' ? 'w:end' : undefined;
        const borderObj = (aliasKey && bordersObj[aliasKey]) || bordersObj[`w:${side}`];
        if (borderObj) {
          borders[side] = this.parseBorderElement(borderObj);
        }
      }
      if (Object.keys(borders).length > 0) result.borders = borders;
    }
    if (propsObj['w:noWrap']) {
      result.noWrap = parseOoxmlBoolean(propsObj['w:noWrap']);
    }
    if (propsObj['w:tcMar']) {
      const tcMar = propsObj['w:tcMar'];
      const margins: any = {};
      if (tcMar['w:top']) margins.top = parseInt(tcMar['w:top']['@_w:w'] || '0', 10);
      if (tcMar['w:bottom']) margins.bottom = parseInt(tcMar['w:bottom']['@_w:w'] || '0', 10);
      const leftSrc = tcMar['w:start'] || tcMar['w:left'];
      if (leftSrc) margins.left = parseInt(leftSrc['@_w:w'] || '0', 10);
      const rightSrc = tcMar['w:end'] || tcMar['w:right'];
      if (rightSrc) margins.right = parseInt(rightSrc['@_w:w'] || '0', 10);
      if (Object.keys(margins).length > 0) result.margins = margins;
    }
    if (propsObj['w:textDirection']) {
      result.textDirection = propsObj['w:textDirection']['@_w:val'];
    }
    if (propsObj['w:tcFitText']) {
      result.fitText = parseOoxmlBoolean(propsObj['w:tcFitText']);
    }
    if (propsObj['w:vAlign']) {
      result.verticalAlignment = propsObj['w:vAlign']['@_w:val'];
    }
    if (propsObj['w:hideMark']) {
      result.hideMark = parseOoxmlBoolean(propsObj['w:hideMark']);
    }
    if (propsObj['w:cnfStyle']) {
      result.cnfStyle = propsObj['w:cnfStyle']['@_w:val'];
    }

    // Shared properties (appear in multiple contexts)
    if (propsObj['w:jc']) {
      const val = propsObj['w:jc']['@_w:val'];
      if (val) {
        result.alignment = val; // Table context
        result.justification = val; // Row context
      }
    }
    if (propsObj['w:shd']) {
      const shading = this.parseShadingFromObj(propsObj['w:shd']);
      if (shading) result.shading = shading;
    }

    return result;
  }

  /**
   * Parse previous section properties from raw XML within w:sectPrChange.
   * Per ECMA-376 Part 1 §17.13.5.32
   */
  private parsePreviousSectionProperties(sectPrXml: string): Record<string, any> {
    if (!sectPrXml) return {};
    const result: Record<string, any> = {};

    // Footnote properties (w:footnotePr) per §17.11.9. The main sectPr parser
    // reads these, and the emitter supports prev.footnotePr, but the
    // sectPrChange parser previously dropped them entirely — tracked history
    // of changes to footnote numbering format / position / start / restart
    // was lost on every round-trip.
    const footnotePrElements = XMLParser.extractElements(sectPrXml, 'w:footnotePr');
    if (footnotePrElements.length > 0 && footnotePrElements[0]) {
      const fnPr = footnotePrElements[0];
      const fnObj: any = {};
      const posElements = XMLParser.extractElements(fnPr, 'w:pos');
      if (posElements[0]) {
        const pos = XMLParser.extractAttribute(posElements[0], 'w:val');
        if (pos) fnObj.position = pos;
      }
      const numFmtElements = XMLParser.extractElements(fnPr, 'w:numFmt');
      if (numFmtElements[0]) {
        const fmt = XMLParser.extractAttribute(numFmtElements[0], 'w:val');
        if (fmt) fnObj.numberFormat = fmt;
      }
      const numStartElements = XMLParser.extractElements(fnPr, 'w:numStart');
      if (numStartElements[0]) {
        const start = XMLParser.extractAttribute(numStartElements[0], 'w:val');
        if (start !== undefined) fnObj.startNumber = parseInt(String(start), 10);
      }
      const numRestartElements = XMLParser.extractElements(fnPr, 'w:numRestart');
      if (numRestartElements[0]) {
        const restart = XMLParser.extractAttribute(numRestartElements[0], 'w:val');
        if (restart) fnObj.restart = restart;
      }
      if (Object.keys(fnObj).length > 0) result.footnotePr = fnObj;
    }

    // Endnote properties (w:endnotePr) per §17.11.5 — mirror of footnotePr.
    const endnotePrElements = XMLParser.extractElements(sectPrXml, 'w:endnotePr');
    if (endnotePrElements.length > 0 && endnotePrElements[0]) {
      const enPr = endnotePrElements[0];
      const enObj: any = {};
      const posElements = XMLParser.extractElements(enPr, 'w:pos');
      if (posElements[0]) {
        const pos = XMLParser.extractAttribute(posElements[0], 'w:val');
        if (pos) enObj.position = pos;
      }
      const numFmtElements = XMLParser.extractElements(enPr, 'w:numFmt');
      if (numFmtElements[0]) {
        const fmt = XMLParser.extractAttribute(numFmtElements[0], 'w:val');
        if (fmt) enObj.numberFormat = fmt;
      }
      const numStartElements = XMLParser.extractElements(enPr, 'w:numStart');
      if (numStartElements[0]) {
        const start = XMLParser.extractAttribute(numStartElements[0], 'w:val');
        if (start !== undefined) enObj.startNumber = parseInt(String(start), 10);
      }
      const numRestartElements = XMLParser.extractElements(enPr, 'w:numRestart');
      if (numRestartElements[0]) {
        const restart = XMLParser.extractAttribute(numRestartElements[0], 'w:val');
        if (restart) enObj.restart = restart;
      }
      if (Object.keys(enObj).length > 0) result.endnotePr = enObj;
    }

    // Paper source (w:paperSrc) per §17.6.12 CT_PaperSource — first-page / other
    // paper tray selection. Both attributes optional per schema.
    const paperSrcElements = XMLParser.extractElements(sectPrXml, 'w:paperSrc');
    if (paperSrcElements.length > 0 && paperSrcElements[0]) {
      const ps = paperSrcElements[0];
      const psObj: any = {};
      const first = XMLParser.extractAttribute(ps, 'w:first');
      if (first !== undefined) psObj.first = parseInt(String(first), 10);
      const other = XMLParser.extractAttribute(ps, 'w:other');
      if (other !== undefined) psObj.other = parseInt(String(other), 10);
      if (Object.keys(psObj).length > 0) result.paperSource = psObj;
    }

    // Page size
    const pgSzElements = XMLParser.extractElements(sectPrXml, 'w:pgSz');
    if (pgSzElements.length > 0 && pgSzElements[0]) {
      const pgSz = pgSzElements[0];
      const width = XMLParser.extractAttribute(pgSz, 'w:w');
      const height = XMLParser.extractAttribute(pgSz, 'w:h');
      const orient = XMLParser.extractAttribute(pgSz, 'w:orient');
      const code = XMLParser.extractAttribute(pgSz, 'w:code');
      if (width || height) {
        result.pageSize = {
          width: width ? parseInt(width, 10) : undefined,
          height: height ? parseInt(height, 10) : undefined,
          orientation: orient === 'landscape' ? 'landscape' : 'portrait',
          code: code ? parseInt(code, 10) : undefined,
        };
      }
    }

    // Margins — full CT_PageMar attribute set (§17.6.11) including w:gutter
    // (the book-binding margin). Previously gutter was dropped on sectPrChange
    // history, so any tracked change to a binding-gutter value lost the
    // previous value on round-trip.
    const pgMarElements = XMLParser.extractElements(sectPrXml, 'w:pgMar');
    if (pgMarElements.length > 0 && pgMarElements[0]) {
      const pgMar = pgMarElements[0];
      const margins: any = {};
      const top = XMLParser.extractAttribute(pgMar, 'w:top');
      if (top) margins.top = parseInt(top, 10);
      const bottom = XMLParser.extractAttribute(pgMar, 'w:bottom');
      if (bottom) margins.bottom = parseInt(bottom, 10);
      const left = XMLParser.extractAttribute(pgMar, 'w:left');
      if (left) margins.left = parseInt(left, 10);
      const right = XMLParser.extractAttribute(pgMar, 'w:right');
      if (right) margins.right = parseInt(right, 10);
      const header = XMLParser.extractAttribute(pgMar, 'w:header');
      if (header) margins.header = parseInt(header, 10);
      const footer = XMLParser.extractAttribute(pgMar, 'w:footer');
      if (footer) margins.footer = parseInt(footer, 10);
      const gutter = XMLParser.extractAttribute(pgMar, 'w:gutter');
      if (gutter) margins.gutter = parseInt(gutter, 10);
      if (Object.keys(margins).length > 0) result.margins = margins;
    }

    // Section type
    const typeElements = XMLParser.extractElements(sectPrXml, 'w:type');
    if (typeElements.length > 0 && typeElements[0]) {
      const val = XMLParser.extractAttribute(typeElements[0], 'w:val');
      if (val) result.type = val;
    }

    // Line numbering
    const lnNumElements = XMLParser.extractElements(sectPrXml, 'w:lnNumType');
    if (lnNumElements.length > 0 && lnNumElements[0]) {
      const ln = lnNumElements[0];
      const lnObj: any = {};
      const countBy = XMLParser.extractAttribute(ln, 'w:countBy');
      if (countBy) lnObj.countBy = parseInt(countBy, 10);
      const start = XMLParser.extractAttribute(ln, 'w:start');
      if (start) lnObj.start = parseInt(start, 10);
      const restart = XMLParser.extractAttribute(ln, 'w:restart');
      if (restart) lnObj.restart = restart;
      const distance = XMLParser.extractAttribute(ln, 'w:distance');
      if (distance) lnObj.distance = parseInt(distance, 10);
      if (Object.keys(lnObj).length > 0) result.lineNumbering = lnObj;
    }

    // Page numbering — full CT_PageNumber attribute set (§17.6.12):
    // fmt / start / chapStyle / chapSep. Previously only fmt+start were read,
    // so tracked-change history of chapter-numbering edits (e.g. switching
    // from "Heading 1" to "Heading 2" as the chapter marker, or changing the
    // chapter separator from hyphen to emDash) lost the previous values.
    // The Section.ts emitter stores chapStyle / chapSep as top-level
    // properties rather than on pageNumbering, so expose them the same way.
    const pgNumElements = XMLParser.extractElements(sectPrXml, 'w:pgNumType');
    if (pgNumElements.length > 0 && pgNumElements[0]) {
      const pn = pgNumElements[0];
      const pnObj: any = {};
      const pnStart = XMLParser.extractAttribute(pn, 'w:start');
      if (pnStart) pnObj.start = parseInt(pnStart, 10);
      const fmt = XMLParser.extractAttribute(pn, 'w:fmt');
      if (fmt) pnObj.format = fmt;
      if (Object.keys(pnObj).length > 0) result.pageNumbering = pnObj;
      // Mirror the main-sectPr parser: chapStyle / chapSep live at the root
      // of the section properties, not inside pageNumbering.
      const chapStyle = XMLParser.extractAttribute(pn, 'w:chapStyle');
      if (chapStyle !== undefined) result.chapStyle = parseInt(String(chapStyle), 10);
      const chapSep = XMLParser.extractAttribute(pn, 'w:chapSep');
      if (chapSep) result.chapSep = chapSep;
    }

    // Columns
    const colsElements = XMLParser.extractElements(sectPrXml, 'w:cols');
    if (colsElements.length > 0 && colsElements[0]) {
      const cols = colsElements[0];
      const num = XMLParser.extractAttribute(cols, 'w:num');
      const space = XMLParser.extractAttribute(cols, 'w:space');
      // Full CT_Columns attribute set (§17.6.4): num / space / equalWidth / sep
      // plus the child <w:col w:w="..." w:space="..."/> entries for per-column
      // widths. Previously only num+space were read, so sectPrChange history of
      // a columns-layout change dropped equalWidth, the separator line, and
      // the entire custom column-width / per-column-space configuration.
      const equalWidth = XMLParser.extractAttribute(cols, 'w:equalWidth');
      const sep = XMLParser.extractAttribute(cols, 'w:sep');

      // Extract individual <w:col> children for non-equal-width layouts.
      const colChildElements = XMLParser.extractElements(cols, 'w:col');
      const columnWidths: number[] = [];
      const columnSpaces: number[] = [];
      let hasColumnSpaces = false;
      for (const col of colChildElements) {
        const width = XMLParser.extractAttribute(col, 'w:w');
        if (width) columnWidths.push(parseInt(width, 10));
        const colSpace = XMLParser.extractAttribute(col, 'w:space');
        if (colSpace) {
          columnSpaces.push(parseInt(colSpace, 10));
          hasColumnSpaces = true;
        } else {
          columnSpaces.push(0);
        }
      }

      if (num) {
        result.columns = {
          count: parseInt(num, 10),
          space: space ? parseInt(space, 10) : undefined,
          equalWidth: equalWidth ? parseOnOffAttribute(equalWidth) : undefined,
          separator: sep ? parseOnOffAttribute(sep) : undefined,
          columnWidths: columnWidths.length > 0 ? columnWidths : undefined,
          columnSpaces: hasColumnSpaces ? columnSpaces : undefined,
        };
      }
    }

    // CT_OnOff sectPr flags — honour w:val per ECMA-376 §17.17.4 (ST_OnOff).
    // Previously these used substring `.includes()`, which both ignored w:val
    // (flipping explicit false to true) and could false-positive on prefix
    // matches (e.g. "<w:bidi" inside "<w:bidiVisual"). Use extractElements +
    // extractAttribute + parseOnOffAttribute instead.
    const parseSectCtOnOff = (tagName: string): boolean | undefined => {
      const els = XMLParser.extractElements(sectPrXml, tagName);
      if (els.length === 0 || !els[0]) return undefined;
      const v = XMLParser.extractAttribute(els[0], 'w:val');
      return parseOnOffAttribute(v, true);
    };

    // Form protection (w:formProt) — CT_OnOff
    const formProtVal = parseSectCtOnOff('w:formProt');
    if (formProtVal !== undefined) result.formProt = formProtVal;

    // Vertical alignment
    const vAlignElements = XMLParser.extractElements(sectPrXml, 'w:vAlign');
    if (vAlignElements.length > 0 && vAlignElements[0]) {
      const val = XMLParser.extractAttribute(vAlignElements[0], 'w:val');
      if (val) result.verticalAlignment = val;
    }

    // Suppress endnotes (w:noEndnote) — CT_OnOff
    const noEndnoteVal = parseSectCtOnOff('w:noEndnote');
    if (noEndnoteVal !== undefined) result.noEndnote = noEndnoteVal;

    // Title page (w:titlePg) — CT_OnOff
    const titlePgVal = parseSectCtOnOff('w:titlePg');
    if (titlePgVal !== undefined) result.titlePage = titlePgVal;

    // Text direction
    const textDirElements = XMLParser.extractElements(sectPrXml, 'w:textDirection');
    if (textDirElements.length > 0 && textDirElements[0]) {
      const val = XMLParser.extractAttribute(textDirElements[0], 'w:val');
      if (val) result.textDirection = val;
    }

    // Bidi section (w:bidi) — CT_OnOff
    const bidiVal = parseSectCtOnOff('w:bidi');
    if (bidiVal !== undefined) result.bidi = bidiVal;

    // RTL gutter (w:rtlGutter) — CT_OnOff
    const rtlGutterVal = parseSectCtOnOff('w:rtlGutter');
    if (rtlGutterVal !== undefined) result.rtlGutter = rtlGutterVal;

    // Document grid
    const docGridElements = XMLParser.extractElements(sectPrXml, 'w:docGrid');
    if (docGridElements.length > 0 && docGridElements[0]) {
      const dg = docGridElements[0];
      const dgObj: any = {};
      const dgType = XMLParser.extractAttribute(dg, 'w:type');
      if (dgType) dgObj.type = dgType;
      const linePitch = XMLParser.extractAttribute(dg, 'w:linePitch');
      if (linePitch) dgObj.linePitch = parseInt(linePitch, 10);
      const charSpace = XMLParser.extractAttribute(dg, 'w:charSpace');
      if (charSpace) dgObj.charSpace = parseInt(charSpace, 10);
      if (Object.keys(dgObj).length > 0) result.docGrid = dgObj;
    }

    // Page borders (w:pgBorders) per ECMA-376 §17.6.10. The main sectPr parser
    // reads these, but the sectPrChange previous-sectPr parser previously
    // didn't — so a tracked-change history of page-border edits lost the
    // entire "previous" border configuration (style, color, themeColor,
    // themeTint, themeShade, shadow, frame) every round-trip. The emitter
    // supports prev.pageBorders already; this is the missing parser half.
    const pgBordersElements = XMLParser.extractElements(sectPrXml, 'w:pgBorders');
    if (pgBordersElements.length > 0 && pgBordersElements[0]) {
      const pgBordersXml = pgBordersElements[0];
      const pageBorders: any = {};
      const offsetFrom = XMLParser.extractAttribute(pgBordersXml, 'w:offsetFrom');
      if (offsetFrom) pageBorders.offsetFrom = offsetFrom;
      const display = XMLParser.extractAttribute(pgBordersXml, 'w:display');
      if (display) pageBorders.display = display;
      const zOrder = XMLParser.extractAttribute(pgBordersXml, 'w:zOrder');
      if (zOrder) pageBorders.zOrder = zOrder;

      // Per-side border parser mirrors the main-sectPr logic — full CT_Border
      // attribute set including themed colors and shadow/frame flags.
      const parsePrevBorder = (sideXml: string): any | undefined => {
        if (!sideXml) return undefined;
        const border: any = {};
        const val = XMLParser.extractAttribute(sideXml, 'w:val');
        if (val) border.style = val;
        const sz = XMLParser.extractAttribute(sideXml, 'w:sz');
        if (sz) border.size = parseInt(sz, 10);
        const color = XMLParser.extractAttribute(sideXml, 'w:color');
        if (color) border.color = color;
        const space = XMLParser.extractAttribute(sideXml, 'w:space');
        if (space) border.space = parseInt(space, 10);
        const shadow = XMLParser.extractAttribute(sideXml, 'w:shadow');
        if (shadow !== undefined) border.shadow = parseOnOffAttribute(shadow, true);
        const frame = XMLParser.extractAttribute(sideXml, 'w:frame');
        if (frame !== undefined) border.frame = parseOnOffAttribute(frame, true);
        const themeColor = XMLParser.extractAttribute(sideXml, 'w:themeColor');
        if (themeColor) border.themeColor = themeColor;
        const themeTint = XMLParser.extractAttribute(sideXml, 'w:themeTint');
        if (themeTint) border.themeTint = themeTint;
        const themeShade = XMLParser.extractAttribute(sideXml, 'w:themeShade');
        if (themeShade) border.themeShade = themeShade;
        const artId = XMLParser.extractAttribute(sideXml, 'w:id');
        if (artId) border.artId = parseInt(artId, 10);
        return Object.keys(border).length > 0 ? border : undefined;
      };

      for (const side of ['top', 'left', 'bottom', 'right']) {
        const sideElements = XMLParser.extractElements(pgBordersXml, `w:${side}`);
        if (sideElements.length > 0 && sideElements[0]) {
          const border = parsePrevBorder(sideElements[0]);
          if (border) pageBorders[side] = border;
        }
      }

      if (Object.keys(pageBorders).length > 0) {
        result.pageBorders = pageBorders;
      }
    }

    return result;
  }
}
