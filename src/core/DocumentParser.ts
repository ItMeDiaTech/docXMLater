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
import { PageNumberFormat, Section, SectionProperties, SectionType } from '../elements/Section';
import { StructuredDocumentTag } from '../elements/StructuredDocumentTag';
import { Table, TableBorder } from '../elements/Table';
import { TableCell } from '../elements/TableCell';
import { TableOfContents } from '../elements/TableOfContents';
import { TableOfContentsElement } from '../elements/TableOfContentsElement';
import { TableRow } from '../elements/TableRow';
import { AbstractNumbering } from '../formatting/AbstractNumbering';
import { NumberingInstance } from '../formatting/NumberingInstance';
import { Style, StyleProperties, StyleType } from '../formatting/Style';
import { logParagraphContent, logParsing, logTextDirection } from '../utils/diagnostics';
import { getGlobalLogger, createScopedLogger, ILogger, defaultLogger } from '../utils/logger';
import { safeParseInt, isExplicitlySet, parseOoxmlBoolean } from '../utils/parsingHelpers';
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
import { DocumentProperties } from './Document';
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
          const bookmark = new Bookmark({
            name: `_end_${id}`,
            id: id,
            skipNormalization: true,
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

      // Parse w14:paraId if present
      const paraId = pElement['w14:paraId'];
      if (paraId) {
        paragraph.formatting.paraId = paraId as string;
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

    // Extract revision XMLs from paragraph content for raw XML parsing
    const insXmls = XMLParser.extractElements(paraContent, 'w:ins');
    const delXmls = XMLParser.extractElements(paraContent, 'w:del');
    const moveFromXmls = XMLParser.extractElements(paraContent, 'w:moveFrom');
    const moveToXmls = XMLParser.extractElements(paraContent, 'w:moveTo');
    const bookmarkStartXmls = XMLParser.extractElements(paraContent, 'w:bookmarkStart');
    const bookmarkEndXmls = XMLParser.extractElements(paraContent, 'w:bookmarkEnd');

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
          const result = this.parseHyperlinkFromObject(
            hyperlinkArray[child.index],
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
      } else if (child.type === 'w:fldSimple') {
        const fields = pElement['w:fldSimple'];
        const fieldArray = Array.isArray(fields) ? fields : fields ? [fields] : [];
        if (child.index < fieldArray.length) {
          const field = this.parseSimpleFieldFromObject(fieldArray[child.index]);
          if (field) {
            paragraph.addField(field);
          }
        }
      } else if (child.type === 'w:ins') {
        if (child.index < insXmls.length) {
          const revisionXml = insXmls[child.index];
          if (revisionXml) {
            const revResult = await this.parseRevisionFromXml(
              revisionXml,
              'w:ins',
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revResult.revision) {
              paragraph.addRevision(revResult.revision);
            }
            // Add any bookmarks found inside the revision to the paragraph
            for (const bookmark of revResult.bookmarkStarts) {
              paragraph.addBookmarkStart(bookmark);
            }
            for (const bookmark of revResult.bookmarkEnds) {
              paragraph.addBookmarkEnd(bookmark);
            }
          }
        }
      } else if (child.type === 'w:del') {
        if (child.index < delXmls.length) {
          const revisionXml = delXmls[child.index];
          if (revisionXml) {
            const revResult = await this.parseRevisionFromXml(
              revisionXml,
              'w:del',
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revResult.revision) {
              paragraph.addRevision(revResult.revision);
            }
            // Add any bookmarks found inside the revision to the paragraph
            for (const bookmark of revResult.bookmarkStarts) {
              paragraph.addBookmarkStart(bookmark);
            }
            for (const bookmark of revResult.bookmarkEnds) {
              paragraph.addBookmarkEnd(bookmark);
            }
          }
        }
      } else if (child.type === 'w:moveFrom') {
        if (child.index < moveFromXmls.length) {
          const revisionXml = moveFromXmls[child.index];
          if (revisionXml) {
            const revResult = await this.parseRevisionFromXml(
              revisionXml,
              'w:moveFrom',
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revResult.revision) {
              paragraph.addRevision(revResult.revision);
            }
            // Add any bookmarks found inside the revision to the paragraph
            for (const bookmark of revResult.bookmarkStarts) {
              paragraph.addBookmarkStart(bookmark);
            }
            for (const bookmark of revResult.bookmarkEnds) {
              paragraph.addBookmarkEnd(bookmark);
            }
          }
        }
      } else if (child.type === 'w:moveTo') {
        if (child.index < moveToXmls.length) {
          const revisionXml = moveToXmls[child.index];
          if (revisionXml) {
            const revResult = await this.parseRevisionFromXml(
              revisionXml,
              'w:moveTo',
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revResult.revision) {
              paragraph.addRevision(revResult.revision);
            }
            // Add any bookmarks found inside the revision to the paragraph
            for (const bookmark of revResult.bookmarkStarts) {
              paragraph.addBookmarkStart(bookmark);
            }
            for (const bookmark of revResult.bookmarkEnds) {
              paragraph.addBookmarkEnd(bookmark);
            }
          }
        }
      } else if (child.type === 'w:bookmarkStart') {
        // Parse bookmark start element
        if (child.index < bookmarkStartXmls.length) {
          const bookmarkXml = bookmarkStartXmls[child.index];
          if (bookmarkXml) {
            const bookmark = this.parseBookmarkStart(bookmarkXml);
            if (bookmark) {
              paragraph.addBookmarkStart(bookmark);
            }
          }
        }
      } else if (child.type === 'w:bookmarkEnd') {
        // Parse bookmark end element
        if (child.index < bookmarkEndXmls.length) {
          const bookmarkXml = bookmarkEndXmls[child.index];
          if (bookmarkXml) {
            const bookmark = this.parseBookmarkEnd(bookmarkXml);
            if (bookmark) {
              paragraph.addBookmarkEnd(bookmark);
            }
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
      // Per ECMA-376, w:done="1" or "true" indicates resolved
      const done = doneAttr === '1' || doneAttr === 'true';

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
      const bookmark = new Bookmark({
        name: nameAttr,
        id: id,
        skipNormalization: true,
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

      // Create a placeholder bookmark for the end marker
      // The name doesn't matter for bookmarkEnd as it only uses the ID
      const bookmark = new Bookmark({
        name: `_end_${id}`,
        id: id,
        skipNormalization: true,
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

      // Parse w14:paraId attribute from paragraph element (Word 2010+ requirement)
      const paraId = paraObj['w14:paraId'];
      if (paraId) {
        paragraph.formatting.paraId = paraId;
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

    // Style
    if (pPrObj['w:pStyle']?.['@_w:val']) {
      paragraph.setStyle(pPrObj['w:pStyle']['@_w:val']);
    }

    // Indentation
    // Note: XMLParser converts numeric strings to numbers, so "0" becomes 0 (falsy)
    // Must use !== undefined check instead of truthy check to handle left="0"
    if (pPrObj['w:ind']) {
      const ind = pPrObj['w:ind'];
      // Use isExplicitlySet and safeParseInt for robust zero-value handling
      if (isExplicitlySet(ind['@_w:left'])) paragraph.setLeftIndent(safeParseInt(ind['@_w:left']));
      if (isExplicitlySet(ind['@_w:right']))
        paragraph.setRightIndent(safeParseInt(ind['@_w:right']));
      if (isExplicitlySet(ind['@_w:firstLine']))
        paragraph.setFirstLineIndent(safeParseInt(ind['@_w:firstLine']));
      // Parse hanging indent per ECMA-376 Part 1 §17.3.1.17
      if (isExplicitlySet(ind['@_w:hanging']))
        paragraph.setHangingIndent(safeParseInt(ind['@_w:hanging']));
    }

    // Spacing
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
    }

    // Keep properties - parse pageBreakBefore FIRST, then apply keep properties
    // This triggers automatic conflict resolution per ECMA-376 v0.28.2
    if (pPrObj['w:pageBreakBefore']) paragraph.formatting.pageBreakBefore = true;

    // Keep properties - these will automatically clear pageBreakBefore if both are set
    if (pPrObj['w:keepNext']) paragraph.setKeepNext(true);
    if (pPrObj['w:keepLines']) paragraph.setKeepLines(true);

    // Contextual spacing
    if (pPrObj['w:contextualSpacing']) paragraph.setContextualSpacing(true);

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

      // Helper function to parse border definition
      const parseBorder = (borderObj: any): any => {
        if (!borderObj) return undefined;
        const border: any = {};
        if (borderObj['@_w:val']) border.style = borderObj['@_w:val'];
        if (borderObj['@_w:sz'] !== undefined) border.size = safeParseInt(borderObj['@_w:sz']);
        if (borderObj['@_w:color']) border.color = borderObj['@_w:color'];
        if (borderObj['@_w:space'] !== undefined)
          border.space = safeParseInt(borderObj['@_w:space']);
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
        if (tabObj['@_w:pos']) tab.position = parseInt(tabObj['@_w:pos'], 10);
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
      const widowControlVal = pPrObj['w:widowControl']?.['@_w:val'];
      // Parse w:val attribute - can be "0"/"1" or "false"/"true"
      if (
        widowControlVal === '0' ||
        widowControlVal === 'false' ||
        widowControlVal === false ||
        widowControlVal === 0
      ) {
        paragraph.setWidowControl(false);
      } else {
        // If w:val is "1", "true", true, 1, or undefined (element present without val), default to true
        paragraph.setWidowControl(true);
      }
    }

    // Outline level per ECMA-376 Part 1 §17.3.1.19
    if (pPrObj['w:outlineLvl'] !== undefined && pPrObj['w:outlineLvl']['@_w:val'] !== undefined) {
      const level = parseInt(pPrObj['w:outlineLvl']['@_w:val'], 10);
      if (!isNaN(level) && level >= 0 && level <= 9) {
        paragraph.setOutlineLevel(level);
      }
    }

    // Suppress line numbers per ECMA-376 Part 1 §17.3.1.34
    if (pPrObj['w:suppressLineNumbers']) {
      paragraph.setSuppressLineNumbers(true);
    }

    // Bidirectional layout per ECMA-376 Part 1 §17.3.1.6
    if (pPrObj['w:bidi'] !== undefined) {
      const bidiVal = pPrObj['w:bidi']?.['@_w:val'];
      if (bidiVal === '0' || bidiVal === 'false' || bidiVal === false || bidiVal === 0) {
        paragraph.setBidi(false);
      } else {
        // Default is true when element present without val attribute or val="1"
        paragraph.setBidi(true);
      }
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
    if (pPrObj['w:mirrorIndents']) {
      paragraph.setMirrorIndents(true);
    }

    // Auto-adjust right indent per ECMA-376 Part 1 §17.3.1.1
    if (pPrObj['w:adjustRightInd'] !== undefined) {
      const adjustRightIndVal = pPrObj['w:adjustRightInd']?.['@_w:val'];
      if (
        adjustRightIndVal === '0' ||
        adjustRightIndVal === 'false' ||
        adjustRightIndVal === false ||
        adjustRightIndVal === 0
      ) {
        paragraph.setAdjustRightInd(false);
      } else {
        // Default is true when element present without val attribute or val="1"
        paragraph.setAdjustRightInd(true);
      }
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
    if (pPrObj['w:suppressAutoHyphens']) {
      paragraph.setSuppressAutoHyphens(true);
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
    if (pPrObj['w:suppressOverlap']) {
      paragraph.setSuppressOverlap(true);
    }

    // Textbox tight wrap per ECMA-376 Part 1 §17.3.1.37
    if (pPrObj['w:textboxTightWrap']) {
      const wrapVal = pPrObj['w:textboxTightWrap']?.['@_w:val'];
      if (wrapVal) {
        paragraph.setTextboxTightWrap(wrapVal);
      }
    }

    // HTML div ID per ECMA-376 Part 1 §17.3.1.9
    if (pPrObj['w:divId']) {
      const divIdVal = pPrObj['w:divId']?.['@_w:val'];
      if (divIdVal) {
        paragraph.setDivId(parseInt(divIdVal, 10));
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

    // Paragraph property change tracking per ECMA-376 Part 1 §17.3.1.27
    if (pPrObj['w:pPrChange']) {
      const changeObj = pPrObj['w:pPrChange'];
      const change: any = {};
      if (changeObj['@_w:author']) change.author = String(changeObj['@_w:author']);
      if (changeObj['@_w:date']) change.date = String(changeObj['@_w:date']);
      if (changeObj['@_w:id']) change.id = String(changeObj['@_w:id']);

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
        if (prevPPr['w:ind']) {
          const ind = prevPPr['w:ind'];
          previousProperties.indentation = {};
          if (ind['@_w:left'] !== undefined)
            previousProperties.indentation.left = parseInt(ind['@_w:left'], 10);
          if (ind['@_w:right'] !== undefined)
            previousProperties.indentation.right = parseInt(ind['@_w:right'], 10);
          if (ind['@_w:firstLine'] !== undefined)
            previousProperties.indentation.firstLine = parseInt(ind['@_w:firstLine'], 10);
          if (ind['@_w:hanging'] !== undefined)
            previousProperties.indentation.hanging = parseInt(ind['@_w:hanging'], 10);
        }

        // Parse previous alignment
        if (prevPPr['w:jc']?.['@_w:val']) {
          previousProperties.alignment = String(prevPPr['w:jc']['@_w:val']);
        }

        // Parse previous spacing
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
        }

        // Parse previous keepNext/keepLines/pageBreakBefore
        if (prevPPr['w:keepNext']) {
          previousProperties.keepNext = prevPPr['w:keepNext']['@_w:val'] !== '0';
        }
        if (prevPPr['w:keepLines']) {
          previousProperties.keepLines = prevPPr['w:keepLines']['@_w:val'] !== '0';
        }
        if (prevPPr['w:pageBreakBefore']) {
          previousProperties.pageBreakBefore = prevPPr['w:pageBreakBefore']['@_w:val'] !== '0';
        }

        // === Extended paragraph property parsing per ECMA-376 Part 1 §17.3.1 ===

        // Parse widowControl (w:widowControl) - orphan/widow control
        if (prevPPr['w:widowControl']) {
          previousProperties.widowControl = prevPPr['w:widowControl']['@_w:val'] !== '0';
        }

        // Parse suppressAutoHyphens (w:suppressAutoHyphens)
        if (prevPPr['w:suppressAutoHyphens']) {
          previousProperties.suppressAutoHyphens =
            prevPPr['w:suppressAutoHyphens']['@_w:val'] !== '0';
        }

        // Parse contextualSpacing (w:contextualSpacing)
        if (prevPPr['w:contextualSpacing']) {
          previousProperties.contextualSpacing = prevPPr['w:contextualSpacing']['@_w:val'] !== '0';
        }

        // Parse mirrorIndents (w:mirrorIndents)
        if (prevPPr['w:mirrorIndents']) {
          previousProperties.mirrorIndents = prevPPr['w:mirrorIndents']['@_w:val'] !== '0';
        }

        // Parse outlineLevel (w:outlineLvl @w:val)
        if (prevPPr['w:outlineLvl']?.['@_w:val'] !== undefined) {
          previousProperties.outlineLevel = parseInt(prevPPr['w:outlineLvl']['@_w:val'], 10);
        }

        // Parse bidi (w:bidi) - right-to-left paragraph
        if (prevPPr['w:bidi']) {
          previousProperties.bidi = prevPPr['w:bidi']['@_w:val'] !== '0';
        }

        // Parse suppressLineNumbers (w:suppressLineNumbers)
        if (prevPPr['w:suppressLineNumbers']) {
          previousProperties.suppressLineNumbers =
            prevPPr['w:suppressLineNumbers']['@_w:val'] !== '0';
        }

        // Parse adjustRightInd (w:adjustRightInd)
        if (prevPPr['w:adjustRightInd']) {
          previousProperties.adjustRightInd = prevPPr['w:adjustRightInd']['@_w:val'] !== '0';
        }

        // Parse snapToGrid (w:snapToGrid)
        if (prevPPr['w:snapToGrid']) {
          previousProperties.snapToGrid = prevPPr['w:snapToGrid']['@_w:val'] !== '0';
        }

        // Parse wordWrap (w:wordWrap)
        if (prevPPr['w:wordWrap']) {
          previousProperties.wordWrap = prevPPr['w:wordWrap']['@_w:val'] !== '0';
        }

        // Parse autoSpaceDE (w:autoSpaceDE) - East Asian/numeric spacing
        if (prevPPr['w:autoSpaceDE']) {
          previousProperties.autoSpaceDE = prevPPr['w:autoSpaceDE']['@_w:val'] !== '0';
        }

        // Parse autoSpaceDN (w:autoSpaceDN) - East Asian/Western spacing
        if (prevPPr['w:autoSpaceDN']) {
          previousProperties.autoSpaceDN = prevPPr['w:autoSpaceDN']['@_w:val'] !== '0';
        }

        // Parse textDirection (w:textDirection @w:val)
        if (prevPPr['w:textDirection']?.['@_w:val']) {
          previousProperties.textDirection = String(prevPPr['w:textDirection']['@_w:val']);
        }

        // Parse textAlignment (w:textAlignment @w:val) per ECMA-376 Part 1 §17.3.1.39
        if (prevPPr['w:textAlignment']?.['@_w:val']) {
          previousProperties.textAlignment = String(prevPPr['w:textAlignment']['@_w:val']);
        }

        // Parse paragraph borders (w:pBdr) per ECMA-376 Part 1 §17.3.1.24
        if (prevPPr['w:pBdr']) {
          const pBdr = prevPPr['w:pBdr'];
          previousProperties.borders = {};

          const parseBorder = (borderObj: any) => {
            if (!borderObj) return undefined;
            return {
              val: borderObj['@_w:val'],
              sz: borderObj['@_w:sz'] !== undefined ? parseInt(borderObj['@_w:sz'], 10) : undefined,
              space:
                borderObj['@_w:space'] !== undefined
                  ? parseInt(borderObj['@_w:space'], 10)
                  : undefined,
              color: borderObj['@_w:color'],
              themeColor: borderObj['@_w:themeColor'],
            };
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
              pos: tab['@_w:pos'] !== undefined ? parseInt(tab['@_w:pos'], 10) : undefined,
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
    let fieldRuns: Run[] = [];
    let fieldRevisions: Revision[] = []; // Track revisions inside field result section
    let instructionRevisions: Revision[] = []; // Track revisions in instruction area
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
                    fieldState = null;
                    break;
                  }

                  // Extract instruction to determine field type
                  let instruction = '';
                  let resultText = '';
                  let resultFormatting: RunFormatting | undefined;
                  let hasSeparate = false;

                  for (const run of fieldRuns) {
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
            // Inside a nested field - collect all runs to preserve raw structure
            fieldRuns.push(item);
          } else if (fieldState === 'separate' || fieldState === 'result') {
            // This run is part of the field result - collect it
            fieldRuns.push(item);
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
          fieldState = null;
          groupedContent.push(item);
        } else {
          // Normal revision outside of any field
          groupedContent.push(item);
        }
      } else {
        // Non-run content (hyperlinks, images, etc.)
        if (nestingDepth > 0) {
          // Inside a nested field - keep collecting
          fieldRuns.push(item as any);
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

      const parseBooleanAttr = (value: any): boolean | undefined => {
        if (value === undefined || value === null) {
          return undefined;
        }
        return value === '1' || value === 1 || value === true || value === 'true';
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
        // w:enabled (presence = true, w:val="0" = false)
        if (ffDataObj['w:enabled'] !== undefined) {
          const enabledVal = ffDataObj['w:enabled']?.['@_w:val'];
          ffd.enabled = enabledVal === '0' || enabledVal === 0 ? false : true;
        }
        // w:calcOnExit
        if (ffDataObj['w:calcOnExit'] !== undefined) {
          const calcVal = ffDataObj['w:calcOnExit']?.['@_w:val'];
          ffd.calcOnExit = calcVal === '1' || calcVal === 1 || calcVal === true;
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
          if (cb['w:default']?.['@_w:val'] !== undefined) {
            checkBox.defaultChecked =
              cb['w:default']['@_w:val'] === '1' || cb['w:default']['@_w:val'] === 1;
          }
          if (cb['w:checked']?.['@_w:val'] !== undefined) {
            checkBox.checked =
              cb['w:checked']['@_w:val'] === '1' || cb['w:checked']['@_w:val'] === 1;
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
              content.push({ type: 'break', breakType });
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

            // Footnote reference (w:footnoteReference) per ECMA-376 Part 1 §17.11.13
            case 'w:footnoteReference': {
              const fnRefElements = toArray(runObj['w:footnoteReference']);
              const fnRef = fnRefElements[elementIndex] || fnRefElements[0];
              const fnId = fnRef?.['@_w:id'];
              content.push({
                type: 'footnoteReference',
                footnoteId: fnId !== undefined ? parseInt(fnId, 10) : undefined,
              });
              break;
            }

            // Endnote reference (w:endnoteReference) per ECMA-376 Part 1 §17.11.2
            case 'w:endnoteReference': {
              const enRefElements = toArray(runObj['w:endnoteReference']);
              const enRef = enRefElements[elementIndex] || enRefElements[0];
              const enId = enRef?.['@_w:id'];
              content.push({
                type: 'endnoteReference',
                endnoteId: enId !== undefined ? parseInt(enId, 10) : undefined,
              });
              break;
            }

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
          content.push({ type: 'break', breakType });
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
        // Footnote/endnote reference fallback
        if (runObj['w:footnoteReference'] !== undefined) {
          const fnRefElements = toArray(runObj['w:footnoteReference']);
          for (const fnRef of fnRefElements) {
            const fnId = fnRef?.['@_w:id'];
            content.push({
              type: 'footnoteReference',
              footnoteId: fnId !== undefined ? parseInt(fnId, 10) : undefined,
            });
          }
        }
        if (runObj['w:endnoteReference'] !== undefined) {
          const enRefElements = toArray(runObj['w:endnoteReference']);
          for (const enRef of enRefElements) {
            const enId = enRef?.['@_w:id'];
            content.push({
              type: 'endnoteReference',
              endnoteId: enId !== undefined ? parseInt(enId, 10) : undefined,
            });
          }
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
          const name = bs['@_w:name'];
          if (id !== undefined && name) {
            const bookmark = new Bookmark({
              name: name,
              id: typeof id === 'number' ? id : parseInt(id, 10),
              skipNormalization: true,
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
            const bookmark = new Bookmark({
              name: `_end_${id}`,
              id: typeof id === 'number' ? id : parseInt(id, 10),
              skipNormalization: true,
            });
            result.bookmarkEnds.push(bookmark);
          }
        }
      }

      // Extract hyperlink attributes
      const relationshipId = hyperlinkObj['@_r:id'];
      const anchor = hyperlinkObj['@_w:anchor'];
      const tooltip = hyperlinkObj['@_w:tooltip'];
      const tgtFrame = hyperlinkObj['@_w:tgtFrame'];
      const history = hyperlinkObj['@_w:history'];
      const docLocation = hyperlinkObj['@_w:docLocation'];

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

      // Handle external hyperlinks with anchor fragments
      // Microsoft Word can store URLs with the base in relationships and fragment in w:anchor
      // Example: rels has "https://example.com/", anchor has "!/view?docid=abc-123"
      // Combined: "https://example.com/#!/view?docid=abc-123"
      // This is common for single-page applications with hash-based routing (theSource, etc.)
      let finalAnchor = anchor;
      let finalRelationshipId = relationshipId;
      if (url && anchor) {
        // Combine URL and anchor for external hyperlinks with fragments
        url = url + '#' + anchor;
        finalAnchor = undefined; // Clear anchor since it's now part of URL
        // Clear relationshipId since the relationship points to the old base URL
        // On save, a new relationship will be created with the combined URL
        finalRelationshipId = undefined;
        defaultLogger.debug(`[DocumentParser] Combined external URL with anchor fragment: ${url}`);
      }

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
      const typeMatch = instruction.trim().match(/^(\w+)/);
      const type = (typeMatch?.[1] || 'PAGE') as import('../elements/Field').FieldType;

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
        instruction,
        formatting,
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

    // Parse character style reference (w:rStyle) per ECMA-376 Part 1 §17.3.2.36
    if (rPrObj['w:rStyle']) {
      const styleId = rPrObj['w:rStyle']['@_w:val'];
      if (styleId) {
        run.setCharacterStyle(styleId);
      }
    }

    // Parse text border (w:bdr) per ECMA-376 Part 1 §17.3.2.5
    if (rPrObj['w:bdr']) {
      const bdr = rPrObj['w:bdr'];
      const border: any = {};
      if (bdr['@_w:val']) border.style = bdr['@_w:val'];
      if (bdr['@_w:sz']) border.size = parseInt(bdr['@_w:sz'], 10);
      if (bdr['@_w:color']) border.color = bdr['@_w:color'];
      if (bdr['@_w:space']) border.space = parseInt(bdr['@_w:space'], 10);
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

    // Parse outline text effect (w:outline) per ECMA-376 Part 1 §17.3.2.23
    if (rPrObj['w:outline']) run.setOutline(true);

    // Parse shadow text effect (w:shadow) per ECMA-376 Part 1 §17.3.2.32
    if (rPrObj['w:shadow']) run.setShadow(true);

    // Parse emboss text effect (w:emboss) per ECMA-376 Part 1 §17.3.2.13
    if (rPrObj['w:emboss']) run.setEmboss(true);

    // Parse imprint text effect (w:imprint) per ECMA-376 Part 1 §17.3.2.18
    if (rPrObj['w:imprint']) run.setImprint(true);

    // Parse no proofing (w:noProof) per ECMA-376 Part 1 §17.3.2.21
    if (rPrObj['w:noProof']) run.setNoProof(true);

    // Parse snap to grid (w:snapToGrid) per ECMA-376 Part 1 §17.3.2.35
    if (rPrObj['w:snapToGrid']) run.setSnapToGrid(true);

    // Parse vanish/hidden (w:vanish) per ECMA-376 Part 1 §17.3.2.42
    if (rPrObj['w:vanish']) run.setVanish(true);

    // Parse special vanish (w:specVanish) per ECMA-376 Part 1 §17.3.2.36
    if (rPrObj['w:specVanish']) run.setSpecVanish(true);

    // Boolean properties - use parseOoxmlBoolean helper
    // Per ECMA-376: <w:b/> or <w:b w:val="1"/> or <w:b w:val="true"/> means true
    // <w:b w:val="0"/> or <w:b w:val="false"/> means false (omit from document)

    // Parse RTL text (w:rtl) per ECMA-376 Part 1 §17.3.2.30
    if (parseOoxmlBoolean(rPrObj['w:rtl'])) run.setRTL(true);

    if (parseOoxmlBoolean(rPrObj['w:b'])) run.setBold(true);
    if (parseOoxmlBoolean(rPrObj['w:bCs'])) run.setComplexScriptBold(true);
    if (parseOoxmlBoolean(rPrObj['w:i'])) run.setItalic(true);
    if (parseOoxmlBoolean(rPrObj['w:iCs'])) run.setComplexScriptItalic(true);
    if (parseOoxmlBoolean(rPrObj['w:strike'])) run.setStrike(true);
    if (parseOoxmlBoolean(rPrObj['w:dstrike'])) {
      (run as any).formatting.dstrike = true;
    }
    if (parseOoxmlBoolean(rPrObj['w:smallCaps'])) run.setSmallCaps(true);
    if (parseOoxmlBoolean(rPrObj['w:caps'])) run.setAllCaps(true);

    // Parse complex script flag (w:cs) per ECMA-376 Part 1 §17.3.2.7
    if (parseOoxmlBoolean(rPrObj['w:cs'])) run.setComplexScript(true);

    // Parse web hidden (w:webHidden) per ECMA-376 Part 1 §17.3.2.44
    if (parseOoxmlBoolean(rPrObj['w:webHidden'])) run.setWebHidden(true);

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

    // Parse character spacing (w:spacing) per ECMA-376 Part 1 §17.3.2.33
    if (rPrObj['w:spacing']) {
      const val = rPrObj['w:spacing']['@_w:val'];
      if (val) run.setCharacterSpacing(parseInt(val, 10));
    }

    // Parse horizontal scaling (w:w) per ECMA-376 Part 1 §17.3.2.43
    if (rPrObj['w:w']) {
      const val = rPrObj['w:w']['@_w:val'];
      if (val) run.setScaling(parseInt(val, 10));
    }

    // Parse vertical position (w:position) per ECMA-376 Part 1 §17.3.2.31
    if (rPrObj['w:position']) {
      const val = rPrObj['w:position']['@_w:val'];
      if (val) run.setPosition(parseInt(val, 10));
    }

    // Parse kerning (w:kern) per ECMA-376 Part 1 §17.3.2.20
    if (rPrObj['w:kern']) {
      const val = rPrObj['w:kern']['@_w:val'];
      if (val) run.setKerning(parseInt(val, 10));
    }

    // Parse language (w:lang) per ECMA-376 Part 1 §17.3.2.20
    if (rPrObj['w:lang']) {
      const val = rPrObj['w:lang']['@_w:val'];
      if (val) run.setLanguage(val);
    }

    // Parse East Asian layout (w:eastAsianLayout) per ECMA-376 Part 1 §17.3.2.10
    if (rPrObj['w:eastAsianLayout']) {
      const layoutObj = rPrObj['w:eastAsianLayout'];
      const layout: any = {};
      if (layoutObj['@_w:id'] !== undefined) layout.id = Number(layoutObj['@_w:id']);
      if (layoutObj['@_w:vert']) layout.vert = true;
      if (layoutObj['@_w:vertCompress']) layout.vertCompress = true;
      if (layoutObj['@_w:combine']) layout.combine = true;
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
      if (val === 'superscript') run.setSuperscript(true);
    }

    if (rPrObj['w:rFonts']) {
      const rFonts = rPrObj['w:rFonts'];
      if (rFonts['@_w:ascii']) run.setFont(rFonts['@_w:ascii']);
      // Parse additional font variants per ECMA-376 Part 1 §17.3.2.26
      if (rFonts['@_w:hAnsi']) run.setFontHAnsi(rFonts['@_w:hAnsi']);
      if (rFonts['@_w:eastAsia']) run.setFontEastAsia(rFonts['@_w:eastAsia']);
      if (rFonts['@_w:cs']) run.setFontCs(rFonts['@_w:cs']);
      if (rFonts['@_w:hint']) run.setFontHint(rFonts['@_w:hint']);
      // Parse theme font references per ECMA-376 Part 1 §17.3.2.26
      if (rFonts['@_w:asciiTheme']) run.setFontAsciiTheme(rFonts['@_w:asciiTheme']);
      if (rFonts['@_w:hAnsiTheme']) run.setFontHAnsiTheme(rFonts['@_w:hAnsiTheme']);
      if (rFonts['@_w:eastAsiaTheme']) run.setFontEastAsiaTheme(rFonts['@_w:eastAsiaTheme']);
      if (rFonts['@_w:cstheme']) run.setFontCsTheme(rFonts['@_w:cstheme']);
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
      // Skip special OOXML values like "auto" (automatic/inherit from style)
      // "auto" is a valid OOXML color that means inherit - not a hex color
      if (colorVal && colorVal !== 'auto') {
        run.setColor(colorVal);
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

        // Parse previous underline
        if (prevRPr['w:u']) {
          const uVal = prevRPr['w:u']['@_w:val'];
          prevProps.underline = uVal || true;
        }

        // Parse previous strikethrough
        if (prevRPr['w:strike']) {
          prevProps.strike = parseOoxmlBoolean(prevRPr['w:strike']);
        }

        // Parse previous font (all w:rFonts attributes per ECMA-376 Part 1 §17.3.2.26)
        if (prevRPr['w:rFonts']) {
          const rFonts = prevRPr['w:rFonts'];
          if (rFonts['@_w:ascii']) prevProps.font = rFonts['@_w:ascii'];
          if (rFonts['@_w:hAnsi']) prevProps.fontHAnsi = rFonts['@_w:hAnsi'];
          if (rFonts['@_w:eastAsia']) prevProps.fontEastAsia = rFonts['@_w:eastAsia'];
          if (rFonts['@_w:cs']) prevProps.fontCs = rFonts['@_w:cs'];
          if (rFonts['@_w:hint']) prevProps.fontHint = rFonts['@_w:hint'];
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
          if (colorVal && colorVal !== 'auto') {
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

        // Parse previous subscript/superscript
        if (prevRPr['w:vertAlign']) {
          const val = prevRPr['w:vertAlign']['@_w:val'];
          if (val === 'subscript') prevProps.subscript = true;
          if (val === 'superscript') prevProps.superscript = true;
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

        // Parse language (w:lang @w:val)
        if (prevRPr['w:lang']) {
          const langVal = prevRPr['w:lang']['@_w:val'];
          if (langVal) {
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

        // Parse text border (w:bdr) per ECMA-376 Part 1 §17.3.2.4
        // Maps to TextBorder interface: style, size, color, space
        if (prevRPr['w:bdr']) {
          const bdrObj = prevRPr['w:bdr'];
          prevProps.border = {
            style: bdrObj['@_w:val'] as import('../elements/Run').TextBorderStyle,
            size: bdrObj['@_w:sz'] !== undefined ? safeParseInt(bdrObj['@_w:sz']) : undefined,
            space:
              bdrObj['@_w:space'] !== undefined ? safeParseInt(bdrObj['@_w:space']) : undefined,
            color: bdrObj['@_w:color'],
          };
        }

        // Parse character shading (w:shd) per ECMA-376 Part 1 §17.3.2.32
        if (prevRPr['w:shd']) {
          const shading = this.parseShadingFromObj(prevRPr['w:shd']);
          if (shading) {
            prevProps.shading = shading;
          }
        }

        // Parse East Asian layout (w:eastAsianLayout) per ECMA-376 Part 1 §17.3.2.10
        if (prevRPr['w:eastAsianLayout']) {
          const eaObj = prevRPr['w:eastAsianLayout'];
          prevProps.eastAsianLayout = {
            id: eaObj['@_w:id'] !== undefined ? safeParseInt(eaObj['@_w:id']) : undefined,
            combine: eaObj['@_w:combine']
              ? parseOoxmlBoolean({ '@_w:val': eaObj['@_w:combine'] })
              : undefined,
            combineBrackets: eaObj['@_w:combineBrackets'],
            vert: eaObj['@_w:vert']
              ? parseOoxmlBoolean({ '@_w:val': eaObj['@_w:vert'] })
              : undefined,
            vertCompress: eaObj['@_w:vertCompress']
              ? parseOoxmlBoolean({ '@_w:val': eaObj['@_w:vertCompress'] })
              : undefined,
          };
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
        name = docPrObj['@_name'] || 'image';
        description = docPrObj['@_descr'] || 'Image';
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
            border = { width: widthEmu / 12700 } as any;
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
          }
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
    const border: TableBorder = {
      style: (borderObj['@_w:val'] || 'single') as TableBorder['style'],
    };
    if (borderObj['@_w:sz'] !== undefined) border.size = safeParseInt(borderObj['@_w:sz']);
    if (borderObj['@_w:space'] !== undefined) border.space = safeParseInt(borderObj['@_w:space']);
    if (borderObj['@_w:color']) border.color = borderObj['@_w:color'];
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

    // Parse table style reference (w:tblStyle)
    if (tblPrObj['w:tblStyle']) {
      const styleId = tblPrObj['w:tblStyle']['@_w:val'];
      if (styleId) {
        table.setStyle(styleId);
      }
    }

    // Parse table look flags (w:tblLook) - for conditional formatting
    // Supports both hex string format (w:val="04A0") and individual attributes
    if (tblPrObj['w:tblLook']) {
      const look = tblPrObj['w:tblLook'];
      if (look['@_w:val']) {
        // Hex string format
        table.setTblLook(look['@_w:val']);
      } else {
        // Individual attribute format - construct hex value
        // Per ECMA-376: bit 0=firstRow, 1=lastRow, 2=firstCol, 3=lastCol, 4=noHBand, 5=noVBand
        let value = 0;
        if (look['@_w:firstRow'] === '1') value |= 0x0020;
        if (look['@_w:lastRow'] === '1') value |= 0x0040;
        if (look['@_w:firstColumn'] === '1') value |= 0x0080;
        if (look['@_w:lastColumn'] === '1') value |= 0x0100;
        if (look['@_w:noHBand'] === '1') value |= 0x0200;
        if (look['@_w:noVBand'] === '1') value |= 0x0400;
        table.setTblLook(value.toString(16).toUpperCase().padStart(4, '0'));
      }
    }

    // Parse table positioning (tblpPr) - for floating tables
    if (tblPrObj['w:tblpPr']) {
      const tblpPr = tblPrObj['w:tblpPr'];
      const position: any = {};

      if (tblpPr['@_w:tblpX']) position.x = parseInt(tblpPr['@_w:tblpX'], 10);
      if (tblpPr['@_w:tblpY']) position.y = parseInt(tblpPr['@_w:tblpY'], 10);
      if (tblpPr['@_w:horzAnchor']) position.horizontalAnchor = tblpPr['@_w:horzAnchor'];
      if (tblpPr['@_w:vertAnchor']) position.verticalAnchor = tblpPr['@_w:vertAnchor'];
      if (tblpPr['@_w:tblpXSpec']) position.horizontalAlignment = tblpPr['@_w:tblpXSpec'];
      if (tblpPr['@_w:tblpYSpec']) position.verticalAlignment = tblpPr['@_w:tblpYSpec'];
      if (tblpPr['@_w:leftFromText'])
        position.leftFromText = parseInt(tblpPr['@_w:leftFromText'], 10);
      if (tblpPr['@_w:rightFromText'])
        position.rightFromText = parseInt(tblpPr['@_w:rightFromText'], 10);
      if (tblpPr['@_w:topFromText']) position.topFromText = parseInt(tblpPr['@_w:topFromText'], 10);
      if (tblpPr['@_w:bottomFromText'])
        position.bottomFromText = parseInt(tblpPr['@_w:bottomFromText'], 10);

      if (Object.keys(position).length > 0) {
        table.setPosition(position);
      }
    }

    // Parse table overlap
    if (tblPrObj['w:tblOverlap']) {
      const val = tblPrObj['w:tblOverlap']['@_w:val'];
      table.setOverlap(val === 'overlap');
    }

    // Parse bidirectional visual layout
    if (tblPrObj['w:bidiVisual']) {
      table.setBidiVisual(true);
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

    // Parse table caption
    if (tblPrObj['w:tblCaption']) {
      const caption = tblPrObj['w:tblCaption']['@_w:val'];
      if (caption) table.setCaption(caption);
    }

    // Parse table description
    if (tblPrObj['w:tblDescription']) {
      const description = tblPrObj['w:tblDescription']['@_w:val'];
      if (description) table.setDescription(description);
    }

    // Parse cell spacing
    if (tblPrObj['w:tblCellSpacing']) {
      const spacing = parseInt(tblPrObj['w:tblCellSpacing']['@_w:w'] || '0', 10);
      const spacingType = tblPrObj['w:tblCellSpacing']['@_w:type'] || 'dxa';
      if (spacing > 0) {
        table.setCellSpacing(spacing);
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
    }

    // Parse table cell margins (w:tblCellMar) per ECMA-376 Part 1 §17.4.42
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
      if (cellMar['w:left']) {
        const w = cellMar['w:left']['@_w:w'];
        if (w !== undefined) margins.left = parseInt(w, 10);
      }
      if (cellMar['w:right']) {
        const w = cellMar['w:right']['@_w:w'];
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

    // Parse table borders (w:tblBorders) per ECMA-376 Part 1 §17.4.40
    if (tblPrObj['w:tblBorders']) {
      const bordersObj = tblPrObj['w:tblBorders'];
      const borders: import('../elements/Table').TableBorders = {};

      if (bordersObj['w:top']) borders.top = this.parseBorderElement(bordersObj['w:top']);
      if (bordersObj['w:bottom']) borders.bottom = this.parseBorderElement(bordersObj['w:bottom']);
      if (bordersObj['w:left']) borders.left = this.parseBorderElement(bordersObj['w:left']);
      if (bordersObj['w:right']) borders.right = this.parseBorderElement(bordersObj['w:right']);
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
        author: changeObj['@_w:author'] || '',
        date: changeObj['@_w:date'] || '',
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

    // Parse row height (w:trHeight) per ECMA-376 Part 1 §17.4.81
    if (trPrObj['w:trHeight']) {
      const heightVal = parseInt(trPrObj['w:trHeight']['@_w:val'] || '0', 10);
      const heightRule = trPrObj['w:trHeight']['@_w:hRule'] || 'atLeast';
      if (heightVal > 0) {
        row.setHeight(heightVal, heightRule);
      }
    }

    // Parse table header row (w:tblHeader) per ECMA-376 Part 1 §17.4.49
    if (trPrObj['w:tblHeader']) {
      row.setHeader(true);
    }

    // Parse can't split (w:cantSplit) per ECMA-376 Part 1 §17.4.5
    if (trPrObj['w:cantSplit']) {
      row.setCantSplit(true);
    }

    // Parse row justification (w:jc) per ECMA-376 Part 1 §17.4.79
    if (trPrObj['w:jc']) {
      const val = trPrObj['w:jc']['@_w:val'];
      if (val) {
        row.setJustification(val);
      }
    }

    // Parse hidden (w:hidden) per ECMA-376 Part 1 §17.4.23
    if (trPrObj['w:hidden']) {
      row.setHidden(true);
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

    // Parse width before (w:wBefore) per ECMA-376 Part 1 §17.4.83
    if (trPrObj['w:wBefore']) {
      const w = parseInt(trPrObj['w:wBefore']['@_w:w'] || '0', 10);
      const type = trPrObj['w:wBefore']['@_w:type'] || 'dxa';
      if (w > 0) {
        row.setWBefore(w, type);
      }
    }

    // Parse width after (w:wAfter) per ECMA-376 Part 1 §17.4.82
    if (trPrObj['w:wAfter']) {
      const w = parseInt(trPrObj['w:wAfter']['@_w:w'] || '0', 10);
      const type = trPrObj['w:wAfter']['@_w:type'] || 'dxa';
      if (w > 0) {
        row.setWAfter(w, type);
      }
    }

    // Parse row-level cell spacing (w:tblCellSpacing)
    if (trPrObj['w:tblCellSpacing']) {
      const w = parseInt(trPrObj['w:tblCellSpacing']['@_w:w'] || '0', 10);
      const type = trPrObj['w:tblCellSpacing']['@_w:type'] || 'dxa';
      if (w > 0) {
        row.setRowCellSpacing(w, type);
      }
    }

    // Parse conditional formatting (w:cnfStyle) per ECMA-376 Part 1 §17.3.1.8
    if (trPrObj['w:cnfStyle']) {
      const val = trPrObj['w:cnfStyle']['@_w:val'];
      if (val) {
        row.setCnfStyle(val);
      }
    }

    // Parse divId (w:divId) per ECMA-376 Part 1 §17.4.9
    if (trPrObj['w:divId']) {
      const val = parseInt(trPrObj['w:divId']['@_w:val'] || '0', 10);
      if (val > 0) {
        row.setDivId(val);
      }
    }

    // Parse table row property change (w:trPrChange) per ECMA-376 Part 1 §17.13.5.38
    if (trPrObj['w:trPrChange']) {
      const changeObj = trPrObj['w:trPrChange'];
      row.setTrPrChange({
        id: String(changeObj['@_w:id'] || '0'),
        author: changeObj['@_w:author'] || '',
        date: changeObj['@_w:date'] || '',
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

    // Parse table width exception (w:tblW)
    if (tblPrExObj['w:tblW']) {
      const widthVal = parseInt(tblPrExObj['w:tblW']['@_w:w'] || '0', 10);
      if (widthVal > 0) {
        exceptions.width = widthVal;
      }
    }

    // Parse table justification exception (w:jc)
    if (tblPrExObj['w:jc']) {
      const val = tblPrExObj['w:jc']['@_w:val'];
      if (val) {
        exceptions.justification = val;
      }
    }

    // Parse cell spacing exception (w:tblCellSpacing)
    if (tblPrExObj['w:tblCellSpacing']) {
      const val = parseInt(tblPrExObj['w:tblCellSpacing']['@_w:w'] || '0', 10);
      if (val > 0) {
        exceptions.cellSpacing = val;
      }
    }

    // Parse table indentation exception (w:tblInd)
    if (tblPrExObj['w:tblInd']) {
      const val = parseInt(tblPrExObj['w:tblInd']['@_w:w'] || '0', 10);
      if (val > 0) {
        exceptions.indentation = val;
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
      const borderKey = `w:${name}`;
      if (bordersObj[borderKey]) {
        const borderObj = bordersObj[borderKey];
        borders[name] = {};

        if (borderObj['@_w:val']) borders[name].style = borderObj['@_w:val'];
        if (borderObj['@_w:sz']) borders[name].size = parseInt(borderObj['@_w:sz'], 10);
        if (borderObj['@_w:space']) borders[name].space = parseInt(borderObj['@_w:space'], 10);
        if (borderObj['@_w:color']) borders[name].color = borderObj['@_w:color'];
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
        // Parse cell width (w:tcW) with type per ECMA-376 Part 1 §17.4.81
        if (tcPr['w:tcW']) {
          const widthVal = parseInt(tcPr['w:tcW']['@_w:w'] || '0', 10);
          const widthType = tcPr['w:tcW']['@_w:type'] || 'dxa';
          if (widthVal > 0 || widthType === 'auto') {
            cell.setWidthType(widthVal, widthType);
          }
        }

        // Parse conditional style (w:cnfStyle) per ECMA-376 Part 1 §17.4.7
        if (tcPr['w:cnfStyle']) {
          const cnfStyle = tcPr['w:cnfStyle']['@_w:val'];
          if (cnfStyle) {
            cell.setConditionalStyle(cnfStyle);
          }
        }

        // Parse cell borders (w:tcBorders)
        if (tcPr['w:tcBorders']) {
          const bordersObj = tcPr['w:tcBorders'];
          const borders: any = {};

          if (bordersObj['w:top']) borders.top = this.parseBorderElement(bordersObj['w:top']);
          if (bordersObj['w:bottom'])
            borders.bottom = this.parseBorderElement(bordersObj['w:bottom']);
          if (bordersObj['w:left']) borders.left = this.parseBorderElement(bordersObj['w:left']);
          if (bordersObj['w:right']) borders.right = this.parseBorderElement(bordersObj['w:right']);
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
        if (tcPr['w:tcMar']) {
          const tcMar = tcPr['w:tcMar'];
          const margins: any = {};

          if (tcMar['w:top']) {
            margins.top = parseInt(tcMar['w:top']['@_w:w'] || '0', 10);
          }
          if (tcMar['w:bottom']) {
            margins.bottom = parseInt(tcMar['w:bottom']['@_w:w'] || '0', 10);
          }
          if (tcMar['w:left']) {
            margins.left = parseInt(tcMar['w:left']['@_w:w'] || '0', 10);
          }
          if (tcMar['w:right']) {
            margins.right = parseInt(tcMar['w:right']['@_w:w'] || '0', 10);
          }

          if (Object.keys(margins).length > 0) {
            cell.setMargins(margins);
          }
        }

        // Parse vertical alignment (w:vAlign)
        if (tcPr['w:vAlign']) {
          const valign = tcPr['w:vAlign']['@_w:val'];
          if (valign && (valign === 'top' || valign === 'center' || valign === 'bottom')) {
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

        // Parse no wrap (w:noWrap) per ECMA-376 Part 1 §17.4.34
        if (tcPr['w:noWrap']) {
          cell.setNoWrap(true);
        }

        // Parse hide mark (w:hideMark) per ECMA-376 Part 1 §17.4.24
        if (tcPr['w:hideMark']) {
          cell.setHideMark(true);
        }

        // Parse headers (w:headers) per ECMA-376 Part 1 §17.4.26
        if (tcPr['w:headers']) {
          const headersVal = tcPr['w:headers']['@_w:val'];
          if (headersVal) {
            cell.setHeaders(headersVal);
          }
        }

        // Parse fit text (w:tcFitText) per ECMA-376 Part 1 §17.4.68
        if (tcPr['w:tcFitText']) {
          cell.setFitText(true);
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
          const id = parseInt(cellIns['@_w:id'] || '0', 10);
          const author = cellIns['@_w:author'] || 'Unknown';
          const dateAttr = cellIns['@_w:date'];
          const date = dateAttr ? new Date(dateAttr) : new Date();

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
          const id = parseInt(cellDel['@_w:id'] || '0', 10);
          const author = cellDel['@_w:author'] || 'Unknown';
          const dateAttr = cellDel['@_w:date'];
          const date = dateAttr ? new Date(dateAttr) : new Date();

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
          const id = parseInt(cellMerge['@_w:id'] || '0', 10);
          const author = cellMerge['@_w:author'] || 'Unknown';
          const dateAttr = cellMerge['@_w:date'];
          const date = dateAttr ? new Date(dateAttr) : new Date();
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
            author: changeObj['@_w:author'] || '',
            date: changeObj['@_w:date'] || '',
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
        // Parse ID
        const idElement = sdtPr['w:id'];
        if (idElement?.['@_w:val']) {
          properties.id = parseInt(idElement['@_w:val'], 10);
        }

        // Parse tag
        const tagElement = sdtPr['w:tag'];
        if (tagElement?.['@_w:val']) {
          properties.tag = tagElement['@_w:val'];
        }

        // Parse lock
        const lockElement = sdtPr['w:lock'];
        if (lockElement?.['@_w:val']) {
          properties.lock = lockElement['@_w:val'];
        }

        // Parse alias
        const aliasElement = sdtPr['w:alias'];
        if (aliasElement?.['@_w:val']) {
          properties.alias = aliasElement['@_w:val'];
        }

        // Parse control type from various elements
        if (sdtPr['w:richText']) {
          properties.controlType = 'richText';
        } else if (sdtPr['w:text']) {
          properties.controlType = 'plainText';
          const textElement = sdtPr['w:text'];
          properties.plainText = {
            multiLine:
              textElement?.['@_w:multiLine'] === '1' || textElement?.['@_w:multiLine'] === 'true',
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
          // Handle both string and numeric values from XML parser
          const checkedVal = checkboxElement?.['w14:checked']?.['@_w14:val'];
          properties.checkbox = {
            checked:
              checkedVal === 1 ||
              checkedVal === '1' ||
              checkedVal === true ||
              checkedVal === 'true',
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

        // Parse showing placeholder flag (w:showingPlcHdr)
        const showingPlcHdr = sdtPr['w:showingPlcHdr'];
        if (showingPlcHdr) {
          const val = showingPlcHdr['@_w:val'];
          properties.showingPlcHdr = val === '1' || val === 'true' || val === true;
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
    properties: any,
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
   * Helper to parse list items for combo box / dropdown
   */
  private parseListItems(element: any): any {
    const items: any[] = [];
    const listItems = element?.['w:listItem'];
    const itemArray = Array.isArray(listItems) ? listItems : listItems ? [listItems] : [];

    for (const item of itemArray) {
      if (item['@_w:displayText'] && item['@_w:value']) {
        items.push({
          displayText: item['@_w:displayText'],
          value: item['@_w:value'],
        });
      }
    }

    return {
      items,
      lastValue: element?.['@_w:lastValue'],
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

          if (width && height) {
            sectionProps.pageSize = {
              width: parseInt(width, 10),
              height: parseInt(height, 10),
              orientation: orient === 'landscape' ? 'landscape' : 'portrait',
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
            const shadow = XMLParser.extractAttribute(sideXml, 'w:shadow');
            if (shadow === '1' || shadow === 'true') border.shadow = true;
            const frame = XMLParser.extractAttribute(sideXml, 'w:frame');
            if (frame === '1' || frame === 'true') border.frame = true;
            const themeColor = XMLParser.extractAttribute(sideXml, 'w:themeColor');
            if (themeColor) border.themeColor = themeColor;
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

      // Parse columns (enhanced with separator and custom widths)
      const colsElements = XMLParser.extractElements(sectPr, 'w:cols');
      if (colsElements.length > 0) {
        const cols = colsElements[0];
        if (cols) {
          const num = XMLParser.extractAttribute(cols, 'w:num');
          const space = XMLParser.extractAttribute(cols, 'w:space');
          const equalWidth = XMLParser.extractAttribute(cols, 'w:equalWidth');
          const sep = XMLParser.extractAttribute(cols, 'w:sep');

          // Extract individual column widths
          const colElements = XMLParser.extractElements(cols, 'w:col');
          const columnWidths: number[] = [];
          for (const col of colElements) {
            const width = XMLParser.extractAttribute(col, 'w:w');
            if (width) {
              columnWidths.push(parseInt(width.toString(), 10));
            }
          }

          // Helper to handle boolean conversion (XMLParser may return string or number)
          const toBool = (val: any) => val === '1' || val === 1 || val === 'true' || val === true;

          if (num) {
            sectionProps.columns = {
              count: parseInt(num.toString(), 10),
              space: space ? parseInt(space.toString(), 10) : undefined,
              equalWidth: equalWidth ? toBool(equalWidth) : undefined,
              separator: sep ? toBool(sep) : undefined,
              columnWidths: columnWidths.length > 0 ? columnWidths : undefined,
            };
          }
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

      // Parse title page flag
      if (XMLParser.hasSelfClosingTag(sectPr, 'w:titlePg')) {
        sectionProps.titlePage = true;
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

      // Parse bidi (right-to-left section layout)
      if (XMLParser.hasSelfClosingTag(sectPr, 'w:bidi')) {
        sectionProps.bidi = true;
      }

      // Parse RTL gutter
      if (XMLParser.hasSelfClosingTag(sectPr, 'w:rtlGutter')) {
        sectionProps.rtlGutter = true;
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

      // Parse noEndnote
      if (XMLParser.hasSelfClosingTag(sectPr, 'w:noEndnote')) {
        sectionProps.noEndnote = true;
      }

      // Parse form protection
      if (XMLParser.hasSelfClosingTag(sectPr, 'w:formProt')) {
        sectionProps.formProt = true;
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

    // Parse metadata properties (Phase 5.3)
    // qFormat - Quick style gallery
    const qFormat = styleXml.includes('<w:qFormat/>') || styleXml.includes('<w:qFormat ');

    // semiHidden - Hide from recommended list
    const semiHidden = styleXml.includes('<w:semiHidden/>') || styleXml.includes('<w:semiHidden ');

    // unhideWhenUsed - Auto-show when applied
    const unhideWhenUsed =
      styleXml.includes('<w:unhideWhenUsed/>') || styleXml.includes('<w:unhideWhenUsed ');

    // locked - Prevent modification
    const locked = styleXml.includes('<w:locked/>') || styleXml.includes('<w:locked ');

    // personal - User-specific style
    const personal = styleXml.includes('<w:personal/>') || styleXml.includes('<w:personal ');

    // personalCompose - Style for composing new messages
    const personalCompose =
      styleXml.includes('<w:personalCompose/>') || styleXml.includes('<w:personalCompose ');

    // personalReply - Style for replying to messages
    const personalReply =
      styleXml.includes('<w:personalReply/>') || styleXml.includes('<w:personalReply ');

    // autoRedefine - Update style from formatting
    const autoRedefine =
      styleXml.includes('<w:autoRedefine/>') || styleXml.includes('<w:autoRedefine ');

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
      isDefault: defaultAttr === '1' || defaultAttr === 'true',
      customStyle: customStyleAttr === '1' || customStyleAttr === 'true',
      paragraphFormatting,
      numPr: styleNumPr,
      runFormatting,
      tableStyle,
      // Metadata properties (Phase 5.3)
      qFormat: qFormat || undefined,
      semiHidden: semiHidden || undefined,
      unhideWhenUsed: unhideWhenUsed || undefined,
      locked: locked || undefined,
      personal: personal || undefined,
      personalCompose: personalCompose || undefined,
      personalReply: personalReply || undefined,
      autoRedefine: autoRedefine || undefined,
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

    // Parse alignment (w:jc)
    const jcElement = XMLParser.extractSelfClosingTag(pPrXml, 'w:jc');
    if (jcElement) {
      const alignment = XMLParser.extractAttribute(`<w:jc${jcElement}`, 'w:val');
      if (alignment) {
        formatting.alignment = alignment as 'left' | 'center' | 'right' | 'justify';
      }
    }

    // Parse spacing (w:spacing)
    const spacingElement = XMLParser.extractSelfClosingTag(pPrXml, 'w:spacing');
    if (spacingElement) {
      const before = XMLParser.extractAttribute(`<w:spacing${spacingElement}`, 'w:before');
      const after = XMLParser.extractAttribute(`<w:spacing${spacingElement}`, 'w:after');
      const line = XMLParser.extractAttribute(`<w:spacing${spacingElement}`, 'w:line');
      const lineRule = XMLParser.extractAttribute(`<w:spacing${spacingElement}`, 'w:lineRule');

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
      };
    }

    // Parse indentation (w:ind)
    const indElement = XMLParser.extractSelfClosingTag(pPrXml, 'w:ind');
    if (indElement) {
      const left = XMLParser.extractAttribute(`<w:ind${indElement}`, 'w:left');
      const right = XMLParser.extractAttribute(`<w:ind${indElement}`, 'w:right');
      const firstLine = XMLParser.extractAttribute(`<w:ind${indElement}`, 'w:firstLine');
      const hanging = XMLParser.extractAttribute(`<w:ind${indElement}`, 'w:hanging');

      formatting.indentation = {
        left: left ? parseInt(left, 10) : undefined,
        right: right ? parseInt(right, 10) : undefined,
        firstLine: firstLine ? parseInt(firstLine, 10) : undefined,
        hanging: hanging ? parseInt(hanging, 10) : undefined,
      };
    }

    // Parse boolean properties
    if (pPrXml.includes('<w:keepNext/>') || pPrXml.includes('<w:keepNext ')) {
      formatting.keepNext = true;
    }
    if (pPrXml.includes('<w:keepLines/>') || pPrXml.includes('<w:keepLines ')) {
      formatting.keepLines = true;
    }
    if (pPrXml.includes('<w:pageBreakBefore/>') || pPrXml.includes('<w:pageBreakBefore ')) {
      formatting.pageBreakBefore = true;
    }

    // Contextual spacing per ECMA-376 Part 1 §17.3.1.8
    // "Don't add space between paragraphs of the same style"
    if (pPrXml.includes('<w:contextualSpacing/>') || pPrXml.includes('<w:contextualSpacing ')) {
      formatting.contextualSpacing = true;
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

    // Parse boolean properties
    if (rPrXml.includes('<w:b/>') || rPrXml.includes('<w:b ')) {
      formatting.bold = true;
    }
    if (rPrXml.includes('<w:i/>') || rPrXml.includes('<w:i ')) {
      formatting.italic = true;
    }
    if (rPrXml.includes('<w:strike/>') || rPrXml.includes('<w:strike ')) {
      formatting.strike = true;
    }
    if (rPrXml.includes('<w:smallCaps/>') || rPrXml.includes('<w:smallCaps ')) {
      formatting.smallCaps = true;
    }
    if (rPrXml.includes('<w:caps/>') || rPrXml.includes('<w:caps ')) {
      formatting.allCaps = true;
    }

    // Parse underline - use extractSelfClosingTag for accuracy
    const uElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:u');
    if (uElement) {
      const uVal = XMLParser.extractAttribute(`<w:u${uElement}`, 'w:val');
      if (
        uVal === 'single' ||
        uVal === 'double' ||
        uVal === 'thick' ||
        uVal === 'dotted' ||
        uVal === 'dash' ||
        uVal === 'none'
      ) {
        formatting.underline = uVal;
      } else {
        formatting.underline = true;
      }
    }

    // Parse subscript/superscript - use extractSelfClosingTag
    const vertAlignElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:vertAlign');
    if (vertAlignElement) {
      const val = XMLParser.extractAttribute(`<w:vertAlign${vertAlignElement}`, 'w:val');
      if (val === 'subscript') {
        formatting.subscript = true;
      } else if (val === 'superscript') {
        formatting.superscript = true;
      }
    }

    // Parse font (w:rFonts) - use extractSelfClosingTag
    const rFontsElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:rFonts');
    if (rFontsElement) {
      const ascii = XMLParser.extractAttribute(`<w:rFonts${rFontsElement}`, 'w:ascii');
      if (ascii) {
        formatting.font = ascii;
      }
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

    // Parse color (w:color)
    // Use extractSelfClosingTag to avoid matching other tags
    const colorElement = XMLParser.extractSelfClosingTag(rPrXml, 'w:color');
    if (colorElement) {
      const val = XMLParser.extractAttribute(`<w:color${colorElement}`, 'w:val');
      if (val && val !== 'auto') {
        formatting.color = val;
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
            | 'white';
        }
      }
    }

    // Parse shading (w:shd) per ECMA-376 Part 1 §17.3.2.32
    const shading = this.parseShadingFromXml(rPrXml);
    if (shading) {
      formatting.shading = shading;
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

    // Parse indent
    if (tblPrXml.includes('<w:tblInd')) {
      const tag = XMLParser.extractSelfClosingTag(tblPrXml, 'w:tblInd');
      if (tag) {
        const w = XMLParser.extractAttribute(`<w:tblInd${tag}`, 'w:w');
        if (w) {
          formatting.indent = parseInt(w, 10);
        }
      }
    }

    // Parse alignment
    if (tblPrXml.includes('<w:jc')) {
      const tag = XMLParser.extractSelfClosingTag(tblPrXml, 'w:jc');
      if (tag) {
        const val = XMLParser.extractAttribute(`<w:jc${tag}`, 'w:val');
        if (val === 'left' || val === 'center' || val === 'right') {
          formatting.alignment = val;
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

    // Parse vertical alignment
    if (tcPrXml.includes('<w:vAlign')) {
      const tag = XMLParser.extractSelfClosingTag(tcPrXml, 'w:vAlign');
      if (tag) {
        const val = XMLParser.extractAttribute(`<w:vAlign${tag}`, 'w:val');
        if (val === 'top' || val === 'center' || val === 'bottom') {
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

    // Parse cantSplit
    if (trPrXml.includes('<w:cantSplit/>') || trPrXml.includes('<w:cantSplit ')) {
      formatting.cantSplit = true;
    }

    // Parse tblHeader (isHeader)
    if (trPrXml.includes('<w:tblHeader/>') || trPrXml.includes('<w:tblHeader ')) {
      formatting.isHeader = true;
    }

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

    const borderTypes = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];
    for (const type of borderTypes) {
      if (bordersXml.includes(`<w:${type}`)) {
        const tag = XMLParser.extractSelfClosingTag(bordersXml, `w:${type}`);
        if (tag) {
          const style = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:val');
          const size = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:sz');
          const space = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:space');
          const color = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:color');

          const border: import('../formatting/Style').BorderProperties = {};
          if (style) border.style = style as any;
          if (size) border.size = parseInt(size, 10);
          if (space) border.space = parseInt(space, 10);
          if (color) border.color = color;

          if (Object.keys(border).length > 0) {
            borders[type] = border;
          }
        }
      }
    }

    // Add diagonal borders for cells
    if (includeDiagonals) {
      const diagonalTypes = ['tl2br', 'tr2bl'];
      for (const type of diagonalTypes) {
        if (bordersXml.includes(`<w:${type}`)) {
          const tag = XMLParser.extractSelfClosingTag(bordersXml, `w:${type}`);
          if (tag) {
            const style = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:val');
            const size = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:sz');
            const space = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:space');
            const color = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:color');

            const border: import('../formatting/Style').BorderProperties = {};
            if (style) border.style = style as any;
            if (size) border.size = parseInt(size, 10);
            if (space) border.space = parseInt(space, 10);
            if (color) border.color = color;

            if (Object.keys(border).length > 0) {
              borders[type] = border;
            }
          }
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
    const shading: ShadingConfig = {};
    if (shd['@_w:val']) shading.pattern = shd['@_w:val'];
    if (shd['@_w:fill']) shading.fill = shd['@_w:fill'];
    if (shd['@_w:color']) shading.color = shd['@_w:color'];
    if (shd['@_w:themeFill']) shading.themeFill = shd['@_w:themeFill'];
    if (shd['@_w:themeColor']) shading.themeColor = shd['@_w:themeColor'];
    if (shd['@_w:themeFillTint']) shading.themeFillTint = shd['@_w:themeFillTint'];
    if (shd['@_w:themeFillShade']) shading.themeFillShade = shd['@_w:themeFillShade'];
    if (shd['@_w:themeTint']) shading.themeTint = shd['@_w:themeTint'];
    if (shd['@_w:themeShade']) shading.themeShade = shd['@_w:themeShade'];
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

    const marginTypes = ['top', 'bottom', 'left', 'right'];
    for (const type of marginTypes) {
      if (marginXml.includes(`<w:${type}`)) {
        const tag = XMLParser.extractSelfClosingTag(marginXml, `w:${type}`);
        if (tag) {
          const w = XMLParser.extractAttribute(`<w:${type}${tag}`, 'w:w');
          if (w) {
            margins[type as keyof import('../formatting/Style').CellMargins] = parseInt(w, 10);
          }
        }
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
      result.style = propsObj['w:tblStyle']['@_w:val'] || '';
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
    }
    if (propsObj['w:tblCellSpacing']) {
      result.cellSpacing = parseInt(propsObj['w:tblCellSpacing']['@_w:w'] || '0', 10);
    }
    if (propsObj['w:tblBorders']) {
      const borders: any = {};
      const bordersObj = propsObj['w:tblBorders'];
      for (const side of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']) {
        if (bordersObj[`w:${side}`]) {
          borders[side] = this.parseBorderElement(bordersObj[`w:${side}`]);
        }
      }
      if (Object.keys(borders).length > 0) result.borders = borders;
    }

    // Row-level properties (w:trPr context)
    if (propsObj['w:trHeight']) {
      result.height = parseInt(propsObj['w:trHeight']['@_w:val'] || '0', 10);
      const rule = propsObj['w:trHeight']['@_w:hRule'];
      if (rule) result.heightRule = rule;
    }
    if (propsObj['w:tblHeader']) {
      result.isHeader = true;
    }
    if (propsObj['w:cantSplit']) {
      result.cantSplit = true;
    }
    if (propsObj['w:hidden']) {
      result.hidden = true;
    }

    // Cell-level properties (w:tcPr context)
    if (propsObj['w:tcW']) {
      result.width = parseInt(propsObj['w:tcW']['@_w:w'] || '0', 10);
      result.widthType = propsObj['w:tcW']['@_w:type'] || 'dxa';
    }
    if (propsObj['w:vAlign']) {
      result.verticalAlignment = propsObj['w:vAlign']['@_w:val'];
    }
    if (propsObj['w:tcBorders']) {
      const borders: any = {};
      const bordersObj = propsObj['w:tcBorders'];
      for (const side of ['top', 'bottom', 'left', 'right', 'tl2br', 'tr2bl']) {
        if (bordersObj[`w:${side}`]) {
          borders[side] = this.parseBorderElement(bordersObj[`w:${side}`]);
        }
      }
      if (Object.keys(borders).length > 0) result.borders = borders;
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

    // Page size
    const pgSzElements = XMLParser.extractElements(sectPrXml, 'w:pgSz');
    if (pgSzElements.length > 0 && pgSzElements[0]) {
      const pgSz = pgSzElements[0];
      const width = XMLParser.extractAttribute(pgSz, 'w:w');
      const height = XMLParser.extractAttribute(pgSz, 'w:h');
      const orient = XMLParser.extractAttribute(pgSz, 'w:orient');
      if (width || height) {
        result.pageSize = {
          width: width ? parseInt(width, 10) : undefined,
          height: height ? parseInt(height, 10) : undefined,
          orientation: orient === 'landscape' ? 'landscape' : 'portrait',
        };
      }
    }

    // Margins
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
      if (Object.keys(margins).length > 0) result.margins = margins;
    }

    // Section type
    const typeElements = XMLParser.extractElements(sectPrXml, 'w:type');
    if (typeElements.length > 0 && typeElements[0]) {
      const val = XMLParser.extractAttribute(typeElements[0], 'w:val');
      if (val) result.type = val;
    }

    // Columns
    const colsElements = XMLParser.extractElements(sectPrXml, 'w:cols');
    if (colsElements.length > 0 && colsElements[0]) {
      const cols = colsElements[0];
      const num = XMLParser.extractAttribute(cols, 'w:num');
      const space = XMLParser.extractAttribute(cols, 'w:space');
      if (num) {
        result.columns = {
          count: parseInt(num, 10),
          space: space ? parseInt(space, 10) : undefined,
        };
      }
    }

    return result;
  }
}
