/**
 * DocumentParser - Handles parsing of DOCX files
 * Extracts content from ZIP archives and converts XML to structured data
 */

import { ComplexField, Field } from "../elements/Field";
import { Hyperlink } from "../elements/Hyperlink";
import { ImageManager } from "../elements/ImageManager";
import { ImageRun } from "../elements/ImageRun";
import { Paragraph, ParagraphFormatting, ParagraphContent } from "../elements/Paragraph";
import { Revision } from "../elements/Revision";
import { BreakType, Run, RunContent, RunFormatting } from "../elements/Run";
import {
  PageNumberFormat,
  Section,
  SectionProperties,
  SectionType,
} from "../elements/Section";
import { StructuredDocumentTag } from "../elements/StructuredDocumentTag";
import { Table } from "../elements/Table";
import { TableCell } from "../elements/TableCell";
import { TableOfContents } from "../elements/TableOfContents";
import { TableOfContentsElement } from "../elements/TableOfContentsElement";
import { TableRow } from "../elements/TableRow";
import { AbstractNumbering } from "../formatting/AbstractNumbering";
import { NumberingInstance } from "../formatting/NumberingInstance";
import { Style, StyleProperties, StyleType } from "../formatting/Style";
import {
  logParagraphContent,
  logParsing,
  logTextDirection,
} from "../utils/diagnostics";
import { getGlobalLogger, createScopedLogger, ILogger, defaultLogger } from "../utils/logger";

// Create scoped logger for DocumentParser operations
function getLogger(): ILogger {
  return createScopedLogger(getGlobalLogger(), 'DocumentParser');
}
import { XMLBuilder } from "../xml/XMLBuilder";
import { XMLParser } from "../xml/XMLParser";
import { ZipHandler } from "../zip/ZipHandler";
import { DOCX_PATHS } from "../zip/types";
import { DocumentProperties } from "./Document";
import { RelationshipManager } from "./RelationshipManager";

/**
 * Parse error tracking
 */
export interface ParseError {
  element: string;
  error: Error;
}

/**
 * Body element types
 */
type BodyElement =
  | Paragraph
  | Table
  | TableOfContentsElement
  | StructuredDocumentTag;

/**
 * DocumentParser handles all document parsing logic
 */
export class DocumentParser {
  private parseErrors: ParseError[] = [];
  private strictParsing: boolean;

  constructor(strictParsing: boolean = false) {
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
    imageManager: ImageManager
  ): Promise<{
    bodyElements: BodyElement[];
    properties: DocumentProperties;
    relationshipManager: RelationshipManager;
    styles: Style[];
    abstractNumberings: AbstractNumbering[];
    numberingInstances: NumberingInstance[];
    section: Section | null;
    namespaces: Record<string, string>;
  }> {
    const logger = getLogger();
    logger.info('Parsing document');

    // Verify the document exists
    const docXml = zipHandler.getFileAsString(DOCX_PATHS.DOCUMENT);
    if (!docXml) {
      logger.error('Invalid document: word/document.xml not found');
      throw new Error("Invalid document: word/document.xml not found");
    }

    logger.info('Parsing document.xml', { xmlSize: docXml.length });

    // Parse existing relationships to avoid ID collisions
    const parsedRelationshipManager = this.parseRelationships(
      zipHandler,
      relationshipManager
    );

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
      instanceCount: numbering.numberingInstances.length
    });

    // Parse section properties from document.xml
    const section = this.parseSectionProperties(docXml);

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
      totalElements: bodyElements.length
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
    while (pos < bodyContent.length) {
      const nextP = this.findNextTopLevelTag(bodyContent, "w:p", pos);
      const nextTbl = this.findNextTopLevelTag(bodyContent, "w:tbl", pos);
      const nextSdt = this.findNextTopLevelTag(bodyContent, "w:sdt", pos);

      const candidates = [];
      if (nextP !== -1) candidates.push({ type: "p", pos: nextP });
      if (nextTbl !== -1) candidates.push({ type: "tbl", pos: nextTbl });
      if (nextSdt !== -1) candidates.push({ type: "sdt", pos: nextSdt });

      if (candidates.length === 0) break;

      candidates.sort((a, b) => a.pos - b.pos);
      const next = candidates[0];

      if (next) {
        if (next.type === "p") {
          const elementXml = this.extractSingleElement(
            bodyContent,
            "w:p",
            next.pos
          );

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
        } else if (next.type === "tbl") {
          const elementXml = this.extractSingleElement(
            bodyContent,
            "w:tbl",
            next.pos
          );
          if (elementXml) {
            // Parse XML to object, then extract the table content
            // IMPORTANT: trimValues: false preserves whitespace from xml:space="preserve" attributes
            const parsed = XMLParser.parseToObject(elementXml, {
              trimValues: false,
            });
            const table = await this.parseTableFromObject(
              parsed["w:tbl"],
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
        } else if (next.type === "sdt") {
          const elementXml = this.extractSingleElement(
            bodyContent,
            "w:sdt",
            next.pos
          );
          if (elementXml) {
            // Parse XML to object, then extract the SDT content
            // IMPORTANT: trimValues: false preserves whitespace from xml:space="preserve" attributes
            const parsed = XMLParser.parseToObject(elementXml, {
              trimValues: false,
            });
            const sdt = await this.parseSDTFromObject(
              parsed["w:sdt"],
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (sdt) bodyElements.push(sdt);
            pos = next.pos + elementXml.length;
          } else {
            pos = next.pos + 1;
          }
        }
      }
    }

    // Validate that we didn't load an empty/corrupted document
    this.validateLoadedContent(bodyElements);

    return bodyElements;
  }

  /**
   * Finds the next occurrence of a tag in the content
   * Returns the position of the opening '<' or -1 if not found
   */
  private findNextTag(
    content: string,
    tagName: string,
    startPos: number
  ): number {
    const tag = `<${tagName}`;
    let pos = content.indexOf(tag, startPos);

    while (pos !== -1) {
      // Verify this is the exact tag (not a prefix match like <w:p matching <w:pPr>)
      // The character after the tag name must be either '>', '/' or whitespace
      const charAfterTag = content[pos + tag.length];
      if (
        charAfterTag &&
        charAfterTag !== ">" &&
        charAfterTag !== "/" &&
        charAfterTag !== " " &&
        charAfterTag !== "\t" &&
        charAfterTag !== "\n" &&
        charAfterTag !== "\r"
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
  private findNextTopLevelTag(
    content: string,
    tagName: string,
    startPos: number
  ): number {
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
   * Checks if a position in the content is inside a table element
   * Returns true if there's an unclosed <w:tbl> before this position
   */
  private isPositionInsideTable(content: string, position: number): boolean {
    // Look backwards from position to find the nearest table-related tag
    const beforeContent = content.substring(0, position);

    // Find all <w:tbl> and </w:tbl> tags before this position
    const openTableTags = (beforeContent.match(/<w:tbl[\s>]/g) || []).length;
    const closeTableTags = (beforeContent.match(/<\/w:tbl>/g) || []).length;

    // If there are more open tags than close tags, we're inside a table
    return openTableTags > closeTableTags;
  }

  /**
   * Extracts a single element from the content starting at the given position
   * Returns the complete element XML including opening and closing tags
   *
   * FIX (v1.3.1): Uses XMLParser.extractElements to ensure consistent extraction
   * behavior and prevent loss of self-closing elements like <w:tab/>
   * This fixes the TOC tab preservation bug where tabs were lost during extraction.
   */
  private extractSingleElement(
    content: string,
    tagName: string,
    startPos: number
  ): string {
    // Extract the substring starting from the position
    const remainingContent = content.substring(startPos);

    // Use XMLParser.extractElements to get all elements of this type
    // This ensures we use the same proven extraction logic throughout
    const elements = XMLParser.extractElements(remainingContent, tagName);

    // Return the first element (which starts at position 0 in remainingContent)
    // This is the element at the specified startPos in the original content
    const extracted = elements.length > 0 ? elements[0]! : "";

    return extracted;
  }

  /**
   * Validates loaded content to detect corrupted or empty documents
   * Adds warnings if the document appears to have lost text content
   */
  private validateLoadedContent(bodyElements: BodyElement[]): void {
    const paragraphs = bodyElements.filter(
      (el): el is Paragraph => el instanceof Paragraph
    );

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
          element: "document-validation",
          error: warning,
        });

        // Always warn to console, even in non-strict mode
        defaultLogger.warn(`\nDocXML Load Warning:\n${warning.message}\n`);
      } else if (emptyPercentage > 50 && emptyRuns > 5) {
        const warning = new Error(
          `Document has ${emptyRuns} out of ${totalRuns} runs (${emptyPercentage.toFixed(
            1
          )}%) with no text. ` +
            `This is higher than normal and may indicate partial data loss.`
        );
        this.parseErrors.push({
          element: "document-validation",
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
      const pElement = paraObj["w:p"] as any;
      if (!pElement) {
        return null;
      }

      // Parse paragraph properties
      this.parseParagraphPropertiesFromObject(pElement["w:pPr"], paragraph);

      // Parse w14:paraId if present
      const paraId = pElement["w14:paraId"];
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

    // NEW: Assemble complex fields from run tokens
    // DISABLED: ComplexField assembly causes parsing failures for multi-paragraph fields (e.g., TOC)
    // TODO: Re-enable when multi-paragraph field support is added
    // this.assembleComplexFields(paragraph);

    // Diagnostic logging for paragraph
      const runs = paragraph.getRuns();
      const runData = runs.map((run) => ({
        text: run.getText(),
        rtl: run.getFormatting().rtl,
      }));
      const bidi = paragraph.getFormatting().bidi;

      logParagraphContent("parsing", -1, runData, bidi);

      if (bidi) {
        logTextDirection(`Paragraph has BiDi enabled`);
      }

      // Merge consecutive hyperlinks with the same URL (handles Google Docs fragmentation)
      this.mergeConsecutiveHyperlinks(paragraph);

      return paragraph;
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: "paragraph", error: err });

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
    const pPrEnd = paraXml.indexOf("</w:pPr>");
    const contentStart = pPrEnd !== -1 ? pPrEnd + 8 : paraXml.indexOf(">") + 1;
    const contentEnd = paraXml.lastIndexOf("</w:p>");

    if (contentEnd <= contentStart) {
      return; // Empty paragraph
    }

    const paraContent = paraXml.substring(contentStart, contentEnd);

    // Track children by scanning XML for opening tags
    interface ChildMarker {
      type: "w:r" | "w:hyperlink" | "w:fldSimple" | "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo";
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

    // Helper to find closing tag position for a given tag name starting from position
    const findClosingTagEnd = (content: string, tagName: string, startPos: number): number => {
      const closingTag = `</${tagName}>`;
      const closingPos = content.indexOf(closingTag, startPos);
      if (closingPos === -1) return startPos; // Fallback if not found
      return closingPos + closingTag.length;
    };

    // Helper to check if tag is self-closing
    const isSelfClosing = (tagContent: string): boolean => {
      return tagContent.endsWith("/");
    };

    // Scan for all first-level child elements in document order
    let searchPos = 0;
    while (searchPos < paraContent.length) {
      // Find the next opening tag
      const tagStart = paraContent.indexOf("<", searchPos);
      if (tagStart === -1) break;

      // Extract tag name
      const tagEnd = paraContent.indexOf(">", tagStart);
      if (tagEnd === -1) break;

      const tagContent = paraContent.substring(tagStart + 1, tagEnd);
      const tagName = tagContent.split(/[\s\/>]/)[0];
      const selfClosing = isSelfClosing(tagContent);

      if (tagName === "w:r") {
        children.push({ type: "w:r", pos: tagStart, index: runIndex++ });
        // Skip past closing tag to avoid counting nested elements
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, "w:r", tagEnd);
      } else if (tagName === "w:hyperlink") {
        children.push({
          type: "w:hyperlink",
          pos: tagStart,
          index: hyperlinkIndex++,
        });
        // Skip past closing tag - hyperlinks contain nested runs we shouldn't count
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, "w:hyperlink", tagEnd);
      } else if (tagName === "w:fldSimple") {
        children.push({
          type: "w:fldSimple",
          pos: tagStart,
          index: fieldIndex++,
        });
        // Skip past closing tag
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, "w:fldSimple", tagEnd);
      } else if (tagName === "w:ins") {
        children.push({
          type: "w:ins",
          pos: tagStart,
          index: insIndex++,
        });
        // Skip past closing tag - ins contains nested runs
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, "w:ins", tagEnd);
      } else if (tagName === "w:del") {
        children.push({
          type: "w:del",
          pos: tagStart,
          index: delIndex++,
        });
        // Skip past closing tag - del contains nested runs
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, "w:del", tagEnd);
      } else if (tagName === "w:moveFrom") {
        children.push({
          type: "w:moveFrom",
          pos: tagStart,
          index: moveFromIndex++,
        });
        // Skip past closing tag
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, "w:moveFrom", tagEnd);
      } else if (tagName === "w:moveTo") {
        children.push({
          type: "w:moveTo",
          pos: tagStart,
          index: moveToIndex++,
        });
        // Skip past closing tag
        searchPos = selfClosing ? tagEnd + 1 : findClosingTagEnd(paraContent, "w:moveTo", tagEnd);
      } else {
        searchPos = tagEnd + 1;
      }
    }

    // Extract revision XMLs from paragraph content for raw XML parsing
    const insXmls = XMLParser.extractElements(paraContent, "w:ins");
    const delXmls = XMLParser.extractElements(paraContent, "w:del");
    const moveFromXmls = XMLParser.extractElements(paraContent, "w:moveFrom");
    const moveToXmls = XMLParser.extractElements(paraContent, "w:moveTo");

    // Now process children in the order they were found
    for (const child of children) {
      if (child.type === "w:r") {
        const runs = pElement["w:r"];
        const runArray = Array.isArray(runs) ? runs : runs ? [runs] : [];
        if (child.index < runArray.length) {
          const runObj = runArray[child.index];
          if (runObj["w:drawing"]) {
            if (zipHandler && imageManager) {
              const imageRun = await this.parseDrawingFromObject(
                runObj["w:drawing"],
                zipHandler,
                relationshipManager,
                imageManager
              );
              if (imageRun) {
                paragraph.addRun(imageRun);
              }
            }
          } else {
            const run = this.parseRunFromObject(runObj);
            if (run) {
              paragraph.addRun(run);
            }
          }
        }
      } else if (child.type === "w:hyperlink") {
        const hyperlinks = pElement["w:hyperlink"];
        const hyperlinkArray = Array.isArray(hyperlinks)
          ? hyperlinks
          : hyperlinks
          ? [hyperlinks]
          : [];
        if (child.index < hyperlinkArray.length) {
          const hyperlink = this.parseHyperlinkFromObject(
            hyperlinkArray[child.index],
            relationshipManager
          );
          if (hyperlink) {
            paragraph.addHyperlink(hyperlink);
          }
        }
      } else if (child.type === "w:fldSimple") {
        const fields = pElement["w:fldSimple"];
        const fieldArray = Array.isArray(fields)
          ? fields
          : fields
          ? [fields]
          : [];
        if (child.index < fieldArray.length) {
          const field = this.parseSimpleFieldFromObject(
            fieldArray[child.index]
          );
          if (field) {
            paragraph.addField(field);
          }
        }
      } else if (child.type === "w:ins") {
        if (child.index < insXmls.length) {
          const revisionXml = insXmls[child.index];
          if (revisionXml) {
            const revision = await this.parseRevisionFromXml(
              revisionXml,
              "w:ins",
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revision) {
              paragraph.addRevision(revision);
            }
          }
        }
      } else if (child.type === "w:del") {
        if (child.index < delXmls.length) {
          const revisionXml = delXmls[child.index];
          if (revisionXml) {
            const revision = await this.parseRevisionFromXml(
              revisionXml,
              "w:del",
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revision) {
              paragraph.addRevision(revision);
            }
          }
        }
      } else if (child.type === "w:moveFrom") {
        if (child.index < moveFromXmls.length) {
          const revisionXml = moveFromXmls[child.index];
          if (revisionXml) {
            const revision = await this.parseRevisionFromXml(
              revisionXml,
              "w:moveFrom",
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revision) {
              paragraph.addRevision(revision);
            }
          }
        }
      } else if (child.type === "w:moveTo") {
        if (child.index < moveToXmls.length) {
          const revisionXml = moveToXmls[child.index];
          if (revisionXml) {
            const revision = await this.parseRevisionFromXml(
              revisionXml,
              "w:moveTo",
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (revision) {
              paragraph.addRevision(revision);
            }
          }
        }
      }
    }
  }

  /**
   * Parses a revision element from raw XML
   * Extracts revision metadata (id, author, date) and contained runs
   * @param revisionXml - Raw XML of the revision element
   * @param tagName - Tag name (w:ins, w:del, w:moveFrom, w:moveTo)
   * @param relationshipManager - Relationship manager for image relationships
   * @param zipHandler - ZIP handler for loading image data
   * @param imageManager - Image manager for registering images
   * @returns Parsed Revision instance or null
   */
  private async parseRevisionFromXml(
    revisionXml: string,
    tagName: string,
    relationshipManager: RelationshipManager,
    zipHandler?: ZipHandler,
    imageManager?: ImageManager
  ): Promise<Revision | null> {
    try {
      // Map XML tag to RevisionType
      let revisionType: import("../elements/Revision").RevisionType;
      switch (tagName) {
        case "w:ins":
          revisionType = "insert";
          break;
        case "w:del":
          revisionType = "delete";
          break;
        case "w:moveFrom":
          revisionType = "moveFrom";
          break;
        case "w:moveTo":
          revisionType = "moveTo";
          break;
        default:
          return null;
      }

      // Extract attributes
      const idAttr = XMLParser.extractAttribute(revisionXml, "w:id");
      const author = XMLParser.extractAttribute(revisionXml, "w:author");
      const dateAttr = XMLParser.extractAttribute(revisionXml, "w:date");
      const moveId = XMLParser.extractAttribute(revisionXml, "w:moveId");

      if (!idAttr || !author) {
        return null; // Required attributes missing
      }

      const id = parseInt(idAttr, 10);
      const date = dateAttr ? new Date(dateAttr) : new Date();

      // Extract content from revision element (runs and hyperlinks)
      const runXmls = XMLParser.extractElements(revisionXml, "w:r");
      const hyperlinkXmls = XMLParser.extractElements(revisionXml, "w:hyperlink");

      // Use RevisionContent to hold both Run and Hyperlink objects
      const content: import('../elements/RevisionContent').RevisionContent[] = [];

      // Parse runs
      for (const runXml of runXmls) {
        // Parse the run object
        const runObj = XMLParser.parseToObject(runXml, { trimValues: false });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const runElement = runObj["w:r"] as any;

        // Check if this run contains a drawing (image)
        if (runElement && runElement["w:drawing"]) {
          if (zipHandler && imageManager) {
            const imageRun = await this.parseDrawingFromObject(
              runElement["w:drawing"],
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
        const hyperlink = this.parseHyperlinkFromObject(
          hyperlinkObj["w:hyperlink"],
          relationshipManager
        );
        if (hyperlink) {
          content.push(hyperlink);
        }
      }

      if (content.length === 0) {
        return null; // No content in revision
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

      return revision;
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse revision:",
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
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
      const paraId = paraObj["w14:paraId"];
      if (paraId) {
        paragraph.formatting.paraId = paraId;
      }

      // Parse paragraph properties
      this.parseParagraphPropertiesFromObject(paraObj["w:pPr"], paragraph);

      // Check if we have ordered children metadata from the enhanced parser
      const orderedChildren = paraObj["_orderedChildren"] as
        | Array<{ type: string; index: number }>
        | undefined;

      if (orderedChildren && orderedChildren.length > 0) {
        // Use the preserved order from the parser
        for (const childInfo of orderedChildren) {
          const elementType = childInfo.type;
          const elementIndex = childInfo.index;

          if (elementType === "w:r") {
            const runs = paraObj["w:r"];
            const runArray = Array.isArray(runs) ? runs : runs ? [runs] : [];
            if (elementIndex < runArray.length) {
              const child = runArray[elementIndex];
              if (child["w:drawing"]) {
                if (zipHandler && imageManager) {
                  const imageRun = await this.parseDrawingFromObject(
                    child["w:drawing"],
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
          } else if (elementType === "w:hyperlink") {
            const hyperlinks = paraObj["w:hyperlink"];
            const hyperlinkArray = Array.isArray(hyperlinks)
              ? hyperlinks
              : hyperlinks
              ? [hyperlinks]
              : [];
            if (elementIndex < hyperlinkArray.length) {
              const hyperlink = this.parseHyperlinkFromObject(
                hyperlinkArray[elementIndex],
                relationshipManager
              );
              if (hyperlink) {
                paragraph.addHyperlink(hyperlink);
              }
            }
          } else if (elementType === "w:fldSimple") {
            const fields = paraObj["w:fldSimple"];
            const fieldArray = Array.isArray(fields)
              ? fields
              : fields
              ? [fields]
              : [];
            if (elementIndex < fieldArray.length) {
              const field = this.parseSimpleFieldFromObject(
                fieldArray[elementIndex]
              );
              if (field) {
                paragraph.addField(field);
              }
            }
          }
        }
      } else {
        // Fallback to sequential processing if no order metadata
        // Handle runs (w:r)
        const runs = paraObj["w:r"];
        const runChildren = Array.isArray(runs) ? runs : runs ? [runs] : [];

        for (const child of runChildren) {
          if (child["w:drawing"]) {
            if (zipHandler && imageManager) {
              // Parse as image run
              const imageRun = await this.parseDrawingFromObject(
                child["w:drawing"],
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
        const hyperlinks = paraObj["w:hyperlink"];
        const hyperlinkChildren = Array.isArray(hyperlinks)
          ? hyperlinks
          : hyperlinks
          ? [hyperlinks]
          : [];

        for (const hyperlinkObj of hyperlinkChildren) {
          const hyperlink = this.parseHyperlinkFromObject(
            hyperlinkObj,
            relationshipManager
          );
          if (hyperlink) {
            paragraph.addHyperlink(hyperlink);
          }
        }

        // Handle simple fields (w:fldSimple)
        const fields = paraObj["w:fldSimple"];
        const fieldChildren = Array.isArray(fields)
          ? fields
          : fields
          ? [fields]
          : [];

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
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: "paragraph", error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse paragraph: ${err.message}`);
      }

      // In lenient mode, log warning and continue
      return null;
    }
  }

  private parseParagraphPropertiesFromObject(
    pPrObj: any,
    paragraph: Paragraph
  ): void {
    if (!pPrObj) return;

    // Paragraph mark run properties (w:rPr within w:pPr) per ECMA-376 Part 1 §17.3.1.29
    // This controls formatting of the paragraph mark (¶ symbol) itself
    if (pPrObj["w:rPr"]) {
      // Create a temporary Run to use the existing parseRunPropertiesFromObject method
      const tempRun = new Run("");
      this.parseRunPropertiesFromObject(pPrObj["w:rPr"], tempRun);
      // Extract the formatting and set it as paragraph mark properties
      paragraph.setParagraphMarkFormatting(tempRun.getFormatting());
    }

    // Alignment
    // XMLParser adds @_ prefix to attributes
    if (pPrObj["w:jc"]?.["@_w:val"]) {
      paragraph.setAlignment(pPrObj["w:jc"]["@_w:val"]);
    }

    // Style
    if (pPrObj["w:pStyle"]?.["@_w:val"]) {
      paragraph.setStyle(pPrObj["w:pStyle"]["@_w:val"]);
    }

    // Indentation
    if (pPrObj["w:ind"]) {
      const ind = pPrObj["w:ind"];
      if (ind["@_w:left"])
        paragraph.setLeftIndent(parseInt(ind["@_w:left"], 10));
      if (ind["@_w:right"])
        paragraph.setRightIndent(parseInt(ind["@_w:right"], 10));
      if (ind["@_w:firstLine"])
        paragraph.setFirstLineIndent(parseInt(ind["@_w:firstLine"], 10));
    }

    // Spacing
    if (pPrObj["w:spacing"]) {
      const spacing = pPrObj["w:spacing"];
      if (spacing["@_w:before"])
        paragraph.setSpaceBefore(parseInt(spacing["@_w:before"], 10));
      if (spacing["@_w:after"])
        paragraph.setSpaceAfter(parseInt(spacing["@_w:after"], 10));
      if (spacing["@_w:line"]) {
        paragraph.setLineSpacing(
          parseInt(spacing["@_w:line"], 10),
          spacing["@_w:lineRule"]
        );
      }
    }

    // Keep properties - parse pageBreakBefore FIRST, then apply keep properties
    // This triggers automatic conflict resolution per ECMA-376 v0.28.2
    if (pPrObj["w:pageBreakBefore"])
      paragraph.formatting.pageBreakBefore = true;

    // Keep properties - these will automatically clear pageBreakBefore if both are set
    if (pPrObj["w:keepNext"]) paragraph.setKeepNext(true);
    if (pPrObj["w:keepLines"]) paragraph.setKeepLines(true);

    // Contextual spacing
    if (pPrObj["w:contextualSpacing"]) paragraph.setContextualSpacing(true);

    // Numbering
    if (pPrObj["w:numPr"]) {
      const numPr = pPrObj["w:numPr"];
      const numId = numPr["w:numId"]?.["@_w:val"];
      const ilvl = numPr["w:ilvl"]?.["@_w:val"] || "0";
      if (numId) {
        paragraph.setNumbering(parseInt(numId, 10), parseInt(ilvl, 10));
      }
    }

    // Borders per ECMA-376 Part 1 §17.3.1.24
    if (pPrObj["w:pBdr"]) {
      const pBdr = pPrObj["w:pBdr"];
      const borders: any = {};

      // Helper function to parse border definition
      const parseBorder = (borderObj: any): any => {
        if (!borderObj) return undefined;
        const border: any = {};
        if (borderObj["@_w:val"]) border.style = borderObj["@_w:val"];
        if (borderObj["@_w:sz"])
          border.size = parseInt(borderObj["@_w:sz"], 10);
        if (borderObj["@_w:color"]) border.color = borderObj["@_w:color"];
        if (borderObj["@_w:space"])
          border.space = parseInt(borderObj["@_w:space"], 10);
        return Object.keys(border).length > 0 ? border : undefined;
      };

      // Parse each border side
      if (pBdr["w:top"]) borders.top = parseBorder(pBdr["w:top"]);
      if (pBdr["w:bottom"]) borders.bottom = parseBorder(pBdr["w:bottom"]);
      if (pBdr["w:left"]) borders.left = parseBorder(pBdr["w:left"]);
      if (pBdr["w:right"]) borders.right = parseBorder(pBdr["w:right"]);
      if (pBdr["w:between"]) borders.between = parseBorder(pBdr["w:between"]);
      if (pBdr["w:bar"]) borders.bar = parseBorder(pBdr["w:bar"]);

      if (Object.keys(borders).length > 0) {
        paragraph.setBorder(borders);
      }
    }

    // Shading per ECMA-376 Part 1 §17.3.1.32
    if (pPrObj["w:shd"]) {
      const shd = pPrObj["w:shd"];
      const shading: any = {};
      if (shd["@_w:fill"]) shading.fill = shd["@_w:fill"];
      if (shd["@_w:color"]) shading.color = shd["@_w:color"];
      if (shd["@_w:val"]) shading.val = shd["@_w:val"];

      if (Object.keys(shading).length > 0) {
        paragraph.setShading(shading);
      }
    }

    // Tab stops per ECMA-376 Part 1 §17.3.1.38
    if (pPrObj["w:tabs"]) {
      const tabsObj = pPrObj["w:tabs"];
      const tabs: any[] = [];

      // Handle both single tab and array of tabs
      const tabElements = Array.isArray(tabsObj["w:tab"])
        ? tabsObj["w:tab"]
        : tabsObj["w:tab"]
        ? [tabsObj["w:tab"]]
        : [];

      for (const tabObj of tabElements) {
        const tab: any = {};
        if (tabObj["@_w:pos"]) tab.position = parseInt(tabObj["@_w:pos"], 10);
        if (tabObj["@_w:val"]) tab.val = tabObj["@_w:val"];
        if (tabObj["@_w:leader"]) tab.leader = tabObj["@_w:leader"];

        if (tab.position !== undefined) {
          tabs.push(tab);
        }
      }

      if (tabs.length > 0) {
        paragraph.setTabs(tabs);
      }
    }

    // Widow control per ECMA-376 Part 1 §17.3.1.40
    if (pPrObj["w:widowControl"] !== undefined) {
      const widowControlVal = pPrObj["w:widowControl"]?.["@_w:val"];
      // Parse w:val attribute - can be "0"/"1" or "false"/"true"
      if (
        widowControlVal === "0" ||
        widowControlVal === "false" ||
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
    if (
      pPrObj["w:outlineLvl"] !== undefined &&
      pPrObj["w:outlineLvl"]["@_w:val"] !== undefined
    ) {
      const level = parseInt(pPrObj["w:outlineLvl"]["@_w:val"], 10);
      if (!isNaN(level) && level >= 0 && level <= 9) {
        paragraph.setOutlineLevel(level);
      }
    }

    // Suppress line numbers per ECMA-376 Part 1 §17.3.1.34
    if (pPrObj["w:suppressLineNumbers"]) {
      paragraph.setSuppressLineNumbers(true);
    }

    // Bidirectional layout per ECMA-376 Part 1 §17.3.1.6
    if (pPrObj["w:bidi"] !== undefined) {
      const bidiVal = pPrObj["w:bidi"]?.["@_w:val"];
      if (
        bidiVal === "0" ||
        bidiVal === "false" ||
        bidiVal === false ||
        bidiVal === 0
      ) {
        paragraph.setBidi(false);
      } else {
        // Default is true when element present without val attribute or val="1"
        paragraph.setBidi(true);
      }
    }

    // Text direction per ECMA-376 Part 1 §17.3.1.36
    if (pPrObj["w:textDirection"]?.["@_w:val"]) {
      paragraph.setTextDirection(pPrObj["w:textDirection"]["@_w:val"]);
    }

    // Text vertical alignment per ECMA-376 Part 1 §17.3.1.35
    if (pPrObj["w:textAlignment"]?.["@_w:val"]) {
      paragraph.setTextAlignment(pPrObj["w:textAlignment"]["@_w:val"]);
    }

    // Mirror indents per ECMA-376 Part 1 §17.3.1.18
    if (pPrObj["w:mirrorIndents"]) {
      paragraph.setMirrorIndents(true);
    }

    // Auto-adjust right indent per ECMA-376 Part 1 §17.3.1.1
    if (pPrObj["w:adjustRightInd"] !== undefined) {
      const adjustRightIndVal = pPrObj["w:adjustRightInd"]?.["@_w:val"];
      if (
        adjustRightIndVal === "0" ||
        adjustRightIndVal === "false" ||
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
    if (pPrObj["w:framePr"]) {
      const framePr = pPrObj["w:framePr"];
      const frameProps: any = {};
      if (framePr["@_w:w"]) frameProps.w = parseInt(framePr["@_w:w"], 10);
      if (framePr["@_w:h"]) frameProps.h = parseInt(framePr["@_w:h"], 10);
      if (framePr["@_w:hRule"]) frameProps.hRule = framePr["@_w:hRule"];
      if (framePr["@_w:x"]) frameProps.x = parseInt(framePr["@_w:x"], 10);
      if (framePr["@_w:y"]) frameProps.y = parseInt(framePr["@_w:y"], 10);
      if (framePr["@_w:xAlign"]) frameProps.xAlign = framePr["@_w:xAlign"];
      if (framePr["@_w:yAlign"]) frameProps.yAlign = framePr["@_w:yAlign"];
      if (framePr["@_w:hAnchor"]) frameProps.hAnchor = framePr["@_w:hAnchor"];
      if (framePr["@_w:vAnchor"]) frameProps.vAnchor = framePr["@_w:vAnchor"];
      if (framePr["@_w:hSpace"])
        frameProps.hSpace = parseInt(framePr["@_w:hSpace"], 10);
      if (framePr["@_w:vSpace"])
        frameProps.vSpace = parseInt(framePr["@_w:vSpace"], 10);
      if (framePr["@_w:wrap"]) frameProps.wrap = framePr["@_w:wrap"];
      if (framePr["@_w:dropCap"]) frameProps.dropCap = framePr["@_w:dropCap"];
      if (framePr["@_w:lines"])
        frameProps.lines = parseInt(framePr["@_w:lines"], 10);
      if (framePr["@_w:anchorLock"] !== undefined) {
        const anchorLockVal = framePr["@_w:anchorLock"];
        frameProps.anchorLock =
          anchorLockVal === "1" ||
          anchorLockVal === "true" ||
          anchorLockVal === true ||
          anchorLockVal === 1;
      }
      if (Object.keys(frameProps).length > 0) {
        paragraph.setFrameProperties(frameProps);
      }
    }

    // Suppress automatic hyphenation per ECMA-376 Part 1 §17.3.1.33
    if (pPrObj["w:suppressAutoHyphens"]) {
      paragraph.setSuppressAutoHyphens(true);
    }

    // Suppress text frame overlap per ECMA-376 Part 1 §17.3.1.34
    if (pPrObj["w:suppressOverlap"]) {
      paragraph.setSuppressOverlap(true);
    }

    // Textbox tight wrap per ECMA-376 Part 1 §17.3.1.37
    if (pPrObj["w:textboxTightWrap"]) {
      const wrapVal = pPrObj["w:textboxTightWrap"]?.["@_w:val"];
      if (wrapVal) {
        paragraph.setTextboxTightWrap(wrapVal);
      }
    }

    // HTML div ID per ECMA-376 Part 1 §17.3.1.9
    if (pPrObj["w:divId"]) {
      const divIdVal = pPrObj["w:divId"]?.["@_w:val"];
      if (divIdVal) {
        paragraph.setDivId(parseInt(divIdVal, 10));
      }
    }

    // Conditional table style formatting per ECMA-376 Part 1 §17.3.1.8
    if (pPrObj["w:cnfStyle"]) {
      const cnfStyleVal = pPrObj["w:cnfStyle"]?.["@_w:val"];
      if (cnfStyleVal !== undefined) {
        // Ensure it's a string and pad to 12 characters (standard bitmask length)
        // XML parser may convert to number, removing leading zeros
        const bitmask = String(cnfStyleVal).padStart(12, "0");
        paragraph.setConditionalFormatting(bitmask);
      }
    }

    // Paragraph property change tracking per ECMA-376 Part 1 §17.3.1.27
    if (pPrObj["w:pPrChange"]) {
      const changeObj = pPrObj["w:pPrChange"];
      const change: any = {};
      if (changeObj["@_w:author"])
        change.author = String(changeObj["@_w:author"]);
      if (changeObj["@_w:date"]) change.date = String(changeObj["@_w:date"]);
      if (changeObj["@_w:id"]) change.id = String(changeObj["@_w:id"]);
      // Note: Full implementation would parse child <w:pPr> for previousProperties
      if (Object.keys(change).length > 0) {
        paragraph.setParagraphPropertiesChange(change);
      }
    }

    // Section properties per ECMA-376 Part 1 §17.3.1.30
    if (pPrObj["w:sectPr"]) {
      // Simplified: store as-is (complex structure)
      // Full implementation would parse complete sectPr structure
      paragraph.setSectionProperties(pPrObj["w:sectPr"]);
    }
  }

  /**
   * NEW: Assemble complex fields from run tokens
   * Groups begin→instr→sep→result→end sequences into ComplexField objects
   *
   * @param paragraph The paragraph containing runs to process
   */
  private assembleComplexFields(paragraph: Paragraph): void {
    const content = paragraph.getContent();
    const groupedContent: any[] = [];
    let fieldRuns: Run[] = [];
    let fieldState:
      | "begin"
      | "instruction"
      | "separate"
      | "result"
      | "end"
      | null = null;

    for (let i = 0; i < content.length; i++) {
      const item = content[i];

      if (item instanceof Run) {
        const runContent = item.getContent();
        const hasFieldContent = runContent.some(
          (c: any) => c.type === "fieldChar" || c.type === "instructionText"
        );

        if (hasFieldContent) {
          // This run is part of a field
          fieldRuns.push(item);
          const fieldChar = runContent.find((c: any) => c.type === "fieldChar");

          if (fieldChar) {
            switch (fieldChar.fieldCharType) {
              case "begin":
                fieldState = "begin";
                break;
              case "separate":
                fieldState = "separate";
                break;
              case "end":
                fieldState = "end";

                // Complete field assembly
                if (fieldState === "end" && fieldRuns.length > 0) {
                  const complexField =
                    this.createComplexFieldFromRuns(fieldRuns);
                  if (complexField) {
                    groupedContent.push(complexField);
                  } else {
                    // If assembly failed, add individual runs
                    fieldRuns.forEach((run) => groupedContent.push(run));
                  }
                  fieldRuns = [];
                  fieldState = null;
                }
                break;
            }
          } else {
            // Instruction text run
            if (fieldState === "begin" || fieldState === "instruction") {
              fieldState = "instruction";
            }
          }
        } else {
          // Regular run - add it as-is
          if (fieldRuns.length > 0) {
            // Incomplete field - add as individual runs
            fieldRuns.forEach((run) => groupedContent.push(run));
            fieldRuns = [];
          }
          groupedContent.push(item);
        }
      } else {
        // Non-run content (hyperlinks, images, etc.)
        if (fieldRuns.length > 0) {
          // Incomplete field - add as individual runs
          fieldRuns.forEach((run) => groupedContent.push(run));
          fieldRuns = [];
        }
        groupedContent.push(item);
      }
    }

    // Handle any remaining incomplete field
    if (fieldRuns.length > 0) {
      fieldRuns.forEach((run) => groupedContent.push(run));
    }

    // Replace paragraph content with grouped content using setContent
    paragraph.setContent(groupedContent as ParagraphContent[]);

    defaultLogger.debug(
      `Assembled ${
        groupedContent.length - content.length
      } complex fields in paragraph`
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
        "Skipping ComplexField assembly: insufficient runs (minimum 2: begin and instr)"
      );
      return null;
    }

    let instruction = "";
    let resultText = "";
    let instructionFormatting: RunFormatting | undefined;
    let resultFormatting: RunFormatting | undefined;
    let hasBegin = false;
    let hasEnd = false;
    let hasSeparate = false;

    for (const run of fieldRuns) {
      const runContent = run.getContent();

      // Check for fieldChar tokens
      const fieldCharToken = runContent.find(
        (c: any) => c.type === "fieldChar"
      );
      if (fieldCharToken) {
        switch (fieldCharToken.fieldCharType) {
          case "begin":
            hasBegin = true;
            // Capture formatting from begin run
            instructionFormatting = run.getFormatting();
            break;
          case "separate":
            hasSeparate = true;
            break;
          case "end":
            hasEnd = true;
            break;
        }
      }

      // Check for instruction text
      const instrText = runContent.find(
        (c: any) => c.type === "instructionText"
      );
      if (instrText) {
        instruction += instrText.value || "";
      }

      // Check for result text (between separate and end)
      const textContent = runContent.find((c: any) => c.type === "text");
      if (textContent && hasSeparate) {
        resultText += textContent.value || "";
        resultFormatting = run.getFormatting();
      }
    }

    // Validate field structure with detailed diagnostics
    if (!hasBegin) {
      const instrPreview = instruction
        ? instruction.substring(0, 50)
        : "<none>";
      defaultLogger.warn(
        `ComplexField missing 'begin' marker. Instruction: "${instrPreview}..."`
      );
      this.parseErrors.push({
        element: "complex-field-structure",
        error: new Error("Missing field begin marker"),
      });
      return null;
    }

    if (!hasEnd) {
      const instrPreview = instruction
        ? instruction.substring(0, 50)
        : "<none>";
      defaultLogger.warn(
        `ComplexField missing 'end' marker. Instruction: "${instrPreview}..."`
      );
      this.parseErrors.push({
        element: "complex-field-structure",
        error: new Error("Missing field end marker"),
      });
      return null;
    }

    if (!instruction.trim()) {
      defaultLogger.warn(`ComplexField has no instruction content`);
      this.parseErrors.push({
        element: "complex-field-structure",
        error: new Error("Empty field instruction"),
      });
      return null;
    }

    // Trim and clean instruction
    instruction = instruction.trim();

    defaultLogger.debug(
      `Created ComplexField: ${instruction.substring(
        0,
        50
      )}... (result: "${resultText}")`
    );

    const properties: any = {
      instruction,
      result: resultText,
      instructionFormatting,
      resultFormatting,
      multiParagraph: false, // Default - can be set later if needed
    };

    return new ComplexField(properties);
  }

  private parseRunFromObject(runObj: any): Run | null {
    try {
      // Extract all run content elements (text, tabs, breaks, etc.)
      // Per ECMA-376 §17.3.3 EG_RunInnerContent, runs can contain multiple content types
      const content: RunContent[] = [];

      const toArray = <T>(value: T | T[] | undefined | null): T[] =>
        Array.isArray(value)
          ? value
          : value !== undefined && value !== null
          ? [value]
          : [];

      const extractTextValue = (node: any): string => {
        if (node === undefined || node === null) {
          return "";
        }
        if (typeof node === "object") {
          return XMLBuilder.unescapeXml(node["#text"] || "");
        }
        return XMLBuilder.unescapeXml(String(node));
      };

      const parseBooleanAttr = (value: any): boolean | undefined => {
        if (value === undefined || value === null) {
          return undefined;
        }
        return (
          value === "1" || value === 1 || value === true || value === "true"
        );
      };

      // Use _orderedChildren to preserve element order (critical for TOC entries)
      // TOC entries have structure: text → tab → text (heading, tab, page number)
      if (runObj["_orderedChildren"]) {
        // Process elements in their original order
        for (const child of runObj["_orderedChildren"]) {
          const elementType = child.type;
          const elementIndex = child.index;

          switch (elementType) {
            case "w:t": {
              const textElements = toArray(runObj["w:t"]);
              const te = textElements[elementIndex];
              if (te !== undefined && te !== null) {
                const text = extractTextValue(te);
                if (text) {
                  content.push({ type: "text", value: text });
                }
              }
              break;
            }

            case "w:instrText": {
              const instrElements = toArray(runObj["w:instrText"]);
              const instr = instrElements[elementIndex];
              if (instr !== undefined && instr !== null) {
                const text = extractTextValue(instr);
                content.push({ type: "instructionText", value: text });
              }
              break;
            }

            case "w:fldChar": {
              const fldChars = toArray(runObj["w:fldChar"]);
              const fldChar = fldChars[elementIndex];
              if (fldChar && typeof fldChar === "object") {
                const charType = (fldChar["@_w:fldCharType"] ||
                  fldChar["@_fldCharType"]) as
                  | "begin"
                  | "separate"
                  | "end"
                  | undefined;
                if (charType) {
                  content.push({
                    type: "fieldChar",
                    fieldCharType: charType,
                    fieldCharDirty: parseBooleanAttr(fldChar["@_w:dirty"]),
                    fieldCharLocked: parseBooleanAttr(
                      fldChar["@_w:fldLock"] ?? fldChar["@_w:lock"]
                    ),
                  });
                }
              }
              break;
            }

            case "w:tab":
              content.push({ type: "tab" });
              break;

            case "w:br": {
              const brElements = toArray(runObj["w:br"]);
              const brElement = brElements[elementIndex] || brElements[0];
              const breakType = brElement?.["@_w:type"] as
                | BreakType
                | undefined;
              content.push({ type: "break", breakType });
              break;
            }

            case "w:cr":
              content.push({ type: "carriageReturn" });
              break;

            case "w:softHyphen":
              content.push({ type: "softHyphen" });
              break;

            case "w:noBreakHyphen":
              content.push({ type: "noBreakHyphen" });
              break;

            // Ignore formatting elements (w:rPr) - handled separately
            case "w:rPr":
              break;
          }
        }
      } else {
        // Fallback: No _orderedChildren (older parser or simple run with single text)
        // Extract text elements (can be array if multiple <w:t> in one run)
        const textElement = runObj["w:t"];
        if (textElement !== undefined && textElement !== null) {
          const textElements = toArray(textElement);

          for (const te of textElements) {
            const text = extractTextValue(te);
            if (text) {
              content.push({ type: "text", value: text });
            }
          }
        }

        const instrTextElement = runObj["w:instrText"];
        if (instrTextElement !== undefined && instrTextElement !== null) {
          const instrElements = toArray(instrTextElement);
          for (const instr of instrElements) {
            const text = extractTextValue(instr);
            content.push({ type: "instructionText", value: text });
          }
        }

        const fldCharElement = runObj["w:fldChar"];
        if (fldCharElement !== undefined && fldCharElement !== null) {
          const fldChars = toArray(fldCharElement);
          for (const fldChar of fldChars) {
            if (fldChar && typeof fldChar === "object") {
              const charType = (fldChar["@_w:fldCharType"] ||
                fldChar["@_fldCharType"]) as
                | "begin"
                | "separate"
                | "end"
                | undefined;
              if (charType) {
                content.push({
                  type: "fieldChar",
                  fieldCharType: charType,
                  fieldCharDirty: parseBooleanAttr(fldChar["@_w:dirty"]),
                  fieldCharLocked: parseBooleanAttr(
                    fldChar["@_w:fldLock"] ?? fldChar["@_w:lock"]
                  ),
                });
              }
            }
          }
        }

        // Extract other elements (order doesn't matter in simple case)
        if (runObj["w:tab"] !== undefined) {
          content.push({ type: "tab" });
        }

        if (runObj["w:br"] !== undefined) {
          const brElement = runObj["w:br"];
          const breakType = brElement?.["@_w:type"] as BreakType | undefined;
          content.push({ type: "break", breakType });
        }

        if (runObj["w:cr"] !== undefined) {
          content.push({ type: "carriageReturn" });
        }

        if (runObj["w:softHyphen"] !== undefined) {
          content.push({ type: "softHyphen" });
        }

        if (runObj["w:noBreakHyphen"] !== undefined) {
          content.push({ type: "noBreakHyphen" });
        }
      }

      // Create run from content elements
      const run = Run.createFromContent(content, { cleanXmlFromText: false });

      // Parse and apply run properties (formatting)
      this.parseRunPropertiesFromObject(runObj["w:rPr"], run);

      // Diagnostic logging
      const text = run.getText();
      const formatting = run.getFormatting();
      if (formatting.rtl) {
        logTextDirection(`Run with RTL: "${text}"`);
      }
      logParsing(
        `Parsed run: "${text}" (${content.length} content element(s))`,
        { rtl: formatting.rtl || false }
      );

      return run;
    } catch (error) {
      return null;
    }
  }

  private parseHyperlinkFromObject(
    hyperlinkObj: any,
    relationshipManager: RelationshipManager
  ): Hyperlink | null {
    try {
      // Extract hyperlink attributes
      const relationshipId = hyperlinkObj["@_r:id"];
      const anchor = hyperlinkObj["@_w:anchor"];
      const tooltip = hyperlinkObj["@_w:tooltip"];

      // Parse runs inside the hyperlink
      const runs = hyperlinkObj["w:r"];
      const runChildren = Array.isArray(runs) ? runs : runs ? [runs] : [];

      // Parse ALL runs to handle multi-run hyperlinks (e.g., varied formatting within one hyperlink)
      // Google Docs often splits hyperlinks by formatting changes, creating multiple runs
      const parsedRuns: Run[] = [];
      let text = "";
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
        const relationship =
          relationshipManager.getRelationship(relationshipId);
        if (relationship) {
          url = relationship.getTarget();
        }
      }

      // Create hyperlink with basic properties
      // NOTE: Do NOT use anchor (bookmark ID) as display text - it should only be used for navigation
      let displayText = text || url || "[Link]";

      // Warn if hyperlink has no display text (possible TOC corruption or malformed hyperlink)
      if (!text && anchor) {
        defaultLogger.warn(
          `[DocumentParser] Hyperlink to anchor "${anchor}" has no display text. ` +
            `Using placeholder "[Link]" to prevent bookmark ID from appearing as visible text. ` +
            `This may indicate a corrupted TOC or malformed hyperlink in the source document.`
        );
      }

      const hyperlink = new Hyperlink({
        url,
        anchor,
        text: displayText,
        formatting,
        tooltip,
        relationshipId,
      });

      // If we successfully parsed a run with tabs/breaks, use it instead of the default run
      // This preserves TOC structure (text → tab → text)
      if (parsedRun && parsedRun.getContent().length > 1) {
        hyperlink.setRun(parsedRun);
      }

      return hyperlink;
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse hyperlink:",
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
  }

  /**
   * Merges hyperlinks with the same URL into a single hyperlink (handles fragmentation)
   * This handles Google Docs-style hyperlinks that are split by formatting changes
   * Now enhanced to merge non-consecutive hyperlinks with the same URL
   * @param paragraph - Paragraph containing hyperlinks to merge
   * @param resetFormatting - Whether to reset hyperlinks to standard formatting
   * @private
   */
  private mergeConsecutiveHyperlinks(
    paragraph: Paragraph,
    resetFormatting: boolean = false
  ): void {
    const content = paragraph.getContent();
    if (!content || content.length < 2) return;

    // First pass: Group all hyperlinks by URL/anchor
    const hyperlinkGroups = new Map<string, any[]>();
    const nonHyperlinkItems: { item: any; index: number }[] = [];
    const hyperlinkIndices = new Map<any, number>();

    for (let i = 0; i < content.length; i++) {
      const item = content[i];

      if (item instanceof Hyperlink) {
        const url = item.getUrl() || "";
        const anchor = item.getAnchor() || "";
        const key = `${url}|${anchor}`; // Unique key for URL+anchor combination

        if (!hyperlinkGroups.has(key)) {
          hyperlinkGroups.set(key, []);
        }
        hyperlinkGroups.get(key)!.push(item);
        hyperlinkIndices.set(item, i);
      } else {
        nonHyperlinkItems.push({ item, index: i });
      }
    }

    // Check if any merging is needed
    let needsMerge = false;
    for (const group of hyperlinkGroups.values()) {
      if (group.length > 1) {
        needsMerge = true;
        break;
      }
    }

    if (!needsMerge && !resetFormatting) {
      return; // Nothing to do
    }

    // Second pass: Build merged content preserving original order
    const mergedContent: any[] = [];
    const processedIndices = new Set<number>();

    for (let i = 0; i < content.length; i++) {
      if (processedIndices.has(i)) {
        continue; // Skip already processed items
      }

      const item = content[i];

      if (item instanceof Hyperlink) {
        const url = item.getUrl() || "";
        const anchor = item.getAnchor() || "";
        const key = `${url}|${anchor}`;
        const group = hyperlinkGroups.get(key)!;

        if (group.length > 1 && group[0] === item) {
          // This is the first hyperlink in a group that needs merging
          // Collect all text from the group
          const mergedText = group.map((h) => h.getText()).join("");

          // Create merged hyperlink using first hyperlink's properties
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

          // Mark all group members as processed
          for (const h of group) {
            processedIndices.add(hyperlinkIndices.get(h)!);
          }

          mergedContent.push(mergedHyperlink);
        } else if (group.length === 1) {
          // Single hyperlink, possibly reset formatting
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
          } else {
            mergedContent.push(item);
          }
          processedIndices.add(i);
        }
      } else {
        // Not a hyperlink, keep as-is
        mergedContent.push(item);
        processedIndices.add(i);
      }
    }

    // Update paragraph content if we changed anything
    if (needsMerge || resetFormatting) {
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
   * Get standard hyperlink formatting (Calibri, blue, underline)
   * @private
   */
  private getStandardHyperlinkFormatting(): any {
    return {
      font: "Verdana",
      color: "0000FF", // Standard hyperlink blue
      underline: "single",
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
      const instruction = fieldObj["@_w:instr"];
      if (!instruction) {
        return null;
      }

      // Extract field type from instruction (first word)
      const typeMatch = instruction.trim().match(/^(\w+)/);
      const type = (typeMatch?.[1] ||
        "PAGE") as import("../elements/Field").FieldType;

      // Parse run formatting from w:rPr if present
      let formatting: RunFormatting | undefined;
      if (fieldObj["w:rPr"]) {
        const tempRun = new Run("");
        this.parseRunPropertiesFromObject(fieldObj["w:rPr"], tempRun);
        formatting = tempRun.getFormatting();
      }

      // Create field with instruction
      const field = Field.create({
        type,
        instruction,
        formatting,
      });

      return field;
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse field:",
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
    if (rPrObj["w:rStyle"]) {
      const styleId = rPrObj["w:rStyle"]["@_w:val"];
      if (styleId) {
        run.setCharacterStyle(styleId);
      }
    }

    // Parse text border (w:bdr) per ECMA-376 Part 1 §17.3.2.5
    if (rPrObj["w:bdr"]) {
      const bdr = rPrObj["w:bdr"];
      const border: any = {};
      if (bdr["@_w:val"]) border.style = bdr["@_w:val"];
      if (bdr["@_w:sz"]) border.size = parseInt(bdr["@_w:sz"], 10);
      if (bdr["@_w:color"]) border.color = bdr["@_w:color"];
      if (bdr["@_w:space"]) border.space = parseInt(bdr["@_w:space"], 10);
      if (Object.keys(border).length > 0) {
        run.setBorder(border);
      }
    }

    // Parse character shading (w:shd) per ECMA-376 Part 1 §17.3.2.32
    if (rPrObj["w:shd"]) {
      const shd = rPrObj["w:shd"];
      const shading: any = {};
      if (shd["@_w:val"]) shading.val = shd["@_w:val"];
      if (shd["@_w:fill"]) shading.fill = shd["@_w:fill"];
      if (shd["@_w:color"]) shading.color = shd["@_w:color"];
      if (Object.keys(shading).length > 0) {
        run.setShading(shading);
      }
    }

    // Parse emphasis marks (w:em) per ECMA-376 Part 1 §17.3.2.13
    if (rPrObj["w:em"]) {
      const val = rPrObj["w:em"]["@_w:val"];
      if (val) run.setEmphasis(val as any);
    }

    // Parse outline text effect (w:outline) per ECMA-376 Part 1 §17.3.2.23
    if (rPrObj["w:outline"]) run.setOutline(true);

    // Parse shadow text effect (w:shadow) per ECMA-376 Part 1 §17.3.2.32
    if (rPrObj["w:shadow"]) run.setShadow(true);

    // Parse emboss text effect (w:emboss) per ECMA-376 Part 1 §17.3.2.13
    if (rPrObj["w:emboss"]) run.setEmboss(true);

    // Parse imprint text effect (w:imprint) per ECMA-376 Part 1 §17.3.2.18
    if (rPrObj["w:imprint"]) run.setImprint(true);

    // Parse no proofing (w:noProof) per ECMA-376 Part 1 §17.3.2.21
    if (rPrObj["w:noProof"]) run.setNoProof(true);

    // Parse snap to grid (w:snapToGrid) per ECMA-376 Part 1 §17.3.2.35
    if (rPrObj["w:snapToGrid"]) run.setSnapToGrid(true);

    // Parse vanish/hidden (w:vanish) per ECMA-376 Part 1 §17.3.2.42
    if (rPrObj["w:vanish"]) run.setVanish(true);

    // Parse special vanish (w:specVanish) per ECMA-376 Part 1 §17.3.2.36
    if (rPrObj["w:specVanish"]) run.setSpecVanish(true);

    // Boolean properties - check w:val attribute
    // Per ECMA-376: <w:b/> or <w:b w:val="1"/> or <w:b w:val="true"/> means true
    // <w:b w:val="0"/> or <w:b w:val="false"/> means false (omit from document)
    const checkBooleanProp = (prop: any): boolean => {
      if (!prop) return false;
      const val = prop["@_w:val"];
      // If no w:val attribute, self-closing tag means true
      if (val === undefined) return true;
      // Check for explicit true/false values (string or number)
      // Note: XMLParser converts "1" to number 1, "0" to number 0
      return val === "1" || val === 1 || val === "true" || val === true;
    };

    // Parse RTL text (w:rtl) per ECMA-376 Part 1 §17.3.2.30
    // FIX: Use checkBooleanProp to correctly handle w:val="0" (LTR) vs w:val="1" (RTL)
    if (checkBooleanProp(rPrObj["w:rtl"])) run.setRTL(true);

    if (checkBooleanProp(rPrObj["w:b"])) run.setBold(true);
    if (checkBooleanProp(rPrObj["w:bCs"])) run.setComplexScriptBold(true);
    if (checkBooleanProp(rPrObj["w:i"])) run.setItalic(true);
    if (checkBooleanProp(rPrObj["w:iCs"])) run.setComplexScriptItalic(true);
    if (checkBooleanProp(rPrObj["w:strike"])) run.setStrike(true);
    if (checkBooleanProp(rPrObj["w:smallCaps"])) run.setSmallCaps(true);
    if (checkBooleanProp(rPrObj["w:caps"])) run.setAllCaps(true);

    if (rPrObj["w:u"]) {
      // XMLParser adds @_ prefix to attributes
      const uVal = rPrObj["w:u"]["@_w:val"];
      run.setUnderline(uVal || true);
    }

    // Parse character spacing (w:spacing) per ECMA-376 Part 1 §17.3.2.33
    if (rPrObj["w:spacing"]) {
      const val = rPrObj["w:spacing"]["@_w:val"];
      if (val) run.setCharacterSpacing(parseInt(val, 10));
    }

    // Parse horizontal scaling (w:w) per ECMA-376 Part 1 §17.3.2.43
    if (rPrObj["w:w"]) {
      const val = rPrObj["w:w"]["@_w:val"];
      if (val) run.setScaling(parseInt(val, 10));
    }

    // Parse vertical position (w:position) per ECMA-376 Part 1 §17.3.2.31
    if (rPrObj["w:position"]) {
      const val = rPrObj["w:position"]["@_w:val"];
      if (val) run.setPosition(parseInt(val, 10));
    }

    // Parse kerning (w:kern) per ECMA-376 Part 1 §17.3.2.20
    if (rPrObj["w:kern"]) {
      const val = rPrObj["w:kern"]["@_w:val"];
      if (val) run.setKerning(parseInt(val, 10));
    }

    // Parse language (w:lang) per ECMA-376 Part 1 §17.3.2.20
    if (rPrObj["w:lang"]) {
      const val = rPrObj["w:lang"]["@_w:val"];
      if (val) run.setLanguage(val);
    }

    // Parse East Asian layout (w:eastAsianLayout) per ECMA-376 Part 1 §17.3.2.10
    if (rPrObj["w:eastAsianLayout"]) {
      const layoutObj = rPrObj["w:eastAsianLayout"];
      const layout: any = {};
      if (layoutObj["@_w:id"] !== undefined)
        layout.id = Number(layoutObj["@_w:id"]);
      if (layoutObj["@_w:vert"]) layout.vert = true;
      if (layoutObj["@_w:vertCompress"]) layout.vertCompress = true;
      if (layoutObj["@_w:combine"]) layout.combine = true;
      if (layoutObj["@_w:combineBrackets"])
        layout.combineBrackets = layoutObj["@_w:combineBrackets"];

      if (Object.keys(layout).length > 0) {
        run.setEastAsianLayout(layout);
      }
    }

    // Parse fit text (w:fitText) per ECMA-376 Part 1 §17.3.2.15
    if (rPrObj["w:fitText"]) {
      const val = rPrObj["w:fitText"]["@_w:val"];
      if (val !== undefined) run.setFitText(Number(val));
    }

    // Parse text effect (w:effect) per ECMA-376 Part 1 §17.3.2.12
    if (rPrObj["w:effect"]) {
      const val = rPrObj["w:effect"]["@_w:val"];
      if (val) run.setEffect(val as any);
    }

    if (rPrObj["w:vertAlign"]) {
      const val = rPrObj["w:vertAlign"]["@_w:val"];
      if (val === "subscript") run.setSubscript(true);
      if (val === "superscript") run.setSuperscript(true);
    }

    if (rPrObj["w:rFonts"]) {
      run.setFont(rPrObj["w:rFonts"]["@_w:ascii"]);
    }

    if (rPrObj["w:sz"]) {
      run.setSize(parseInt(rPrObj["w:sz"]["@_w:val"], 10) / 2);
    }

    if (rPrObj["w:color"]) {
      run.setColor(rPrObj["w:color"]["@_w:val"]);
    }

    if (rPrObj["w:highlight"]) {
      run.setHighlight(rPrObj["w:highlight"]["@_w:val"]);
    }
  }

  private async parseDrawingFromObject(
    drawingObj: any,
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager,
    imageManager: ImageManager
  ): Promise<ImageRun | null> {
    try {
      // Drawing can contain either wp:inline (inline image) or wp:anchor (floating image)
      const inlineObj = drawingObj["wp:inline"];
      const anchorObj = drawingObj["wp:anchor"];
      const imageObj = inlineObj || anchorObj;

      if (!imageObj) {
        return null;
      }

      const isFloating = !!anchorObj;

      // Extract dimensions from wp:extent
      const extentObj = imageObj["wp:extent"];
      let width = 0;
      let height = 0;
      if (extentObj) {
        width = parseInt(extentObj["@_cx"] || "0", 10);
        height = parseInt(extentObj["@_cy"] || "0", 10);
      }

      // Extract effect extent
      const effectExtentObj = imageObj["wp:effectExtent"];
      let effectExtent = undefined;
      if (effectExtentObj) {
        effectExtent = {
          left: parseInt(effectExtentObj["@_l"] || "0", 10),
          top: parseInt(effectExtentObj["@_t"] || "0", 10),
          right: parseInt(effectExtentObj["@_r"] || "0", 10),
          bottom: parseInt(effectExtentObj["@_b"] || "0", 10),
        };
      }

      // Extract name, description, and ID from wp:docPr
      const docPrObj = imageObj["wp:docPr"];
      let name = "image";
      let description = "Image";
      let docPrId = 1;
      if (docPrObj) {
        name = docPrObj["@_name"] || "image";
        description = docPrObj["@_descr"] || "Image";
        // Parse docPr id to preserve unique image IDs
        const idAttr = docPrObj["@_id"];
        if (idAttr) {
          docPrId = parseInt(String(idAttr), 10);
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
        // XMLParser converts attribute values - handle both string and number
        const toBool = (val: any) => val === "1" || val === 1 || val === true;

        anchor = {
          behindDoc: toBool(anchorObj["@_behindDoc"]),
          locked: toBool(anchorObj["@_locked"]),
          layoutInCell: toBool(anchorObj["@_layoutInCell"]),
          allowOverlap: toBool(anchorObj["@_allowOverlap"]),
          relativeHeight: parseInt(
            anchorObj["@_relativeHeight"] || "251658240",
            10
          ),
        };
      }

      // Navigate through the graphic structure to find the relationship ID
      // Structure: a:graphic → a:graphicData → pic:pic → pic:blipFill → a:blip
      const graphicObj = imageObj["a:graphic"];
      if (!graphicObj) {
        return null;
      }

      const graphicDataObj = graphicObj["a:graphicData"];
      if (!graphicDataObj) {
        return null;
      }

      const picPicObj = graphicDataObj["pic:pic"];
      if (!picPicObj) {
        return null;
      }

      const blipFillObj = picPicObj["pic:blipFill"];
      if (!blipFillObj) {
        return null;
      }

      const blipObj = blipFillObj["a:blip"];
      if (!blipObj) {
        return null;
      }

      // Parse crop settings (a:srcRect is sibling of a:blip in blipFill, not child of blip)
      const crop = this.parseImageCrop(blipFillObj);

      // Parse effects (effects are children of a:blip)
      const effects = this.parseImageEffects(blipObj);

      // Extract relationship ID (r:embed)
      const relationshipId = blipObj["@_r:embed"];
      if (!relationshipId) {
        return null;
      }

      // Get the image from the relationship
      const relationship = relationshipManager.getRelationship(relationshipId);
      if (!relationship) {
        defaultLogger.warn(
          `[DocumentParser] Image relationship not found: ${relationshipId}`
        );
        return null;
      }

      const imageTarget = relationship.getTarget();
      if (!imageTarget) {
        defaultLogger.warn(
          `[DocumentParser] Image relationship has no target: ${relationshipId}`
        );
        return null;
      }

      // Read image data from zip
      const imagePath = `word/${imageTarget}`;
      const imageData = zipHandler.getFileAsBuffer(imagePath);
      if (!imageData) {
        defaultLogger.warn(
          `[DocumentParser] Image file not found: ${imagePath}`
        );
        return null;
      }

      // Detect image extension from path
      const extension = imagePath.split(".").pop()?.toLowerCase() || "png";

      // Create image from buffer with all properties
      const { Image: ImageClass } = await import("../elements/Image");
      const image = await ImageClass.create({
        source: imageData,
        width,
        height,
        name,
        description,
        effectExtent,
        wrap,
        position,
        anchor,
        crop,
        effects,
      });

      // Register image with ImageManager (reuse existing relationship)
      imageManager.registerImage(image, relationshipId);
      image.setRelationshipId(relationshipId);
      
      // Preserve the original docPr ID to prevent corruption
      image.setDocPrId(docPrId);

      // Create and return ImageRun
      return new ImageRun(image);
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse drawing:",
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
    const wrapSquare = anchorObj["wp:wrapSquare"];
    const wrapTight = anchorObj["wp:wrapTight"];
    const wrapThrough = anchorObj["wp:wrapThrough"];
    const wrapTopBottom = anchorObj["wp:wrapTopAndBottom"];
    const wrapNone = anchorObj["wp:wrapNone"];

    const wrapObj =
      wrapSquare || wrapTight || wrapThrough || wrapTopBottom || wrapNone;
    if (!wrapObj) {
      return undefined;
    }

    // Determine wrap type
    let type: any = "square";
    if (wrapTight) type = "tight";
    else if (wrapThrough) type = "through";
    else if (wrapTopBottom) type = "topAndBottom";
    else if (wrapNone) type = "none";

    return {
      type,
      side: wrapObj["@_wrapText"] || "bothSides",
      // Distance attributes are on the wrap element, not the anchor
      distanceTop: wrapObj["@_distT"]
        ? parseInt(wrapObj["@_distT"], 10)
        : undefined,
      distanceBottom: wrapObj["@_distB"]
        ? parseInt(wrapObj["@_distB"], 10)
        : undefined,
      distanceLeft: wrapObj["@_distL"]
        ? parseInt(wrapObj["@_distL"], 10)
        : undefined,
      distanceRight: wrapObj["@_distR"]
        ? parseInt(wrapObj["@_distR"], 10)
        : undefined,
    };
  }

  /**
   * Parses image position from anchor object
   * @private
   */
  private parseImagePosition(anchorObj: any): any {
    const posH = anchorObj["wp:positionH"];
    const posV = anchorObj["wp:positionV"];

    if (!posH || !posV) {
      return undefined;
    }

    // Parse horizontal position
    const horizontal: any = {
      anchor: posH["@_relativeFrom"] || "page",
    };

    if (posH["wp:posOffset"]) {
      const offsetText = Array.isArray(posH["wp:posOffset"])
        ? posH["wp:posOffset"][0]
        : posH["wp:posOffset"];
      horizontal.offset = parseInt(
        typeof offsetText === "string"
          ? offsetText
          : offsetText?.["#text"] || "0",
        10
      );
    } else if (posH["wp:align"]) {
      const alignText = Array.isArray(posH["wp:align"])
        ? posH["wp:align"][0]
        : posH["wp:align"];
      horizontal.alignment =
        typeof alignText === "string" ? alignText : alignText?.["#text"];
    }

    // Parse vertical position
    const vertical: any = {
      anchor: posV["@_relativeFrom"] || "page",
    };

    if (posV["wp:posOffset"]) {
      const offsetText = Array.isArray(posV["wp:posOffset"])
        ? posV["wp:posOffset"][0]
        : posV["wp:posOffset"];
      vertical.offset = parseInt(
        typeof offsetText === "string"
          ? offsetText
          : offsetText?.["#text"] || "0",
        10
      );
    } else if (posV["wp:align"]) {
      const alignText = Array.isArray(posV["wp:align"])
        ? posV["wp:align"][0]
        : posV["wp:align"];
      vertical.alignment =
        typeof alignText === "string" ? alignText : alignText?.["#text"];
    }

    return { horizontal, vertical };
  }

  /**
   * Parses image crop from blip object
   * @private
   */
  private parseImageCrop(blipObj: any): any {
    const srcRect = blipObj["a:srcRect"];
    if (!srcRect) {
      return undefined;
    }

    return {
      left: parseInt(srcRect["@_l"] || "0", 10) / 1000, // Convert from per-mille to percentage
      top: parseInt(srcRect["@_t"] || "0", 10) / 1000,
      right: parseInt(srcRect["@_r"] || "0", 10) / 1000,
      bottom: parseInt(srcRect["@_b"] || "0", 10) / 1000,
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
    const lum = blipObj["a:lum"];
    if (lum) {
      if (lum["@_bright"]) {
        effects.brightness = parseInt(lum["@_bright"], 10) / 1000; // Convert from per-mille
      }
      if (lum["@_contrast"]) {
        effects.contrast = parseInt(lum["@_contrast"], 10) / 1000;
      }
    }

    // Check for grayscale (direct child of a:blip)
    if (blipObj["a:grayscl"]) {
      effects.grayscale = true;
    }

    return Object.keys(effects).length > 0 ? effects : undefined;
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
      if (tableObj["w:tblPr"]) {
        this.parseTablePropertiesFromObject(tableObj["w:tblPr"], table);
      }

      // Parse table grid (w:tblGrid) - column widths
      if (tableObj["w:tblGrid"] && tableObj["w:tblGrid"]["w:gridCol"]) {
        const gridCols = tableObj["w:tblGrid"]["w:gridCol"];
        const gridColArray = Array.isArray(gridCols) ? gridCols : [gridCols];
        const widths = gridColArray.map((col: any) => {
          const w = col["@_w:w"];
          return w ? parseInt(w, 10) : 2880; // default to 2 inches
        });
        if (widths.length > 0) {
          table.setTableGrid(widths);
        }
      }

      // Parse table rows (w:tr)
      const rows = tableObj["w:tr"];
      const rowChildren = Array.isArray(rows) ? rows : rows ? [rows] : [];

      // Extract row XMLs from raw table XML if available
      let rowXmls: string[] = [];
      if (rawTableXml) {
        rowXmls = XMLParser.extractElements(rawTableXml, "w:tr");
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
        }
      }

      return table;
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse table:",
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

    // Parse table positioning (tblpPr) - for floating tables
    if (tblPrObj["w:tblpPr"]) {
      const tblpPr = tblPrObj["w:tblpPr"];
      const position: any = {};

      if (tblpPr["@_w:tblpX"]) position.x = parseInt(tblpPr["@_w:tblpX"], 10);
      if (tblpPr["@_w:tblpY"]) position.y = parseInt(tblpPr["@_w:tblpY"], 10);
      if (tblpPr["@_w:tblpXSpec"])
        position.horizontalAnchor = tblpPr["@_w:tblpXSpec"];
      if (tblpPr["@_w:tblpYSpec"])
        position.verticalAnchor = tblpPr["@_w:tblpYSpec"];
      if (tblpPr["@_w:tblpXAlign"])
        position.horizontalAlignment = tblpPr["@_w:tblpXAlign"];
      if (tblpPr["@_w:tblpYAlign"])
        position.verticalAlignment = tblpPr["@_w:tblpYAlign"];
      if (tblpPr["@_w:leftFromText"])
        position.leftFromText = parseInt(tblpPr["@_w:leftFromText"], 10);
      if (tblpPr["@_w:rightFromText"])
        position.rightFromText = parseInt(tblpPr["@_w:rightFromText"], 10);
      if (tblpPr["@_w:topFromText"])
        position.topFromText = parseInt(tblpPr["@_w:topFromText"], 10);
      if (tblpPr["@_w:bottomFromText"])
        position.bottomFromText = parseInt(tblpPr["@_w:bottomFromText"], 10);

      if (Object.keys(position).length > 0) {
        table.setPosition(position);
      }
    }

    // Parse table overlap
    if (tblPrObj["w:tblOverlap"]) {
      const val = tblPrObj["w:tblOverlap"]["@_w:val"];
      table.setOverlap(val === "overlap");
    }

    // Parse bidirectional visual layout
    if (tblPrObj["w:bidiVisual"]) {
      table.setBidiVisual(true);
    }

    // Parse table width
    if (tblPrObj["w:tblW"]) {
      const width = parseInt(tblPrObj["w:tblW"]["@_w:w"] || "0", 10);
      const widthType = tblPrObj["w:tblW"]["@_w:type"] || "dxa";
      if (width > 0) {
        table.setWidth(width);
        table.setWidthType(widthType as any);
      }
    }

    // Parse table caption
    if (tblPrObj["w:tblCaption"]) {
      const caption = tblPrObj["w:tblCaption"]["@_w:val"];
      if (caption) table.setCaption(caption);
    }

    // Parse table description
    if (tblPrObj["w:tblDescription"]) {
      const description = tblPrObj["w:tblDescription"]["@_w:val"];
      if (description) table.setDescription(description);
    }

    // Parse cell spacing
    if (tblPrObj["w:tblCellSpacing"]) {
      const spacing = parseInt(
        tblPrObj["w:tblCellSpacing"]["@_w:w"] || "0",
        10
      );
      const spacingType = tblPrObj["w:tblCellSpacing"]["@_w:type"] || "dxa";
      if (spacing > 0) {
        table.setCellSpacing(spacing);
        table.setCellSpacingType(spacingType as any);
      }
    }

    // Parse table alignment (w:jc) - IMPORTANT for preserving table centering
    if (tblPrObj["w:jc"]) {
      const alignment = tblPrObj["w:jc"]["@_w:val"];
      if (alignment) {
        table.setAlignment(alignment as "left" | "center" | "right");
      }
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
      const trPr = rowObj["w:trPr"];
      if (trPr) {
        this.parseTableRowPropertiesFromObject(trPr, row);
      }

      // Parse table property exceptions (w:tblPrEx) per ECMA-376 Part 1 §17.4.61
      const tblPrEx = rowObj["w:tblPrEx"];
      if (tblPrEx) {
        const exceptions = this.parseTablePropertyExceptionsFromObject(tblPrEx);
        if (exceptions) {
          row.setTablePropertyExceptions(exceptions);
        }
      }

      // Parse table cells (w:tc)
      const cells = rowObj["w:tc"];
      const cellChildren = Array.isArray(cells) ? cells : cells ? [cells] : [];

      // Extract cell XMLs from raw row XML if available
      let cellXmls: string[] = [];
      if (rawRowXml) {
        cellXmls = XMLParser.extractElements(rawRowXml, "w:tc");
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
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse table row:",
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
    if (trPrObj["w:trHeight"]) {
      const heightVal = parseInt(trPrObj["w:trHeight"]["@_w:val"] || "0", 10);
      const heightRule = trPrObj["w:trHeight"]["@_w:hRule"] || "atLeast";
      if (heightVal > 0) {
        row.setHeight(heightVal, heightRule as any);
      }
    }

    // Parse table header row (w:tblHeader) per ECMA-376 Part 1 §17.4.49
    if (trPrObj["w:tblHeader"]) {
      row.setHeader(true);
    }

    // Parse can't split (w:cantSplit) per ECMA-376 Part 1 §17.4.5
    if (trPrObj["w:cantSplit"]) {
      row.setCantSplit(true);
    }

    // Parse row justification (w:jc) per ECMA-376 Part 1 §17.4.79
    if (trPrObj["w:jc"]) {
      const val = trPrObj["w:jc"]["@_w:val"];
      if (val) {
        row.setJustification(val as any);
      }
    }

    // Parse hidden (w:hidden) per ECMA-376 Part 1 §17.4.23
    if (trPrObj["w:hidden"]) {
      row.setHidden(true);
    }

    // Parse grid before (w:gridBefore) per ECMA-376 Part 1 §17.4.15
    if (trPrObj["w:gridBefore"]) {
      const val = parseInt(trPrObj["w:gridBefore"]["@_w:val"] || "0", 10);
      if (val > 0) {
        row.setGridBefore(val);
      }
    }

    // Parse grid after (w:gridAfter) per ECMA-376 Part 1 §17.4.14
    if (trPrObj["w:gridAfter"]) {
      const val = parseInt(trPrObj["w:gridAfter"]["@_w:val"] || "0", 10);
      if (val > 0) {
        row.setGridAfter(val);
      }
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
    if (tblPrExObj["w:tblW"]) {
      const widthVal = parseInt(tblPrExObj["w:tblW"]["@_w:w"] || "0", 10);
      if (widthVal > 0) {
        exceptions.width = widthVal;
      }
    }

    // Parse table justification exception (w:jc)
    if (tblPrExObj["w:jc"]) {
      const val = tblPrExObj["w:jc"]["@_w:val"];
      if (val) {
        exceptions.justification = val;
      }
    }

    // Parse cell spacing exception (w:tblCellSpacing)
    if (tblPrExObj["w:tblCellSpacing"]) {
      const val = parseInt(tblPrExObj["w:tblCellSpacing"]["@_w:w"] || "0", 10);
      if (val > 0) {
        exceptions.cellSpacing = val;
      }
    }

    // Parse table indentation exception (w:tblInd)
    if (tblPrExObj["w:tblInd"]) {
      const val = parseInt(tblPrExObj["w:tblInd"]["@_w:w"] || "0", 10);
      if (val > 0) {
        exceptions.indentation = val;
      }
    }

    // Parse table borders exception (w:tblBorders)
    if (tblPrExObj["w:tblBorders"]) {
      exceptions.borders = this.parseTableBordersFromObject(
        tblPrExObj["w:tblBorders"]
      );
    }

    // Parse shading exception (w:shd)
    if (tblPrExObj["w:shd"]) {
      const shdObj = tblPrExObj["w:shd"];
      exceptions.shading = {};
      if (shdObj["@_w:fill"]) exceptions.shading.fill = shdObj["@_w:fill"];
      if (shdObj["@_w:color"]) exceptions.shading.color = shdObj["@_w:color"];
      if (shdObj["@_w:val"]) exceptions.shading.pattern = shdObj["@_w:val"];
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
    const borderNames = [
      "top",
      "bottom",
      "left",
      "right",
      "insideH",
      "insideV",
    ];

    for (const name of borderNames) {
      const borderKey = `w:${name}`;
      if (bordersObj[borderKey]) {
        const borderObj = bordersObj[borderKey];
        borders[name] = {};

        if (borderObj["@_w:val"]) borders[name].style = borderObj["@_w:val"];
        if (borderObj["@_w:sz"])
          borders[name].size = parseInt(borderObj["@_w:sz"], 10);
        if (borderObj["@_w:space"])
          borders[name].space = parseInt(borderObj["@_w:space"], 10);
        if (borderObj["@_w:color"])
          borders[name].color = borderObj["@_w:color"];
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
      const tcPr = cellObj["w:tcPr"];
      if (tcPr) {
        // Parse cell width (w:tcW) with type per ECMA-376 Part 1 §17.4.81
        if (tcPr["w:tcW"]) {
          const widthVal = parseInt(tcPr["w:tcW"]["@_w:w"] || "0", 10);
          const widthType = tcPr["w:tcW"]["@_w:type"] || "dxa";
          if (widthVal > 0 || widthType === "auto") {
            cell.setWidthType(widthVal, widthType as any);
          }
        }

        // Parse conditional style (w:cnfStyle) per ECMA-376 Part 1 §17.4.7
        if (tcPr["w:cnfStyle"]) {
          const cnfStyle = tcPr["w:cnfStyle"]["@_w:val"];
          if (cnfStyle) {
            cell.setConditionalStyle(cnfStyle);
          }
        }

        // Parse cell borders (w:tcBorders)
        if (tcPr["w:tcBorders"]) {
          const bordersObj = tcPr["w:tcBorders"];
          const borders: any = {};

          const parseBorder = (borderObj: any) => {
            if (!borderObj) return undefined;
            return {
              style: borderObj["@_w:val"] || "single",
              size: borderObj["@_w:sz"]
                ? parseInt(borderObj["@_w:sz"], 10)
                : undefined,
              color: borderObj["@_w:color"] || undefined,
            };
          };

          if (bordersObj["w:top"])
            borders.top = parseBorder(bordersObj["w:top"]);
          if (bordersObj["w:bottom"])
            borders.bottom = parseBorder(bordersObj["w:bottom"]);
          if (bordersObj["w:left"])
            borders.left = parseBorder(bordersObj["w:left"]);
          if (bordersObj["w:right"])
            borders.right = parseBorder(bordersObj["w:right"]);

          if (Object.keys(borders).length > 0) {
            cell.setBorders(borders);
          }
        }

        // Parse cell shading (w:shd)
        if (tcPr["w:shd"]) {
          const shd = tcPr["w:shd"];
          const shading: any = {};
          if (shd["@_w:fill"]) shading.fill = shd["@_w:fill"];
          if (shd["@_w:color"]) shading.color = shd["@_w:color"];
          if (Object.keys(shading).length > 0) {
            cell.setShading(shading);
          }
        }

        // Parse cell margins (w:tcMar) per ECMA-376 Part 1 §17.4.43
        if (tcPr["w:tcMar"]) {
          const tcMar = tcPr["w:tcMar"];
          const margins: any = {};

          if (tcMar["w:top"]) {
            margins.top = parseInt(tcMar["w:top"]["@_w:w"] || "0", 10);
          }
          if (tcMar["w:bottom"]) {
            margins.bottom = parseInt(tcMar["w:bottom"]["@_w:w"] || "0", 10);
          }
          if (tcMar["w:left"]) {
            margins.left = parseInt(tcMar["w:left"]["@_w:w"] || "0", 10);
          }
          if (tcMar["w:right"]) {
            margins.right = parseInt(tcMar["w:right"]["@_w:w"] || "0", 10);
          }

          if (Object.keys(margins).length > 0) {
            cell.setMargins(margins);
          }
        }

        // Parse vertical alignment (w:vAlign)
        if (tcPr["w:vAlign"]) {
          const valign = tcPr["w:vAlign"]["@_w:val"];
          if (
            valign &&
            (valign === "top" || valign === "center" || valign === "bottom")
          ) {
            cell.setVerticalAlignment(valign);
          }
        }

        // Parse column span (w:gridSpan)
        if (tcPr["w:gridSpan"]) {
          const span = parseInt(tcPr["w:gridSpan"]["@_w:val"] || "1", 10);
          if (span > 1) {
            cell.setColumnSpan(span);
          }
        }

        // Parse text direction (w:textDirection) per ECMA-376 Part 1 §17.4.72
        if (tcPr["w:textDirection"]) {
          const direction = tcPr["w:textDirection"]["@_w:val"];
          if (direction) {
            cell.setTextDirection(direction as any);
          }
        }

        // Parse no wrap (w:noWrap) per ECMA-376 Part 1 §17.4.34
        if (tcPr["w:noWrap"]) {
          cell.setNoWrap(true);
        }

        // Parse hide mark (w:hideMark) per ECMA-376 Part 1 §17.4.24
        if (tcPr["w:hideMark"]) {
          cell.setHideMark(true);
        }

        // Parse fit text (w:tcFitText) per ECMA-376 Part 1 §17.4.68
        if (tcPr["w:tcFitText"]) {
          cell.setFitText(true);
        }

        // Parse vertical merge (w:vMerge) per ECMA-376 Part 1 §17.4.85
        if (tcPr["w:vMerge"]) {
          const vMergeVal = tcPr["w:vMerge"]["@_w:val"];
          // Empty element or "continue" means continue, "restart" means restart
          if (vMergeVal === "restart") {
            cell.setVerticalMerge("restart");
          } else {
            cell.setVerticalMerge("continue");
          }
        }
      }

    // Parse paragraphs in cell (w:p)
    const paragraphs = cellObj["w:p"];
    const paraChildren = Array.isArray(paragraphs)
      ? paragraphs
      : paragraphs
      ? [paragraphs]
      : [];

    // Extract paragraph XMLs from raw cell XML if available
    let paraXmls: string[] = [];
    if (rawCellXml) {
      paraXmls = XMLParser.extractElements(rawCellXml, "w:p");
    }

    for (let i = 0; i < paraChildren.length; i++) {
      const paraObj = paraChildren[i];
      const rawParaXml = i < paraXmls.length ? paraXmls[i] : undefined;
      
      // CRITICAL FIX: Use parseParagraphWithOrder to preserve revisions in table cells
      // parseParagraphFromObject doesn't scan for revisions, causing 62+ revisions to be lost
      let paragraph;
      if (rawParaXml) {
        paragraph = await this.parseParagraphWithOrder(
          rawParaXml,
          relationshipManager,
          zipHandler,
          imageManager
        );
      } else {
        paragraph = await this.parseParagraphFromObject(
          paraObj,
          relationshipManager,
          zipHandler,
          imageManager
        );
      }
      
      if (paragraph) {
        cell.addParagraph(paragraph);
      }
    }

      return cell;
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse table cell:",
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
      return null;
    }
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
      const sdtPr = sdtObj["w:sdtPr"];
      if (sdtPr) {
        // Parse ID
        const idElement = sdtPr["w:id"];
        if (idElement && idElement["@_w:val"]) {
          properties.id = parseInt(idElement["@_w:val"], 10);
        }

        // Parse tag
        const tagElement = sdtPr["w:tag"];
        if (tagElement && tagElement["@_w:val"]) {
          properties.tag = tagElement["@_w:val"];
        }

        // Parse lock
        const lockElement = sdtPr["w:lock"];
        if (lockElement && lockElement["@_w:val"]) {
          properties.lock = lockElement["@_w:val"];
        }

        // Parse alias
        const aliasElement = sdtPr["w:alias"];
        if (aliasElement && aliasElement["@_w:val"]) {
          properties.alias = aliasElement["@_w:val"];
        }

        // Parse control type from various elements
        if (sdtPr["w:richText"]) {
          properties.controlType = "richText";
        } else if (sdtPr["w:text"]) {
          properties.controlType = "plainText";
          const textElement = sdtPr["w:text"];
          properties.plainText = {
            multiLine:
              textElement?.["@_w:multiLine"] === "1" ||
              textElement?.["@_w:multiLine"] === "true",
          };
        } else if (sdtPr["w:comboBox"]) {
          properties.controlType = "comboBox";
          const comboBoxElement = sdtPr["w:comboBox"];
          properties.comboBox = this.parseListItems(comboBoxElement);
        } else if (sdtPr["w:dropDownList"]) {
          properties.controlType = "dropDownList";
          const dropDownElement = sdtPr["w:dropDownList"];
          properties.dropDownList = this.parseListItems(dropDownElement);
        } else if (sdtPr["w:date"]) {
          properties.controlType = "datePicker";
          const dateElement = sdtPr["w:date"];
          properties.datePicker = {
            dateFormat: dateElement?.["w:dateFormat"]?.["@_w:val"],
            fullDate: dateElement?.["w:fullDate"]?.["@_w:val"]
              ? new Date(dateElement["w:fullDate"]["@_w:val"])
              : undefined,
            lid: dateElement?.["w:lid"]?.["@_w:val"],
            calendar: dateElement?.["w:calendar"]?.["@_w:val"],
          };
        } else if (sdtPr["w14:checkbox"]) {
          properties.controlType = "checkbox";
          const checkboxElement = sdtPr["w14:checkbox"];
          properties.checkbox = {
            checked:
              checkboxElement?.["w14:checked"]?.["@_w14:val"] === "1" ||
              checkboxElement?.["w14:checked"]?.["@_w14:val"] === "true",
            checkedState: checkboxElement?.["w14:checkedState"]?.["@_w14:val"],
            uncheckedState:
              checkboxElement?.["w14:uncheckedState"]?.["@_w14:val"],
          };
        } else if (sdtPr["w:picture"]) {
          properties.controlType = "picture";
        } else if (sdtPr["w:docPartObj"]) {
          properties.controlType = "buildingBlock";
          const docPartObj = sdtPr["w:docPartObj"];
          properties.buildingBlock = {
            gallery: docPartObj?.["w:docPartGallery"]?.["@_w:val"],
            category: docPartObj?.["w:docPartCategory"]?.["@_w:val"],
          };
        } else if (sdtPr["w:group"]) {
          properties.controlType = "group";
        }
      }

      // Parse SDT content (sdtContent)
      const content: any[] = [];
      const sdtContent = sdtObj["w:sdtContent"];
      if (sdtContent) {
        // Check for ordered children (preserves element order)
        const orderedChildren = sdtContent["_orderedChildren"] as
          | Array<{ type: string; index: number }>
          | undefined;

        if (orderedChildren && orderedChildren.length > 0) {
          // Process in original order
          for (const childInfo of orderedChildren) {
            const elementType = childInfo.type;
            const elementIndex = childInfo.index;

            if (elementType === "w:p") {
              const paragraphs = sdtContent["w:p"];
              const paraArray = Array.isArray(paragraphs)
                ? paragraphs
                : paragraphs
                ? [paragraphs]
                : [];
              if (elementIndex < paraArray.length) {
                // Reconstruct XML for paragraph parsing
                const paraXml = this.objectToXml({
                  "w:p": paraArray[elementIndex],
                });
                const para = await this.parseParagraphWithOrder(
                  paraXml,
                  relationshipManager,
                  zipHandler,
                  imageManager
                );
                if (para) content.push(para);
              }
            } else if (elementType === "w:tbl") {
              const tables = sdtContent["w:tbl"];
              const tableArray = Array.isArray(tables)
                ? tables
                : tables
                ? [tables]
                : [];
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
            } else if (elementType === "w:sdt") {
              const sdts = sdtContent["w:sdt"];
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
          const paragraphs = sdtContent["w:p"];
          const paraArray = Array.isArray(paragraphs)
            ? paragraphs
            : paragraphs
            ? [paragraphs]
            : [];
          for (const paraObj of paraArray) {
            const paraXml = this.objectToXml({ "w:p": paraObj });
            const para = await this.parseParagraphWithOrder(
              paraXml,
              relationshipManager,
              zipHandler,
              imageManager
            );
            if (para) content.push(para);
          }

          // Parse tables
          const tables = sdtContent["w:tbl"];
          const tableArray = Array.isArray(tables)
            ? tables
            : tables
            ? [tables]
            : [];
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
          const nestedSdts = sdtContent["w:sdt"];
          const sdtArray = Array.isArray(nestedSdts)
            ? nestedSdts
            : nestedSdts
            ? [nestedSdts]
            : [];
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
      if (properties.buildingBlock?.gallery === "Table of Contents") {
        // This is a TOC - create TableOfContentsElement instead
        const toc = this.parseTOCFromSDTContent(
          content,
          properties,
          sdtContent
        );
        if (toc) {
          return new TableOfContentsElement(toc);
        }
      }

      return new StructuredDocumentTag(properties, content);
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse SDT:",
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
            if (instruction && instruction.trim().startsWith("TOC")) {
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
      "[TOC Parser] No ComplexField with TOC instruction found in assembled content"
    );
    return undefined;
  }

  /**
   * Fallback: Extracts TOC instruction from raw XML when ComplexField not available
   * Handles instructions split across multiple runs by concatenating all runs between begin/separate markers
   * @private
   */
  private extractInstructionFromRawXML(sdtContent: any): string | undefined {
    const paragraphs = sdtContent["w:p"];
    const paraArray = Array.isArray(paragraphs)
      ? paragraphs
      : paragraphs
      ? [paragraphs]
      : [];

    defaultLogger.debug(
      `[TOC Parser] Fallback: Parsing raw XML from ${paraArray.length} paragraph(s)`
    );

    // Track field state across paragraphs (TOC fields can span multiple paragraphs)
    let inField = false;
    let instructionParts: string[] = [];
    let foundTOCInstruction: string | undefined;

    for (let pIdx = 0; pIdx < paraArray.length; pIdx++) {
      const paraObj = paraArray[pIdx];
      const runs = paraObj["w:r"];
      const runArray = Array.isArray(runs) ? runs : runs ? [runs] : [];

      defaultLogger.debug(
        `[TOC Parser] Paragraph ${pIdx + 1}: ${runArray.length} runs`
      );

      for (let rIdx = 0; rIdx < runArray.length; rIdx++) {
        const runObj = runArray[rIdx];

        // Check for field character(s) - can be single object or array
        const fldChar = runObj["w:fldChar"];
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
            const charType =
              fldCharObj["@_w:fldCharType"] || fldCharObj["@_fldCharType"];

            if (!charType) {
              defaultLogger.debug(
                `[TOC Parser] Warning: fldChar without charType attribute`
              );
              continue;
            }

            defaultLogger.debug(
              `[TOC Parser] Paragraph ${pIdx + 1}, Run ${
                rIdx + 1
              }: fldChar type = "${charType}"`
            );

            if (charType === "begin") {
              inField = true;
              instructionParts = [];
              defaultLogger.debug(
                "[TOC Parser] Found field begin marker, starting instruction collection"
              );
              continue;
            }

            // Check for field end/separate marker
            if (charType === "end" || charType === "separate") {
              // If we collected instruction parts, join and check if TOC
              const fullInstruction = instructionParts.join("").trim();
              defaultLogger.debug(
                `[TOC Parser] Field ${charType} marker found. Collected instruction: "${fullInstruction.substring(
                  0,
                  100
                )}..."`
              );

              if (fullInstruction.startsWith("TOC")) {
                foundTOCInstruction = fullInstruction;
                defaultLogger.debug(
                  `[TOC Parser] ✓ Extracted complete TOC instruction from ${instructionParts.length} part(s)`
                );
                // Don't return yet - field might have more content after separate
                // But save it in case we don't find end marker
              }

              // Only reset on "end", not "separate" (there's content between separate and end)
              if (charType === "end") {
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
          const instrText = runObj["w:instrText"];
          if (instrText) {
            // Extract text value and decode XML entities
            let text = "";
            if (typeof instrText === "string") {
              text = instrText;
            } else if (instrText["#text"]) {
              text = instrText["#text"];
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

    defaultLogger.debug(
      "[TOC Parser] No TOC instruction found in raw XML fallback"
    );
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
    if (!instruction.includes("\\t")) {
      return instruction; // Already using \o or no style switch
    }

    // Extract all \t switches
    const tSwitchPattern = /\\t\s+"([^"]+)"/g;
    const matches = [...instruction.matchAll(tSwitchPattern)];
    
    if (matches.length === 0) {
      return instruction; // No \t switches found
    }

    // Parse all styles from \t switches
    const styles: Array<{ styleName: string; level: number }> = [];
    for (const match of matches) {
      const stylesStr = match[1];
      if (!stylesStr) continue;

      const parts = stylesStr.split(",").filter((p: string) => p.trim());
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
        const match = style.styleName.match(standardHeadingPattern);
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
    let normalized = instruction.replace(tSwitchPattern, "").trim();

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
          if (style === "TOCHeading") {
            // Extract title text
            const runs = element.getRuns();
            title = runs.map((r) => r.getText()).join("");
          }
        }
      }

      // NEW: Extract field instruction from assembled ComplexField objects first
      fieldInstruction = this.extractInstructionFromContent(content);

      // FALLBACK: Use raw XML parsing if no ComplexField found
      if (!fieldInstruction) {
        defaultLogger.debug(
          "[TOC Parser] ComplexField extraction failed, falling back to raw XML parsing"
        );
        fieldInstruction = this.extractInstructionFromRawXML(sdtContent);
      }

      if (!fieldInstruction) {
        defaultLogger.warn(
          "[DocumentParser] No TOC field instruction found in SDT content (tried both ComplexField and raw XML)"
        );
        return null;
      }

      defaultLogger.debug(
        `[TOC Parser] Successfully extracted instruction: "${fieldInstruction}"`
      );

      // Decode HTML entities before parsing switches
      // XML stores quotes as &quot; which need to be converted for regex matching
      const decodedInstruction = fieldInstruction
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&amp;/g, "&");

      defaultLogger.debug(
        `[TOC Parser] Decoded instruction: "${decodedInstruction}"`
      );

      // NORMALIZATION: Convert incomplete \t switches to \o format for standard headings
      let normalizedInstruction = this.normalizeTOCFieldInstruction(decodedInstruction);

      // Parse field switches from normalized instruction
      const tocOptions: any = {
        title,
        originalFieldInstruction: normalizedInstruction.trim(), // Store normalized version
      };

      // Check for \h (hyperlinks)
      if (normalizedInstruction.includes("\\h")) {
        tocOptions.useHyperlinks = true;
      }

      // Check for \n (omit page numbers)
      if (normalizedInstruction.includes("\\n")) {
        tocOptions.showPageNumbers = false;
      }

      // Check for \z (hide in web layout)
      if (normalizedInstruction.includes("\\z")) {
        tocOptions.hideInWebLayout = true;
      }

      // Check for \o "x-y" (outline levels) - supports quoted and unquoted formats
      const outlineMatch = normalizedInstruction.match(/\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/);
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
      const stylesMatch = normalizedInstruction.match(/\\t\s+"([^"]+)"/);
      if (stylesMatch && stylesMatch[1]) {
        const stylesStr = stylesMatch[1];
        const styles: Array<{ styleName: string; level: number }> = [];

        // Parse "StyleName,Level,StyleName2,Level2,..."
        const parts = stylesStr.split(",").filter((p) => p.trim());
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
    } catch (error) {
      defaultLogger.warn(
        "[DocumentParser] Failed to parse TOC from SDT content:",
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
    const listItems = element?.["w:listItem"];
    const itemArray = Array.isArray(listItems)
      ? listItems
      : listItems
      ? [listItems]
      : [];

    for (const item of itemArray) {
      if (item["@_w:displayText"] && item["@_w:value"]) {
        items.push({
          displayText: item["@_w:displayText"],
          value: item["@_w:value"],
        });
      }
    }

    return {
      items,
      lastValue: element?.["@_w:lastValue"],
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
      if (typeof o === "string") return o;
      if (typeof o !== "object") return String(o);

      const keys = Object.keys(o);

      // FIX: If a name is provided, we're building a specific element (possibly self-closing)
      // Don't return empty string for empty objects with a name - they should become self-closing tags
      if (keys.length === 0 && !name) return "";

      const tagName = name || keys[0]!; // keys[0] is guaranteed to exist due to length check (or name is provided)
      const element = name ? o : o[tagName];

      let xml = `<${tagName}`;

      // Add attributes
      if (element && typeof element === "object") {
        for (const key of Object.keys(element)) {
          if (key.startsWith("@_")) {
            const attrName = key.substring(2);
            xml += ` ${attrName}="${element[key]}"`;
          }
        }
      }

      // Check for children
      const hasChildren =
        element &&
        typeof element === "object" &&
        Object.keys(element).some(
          (k) =>
            !k.startsWith("@_") && k !== "#text" && k !== "_orderedChildren"
        );

      if (!hasChildren && (!element || !element["#text"])) {
        xml += "/>";
      } else {
        xml += ">";

        // Add text content
        if (element && element["#text"]) {
          xml += element["#text"];
        }

        // Add child elements using _orderedChildren if available
        if (element && typeof element === "object") {
          const orderedChildren = element["_orderedChildren"] as
            | Array<{ type: string; index: number }>
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
              if (
                !key.startsWith("@_") &&
                key !== "#text" &&
                key !== "_orderedChildren"
              ) {
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
    const relsPath = "word/_rels/document.xml.rels";
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
      const tagContent = XMLParser.extractBetweenTags(
        xml,
        `<${tag}`,
        `</${tag}>`
      );
      return tagContent ? XMLBuilder.unescapeXml(tagContent) : undefined;
    };

    const properties: DocumentProperties = {};

    // Parse core.xml (core properties)
    const coreXml = zipHandler.getFileAsString(DOCX_PATHS.CORE_PROPS);
    if (coreXml) {
      properties.title = extractTag(coreXml, "dc:title");
      properties.subject = extractTag(coreXml, "dc:subject");
      properties.creator = extractTag(coreXml, "dc:creator");
      properties.keywords = extractTag(coreXml, "cp:keywords");
      properties.description = extractTag(coreXml, "dc:description");
      properties.lastModifiedBy = extractTag(coreXml, "cp:lastModifiedBy");

      // Phase 5.5 - Extended core properties
      properties.category = extractTag(coreXml, "cp:category");
      properties.contentStatus = extractTag(coreXml, "cp:contentStatus");
      properties.language = extractTag(coreXml, "dc:language");

      // Parse revision as number
      const revisionStr = extractTag(coreXml, "cp:revision");
      if (revisionStr) {
        properties.revision = parseInt(revisionStr, 10);
      }

      // Parse dates
      const createdStr = extractTag(coreXml, "dcterms:created");
      if (createdStr) {
        properties.created = new Date(createdStr);
      }

      const modifiedStr = extractTag(coreXml, "dcterms:modified");
      if (modifiedStr) {
        properties.modified = new Date(modifiedStr);
      }
    }

    // Parse app.xml (extended properties)
    const appXml = zipHandler.getFileAsString(DOCX_PATHS.APP_PROPS);
    if (appXml) {
      properties.application = extractTag(appXml, "Application");
      properties.appVersion = extractTag(appXml, "AppVersion");
      properties.company = extractTag(appXml, "Company");
      properties.manager = extractTag(appXml, "Manager");

      // Also check for version field
      if (!properties.appVersion) {
        properties.version = extractTag(appXml, "Version");
      }
    }

    // Parse custom.xml (custom properties)
    const customXml = zipHandler.getFileAsString("docProps/custom.xml");
    if (customXml) {
      properties.customProperties = this.parseCustomProperties(customXml);
    }

    return properties;
  }

  /**
   * Parses custom properties from custom.xml
   */
  private parseCustomProperties(
    xml: string
  ): Record<string, string | number | boolean | Date> {
    const customProps: Record<string, string | number | boolean | Date> = {};

    // Extract all property elements
    const propertyElements = XMLParser.extractElements(xml, "property");

    for (const propXml of propertyElements) {
      // Extract name attribute
      const nameMatch = propXml.match(/name="([^"]+)"/);
      if (!nameMatch || !nameMatch[1]) continue;
      const name = XMLBuilder.unescapeXml(nameMatch[1]);

      // Determine value type and extract value
      if (propXml.includes("<vt:lpwstr>")) {
        const value = XMLParser.extractBetweenTags(
          propXml,
          "<vt:lpwstr>",
          "</vt:lpwstr>"
        );
        if (value !== undefined) {
          customProps[name] = XMLBuilder.unescapeXml(value);
        }
      } else if (propXml.includes("<vt:r8>")) {
        const value = XMLParser.extractBetweenTags(
          propXml,
          "<vt:r8>",
          "</vt:r8>"
        );
        if (value !== undefined) {
          customProps[name] = parseFloat(value);
        }
      } else if (propXml.includes("<vt:bool>")) {
        const value = XMLParser.extractBetweenTags(
          propXml,
          "<vt:bool>",
          "</vt:bool>"
        );
        if (value !== undefined) {
          customProps[name] = value === "true";
        }
      } else if (propXml.includes("<vt:filetime>")) {
        const value = XMLParser.extractBetweenTags(
          propXml,
          "<vt:filetime>",
          "</vt:filetime>"
        );
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
      const styleElements = XMLParser.extractElements(stylesXml, "w:style");

      for (const styleXml of styleElements) {
        try {
          const style = this.parseStyle(styleXml);
          if (style) {
            styles.push(style);
          }
        } catch (error) {
          const err = error instanceof Error ? error : new Error(String(error));
          this.parseErrors.push({ element: "style", error: err });

          if (this.strictParsing) {
            throw error;
          }
          // In lenient mode, continue parsing other styles
        }
      }
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: "styles.xml", error: err });

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
      const abstractNumElements = XMLParser.extractElements(
        numberingXml,
        "w:abstractNum"
      );

      for (const abstractNumXml of abstractNumElements) {
        try {
          const abstractNum = AbstractNumbering.fromXML(abstractNumXml);
          abstractNumberings.push(abstractNum);
        } catch (error) {
          const err = error instanceof Error ? error : new Error(String(error));
          this.parseErrors.push({ element: "abstractNum", error: err });

          if (this.strictParsing) {
            throw error;
          }
          // In lenient mode, continue parsing other abstract numberings
        }
      }

      // Extract all <w:num> elements (numbering instances)
      const numElements = XMLParser.extractElements(numberingXml, "w:num");

      for (const numXml of numElements) {
        try {
          const instance = NumberingInstance.fromXML(numXml);
          numberingInstances.push(instance);
        } catch (error) {
          const err = error instanceof Error ? error : new Error(String(error));
          this.parseErrors.push({ element: "num", error: err });

          if (this.strictParsing) {
            throw error;
          }
          // In lenient mode, continue parsing other instances
        }
      }
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: "numbering.xml", error: err });

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
      const bodyElements = XMLParser.extractElements(docXml, "w:body");
      if (bodyElements.length === 0) {
        return null;
      }
      const bodyContent = bodyElements[0];
      if (!bodyContent) {
        return null;
      }

      const sectPrElements = XMLParser.extractElements(bodyContent, "w:sectPr");
      if (sectPrElements.length === 0) {
        return null;
      }

      // Use the last sectPr (document-level section properties)
      const sectPr = sectPrElements[sectPrElements.length - 1];
      if (!sectPr) {
        return null;
      }

      const sectionProps: SectionProperties = {};

      // Parse page size
      const pgSzElements = XMLParser.extractElements(sectPr, "w:pgSz");
      if (pgSzElements.length > 0) {
        const pgSz = pgSzElements[0];
        if (pgSz) {
          const width = XMLParser.extractAttribute(pgSz, "w:w");
          const height = XMLParser.extractAttribute(pgSz, "w:h");
          const orient = XMLParser.extractAttribute(pgSz, "w:orient");

          if (width && height) {
            sectionProps.pageSize = {
              width: parseInt(width, 10),
              height: parseInt(height, 10),
              orientation: orient === "landscape" ? "landscape" : "portrait",
            };
          }
        }
      }

      // Parse margins
      const pgMarElements = XMLParser.extractElements(sectPr, "w:pgMar");
      if (pgMarElements.length > 0) {
        const pgMar = pgMarElements[0];
        if (pgMar) {
          const top = XMLParser.extractAttribute(pgMar, "w:top");
          const bottom = XMLParser.extractAttribute(pgMar, "w:bottom");
          const left = XMLParser.extractAttribute(pgMar, "w:left");
          const right = XMLParser.extractAttribute(pgMar, "w:right");
          const header = XMLParser.extractAttribute(pgMar, "w:header");
          const footer = XMLParser.extractAttribute(pgMar, "w:footer");
          const gutter = XMLParser.extractAttribute(pgMar, "w:gutter");

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

      // Parse columns (enhanced with separator and custom widths)
      const colsElements = XMLParser.extractElements(sectPr, "w:cols");
      if (colsElements.length > 0) {
        const cols = colsElements[0];
        if (cols) {
          const num = XMLParser.extractAttribute(cols, "w:num");
          const space = XMLParser.extractAttribute(cols, "w:space");
          const equalWidth = XMLParser.extractAttribute(cols, "w:equalWidth");
          const sep = XMLParser.extractAttribute(cols, "w:sep");

          // Extract individual column widths
          const colElements = XMLParser.extractElements(cols, "w:col");
          const columnWidths: number[] = [];
          for (const col of colElements) {
            const width = XMLParser.extractAttribute(col, "w:w");
            if (width) {
              columnWidths.push(parseInt(width.toString(), 10));
            }
          }

          // Helper to handle boolean conversion (XMLParser may return string or number)
          const toBool = (val: any) =>
            val === "1" || val === 1 || val === "true" || val === true;

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
      const typeElements = XMLParser.extractElements(sectPr, "w:type");
      if (typeElements.length > 0) {
        const type = typeElements[0];
        if (type) {
          const typeVal = XMLParser.extractAttribute(
            type,
            "w:val"
          ) as SectionType;
          if (typeVal) {
            sectionProps.type = typeVal;
          }
        }
      }

      // Parse page numbering
      const pgNumTypeElements = XMLParser.extractElements(
        sectPr,
        "w:pgNumType"
      );
      if (pgNumTypeElements.length > 0) {
        const pgNumType = pgNumTypeElements[0];
        if (pgNumType) {
          const start = XMLParser.extractAttribute(pgNumType, "w:start");
          const fmt = XMLParser.extractAttribute(
            pgNumType,
            "w:fmt"
          ) as PageNumberFormat;

          sectionProps.pageNumbering = {
            start: start ? parseInt(start, 10) : undefined,
            format: fmt,
          };
        }
      }

      // Parse title page flag
      if (XMLParser.hasSelfClosingTag(sectPr, "w:titlePg")) {
        sectionProps.titlePage = true;
      }

      // Parse header references
      const headerRefs = XMLParser.extractElements(sectPr, "w:headerReference");
      if (headerRefs.length > 0) {
        sectionProps.headers = {};
        for (const headerRef of headerRefs) {
          const type = XMLParser.extractAttribute(headerRef, "w:type");
          const rId = XMLParser.extractAttribute(headerRef, "r:id");
          if (type && rId) {
            if (type === "default") sectionProps.headers.default = rId;
            else if (type === "first") sectionProps.headers.first = rId;
            else if (type === "even") sectionProps.headers.even = rId;
          }
        }
      }

      // Parse footer references
      const footerRefs = XMLParser.extractElements(sectPr, "w:footerReference");
      if (footerRefs.length > 0) {
        sectionProps.footers = {};
        for (const footerRef of footerRefs) {
          const type = XMLParser.extractAttribute(footerRef, "w:type");
          const rId = XMLParser.extractAttribute(footerRef, "r:id");
          if (type && rId) {
            if (type === "default") sectionProps.footers.default = rId;
            else if (type === "first") sectionProps.footers.first = rId;
            else if (type === "even") sectionProps.footers.even = rId;
          }
        }
      }

      // Parse vertical alignment
      const vAlignElements = XMLParser.extractElements(sectPr, "w:vAlign");
      if (vAlignElements.length > 0) {
        const vAlign = vAlignElements[0];
        if (vAlign) {
          const val = XMLParser.extractAttribute(vAlign, "w:val");
          if (val) {
            sectionProps.verticalAlignment = val as
              | "top"
              | "center"
              | "bottom"
              | "both";
          }
        }
      }

      // Parse paper source
      const paperSrcElements = XMLParser.extractElements(sectPr, "w:paperSrc");
      if (paperSrcElements.length > 0) {
        const paperSrc = paperSrcElements[0];
        if (paperSrc) {
          const first = XMLParser.extractAttribute(paperSrc, "w:first");
          const other = XMLParser.extractAttribute(paperSrc, "w:other");

          if (first || other) {
            sectionProps.paperSource = {
              first: first ? parseInt(first.toString(), 10) : undefined,
              other: other ? parseInt(other.toString(), 10) : undefined,
            };
          }
        }
      }

      // Parse text direction
      const textDirElements = XMLParser.extractElements(
        sectPr,
        "w:textDirection"
      );
      if (textDirElements.length > 0) {
        const textDir = textDirElements[0];
        if (textDir) {
          const val = XMLParser.extractAttribute(textDir, "w:val");
          if (val) {
            sectionProps.textDirection = val as "ltr" | "rtl" | "tbRl" | "btLr";
          }
        }
      }

      return new Section(sectionProps);
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: "sectPr", error: err });

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
    const typeAttr = XMLParser.extractAttribute(
      styleXml,
      "w:type"
    ) as StyleType;
    const styleId = XMLParser.extractAttribute(styleXml, "w:styleId") || "";
    const defaultAttr = XMLParser.extractAttribute(styleXml, "w:default");
    const customStyleAttr = XMLParser.extractAttribute(
      styleXml,
      "w:customStyle"
    );

    if (!styleId || !typeAttr) {
      return null; // Invalid style, missing required attributes
    }

    // Extract style name
    const nameElement = XMLParser.extractBetweenTags(
      styleXml,
      "<w:name",
      "</w:name>"
    );
    const name = nameElement
      ? XMLParser.extractAttribute(`<w:name${nameElement}`, "w:val") || styleId
      : styleId;

    // Extract basedOn
    const basedOnElement = XMLParser.extractBetweenTags(
      styleXml,
      "<w:basedOn",
      "</w:basedOn>"
    );
    const basedOn = basedOnElement
      ? XMLParser.extractAttribute(`<w:basedOn${basedOnElement}`, "w:val")
      : undefined;

    // Extract next
    const nextElement = XMLParser.extractBetweenTags(
      styleXml,
      "<w:next",
      "</w:next>"
    );
    const next = nextElement
      ? XMLParser.extractAttribute(`<w:next${nextElement}`, "w:val")
      : undefined;

    // Parse paragraph formatting (w:pPr)
    let paragraphFormatting: ParagraphFormatting | undefined;
    const pPrXml = XMLParser.extractBetweenTags(
      styleXml,
      "<w:pPr>",
      "</w:pPr>"
    );
    if (pPrXml) {
      paragraphFormatting = this.parseParagraphFormattingFromXml(pPrXml);
    }

    // Parse run formatting (w:rPr)
    let runFormatting: RunFormatting | undefined;
    const rPrXml = XMLParser.extractBetweenTags(
      styleXml,
      "<w:rPr>",
      "</w:rPr>"
    );
    if (rPrXml) {
      runFormatting = this.parseRunFormattingFromXml(rPrXml);
    }

    // Parse metadata properties (Phase 5.3)
    // qFormat - Quick style gallery
    const qFormat =
      styleXml.includes("<w:qFormat/>") || styleXml.includes("<w:qFormat ");

    // semiHidden - Hide from recommended list
    const semiHidden =
      styleXml.includes("<w:semiHidden/>") ||
      styleXml.includes("<w:semiHidden ");

    // unhideWhenUsed - Auto-show when applied
    const unhideWhenUsed =
      styleXml.includes("<w:unhideWhenUsed/>") ||
      styleXml.includes("<w:unhideWhenUsed ");

    // locked - Prevent modification
    const locked =
      styleXml.includes("<w:locked/>") || styleXml.includes("<w:locked ");

    // personal - User-specific style
    const personal =
      styleXml.includes("<w:personal/>") || styleXml.includes("<w:personal ");

    // autoRedefine - Update style from formatting
    const autoRedefine =
      styleXml.includes("<w:autoRedefine/>") ||
      styleXml.includes("<w:autoRedefine ");

    // uiPriority - Sort order
    let uiPriority: number | undefined;
    if (styleXml.includes("<w:uiPriority")) {
      const uiPriorityStart = styleXml.indexOf("<w:uiPriority");
      const uiPriorityEnd = styleXml.indexOf("/>", uiPriorityStart);
      if (uiPriorityEnd !== -1) {
        const uiPriorityTag = styleXml.substring(
          uiPriorityStart,
          uiPriorityEnd + 2
        );
        const valStr = XMLParser.extractAttribute(uiPriorityTag, "w:val");
        if (valStr) {
          uiPriority = parseInt(valStr, 10);
        }
      }
    }

    // link - Linked character/paragraph style
    let link: string | undefined;
    if (styleXml.includes("<w:link")) {
      const linkStart = styleXml.indexOf("<w:link");
      const linkEnd = styleXml.indexOf("/>", linkStart);
      if (linkEnd !== -1) {
        const linkTag = styleXml.substring(linkStart, linkEnd + 2);
        link = XMLParser.extractAttribute(linkTag, "w:val") || undefined;
      }
    }

    // aliases - Alternative names
    let aliases: string | undefined;
    if (styleXml.includes("<w:aliases")) {
      const aliasesStart = styleXml.indexOf("<w:aliases");
      const aliasesEnd = styleXml.indexOf("/>", aliasesStart);
      if (aliasesEnd !== -1) {
        const aliasesTag = styleXml.substring(aliasesStart, aliasesEnd + 2);
        aliases = XMLParser.extractAttribute(aliasesTag, "w:val") || undefined;
      }
    }

    // Parse table style properties (Phase 5.1)
    let tableStyle:
      | import("../formatting/Style").TableStyleProperties
      | undefined;
    if (typeAttr === "table") {
      tableStyle = this.parseTableStyleProperties(styleXml);
    }

    // Create style properties
    const properties: StyleProperties = {
      styleId,
      name,
      type: typeAttr,
      basedOn,
      next,
      isDefault: defaultAttr === "1" || defaultAttr === "true",
      customStyle: customStyleAttr === "1" || customStyleAttr === "true",
      paragraphFormatting,
      runFormatting,
      tableStyle,
      // Metadata properties (Phase 5.3)
      qFormat: qFormat || undefined,
      semiHidden: semiHidden || undefined,
      unhideWhenUsed: unhideWhenUsed || undefined,
      locked: locked || undefined,
      personal: personal || undefined,
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
    const jcElement = XMLParser.extractBetweenTags(pPrXml, "<w:jc", "/>");
    if (jcElement) {
      const alignment = XMLParser.extractAttribute(
        `<w:jc${jcElement}`,
        "w:val"
      );
      if (alignment) {
        formatting.alignment = alignment as
          | "left"
          | "center"
          | "right"
          | "justify";
      }
    }

    // Parse spacing (w:spacing)
    const spacingElement = XMLParser.extractBetweenTags(
      pPrXml,
      "<w:spacing",
      "/>"
    );
    if (spacingElement) {
      const before = XMLParser.extractAttribute(
        `<w:spacing${spacingElement}`,
        "w:before"
      );
      const after = XMLParser.extractAttribute(
        `<w:spacing${spacingElement}`,
        "w:after"
      );
      const line = XMLParser.extractAttribute(
        `<w:spacing${spacingElement}`,
        "w:line"
      );
      const lineRule = XMLParser.extractAttribute(
        `<w:spacing${spacingElement}`,
        "w:lineRule"
      );

      // Validate lineRule
      let validatedLineRule: "auto" | "exact" | "atLeast" | undefined;
      if (lineRule) {
        const validLineRules = ["auto", "exact", "atLeast"];
        if (validLineRules.includes(lineRule)) {
          validatedLineRule = lineRule as "auto" | "exact" | "atLeast";
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
    const indElement = XMLParser.extractBetweenTags(pPrXml, "<w:ind", "/>");
    if (indElement) {
      const left = XMLParser.extractAttribute(`<w:ind${indElement}`, "w:left");
      const right = XMLParser.extractAttribute(
        `<w:ind${indElement}`,
        "w:right"
      );
      const firstLine = XMLParser.extractAttribute(
        `<w:ind${indElement}`,
        "w:firstLine"
      );
      const hanging = XMLParser.extractAttribute(
        `<w:ind${indElement}`,
        "w:hanging"
      );

      formatting.indentation = {
        left: left ? parseInt(left, 10) : undefined,
        right: right ? parseInt(right, 10) : undefined,
        firstLine: firstLine ? parseInt(firstLine, 10) : undefined,
        hanging: hanging ? parseInt(hanging, 10) : undefined,
      };
    }

    // Parse boolean properties
    if (pPrXml.includes("<w:keepNext/>") || pPrXml.includes("<w:keepNext ")) {
      formatting.keepNext = true;
    }
    if (pPrXml.includes("<w:keepLines/>") || pPrXml.includes("<w:keepLines ")) {
      formatting.keepLines = true;
    }
    if (
      pPrXml.includes("<w:pageBreakBefore/>") ||
      pPrXml.includes("<w:pageBreakBefore ")
    ) {
      formatting.pageBreakBefore = true;
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
    if (rPrXml.includes("<w:b/>") || rPrXml.includes("<w:b ")) {
      formatting.bold = true;
    }
    if (rPrXml.includes("<w:i/>") || rPrXml.includes("<w:i ")) {
      formatting.italic = true;
    }
    if (rPrXml.includes("<w:strike/>") || rPrXml.includes("<w:strike ")) {
      formatting.strike = true;
    }
    if (rPrXml.includes("<w:smallCaps/>") || rPrXml.includes("<w:smallCaps ")) {
      formatting.smallCaps = true;
    }
    if (rPrXml.includes("<w:caps/>") || rPrXml.includes("<w:caps ")) {
      formatting.allCaps = true;
    }

    // Parse underline - use extractSelfClosingTag for accuracy
    const uElement = XMLParser.extractSelfClosingTag(rPrXml, "w:u");
    if (uElement) {
      const uVal = XMLParser.extractAttribute(`<w:u${uElement}`, "w:val");
      if (
        uVal === "single" ||
        uVal === "double" ||
        uVal === "thick" ||
        uVal === "dotted" ||
        uVal === "dash"
      ) {
        formatting.underline = uVal as
          | "single"
          | "double"
          | "thick"
          | "dotted"
          | "dash";
      } else {
        formatting.underline = true;
      }
    }

    // Parse subscript/superscript - use extractSelfClosingTag
    const vertAlignElement = XMLParser.extractSelfClosingTag(
      rPrXml,
      "w:vertAlign"
    );
    if (vertAlignElement) {
      const val = XMLParser.extractAttribute(
        `<w:vertAlign${vertAlignElement}`,
        "w:val"
      );
      if (val === "subscript") {
        formatting.subscript = true;
      } else if (val === "superscript") {
        formatting.superscript = true;
      }
    }

    // Parse font (w:rFonts) - use extractSelfClosingTag
    const rFontsElement = XMLParser.extractSelfClosingTag(rPrXml, "w:rFonts");
    if (rFontsElement) {
      const ascii = XMLParser.extractAttribute(
        `<w:rFonts${rFontsElement}`,
        "w:ascii"
      );
      if (ascii) {
        formatting.font = ascii;
      }
    }

    // Parse size (w:sz) - size is in half-points
    // Use extractSelfClosingTag to avoid matching w:szCs
    const szElement = XMLParser.extractSelfClosingTag(rPrXml, "w:sz");
    if (szElement) {
      const val = XMLParser.extractAttribute(`<w:sz${szElement}`, "w:val");
      if (val) {
        formatting.size = parseInt(val, 10) / 2; // Convert half-points to points
      }
    }

    // Parse color (w:color)
    // Use extractSelfClosingTag to avoid matching other tags
    const colorElement = XMLParser.extractSelfClosingTag(rPrXml, "w:color");
    if (colorElement) {
      const val = XMLParser.extractAttribute(
        `<w:color${colorElement}`,
        "w:val"
      );
      if (val && val !== "auto") {
        formatting.color = val;
      }
    }

    // Parse highlight (w:highlight) - use extractSelfClosingTag
    const highlightElement = XMLParser.extractSelfClosingTag(
      rPrXml,
      "w:highlight"
    );
    if (highlightElement) {
      const val = XMLParser.extractAttribute(
        `<w:highlight${highlightElement}`,
        "w:val"
      );
      if (val) {
        const validHighlights = [
          "yellow",
          "green",
          "cyan",
          "magenta",
          "blue",
          "red",
          "darkBlue",
          "darkCyan",
          "darkGreen",
          "darkMagenta",
          "darkRed",
          "darkYellow",
          "darkGray",
          "lightGray",
          "black",
          "white",
        ];
        if (validHighlights.includes(val)) {
          formatting.highlight = val as
            | "yellow"
            | "green"
            | "cyan"
            | "magenta"
            | "blue"
            | "red"
            | "darkBlue"
            | "darkCyan"
            | "darkGreen"
            | "darkMagenta"
            | "darkRed"
            | "darkYellow"
            | "darkGray"
            | "lightGray"
            | "black"
            | "white";
        }
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
  ): import("../formatting/Style").TableStyleProperties {
    const tableStyle: import("../formatting/Style").TableStyleProperties = {};

    // Parse tblPr (table properties)
    const tblPrXml = XMLParser.extractBetweenTags(
      styleXml,
      "<w:tblPr>",
      "</w:tblPr>"
    );
    if (tblPrXml) {
      tableStyle.table = this.parseTableFormattingFromXml(tblPrXml);

      // Row band size
      if (tblPrXml.includes("<w:tblStyleRowBandSize")) {
        const tag = XMLParser.extractSelfClosingTag(
          tblPrXml,
          "w:tblStyleRowBandSize"
        );
        if (tag) {
          const val = XMLParser.extractAttribute(
            `<w:tblStyleRowBandSize${tag}`,
            "w:val"
          );
          if (val) {
            tableStyle.rowBandSize = parseInt(val, 10);
          }
        }
      }

      // Column band size
      if (tblPrXml.includes("<w:tblStyleColBandSize")) {
        const tag = XMLParser.extractSelfClosingTag(
          tblPrXml,
          "w:tblStyleColBandSize"
        );
        if (tag) {
          const val = XMLParser.extractAttribute(
            `<w:tblStyleColBandSize${tag}`,
            "w:val"
          );
          if (val) {
            tableStyle.colBandSize = parseInt(val, 10);
          }
        }
      }
    }

    // Parse tcPr (cell properties)
    const tcPrXml = XMLParser.extractBetweenTags(
      styleXml,
      "<w:tcPr>",
      "</w:tcPr>"
    );
    if (tcPrXml) {
      tableStyle.cell = this.parseTableCellFormattingFromXml(tcPrXml);
    }

    // Parse trPr (row properties)
    const trPrXml = XMLParser.extractBetweenTags(
      styleXml,
      "<w:trPr>",
      "</w:trPr>"
    );
    if (trPrXml) {
      tableStyle.row = this.parseTableRowFormattingFromXml(trPrXml);
    }

    // Parse tblStylePr (conditional formatting)
    tableStyle.conditionalFormatting =
      this.parseConditionalFormattingFromXml(styleXml);

    return tableStyle;
  }

  /**
   * Parses table formatting from tblPr XML (Phase 5.1)
   */
  private parseTableFormattingFromXml(
    tblPrXml: string
  ): import("../formatting/Style").TableStyleFormatting {
    const formatting: import("../formatting/Style").TableStyleFormatting = {};

    // Parse indent
    if (tblPrXml.includes("<w:tblInd")) {
      const tag = XMLParser.extractSelfClosingTag(tblPrXml, "w:tblInd");
      if (tag) {
        const w = XMLParser.extractAttribute(`<w:tblInd${tag}`, "w:w");
        if (w) {
          formatting.indent = parseInt(w, 10);
        }
      }
    }

    // Parse alignment
    if (tblPrXml.includes("<w:jc")) {
      const tag = XMLParser.extractSelfClosingTag(tblPrXml, "w:jc");
      if (tag) {
        const val = XMLParser.extractAttribute(`<w:jc${tag}`, "w:val");
        if (val === "left" || val === "center" || val === "right") {
          formatting.alignment = val;
        }
      }
    }

    // Parse cell spacing
    if (tblPrXml.includes("<w:tblCellSpacing")) {
      const tag = XMLParser.extractSelfClosingTag(tblPrXml, "w:tblCellSpacing");
      if (tag) {
        const w = XMLParser.extractAttribute(`<w:tblCellSpacing${tag}`, "w:w");
        if (w) {
          formatting.cellSpacing = parseInt(w, 10);
        }
      }
    }

    // Parse borders
    const bordersXml = XMLParser.extractBetweenTags(
      tblPrXml,
      "<w:tblBorders>",
      "</w:tblBorders>"
    );
    if (bordersXml) {
      formatting.borders = this.parseBordersFromXml(bordersXml, false);
    }

    // Parse shading
    if (tblPrXml.includes("<w:shd")) {
      formatting.shading = this.parseShadingFromXml(tblPrXml);
    }

    // Parse cell margins
    const marginXml = XMLParser.extractBetweenTags(
      tblPrXml,
      "<w:tblCellMar>",
      "</w:tblCellMar>"
    );
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
  ): import("../formatting/Style").TableCellStyleFormatting {
    const formatting: import("../formatting/Style").TableCellStyleFormatting =
      {};

    // Parse borders
    const bordersXml = XMLParser.extractBetweenTags(
      tcPrXml,
      "<w:tcBorders>",
      "</w:tcBorders>"
    );
    if (bordersXml) {
      formatting.borders = this.parseBordersFromXml(
        bordersXml,
        true
      ) as import("../formatting/Style").CellBorders;
    }

    // Parse shading
    if (tcPrXml.includes("<w:shd")) {
      formatting.shading = this.parseShadingFromXml(tcPrXml);
    }

    // Parse margins
    const marginXml = XMLParser.extractBetweenTags(
      tcPrXml,
      "<w:tcMar>",
      "</w:tcMar>"
    );
    if (marginXml) {
      formatting.margins = this.parseCellMarginsFromXml(marginXml);
    }

    // Parse vertical alignment
    if (tcPrXml.includes("<w:vAlign")) {
      const tag = XMLParser.extractSelfClosingTag(tcPrXml, "w:vAlign");
      if (tag) {
        const val = XMLParser.extractAttribute(`<w:vAlign${tag}`, "w:val");
        if (val === "top" || val === "center" || val === "bottom") {
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
  ): import("../formatting/Style").TableRowStyleFormatting {
    const formatting: import("../formatting/Style").TableRowStyleFormatting =
      {};

    // Parse height
    if (trPrXml.includes("<w:trHeight")) {
      const tag = XMLParser.extractSelfClosingTag(trPrXml, "w:trHeight");
      if (tag) {
        const val = XMLParser.extractAttribute(`<w:trHeight${tag}`, "w:val");
        const hRule = XMLParser.extractAttribute(
          `<w:trHeight${tag}`,
          "w:hRule"
        );
        if (val) {
          formatting.height = parseInt(val, 10);
        }
        if (hRule === "auto" || hRule === "exact" || hRule === "atLeast") {
          formatting.heightRule = hRule;
        }
      }
    }

    // Parse cantSplit
    if (
      trPrXml.includes("<w:cantSplit/>") ||
      trPrXml.includes("<w:cantSplit ")
    ) {
      formatting.cantSplit = true;
    }

    // Parse tblHeader (isHeader)
    if (
      trPrXml.includes("<w:tblHeader/>") ||
      trPrXml.includes("<w:tblHeader ")
    ) {
      formatting.isHeader = true;
    }

    return formatting;
  }

  /**
   * Parses conditional formatting from style XML (Phase 5.1)
   */
  private parseConditionalFormattingFromXml(
    styleXml: string
  ): import("../formatting/Style").ConditionalTableFormatting[] | undefined {
    const conditionalFormatting: import("../formatting/Style").ConditionalTableFormatting[] =
      [];

    // Find all tblStylePr elements
    let searchFrom = 0;
    while (true) {
      const startIdx = styleXml.indexOf("<w:tblStylePr", searchFrom);
      if (startIdx === -1) break;

      const endIdx = styleXml.indexOf("</w:tblStylePr>", startIdx);
      if (endIdx === -1) break;

      const tblStylePrXml = styleXml.substring(startIdx, endIdx + 15); // 15 = length of "</w:tblStylePr>"

      // Extract type attribute
      const typeAttr = XMLParser.extractAttribute(tblStylePrXml, "w:type");
      if (typeAttr) {
        const conditional: import("../formatting/Style").ConditionalTableFormatting =
          {
            type: typeAttr as import("../formatting/Style").ConditionalFormattingType,
          };

        // Parse pPr
        const pPrXml = XMLParser.extractBetweenTags(
          tblStylePrXml,
          "<w:pPr>",
          "</w:pPr>"
        );
        if (pPrXml) {
          conditional.paragraphFormatting =
            this.parseParagraphFormattingFromXml(pPrXml);
        }

        // Parse rPr
        const rPrXml = XMLParser.extractBetweenTags(
          tblStylePrXml,
          "<w:rPr>",
          "</w:rPr>"
        );
        if (rPrXml) {
          conditional.runFormatting = this.parseRunFormattingFromXml(rPrXml);
        }

        // Parse tblPr
        const tblPrXml = XMLParser.extractBetweenTags(
          tblStylePrXml,
          "<w:tblPr>",
          "</w:tblPr>"
        );
        if (tblPrXml) {
          conditional.tableFormatting =
            this.parseTableFormattingFromXml(tblPrXml);
        }

        // Parse tcPr
        const tcPrXml = XMLParser.extractBetweenTags(
          tblStylePrXml,
          "<w:tcPr>",
          "</w:tcPr>"
        );
        if (tcPrXml) {
          conditional.cellFormatting =
            this.parseTableCellFormattingFromXml(tcPrXml);
        }

        // Parse trPr
        const trPrXml = XMLParser.extractBetweenTags(
          tblStylePrXml,
          "<w:trPr>",
          "</w:trPr>"
        );
        if (trPrXml) {
          conditional.rowFormatting =
            this.parseTableRowFormattingFromXml(trPrXml);
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
  ):
    | import("../formatting/Style").TableBorders
    | import("../formatting/Style").CellBorders {
    const borders: any = {};

    const borderTypes = [
      "top",
      "bottom",
      "left",
      "right",
      "insideH",
      "insideV",
    ];
    for (const type of borderTypes) {
      if (bordersXml.includes(`<w:${type}`)) {
        const tag = XMLParser.extractSelfClosingTag(bordersXml, `w:${type}`);
        if (tag) {
          const style = XMLParser.extractAttribute(`<w:${type}${tag}`, "w:val");
          const size = XMLParser.extractAttribute(`<w:${type}${tag}`, "w:sz");
          const space = XMLParser.extractAttribute(
            `<w:${type}${tag}`,
            "w:space"
          );
          const color = XMLParser.extractAttribute(
            `<w:${type}${tag}`,
            "w:color"
          );

          const border: import("../formatting/Style").BorderProperties = {};
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
      const diagonalTypes = ["tl2br", "tr2bl"];
      for (const type of diagonalTypes) {
        if (bordersXml.includes(`<w:${type}`)) {
          const tag = XMLParser.extractSelfClosingTag(bordersXml, `w:${type}`);
          if (tag) {
            const style = XMLParser.extractAttribute(
              `<w:${type}${tag}`,
              "w:val"
            );
            const size = XMLParser.extractAttribute(`<w:${type}${tag}`, "w:sz");
            const space = XMLParser.extractAttribute(
              `<w:${type}${tag}`,
              "w:space"
            );
            const color = XMLParser.extractAttribute(
              `<w:${type}${tag}`,
              "w:color"
            );

            const border: import("../formatting/Style").BorderProperties = {};
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
   * Parses shading from XML (Phase 5.1)
   */
  private parseShadingFromXml(
    xml: string
  ): import("../formatting/Style").ShadingProperties | undefined {
    const tag = XMLParser.extractSelfClosingTag(xml, "w:shd");
    if (!tag) return undefined;

    const shading: import("../formatting/Style").ShadingProperties = {};
    const val = XMLParser.extractAttribute(`<w:shd${tag}`, "w:val");
    const color = XMLParser.extractAttribute(`<w:shd${tag}`, "w:color");
    const fill = XMLParser.extractAttribute(`<w:shd${tag}`, "w:fill");

    if (val) shading.val = val as any;
    if (color) shading.color = color;
    if (fill) shading.fill = fill;

    return Object.keys(shading).length > 0 ? shading : undefined;
  }

  /**
   * Parses cell margins from XML (Phase 5.1)
   */
  private parseCellMarginsFromXml(
    marginXml: string
  ): import("../formatting/Style").CellMargins | undefined {
    const margins: import("../formatting/Style").CellMargins = {};

    const marginTypes = ["top", "bottom", "left", "right"];
    for (const type of marginTypes) {
      if (marginXml.includes(`<w:${type}`)) {
        const tag = XMLParser.extractSelfClosingTag(marginXml, `w:${type}`);
        if (tag) {
          const w = XMLParser.extractAttribute(`<w:${type}${tag}`, "w:w");
          if (w) {
            margins[type as keyof import("../formatting/Style").CellMargins] =
              parseInt(w, 10);
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
      if (typeof file.content === "string") {
        return file.content;
      }

      // If Buffer, decode as UTF-8
      if (Buffer.isBuffer(file.content)) {
        return file.content.toString("utf8");
      }

      return null;
    } catch (error) {
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
  static setRawXml(
    zipHandler: ZipHandler,
    partName: string,
    xmlContent: string
  ): boolean {
    try {
      if (typeof xmlContent !== "string") {
        return false;
      }

      // Add or update the file in the ZIP handler
      // Convert string to UTF-8 Buffer for consistent encoding
      zipHandler.addFile(partName, Buffer.from(xmlContent, "utf8"), {
        binary: true,
      });
      return true;
    } catch (error) {
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
  ): Array<{
    id?: string;
    type?: string;
    target?: string;
    targetMode?: string;
  }> {
    try {
      // Construct the .rels path from the part name
      // For 'word/document.xml' -> 'word/_rels/document.xml.rels'
      const lastSlash = partName.lastIndexOf("/");
      const relsPath =
        lastSlash === -1
          ? `_rels/${partName}.rels`
          : `${partName.substring(0, lastSlash)}/_rels/${partName.substring(
              lastSlash + 1
            )}.rels`;

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
        const idMatch = attrs.match(/Id="([^"]+)"/);
        const typeMatch = attrs.match(/Type="([^"]+)"/);
        const targetMatch = attrs.match(/Target="([^"]+)"/);
        const modeMatch = attrs.match(/TargetMode="([^"]+)"/);

        if (idMatch) rel.id = idMatch[1];
        if (typeMatch) rel.type = typeMatch[1];
        if (targetMatch) rel.target = targetMatch[1];
        if (modeMatch) rel.targetMode = modeMatch[1];

        relationships.push(rel);
      }

      return relationships;
    } catch (error) {
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
    const docTagMatch = docXml.match(/<w:document([^>]+)>/);

    if (docTagMatch && docTagMatch[1]) {
      const attributes = docTagMatch[1];
      const nsPattern = /xmlns:([^=]+)="([^"]+)"/g;
      let match;

      while ((match = nsPattern.exec(attributes)) !== null) {
        if (match[1] && match[2]) {
          namespaces[`xmlns:${match[1]}`] = match[2];
        }
      }
    }

    return namespaces;
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
    headers: Array<{
      header: import("../elements/Header").Header;
      relationshipId: string;
      filename: string;
    }>;
    footers: Array<{
      footer: import("../elements/Footer").Footer;
      relationshipId: string;
      filename: string;
    }>;
  }> {
    const headers: Array<{
      header: import("../elements/Header").Header;
      relationshipId: string;
      filename: string;
    }> = [];
    const footers: Array<{
      footer: import("../elements/Footer").Footer;
      relationshipId: string;
      filename: string;
    }> = [];

    if (!section) {
      return { headers, footers };
    }

    const sectionProps = section.getProperties();

    // Parse headers
    if (sectionProps.headers) {
      for (const [type, rId] of Object.entries(sectionProps.headers)) {
        if (!rId) continue;

        // Get relationship to find header filename
        const rel = relationshipManager.getRelationship(rId);
        if (!rel) continue;

        const headerPath = `word/${rel.getTarget()}`;
        const headerXml = zipHandler.getFileAsString(headerPath);
        if (!headerXml) continue;

        // Create Header object
        const header = await this.parseHeader(
          headerXml,
          type as "default" | "first" | "even",
          zipHandler,
          relationshipManager,
          imageManager
        );
        if (header) {
          headers.push({
            header,
            relationshipId: rId,
            filename: rel.getTarget(),
          });
        }
      }
    }

    // Parse footers
    if (sectionProps.footers) {
      for (const [type, rId] of Object.entries(sectionProps.footers)) {
        if (!rId) continue;

        // Get relationship to find footer filename
        const rel = relationshipManager.getRelationship(rId);
        if (!rel) continue;

        const footerPath = `word/${rel.getTarget()}`;
        const footerXml = zipHandler.getFileAsString(footerPath);
        if (!footerXml) continue;

        // Create Footer object
        const footer = await this.parseFooter(
          footerXml,
          type as "default" | "first" | "even",
          zipHandler,
          relationshipManager,
          imageManager
        );
        if (footer) {
          footers.push({
            footer,
            relationshipId: rId,
            filename: rel.getTarget(),
          });
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
    type: "default" | "first" | "even",
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager,
    imageManager: ImageManager
  ): Promise<import("../elements/Header").Header | null> {
    try {
      const { Header } = require("../elements/Header");
      const header = new Header({ type });

      // Store raw XML for preservation when saving
      header.setRawXML(headerXml);

      // Extract w:hdr content
      const hdrContent = XMLParser.extractBetweenTags(
        headerXml,
        "<w:hdr",
        "</w:hdr>"
      );
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
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: "header", error: err });

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
    type: "default" | "first" | "even",
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager,
    imageManager: ImageManager
  ): Promise<import("../elements/Footer").Footer | null> {
    try {
      const { Footer } = require("../elements/Footer");
      const footer = new Footer({ type });

      // Store raw XML for preservation when saving
      footer.setRawXML(footerXml);

      // Extract w:ftr content
      const ftrContent = XMLParser.extractBetweenTags(
        footerXml,
        "<w:ftr",
        "</w:ftr>"
      );
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
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: "footer", error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse footer: ${err.message}`);
      }

      return null;
    }
  }
}
