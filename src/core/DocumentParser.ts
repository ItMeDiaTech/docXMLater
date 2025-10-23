/**
 * DocumentParser - Handles parsing of DOCX files
 * Extracts content from ZIP archives and converts XML to structured data
 */

import { ZipHandler } from "../zip/ZipHandler";
import { DOCX_PATHS } from "../zip/types";
import { Paragraph, ParagraphFormatting } from "../elements/Paragraph";
import { Run, RunFormatting } from "../elements/Run";
import { Hyperlink } from "../elements/Hyperlink";
import { Table } from "../elements/Table";
import { TableRow } from "../elements/TableRow";
import { TableCell } from "../elements/TableCell";
import { TableOfContentsElement } from "../elements/TableOfContentsElement";
import { StructuredDocumentTag } from "../elements/StructuredDocumentTag";
import { ImageManager } from "../elements/ImageManager";
import { ImageRun } from "../elements/ImageRun";
import {
  Section,
  SectionProperties,
  SectionType,
  PageNumberFormat,
} from "../elements/Section";
import { XMLBuilder } from "../xml/XMLBuilder";
import { XMLParser } from "../xml/XMLParser";
import { RelationshipManager } from "./RelationshipManager";
import { DocumentProperties } from "./Document";
import { Style, StyleProperties, StyleType } from "../formatting/Style";
import { AbstractNumbering } from "../formatting/AbstractNumbering";
import { NumberingInstance } from "../formatting/NumberingInstance";

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
    // Verify the document exists
    const docXml = zipHandler.getFileAsString(DOCX_PATHS.DOCUMENT);
    if (!docXml) {
      throw new Error("Invalid document: word/document.xml not found");
    }

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

    // Parse numbering from numbering.xml
    const numbering = this.parseNumbering(zipHandler);

    // Parse section properties from document.xml
    const section = this.parseSectionProperties(docXml);

    // Parse and preserve namespaces from the root <w:document> tag
    const namespaces = this.parseNamespaces(docXml);

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
            // Parse XML to object, then extract the paragraph content
            // XMLParser.parseToObject returns { "w:p": { ... } }, we need just the content
            // IMPORTANT: trimValues: false preserves whitespace from xml:space="preserve" attributes
            const parsed = XMLParser.parseToObject(elementXml, { trimValues: false });
            const paragraph = await this.parseParagraphFromObject(
              parsed["w:p"],
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
            const parsed = XMLParser.parseToObject(elementXml, { trimValues: false });
            const table = await this.parseTableFromObject(
              parsed["w:tbl"],
              relationshipManager,
              zipHandler,
              imageManager
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
            const parsed = XMLParser.parseToObject(elementXml, { trimValues: false });
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
   */
  private extractSingleElement(
    content: string,
    tagName: string,
    startPos: number
  ): string {
    const openTag = `<${tagName}`;
    const closeTag = `</${tagName}>`;
    const selfClosingEnd = "/>";

    // Verify we're at the right position
    if (!content.substring(startPos).startsWith(openTag)) {
      return "";
    }

    // Find the end of the opening tag
    const openEnd = content.indexOf(">", startPos);
    if (openEnd === -1) {
      return "";
    }

    // Check if it's a self-closing tag
    if (content.substring(openEnd - 1, openEnd + 1) === selfClosingEnd) {
      return content.substring(startPos, openEnd + 1);
    }

    // Find the matching closing tag (with depth tracking for nested elements)
    let depth = 1;
    let pos = openEnd + 1;

    while (pos < content.length && depth > 0) {
      const nextOpen = content.indexOf(openTag, pos);
      const nextClose = content.indexOf(closeTag, pos);

      if (nextClose === -1) {
        // No closing tag found
        return "";
      }

      if (nextOpen !== -1 && nextOpen < nextClose) {
        // Found another opening tag before the closing tag
        // Verify it's an exact match (not a prefix like <w:pPr>)
        const charAfter = content[nextOpen + openTag.length];
        if (
          charAfter === ">" ||
          charAfter === "/" ||
          charAfter === " " ||
          charAfter === "\t" ||
          charAfter === "\n" ||
          charAfter === "\r"
        ) {
          depth++;
          pos = nextOpen + openTag.length;
        } else {
          // Prefix match, skip it
          pos = nextOpen + openTag.length;
        }
      } else {
        // Found the closing tag
        depth--;
        pos = nextClose + closeTag.length;
        if (depth === 0) {
          return content.substring(startPos, pos);
        }
      }
    }

    return "";
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
        console.warn(`\nDocXML Load Warning:\n${warning.message}\n`);
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
        console.warn(`\nDocXML Load Warning:\n${warning.message}\n`);
      }
    }
  }

  /**
   * Extracts paragraph children (runs and hyperlinks) in document order
   * This ensures we preserve the original ordering of text and hyperlinks
   */

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

      // Parse children (runs, hyperlinks, and bookmarks) in document order
      // This preserves the original ordering of all elements

      // Handle runs (w:r)
      const runs = paraObj["w:r"];
      const runChildren = Array.isArray(runs) ? runs : (runs ? [runs] : []);

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
      const hyperlinkChildren = Array.isArray(hyperlinks) ? hyperlinks : (hyperlinks ? [hyperlinks] : []);

      for (const hyperlinkObj of hyperlinkChildren) {
        const hyperlink = this.parseHyperlinkFromObject(hyperlinkObj, relationshipManager);
        if (hyperlink) {
          paragraph.addHyperlink(hyperlink);
        }
      }

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
      if (ind["@_w:left"]) paragraph.setLeftIndent(parseInt(ind["@_w:left"], 10));
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
        if (borderObj["@_w:sz"]) border.size = parseInt(borderObj["@_w:sz"], 10);
        if (borderObj["@_w:color"]) border.color = borderObj["@_w:color"];
        if (borderObj["@_w:space"]) border.space = parseInt(borderObj["@_w:space"], 10);
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
        : tabsObj["w:tab"] ? [tabsObj["w:tab"]] : [];

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
      if (widowControlVal === "0" || widowControlVal === "false" || widowControlVal === false || widowControlVal === 0) {
        paragraph.setWidowControl(false);
      } else {
        // If w:val is "1", "true", true, 1, or undefined (element present without val), default to true
        paragraph.setWidowControl(true);
      }
    }

    // Outline level per ECMA-376 Part 1 §17.3.1.19
    if (pPrObj["w:outlineLvl"] !== undefined && pPrObj["w:outlineLvl"]["@_w:val"] !== undefined) {
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
      if (bidiVal === "0" || bidiVal === "false" || bidiVal === false || bidiVal === 0) {
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
      if (adjustRightIndVal === "0" || adjustRightIndVal === "false" || adjustRightIndVal === false || adjustRightIndVal === 0) {
        paragraph.setAdjustRightInd(false);
      } else {
        // Default is true when element present without val attribute or val="1"
        paragraph.setAdjustRightInd(true);
      }
    }
  }

  private parseRunFromObject(runObj: any): Run | null {
    try {
      // Extract text content from w:t element
      // XMLParser.parseToObject() returns: { "w:t": { "#text": "actual text", "@_xml:space": "preserve" } }
      // So we need to access the #text property for the actual text content
      const textElement = runObj["w:t"];
      let text =
        (typeof textElement === 'object' && textElement !== null)
          ? (textElement["#text"] || "")
          : (textElement || "");

      // Unescape XML entities (&lt; → <, &amp; → &, etc.)
      text = XMLBuilder.unescapeXml(text);

      const run = new Run(text, { cleanXmlFromText: false });

      this.parseRunPropertiesFromObject(runObj["w:rPr"], run);

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
      const runChildren = Array.isArray(runs) ? runs : (runs ? [runs] : []);

      // Extract text from all runs
      const text = runChildren
        .map((runObj: any) => {
          const textElement = runObj["w:t"];
          let runText =
            (typeof textElement === 'object' && textElement !== null)
              ? (textElement["#text"] || "")
              : (textElement || "");
          return XMLBuilder.unescapeXml(runText);
        })
        .join('');

      // Parse formatting from first run if present
      let formatting: RunFormatting = {};
      if (runChildren.length > 0 && runChildren[0]["w:rPr"]) {
        const tempRun = new Run('');
        this.parseRunPropertiesFromObject(runChildren[0]["w:rPr"], tempRun);
        formatting = tempRun.getFormatting();
      }

      // Resolve URL from relationship if external hyperlink
      let url: string | undefined;
      if (relationshipId) {
        const relationship = relationshipManager.getRelationship(relationshipId);
        if (relationship) {
          url = relationship.getTarget();
        }
      }

      // Create hyperlink with all properties including relationshipId
      // Use constructor directly to preserve relationship ID through save/load cycles
      const hyperlink = new Hyperlink({
        url,
        anchor,
        text: text || url || anchor || 'Link',
        formatting,
        tooltip,
        relationshipId,
      });

      return hyperlink;
    } catch (error) {
      console.warn('[DocumentParser] Failed to parse hyperlink:', error);
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

    // Parse RTL text (w:rtl) per ECMA-376 Part 1 §17.3.2.30
    if (rPrObj["w:rtl"]) run.setRTL(true);

    if (rPrObj["w:b"]) run.setBold(true);
    if (rPrObj["w:bCs"]) run.setComplexScriptBold(true);
    if (rPrObj["w:i"]) run.setItalic(true);
    if (rPrObj["w:iCs"]) run.setComplexScriptItalic(true);
    if (rPrObj["w:strike"]) run.setStrike(true);
    if (rPrObj["w:smallCaps"]) run.setSmallCaps(true);
    if (rPrObj["w:caps"]) run.setAllCaps(true);

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
      if (layoutObj["@_w:id"] !== undefined) layout.id = Number(layoutObj["@_w:id"]);
      if (layoutObj["@_w:vert"]) layout.vert = true;
      if (layoutObj["@_w:vertCompress"]) layout.vertCompress = true;
      if (layoutObj["@_w:combine"]) layout.combine = true;
      if (layoutObj["@_w:combineBrackets"]) layout.combineBrackets = layoutObj["@_w:combineBrackets"];

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
      // For now, we'll focus on wp:inline which is more common
      const inlineObj = drawingObj["wp:inline"];
      if (!inlineObj) {
        // Could be wp:anchor for floating images, but we'll skip those for now
        return null;
      }

      // Extract dimensions from wp:extent
      const extentObj = inlineObj["wp:extent"];
      let width = 0;
      let height = 0;
      if (extentObj) {
        width = parseInt(extentObj["@_cx"] || "0", 10);
        height = parseInt(extentObj["@_cy"] || "0", 10);
      }

      // Extract name and description from wp:docPr
      const docPrObj = inlineObj["wp:docPr"];
      let name = "image";
      let description = "Image";
      if (docPrObj) {
        name = docPrObj["@_name"] || "image";
        description = docPrObj["@_descr"] || "Image";
      }

      // Navigate through the graphic structure to find the relationship ID
      // Structure: a:graphic → a:graphicData → pic:pic → pic:blipFill → a:blip
      const graphicObj = inlineObj["a:graphic"];
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

      // Extract relationship ID (r:embed)
      const relationshipId = blipObj["@_r:embed"];
      if (!relationshipId) {
        return null;
      }

      // Get the image from the relationship
      const relationship = relationshipManager.getRelationship(relationshipId);
      if (!relationship) {
        console.warn(`[DocumentParser] Image relationship not found: ${relationshipId}`);
        return null;
      }

      const imageTarget = relationship.getTarget();
      if (!imageTarget) {
        console.warn(`[DocumentParser] Image relationship has no target: ${relationshipId}`);
        return null;
      }

      // Read image data from zip
      const imagePath = `word/${imageTarget}`;
      const imageData = zipHandler.getFileAsBuffer(imagePath);
      if (!imageData) {
        console.warn(`[DocumentParser] Image file not found: ${imagePath}`);
        return null;
      }

      // Detect image extension from path
      const extension = imagePath.split('.').pop()?.toLowerCase() || 'png';

      // Create image from buffer with name and description
      const { Image: ImageClass } = await import('../elements/Image');
      const image = await ImageClass.create({
        source: imageData,
        width,
        height,
        name,
        description,
      });

      // Register image with ImageManager (reuse existing relationship)
      imageManager.registerImage(image, relationshipId);
      image.setRelationshipId(relationshipId);

      // Create and return ImageRun
      return new ImageRun(image);
    } catch (error) {
      console.warn('[DocumentParser] Failed to parse drawing:', error);
      return null;
    }
  }

  private async parseTableFromObject(
    tableObj: any,
    relationshipManager: RelationshipManager,
    zipHandler: ZipHandler,
    imageManager: ImageManager
  ): Promise<Table | null> {
    try {
      // Create empty table
      const table = new Table();

      // Parse table rows (w:tr)
      const rows = tableObj["w:tr"];
      const rowChildren = Array.isArray(rows) ? rows : (rows ? [rows] : []);

      for (const rowObj of rowChildren) {
        const row = await this.parseTableRowFromObject(
          rowObj,
          relationshipManager,
          zipHandler,
          imageManager
        );
        if (row) {
          table.addRow(row);
        }
      }

      return table;
    } catch (error) {
      console.warn('[DocumentParser] Failed to parse table:', error);
      return null;
    }
  }

  private async parseTableRowFromObject(
    rowObj: any,
    relationshipManager: RelationshipManager,
    zipHandler: ZipHandler,
    imageManager: ImageManager
  ): Promise<TableRow | null> {
    try {
      // Create empty row
      const row = new TableRow();

      // Parse table cells (w:tc)
      const cells = rowObj["w:tc"];
      const cellChildren = Array.isArray(cells) ? cells : (cells ? [cells] : []);

      for (const cellObj of cellChildren) {
        const cell = await this.parseTableCellFromObject(
          cellObj,
          relationshipManager,
          zipHandler,
          imageManager
        );
        if (cell) {
          row.addCell(cell);
        }
      }

      return row;
    } catch (error) {
      console.warn('[DocumentParser] Failed to parse table row:', error);
      return null;
    }
  }

  private async parseTableCellFromObject(
    cellObj: any,
    relationshipManager: RelationshipManager,
    zipHandler: ZipHandler,
    imageManager: ImageManager
  ): Promise<TableCell | null> {
    try {
      // Create empty cell
      const cell = new TableCell();

      // Parse cell properties (w:tcPr) per ECMA-376 Part 1 §17.4.42
      const tcPr = cellObj["w:tcPr"];
      if (tcPr) {
        // Parse cell width (w:tcW)
        if (tcPr["w:tcW"]) {
          const widthVal = parseInt(tcPr["w:tcW"]["@_w:w"] || "0", 10);
          if (widthVal > 0) {
            cell.setWidth(widthVal);
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
              size: borderObj["@_w:sz"] ? parseInt(borderObj["@_w:sz"], 10) : undefined,
              color: borderObj["@_w:color"] || undefined,
            };
          };

          if (bordersObj["w:top"]) borders.top = parseBorder(bordersObj["w:top"]);
          if (bordersObj["w:bottom"]) borders.bottom = parseBorder(bordersObj["w:bottom"]);
          if (bordersObj["w:left"]) borders.left = parseBorder(bordersObj["w:left"]);
          if (bordersObj["w:right"]) borders.right = parseBorder(bordersObj["w:right"]);

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
          if (valign && (valign === "top" || valign === "center" || valign === "bottom")) {
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
      }

      // Parse paragraphs in cell (w:p)
      const paragraphs = cellObj["w:p"];
      const paraChildren = Array.isArray(paragraphs) ? paragraphs : (paragraphs ? [paragraphs] : []);

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

      return cell;
    } catch (error) {
      console.warn('[DocumentParser] Failed to parse table cell:', error);
      return null;
    }
  }

  private async parseSDTFromObject(
    _sdtObj: any,
    _relationshipManager: RelationshipManager,
    _zipHandler: ZipHandler,
    _imageManager: ImageManager
  ): Promise<StructuredDocumentTag | null> {
    // TODO: Implement SDT parsing from object
    return null;
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
   * Parses document properties from core.xml
   */
  private parseProperties(zipHandler: ZipHandler): DocumentProperties {
    const coreXml = zipHandler.getFileAsString(DOCX_PATHS.CORE_PROPS);
    if (!coreXml) {
      return {};
    }

    // Extract document properties using XMLParser
    const extractTag = (xml: string, tag: string): string | undefined => {
      const tagContent = XMLParser.extractBetweenTags(
        xml,
        `<${tag}`,
        `</${tag}>`
      );
      return tagContent ? XMLBuilder.unescapeXml(tagContent) : undefined;
    };

    const properties: DocumentProperties = {
      title: extractTag(coreXml, "dc:title"),
      subject: extractTag(coreXml, "dc:subject"),
      creator: extractTag(coreXml, "dc:creator"),
      keywords: extractTag(coreXml, "cp:keywords"),
      description: extractTag(coreXml, "dc:description"),
      lastModifiedBy: extractTag(coreXml, "cp:lastModifiedBy"),
    };

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

    return properties;
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

      // Parse columns
      const colsElements = XMLParser.extractElements(sectPr, "w:cols");
      if (colsElements.length > 0) {
        const cols = colsElements[0];
        if (cols) {
          const num = XMLParser.extractAttribute(cols, "w:num");
          const space = XMLParser.extractAttribute(cols, "w:space");
          const equalWidth = XMLParser.extractAttribute(cols, "w:equalWidth");

          if (num) {
            sectionProps.columns = {
              count: parseInt(num, 10),
              space: space ? parseInt(space, 10) : undefined,
              equalWidth: equalWidth === "1" || equalWidth === "true",
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
   * @param docXml - The full XML content of word/document.xml
   * @returns A record of namespace prefixes to their URIs
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
          namespaces[match[1]] = match[2];
        }
      }

      // Extract default namespace (without prefix)
      const defaultNsMatch = attributes.match(/xmlns="([^"]+)"/);
      if (defaultNsMatch && defaultNsMatch[1]) {
        namespaces["xmlns"] = defaultNsMatch[1];
      }
    }

    return namespaces;
  }
}
