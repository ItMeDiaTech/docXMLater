/**
 * DocumentParser - Handles parsing of DOCX files
 * Extracts content from ZIP archives and converts XML to structured data
 */

import { ZipHandler } from '../zip/ZipHandler';
import { DOCX_PATHS } from '../zip/types';
import { Paragraph } from '../elements/Paragraph';
import { Run, RunFormatting } from '../elements/Run';
import { Table } from '../elements/Table';
import { TableCell } from '../elements/TableCell';
import { Hyperlink } from '../elements/Hyperlink';
import { XMLBuilder } from '../xml/XMLBuilder';
import { XMLParser } from '../xml/XMLParser';
import { RelationshipManager } from './RelationshipManager';
import { DocumentProperties } from './Document';

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
type BodyElement = Paragraph | Table; // Table and TableOfContentsElement will be added later

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
   * @returns Parsed body elements, properties, and updated relationship manager
   */
  async parseDocument(
    zipHandler: ZipHandler,
    relationshipManager: RelationshipManager
  ): Promise<{
    bodyElements: BodyElement[];
    properties: DocumentProperties;
    relationshipManager: RelationshipManager;
  }> {
    // Verify the document exists
    const docXml = zipHandler.getFileAsString(DOCX_PATHS.DOCUMENT);
    if (!docXml) {
      throw new Error('Invalid document: word/document.xml not found');
    }

    // Parse existing relationships to avoid ID collisions
    const parsedRelationshipManager = this.parseRelationships(zipHandler, relationshipManager);

    // Parse document properties
    const properties = this.parseProperties(zipHandler);

    // Parse body elements (paragraphs and tables)
    const bodyElements = this.parseBodyElements(docXml, parsedRelationshipManager);

    return { bodyElements, properties, relationshipManager: parsedRelationshipManager };
  }

  /**
   * Parses body elements from document XML
   * Extracts paragraphs and tables with their formatting
   * Uses XMLParser for safe position-based parsing (prevents ReDoS)
   */
  private parseBodyElements(
    docXml: string,
    relationshipManager: RelationshipManager
  ): BodyElement[] {
    const bodyElements: BodyElement[] = [];

    // Validate input size to prevent excessive memory usage
    try {
      XMLParser.validateSize(docXml);
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'document', error: err });
      if (this.strictParsing) {
        throw err;
      }
      return bodyElements;
    }

    // Extract the body content using safe position-based parsing
    const bodyContent = XMLParser.extractBody(docXml);
    if (!bodyContent) {
      return bodyElements;
    }

    // Parse paragraphs using XMLParser (no regex backtracking)
    const paragraphXmls = XMLParser.extractElements(bodyContent, 'w:p');

    for (const paraXml of paragraphXmls) {
      const paragraph = this.parseParagraph(paraXml, relationshipManager);
      if (paragraph) {
        bodyElements.push(paragraph);
      }
    }

    // Parse tables (w:tbl elements)
    const tableRegex = /<w:tbl[^>]*>[\s\S]*?<\/w:tbl>/g;
    let tableMatch;
    while ((tableMatch = tableRegex.exec(bodyContent)) !== null) {
      try {
        const table = this.parseTable(tableMatch[0], relationshipManager);
        if (table) {
          bodyElements.push(table);
        }
      } catch (error) {
        this.parseErrors.push({
          element: 'table',
          error: error instanceof Error ? error : new Error(String(error))
        });
        if (this.strictParsing) {
          throw error;
        }
      }
    }

    // Validate that we didn't load an empty/corrupted document
    this.validateLoadedContent(bodyElements);

    return bodyElements;
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
            `${emptyRuns} out of ${totalRuns} runs (${emptyPercentage.toFixed(1)}%) have no text content. ` +
            `This may indicate:\n` +
            `  - The document was already corrupted before loading\n` +
            `  - Text content was stripped by another application\n` +
            `  - Encoding issues during document creation\n` +
            `Original document structure is preserved, but text may be lost.`
        );
        this.parseErrors.push({ element: 'document-validation', error: warning });

        // Always warn to console, even in non-strict mode
        console.warn(`\nDocXML Load Warning:\n${warning.message}\n`);
      } else if (emptyPercentage > 50 && emptyRuns > 5) {
        const warning = new Error(
          `Document has ${emptyRuns} out of ${totalRuns} runs (${emptyPercentage.toFixed(1)}%) with no text. ` +
            `This is higher than normal and may indicate partial data loss.`
        );
        this.parseErrors.push({ element: 'document-validation', error: warning });
        console.warn(`\nDocXML Load Warning:\n${warning.message}\n`);
      }
    }
  }

  /**
   * Parses a single paragraph from XML
   */
  private parseParagraph(
    paraXml: string,
    relationshipManager: RelationshipManager
  ): Paragraph | null {
    try {
      const paragraph = new Paragraph();

      // Parse paragraph properties
      this.parseParagraphProperties(paraXml, paragraph);

      // Parse hyperlinks first using XMLParser
      const hyperlinkXmls = XMLParser.extractElements(paraXml, 'w:hyperlink');

      // Add hyperlinks to paragraph
      for (const hyperlinkXml of hyperlinkXmls) {
        const hyperlink = this.parseHyperlink(hyperlinkXml, relationshipManager);
        if (hyperlink) {
          paragraph.addHyperlink(hyperlink);
        }
      }

      // Parse runs using XMLParser (safe position-based parsing)
      // Note: We need to exclude runs that are inside hyperlinks
      // Remove all hyperlink tags from the XML before extracting runs
      let paraXmlWithoutHyperlinks = paraXml;
      for (const hyperlinkXml of hyperlinkXmls) {
        paraXmlWithoutHyperlinks = paraXmlWithoutHyperlinks.replace(hyperlinkXml, '');
      }

      const runXmls = XMLParser.extractElements(paraXmlWithoutHyperlinks, 'w:r');

      // Add runs to paragraph
      for (const runXml of runXmls) {
        const run = this.parseRun(runXml);
        if (run) {
          paragraph.addRun(run);
        }
      }

      return paragraph;
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'paragraph', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse paragraph: ${err.message}`);
      }

      // In lenient mode, log warning and continue
      return null;
    }
  }

  /**
   * Parses paragraph properties and applies them
   */
  private parseParagraphProperties(paraXml: string, paragraph: Paragraph): void {
    // Use XMLParser to extract paragraph properties element
    const pPr = XMLParser.extractBetweenTags(paraXml, '<w:pPr', '</w:pPr>');
    if (!pPr) {
      return;
    }

    // Alignment - extract w:jc element and get w:val attribute
    const jcElements = XMLParser.extractElements(pPr, 'w:jc');
    if (jcElements.length > 0) {
      // @ts-ignore
      const value = XMLParser.extractAttribute(jcElements[0], 'w:val');
      if (value) {
        const validAlignments = ['left', 'center', 'right', 'justify'];
        if (validAlignments.includes(value)) {
          paragraph.setAlignment(value as 'left' | 'center' | 'right' | 'justify');
        }
      }
    }

    // Style - extract w:pStyle element and get w:val attribute
    const styleElements = XMLParser.extractElements(pPr, 'w:pStyle');
    if (styleElements.length > 0) {
      // @ts-ignore
      const styleId = XMLParser.extractAttribute(styleElements[0], 'w:val');
      if (styleId) {
        paragraph.setStyle(styleId!);
      }
    }

    // Indentation - extract w:ind element and get attributes
    const indElements = XMLParser.extractElements(pPr, 'w:ind');
    if (indElements.length > 0) {
      const indElement = indElements[0];
      // @ts-ignore
      const left = XMLParser.extractAttribute(indElement, 'w:left');
      // @ts-ignore
      const right = XMLParser.extractAttribute(indElement, 'w:right');
      // @ts-ignore
      const firstLine = XMLParser.extractAttribute(indElement, 'w:firstLine');

      if (left) {
        paragraph.setLeftIndent(parseInt(left!, 10));
      }
      if (right) {
        paragraph.setRightIndent(parseInt(right!, 10));
      }
      if (firstLine) {
        paragraph.setFirstLineIndent(parseInt(firstLine!, 10));
      }
    }

    // Spacing - extract w:spacing element and get attributes
    const spacingElements = XMLParser.extractElements(pPr, 'w:spacing');
    if (spacingElements.length > 0) {
      const spacingElement = spacingElements[0];
      // @ts-ignore
      const before = XMLParser.extractAttribute(spacingElement, 'w:before');
      // @ts-ignore
      const after = XMLParser.extractAttribute(spacingElement, 'w:after');
      // @ts-ignore
      const line = XMLParser.extractAttribute(spacingElement, 'w:line');
      // @ts-ignore
      const lineRule = XMLParser.extractAttribute(spacingElement, 'w:lineRule');

      if (before) {
        paragraph.setSpaceBefore(parseInt(before!, 10));
      }
      if (after) {
        paragraph.setSpaceAfter(parseInt(after!, 10));
      }
      if (line) {
        let validatedLineRule: 'auto' | 'exact' | 'atLeast' | undefined;
        if (lineRule) {
          const validLineRules = ['auto', 'exact', 'atLeast'];
          if (validLineRules.includes(lineRule)) {
            validatedLineRule = lineRule as 'auto' | 'exact' | 'atLeast';
          }
        }
        paragraph.setLineSpacing(parseInt(line!, 10), validatedLineRule);
      }
    }

    // Keep properties - use XMLParser helper
    if (XMLParser.hasSelfClosingTag(pPr, 'w:keepNext')) paragraph.setKeepNext(true);
    if (XMLParser.hasSelfClosingTag(pPr, 'w:keepLines')) paragraph.setKeepLines(true);
    if (XMLParser.hasSelfClosingTag(pPr, 'w:pageBreakBefore')) paragraph.setPageBreakBefore(true);

    // Contextual spacing per ECMA-376 Part 1 §17.3.1.8
    if (XMLParser.hasSelfClosingTag(pPr, 'w:contextualSpacing')) {
      paragraph.setContextualSpacing(true);
    }
  }

  /**
   * Parses a single run from XML
   */
  private parseRun(runXml: string): Run | null {
    try {
      // Extract text content using XMLParser (safe parsing)
      const text = XMLBuilder.unescapeXml(XMLParser.extractText(runXml));

      // Create run with text
      const run = new Run(text);

      // Parse run properties
      this.parseRunProperties(runXml, run);

      return run;
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'run', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse run: ${err.message}`);
      }

      // In lenient mode, log warning and continue
      return null;
    }
  }

  /**
   * Parses run properties and applies them
   */
  private parseRunProperties(runXml: string, run: Run): void {
    // Use XMLParser to extract run properties element
    const rPr = XMLParser.extractBetweenTags(runXml, '<w:rPr', '</w:rPr>');
    if (!rPr) {
      return;
    }

    // Bold - use XMLParser helper
    if (XMLParser.hasSelfClosingTag(rPr, 'w:b')) {
      run.setBold(true);
    }

    // Italic - use XMLParser helper
    if (XMLParser.hasSelfClosingTag(rPr, 'w:i')) {
      run.setItalic(true);
    }

    // Underline - extract w:u element
    const underlineElements = XMLParser.extractElements(rPr, 'w:u');
    if (underlineElements.length > 0) {
      // @ts-ignore
      const value = XMLParser.extractAttribute(underlineElements[0], 'w:val');
      if (value) {
        const validUnderlineStyles = [
          'single',
          'double',
          'thick',
          'dotted',
          'dash',
          'dotDash',
          'dotDotDash',
          'wave',
        ];
        if (
          validUnderlineStyles.includes(value) ||
          value === 'true' ||
          value === 'false'
        ) {
          const underlineStyle = value as RunFormatting['underline'];
          run.setUnderline(underlineStyle!);
        }
      } else {
        // Self-closing <w:u/> means single underline
        run.setUnderline(true);
      }
    }

    // Strike - use XMLParser helper
    if (XMLParser.hasSelfClosingTag(rPr, 'w:strike')) {
      run.setStrike(true);
    }

    // Subscript/Superscript - extract w:vertAlign element
    const vertAlignElements = XMLParser.extractElements(rPr, 'w:vertAlign');
    if (vertAlignElements.length > 0) {
      // @ts-ignore
      const value = XMLParser.extractAttribute(vertAlignElements[0], 'w:val');
      if (value === 'subscript') {
        run.setSubscript(true);
      } else if (value === 'superscript') {
        run.setSuperscript(true);
      }
    }

    // Font - extract w:rFonts element
    const fontElements = XMLParser.extractElements(rPr, 'w:rFonts');
    if (fontElements.length > 0) {
      // @ts-ignore
      const fontName = XMLParser.extractAttribute(fontElements[0], 'w:ascii');
      if (fontName) {
        run.setFont(fontName!);
      }
    }

    // Size (in half-points, convert to points) - extract w:sz element
    const sizeElements = XMLParser.extractElements(rPr, 'w:sz');
    if (sizeElements.length > 0) {
      // @ts-ignore
      const halfPoints = XMLParser.extractAttribute(sizeElements[0], 'w:val');
      if (halfPoints) {
        run.setSize(parseInt(halfPoints!, 10) / 2);
      }
    }

    // Color - extract w:color element
    const colorElements = XMLParser.extractElements(rPr, 'w:color');
    if (colorElements.length > 0) {
      // @ts-ignore
      const colorValue = XMLParser.extractAttribute(colorElements[0], 'w:val');
      if (colorValue) {
        run.setColor(colorValue!);
      }
    }

    // Highlight - extract w:highlight element
    const highlightElements = XMLParser.extractElements(rPr, 'w:highlight');
    if (highlightElements.length > 0) {
      // @ts-ignore
      const value = XMLParser.extractAttribute(highlightElements[0], 'w:val');
      if (value) {
        const validHighlightColors = [
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
        if (validHighlightColors.includes(value)) {
          const highlightColor = value as RunFormatting['highlight'];
          run.setHighlight(highlightColor);
        }
      }
    }

    // Small caps - use XMLParser helper
    if (XMLParser.hasSelfClosingTag(rPr, 'w:smallCaps')) {
      run.setSmallCaps(true);
    }

    // All caps - use XMLParser helper
    if (XMLParser.hasSelfClosingTag(rPr, 'w:caps')) {
      run.setAllCaps(true);
    }
  }

  /**
   * Parses a single hyperlink from XML
   * Extracts URL from relationships, anchor for internal links, and text formatting
   */
  private parseHyperlink(
    hyperlinkXml: string,
    relationshipManager: RelationshipManager
  ): Hyperlink | null {
    try {
      // Extract hyperlink attributes using XMLParser
      const relationshipId = XMLParser.extractAttribute(hyperlinkXml, 'r:id');
      const anchor = XMLParser.extractAttribute(hyperlinkXml, 'w:anchor');
      const tooltip = XMLParser.extractAttribute(hyperlinkXml, 'w:tooltip');

      // Hyperlink must have either a relationship ID (external) or anchor (internal)
      if (!relationshipId && !anchor) {
        return null;
      }

      // Extract runs within the hyperlink for text and formatting
      const runXmls = XMLParser.extractElements(hyperlinkXml, 'w:r');
      let text = '';
      let formatting: RunFormatting | undefined;

      for (const runXml of runXmls) {
        // Accumulate text from all runs
        text += XMLBuilder.unescapeXml(XMLParser.extractText(runXml));

        // Get formatting from first run
        if (!formatting) {
          const run = this.parseRun(runXml);
          if (run) {
            formatting = run.getFormatting();
          }
        }
      }

      // Resolve URL from relationship manager for external links
      let url: string | undefined;
      if (relationshipId) {
        const relationship = relationshipManager.getRelationship(relationshipId);
        if (relationship && relationship.getType().includes('hyperlink')) {
          url = relationship.getTarget();
        }
      }

      // Create hyperlink with extracted properties
      // Improved fallback: text → url → anchor → 'Link'
      return new Hyperlink({
        url,
        anchor,
        text: text || url || anchor || 'Link',
        formatting,
        tooltip,
        relationshipId,
      });
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'hyperlink', error: err });

      if (this.strictParsing) {
        throw new Error(`Failed to parse hyperlink: ${err.message}`);
      }

      // In lenient mode, log warning and continue
      return null;
    }
  }

  /**
   * Parses a table element from XML
   * Extracts rows and cells with their content
   */
  private parseTable(tableXml: string, relationshipManager: RelationshipManager): Table | null {
    try {
      // Extract table rows (w:tr)
      const rowRegex = /<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g;
      const rows: Array<Array<string>> = [];
      let rowMatch;

      while ((rowMatch = rowRegex.exec(tableXml)) !== null) {
        const rowContent = rowMatch[1] || '';
        const cells: string[] = [];

        // Extract cells (w:tc)
        const cellRegex = /<w:tc[^>]*>([\s\S]*?)<\/w:tc>/g;
        let cellMatch;
        while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
          cells.push(cellMatch[1] || '');
        }

        if (cells.length > 0) {
          rows.push(cells);
        }
      }

      if (rows.length === 0) {
        return null; // No valid rows found
      }

      // Create table with row/column dimensions from first row
      const firstRow = rows[0];
      if (!firstRow) {
        return null;
      }

      const colCount = firstRow.length;
      const rowCount = rows.length;

      if (colCount === 0) {
        return null;
      }

      const table = new Table(rowCount, colCount);

      // Populate table cells with content
      for (let rIdx = 0; rIdx < rows.length; rIdx++) {
        const row = table.getRows()[rIdx];
        if (!row) continue;

        const cells = row.getCells();
        const cellContents = rows[rIdx];
        if (!cellContents) continue;

        for (let cIdx = 0; cIdx < cellContents.length && cIdx < cells.length; cIdx++) {
          const cell = cells[cIdx];
          if (cell instanceof TableCell) {
            // Parse paragraphs inside cell
            const cellXml = cellContents[cIdx] || '';
            const paraRegex = /<w:p[^>]*>([\s\S]*?)<\/w:p>/g;
            let paraMatch;

            while ((paraMatch = paraRegex.exec(cellXml)) !== null) {
              const paraContent = paraMatch[1] || '';
              const para = this.parseParagraph(`<w:p>${paraContent}</w:p>`, relationshipManager);
              if (para) {
                cell.addParagraph(para);
              }
            }
          }
        }
      }

      return table;
    } catch (error) {
      if (this.strictParsing) {
        throw error;
      }
      return null;
    }
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
   * Parses document properties from core.xml
   */
  private parseProperties(zipHandler: ZipHandler): DocumentProperties {
    const coreXml = zipHandler.getFileAsString(DOCX_PATHS.CORE_PROPS);
    if (!coreXml) {
      return {};
    }

    // Simple XML parsing using regex (sufficient for Phase 3)
    const extractTag = (xml: string, tag: string): string | undefined => {
      const match = xml.match(new RegExp(`<${tag}[^>]*>([^<]*)</${tag}>`));
      return match && match[1] ? XMLBuilder.unescapeXml(match[1]) : undefined;
    };

    const properties: DocumentProperties = {
      title: extractTag(coreXml, 'dc:title'),
      subject: extractTag(coreXml, 'dc:subject'),
      creator: extractTag(coreXml, 'dc:creator'),
      keywords: extractTag(coreXml, 'cp:keywords'),
      description: extractTag(coreXml, 'dc:description'),
      lastModifiedBy: extractTag(coreXml, 'cp:lastModifiedBy'),
    };

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

    return properties;
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
  static setRawXml(zipHandler: ZipHandler, partName: string, xmlContent: string): boolean {
    try {
      if (typeof xmlContent !== 'string') {
        return false;
      }

      // Add or update the file in the ZIP handler
      // Convert string to UTF-8 Buffer for consistent encoding
      zipHandler.addFile(partName, Buffer.from(xmlContent, 'utf8'), { binary: true });
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
  static getRelationships(zipHandler: ZipHandler, partName: string): Array<{ id?: string; type?: string; target?: string; targetMode?: string }> {
    try {
      // Construct the .rels path from the part name
      // For 'word/document.xml' -> 'word/_rels/document.xml.rels'
      const lastSlash = partName.lastIndexOf('/');
      const relsPath = lastSlash === -1
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
}
