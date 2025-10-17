/**
 * DocumentParser - Handles parsing of DOCX files
 * Extracts content from ZIP archives and converts XML to structured data
 */

import { ZipHandler } from '../zip/ZipHandler';
import { DOCX_PATHS } from '../zip/types';
import { Paragraph } from '../elements/Paragraph';
import { Run, RunFormatting } from '../elements/Run';
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
type BodyElement = Paragraph; // Table and TableOfContentsElement will be added later

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

    // Check for tables (not yet implemented)
    const hasTable = bodyContent.includes('<w:tbl');
    if (hasTable) {
      const err = new Error(
        'Document contains tables which are not yet fully supported in Phase 2. Tables will be ignored.'
      );
      this.parseErrors.push({ element: 'table', error: err });
      if (this.strictParsing) {
        throw err;
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
    const pPrMatch = paraXml.match(/<w:pPr[^>]*>([\s\S]*?)<\/w:pPr>/);
    if (!pPrMatch || !pPrMatch[1]) {
      return;
    }

    const pPr = pPrMatch[1];

    // Alignment
    const alignMatch = pPr.match(/<w:jc\s+w:val="([^"]+)"/);
    if (alignMatch && alignMatch[1]) {
      // Validate alignment value before applying
      const value = alignMatch[1];
      const validAlignments = ['left', 'center', 'right', 'justify'];
      if (validAlignments.includes(value)) {
        const alignment = value as 'left' | 'center' | 'right' | 'justify';
        paragraph.setAlignment(alignment);
      }
      // Invalid values are silently ignored to prevent crashes
    }

    // Style
    const styleMatch = pPr.match(/<w:pStyle\s+w:val="([^"]+)"/);
    if (styleMatch && styleMatch[1]) {
      paragraph.setStyle(styleMatch[1]);
    }

    // Indentation
    const indMatch = pPr.match(/<w:ind([^>]+)\/>/);
    if (indMatch && indMatch[1]) {
      const indStr = indMatch[1];
      const leftMatch = indStr.match(/w:left="(\d+)"/);
      const rightMatch = indStr.match(/w:right="(\d+)"/);
      const firstLineMatch = indStr.match(/w:firstLine="(\d+)"/);

      if (leftMatch && leftMatch[1]) {
        paragraph.setLeftIndent(parseInt(leftMatch[1], 10));
      }
      if (rightMatch && rightMatch[1]) {
        paragraph.setRightIndent(parseInt(rightMatch[1], 10));
      }
      if (firstLineMatch && firstLineMatch[1]) {
        paragraph.setFirstLineIndent(parseInt(firstLineMatch[1], 10));
      }
    }

    // Spacing
    const spacingMatch = pPr.match(/<w:spacing([^>]+)\/>/);
    if (spacingMatch && spacingMatch[1]) {
      const spacingStr = spacingMatch[1];
      const beforeMatch = spacingStr.match(/w:before="(\d+)"/);
      const afterMatch = spacingStr.match(/w:after="(\d+)"/);
      const lineMatch = spacingStr.match(/w:line="(\d+)"/);

      if (beforeMatch && beforeMatch[1]) {
        paragraph.setSpaceBefore(parseInt(beforeMatch[1], 10));
      }
      if (afterMatch && afterMatch[1]) {
        paragraph.setSpaceAfter(parseInt(afterMatch[1], 10));
      }
      if (lineMatch && lineMatch[1]) {
        const lineRule = spacingStr.match(/w:lineRule="([^"]+)"/);
        // Validate lineRule value before applying
        let validatedLineRule: 'auto' | 'exact' | 'atLeast' | undefined;
        if (lineRule && lineRule[1]) {
          const value = lineRule[1];
          const validLineRules = ['auto', 'exact', 'atLeast'];
          if (validLineRules.includes(value)) {
            validatedLineRule = value as 'auto' | 'exact' | 'atLeast';
          }
        }
        paragraph.setLineSpacing(parseInt(lineMatch[1], 10), validatedLineRule);
      }
    }

    // Keep properties
    if (pPr.includes('<w:keepNext')) paragraph.setKeepNext(true);
    if (pPr.includes('<w:keepLines')) paragraph.setKeepLines(true);
    if (pPr.includes('<w:pageBreakBefore')) paragraph.setPageBreakBefore(true);
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
    const rPrMatch = runXml.match(/<w:rPr[^>]*>([\s\S]*?)<\/w:rPr>/);
    if (!rPrMatch || !rPrMatch[1]) {
      return;
    }

    const rPr = rPrMatch[1];

    // Bold
    if (rPr.includes('<w:b/>') || rPr.includes('<w:b ')) {
      run.setBold(true);
    }

    // Italic
    if (rPr.includes('<w:i/>') || rPr.includes('<w:i ')) {
      run.setItalic(true);
    }

    // Underline
    const underlineMatch = rPr.match(/<w:u\s+w:val="([^"]+)"/);
    if (underlineMatch && underlineMatch[1]) {
      // Validate underline style before applying
      const value = underlineMatch[1];
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
        run.setUnderline(underlineStyle);
      }
      // Invalid values are silently ignored to prevent crashes
    } else if (rPr.includes('<w:u/>')) {
      run.setUnderline(true);
    }

    // Strike
    if (rPr.includes('<w:strike/>') || rPr.includes('<w:strike ')) {
      run.setStrike(true);
    }

    // Subscript/Superscript
    const vertAlignMatch = rPr.match(/<w:vertAlign\s+w:val="([^"]+)"/);
    if (vertAlignMatch && vertAlignMatch[1]) {
      if (vertAlignMatch[1] === 'subscript') {
        run.setSubscript(true);
      } else if (vertAlignMatch[1] === 'superscript') {
        run.setSuperscript(true);
      }
    }

    // Font
    const fontMatch = rPr.match(/<w:rFonts[^>]+w:ascii="([^"]+)"/);
    if (fontMatch && fontMatch[1]) {
      run.setFont(fontMatch[1]);
    }

    // Size (in half-points, convert to points)
    const sizeMatch = rPr.match(/<w:sz\s+w:val="(\d+)"/);
    if (sizeMatch && sizeMatch[1]) {
      const halfPoints = parseInt(sizeMatch[1], 10);
      run.setSize(halfPoints / 2);
    }

    // Color
    const colorMatch = rPr.match(/<w:color\s+w:val="([^"]+)"/);
    if (colorMatch && colorMatch[1]) {
      run.setColor(colorMatch[1]);
    }

    // Highlight
    const highlightMatch = rPr.match(/<w:highlight\s+w:val="([^"]+)"/);
    if (highlightMatch && highlightMatch[1]) {
      // Validate highlight color before applying
      const value = highlightMatch[1];
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
      // Invalid values are silently ignored to prevent crashes
    }

    // Small caps
    if (rPr.includes('<w:smallCaps/>') || rPr.includes('<w:smallCaps ')) {
      run.setSmallCaps(true);
    }

    // All caps
    if (rPr.includes('<w:caps/>') || rPr.includes('<w:caps ')) {
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
}
