/**
 * Document - High-level API for creating and managing Word documents
 * Provides a simple interface for creating DOCX files without managing ZIP and XML manually
 */

import { ZipHandler } from '../zip/ZipHandler';
import { DOCX_PATHS } from '../zip/types';
import { Paragraph } from '../elements/Paragraph';
import { Table } from '../elements/Table';
import { Section } from '../elements/Section';
import { Image } from '../elements/Image';
import { ImageManager } from '../elements/ImageManager';
import { Header } from '../elements/Header';
import { Footer } from '../elements/Footer';
import { HeaderFooterManager } from '../elements/HeaderFooterManager';
import { Hyperlink } from '../elements/Hyperlink';
import { TableOfContents } from '../elements/TableOfContents';
import { TableOfContentsElement } from '../elements/TableOfContentsElement';
import { Bookmark } from '../elements/Bookmark';
import { BookmarkManager } from '../elements/BookmarkManager';
import { Revision, RevisionType } from '../elements/Revision';
import { RevisionManager } from '../elements/RevisionManager';
import { Comment } from '../elements/Comment';
import { CommentManager } from '../elements/CommentManager';
import { Run, RunFormatting } from '../elements/Run';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { XMLParser } from '../xml/XMLParser';
import { StylesManager } from '../formatting/StylesManager';
import { Style } from '../formatting/Style';
import { NumberingManager } from '../formatting/NumberingManager';
import { RelationshipManager } from './RelationshipManager';

/**
 * Document properties
 */
export interface DocumentProperties {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  revision?: number;
  created?: Date;
  modified?: Date;
}

/**
 * Document creation options
 */
export interface DocumentOptions {
  properties?: DocumentProperties;
  /** Maximum memory usage percentage (0-100) before throwing error. Default: 80 */
  maxMemoryUsagePercent?: number;
  /** Strict parsing mode - throw errors instead of collecting warnings. Default: false */
  strictParsing?: boolean;
}

/**
 * Body element - can be a Paragraph, Table, or TableOfContentsElement
 */
type BodyElement = Paragraph | Table | TableOfContentsElement;

/**
 * Renderable interface for objects that can generate XML
 * Ensures type safety for paragraph content
 */
interface Renderable {
  toXML(): XMLElement;
}

/**
 * ImageRun - A run that contains an image (drawing)
 * Implements Renderable interface for type-safe paragraph content
 */
class ImageRun implements Renderable {
  constructor(private image: Image) {}

  toXML(): XMLElement {
    const drawing = this.image.toXML();
    return {
      name: 'w:r',
      children: [drawing]
    };
  }
}

/**
 * Represents a Word document
 */
export class Document {
  private zipHandler: ZipHandler;
  private bodyElements: BodyElement[] = [];
  private properties: DocumentProperties;
  private stylesManager: StylesManager;
  private numberingManager: NumberingManager;
  private section: Section;
  private imageManager: ImageManager;
  private relationshipManager: RelationshipManager;
  private headerFooterManager: HeaderFooterManager;
  private bookmarkManager: BookmarkManager;
  private revisionManager: RevisionManager;
  private commentManager: CommentManager;
  private maxMemoryUsagePercent: number;
  private strictParsing: boolean;
  private parseErrors: Array<{ element: string; error: Error }> = [];

  /**
   * Private constructor - use Document.create() or Document.load()
   * @param zipHandler Optional ZIP handler (for loading existing documents)
   * @param options Document options
   * @param initDefaults Whether to initialize with default relationships (false for loaded docs)
   */
  private constructor(zipHandler?: ZipHandler, options: DocumentOptions = {}, initDefaults: boolean = true) {
    this.zipHandler = zipHandler || new ZipHandler();
    this.properties = options.properties || {};
    this.maxMemoryUsagePercent = options.maxMemoryUsagePercent ?? 80;
    this.strictParsing = options.strictParsing ?? false;
    this.stylesManager = StylesManager.create();
    this.numberingManager = NumberingManager.create();
    this.section = Section.createLetter(); // Default Letter-sized section
    this.imageManager = ImageManager.create();
    this.relationshipManager = RelationshipManager.create();
    this.headerFooterManager = HeaderFooterManager.create();
    this.bookmarkManager = BookmarkManager.create();
    this.revisionManager = RevisionManager.create();
    this.commentManager = CommentManager.create();

    // Add default relationships only for new documents
    if (initDefaults) {
      this.relationshipManager.addStyles();
      this.relationshipManager.addNumbering();
    }
  }

  /**
   * Creates a new empty document
   * @param options - Document options
   * @returns New Document instance
   */
  static create(options?: DocumentOptions): Document {
    const doc = new Document(undefined, options);
    doc.initializeRequiredFiles();
    return doc;
  }

  /**
   * Loads a document from a file
   * @param filePath - Path to the DOCX file
   * @param options - Document options
   * @returns Document instance
   */
  static async load(filePath: string, options?: DocumentOptions): Promise<Document> {
    const zipHandler = new ZipHandler();
    await zipHandler.load(filePath);

    // Create document without default relationships (will parse from file)
    const doc = new Document(zipHandler, options, false);
    await doc.parseDocument();

    return doc;
  }

  /**
   * Loads a document from a buffer
   * @param buffer - DOCX file buffer
   * @param options - Document options
   * @returns Document instance
   */
  static async loadFromBuffer(buffer: Buffer, options?: DocumentOptions): Promise<Document> {
    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer);

    // Create document without default relationships (will parse from file)
    const doc = new Document(zipHandler, options, false);
    await doc.parseDocument();

    return doc;
  }

  /**
   * Initializes all required DOCX files with minimal valid content
   */
  private initializeRequiredFiles(): void {
    // [Content_Types].xml
    this.zipHandler.addFile(DOCX_PATHS.CONTENT_TYPES, this.generateContentTypes());

    // _rels/.rels
    this.zipHandler.addFile(DOCX_PATHS.RELS, this.generateRels());

    // word/document.xml (will be updated when saving)
    this.zipHandler.addFile(DOCX_PATHS.DOCUMENT, this.generateDocumentXml());

    // word/_rels/document.xml.rels
    this.zipHandler.addFile('word/_rels/document.xml.rels', this.relationshipManager.generateXml());

    // word/styles.xml
    this.zipHandler.addFile(DOCX_PATHS.STYLES, this.stylesManager.generateStylesXml());

    // word/numbering.xml
    this.zipHandler.addFile(DOCX_PATHS.NUMBERING, this.numberingManager.generateNumberingXml());

    // docProps/core.xml
    this.zipHandler.addFile(DOCX_PATHS.CORE_PROPS, this.generateCoreProps());

    // docProps/app.xml
    this.zipHandler.addFile(DOCX_PATHS.APP_PROPS, this.generateAppProps());
  }

  /**
   * Parses the document XML and extracts paragraphs, runs, and tables
   */
  private async parseDocument(): Promise<void> {
    // Verify the document exists
    const docXml = this.zipHandler.getFileAsString(DOCX_PATHS.DOCUMENT);
    if (!docXml) {
      throw new Error('Invalid document: word/document.xml not found');
    }

    // Parse existing relationships to avoid ID collisions
    this.parseRelationships();

    // Parse document properties
    this.parseProperties();

    // Parse body elements (paragraphs and tables)
    this.parseBodyElements(docXml);
  }

  /**
   * Parses body elements from document XML
   * Extracts paragraphs and tables with their formatting
   * Uses XMLParser for safe position-based parsing (prevents ReDoS)
   */
  private parseBodyElements(docXml: string): void {
    this.bodyElements = [];

    // Validate input size to prevent excessive memory usage
    try {
      XMLParser.validateSize(docXml);
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.parseErrors.push({ element: 'document', error: err });
      if (this.strictParsing) {
        throw err;
      }
      return;
    }

    // Extract the body content using safe position-based parsing
    const bodyContent = XMLParser.extractBody(docXml);
    if (!bodyContent) {
      return;
    }

    // Parse paragraphs using XMLParser (no regex backtracking)
    const paragraphXmls = XMLParser.extractElements(bodyContent, 'w:p');

    for (const paraXml of paragraphXmls) {
      const paragraph = this.parseParagraph(paraXml);
      if (paragraph) {
        this.bodyElements.push(paragraph);
      }
    }

    // Check for tables (not yet implemented)
    const hasTable = bodyContent.includes('<w:tbl');
    if (hasTable) {
      const err = new Error('Document contains tables which are not yet fully supported in Phase 2. Tables will be ignored.');
      this.parseErrors.push({ element: 'table', error: err });
      if (this.strictParsing) {
        throw err;
      }
    }

    // Validate that we didn't load an empty/corrupted document
    this.validateLoadedContent();
  }

  /**
   * Validates loaded content to detect corrupted or empty documents
   * Adds warnings if the document appears to have lost text content
   */
  private validateLoadedContent(): void {
    const paragraphs = this.bodyElements.filter((el): el is Paragraph => el instanceof Paragraph);

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
  private parseParagraph(paraXml: string): Paragraph | null {
    try {
      const paragraph = new Paragraph();

      // Parse paragraph properties
      this.parseParagraphProperties(paraXml, paragraph);

      // Parse hyperlinks first using XMLParser
      const hyperlinkXmls = XMLParser.extractElements(paraXml, 'w:hyperlink');

      // Add hyperlinks to paragraph
      for (const hyperlinkXml of hyperlinkXmls) {
        const hyperlink = this.parseHyperlink(hyperlinkXml);
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
      const alignment = alignMatch[1] as 'left' | 'center' | 'right' | 'justify';
      paragraph.setAlignment(alignment);
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
        paragraph.setLineSpacing(
          parseInt(lineMatch[1], 10),
          lineRule && lineRule[1] ? lineRule[1] as 'auto' | 'exact' | 'atLeast' : undefined
        );
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
      // Accept underline style from XML (cast to expected type union)
      const underlineStyle = underlineMatch[1] as RunFormatting['underline'];
      run.setUnderline(underlineStyle);
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
      // Accept highlight color from XML (cast to expected type union)
      const highlightColor = highlightMatch[1] as RunFormatting['highlight'];
      run.setHighlight(highlightColor);
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
  private parseHyperlink(hyperlinkXml: string): Hyperlink | null {
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
        const relationship = this.relationshipManager.getRelationship(relationshipId);
        if (relationship && relationship.getType().includes('hyperlink')) {
          url = relationship.getTarget();
        }
      }

      // Create hyperlink with extracted properties
      return new Hyperlink({
        url,
        anchor,
        text: text || 'Link', // Default text if empty
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
   */
  private parseRelationships(): void {
    const relsPath = 'word/_rels/document.xml.rels';
    const relsXml = this.zipHandler.getFileAsString(relsPath);

    if (relsXml) {
      // Parse and replace the relationship manager with populated one
      this.relationshipManager = RelationshipManager.fromXml(relsXml);
    } else {
      // No existing relationships - add defaults
      this.relationshipManager.addStyles();
      this.relationshipManager.addNumbering();
    }
  }

  /**
   * Parses document properties from core.xml
   */
  private parseProperties(): void {
    const coreXml = this.zipHandler.getFileAsString(DOCX_PATHS.CORE_PROPS);
    if (!coreXml) {
      return;
    }

    // Simple XML parsing using regex (sufficient for Phase 3)
    const extractTag = (xml: string, tag: string): string | undefined => {
      const match = xml.match(new RegExp(`<${tag}[^>]*>([^<]*)</${tag}>`));
      return match && match[1] ? XMLBuilder.unescapeXml(match[1]) : undefined;
    };

    this.properties = {
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
      this.properties.revision = parseInt(revisionStr, 10);
    }

    // Parse dates
    const createdStr = extractTag(coreXml, 'dcterms:created');
    if (createdStr) {
      this.properties.created = new Date(createdStr);
    }

    const modifiedStr = extractTag(coreXml, 'dcterms:modified');
    if (modifiedStr) {
      this.properties.modified = new Date(modifiedStr);
    }
  }


  /**
   * Adds a paragraph to the document
   * @param paragraph - Paragraph to add
   * @returns This document for chaining
   */
  addParagraph(paragraph: Paragraph): this {
    this.bodyElements.push(paragraph);
    return this;
  }

  /**
   * Creates and adds a new paragraph with text
   * @param text - Text content
   * @param formatting - Paragraph and run formatting
   * @returns The created paragraph
   */
  createParagraph(text?: string): Paragraph {
    const para = new Paragraph();
    if (text) {
      para.addText(text);
    }
    this.bodyElements.push(para);
    return para;
  }

  /**
   * Adds a table to the document
   * @param table - Table to add
   * @returns This document for chaining
   */
  addTable(table: Table): this {
    this.bodyElements.push(table);
    return this;
  }

  /**
   * Creates and adds a new table
   * @param rows - Number of rows
   * @param columns - Number of columns
   * @returns The created table
   */
  createTable(rows: number, columns: number): Table {
    const table = new Table(rows, columns);
    this.bodyElements.push(table);
    return table;
  }

  /**
   * Adds a Table of Contents to the document
   * @param toc - TableOfContents or TableOfContentsElement to add
   * @returns This document for chaining
   */
  addTableOfContents(toc?: TableOfContents | TableOfContentsElement): this {
    // Wrap in TableOfContentsElement if plain TableOfContents provided
    const tocElement = toc instanceof TableOfContentsElement
      ? toc
      : new TableOfContentsElement(toc || TableOfContents.createStandard());

    this.bodyElements.push(tocElement);
    return this;
  }

  /**
   * Creates and adds a standard Table of Contents
   * @param title - Optional custom title
   * @returns This document for chaining
   */
  createTableOfContents(title?: string): this {
    const tocElement = TableOfContentsElement.createStandard(title);
    return this.addTableOfContents(tocElement);
  }

  /**
   * Gets all paragraphs in the document
   * @returns Array of paragraphs
   */
  getParagraphs(): Paragraph[] {
    return this.bodyElements.filter((el): el is Paragraph => el instanceof Paragraph);
  }

  /**
   * Gets all tables in the document
   * @returns Array of tables
   */
  getTables(): Table[] {
    return this.bodyElements.filter((el): el is Table => el instanceof Table);
  }

  /**
   * Gets all Table of Contents elements in the document
   * @returns Array of TableOfContentsElement
   */
  getTableOfContentsElements(): TableOfContentsElement[] {
    return this.bodyElements.filter((el): el is TableOfContentsElement => el instanceof TableOfContentsElement);
  }

  /**
   * Gets all body elements (paragraphs and tables)
   * @returns Array of body elements
   */
  getBodyElements(): BodyElement[] {
    return [...this.bodyElements];
  }

  /**
   * Gets the number of paragraphs
   * @returns Number of paragraphs
   */
  getParagraphCount(): number {
    return this.getParagraphs().length;
  }

  /**
   * Gets the number of tables
   * @returns Number of tables
   */
  getTableCount(): number {
    return this.getTables().length;
  }

  /**
   * Removes all paragraphs and tables
   * @returns This document for chaining
   */
  clearParagraphs(): this {
    this.bodyElements = [];
    return this;
  }

  /**
   * Sets document properties
   * @param properties - Document properties
   * @returns This document for chaining
   */
  setProperties(properties: DocumentProperties): this {
    this.properties = { ...this.properties, ...properties };
    return this;
  }

  /**
   * Gets document properties
   * @returns Document properties
   */
  getProperties(): DocumentProperties {
    return { ...this.properties };
  }

  /**
   * Validates that the document has meaningful content before saving
   * Warns if the document appears to be empty or corrupted
   */
  private validateBeforeSave(): void {
    const paragraphs = this.getParagraphs();

    if (paragraphs.length === 0) {
      console.warn(
        '\nDocXML Save Warning:\n' +
        'Document has no paragraphs. You are saving an empty document.\n'
      );
      return;
    }

    // Count runs with text
    let totalRuns = 0;
    let emptyRuns = 0;

    for (const para of paragraphs) {
      const runs = para.getRuns();
      totalRuns += runs.length;

      for (const run of runs) {
        if (run.getText().length === 0) {
          emptyRuns++;
        }
      }
    }

    if (totalRuns > 0) {
      const emptyPercentage = (emptyRuns / totalRuns) * 100;

      if (emptyPercentage > 90 && emptyRuns > 10) {
        console.warn(
          '\nDocXML Save Warning:\n' +
          `You are about to save a document where ${emptyRuns} out of ${totalRuns} runs (${emptyPercentage.toFixed(1)}%) are empty.\n` +
          'This may result in a document with no visible text content.\n' +
          'If this is unintentional, please review the document before saving.\n'
        );
      }
    }
  }

  /**
   * Saves the document to a file
   * @param filePath - Output file path
   */
  async save(filePath: string): Promise<void> {
    // Use atomic save pattern: save to temp file, then rename
    // This prevents partial/corrupted saves if operation fails mid-way
    const tempPath = `${filePath}.tmp.${Date.now()}`;

    try {
      // Validate before saving to prevent data loss
      this.validateBeforeSave();

      // Check memory usage before starting
      this.checkMemoryThreshold();

      // Load all image data before saving (now async)
      await this.imageManager.loadAllImageData();

      // Check memory again after loading images
      this.checkMemoryThreshold();

      // Check document size and warn if too large
      const sizeInfo = this.estimateSize();
      if (sizeInfo.warning) {
        console.warn(`DocXML Warning: ${sizeInfo.warning}`);
      }

      this.processHyperlinks();
      this.updateDocumentXml();
      this.updateStylesXml();
      this.updateNumberingXml();
      this.updateCoreProps();
      this.saveImages();
      this.saveHeaders();
      this.saveFooters();
      this.saveComments();
      this.updateRelationships();
      this.updateContentTypesWithImagesHeadersFootersAndComments();

      // Save to temporary file first
      await this.zipHandler.save(tempPath);

      // Atomic rename - only if save succeeded
      const { promises: fs } = await import('fs');
      await fs.rename(tempPath, filePath);
    } catch (error) {
      // Cleanup temporary file on error
      try {
        const { promises: fs } = await import('fs');
        await fs.unlink(tempPath);
      } catch {
        // Ignore cleanup errors
      }
      throw error; // Re-throw original error
    } finally {
      // Release image data to free memory
      this.imageManager.releaseAllImageData();
    }
  }

  /**
   * Generates the document as a buffer
   * @returns Document buffer
   */
  async toBuffer(): Promise<Buffer> {
    try {
      // Validate before saving to prevent data loss
      this.validateBeforeSave();

      // Check memory usage before starting
      this.checkMemoryThreshold();

      // Load all image data before saving (now async)
      await this.imageManager.loadAllImageData();

      // Check memory again after loading images
      this.checkMemoryThreshold();

      // Check document size and warn if too large
      const sizeInfo = this.estimateSize();
      if (sizeInfo.warning) {
        console.warn(`DocXML Warning: ${sizeInfo.warning}`);
      }

      this.processHyperlinks();
      this.updateDocumentXml();
      this.updateStylesXml();
      this.updateNumberingXml();
      this.updateCoreProps();
      this.saveImages();
      this.saveHeaders();
      this.saveFooters();
      this.saveComments();
      this.updateRelationships();
      this.updateContentTypesWithImagesHeadersFootersAndComments();
      return await this.zipHandler.toBuffer();
    } finally {
      // Release image data to free memory
      this.imageManager.releaseAllImageData();
    }
  }

  /**
   * Updates the document.xml file with current paragraphs
   */
  private updateDocumentXml(): void {
    const xml = this.generateDocumentXml();
    this.zipHandler.updateFile(DOCX_PATHS.DOCUMENT, xml);
  }

  /**
   * Updates the core properties with current values
   */
  private updateCoreProps(): void {
    const xml = this.generateCoreProps();
    this.zipHandler.updateFile(DOCX_PATHS.CORE_PROPS, xml);
  }

  /**
   * Updates the styles.xml file with current styles
   */
  private updateStylesXml(): void {
    const xml = this.stylesManager.generateStylesXml();
    this.zipHandler.updateFile(DOCX_PATHS.STYLES, xml);
  }

  /**
   * Updates the numbering.xml file with current numbering definitions
   */
  private updateNumberingXml(): void {
    const xml = this.numberingManager.generateNumberingXml();
    this.zipHandler.updateFile(DOCX_PATHS.NUMBERING, xml);
  }

  /**
   * Generates [Content_Types].xml
   */
  private generateContentTypes(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
  }

  /**
   * Generates _rels/.rels
   */
  private generateRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  }

  /**
   * Generates word/document.xml with current body elements
   */
  private generateDocumentXml(): string {
    const bodyXmls: XMLElement[] = [];

    // Generate XML for each body element
    // Note: TableOfContentsElement.toXML() returns an array
    for (const element of this.bodyElements) {
      const xml = element.toXML();
      if (Array.isArray(xml)) {
        // TableOfContentsElement returns array of XMLElements
        bodyXmls.push(...xml);
      } else {
        // Paragraph and Table return single XMLElement
        bodyXmls.push(xml);
      }
    }

    // Add section properties at the end
    bodyXmls.push(this.section.toXML());
    return XMLBuilder.createDocument(bodyXmls);
  }

  /**
   * Generates docProps/core.xml
   */
  private generateCoreProps(): string {
    const now = new Date();
    const created = this.properties.created || now;
    const modified = this.properties.modified || now;

    const formatDate = (date: Date): string => {
      return date.toISOString();
    };

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${XMLBuilder.escapeXmlText(this.properties.title || '')}</dc:title>
  <dc:subject>${XMLBuilder.escapeXmlText(this.properties.subject || '')}</dc:subject>
  <dc:creator>${XMLBuilder.escapeXmlText(this.properties.creator || 'DocXML')}</dc:creator>
  <cp:keywords>${XMLBuilder.escapeXmlText(this.properties.keywords || '')}</cp:keywords>
  <dc:description>${XMLBuilder.escapeXmlText(this.properties.description || '')}</dc:description>
  <cp:lastModifiedBy>${XMLBuilder.escapeXmlText(this.properties.lastModifiedBy || this.properties.creator || 'DocXML')}</cp:lastModifiedBy>
  <cp:revision>${this.properties.revision || 1}</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">${formatDate(created)}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${formatDate(modified)}</dcterms:modified>
</cp:coreProperties>`;
  }

  /**
   * Generates docProps/app.xml
   */
  private generateAppProps(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>DocXML</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>0.1.0</AppVersion>
</Properties>`;
  }


  /**
   * Gets the StylesManager
   * @returns StylesManager instance
   */
  getStylesManager(): StylesManager {
    return this.stylesManager;
  }

  /**
   * Adds a style to the document
   * @param style - Style to add
   * @returns This document for chaining
   */
  addStyle(style: Style): this {
    this.stylesManager.addStyle(style);
    return this;
  }

  /**
   * Gets a style by ID
   * @param styleId - Style ID
   * @returns The style, or undefined if not found
   */
  getStyle(styleId: string): Style | undefined {
    return this.stylesManager.getStyle(styleId);
  }

  /**
   * Checks if a style exists
   * @param styleId - Style ID
   * @returns True if the style exists
   */
  hasStyle(styleId: string): boolean {
    return this.stylesManager.hasStyle(styleId);
  }

  /**
   * Gets the underlying ZipHandler (for advanced use cases)
   * @returns ZipHandler instance
   */
  getZipHandler(): ZipHandler {
    return this.zipHandler;
  }

  /**
   * Gets the NumberingManager
   * @returns NumberingManager instance
   */
  getNumberingManager(): NumberingManager {
    return this.numberingManager;
  }

  /**
   * Creates a new bullet list and returns its numId
   * @param levels - Number of levels (default: 3)
   * @param bullets - Array of bullet characters
   * @returns The numId to use with setNumbering()
   */
  createBulletList(levels: number = 3, bullets?: string[]): number {
    return this.numberingManager.createBulletList(levels, bullets);
  }

  /**
   * Creates a new numbered list and returns its numId
   * @param levels - Number of levels (default: 3)
   * @param formats - Array of formats for each level
   * @returns The numId to use with setNumbering()
   */
  createNumberedList(
    levels: number = 3,
    formats?: Array<'decimal' | 'lowerLetter' | 'lowerRoman'>
  ): number {
    return this.numberingManager.createNumberedList(levels, formats);
  }

  /**
   * Creates a new multi-level list and returns its numId
   * @returns The numId to use with setNumbering()
   */
  createMultiLevelList(): number {
    return this.numberingManager.createMultiLevelList();
  }

  /**
   * Gets the document section
   * @returns The section
   */
  getSection(): Section {
    return this.section;
  }

  /**
   * Sets the document section
   * @param section The section to set
   * @returns This document for chaining
   */
  setSection(section: Section): this {
    this.section = section;
    return this;
  }

  /**
   * Sets page size
   * @param width Width in twips
   * @param height Height in twips
   * @param orientation Page orientation
   * @returns This document for chaining
   */
  setPageSize(width: number, height: number, orientation?: 'portrait' | 'landscape'): this {
    this.section.setPageSize(width, height, orientation);
    return this;
  }

  /**
   * Sets page orientation
   * @param orientation Page orientation
   * @returns This document for chaining
   */
  setPageOrientation(orientation: 'portrait' | 'landscape'): this {
    this.section.setOrientation(orientation);
    return this;
  }

  /**
   * Sets margins
   * @param margins Margin properties
   * @returns This document for chaining
   */
  setMargins(margins: { top: number; bottom: number; left: number; right: number; header?: number; footer?: number; gutter?: number }): this {
    this.section.setMargins(margins);
    return this;
  }

  /**
   * Sets the default header for the document
   * @param header The header to set
   * @returns This document for chaining
   */
  setHeader(header: Header): this {
    // Generate relationship for header
    const relationship = this.relationshipManager.addHeader(`${header.getFilename(1)}`);

    // Register with manager
    this.headerFooterManager.registerHeader(header, relationship.getId());

    // Link to section
    this.section.setHeaderReference('default', relationship.getId());

    return this;
  }

  /**
   * Sets the first page header
   * @param header The header to set
   * @returns This document for chaining
   */
  setFirstPageHeader(header: Header): this {
    // Enable title page
    this.section.setTitlePage(true);

    // Generate relationship for header
    const relationship = this.relationshipManager.addHeader(`${header.getFilename(this.headerFooterManager.getHeaderCount() + 1)}`);

    // Register with manager
    this.headerFooterManager.registerHeader(header, relationship.getId());

    // Link to section
    this.section.setHeaderReference('first', relationship.getId());

    return this;
  }

  /**
   * Sets the even page header (requires different odd/even pages)
   * @param header The header to set
   * @returns This document for chaining
   */
  setEvenPageHeader(header: Header): this {
    // Generate relationship for header
    const relationship = this.relationshipManager.addHeader(`${header.getFilename(this.headerFooterManager.getHeaderCount() + 1)}`);

    // Register with manager
    this.headerFooterManager.registerHeader(header, relationship.getId());

    // Link to section
    this.section.setHeaderReference('even', relationship.getId());

    return this;
  }

  /**
   * Sets the default footer for the document
   * @param footer The footer to set
   * @returns This document for chaining
   */
  setFooter(footer: Footer): this {
    // Generate relationship for footer
    const relationship = this.relationshipManager.addFooter(`${footer.getFilename(1)}`);

    // Register with manager
    this.headerFooterManager.registerFooter(footer, relationship.getId());

    // Link to section
    this.section.setFooterReference('default', relationship.getId());

    return this;
  }

  /**
   * Sets the first page footer
   * @param footer The footer to set
   * @returns This document for chaining
   */
  setFirstPageFooter(footer: Footer): this {
    // Enable title page
    this.section.setTitlePage(true);

    // Generate relationship for footer
    const relationship = this.relationshipManager.addFooter(`${footer.getFilename(this.headerFooterManager.getFooterCount() + 1)}`);

    // Register with manager
    this.headerFooterManager.registerFooter(footer, relationship.getId());

    // Link to section
    this.section.setFooterReference('first', relationship.getId());

    return this;
  }

  /**
   * Sets the even page footer (requires different odd/even pages)
   * @param footer The footer to set
   * @returns This document for chaining
   */
  setEvenPageFooter(footer: Footer): this {
    // Generate relationship for footer
    const relationship = this.relationshipManager.addFooter(`${footer.getFilename(this.headerFooterManager.getFooterCount() + 1)}`);

    // Register with manager
    this.headerFooterManager.registerFooter(footer, relationship.getId());

    // Link to section
    this.section.setFooterReference('even', relationship.getId());

    return this;
  }

  /**
   * Gets the HeaderFooterManager
   * @returns HeaderFooterManager instance
   */
  getHeaderFooterManager(): HeaderFooterManager {
    return this.headerFooterManager;
  }

  /**
   * Adds an image to the document inside a paragraph
   * @param image The image to add
   * @returns This document for chaining
   */
  addImage(image: Image): this {
    // Generate relationship ID
    const target = `media/image${this.imageManager.getImageCount() + 1}.${image.getExtension()}`;
    const relationship = this.relationshipManager.addImage(target);

    // Register image with manager
    this.imageManager.registerImage(image, relationship.getId());

    // Create a paragraph containing the image
    const para = new Paragraph();
    // Add image as a run (images use w:drawing in w:r)
    const imageRun = this.createImageRun(image);
    // Cast to Run for type compatibility (ImageRun implements same toXML interface)
    para.addRun(imageRun as any as Run);

    this.bodyElements.push(para);
    return this;
  }

  /**
   * Creates a run containing an image
   * @param image The image
   * @returns A run with the image
   */
  private createImageRun(image: Image): Renderable {
    // Create a special run that contains the drawing
    // Implements Renderable interface for type safety
    return new ImageRun(image);
  }

  /**
   * Gets the ImageManager
   * @returns ImageManager instance
   */
  getImageManager(): ImageManager {
    return this.imageManager;
  }

  /**
   * Gets the RelationshipManager
   * @returns RelationshipManager instance
   */
  getRelationshipManager(): RelationshipManager {
    return this.relationshipManager;
  }

  /**
   * Processes all hyperlinks in paragraphs and registers them with RelationshipManager
   */
  private processHyperlinks(): void {
    // Get all paragraphs (from body and from headers/footers)
    const paragraphs = this.getParagraphs();

    // Also check headers and footers
    const headers = this.headerFooterManager.getAllHeaders();
    const footers = this.headerFooterManager.getAllFooters();

    for (const header of headers) {
      for (const element of header.header.getElements()) {
        if (element instanceof Paragraph) {
          this.processHyperlinksInParagraph(element);
        }
      }
    }

    for (const footer of footers) {
      for (const element of footer.footer.getElements()) {
        if (element instanceof Paragraph) {
          this.processHyperlinksInParagraph(element);
        }
      }
    }

    // Process body paragraphs
    for (const para of paragraphs) {
      this.processHyperlinksInParagraph(para);
    }
  }

  /**
   * Processes hyperlinks in a single paragraph
   */
  private processHyperlinksInParagraph(paragraph: Paragraph): void {
    const content = paragraph.getContent();

    for (const item of content) {
      if (item instanceof Hyperlink && item.isExternal() && !item.getRelationshipId()) {
        // Register external hyperlink with relationship manager
        const url = item.getUrl();
        if (url) {
          const relationship = this.relationshipManager.addHyperlink(url);
          item.setRelationshipId(relationship.getId());
        }
      }
    }
  }

  /**
   * Saves all images to the ZIP archive
   */
  private saveImages(): void {
    const images = this.imageManager.getAllImages();

    for (const entry of images) {
      const imageData = entry.image.getImageData();
      if (imageData && imageData.length > 0) {
        // Save to word/media/
        const path = `word/media/${entry.filename}`;
        this.zipHandler.addFile(path, imageData);
      }
    }
  }

  /**
   * Saves all headers to the ZIP archive
   */
  private saveHeaders(): void {
    const headers = this.headerFooterManager.getAllHeaders();

    for (const entry of headers) {
      const xml = entry.header.toXML();
      const path = `word/${entry.filename}`;
      this.zipHandler.addFile(path, xml);
    }
  }

  /**
   * Saves all footers to the ZIP archive
   */
  private saveFooters(): void {
    const footers = this.headerFooterManager.getAllFooters();

    for (const entry of footers) {
      const xml = entry.footer.toXML();
      const path = `word/${entry.filename}`;
      this.zipHandler.addFile(path, xml);
    }
  }

  /**
   * Updates the word/_rels/document.xml.rels file with current relationships
   */
  private updateRelationships(): void {
    const xml = this.relationshipManager.generateXml();
    this.zipHandler.updateFile('word/_rels/document.xml.rels', xml);
  }

  /**
   * Saves all comments to the ZIP archive
   */
  private saveComments(): void {
    // Only save comments.xml if there are comments
    if (this.commentManager.getCount() > 0) {
      const xml = this.commentManager.generateCommentsXml();
      this.zipHandler.addFile('word/comments.xml', xml);

      // Add comments relationship
      this.relationshipManager.addComments();
    }
  }

  /**
   * Updates [Content_Types].xml to include image extensions, headers/footers, and comments
   */
  private updateContentTypesWithImagesHeadersFootersAndComments(): void {
    const contentTypes = this.generateContentTypesWithImagesHeadersFootersAndComments();
    this.zipHandler.updateFile(DOCX_PATHS.CONTENT_TYPES, contentTypes);
  }

  /**
   * Generates [Content_Types].xml with image extensions, headers/footers, and comments
   */
  private generateContentTypesWithImagesHeadersFootersAndComments(): string {
    const images = this.imageManager.getAllImages();
    const headers = this.headerFooterManager.getAllHeaders();
    const footers = this.headerFooterManager.getAllFooters();
    const hasComments = this.commentManager.getCount() > 0;

    // Collect unique extensions
    const extensions = new Set<string>();
    for (const entry of images) {
      extensions.add(entry.image.getExtension());
    }

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n';

    // Default types
    xml += '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n';
    xml += '  <Default Extension="xml" ContentType="application/xml"/>\n';

    // Image extensions
    for (const ext of extensions) {
      const mimeType = ImageManager.getMimeType(ext);
      xml += `  <Default Extension="${ext}" ContentType="${mimeType}"/>\n`;
    }

    // Override types
    xml += '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n';
    xml += '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n';
    xml += '  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n';

    // Header overrides
    for (const entry of headers) {
      xml += `  <Override PartName="/word/${entry.filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>\n`;
    }

    // Footer overrides
    for (const entry of footers) {
      xml += `  <Override PartName="/word/${entry.filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>\n`;
    }

    // Comments override
    if (hasComments) {
      xml += '  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>\n';
    }

    xml += '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n';
    xml += '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n';
    xml += '</Types>';

    return xml;
  }

  /**
   * Gets the BookmarkManager
   * @returns BookmarkManager instance
   */
  getBookmarkManager(): BookmarkManager {
    return this.bookmarkManager;
  }

  /**
   * Creates and registers a new bookmark with a unique name
   * @param name - Desired bookmark name
   * @returns The created and registered bookmark
   */
  createBookmark(name: string): Bookmark {
    return this.bookmarkManager.createBookmark(name);
  }

  /**
   * Creates and registers a bookmark for a heading
   * Automatically generates a unique name from the heading text
   * @param headingText - The text of the heading
   * @returns The created and registered bookmark
   */
  createHeadingBookmark(headingText: string): Bookmark {
    return this.bookmarkManager.createHeadingBookmark(headingText);
  }

  /**
   * Gets a bookmark by name
   * @param name - Bookmark name
   * @returns The bookmark, or undefined if not found
   */
  getBookmark(name: string): Bookmark | undefined {
    return this.bookmarkManager.getBookmark(name);
  }

  /**
   * Checks if a bookmark exists
   * @param name - Bookmark name
   * @returns True if the bookmark exists
   */
  hasBookmark(name: string): boolean {
    return this.bookmarkManager.hasBookmark(name);
  }

  /**
   * Adds a bookmark to a paragraph (wraps the entire paragraph)
   * Creates the bookmark if a name is provided, or uses an existing bookmark object
   * @param paragraph - The paragraph to bookmark
   * @param bookmarkOrName - Bookmark object or bookmark name
   * @returns The bookmark that was added
   */
  addBookmarkToParagraph(paragraph: Paragraph, bookmarkOrName: Bookmark | string): Bookmark {
    const bookmark = typeof bookmarkOrName === 'string'
      ? this.createBookmark(bookmarkOrName)
      : bookmarkOrName;

    paragraph.addBookmark(bookmark);
    return bookmark;
  }

  /**
   * Gets the RevisionManager
   * @returns RevisionManager instance
   */
  getRevisionManager(): RevisionManager {
    return this.revisionManager;
  }

  /**
   * Creates and registers a new insertion revision
   * @param author - Author who made the insertion
   * @param content - Inserted content (Run or array of Runs)
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createInsertion(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    const revision = Revision.createInsertion(author, content, date);
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a new deletion revision
   * @param author - Author who made the deletion
   * @param content - Deleted content (Run or array of Runs)
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createDeletion(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    const revision = Revision.createDeletion(author, content, date);
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a revision from text
   * Convenience method that creates a Run from the text
   * @param type - Revision type ('insert' or 'delete')
   * @param author - Author who made the change
   * @param text - Text content
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createRevisionFromText(
    type: RevisionType,
    author: string,
    text: string,
    date?: Date
  ): Revision {
    const revision = Revision.fromText(type, author, text, date);
    return this.revisionManager.register(revision);
  }

  /**
   * Adds a tracked insertion to a paragraph
   * @param paragraph - The paragraph to add the insertion to
   * @param author - Author who made the insertion
   * @param text - Inserted text
   * @param date - Optional date (defaults to now)
   * @returns The created revision
   */
  trackInsertion(
    paragraph: Paragraph,
    author: string,
    text: string,
    date?: Date
  ): Revision {
    const revision = this.createRevisionFromText('insert', author, text, date);
    paragraph.addRevision(revision);
    return revision;
  }

  /**
   * Adds a tracked deletion to a paragraph
   * @param paragraph - The paragraph to add the deletion to
   * @param author - Author who made the deletion
   * @param text - Deleted text
   * @param date - Optional date (defaults to now)
   * @returns The created revision
   */
  trackDeletion(
    paragraph: Paragraph,
    author: string,
    text: string,
    date?: Date
  ): Revision {
    const revision = this.createRevisionFromText('delete', author, text, date);
    paragraph.addRevision(revision);
    return revision;
  }

  /**
   * Checks if track changes is enabled (has any revisions)
   * @returns True if there are revisions
   */
  isTrackingChanges(): boolean {
    return this.revisionManager.isTrackingChanges();
  }

  /**
   * Gets statistics about tracked changes
   * @returns Object with revision statistics
   */
  getRevisionStats(): {
    total: number;
    insertions: number;
    deletions: number;
    authors: string[];
    nextId: number;
  } {
    return this.revisionManager.getStats();
  }

  /**
   * Gets the CommentManager
   * @returns CommentManager instance
   */
  getCommentManager(): CommentManager {
    return this.commentManager;
  }

  /**
   * Creates and registers a new comment
   * @param author - Comment author
   * @param content - Comment content (text or runs)
   * @param initials - Optional author initials
   * @returns The created and registered comment
   */
  createComment(
    author: string,
    content: string | Run | Run[],
    initials?: string
  ): Comment {
    return this.commentManager.createComment(author, content, initials);
  }

  /**
   * Creates and registers a reply to an existing comment
   * @param parentCommentId - ID of the parent comment
   * @param author - Reply author
   * @param content - Reply content (text or runs)
   * @param initials - Optional author initials
   * @returns The created and registered reply
   */
  createReply(
    parentCommentId: number,
    author: string,
    content: string | Run | Run[],
    initials?: string
  ): Comment {
    return this.commentManager.createReply(parentCommentId, author, content, initials);
  }

  /**
   * Gets a comment by ID
   * @param id - Comment ID
   * @returns The comment, or undefined if not found
   */
  getComment(id: number): Comment | undefined {
    return this.commentManager.getComment(id);
  }

  /**
   * Gets all comments (top-level only, not replies)
   * @returns Array of all top-level comments
   */
  getAllComments(): Comment[] {
    return this.commentManager.getAllComments();
  }

  /**
   * Adds a comment to a paragraph (wraps the entire paragraph)
   * Creates the comment if text is provided, or uses an existing comment object
   * @param paragraph - The paragraph to comment
   * @param commentOrAuthor - Comment object, or author name if creating new comment
   * @param content - Comment content (required if creating new comment)
   * @param initials - Optional author initials (for new comments)
   * @returns The comment that was added
   */
  addCommentToParagraph(
    paragraph: Paragraph,
    commentOrAuthor: Comment | string,
    content?: string | Run | Run[],
    initials?: string
  ): Comment {
    const comment = typeof commentOrAuthor === 'string'
      ? this.createComment(commentOrAuthor, content!, initials)
      : commentOrAuthor;

    paragraph.addComment(comment);
    return comment;
  }

  /**
   * Gets statistics about comments
   * @returns Object with comment statistics
   */
  getCommentStats(): {
    total: number;
    topLevel: number;
    replies: number;
    authors: string[];
    nextId: number;
  } {
    return this.commentManager.getStats();
  }

  /**
   * Checks if there are any comments in the document
   * @returns True if there are no comments
   */
  hasNoComments(): boolean {
    return this.commentManager.isEmpty();
  }

  /**
   * Checks if there are comments in the document
   * @returns True if there are comments
   */
  hasComments(): boolean {
    return !this.commentManager.isEmpty();
  }

  /**
   * Gets a comment thread (comment and all its replies)
   * @param commentId - ID of the top-level comment
   * @returns Object with the comment and its replies, or undefined if not found
   */
  getCommentThread(commentId: number): { comment: Comment; replies: Comment[] } | undefined {
    return this.commentManager.getCommentThread(commentId);
  }

  /**
   * Searches comments by text content
   * @param searchText - Text to search for (case-insensitive)
   * @returns Array of comments containing the search text
   */
  findCommentsByText(searchText: string): Comment[] {
    return this.commentManager.findCommentsByText(searchText);
  }

  /**
   * Gets the most recent comments
   * @param count - Number of recent comments to return
   * @returns Array of most recent comments
   */
  getRecentComments(count: number): Comment[] {
    return this.commentManager.getRecentComments(count);
  }

  /**
   * Checks if there are any revisions in the document
   * @returns True if there are no revisions
   */
  hasNoRevisions(): boolean {
    return this.revisionManager.isEmpty();
  }

  /**
   * Checks if there are revisions in the document
   * @returns True if there are revisions
   */
  hasRevisions(): boolean {
    return !this.revisionManager.isEmpty();
  }

  /**
   * Gets the most recent revisions
   * @param count - Number of recent revisions to return
   * @returns Array of most recent revisions
   */
  getRecentRevisions(count: number): Revision[] {
    return this.revisionManager.getRecentRevisions(count);
  }

  /**
   * Searches revisions by text content
   * @param searchText - Text to search for (case-insensitive)
   * @returns Array of revisions containing the search text
   */
  findRevisionsByText(searchText: string): Revision[] {
    return this.revisionManager.findRevisionsByText(searchText);
  }

  /**
   * Gets all insertion revisions
   * @returns Array of insertion revisions
   */
  getAllInsertions(): Revision[] {
    return this.revisionManager.getAllInsertions();
  }

  /**
   * Gets all deletion revisions
   * @returns Array of deletion revisions
   */
  getAllDeletions(): Revision[] {
    return this.revisionManager.getAllDeletions();
  }

  /**
   * Gets parse warnings collected during document loading
   * Only populated when loading existing documents in lenient mode
   * @returns Array of parse errors/warnings
   */
  getParseWarnings(): Array<{ element: string; error: Error }> {
    return [...this.parseErrors];
  }

  /**
   * Checks current memory usage and throws if above threshold
   * Prevents out-of-memory errors by failing early
   * @throws {Error} If memory usage exceeds configured percentage
   */
  private checkMemoryThreshold(): void {
    const { heapUsed, heapTotal } = process.memoryUsage();
    const usagePercent = (heapUsed / heapTotal) * 100;

    if (usagePercent > this.maxMemoryUsagePercent) {
      throw new Error(
        `Memory usage critical (${usagePercent.toFixed(1)}% of ${(heapTotal / 1024 / 1024).toFixed(0)}MB heap). ` +
        `Cannot process document safely. Consider:\n` +
        `- Reducing document size\n` +
        `- Optimizing/compressing images\n` +
        `- Splitting into multiple documents\n` +
        `- Increasing Node.js heap size (--max-old-space-size)`
      );
    }
  }

  /**
   * Estimates the size of the document
   * Provides breakdown by component and warnings if size is too large
   * @returns Size estimation with breakdown and optional warning
   */
  estimateSize(): {
    paragraphs: number;
    tables: number;
    images: number;
    estimatedXmlBytes: number;
    imageBytes: number;
    totalEstimatedBytes: number;
    totalEstimatedMB: number;
    warning?: string;
  } {
    // Count elements
    const paragraphCount = this.getParagraphCount();
    const tableCount = this.getTableCount();
    const imageCount = this.imageManager.getImageCount();

    // Estimate XML size
    // Average: 200 bytes per paragraph, 1000 bytes per table
    const estimatedXml = (paragraphCount * 200) + (tableCount * 1000) + 50000; // +50KB for structure

    // Get actual image sizes
    const imageBytes = this.imageManager.getTotalSize();

    // Total estimate
    const totalBytes = estimatedXml + imageBytes;
    const totalMB = totalBytes / (1024 * 1024);

    // Thresholds
    const WARNING_MB = 50;
    const ERROR_MB = 100;

    let warning: string | undefined;

    if (totalMB > ERROR_MB) {
      warning = `Document size (${totalMB.toFixed(1)}MB) exceeds recommended maximum of ${ERROR_MB}MB. ` +
        `This may cause memory issues. Consider splitting into multiple documents or optimizing images.`;
    } else if (totalMB > WARNING_MB) {
      warning = `Document size (${totalMB.toFixed(1)}MB) exceeds ${WARNING_MB}MB. ` +
        `Large documents may take longer to process and use significant memory.`;
    }

    return {
      paragraphs: paragraphCount,
      tables: tableCount,
      images: imageCount,
      estimatedXmlBytes: estimatedXml,
      imageBytes,
      totalEstimatedBytes: totalBytes,
      totalEstimatedMB: parseFloat(totalMB.toFixed(2)),
      warning,
    };
  }

  /**
   * Cleans up resources and clears all managers
   * Call this after saving in long-running processes to free memory
   * Especially important for API servers processing many documents
   */
  dispose(): void {
    // Clear all managers to free memory
    this.bodyElements = [];
    this.parseErrors = [];
    this.stylesManager = StylesManager.create();
    this.numberingManager = NumberingManager.create();
    this.imageManager.clear();
    this.imageManager.releaseAllImageData();
    this.relationshipManager = RelationshipManager.create();
    this.headerFooterManager = HeaderFooterManager.create();
    this.bookmarkManager.clear();
    this.revisionManager.clear();
    this.commentManager.clear();
  }

  /**
   * Gets size statistics for the document
   * @returns Size statistics
   */
  getSizeStats(): {
    elements: { paragraphs: number; tables: number; images: number };
    size: { xml: string; images: string; total: string };
    warnings: string[];
  } {
    const estimate = this.estimateSize();
    const warnings: string[] = [];

    if (estimate.warning) {
      warnings.push(estimate.warning);
    }

    // Format sizes for display
    const formatBytes = (bytes: number): string => {
      if (bytes < 1024) return `${bytes} B`;
      if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
      return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
    };

    return {
      elements: {
        paragraphs: estimate.paragraphs,
        tables: estimate.tables,
        images: estimate.images,
      },
      size: {
        xml: formatBytes(estimate.estimatedXmlBytes),
        images: formatBytes(estimate.imageBytes),
        total: formatBytes(estimate.totalEstimatedBytes),
      },
      warnings,
    };
  }
}
