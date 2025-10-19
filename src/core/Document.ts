/**
 * Document - High-level API for creating and managing Word documents
 * Provides a simple interface for creating DOCX files without managing ZIP and XML manually
 */

import { ZipHandler } from '../zip/ZipHandler';
import { DOCX_PATHS } from '../zip/types';
import { Paragraph } from '../elements/Paragraph';
import { Table } from '../elements/Table';
import { TableCell } from '../elements/TableCell';
import { Section } from '../elements/Section';
import { Image } from '../elements/Image';
import { ImageManager } from '../elements/ImageManager';
import { Header } from '../elements/Header';
import { Footer } from '../elements/Footer';
import { HeaderFooterManager } from '../elements/HeaderFooterManager';
import { TableOfContents } from '../elements/TableOfContents';
import { TableOfContentsElement } from '../elements/TableOfContentsElement';
import { Bookmark } from '../elements/Bookmark';
import { BookmarkManager } from '../elements/BookmarkManager';
import { Revision, RevisionType } from '../elements/Revision';
import { RevisionManager } from '../elements/RevisionManager';
import { Comment } from '../elements/Comment';
import { CommentManager } from '../elements/CommentManager';
import { FootnoteManager } from '../elements/FootnoteManager';
import { EndnoteManager } from '../elements/EndnoteManager';
import { Run } from '../elements/Run';
import { Hyperlink } from '../elements/Hyperlink';
import { XMLElement } from '../xml/XMLBuilder';
import { XMLParser } from '../xml/XMLParser';
import { StylesManager } from '../formatting/StylesManager';
import { Style, StyleProperties } from '../formatting/Style';
import { NumberingManager } from '../formatting/NumberingManager';
import { RelationshipManager } from './RelationshipManager';
import { DocumentParser } from './DocumentParser';
import { DocumentGenerator } from './DocumentGenerator';
import { DocumentValidator } from './DocumentValidator';

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
  language?: string;
}

/**
 * Document part representation
 * Represents any part within a DOCX package (XML, binary, etc.)
 */
export interface DocumentPart {
  /** Part name/path within the package */
  name: string;
  /** Part content (string for XML/text, Buffer for binary) */
  content: string | Buffer;
  /** MIME content type */
  contentType?: string;
  /** Whether the part is binary */
  isBinary?: boolean;
  /** Part size in bytes */
  size?: number;
}

/**
 * Document creation options
 */
export interface DocumentOptions {
  properties?: DocumentProperties;
  /** Maximum memory usage percentage (0-100) before throwing error. Default: 80 */
  maxMemoryUsagePercent?: number;
  /** Maximum absolute RSS (Resident Set Size) in MB. Default: 2048 (2GB) */
  maxRssMB?: number;
  /** Enable absolute RSS limit checking. Default: true */
  useAbsoluteMemoryLimit?: boolean;
  /** Strict parsing mode - throw errors instead of collecting warnings. Default: false */
  strictParsing?: boolean;
  /** Maximum number of images allowed in document. Default: 20 */
  maxImageCount?: number;
  /** Maximum total size of all images in MB. Default: 100 */
  maxTotalImageSizeMB?: number;
  /** Maximum size of a single image in MB. Default: 20 */
  maxSingleImageSizeMB?: number;
}

/**
 * Body element - can be a Paragraph, Table, or TableOfContentsElement
 */
type BodyElement = Paragraph | Table | TableOfContentsElement;

/**
 * ImageRun - A run that contains an image (drawing)
 * Extends Run class for type-safe paragraph content
 *
 * Note: This is a specialized Run that contains a drawing instead of text.
 * It generates proper w:r (run) XML with w:drawing child element.
 */
class ImageRun extends Run {
  private imageElement: Image;

  constructor(image: Image) {
    // Call parent constructor with empty text
    // The text is irrelevant for image runs
    super('');
    this.imageElement = image;
  }

  /**
   * Override toXML to generate image-specific XML
   * @returns XMLElement with w:r containing w:drawing
   */
  toXML(): XMLElement {
    const drawing = this.imageElement.toXML();
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
  // @ts-ignore - Reserved for future implementation
  private _footnoteManager: FootnoteManager;
  // @ts-ignore - Reserved for future implementation
  private _endnoteManager: EndnoteManager;

  // Helper classes for parsing, generation, and validation
  private parser: DocumentParser;
  private generator: DocumentGenerator;
  private validator: DocumentValidator;

  /**
   * Private constructor - use Document.create() or Document.load()
   * @param zipHandler Optional ZIP handler (for loading existing documents)
   * @param options Document options
   * @param initDefaults Whether to initialize with default relationships (false for loaded docs)
   */
  private constructor(zipHandler?: ZipHandler, options: DocumentOptions = {}, initDefaults: boolean = true) {
    this.zipHandler = zipHandler || new ZipHandler();

    // Initialize helper classes
    const strictParsing = options.strictParsing ?? false;
    const memoryPercent = options.maxMemoryUsagePercent ?? 80;

    this.parser = new DocumentParser(strictParsing);
    this.generator = new DocumentGenerator();
    this.validator = new DocumentValidator(memoryPercent, {
      maxMemoryUsagePercent: options.maxMemoryUsagePercent,
      maxRssMB: options.maxRssMB,
      useAbsoluteLimit: options.useAbsoluteMemoryLimit,
    });

    // Validate and sanitize properties before storing
    this.properties = options.properties ? DocumentValidator.validateProperties(options.properties) : {};

    // Initialize managers
    this.stylesManager = StylesManager.create();
    this.numberingManager = NumberingManager.create();
    this.section = Section.createLetter(); // Default Letter-sized section
    this.imageManager = ImageManager.create({
      maxImageCount: options.maxImageCount,
      maxTotalImageSizeMB: options.maxTotalImageSizeMB,
      maxSingleImageSizeMB: options.maxSingleImageSizeMB,
    });
    this.relationshipManager = RelationshipManager.create();
    this.headerFooterManager = HeaderFooterManager.create();
    this.bookmarkManager = BookmarkManager.create();
    this.revisionManager = RevisionManager.create();
    this.commentManager = CommentManager.create();
    this._footnoteManager = FootnoteManager.create();
    this._endnoteManager = EndnoteManager.create();

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
    this.zipHandler.addFile(DOCX_PATHS.CONTENT_TYPES, this.generator.generateContentTypes());

    // _rels/.rels
    this.zipHandler.addFile(DOCX_PATHS.RELS, this.generator.generateRels());

    // word/document.xml (will be updated when saving)
    this.zipHandler.addFile(DOCX_PATHS.DOCUMENT, this.generator.generateDocumentXml(this.bodyElements, this.section));

    // word/_rels/document.xml.rels
    this.zipHandler.addFile('word/_rels/document.xml.rels', this.relationshipManager.generateXml());

    // word/styles.xml
    this.zipHandler.addFile(DOCX_PATHS.STYLES, this.stylesManager.generateStylesXml());

    // word/numbering.xml
    this.zipHandler.addFile(DOCX_PATHS.NUMBERING, this.numberingManager.generateNumberingXml());

    // docProps/core.xml
    this.zipHandler.addFile(DOCX_PATHS.CORE_PROPS, this.generator.generateCoreProps(this.properties));

    // docProps/app.xml
    this.zipHandler.addFile(DOCX_PATHS.APP_PROPS, this.generator.generateAppProps());
  }

  /**
   * Parses the document XML and extracts paragraphs, runs, and tables
   */
  private async parseDocument(): Promise<void> {
    const result = await this.parser.parseDocument(this.zipHandler, this.relationshipManager);
    this.bodyElements = result.bodyElements;
    this.properties = result.properties;
    this.relationshipManager = result.relationshipManager;
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
    // Validate and sanitize properties before storing
    const validated = DocumentValidator.validateProperties(properties);
    this.properties = { ...this.properties, ...validated };
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
   * Saves the document to a file
   * @param filePath - Output file path
   */
  async save(filePath: string): Promise<void> {
    // Use atomic save pattern: save to temp file, then rename
    // This prevents partial/corrupted saves if operation fails mid-way
    const tempPath = `${filePath}.tmp.${Date.now()}`;

    try {
      // Validate before saving to prevent data loss
      this.validator.validateBeforeSave(this.bodyElements);

      // Check memory usage before starting
      this.validator.checkMemoryThreshold();

      // Load all image data before saving (now async)
      await this.imageManager.loadAllImageData();

      // Check memory again after loading images
      this.validator.checkMemoryThreshold();

      // Check document size and warn if too large
      const sizeInfo = this.validator.estimateSize(this.bodyElements, this.imageManager);
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
      this.validator.validateBeforeSave(this.bodyElements);

      // Check memory usage before starting
      this.validator.checkMemoryThreshold();

      // Load all image data before saving (now async)
      await this.imageManager.loadAllImageData();

      // Check memory again after loading images
      this.validator.checkMemoryThreshold();

      // Check document size and warn if too large
      const sizeInfo = this.validator.estimateSize(this.bodyElements, this.imageManager);
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
    const xml = this.generator.generateDocumentXml(this.bodyElements, this.section);
    this.zipHandler.updateFile(DOCX_PATHS.DOCUMENT, xml);
  }

  /**
   * Updates the core properties with current values
   */
  private updateCoreProps(): void {
    const xml = this.generator.generateCoreProps(this.properties);
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
    // Update styles XML immediately so it's reflected in getStylesXml()
    this.updateStylesXml();
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
   * Gets all styles in the document
   * @returns Array of all style definitions
   */
  getStyles(): Style[] {
    return this.stylesManager.getAllStyles();
  }

  /**
   * Removes a style from the document
   * @param styleId - Style ID to remove
   * @returns True if the style was removed, false if not found
   */
  removeStyle(styleId: string): boolean {
    return this.stylesManager.removeStyle(styleId);
  }

  /**
   * Updates an existing style with new properties
   * @param styleId - Style ID to update
   * @param properties - Properties to update
   * @returns True if the style was updated, false if not found
   */
  updateStyle(styleId: string, properties: Partial<StyleProperties>): boolean {
    const style = this.stylesManager.getStyle(styleId);
    if (!style) {
      return false;
    }

    // Update the style properties
    const currentProps = style.getProperties();
    const updatedProps = { ...currentProps, ...properties, styleId }; // Preserve styleId

    // Create new style with updated properties
    const updatedStyle = Style.create(updatedProps);

    // Replace in manager
    this.stylesManager.addStyle(updatedStyle);
    return true;
  }

  /**
   * Gets the raw styles.xml content as a string
   * @returns The raw XML content of styles.xml
   */
  getStylesXml(): string {
    const stylesFile = this.zipHandler.getFileAsString(DOCX_PATHS.STYLES);
    return stylesFile || this.stylesManager.generateStylesXml();
  }

  /**
   * Sets the raw styles.xml content
   *
   * **Warning:** This directly sets the XML content without validation.
   * Invalid XML may corrupt the document. Use StylesManager.validate()
   * to check the XML before setting.
   *
   * @param xml - The raw XML content to set
   */
  setStylesXml(xml: string): void {
    this.zipHandler.updateFile(DOCX_PATHS.STYLES, xml);

    // Clear the styles manager to force reload on next access
    this.stylesManager.clear();
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
    // Add image as a run (ImageRun extends Run and generates w:drawing in w:r)
    const imageRun = this.createImageRun(image);
    para.addRun(imageRun);

    this.bodyElements.push(para);
    return this;
  }

  /**
   * Creates a run containing an image
   * @param image The image
   * @returns ImageRun (extends Run) with the image
   */
  private createImageRun(image: Image): ImageRun {
    // ImageRun extends Run, so it's type-safe to add to paragraphs
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
    this.generator.processHyperlinks(
      this.bodyElements,
      this.headerFooterManager,
      this.relationshipManager
    );
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
    const contentTypes = this.generator.generateContentTypesWithImagesHeadersFootersAndComments(
      this.imageManager,
      this.headerFooterManager,
      this.commentManager
    );
    this.zipHandler.updateFile(DOCX_PATHS.CONTENT_TYPES, contentTypes);
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
    return this.parser.getParseErrors();
  }

  /**
   * Updates hyperlink URLs in the document using a URL mapping
   *
   * This method finds all external hyperlinks in the document and updates their URLs
   * according to the provided map. The relationships are updated in-place to maintain
   * document integrity and prevent orphaned relationships per ECMA-376 §17.16.22.
   *
   * **Important Notes:**
   * - Only updates external hyperlinks (not internal bookmarks)
   * - Only updates the URL, not the display text
   * - Relationships are updated in-place to maintain IDs
   * - To update text too, manually iterate and call setText() on hyperlinks
   *
   * **OpenXML Compliance:**
   * This implementation ensures proper OpenXML structure by:
   * 1. Updating existing relationship targets in-place (prevents orphaned relationships)
   * 2. Maintaining relationship IDs for document integrity
   * 3. Maintaining TargetMode="External" for all web links (per ECMA-376 §17.16.22)
   *
   * @param urlMap - Map of old URLs to new URLs
   * @returns Number of hyperlinks updated
   *
   * @example
   * ```typescript
   * // Load existing document
   * const doc = await Document.load('document.docx');
   *
   * // Define URL mappings (old URL → new URL)
   * const urlMap = new Map([
   *   ['https://old-site.com', 'https://new-site.com'],
   *   ['https://example.org', 'https://example.com']
   * ]);
   *
   * // Update hyperlink URLs
   * const updated = doc.updateHyperlinkUrls(urlMap);
   * console.log(`Updated ${updated} hyperlink(s)`);
   *
   * // Save with updated relationships
   * await doc.save('updated-document.docx');
   * ```
   *
   * @see {@link https://www.ecma-international.org/publications-and-standards/standards/ecma-376/ | ECMA-376 Part 1 §17.16.22}
   */
  updateHyperlinkUrls(urlMap: Map<string, string>): number {
    // Early exit if no URLs to update
    if (urlMap.size === 0) {
      return 0;
    }

    // Two-phase update to handle circular URL swaps correctly
    // Phase 1: Collect all updates without modifying hyperlinks
    const updates: Array<{ hyperlink: Hyperlink; newUrl: string; relationshipId?: string }> = [];

    // Iterate through all paragraphs in document body
    for (const para of this.getParagraphs()) {
      // Get all content items (runs, hyperlinks, fields, revisions)
      for (const content of para.getContent()) {
        // Check if content is a Hyperlink and is external
        if (content instanceof Hyperlink && content.isExternal()) {
          const currentUrl = content.getUrl();

          // If current URL is in the map, collect the update
          if (currentUrl && urlMap.has(currentUrl)) {
            const newUrl = urlMap.get(currentUrl)!;
            updates.push({
              hyperlink: content,
              newUrl,
              relationshipId: content.getRelationshipId()
            });
          }
        }
      }
    }

    // Phase 2: Apply all updates atomically
    // This prevents circular swap issues (e.g., A→B, B→A becomes B→A, A→B)
    for (const { hyperlink, newUrl, relationshipId } of updates) {
      // Update the hyperlink URL (maintains relationship ID)
      hyperlink.setUrl(newUrl);

      // Update the relationship target in-place if relationship exists
      if (relationshipId) {
        this.relationshipManager.updateHyperlinkTarget(relationshipId, newUrl);
      }
    }

    // Note: This implementation updates relationships in-place,
    // maintaining document integrity per ECMA-376

    return updates.length;
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
    return this.validator.estimateSize(this.bodyElements, this.imageManager);
  }

  /**
   * Cleans up resources and clears all managers
   * Call this after saving in long-running processes to free memory
   * Especially important for API servers processing many documents
   */
  dispose(): void {
    // Clear all managers to free memory
    this.bodyElements = [];
    this.parser.clearParseErrors();
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
    return this.validator.getSizeStats(this.bodyElements, this.imageManager);
  }

  // ==================== DOCUMENT PART ACCESS METHODS ====================
  // These methods provide low-level access to document package parts,
  // enabling advanced operations not covered by the high-level API.

  /**
   * Gets a specific document part from the package
   *
   * Provides direct access to any part within the DOCX package, including
   * XML parts, binary files, and custom parts. This enables advanced scenarios
   * not covered by the high-level API.
   *
   * @param partName - The part name/path (e.g., 'word/document.xml', '[Content_Types].xml')
   * @returns The document part with content and metadata, or null if not found
   *
   * @example
   * ```typescript
   * // Get the main document XML
   * const docPart = await doc.getPart('word/document.xml');
   * if (docPart) {
   *   console.log(docPart.content); // XML content as string
   * }
   *
   * // Get an image
   * const imagePart = await doc.getPart('word/media/image1.png');
   * if (imagePart) {
   *   console.log(imagePart.isBinary); // true
   *   // imagePart.content is a Buffer
   * }
   * ```
   */
  async getPart(partName: string): Promise<DocumentPart | null> {
    try {
      const file = this.zipHandler.getFile(partName);
      if (!file) {
        return null;
      }

      return {
        name: partName,
        content: file.content,
        contentType: this.getContentTypeForPart(partName),
        isBinary: file.isBinary,
        size: file.size,
      };
    } catch (error) {
      // Return null for any errors (file not found, etc.)
      return null;
    }
  }

  /**
   * Sets or updates a document part in the package
   *
   * Allows adding or updating any part within the DOCX package. Use with caution
   * as incorrect modifications can corrupt the document structure.
   *
   * **Important:** This method does not automatically update relationships or
   * content types. You may need to manually update these for new parts.
   *
   * @param partName - The part name/path
   * @param content - The part content (string for XML/text, Buffer for binary)
   * @returns Promise that resolves when the part is set
   *
   * @example
   * ```typescript
   * // Update custom XML part
   * await doc.setPart('customXml/item1.xml', '<data>Custom content</data>');
   *
   * // Add a new image (remember to update relationships and content types)
   * const imageBuffer = await fs.readFile('image.png');
   * await doc.setPart('word/media/image2.png', imageBuffer);
   * ```
   */
  async setPart(partName: string, content: string | Buffer): Promise<void> {
    // Determine if content is binary
    const isBinary = Buffer.isBuffer(content);

    // Add or update the file in the ZIP handler
    this.zipHandler.addFile(partName, content, { binary: isBinary });
  }

  /**
   * Removes a document part from the package
   *
   * **Warning:** Removing required parts can corrupt the document.
   * This method does not update relationships or content types that may
   * reference the removed part.
   *
   * @param partName - The part name/path to remove
   * @returns True if the part was removed, false if it didn't exist
   *
   * @example
   * ```typescript
   * // Remove a custom part
   * const removed = await doc.removePart('customXml/item1.xml');
   * console.log(removed ? 'Part removed' : 'Part not found');
   * ```
   */
  async removePart(partName: string): Promise<boolean> {
    return this.zipHandler.removeFile(partName);
  }

  /**
   * Lists all parts in the document package
   *
   * Returns an array of all part names/paths in the DOCX package,
   * useful for debugging, analysis, or discovering custom parts.
   *
   * @returns Array of part names
   *
   * @example
   * ```typescript
   * const parts = await doc.listParts();
   * console.log('Document contains', parts.length, 'parts');
   * parts.forEach(part => console.log(part));
   * ```
   */
  async listParts(): Promise<string[]> {
    return this.zipHandler.getFilePaths();
  }

  /**
   * Checks if a part exists in the document package
   *
   * @param partName - The part name/path to check
   * @returns True if the part exists, false otherwise
   *
   * @example
   * ```typescript
   * if (await doc.partExists('word/glossary/document.xml')) {
   *   console.log('Document has a glossary');
   * }
   * ```
   */
  async partExists(partName: string): Promise<boolean> {
    return this.zipHandler.hasFile(partName);
  }

  /**
   * Gets all content types from [Content_Types].xml
   *
   * Returns a map of part names/extensions to their MIME content types.
   * This is useful for understanding the document structure and registering
   * new content types for custom parts.
   *
   * @returns Map of part names/extensions to content types
   *
   * @example
   * ```typescript
   * const contentTypes = await doc.getContentTypes();
   * contentTypes.forEach((type, name) => {
   *   console.log(`${name}: ${type}`);
   * });
   * ```
   */
  async getContentTypes(): Promise<Map<string, string>> {
    const contentTypes = new Map<string, string>();

    try {
      const contentTypesXml = this.zipHandler.getFileAsString('[Content_Types].xml');
      if (!contentTypesXml) {
        return contentTypes;
      }

      // Parse content types XML
      // Match Default elements (by extension)
      const defaultPattern = /<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"/g;
      let match;
      while ((match = defaultPattern.exec(contentTypesXml)) !== null) {
        if (match[1] && match[2]) {
          contentTypes.set(`.${match[1]}`, match[2]);
        }
      }

      // Match Override elements (by part name)
      const overridePattern = /<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"/g;
      while ((match = overridePattern.exec(contentTypesXml)) !== null) {
        if (match[1] && match[2]) {
          contentTypes.set(match[1], match[2]);
        }
      }
    } catch (error) {
      // Return empty map on error
    }

    return contentTypes;
  }

  /**
   * Gets the raw XML content of a document part without any processing
   *
   * Returns the unparsed XML string for any part in the document package.
   * This is useful for advanced manipulation, debugging, or accessing
   * content types that don't have dedicated APIs.
   *
   * **Note**: For binary parts (images, fonts), this converts the Buffer to UTF-8
   * string, which may not be appropriate. Check the content type first.
   *
   * @param partName - Part path (e.g., 'word/document.xml', '[Content_Types].xml')
   * @returns Raw XML string, or null if part not found
   *
   * @example
   * ```typescript
   * // Get raw document XML
   * const xml = await doc.getRawXml('word/document.xml');
   * console.log(xml); // Complete XML as string
   *
   * // Get raw styles XML
   * const stylesXml = await doc.getRawXml('word/styles.xml');
   *
   * // Get package metadata
   * const coreProps = await doc.getRawXml('docProps/core.xml');
   * ```
   */
  async getRawXml(partName: string): Promise<string | null> {
    try {
      const part = await this.getPart(partName);
      if (!part) {
        return null;
      }

      // If already a string, return as-is
      if (typeof part.content === 'string') {
        return part.content;
      }

      // If Buffer, decode as UTF-8 (standard for XML files)
      if (Buffer.isBuffer(part.content)) {
        return part.content.toString('utf8');
      }

      return null;
    } catch (error) {
      return null;
    }
  }

  /**
   * Gets raw XML content for all text-based parts (non-binary files)
   *
   * Returns a map of part names to their raw XML content, excluding binary files
   * like images and fonts. Useful for debugging or batch processing.
   *
   * @returns Map of part names to raw XML content
   *
   * @example
   * ```typescript
   * // Get all XML parts
   * const allXml = await doc.getAllRawXml();
   * for (const [partName, xml] of allXml) {
   *   console.log(`${partName}: ${xml.length} bytes`);
   * }
   *
   * // Validate all XML parts
   * for (const [partName, xml] of allXml) {
   *   try {
   *     // Parse and validate XML
   *     const parser = new DOMParser();
   *     parser.parseFromString(xml, 'text/xml');
   *   } catch (e) {
   *     console.error(`Invalid XML in ${partName}:`, e);
   *   }
   * }
   * ```
   */
  async getAllRawXml(): Promise<Map<string, string>> {
    const xmlMap = new Map<string, string>();

    try {
      const parts = await this.listParts();

      for (const partName of parts) {
        // Skip binary files (images, fonts, etc.)
        if (partName.match(/\.(png|jpg|jpeg|gif|woff|woff2|ttf|otf|bin)$/i)) {
          continue;
        }

        const xml = await this.getRawXml(partName);
        if (xml) {
          xmlMap.set(partName, xml);
        }
      }
    } catch (error) {
      // Return partial results on error
    }

    return xmlMap;
  }

  /**
   * Sets or updates raw XML content for a document part
   *
   * Convenience method for updating XML content in document parts.
   * Automatically detects and handles text/XML content.
   *
   * **Note**: This method does not automatically update relationships or
   * content types. You may need to manually update these if adding new parts.
   *
   * @param partName - Part path (e.g., 'word/document.xml')
   * @param xmlContent - Raw XML string to set
   * @returns Promise that resolves when the part is updated
   *
   * @example
   * ```typescript
   * // Update document XML
   * const newXml = '<?xml version="1.0"?><w:document>...</w:document>';
   * await doc.setRawXml('word/document.xml', newXml);
   *
   * // Update styles
   * const stylesXml = await doc.getRawXml('word/styles.xml');
   * const modified = stylesXml.replace('Old Style', 'New Style');
   * await doc.setRawXml('word/styles.xml', modified);
   * ```
   */
  async setRawXml(partName: string, xmlContent: string): Promise<void> {
    if (typeof xmlContent !== 'string') {
      throw new Error('XML content must be a string');
    }

    // Use setPart to update the part (handles both string and binary detection)
    await this.setPart(partName, xmlContent);
  }

  /**
   * Adds or updates a content type registration
   *
   * Registers a new content type in [Content_Types].xml. This is required
   * when adding new types of parts to the document package.
   *
   * @param partNameOrExtension - Part name (e.g., '/word/custom.xml') or extension (e.g., '.xml')
   * @param contentType - MIME content type (e.g., 'application/xml')
   * @returns True if successful
   *
   * @example
   * ```typescript
   * // Register a custom XML part
   * await doc.addContentType('/customXml/item1.xml', 'application/xml');
   *
   * // Register a new file extension
   * await doc.addContentType('.json', 'application/json');
   * ```
   */
  async addContentType(partNameOrExtension: string, contentType: string): Promise<boolean> {
    try {
      let contentTypesXml = this.zipHandler.getFileAsString('[Content_Types].xml');
      if (!contentTypesXml) {
        return false;
      }

      const isExtension = partNameOrExtension.startsWith('.');

      if (isExtension) {
        // Add as Default element (for extensions)
        const extension = partNameOrExtension.substring(1);

        // Check if already exists
        const existingPattern = new RegExp(`<Default\\s+Extension="${extension}"\\s+ContentType="[^"]+"/?>`, 'g');
        if (existingPattern.test(contentTypesXml)) {
          // Update existing
          contentTypesXml = contentTypesXml.replace(
            existingPattern,
            `<Default Extension="${extension}" ContentType="${contentType}"/>`
          );
        } else {
          // Add new before closing tag
          contentTypesXml = contentTypesXml.replace(
            '</Types>',
            `  <Default Extension="${extension}" ContentType="${contentType}"/>\n</Types>`
          );
        }
      } else {
        // Add as Override element (for specific parts)
        const partName = partNameOrExtension.startsWith('/') ? partNameOrExtension : `/${partNameOrExtension}`;

        // Check if already exists
        const existingPattern = new RegExp(`<Override\\s+PartName="${partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"\\s+ContentType="[^"]+"/?>`, 'g');
        if (existingPattern.test(contentTypesXml)) {
          // Update existing
          contentTypesXml = contentTypesXml.replace(
            existingPattern,
            `<Override PartName="${partName}" ContentType="${contentType}"/>`
          );
        } else {
          // Add new before closing tag
          contentTypesXml = contentTypesXml.replace(
            '</Types>',
            `  <Override PartName="${partName}" ContentType="${contentType}"/>\n</Types>`
          );
        }
      }

      // Update the content types file
      this.zipHandler.updateFile('[Content_Types].xml', contentTypesXml);
      return true;
    } catch (error) {
      return false;
    }
  }

  /**
   * Gets all relationships for the document
   *
   * Returns a map of relationship file paths to their relationships.
   * This includes document relationships, part relationships, etc.
   *
   * @returns Map of relationship file paths to relationship arrays
   *
   * @example
   * ```typescript
   * const relationships = await doc.getAllRelationships();
   * relationships.forEach((rels, path) => {
   *   console.log(`${path}: ${rels.length} relationships`);
   * });
   * ```
   */
  async getAllRelationships(): Promise<Map<string, any[]>> {
    const relationships = new Map<string, any[]>();

    try {
      // Get all .rels files
      const relsPaths = this.zipHandler.getFilePaths().filter(path => path.endsWith('.rels'));

      for (const relsPath of relsPaths) {
        const relsContent = this.zipHandler.getFileAsString(relsPath);
        if (relsContent) {
          interface ParsedRelationship {
            id?: string;
            type?: string;
            target?: string;
            targetMode?: string;
          }

          const rels: ParsedRelationship[] = [];

          // Use XMLParser to extract all Relationship elements
          const relationshipElements = XMLParser.extractElements(relsContent, 'Relationship');

          for (const relElement of relationshipElements) {
            const rel: ParsedRelationship = {};

            // Extract attributes using XMLParser
            const id = XMLParser.extractAttribute(relElement, 'Id');
            const type = XMLParser.extractAttribute(relElement, 'Type');
            const target = XMLParser.extractAttribute(relElement, 'Target');
            const targetMode = XMLParser.extractAttribute(relElement, 'TargetMode');

            if (id) rel.id = id;
            if (type) rel.type = type;
            if (target) rel.target = target;
            if (targetMode) rel.targetMode = targetMode;

            rels.push(rel);
          }

          relationships.set(relsPath, rels);
        }
      }
    } catch (error) {
      // Return empty map on error
    }

    return relationships;
  }

  /**
   * Gets relationships for a specific document part
   *
   * Retrieves all relationships defined for a specific part's .rels file.
   * For example, calling with 'word/document.xml' returns relationships
   * from 'word/_rels/document.xml.rels'.
   *
   * @param partName - The part name to get relationships for (e.g., 'word/document.xml')
   * @returns Array of relationships for that part, or empty array if none found
   *
   * @example
   * ```typescript
   * // Get relationships for document
   * const docRels = await doc.getRelationships('word/document.xml');
   * for (const rel of docRels) {
   *   if (rel.type.includes('hyperlink')) {
   *     console.log('Hyperlink target:', rel.target);
   *   }
   * }
   *
   * // Get relationships for styles
   * const styleRels = await doc.getRelationships('word/styles.xml');
   *
   * // Get relationships for headers/footers
   * const headerRels = await doc.getRelationships('word/header1.xml');
   * ```
   */
  async getRelationships(partName: string): Promise<Array<{ id?: string; type?: string; target?: string; targetMode?: string }>> {
    try {
      // Construct the .rels path from the part name
      // For 'word/document.xml' -> 'word/_rels/document.xml.rels'
      const lastSlash = partName.lastIndexOf('/');
      const relsPath = lastSlash === -1
        ? `_rels/${partName}.rels`
        : `${partName.substring(0, lastSlash)}/_rels/${partName.substring(lastSlash + 1)}.rels`;

      const relsContent = this.zipHandler.getFileAsString(relsPath);
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

      // Use XMLParser to extract all Relationship elements
      const relationshipElements = XMLParser.extractElements(relsContent, 'Relationship');

      for (const relElement of relationshipElements) {
        const rel: ParsedRelationship = {};

        // Extract attributes using XMLParser
        const id = XMLParser.extractAttribute(relElement, 'Id');
        const type = XMLParser.extractAttribute(relElement, 'Type');
        const target = XMLParser.extractAttribute(relElement, 'Target');
        const targetMode = XMLParser.extractAttribute(relElement, 'TargetMode');

        if (id) rel.id = id;
        if (type) rel.type = type;
        if (target) rel.target = target;
        if (targetMode) rel.targetMode = targetMode;

        relationships.push(rel);
      }

      return relationships;
    } catch (error) {
      // Return empty array on error
      return [];
    }
  }

  /**
   * Gets the content type for a specific part
   * Helper method used internally by getPart
   */
  private getContentTypeForPart(partName: string): string | undefined {
    try {
      const contentTypesXml = this.zipHandler.getFileAsString('[Content_Types].xml');
      if (!contentTypesXml) {
        return undefined;
      }

      // Check for specific override
      const overridePattern = new RegExp(`<Override\\s+PartName="${partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"\\s+ContentType="([^"]+)"`, 'i');
      const overrideMatch = contentTypesXml.match(overridePattern);
      if (overrideMatch) {
        return overrideMatch[1];
      }

      // Check for extension default
      const ext = partName.substring(partName.lastIndexOf('.'));
      if (ext) {
        const defaultPattern = new RegExp(`<Default\\s+Extension="${ext.substring(1)}"\\s+ContentType="([^"]+)"`, 'i');
        const defaultMatch = contentTypesXml.match(defaultPattern);
        if (defaultMatch) {
          return defaultMatch[1];
        }
      }
    } catch (error) {
      // Return undefined on error
    }

    return undefined;
  }

  /**
   * Finds all occurrences of text in the document
   * @param text - Text to search for
   * @param options - Search options
   * @returns Array of search results with paragraph and run information
   */
  findText(text: string, options?: { caseSensitive?: boolean; wholeWord?: boolean }): Array<{
    paragraph: Paragraph;
    paragraphIndex: number;
    run: Run;
    runIndex: number;
    text: string;
    startIndex: number;
  }> {
    const results: Array<{
      paragraph: Paragraph;
      paragraphIndex: number;
      run: Run;
      runIndex: number;
      text: string;
      startIndex: number;
    }> = [];

    const caseSensitive = options?.caseSensitive ?? false;
    const wholeWord = options?.wholeWord ?? false;
    const searchText = caseSensitive ? text : text.toLowerCase();

    const paragraphs = this.getParagraphs();
    for (let pIndex = 0; pIndex < paragraphs.length; pIndex++) {
      const paragraph = paragraphs[pIndex];
      if (!paragraph) continue;
      const runs = paragraph.getRuns();

      for (let rIndex = 0; rIndex < runs.length; rIndex++) {
        const run = runs[rIndex];
        if (!run) continue;
        const runText = run.getText();
        const compareText = caseSensitive ? runText : runText.toLowerCase();

        if (wholeWord) {
          // Create word boundary regex
          const wordPattern = new RegExp(`\\b${searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, caseSensitive ? 'g' : 'gi');
          let match;
          while ((match = wordPattern.exec(runText)) !== null) {
            results.push({
              paragraph,
              paragraphIndex: pIndex,
              run,
              runIndex: rIndex,
              text: match[0],
              startIndex: match.index,
            });
          }
        } else {
          // Simple substring search
          let startIndex = 0;
          while ((startIndex = compareText.indexOf(searchText, startIndex)) !== -1) {
            results.push({
              paragraph,
              paragraphIndex: pIndex,
              run,
              runIndex: rIndex,
              text: runText.substr(startIndex, text.length),
              startIndex,
            });
            startIndex += text.length;
          }
        }
      }
    }

    // Also search in tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          if (cell instanceof TableCell) {
            const cellParagraphs = cell.getParagraphs();
            for (let pIndex = 0; pIndex < cellParagraphs.length; pIndex++) {
              const paragraph = cellParagraphs[pIndex];
              if (!paragraph) continue;
              const runs = paragraph.getRuns();

              for (let rIndex = 0; rIndex < runs.length; rIndex++) {
                const run = runs[rIndex];
                if (!run) continue;
                const runText = run.getText();
                const compareText = caseSensitive ? runText : runText.toLowerCase();

                if (wholeWord) {
                  // Create word boundary regex
                  const wordPattern = new RegExp(`\\b${searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, caseSensitive ? 'g' : 'gi');
                  let match;
                  while ((match = wordPattern.exec(runText)) !== null) {
                    results.push({
                      paragraph,
                      paragraphIndex: -1, // Not in main body, in table
                      run,
                      runIndex: rIndex,
                      text: match[0],
                      startIndex: match.index,
                    });
                  }
                } else {
                  // Simple substring search
                  let startIndex = 0;
                  while ((startIndex = compareText.indexOf(searchText, startIndex)) !== -1) {
                    results.push({
                      paragraph,
                      paragraphIndex: -1, // Not in main body, in table
                      run,
                      runIndex: rIndex,
                      text: runText.substr(startIndex, text.length),
                      startIndex,
                    });
                    startIndex += text.length;
                  }
                }
              }
            }
          }
        }
      }
    }

    return results;
  }

  /**
   * Replaces all occurrences of text in the document
   * @param find - Text to find
   * @param replace - Text to replace with
   * @param options - Replace options
   * @returns Number of replacements made
   */
  replaceText(find: string, replace: string, options?: { caseSensitive?: boolean; wholeWord?: boolean }): number {
    let replacementCount = 0;
    const caseSensitive = options?.caseSensitive ?? false;
    const wholeWord = options?.wholeWord ?? false;

    const paragraphs = this.getParagraphs();
    for (const paragraph of paragraphs) {
      const runs = paragraph.getRuns();

      for (const run of runs) {
        const originalText = run.getText();
        let newText = originalText;

        if (wholeWord) {
          // Use word boundary regex for whole word replacement
          const wordPattern = new RegExp(
            `\\b${find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`,
            caseSensitive ? 'g' : 'gi'
          );
          const matches = originalText.match(wordPattern);
          if (matches) {
            replacementCount += matches.length;
            newText = originalText.replace(wordPattern, replace);
          }
        } else {
          // Simple substring replacement
          const searchPattern = new RegExp(
            find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
            caseSensitive ? 'g' : 'gi'
          );
          const matches = originalText.match(searchPattern);
          if (matches) {
            replacementCount += matches.length;
            newText = originalText.replace(searchPattern, replace);
          }
        }

        if (newText !== originalText) {
          run.setText(newText);
        }
      }
    }

    return replacementCount;
  }

  /**
   * Gets the total word count in the document
   * @returns Total number of words
   */
  getWordCount(): number {
    let totalWords = 0;

    const paragraphs = this.getParagraphs();
    for (const paragraph of paragraphs) {
      const text = paragraph.getText().trim();
      if (text) {
        // Split by whitespace and filter out empty strings
        const words = text.split(/\s+/).filter(word => word.length > 0);
        totalWords += words.length;
      }
    }

    // Also count words in tables
    const tables = this.getTables();
    for (const table of tables) {
      const rows = table.getRows();
      for (const row of rows) {
        const cells = row.getCells();
        for (const cell of cells) {
          const cellParas = cell.getParagraphs();
          for (const para of cellParas) {
            const text = para.getText().trim();
            if (text) {
              const words = text.split(/\s+/).filter(word => word.length > 0);
              totalWords += words.length;
            }
          }
        }
      }
    }

    return totalWords;
  }

  /**
   * Gets the total character count in the document
   * @param includeSpaces - Whether to include spaces in the count
   * @returns Total number of characters
   */
  getCharacterCount(includeSpaces: boolean = true): number {
    let totalChars = 0;

    const paragraphs = this.getParagraphs();
    for (const paragraph of paragraphs) {
      const text = paragraph.getText();
      if (includeSpaces) {
        totalChars += text.length;
      } else {
        totalChars += text.replace(/\s/g, '').length;
      }
    }

    // Also count characters in tables
    const tables = this.getTables();
    for (const table of tables) {
      const rows = table.getRows();
      for (const row of rows) {
        const cells = row.getCells();
        for (const cell of cells) {
          const cellParas = cell.getParagraphs();
          for (const para of cellParas) {
            const text = para.getText();
            if (includeSpaces) {
              totalChars += text.length;
            } else {
              totalChars += text.replace(/\s/g, '').length;
            }
          }
        }
      }
    }

    return totalChars;
  }

  /**
   * Removes a paragraph from the document
   * @param paragraphOrIndex - The paragraph object or its index
   * @returns True if the paragraph was removed, false otherwise
   */
  removeParagraph(paragraphOrIndex: Paragraph | number): boolean {
    let index: number;

    if (typeof paragraphOrIndex === 'number') {
      index = paragraphOrIndex;
    } else {
      // Find the index of the paragraph
      index = this.bodyElements.indexOf(paragraphOrIndex);
    }

    if (index >= 0 && index < this.bodyElements.length) {
      const element = this.bodyElements[index];
      if (element instanceof Paragraph) {
        this.bodyElements.splice(index, 1);
        return true;
      }
    }

    return false;
  }

  /**
   * Removes a table from the document
   * @param tableOrIndex - The table object or its index
   * @returns True if the table was removed, false otherwise
   */
  removeTable(tableOrIndex: Table | number): boolean {
    let index: number;

    if (typeof tableOrIndex === 'number') {
      // If number provided, find the nth table
      const tables = this.getTables();
      if (tableOrIndex >= 0 && tableOrIndex < tables.length) {
        const table = tables[tableOrIndex];
        if (!table) return false;
        index = this.bodyElements.indexOf(table);
      } else {
        return false;
      }
    } else {
      // Find the index of the table
      index = this.bodyElements.indexOf(tableOrIndex);
    }

    if (index >= 0 && index < this.bodyElements.length) {
      const element = this.bodyElements[index];
      if (element instanceof Table) {
        this.bodyElements.splice(index, 1);
        return true;
      }
    }

    return false;
  }

  /**
   * Inserts a paragraph at a specific position
   * @param index - The position to insert at (0-based)
   * @param paragraph - The paragraph to insert
   * @returns This document for chaining
   */
  insertParagraphAt(index: number, paragraph: Paragraph): this {
    if (index < 0) {
      index = 0;
    } else if (index > this.bodyElements.length) {
      index = this.bodyElements.length;
    }

    this.bodyElements.splice(index, 0, paragraph);
    return this;
  }

  /**
   * Gets all hyperlinks in the document
   * @returns Array of hyperlinks with their containing paragraph
   */
  getHyperlinks(): Array<{ hyperlink: Hyperlink; paragraph: Paragraph }> {
    const hyperlinks: Array<{ hyperlink: Hyperlink; paragraph: Paragraph }> = [];

    for (const paragraph of this.getParagraphs()) {
      for (const content of paragraph.getContent()) {
        if (content instanceof Hyperlink) {
          hyperlinks.push({ hyperlink: content, paragraph });
        }
      }
    }

    // Also check in tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          // TableCell has getParagraphs method
          const cellParagraphs = cell instanceof TableCell ? cell.getParagraphs() : [];
          for (const para of cellParagraphs) {
            for (const content of para.getContent()) {
              if (content instanceof Hyperlink) {
                hyperlinks.push({ hyperlink: content, paragraph: para });
              }
            }
          }
        }
      }
    }

    return hyperlinks;
  }

  /**
   * Gets all bookmarks in the document
   * @returns Array of bookmarks with their containing paragraph
   */
  getBookmarks(): Array<{ bookmark: Bookmark; paragraph: Paragraph }> {
    const bookmarks: Array<{ bookmark: Bookmark; paragraph: Paragraph }> = [];

    for (const paragraph of this.getParagraphs()) {
      // Get bookmarks that start in this paragraph
      for (const bookmark of paragraph.getBookmarksStart()) {
        bookmarks.push({ bookmark, paragraph });
      }
    }

    // Also check in tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            for (const bookmark of para.getBookmarksStart()) {
              bookmarks.push({ bookmark, paragraph: para });
            }
          }
        }
      }
    }

    return bookmarks;
  }

  /**
   * Gets all images in the document
   * @returns Array of images with their metadata
   */
  getImages(): Array<{
    image: Image;
    relationshipId: string;
    filename: string;
  }> {
    return this.imageManager.getAllImages();
  }

  /**
   * Sets the document language
   * @param language - Language code (e.g., 'en-US', 'es-ES', 'fr-FR')
   * @returns This document for chaining
   */
  setLanguage(language: string): this {
    // Store language in properties for core.xml
    if (!this.properties) {
      this.properties = {};
    }
    this.properties.language = language;

    return this;
  }

  /**
   * Gets the document language
   * @returns Language code or undefined if not set
   */
  getLanguage(): string | undefined {
    return this.properties?.language;
  }

  /**
   * Creates an empty document with minimal structure
   *
   * Creates a new document with only the essential parts required
   * for a valid DOCX file, without any default content or styling.
   * Useful for building documents from scratch programmatically.
   *
   * @returns New empty Document instance
   *
   * @example
   * ```typescript
   * const doc = Document.createEmpty();
   * // Document has minimal structure, ready for content
   * doc.createParagraph('First paragraph');
   * await doc.save('minimal.docx');
   * ```
   */
  static createEmpty(): Document {
    const doc = new Document(undefined, {}, false); // Don't init defaults

    // Add only the absolute minimum required files
    const zipHandler = doc.getZipHandler();

    // [Content_Types].xml - minimal
    zipHandler.addFile('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n' +
      '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n' +
      '  <Default Extension="xml" ContentType="application/xml"/>\n' +
      '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n' +
      '</Types>'
    );

    // _rels/.rels - minimal
    zipHandler.addFile('_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
      '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n' +
      '</Relationships>'
    );

    // word/document.xml - empty body
    zipHandler.addFile('word/document.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n' +
      '  <w:body/>\n' +
      '</w:document>'
    );

    // word/_rels/document.xml.rels - empty relationships
    zipHandler.addFile('word/_rels/document.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    );

    return doc;
  }
}
