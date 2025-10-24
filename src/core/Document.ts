/**
 * Document - High-level API for creating and managing Word documents
 * Provides a simple interface for creating DOCX files without managing ZIP and XML manually
 */

import { ZipHandler } from "../zip/ZipHandler";
import { DOCX_PATHS } from "../zip/types";
import { Paragraph } from "../elements/Paragraph";
import { Table } from "../elements/Table";
import { TableCell } from "../elements/TableCell";
import { Section } from "../elements/Section";
import { Image } from "../elements/Image";
import { ImageManager } from "../elements/ImageManager";
import { ImageRun } from "../elements/ImageRun";
import { Header } from "../elements/Header";
import { Footer } from "../elements/Footer";
import { HeaderFooterManager } from "../elements/HeaderFooterManager";
import { TableOfContents } from "../elements/TableOfContents";
import { TableOfContentsElement } from "../elements/TableOfContentsElement";
import { Bookmark } from "../elements/Bookmark";
import { StructuredDocumentTag } from "../elements/StructuredDocumentTag";
import { BookmarkManager } from "../elements/BookmarkManager";
import { Revision, RevisionType } from "../elements/Revision";
import { RevisionManager } from "../elements/RevisionManager";
import { Comment } from "../elements/Comment";
import { CommentManager } from "../elements/CommentManager";
import { FootnoteManager } from "../elements/FootnoteManager";
import { EndnoteManager } from "../elements/EndnoteManager";
import { Run } from "../elements/Run";
import { Hyperlink } from "../elements/Hyperlink";
import { XMLParser } from "../xml/XMLParser";
import { StylesManager } from "../formatting/StylesManager";
import { Style, StyleProperties } from "../formatting/Style";
import { NumberingManager } from "../formatting/NumberingManager";
import { RelationshipManager } from "./RelationshipManager";
import { DocumentParser } from "./DocumentParser";
import { DocumentGenerator } from "./DocumentGenerator";
import { DocumentValidator } from "./DocumentValidator";
import { ILogger, defaultLogger } from "../utils/logger";

/**
 * Document properties (core and extended)
 */
export interface DocumentProperties {
  // Core Properties (docProps/core.xml)
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
  category?: string;
  contentStatus?: string;

  // Extended Properties (docProps/app.xml)
  application?: string;
  appVersion?: string;
  company?: string;
  manager?: string;
  version?: string;

  // Custom Properties (docProps/custom.xml)
  customProperties?: Record<string, string | number | boolean | Date>;
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
  /**
   * Logger instance for framework messages
   * Allows control over how warnings, info, and debug messages are handled
   * If not provided, uses ConsoleLogger with WARN minimum level
   * Use SilentLogger to suppress all logging
   * @example
   * // Use custom logger
   * const doc = Document.create({ logger: myCustomLogger });
   *
   * // Suppress all logging
   * import { SilentLogger } from 'docxmlater';
   * const doc = Document.create({ logger: new SilentLogger() });
   */
  logger?: ILogger;
}

/**
 * Body element - can be a Paragraph, Table, or TableOfContentsElement
 */
type BodyElement =
  | Paragraph
  | Table
  | TableOfContentsElement
  | StructuredDocumentTag;

/**
 * Represents a Word document
 */
export class Document {
  private zipHandler: ZipHandler;
  private bodyElements: BodyElement[] = [];
  private properties: DocumentProperties;
  public namespaces: Record<string, string> = {};
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
  private logger: ILogger;

  /**
   * Private constructor - use Document.create() or Document.load()
   * @param zipHandler Optional ZIP handler (for loading existing documents)
   * @param options Document options
   * @param initDefaults Whether to initialize with default relationships (false for loaded docs)
   */
  private constructor(
    zipHandler?: ZipHandler,
    options: DocumentOptions = {},
    initDefaults: boolean = true
  ) {
    this.zipHandler = zipHandler || new ZipHandler();

    // Initialize logger (use provided or default)
    this.logger = options.logger || defaultLogger;

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
    this.properties = options.properties
      ? DocumentValidator.validateProperties(options.properties)
      : {};

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
      this.relationshipManager.addFontTable();
      this.relationshipManager.addSettings();
      this.relationshipManager.addTheme();
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
  static async load(
    filePath: string,
    options?: DocumentOptions
  ): Promise<Document> {
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
  static async loadFromBuffer(
    buffer: Buffer,
    options?: DocumentOptions
  ): Promise<Document> {
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
    this.zipHandler.addFile(
      DOCX_PATHS.CONTENT_TYPES,
      this.generator.generateContentTypes()
    );

    // _rels/.rels
    this.zipHandler.addFile(DOCX_PATHS.RELS, this.generator.generateRels());

    // word/document.xml (will be updated when saving)
    this.zipHandler.addFile(
      DOCX_PATHS.DOCUMENT,
      this.generator.generateDocumentXml(
        this.bodyElements,
        this.section,
        this.namespaces
      )
    );

    // word/_rels/document.xml.rels
    this.zipHandler.addFile(
      "word/_rels/document.xml.rels",
      this.relationshipManager.generateXml()
    );

    // word/styles.xml
    this.zipHandler.addFile(
      DOCX_PATHS.STYLES,
      this.stylesManager.generateStylesXml()
    );

    // word/numbering.xml
    this.zipHandler.addFile(
      DOCX_PATHS.NUMBERING,
      this.numberingManager.generateNumberingXml()
    );

    // word/fontTable.xml (REQUIRED for DOCX compliance)
    this.zipHandler.addFile(
      "word/fontTable.xml",
      this.generator.generateFontTable()
    );

    // word/settings.xml (REQUIRED for DOCX compliance)
    this.zipHandler.addFile(
      "word/settings.xml",
      this.generator.generateSettings()
    );

    // word/theme/theme1.xml (REQUIRED for DOCX compliance)
    this.zipHandler.addFile(
      "word/theme/theme1.xml",
      this.generator.generateTheme()
    );

    // docProps/core.xml
    this.zipHandler.addFile(
      DOCX_PATHS.CORE_PROPS,
      this.generator.generateCoreProps(this.properties)
    );

    // docProps/app.xml
    this.zipHandler.addFile(
      DOCX_PATHS.APP_PROPS,
      this.generator.generateAppProps(this.properties)
    );

    // Note: docProps/custom.xml is added during save() if custom properties exist
  }

  /**
   * Parses the document XML and extracts paragraphs, runs, tables, images, and styles
   */
  private async parseDocument(): Promise<void> {
    const result = await this.parser.parseDocument(
      this.zipHandler,
      this.relationshipManager,
      this.imageManager
    );
    this.bodyElements = result.bodyElements;
    this.properties = result.properties;
    this.relationshipManager = result.relationshipManager;
    this.namespaces = result.namespaces;

    // Load parsed section properties (preserves page size, margins, headers, etc.)
    if (result.section) {
      this.section = result.section;
    }

    // Load parsed styles into StylesManager
    // This replaces built-in styles with document-specific styles
    if (result.styles && result.styles.length > 0) {
      for (const style of result.styles) {
        this.stylesManager.addStyle(style);
      }
    }

    // Load parsed numbering into NumberingManager
    // This preserves existing list definitions from the document
    if (result.abstractNumberings && result.abstractNumberings.length > 0) {
      for (const abstractNum of result.abstractNumberings) {
        this.numberingManager.addAbstractNumbering(abstractNum);
      }
    }

    if (result.numberingInstances && result.numberingInstances.length > 0) {
      for (const instance of result.numberingInstances) {
        this.numberingManager.addInstance(instance);
      }
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
   * Adds a structured document tag (content control) to the document
   * @param sdt - Structured document tag to add
   * @returns This document for chaining
   */
  addStructuredDocumentTag(sdt: StructuredDocumentTag): this {
    this.bodyElements.push(sdt);
    return this;
  }

  /**
   * Adds a Table of Contents to the document
   * @param toc - TableOfContents or TableOfContentsElement to add
   * @returns This document for chaining
   */
  addTableOfContents(toc?: TableOfContents | TableOfContentsElement): this {
    // Wrap in TableOfContentsElement if plain TableOfContents provided
    const tocElement =
      toc instanceof TableOfContentsElement
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
    return this.bodyElements.filter(
      (el): el is Paragraph => el instanceof Paragraph
    );
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
    return this.bodyElements.filter(
      (el): el is TableOfContentsElement => el instanceof TableOfContentsElement
    );
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
   * Sets document property value
   * @param key - Property key
   * @param value - Property value
   * @returns This document for chaining
   */
  setProperty(key: keyof DocumentProperties, value: any): this {
    (this.properties as any)[key] = value;
    return this;
  }

  /**
   * Sets the document title
   * @param title - Document title
   * @returns This document for chaining
   */
  setTitle(title: string): this {
    this.properties.title = title;
    return this;
  }

  /**
   * Sets the document subject
   * @param subject - Document subject
   * @returns This document for chaining
   */
  setSubject(subject: string): this {
    this.properties.subject = subject;
    return this;
  }

  /**
   * Sets the document creator/author
   * @param creator - Document creator
   * @returns This document for chaining
   */
  setCreator(creator: string): this {
    this.properties.creator = creator;
    return this;
  }

  /**
   * Sets the document keywords
   * @param keywords - Document keywords (comma-separated)
   * @returns This document for chaining
   */
  setKeywords(keywords: string): this {
    this.properties.keywords = keywords;
    return this;
  }

  /**
   * Sets the document description
   * @param description - Document description
   * @returns This document for chaining
   */
  setDescription(description: string): this {
    this.properties.description = description;
    return this;
  }

  /**
   * Sets the document category
   * @param category - Document category
   * @returns This document for chaining
   */
  setCategory(category: string): this {
    this.properties.category = category;
    return this;
  }

  /**
   * Sets the document content status
   * @param status - Content status (e.g., "Draft", "Final", "In Review")
   * @returns This document for chaining
   */
  setContentStatus(status: string): this {
    this.properties.contentStatus = status;
    return this;
  }

  /**
   * Sets the application name
   * @param application - Application name
   * @returns This document for chaining
   */
  setApplication(application: string): this {
    this.properties.application = application;
    return this;
  }

  /**
   * Sets the application version
   * @param version - Application version
   * @returns This document for chaining
   */
  setAppVersion(version: string): this {
    this.properties.appVersion = version;
    return this;
  }

  /**
   * Sets the company name
   * @param company - Company name
   * @returns This document for chaining
   */
  setCompany(company: string): this {
    this.properties.company = company;
    return this;
  }

  /**
   * Sets the manager name
   * @param manager - Manager name
   * @returns This document for chaining
   */
  setManager(manager: string): this {
    this.properties.manager = manager;
    return this;
  }

  /**
   * Sets a custom property
   * @param name - Property name
   * @param value - Property value (string, number, boolean, or Date)
   * @returns This document for chaining
   */
  setCustomProperty(
    name: string,
    value: string | number | boolean | Date
  ): this {
    if (!this.properties.customProperties) {
      this.properties.customProperties = {};
    }
    this.properties.customProperties[name] = value;
    return this;
  }

  /**
   * Sets multiple custom properties
   * @param properties - Object containing custom properties
   * @returns This document for chaining
   */
  setCustomProperties(
    properties: Record<string, string | number | boolean | Date>
  ): this {
    this.properties.customProperties = { ...properties };
    return this;
  }

  /**
   * Gets a custom property value
   * @param name - Property name
   * @returns Property value or undefined
   */
  getCustomProperty(
    name: string
  ): string | number | boolean | Date | undefined {
    return this.properties.customProperties?.[name];
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
      const sizeInfo = this.validator.estimateSize(
        this.bodyElements,
        this.imageManager
      );
      if (sizeInfo.warning) {
        this.logger.warn(sizeInfo.warning, {
          totalMB: sizeInfo.totalEstimatedMB,
          paragraphs: sizeInfo.paragraphs,
          tables: sizeInfo.tables,
          images: sizeInfo.images,
        });
      }

      this.processHyperlinks();
      this.updateDocumentXml();
      this.updateStylesXml();
      this.updateNumberingXml();
      this.updateCoreProps();
      this.updateAppProps(); // Update app.xml with current property values
      this.saveImages();
      this.saveHeaders();
      this.saveFooters();
      this.saveComments();
      this.saveCustomProperties(); // Add custom.xml if custom properties exist
      this.updateRelationships();
      this.updateContentTypesWithImagesHeadersFootersAndComments();

      // Save to temporary file first
      await this.zipHandler.save(tempPath);

      // Atomic rename - only if save succeeded
      const { promises: fs } = await import("fs");
      await fs.rename(tempPath, filePath);
    } catch (error) {
      // Cleanup temporary file on error
      try {
        const { promises: fs } = await import("fs");
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
      const sizeInfo = this.validator.estimateSize(
        this.bodyElements,
        this.imageManager
      );
      if (sizeInfo.warning) {
        this.logger.warn(sizeInfo.warning, {
          totalMB: sizeInfo.totalEstimatedMB,
          paragraphs: sizeInfo.paragraphs,
          tables: sizeInfo.tables,
          images: sizeInfo.images,
        });
      }

      this.processHyperlinks();
      this.updateDocumentXml();
      this.updateStylesXml();
      this.updateNumberingXml();
      this.updateCoreProps();
      this.updateAppProps(); // Update app.xml with current property values
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
    const xml = this.generator.generateDocumentXml(
      this.bodyElements,
      this.section,
      this.namespaces
    );
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
   * Updates the app properties with current values
   */
  private updateAppProps(): void {
    const xml = this.generator.generateAppProps(this.properties);
    this.zipHandler.updateFile(DOCX_PATHS.APP_PROPS, xml);
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

    // Deep merge nested properties (paragraphFormatting, runFormatting)
    const updatedProps: StyleProperties = {
      ...currentProps,
      ...properties,
      styleId, // Preserve styleId
      // Deep merge paragraph formatting
      paragraphFormatting: properties.paragraphFormatting
        ? {
            ...currentProps.paragraphFormatting,
            ...properties.paragraphFormatting,
            // Deep merge nested spacing and indentation
            spacing: properties.paragraphFormatting.spacing
              ? {
                  ...currentProps.paragraphFormatting?.spacing,
                  ...properties.paragraphFormatting.spacing,
                }
              : currentProps.paragraphFormatting?.spacing,
            indentation: properties.paragraphFormatting.indentation
              ? {
                  ...currentProps.paragraphFormatting?.indentation,
                  ...properties.paragraphFormatting.indentation,
                }
              : currentProps.paragraphFormatting?.indentation,
          }
        : currentProps.paragraphFormatting,
      // Deep merge run formatting
      runFormatting: properties.runFormatting
        ? { ...currentProps.runFormatting, ...properties.runFormatting }
        : currentProps.runFormatting,
    };

    // Create new style with updated properties
    const updatedStyle = Style.create(updatedProps);

    // Replace in manager
    this.stylesManager.addStyle(updatedStyle);
    return true;
  }

  /**
   * Applies a style to all elements matching a predicate
   * @param styleId - Style ID to apply
   * @param predicate - Function to test each element
   * @returns Number of elements updated
   * @example
   * ```typescript
   * // Apply Heading1 style to all paragraphs containing "Chapter"
   * const count = doc.applyStyleToAll('Heading1', (el) => {
   *   return el instanceof Paragraph && el.getText().includes('Chapter');
   * });
   * console.log(`Updated ${count} elements`);
   * ```
   */
  applyStyleToAll(
    styleId: string,
    predicate: (
      element:
        | Paragraph
        | Table
        | TableOfContentsElement
        | StructuredDocumentTag
    ) => boolean
  ): number {
    let count = 0;

    for (const element of this.bodyElements) {
      if (predicate(element)) {
        if (element instanceof Paragraph) {
          element.setStyle(styleId);
          count++;
        }
      }
    }

    // Also check paragraphs inside tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            if (predicate(para)) {
              para.setStyle(styleId);
              count++;
            }
          }
        }
      }
    }

    return count;
  }

  /**
   * Finds all elements using a specific style
   * @param styleId - Style ID to search for
   * @returns Array of paragraphs and table cells using this style
   * @example
   * ```typescript
   * const heading1Elements = doc.findElementsByStyle('Heading1');
   * console.log(`Found ${heading1Elements.length} Heading1 elements`);
   * ```
   */
  findElementsByStyle(styleId: string): Array<Paragraph | TableCell> {
    const results: Array<Paragraph | TableCell> = [];

    // Check body paragraphs
    for (const element of this.bodyElements) {
      if (element instanceof Paragraph) {
        const formatting = element.getFormatting();
        if (formatting.style === styleId) {
          results.push(element);
        }
      }
    }

    // Check paragraphs inside tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            const formatting = para.getFormatting();
            if (formatting.style === styleId) {
              results.push(para);
            }
          }

          // Include the cell itself if it has styled paragraphs
          const hasStyledParagraph = cell
            .getParagraphs()
            .some((p) => p.getFormatting().style === styleId);
          if (hasStyledParagraph) {
            results.push(cell);
          }
        }
      }
    }

    return results;
  }

  /**
   * Applies a new style to all paragraphs currently using a specific style
   * Useful for bulk style updates across the document
   * @param currentStyleId - The style ID currently applied to paragraphs
   * @param newStyleId - The style ID to apply
   * @returns Number of paragraphs updated
   * @example
   * ```typescript
   * // Change all Normal paragraphs to BodyText
   * const count = doc.applyStyleToAllParagraphsWithStyle('Normal', 'BodyText');
   * console.log(`Updated ${count} paragraphs`);
   * ```
   */
  applyStyleToAllParagraphsWithStyle(
    currentStyleId: string,
    newStyleId: string
  ): number {
    let count = 0;

    // Check body paragraphs
    for (const element of this.bodyElements) {
      if (element instanceof Paragraph) {
        const formatting = element.getFormatting();
        if (formatting.style === currentStyleId) {
          element.setStyle(newStyleId);
          count++;
        }
      }
    }

    // Check paragraphs inside tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            const formatting = para.getFormatting();
            if (formatting.style === currentStyleId) {
              para.setStyle(newStyleId);
              count++;
            }
          }
        }
      }
    }

    return count;
  }

  /**
   * Applies a style to all paragraphs in the document with optional filtering and formatting control
   *
   * This method provides flexible bulk style application with the ability to:
   * - Apply style to all paragraphs or filter by current style
   * - Clear direct run formatting so the style takes full effect
   * - Clear only specific formatting properties
   *
   * @param styleId - The style ID to apply (e.g., 'Normal', 'Heading1')
   * @param options - Optional configuration
   * @param options.currentStyleId - Only apply to paragraphs with this style (undefined = all paragraphs)
   * @param options.clearFormatting - Whether to clear direct run formatting (default: false)
   * @param options.clearProperties - Specific properties to clear (default: all if clearFormatting=true)
   * @returns Number of paragraphs updated
   * @example
   * ```typescript
   * // Apply Normal to all paragraphs and clear all direct formatting
   * doc.applyStyleToAllParagraphs('Normal', { clearFormatting: true });
   *
   * // Apply Heading1 to all Normal paragraphs, clear only font and color
   * doc.applyStyleToAllParagraphs('Heading1', {
   *   currentStyleId: 'Normal',
   *   clearFormatting: true,
   *   clearProperties: ['font', 'color']
   * });
   *
   * // Apply BodyText to all paragraphs without clearing formatting
   * doc.applyStyleToAllParagraphs('BodyText');
   * ```
   */
  applyStyleToAllParagraphs(
    styleId: string,
    options?: {
      currentStyleId?: string;
      clearFormatting?: boolean;
      clearProperties?: string[];
    }
  ): number {
    let count = 0;
    const clearFormatting = options?.clearFormatting || false;
    const clearProperties = options?.clearProperties;
    const currentStyleId = options?.currentStyleId;

    // Helper to apply style to a paragraph
    const applyToParagraph = (para: Paragraph): void => {
      const formatting = para.getFormatting();

      // Check if this paragraph matches the filter
      if (currentStyleId !== undefined && formatting.style !== currentStyleId) {
        return; // Skip this paragraph
      }

      // Apply style
      if (clearFormatting) {
        para.applyStyleAndClearFormatting(
          styleId,
          clearProperties === undefined ? [] : clearProperties
        );
      } else {
        para.setStyle(styleId);
      }

      count++;
    };

    // Process body paragraphs
    for (const element of this.bodyElements) {
      if (element instanceof Paragraph) {
        applyToParagraph(element);
      }
    }

    // Process paragraphs inside tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            applyToParagraph(para);
          }
        }
      }
    }

    return count;
  }

  /**
   * Updates the color of all hyperlinks in the document
   * @param color - Hex color without # (e.g., '0000FF' for blue)
   * @returns Number of hyperlinks updated
   * @example
   * ```typescript
   * // Make all hyperlinks blue
   * const count = doc.updateAllHyperlinkColors('0000FF');
   * console.log(`Updated ${count} hyperlinks`);
   * ```
   */
  updateAllHyperlinkColors(color: string): number {
    const hyperlinks = this.getHyperlinks();

    for (const { hyperlink } of hyperlinks) {
      const currentFormatting = hyperlink.getFormatting();
      hyperlink.setFormatting({
        ...currentFormatting,
        color: color,
      });
    }

    return hyperlinks.length;
  }

  /**
   * Sets layout for all tables in the document
   * @param layout - Layout type ('auto' for fit to window, 'fixed' for fixed width)
   * @returns Number of tables updated
   * @example
   * ```typescript
   * // Make all tables fit to window
   * const count = doc.setAllTablesLayout('auto');
   * console.log(`Updated ${count} tables`);
   * ```
   */
  setAllTablesLayout(layout: "auto" | "fixed"): number {
    const tables = this.getTables();

    for (const table of tables) {
      table.setLayout(layout);
    }

    return tables.length;
  }

  /**
   * Applies different shading to tables based on their size
   * 1x1 tables get one color, multi-cell tables get another color on first row
   * @param singleCellShading - Hex color for 1x1 tables (e.g., 'BFBFBF')
   * @param multiCellFirstRowShading - Hex color for first row of multi-cell tables (e.g., 'DFDFDF')
   * @returns Object with counts of single-cell and multi-cell tables updated
   * @example
   * ```typescript
   * const result = doc.applyTableFormattingBySize('BFBFBF', 'DFDFDF');
   * console.log(`Updated ${result.singleCellCount} 1x1 tables and ${result.multiCellCount} multi-cell tables`);
   * ```
   */
  applyTableFormattingBySize(
    singleCellShading: string,
    multiCellFirstRowShading: string
  ): { singleCellCount: number; multiCellCount: number } {
    const tables = this.getTables();
    let singleCellCount = 0;
    let multiCellCount = 0;

    for (const table of tables) {
      const rowCount = table.getRowCount();
      const colCount = table.getColumnCount();

      if (rowCount === 1 && colCount === 1) {
        // 1x1 table - shade the single cell
        const cell = table.getCell(0, 0);
        if (cell) {
          cell.setShading({ fill: singleCellShading });
          singleCellCount++;
        }
      } else {
        // Multi-cell table - shade first row
        const firstRow = table.getRow(0);
        if (firstRow) {
          for (const cell of firstRow.getCells()) {
            cell.setShading({ fill: multiCellFirstRowShading });
          }
          multiCellCount++;
        }
      }
    }

    return { singleCellCount, multiCellCount };
  }

  /**
   * Fixes all "Top of Document" hyperlinks to use standard formatting and anchor
   *
   * Changes made to matching hyperlinks:
   * - Text: Any variation → "Top of the Document"
   * - Formatting: Verdana 12pt, underline, blue (0000FF)
   * - Paragraph: Right aligned
   * - Anchor: Points to "_top" bookmark at document start
   *
   * Creates "_top" bookmark at first paragraph if it doesn't exist
   *
   * @returns Number of hyperlinks updated
   * @example
   * ```typescript
   * const count = doc.fixTODHyperlinks();
   * console.log(`Fixed ${count} "Top of Document" links`);
   * ```
   */
  fixTODHyperlinks(): number {
    let count = 0;

    // Ensure _top bookmark exists at document start
    if (!this.hasBookmark("_top")) {
      const paragraphs = this.getParagraphs();
      if (paragraphs.length > 0) {
        const firstPara = paragraphs[0];
        if (firstPara) {
          const bookmark = new Bookmark({ name: "_top" });
          const registered = this.bookmarkManager.register(bookmark);
          firstPara.addBookmark(registered);
        }
      }
    }

    // Find and fix all "Top of Document" hyperlinks
    const hyperlinks = this.getHyperlinks();

    for (const { hyperlink, paragraph } of hyperlinks) {
      const text = hyperlink.getText().toLowerCase();

      // Match variations: "top of document", "top of the document", etc.
      if (text.includes("top") && text.includes("document")) {
        // Update text
        hyperlink.setText("Top of the Document");

        // Update formatting
        hyperlink.setFormatting({
          font: "Verdana",
          size: 12,
          underline: "single",
          color: "0000FF",
        });

        // Update anchor to _top
        hyperlink.setAnchor("_top");

        // Set paragraph alignment to right
        paragraph.setAlignment("right");

        count++;
      }
    }

    return count;
  }

  /**
   * Sets width to 5% for first column of tables containing "If"
   *
   * For tables with 2+ columns where first column text contains "If",
   * sets the first column width to 5% of table width.
   *
   * Common use case: If/Then decision tables
   *
   * @returns Number of table cells updated
   * @example
   * ```typescript
   * const count = doc.setIfColumnWidth();
   * console.log(`Updated ${count} If column cells`);
   * ```
   */
  setIfColumnWidth(): number {
    let count = 0;
    const tables = this.getTables();

    for (const table of tables) {
      const columnCount = table.getColumnCount();

      // Only process tables with at least 2 columns
      if (columnCount < 2) continue;

      // Check all rows for "If" in first column
      let hasIfColumn = false;
      const rows = table.getRows();

      for (const row of rows) {
        const cells = row.getCells();
        const firstCell = cells[0];
        if (firstCell) {
          const text = firstCell.getText().toLowerCase();
          if (text.includes("if")) {
            hasIfColumn = true;
            break;
          }
        }
      }

      // If "If" found, set width on all first column cells
      if (hasIfColumn) {
        const tableWidth = table.getFormatting().width || 12960; // Default page width in twips
        const targetWidth = Math.round(tableWidth * 0.05); // 5%

        for (const row of rows) {
          const cells = row.getCells();
          const firstCell = cells[0];
          if (firstCell) {
            firstCell.setWidth(targetWidth);
            count++;
          }
        }
      }
    }

    return count;
  }

  /**
   * Centers all images greater than specified pixel size
   *
   * Actually centers the paragraph containing the image, since images
   * themselves don't have alignment properties in Word.
   *
   * Conversion: 100 pixels ≈ 952,500 EMUs (at 96 DPI)
   *
   * @param minPixels Minimum size in pixels for both width and height (default: 100)
   * @returns Number of images centered
   * @example
   * ```typescript
   * const count = doc.centerLargeImages(100);
   * console.log(`Centered ${count} large images`);
   * ```
   */
  centerLargeImages(minPixels: number = 100): number {
    let count = 0;

    // Convert pixels to EMUs (914400 EMUs per inch, 96 DPI)
    // Formula: pixels * (914400 / 96) = pixels * 9525
    const minEmus = Math.round(minPixels * 9525);

    // Get all images with metadata
    const images = this.imageManager.getAllImages();

    // Create a Set of image IDs that meet size criteria
    const largeImageIds = new Set<string>();

    for (const entry of images) {
      const image = entry.image;
      const width = image.getWidth();
      const height = image.getHeight();

      // Check if both dimensions meet minimum
      if (width >= minEmus && height >= minEmus) {
        // Track this image's relationship ID
        const relId = image.getRelationshipId();
        if (relId) {
          largeImageIds.add(relId);
        }
      }
    }

    // Find paragraphs containing these large images and center them
    // Note: Images are embedded in paragraphs as ImageRun elements
    for (const paragraph of this.getParagraphs()) {
      const content = paragraph.getContent();

      for (const item of content) {
        // Check if this is an ImageRun (subclass of Run)
        if (item instanceof ImageRun) {
          const image = item.getImageElement();
          const relId = image.getRelationshipId();

          if (relId && largeImageIds.has(relId)) {
            // Center this paragraph
            paragraph.setAlignment("center");
            count++;
            break; // Only count paragraph once
          }
        }
      }
    }

    return count;
  }

  /**
   * Sets line spacing for all list items (numbered or bulleted)
   *
   * @param spacingTwips Line spacing in twips (default: 240 = 12pt)
   * @returns Number of list items updated
   * @example
   * ```typescript
   * const count = doc.setListLineSpacing(240);
   * console.log(`Updated ${count} list items`);
   * ```
   */
  setListLineSpacing(spacingTwips: number = 240): number {
    let count = 0;

    for (const paragraph of this.getParagraphs()) {
      const numbering = paragraph.getNumbering();

      if (numbering) {
        // Has numbering - it's a list item
        paragraph.setLineSpacing(spacingTwips, "auto");
        count++;
      }
    }

    return count;
  }

  /**
   * Normalizes all numbered lists in the document to use consistent formatting
   *
   * Creates a standard numbered list format and applies it to all numbered lists:
   * - Level 0: 1., 2., 3., ... (decimal)
   * - Level 1: a., b., c., ... (lowerLetter)
   * - Level 2: i., ii., iii., ... (lowerRoman)
   * - Consistent indentation and spacing
   *
   * @returns Number of paragraphs updated
   * @example
   * ```typescript
   * const count = doc.normalizeNumberedLists();
   * console.log(`Normalized ${count} numbered list items`);
   * ```
   */
  normalizeNumberedLists(): number {
    let count = 0;

    // Create a standard numbered list
    const standardNumId = this.numberingManager.createNumberedList(3, [
      "decimal",
      "lowerLetter",
      "lowerRoman",
    ]);

    // Collect all paragraphs with numbering and identify numbered lists
    const paragraphs = this.getParagraphs();
    const numberedParas: { para: Paragraph; level: number }[] = [];

    for (const para of paragraphs) {
      const numbering = para.getNumbering();
      if (!numbering) continue;

      // Get the abstract numbering for this numId
      const instance = this.numberingManager.getInstance(numbering.numId);
      if (!instance) continue;

      const abstractNum = this.numberingManager.getAbstractNumbering(
        instance.getAbstractNumId()
      );
      if (!abstractNum) continue;

      // Check if level 0 is a numbered format (not bullet)
      const level0 = abstractNum.getLevel(0);
      if (!level0) continue;

      const format = level0.getFormat();
      // Numbered formats: decimal, lowerRoman, upperRoman, lowerLetter, upperLetter, etc.
      if (format !== "bullet") {
        numberedParas.push({ para, level: numbering.level });
      }
    }

    // Apply standard numbering to all numbered paragraphs
    for (const { para, level } of numberedParas) {
      para.setNumbering(standardNumId, level);
      count++;
    }

    // Clean up orphaned numbering definitions
    this.cleanupUnusedNumbering();

    return count;
  }

  /**
   * Normalizes all bullet lists in the document to use consistent formatting
   *
   * Creates a standard bullet list format and applies it to all bullet lists:
   * - Level 0: • (bullet)
   * - Level 1: ○ (circle)
   * - Level 2: ■ (square)
   * - Consistent indentation and spacing
   *
   * @returns Number of paragraphs updated
   * @example
   * ```typescript
   * const count = doc.normalizeBulletLists();
   * console.log(`Normalized ${count} bullet list items`);
   * ```
   */
  normalizeBulletLists(): number {
    let count = 0;

    // Create a standard bullet list with custom bullets
    const standardNumId = this.numberingManager.createBulletList(3, [
      "•",
      "○",
      "■",
    ]);

    // Collect all paragraphs with numbering and identify bullet lists
    const paragraphs = this.getParagraphs();
    const bulletParas: { para: Paragraph; level: number }[] = [];

    for (const para of paragraphs) {
      const numbering = para.getNumbering();
      if (!numbering) continue;

      // Get the abstract numbering for this numId
      const instance = this.numberingManager.getInstance(numbering.numId);
      if (!instance) continue;

      const abstractNum = this.numberingManager.getAbstractNumbering(
        instance.getAbstractNumId()
      );
      if (!abstractNum) continue;

      // Check if level 0 is a bullet format
      const level0 = abstractNum.getLevel(0);
      if (!level0) continue;

      const format = level0.getFormat();
      if (format === "bullet") {
        bulletParas.push({ para, level: numbering.level });
      }
    }

    // Apply standard bullet numbering to all bullet paragraphs
    for (const { para, level } of bulletParas) {
      para.setNumbering(standardNumId, level);
      count++;
    }

    // Clean up orphaned numbering definitions
    this.cleanupUnusedNumbering();

    return count;
  }

  /**
   * Cleans up unused numbering definitions
   *
   * Removes numbering instances and abstract numberings that are no longer
   * referenced by any paragraphs in the document. This prevents corruption
   * from orphaned numbering definitions.
   *
   * @private
   */
  private cleanupUnusedNumbering(): void {
    // Collect all numIds currently used by paragraphs
    const usedNumIds = new Set<number>();
    const paragraphs = this.getParagraphs();

    for (const para of paragraphs) {
      const numbering = para.getNumbering();
      if (numbering) {
        usedNumIds.add(numbering.numId);
      }
    }

    // Clean up unused numbering definitions
    this.numberingManager.cleanupUnusedNumbering(usedNumIds);
  }

  /**
   * Removes all headers and footers from the document
   *
   * Clears all header and footer references including:
   * - Default header/footer
   * - First page header/footer
   * - Even page header/footer
   *
   * @returns Number of headers and footers removed
   * @example
   * ```typescript
   * const count = doc.removeAllHeadersFooters();
   * console.log(`Removed ${count} headers and footers`);
   * ```
   */
  removeAllHeadersFooters(): number {
    let totalCount = 0;

    // Step 1: Remove relationship entries for headers and footers
    const headerRels = this.relationshipManager.getRelationshipsByType(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    );
    const footerRels = this.relationshipManager.getRelationshipsByType(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    );

    for (const rel of [...headerRels, ...footerRels]) {
      this.relationshipManager.removeRelationship(rel.getId());
      totalCount++;
    }

    // Step 2: Find and delete all header/footer XML files from ZIP archive
    // Scan for word/header*.xml and word/footer*.xml files
    const allFiles = this.zipHandler.getFilePaths();
    const headerFooterFiles = allFiles.filter((path) =>
      path.match(/^word\/(header|footer)\d+\.xml$/i)
    );

    for (const filePath of headerFooterFiles) {
      this.zipHandler.removeFile(filePath);
    }

    // Step 3: Clear internal references
    this.headerFooterManager.clear();

    // Clear section header/footer references
    const section = this.section;
    const sectionProps = (section as any).properties as any;

    if (sectionProps.headers) {
      sectionProps.headers = {};
    }

    if (sectionProps.footers) {
      sectionProps.footers = {};
    }

    // Disable title page if it was enabled for first page headers/footers
    if (sectionProps.titlePage) {
      sectionProps.titlePage = false;
    }

    return totalCount;
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
    formats?: Array<"decimal" | "lowerLetter" | "lowerRoman">
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
   * Gets the framework's standard indentation for a list level
   *
   * The framework uses a consistent indentation scheme:
   * - leftIndent: 720 * (level + 1) twips
   * - hangingIndent: 360 twips
   *
   * @param level The level (0-8)
   * @returns Object with leftIndent and hangingIndent in twips
   * @example
   * ```typescript
   * const indent = doc.getStandardIndentation(0);
   * // Returns: { leftIndent: 720, hangingIndent: 360 }
   * ```
   */
  getStandardIndentation(level: number): {
    leftIndent: number;
    hangingIndent: number;
  } {
    return this.numberingManager.getStandardIndentation(level);
  }

  /**
   * Sets custom indentation for a specific level in a numbering definition
   *
   * This updates the indentation for a specific level across ALL paragraphs
   * that use this numId and level combination.
   *
   * @param numId The numbering instance ID
   * @param level The level to modify (0-8)
   * @param leftIndent Left indentation in twips
   * @param hangingIndent Hanging indentation in twips (optional, defaults to 360)
   * @returns This document for chaining
   * @example
   * ```typescript
   * // Set level 0 to 0.5 inch left, 0.25 inch hanging
   * doc.setListIndentation(1, 0, 720, 360);
   * ```
   */
  setListIndentation(
    numId: number,
    level: number,
    leftIndent: number,
    hangingIndent?: number
  ): this {
    this.numberingManager.setListIndentation(
      numId,
      level,
      leftIndent,
      hangingIndent
    );
    return this;
  }

  /**
   * Normalizes indentation for all lists in the document
   *
   * Applies standard indentation to every numbering instance:
   * - leftIndent: 720 * (level + 1) twips
   * - hangingIndent: 360 twips
   *
   * This ensures consistent spacing across all lists in the document.
   *
   * @returns Number of numbering instances updated
   * @example
   * ```typescript
   * const count = doc.normalizeAllListIndentation();
   * console.log(`Normalized ${count} lists`);
   * ```
   */
  normalizeAllListIndentation(): number {
    return this.numberingManager.normalizeAllListIndentation();
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
  setPageSize(
    width: number,
    height: number,
    orientation?: "portrait" | "landscape"
  ): this {
    this.section.setPageSize(width, height, orientation);
    return this;
  }

  /**
   * Sets page orientation
   * @param orientation Page orientation
   * @returns This document for chaining
   */
  setPageOrientation(orientation: "portrait" | "landscape"): this {
    this.section.setOrientation(orientation);
    return this;
  }

  /**
   * Sets margins
   * @param margins Margin properties
   * @returns This document for chaining
   */
  setMargins(margins: {
    top: number;
    bottom: number;
    left: number;
    right: number;
    header?: number;
    footer?: number;
    gutter?: number;
  }): this {
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
    const relationship = this.relationshipManager.addHeader(
      `${header.getFilename(1)}`
    );

    // Register with manager
    this.headerFooterManager.registerHeader(header, relationship.getId());

    // Link to section
    this.section.setHeaderReference("default", relationship.getId());

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
    const relationship = this.relationshipManager.addHeader(
      `${header.getFilename(this.headerFooterManager.getHeaderCount() + 1)}`
    );

    // Register with manager
    this.headerFooterManager.registerHeader(header, relationship.getId());

    // Link to section
    this.section.setHeaderReference("first", relationship.getId());

    return this;
  }

  /**
   * Sets the even page header (requires different odd/even pages)
   * @param header The header to set
   * @returns This document for chaining
   */
  setEvenPageHeader(header: Header): this {
    // Generate relationship for header
    const relationship = this.relationshipManager.addHeader(
      `${header.getFilename(this.headerFooterManager.getHeaderCount() + 1)}`
    );

    // Register with manager
    this.headerFooterManager.registerHeader(header, relationship.getId());

    // Link to section
    this.section.setHeaderReference("even", relationship.getId());

    return this;
  }

  /**
   * Sets the default footer for the document
   * @param footer The footer to set
   * @returns This document for chaining
   */
  setFooter(footer: Footer): this {
    // Generate relationship for footer
    const relationship = this.relationshipManager.addFooter(
      `${footer.getFilename(1)}`
    );

    // Register with manager
    this.headerFooterManager.registerFooter(footer, relationship.getId());

    // Link to section
    this.section.setFooterReference("default", relationship.getId());

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
    const relationship = this.relationshipManager.addFooter(
      `${footer.getFilename(this.headerFooterManager.getFooterCount() + 1)}`
    );

    // Register with manager
    this.headerFooterManager.registerFooter(footer, relationship.getId());

    // Link to section
    this.section.setFooterReference("first", relationship.getId());

    return this;
  }

  /**
   * Sets the even page footer (requires different odd/even pages)
   * @param footer The footer to set
   * @returns This document for chaining
   */
  setEvenPageFooter(footer: Footer): this {
    // Generate relationship for footer
    const relationship = this.relationshipManager.addFooter(
      `${footer.getFilename(this.headerFooterManager.getFooterCount() + 1)}`
    );

    // Register with manager
    this.headerFooterManager.registerFooter(footer, relationship.getId());

    // Link to section
    this.section.setFooterReference("even", relationship.getId());

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
    const target = `media/image${
      this.imageManager.getImageCount() + 1
    }.${image.getExtension()}`;
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
    this.zipHandler.updateFile("word/_rels/document.xml.rels", xml);
  }

  /**
   * Saves all comments to the ZIP archive
   */
  private saveComments(): void {
    // Only save comments.xml if there are comments
    if (this.commentManager.getCount() > 0) {
      const xml = this.commentManager.generateCommentsXml();
      this.zipHandler.addFile("word/comments.xml", xml);

      // Add comments relationship
      this.relationshipManager.addComments();
    }
  }

  /**
   * Saves custom properties to the ZIP archive
   */
  private saveCustomProperties(): void {
    // Only save custom.xml if there are custom properties
    if (
      this.properties.customProperties &&
      Object.keys(this.properties.customProperties).length > 0
    ) {
      const customXml = this.generator.generateCustomProps(
        this.properties.customProperties
      );
      if (customXml) {
        this.zipHandler.addFile("docProps/custom.xml", customXml);
      }
    }
  }

  /**
   * Updates [Content_Types].xml to include image extensions, headers/footers, comments, and custom properties
   * Preserves entries for files that exist in the loaded document
   */
  private updateContentTypesWithImagesHeadersFootersAndComments(): void {
    const hasCustomProps =
      this.properties.customProperties &&
      Object.keys(this.properties.customProperties).length > 0;

    const contentTypes =
      this.generator.generateContentTypesWithImagesHeadersFootersAndComments(
        this.imageManager,
        this.headerFooterManager,
        this.commentManager,
        this.zipHandler, // Pass zipHandler to check file existence
        undefined, // fontManager (optional)
        hasCustomProps // Flag to include custom.xml override
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
  addBookmarkToParagraph(
    paragraph: Paragraph,
    bookmarkOrName: Bookmark | string
  ): Bookmark {
    const bookmark =
      typeof bookmarkOrName === "string"
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
  createInsertion(author: string, content: Run | Run[], date?: Date): Revision {
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
  createDeletion(author: string, content: Run | Run[], date?: Date): Revision {
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
    const revision = this.createRevisionFromText("insert", author, text, date);
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
    const revision = this.createRevisionFromText("delete", author, text, date);
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
    return this.commentManager.createReply(
      parentCommentId,
      author,
      content,
      initials
    );
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
    const comment =
      typeof commentOrAuthor === "string"
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
  getCommentThread(
    commentId: number
  ): { comment: Comment; replies: Comment[] } | undefined {
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
    const updates: Array<{
      hyperlink: Hyperlink;
      newUrl: string;
      relationshipId?: string;
    }> = [];

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
              relationshipId: content.getRelationshipId(),
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

      // Convert Buffer to string for text files
      // ZipWriter stores all content as Buffer internally, but DocumentPart expects string for text
      let content: string | Buffer = file.content;
      if (!file.isBinary && Buffer.isBuffer(file.content)) {
        content = file.content.toString('utf-8');
      }

      return {
        name: partName,
        content,
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
      const contentTypesXml = this.zipHandler.getFileAsString(
        "[Content_Types].xml"
      );
      if (!contentTypesXml) {
        return contentTypes;
      }

      // Parse content types XML
      // Match Default elements (by extension)
      const defaultPattern =
        /<Default\s+Extension="([^"]+)"\s+ContentType="([^"]+)"/g;
      let match;
      while ((match = defaultPattern.exec(contentTypesXml)) !== null) {
        if (match[1] && match[2]) {
          contentTypes.set(`.${match[1]}`, match[2]);
        }
      }

      // Match Override elements (by part name)
      const overridePattern =
        /<Override\s+PartName="([^"]+)"\s+ContentType="([^"]+)"/g;
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
      if (typeof part.content === "string") {
        return part.content;
      }

      // If Buffer, decode as UTF-8 (standard for XML files)
      if (Buffer.isBuffer(part.content)) {
        return part.content.toString("utf8");
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
    if (typeof xmlContent !== "string") {
      throw new Error("XML content must be a string");
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
  async addContentType(
    partNameOrExtension: string,
    contentType: string
  ): Promise<boolean> {
    try {
      let contentTypesXml = this.zipHandler.getFileAsString(
        "[Content_Types].xml"
      );
      if (!contentTypesXml) {
        return false;
      }

      const isExtension = partNameOrExtension.startsWith(".");

      if (isExtension) {
        // Add as Default element (for extensions)
        const extension = partNameOrExtension.substring(1);

        // Check if already exists
        const existingPattern = new RegExp(
          `<Default\\s+Extension="${extension}"\\s+ContentType="[^"]+"/?>`,
          "g"
        );
        if (existingPattern.test(contentTypesXml)) {
          // Update existing
          contentTypesXml = contentTypesXml.replace(
            existingPattern,
            `<Default Extension="${extension}" ContentType="${contentType}"/>`
          );
        } else {
          // Add new before closing tag
          contentTypesXml = contentTypesXml.replace(
            "</Types>",
            `  <Default Extension="${extension}" ContentType="${contentType}"/>\n</Types>`
          );
        }
      } else {
        // Add as Override element (for specific parts)
        const partName = partNameOrExtension.startsWith("/")
          ? partNameOrExtension
          : `/${partNameOrExtension}`;

        // Check if already exists
        const existingPattern = new RegExp(
          `<Override\\s+PartName="${partName.replace(
            /[.*+?^${}()|[\]\\]/g,
            "\\$&"
          )}"\\s+ContentType="[^"]+"/?>`,
          "g"
        );
        if (existingPattern.test(contentTypesXml)) {
          // Update existing
          contentTypesXml = contentTypesXml.replace(
            existingPattern,
            `<Override PartName="${partName}" ContentType="${contentType}"/>`
          );
        } else {
          // Add new before closing tag
          contentTypesXml = contentTypesXml.replace(
            "</Types>",
            `  <Override PartName="${partName}" ContentType="${contentType}"/>\n</Types>`
          );
        }
      }

      // Update the content types file
      this.zipHandler.updateFile("[Content_Types].xml", contentTypesXml);
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
      const relsPaths = this.zipHandler
        .getFilePaths()
        .filter((path) => path.endsWith(".rels"));

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
          const relationshipElements = XMLParser.extractElements(
            relsContent,
            "Relationship"
          );

          for (const relElement of relationshipElements) {
            const rel: ParsedRelationship = {};

            // Extract attributes using XMLParser
            const id = XMLParser.extractAttribute(relElement, "Id");
            const type = XMLParser.extractAttribute(relElement, "Type");
            const target = XMLParser.extractAttribute(relElement, "Target");
            const targetMode = XMLParser.extractAttribute(
              relElement,
              "TargetMode"
            );

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
  async getRelationships(
    partName: string
  ): Promise<
    Array<{ id?: string; type?: string; target?: string; targetMode?: string }>
  > {
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
      const relationshipElements = XMLParser.extractElements(
        relsContent,
        "Relationship"
      );

      for (const relElement of relationshipElements) {
        const rel: ParsedRelationship = {};

        // Extract attributes using XMLParser
        const id = XMLParser.extractAttribute(relElement, "Id");
        const type = XMLParser.extractAttribute(relElement, "Type");
        const target = XMLParser.extractAttribute(relElement, "Target");
        const targetMode = XMLParser.extractAttribute(relElement, "TargetMode");

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
      const contentTypesXml = this.zipHandler.getFileAsString(
        "[Content_Types].xml"
      );
      if (!contentTypesXml) {
        return undefined;
      }

      // Check for specific override
      const overridePattern = new RegExp(
        `<Override\\s+PartName="${partName.replace(
          /[.*+?^${}()|[\]\\]/g,
          "\\$&"
        )}"\\s+ContentType="([^"]+)"`,
        "i"
      );
      const overrideMatch = contentTypesXml.match(overridePattern);
      if (overrideMatch) {
        return overrideMatch[1];
      }

      // Check for extension default
      const ext = partName.substring(partName.lastIndexOf("."));
      if (ext) {
        const defaultPattern = new RegExp(
          `<Default\\s+Extension="${ext.substring(
            1
          )}"\\s+ContentType="([^"]+)"`,
          "i"
        );
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
  findText(
    text: string,
    options?: { caseSensitive?: boolean; wholeWord?: boolean }
  ): Array<{
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
          const wordPattern = new RegExp(
            `\\b${searchText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`,
            caseSensitive ? "g" : "gi"
          );
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
          while (
            (startIndex = compareText.indexOf(searchText, startIndex)) !== -1
          ) {
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
                const compareText = caseSensitive
                  ? runText
                  : runText.toLowerCase();

                if (wholeWord) {
                  // Create word boundary regex
                  const wordPattern = new RegExp(
                    `\\b${searchText.replace(
                      /[.*+?^${}()|[\]\\]/g,
                      "\\$&"
                    )}\\b`,
                    caseSensitive ? "g" : "gi"
                  );
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
                  while (
                    (startIndex = compareText.indexOf(
                      searchText,
                      startIndex
                    )) !== -1
                  ) {
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
  replaceText(
    find: string,
    replace: string,
    options?: { caseSensitive?: boolean; wholeWord?: boolean }
  ): number {
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
            `\\b${find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`,
            caseSensitive ? "g" : "gi"
          );
          const matches = originalText.match(wordPattern);
          if (matches) {
            replacementCount += matches.length;
            newText = originalText.replace(wordPattern, replace);
          }
        } else {
          // Simple substring replacement
          const searchPattern = new RegExp(
            find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
            caseSensitive ? "g" : "gi"
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
        const words = text.split(/\s+/).filter((word) => word.length > 0);
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
              const words = text.split(/\s+/).filter((word) => word.length > 0);
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
        totalChars += text.replace(/\s/g, "").length;
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
              totalChars += text.replace(/\s/g, "").length;
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

    if (typeof paragraphOrIndex === "number") {
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

    if (typeof tableOrIndex === "number") {
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
   * Inserts a table at a specific position
   * @param index - The position to insert at (0-based)
   * @param table - The table to insert
   * @returns This document for chaining
   * @example
   * ```typescript
   * const table = new Table(2, 3);
   * doc.insertTableAt(5, table);  // Insert at position 5
   * ```
   */
  insertTableAt(index: number, table: Table): this {
    if (index < 0) {
      index = 0;
    } else if (index > this.bodyElements.length) {
      index = this.bodyElements.length;
    }

    this.bodyElements.splice(index, 0, table);
    return this;
  }

  /**
   * Inserts a Table of Contents at a specific position
   * @param index - The position to insert at (0-based)
   * @param toc - The TableOfContentsElement to insert
   * @returns This document for chaining
   * @example
   * ```typescript
   * const toc = TableOfContentsElement.createStandard();
   * doc.insertTocAt(0, toc);  // Insert at beginning
   * ```
   */
  insertTocAt(index: number, toc: TableOfContentsElement): this {
    if (index < 0) {
      index = 0;
    } else if (index > this.bodyElements.length) {
      index = this.bodyElements.length;
    }

    this.bodyElements.splice(index, 0, toc);
    return this;
  }

  /**
   * Replaces a paragraph at a specific position
   * @param index - The position to replace at (0-based)
   * @param paragraph - The new paragraph
   * @returns True if replaced, false if index invalid
   * @example
   * ```typescript
   * const newPara = new Paragraph();
   * newPara.addText('Replacement text');
   * doc.replaceParagraphAt(3, newPara);
   * ```
   */
  replaceParagraphAt(index: number, paragraph: Paragraph): boolean {
    if (index >= 0 && index < this.bodyElements.length) {
      const element = this.bodyElements[index];
      if (element instanceof Paragraph) {
        this.bodyElements[index] = paragraph;
        return true;
      }
    }
    return false;
  }

  /**
   * Replaces a table at a specific position
   * @param index - The position to replace at (0-based)
   * @param table - The new table
   * @returns True if replaced, false if index invalid or not a table
   * @example
   * ```typescript
   * const newTable = new Table(3, 4);
   * doc.replaceTableAt(2, newTable);
   * ```
   */
  replaceTableAt(index: number, table: Table): boolean {
    if (index >= 0 && index < this.bodyElements.length) {
      const element = this.bodyElements[index];
      if (element instanceof Table) {
        this.bodyElements[index] = table;
        return true;
      }
    }
    return false;
  }

  /**
   * Moves an element from one position to another
   * @param fromIndex - Current position (0-based)
   * @param toIndex - Target position (0-based)
   * @returns True if moved, false if indices invalid
   * @example
   * ```typescript
   * doc.moveElement(5, 2);  // Move element from position 5 to position 2
   * ```
   */
  moveElement(fromIndex: number, toIndex: number): boolean {
    if (
      fromIndex < 0 ||
      fromIndex >= this.bodyElements.length ||
      toIndex < 0 ||
      toIndex >= this.bodyElements.length
    ) {
      return false;
    }

    const [element] = this.bodyElements.splice(fromIndex, 1);
    this.bodyElements.splice(toIndex, 0, element!);
    return true;
  }

  /**
   * Swaps two elements' positions
   * @param index1 - First element position (0-based)
   * @param index2 - Second element position (0-based)
   * @returns True if swapped, false if indices invalid
   * @example
   * ```typescript
   * doc.swapElements(2, 5);  // Swap elements at positions 2 and 5
   * ```
   */
  swapElements(index1: number, index2: number): boolean {
    if (
      index1 < 0 ||
      index1 >= this.bodyElements.length ||
      index2 < 0 ||
      index2 >= this.bodyElements.length
    ) {
      return false;
    }

    const temp = this.bodyElements[index1];
    this.bodyElements[index1] = this.bodyElements[index2]!;
    this.bodyElements[index2] = temp!;
    return true;
  }

  /**
   * Removes a Table of Contents element at a specific position
   * @param index - The position to remove (0-based)
   * @returns True if removed, false if index invalid or not a TOC
   * @example
   * ```typescript
   * doc.removeTocAt(0);  // Remove TOC at beginning
   * ```
   */
  removeTocAt(index: number): boolean {
    if (index >= 0 && index < this.bodyElements.length) {
      const element = this.bodyElements[index];
      if (element instanceof TableOfContentsElement) {
        this.bodyElements.splice(index, 1);
        return true;
      }
    }
    return false;
  }

  /**
   * Gets all hyperlinks in the document
   * @returns Array of hyperlinks with their containing paragraph
   */
  getHyperlinks(): Array<{ hyperlink: Hyperlink; paragraph: Paragraph }> {
    const hyperlinks: Array<{ hyperlink: Hyperlink; paragraph: Paragraph }> =
      [];

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
          const cellParagraphs =
            cell instanceof TableCell ? cell.getParagraphs() : [];
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
    zipHandler.addFile(
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n' +
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n' +
        '  <Default Extension="xml" ContentType="application/xml"/>\n' +
        '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n' +
        "</Types>"
    );

    // _rels/.rels - minimal
    zipHandler.addFile(
      "_rels/.rels",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' +
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n' +
        "</Relationships>"
    );

    // word/document.xml - empty body
    zipHandler.addFile(
      "word/document.xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n' +
        "  <w:body/>\n" +
        "</w:document>"
    );

    // word/_rels/document.xml.rels - empty relationships
    zipHandler.addFile(
      "word/_rels/document.xml.rels",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    );

    return doc;
  }
}
