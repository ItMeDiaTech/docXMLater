/**
 * Document - High-level API for creating and managing Word documents
 * Provides a simple interface for creating DOCX files without managing ZIP and XML manually
 */

import { Bookmark } from "../elements/Bookmark";
import { BookmarkManager } from "../elements/BookmarkManager";
import { Comment } from "../elements/Comment";
import { CommentManager } from "../elements/CommentManager";
import { EndnoteManager } from "../elements/EndnoteManager";
import { Footer } from "../elements/Footer";
import { FootnoteManager } from "../elements/FootnoteManager";
import { Header } from "../elements/Header";
import { HeaderFooterManager } from "../elements/HeaderFooterManager";
import { Hyperlink } from "../elements/Hyperlink";
import { Image } from "../elements/Image";
import { ImageManager } from "../elements/ImageManager";
import { ImageRun } from "../elements/ImageRun";
import { Paragraph } from "../elements/Paragraph";
import { RangeMarker } from "../elements/RangeMarker";
import { Revision, RevisionType } from "../elements/Revision";
import { RevisionManager } from "../elements/RevisionManager";
import { Run, RunFormatting } from "../elements/Run";
import { Section } from "../elements/Section";
import { StructuredDocumentTag } from "../elements/StructuredDocumentTag";
import { Table, TableBorder } from "../elements/Table";
import { TableCell } from "../elements/TableCell";
import { TableOfContents, TOCProperties } from "../elements/TableOfContents";
import { TableOfContentsElement } from "../elements/TableOfContentsElement";
import { NumberingManager } from "../formatting/NumberingManager";
import { Style, StyleProperties } from "../formatting/Style";
import { StylesManager } from "../formatting/StylesManager";
import { FormatOptions, StyleApplyOptions } from "../types/formatting";
import {
  ApplyCustomFormattingOptions,
  Heading2Config,
  StyleConfig,
} from "../types/styleConfig";
import { ILogger, defaultLogger } from "../utils/logger";
import { XMLParser } from "../xml/XMLParser";
import { ZipHandler } from "../zip/ZipHandler";
import { DOCX_PATHS } from "../zip/types";
import { DocumentGenerator } from "./DocumentGenerator";
import { DocumentParser } from "./DocumentParser";
import { DocumentValidator } from "./DocumentValidator";
import { RelationshipManager } from "./RelationshipManager";

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
  // Reserved for future implementation - placeholder properties
  private _footnoteManager?: any;
  private _endnoteManager?: any;

  // Helper classes for parsing, generation, and validation
  private parser: DocumentParser;
  private generator: DocumentGenerator;
  private validator: DocumentValidator;
  private logger: ILogger;

  // Track changes settings
  private trackChangesEnabled: boolean = false;
  private trackFormatting: boolean = true;
  private revisionViewSettings: {
    showInsertionsAndDeletions: boolean;
    showFormatting: boolean;
    showInkAnnotations: boolean;
  } = {
    showInsertionsAndDeletions: true,
    showFormatting: true,
    showInkAnnotations: true,
  };

  // TOC auto-population setting
  private autoPopulateTOCs: boolean = false;

  private rsidRoot?: string;
  private rsids: Set<string> = new Set();
  private documentProtection?: {
    edit: "readOnly" | "comments" | "trackedChanges" | "forms";
    enforcement: boolean;
    cryptProviderType?: string;
    cryptAlgorithmClass?: string;
    cryptAlgorithmType?: string;
    cryptAlgorithmSid?: number;
    cryptSpinCount?: number;
    hash?: string;
    salt?: string;
  };

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
      this.generator.generateSettings({
        trackChangesEnabled: this.trackChangesEnabled,
        trackFormatting: this.trackFormatting,
        revisionView: this.revisionViewSettings,
        rsidRoot: this.rsidRoot,
        rsids: this.getRsids(),
        documentProtection: this.documentProtection,
      })
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
    // Clear built-in styles and use document-specific styles
    if (result.styles && result.styles.length > 0) {
      // Clear all existing styles to avoid conflicts with built-in styles
      this.stylesManager.clear();

      // Add all parsed styles from the document
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

    // Load parsed headers and footers into HeaderFooterManager
    // This preserves existing headers/footers from the document
    const { headers, footers } = await this.parser.parseHeadersAndFooters(
      this.zipHandler,
      result.section,
      this.relationshipManager,
      this.imageManager
    );

    // Register parsed headers with manager
    for (const { header, relationshipId, filename } of headers) {
      // Manually register without creating new relationships
      // (relationships already exist from loaded document)
      this.headerFooterManager.registerHeader(header, relationshipId);
      header.setHeaderId(relationshipId);
    }

    // Register parsed footers with manager
    for (const { footer, relationshipId, filename } of footers) {
      // Manually register without creating new relationships
      // (relationships already exist from loaded document)
      this.headerFooterManager.registerFooter(footer, relationshipId);
      footer.setFooterId(relationshipId);
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
   * Creates and adds a pre-populated Table of Contents
   * The TOC will display actual heading entries when document is first opened in Word
   * Field structure is preserved so users can still right-click "Update Field"
   *
   * @param title - Optional TOC title
   * @param options - Optional TOC configuration
   * @returns This document for chaining
   *
   * @example
   * const doc = Document.create();
   * doc.createParagraph('Chapter 1').setStyle('Heading1');
   * doc.createParagraph('Section 1.1').setStyle('Heading2');
   *
   * // Create pre-populated TOC
   * doc.createPrePopulatedTableOfContents('Contents', {
   *   levels: 3,
   *   useHyperlinks: true
   * });
   *
   * await doc.save('output.docx');
   * // TOC entries are already visible when opened!
   */
  createPrePopulatedTableOfContents(
    title?: string,
    options?: Partial<TOCProperties>
  ): this {
    const toc = TableOfContents.create({
      title: title || "Table of Contents",
      ...options,
    });
    this.addTableOfContents(toc);
    this.setAutoPopulateTOCs(true);
    return this;
  }

  /**
   * Enables or disables automatic TOC population during save
   * When enabled, TOCs will be pre-populated with heading entries
   *
   * @param enabled - Whether to auto-populate TOCs
   * @returns This document for chaining
   *
   * @example
   * doc.setAutoPopulateTOCs(true);
   * doc.createTableOfContents();
   * await doc.save('output.docx');
   * // TOC is pre-populated with entries
   */
  setAutoPopulateTOCs(enabled: boolean): this {
    this.autoPopulateTOCs = enabled;
    return this;
  }

  /**
   * Checks if automatic TOC population is enabled
   * @returns True if TOCs will be auto-populated on save
   */
  isAutoPopulateTOCsEnabled(): boolean {
    return this.autoPopulateTOCs;
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
   * Gets all paragraphs in the document recursively
   * Includes paragraphs in tables and SDTs
   * @returns Array of all paragraphs
   */
  getAllParagraphs(): Paragraph[] {
    const result: Paragraph[] = [];

    for (const element of this.bodyElements) {
      if (element instanceof Paragraph) {
        result.push(element);
      } else if (element instanceof Table) {
        // Recurse into table cells
        for (const row of element.getRows()) {
          for (const cell of row.getCells()) {
            result.push(...cell.getParagraphs());
          }
        }
      } else if (element instanceof StructuredDocumentTag) {
        // Recurse into SDT content
        for (const content of element.getContent()) {
          if (content instanceof Paragraph) {
            result.push(content);
          } else if (content instanceof Table) {
            // Recurse into tables inside SDTs
            for (const row of content.getRows()) {
              for (const cell of row.getCells()) {
                result.push(...cell.getParagraphs());
              }
            }
          }
          // Handle nested SDTs recursively
          // Note: This could be extended to handle deeply nested SDTs
        }
      }
    }

    return result;
  }

  /**
   * Gets all tables in the document recursively
   * Includes tables inside SDTs
   * @returns Array of all tables
   */
  getAllTables(): Table[] {
    const result: Table[] = [];

    for (const element of this.bodyElements) {
      if (element instanceof Table) {
        result.push(element);
      } else if (element instanceof StructuredDocumentTag) {
        // Recurse into SDT content
        for (const content of element.getContent()) {
          if (content instanceof Table) {
            result.push(content);
          }
          // Note: Could extend to handle tables nested in SDTs inside SDTs
        }
      }
    }

    return result;
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
   * Adds a body element (paragraph, table, SDT, etc.) to the document
   * @param element - The body element to add
   * @returns This document for chaining
   */
  addBodyElement(element: BodyElement): this {
    this.bodyElements.push(element);
    return this;
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

      // Auto-populate TOCs if enabled
      if (this.autoPopulateTOCs) {
        await this.populateTOCsInFile(tempPath);
      }

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

      // Auto-populate TOCs if enabled
      if (this.autoPopulateTOCs) {
        const docXml = this.zipHandler.getFileAsString("word/document.xml");
        if (docXml) {
          const populatedXml = this.populateAllTOCsInXML(docXml);
          if (populatedXml !== docXml) {
            this.zipHandler.updateFile("word/document.xml", populatedXml);
          }
        }
      }

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
   * Applies borders to all cells in all tables throughout the document
   * Useful for ensuring consistent border styling across all tables
   * @param border - Border definition to apply to all sides of every cell
   * @returns Number of tables updated
   * @example
   * ```typescript
   * // Add single black borders to all tables
   * const count = doc.applyBordersToAllTables({
   *   style: 'single',
   *   size: 8,
   *   color: '000000'
   * });
   * console.log(`Applied borders to ${count} tables`);
   * ```
   */
  applyBordersToAllTables(border: TableBorder): number {
    const tables = this.getAllTables();

    for (const table of tables) {
      table.setAllBorders(border);
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
   * Applies comprehensive formatting to all tables in the document
   *
   * This helper function provides a one-call solution for standardizing table formatting:
   * - Apply black borders to all cells (always applied to all tables)
   * - Set table width to autofit to window (always applied to all tables)
   * - Format first row as header (shading, bold, centered, custom font/spacing)
   * - Apply consistent cell margins to all cells
   * - Recolor and format cells with existing shading
   * - Optionally skip shading/formatting for single-cell (1x1) tables
   *
   * Shading and formatting logic (for tables > 1x1):
   * - Row 0: Apply the specified color + full formatting (bold, Verdana 12pt, centered, 3pt spacing)
   * - Other rows: Cells with existing color (NOT white) receive the same color + formatting
   * - Cells with no color or white color remain unchanged
   *
   * @param colorOrOptions Hex color for multi-cell tables, or options object
   * @param multiCellColor Optional second color for multi-cell tables (when first param is 1x1 color)
   * @returns Statistics about tables processed
   *
   * @example
   * // Use default gray color (E9E9E9) for multi-cell tables, skip 1x1 tables
   * const result = doc.applyStandardTableFormatting();
   *
   * @example
   * // Custom color for multi-cell tables only
   * const result = doc.applyStandardTableFormatting('D9D9D9');
   *
   * @example
   * // Two colors: first for 1x1 tables, second for multi-cell tables
   * const result = doc.applyStandardTableFormatting('BFBFBF', 'E9E9E9');
   *
   * @example
   * // Advanced: Full customization
   * const result = doc.applyStandardTableFormatting({
   *   singleCellShading: 'BFBFBF',  // Gray for 1x1 tables
   *   headerRowShading: '4472C4',   // Blue for multi-cell table headers
   *   headerRowFormatting: {
   *     bold: true,
   *     alignment: 'center',
   *     font: 'Arial',
   *     size: 14
   *   }
   * });
   * console.log(`Processed ${result.tablesProcessed} tables`);
   * console.log(`Formatted ${result.headerRowsFormatted} header rows`);
   * console.log(`Recolored ${result.cellsRecolored} cells`);
   * console.log(`Shaded ${result.singleCellTablesShaded} single-cell tables`);
   */
  public applyStandardTableFormatting(
    colorOrOptions?:
      | string
      | {
          /** Autofit tables to window width (DEPRECATED: now always enabled) */
          autofitToWindow?: boolean;
          /** Single-cell (1x1) table shading color */
          singleCellShading?: string;
          /** Header row background color for multi-cell tables (default: 'E9E9E9') */
          headerRowShading?: string;
          /** Header row text formatting */
          headerRowFormatting?: {
            bold?: boolean;
            alignment?: "left" | "center" | "right" | "justify";
            font?: string;
            size?: number;
            color?: string;
            spacingBefore?: number;
            spacingAfter?: number;
          };
          /** Cell margins for all cells in twips */
          cellMargins?: {
            top?: number;
            bottom?: number;
            left?: number;
            right?: number;
          };
          /** Skip 1x1 tables (DEPRECATED: use singleCellShading instead) */
          skipSingleCellTables?: boolean;
        },
    multiCellColor?: string
  ): {
    tablesProcessed: number;
    headerRowsFormatted: number;
    cellsRecolored: number;
    singleCellTablesShaded: number;
  } {
    // Handle different parameter combinations
    let options: any;
    if (typeof colorOrOptions === "string") {
      if (multiCellColor) {
        // Two colors provided: applyStandardTableFormatting('BFBFBF', 'E9E9E9')
        options = {
          singleCellShading: colorOrOptions,
          headerRowShading: multiCellColor,
        };
      } else {
        // One color provided: backwards compatible - for multi-cell only
        options = { headerRowShading: colorOrOptions };
      }
    } else {
      options = colorOrOptions;
    }

    // Default values
    const singleCellShading = options?.singleCellShading?.toUpperCase();
    const headerRowShading = (
      options?.headerRowShading || "E9E9E9"
    ).toUpperCase();
    const headerRowFormatting = {
      bold: options?.headerRowFormatting?.bold !== false,
      alignment: options?.headerRowFormatting?.alignment || ("center" as const),
      font: options?.headerRowFormatting?.font || "Verdana",
      size: options?.headerRowFormatting?.size || 12,
      color: options?.headerRowFormatting?.color || "000000",
      spacingBefore: options?.headerRowFormatting?.spacingBefore ?? 60,
      spacingAfter: options?.headerRowFormatting?.spacingAfter ?? 60,
    };
    const cellMargins = {
      top: options?.cellMargins?.top ?? 0,
      bottom: options?.cellMargins?.bottom ?? 0,
      left: options?.cellMargins?.left ?? 115, // 0.08 inches
      right: options?.cellMargins?.right ?? 115, // 0.08 inches
    };
    const skipSingleCellTables =
      options?.skipSingleCellTables !== false && !singleCellShading;

    // Statistics
    let tablesProcessed = 0;
    let headerRowsFormatted = 0;
    let cellsRecolored = 0;
    let singleCellTablesShaded = 0;

    // Get all tables
    const tables = this.getAllTables();

    for (const table of tables) {
      const rowCount = table.getRowCount();
      const columnCount = table.getColumnCount();

      // Apply borders to all cells (always applied to all tables)
      table.setAllBorders({
        style: "single",
        size: 4,
        color: "000000",
      });

      // Set table width to autofit to window (always applied to all tables)
      table.setLayout("auto");
      table.setWidthType("pct");
      table.setWidth(5000);

      // Handle 1x1 (single-cell) tables separately
      const is1x1Table = rowCount === 1 && columnCount === 1;
      if (is1x1Table) {
        if (singleCellShading) {
          // Apply single-cell shading color
          const singleCell = table.getRow(0)?.getCell(0);
          if (singleCell) {
            singleCell.setShading({ fill: singleCellShading });
            singleCellTablesShaded++;
          }
        }
        // Skip further processing for 1x1 tables
        tablesProcessed++;
        continue;
      }

      // Format first row (header) for multi-cell tables
      const firstRow = table.getRow(0);
      if (firstRow) {
        for (const cell of firstRow.getCells()) {
          // Set header shading
          cell.setShading({ fill: headerRowShading });

          // Set margins
          cell.setMargins(cellMargins);

          // Format paragraphs and runs in header (skip list paragraphs)
          for (const para of cell.getParagraphs()) {
            // Skip paragraphs that are part of numbered or bulleted lists
            const numPr = para.getFormatting().numbering;
            if (
              numPr &&
              (numPr.level !== undefined || numPr.numId !== undefined)
            ) {
              continue; // Preserve list formatting
            }

            para.setAlignment(headerRowFormatting.alignment);
            para.setSpaceBefore(headerRowFormatting.spacingBefore);
            para.setSpaceAfter(headerRowFormatting.spacingAfter);

            for (const run of para.getRuns()) {
              if (headerRowFormatting.bold) run.setBold(true);
              run.setFont(headerRowFormatting.font, headerRowFormatting.size);
              run.setColor(headerRowFormatting.color);
            }
          }
        }
        headerRowsFormatted++;
      }

      // Format remaining rows (data rows)
      for (let i = 1; i < rowCount; i++) {
        const row = table.getRow(i);
        if (!row) continue;

        for (const cell of row.getCells()) {
          // Always apply margins
          cell.setMargins(cellMargins);

          // Apply shading and formatting to cells with existing color (not white)
          const currentShading = cell.getFormatting().shading;
          const currentColor = currentShading?.fill?.toUpperCase();

          // Check if color is a valid 6-character hex code (not 'auto' or other special values)
          const isValidHexColor = /^[0-9A-F]{6}$/i.test(currentColor || "");

          if (currentColor && currentColor !== "FFFFFF" && isValidHexColor) {
            // Apply the color passed to the method
            cell.setShading({ fill: headerRowShading });
            cellsRecolored++;

            // Always apply formatting when shading is applied (but skip list paragraphs)
            for (const para of cell.getParagraphs()) {
              // Skip paragraphs that are part of numbered or bulleted lists
              const numPr = para.getFormatting().numbering;
              if (
                numPr &&
                (numPr.level !== undefined || numPr.numId !== undefined)
              ) {
                continue; // Preserve list formatting
              }

              para.setAlignment("center");
              para.setSpaceBefore(60); // 3pt
              para.setSpaceAfter(60); // 3pt

              for (const run of para.getRuns()) {
                run.setBold(true);
                run.setFont("Verdana", 12);
                run.setColor("000000");
              }
            }
          }
        }
      }

      tablesProcessed++;
    }

    return {
      tablesProcessed,
      headerRowsFormatted,
      cellsRecolored,
      singleCellTablesShaded,
    };
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
   * Applies standard formatting to all bullet lists in the document
   *
   * Standardizes bullet lists with:
   * - Alternating bullet symbols: • (solid) for even levels, ○ (open) for odd levels
   * - Indentation: 0.5" increments (720 twips per level)
   * - Hanging indent: 0.25" (360 twips)
   * - Paragraph text: Verdana 12pt
   * - Spacing: 0pt before, 3pt after (60 twips)
   * - Contextual spacing enabled (no spacing between same-type paragraphs)
   *
   * @returns Number of bullet lists updated
   * @example
   * ```typescript
   * const doc = await Document.load('document.docx');
   * const count = doc.applyStandardListFormatting();
   * console.log(`Standardized ${count} bullet lists`);
   * await doc.save('document-formatted.docx');
   * ```
   */
  applyStandardListFormatting(): number {
    const instances = this.numberingManager.getAllInstances();
    let count = 0;

    for (const instance of instances) {
      const abstractNumId = instance.getAbstractNumId();
      const abstractNum =
        this.numberingManager.getAbstractNumbering(abstractNumId);

      if (!abstractNum) continue;

      // Only process bullet lists (skip numbered lists)
      const level0 = abstractNum.getLevel(0);
      if (!level0 || level0.getFormat() !== "bullet") continue;

      // Update all 9 levels (0-8) with standard formatting
      for (let levelIndex = 0; levelIndex < 9; levelIndex++) {
        const numLevel = abstractNum.getLevel(levelIndex);
        if (!numLevel) continue;

        // Alternate bullets: even levels = solid (•), odd levels = open (○)
        const bullet = levelIndex % 2 === 0 ? "•" : "○";
        numLevel.setText(bullet);

        // Set bullet font to Arial (Unicode bullets require a regular font, not Symbol)
        numLevel.setFont("Arial");

        // Set bullet size to 12pt (24 half-points)
        numLevel.setFontSize(24);

        // Indentation: 0.5" per level (720 twips)
        // Level 0 = 720, Level 1 = 1440, Level 2 = 2160, etc.
        numLevel.setLeftIndent(720 * (levelIndex + 1));

        // Hanging indent: 0.25" (360 twips) for all levels
        numLevel.setHangingIndent(360);
      }

      // Apply paragraph formatting to all paragraphs using this list
      this.applyFormattingToListParagraphs(instance.getNumId());
      count++;
    }

    return count;
  }

  /**
   * Applies standard formatting to all numbered lists in the document
   *
   * Standardizes numbered lists with (preserves existing numbering format):
   * - Indentation: 0.5" increments (720 twips per level)
   * - Hanging indent: 0.25" (360 twips)
   * - Number font: Verdana 12pt
   * - Paragraph text: Verdana 12pt
   * - Spacing: 0pt before, 3pt after (60 twips)
   * - Contextual spacing enabled (no spacing between same-type paragraphs)
   *
   * Note: This preserves the existing numbering format (decimal, roman, etc.)
   * and only standardizes the visual formatting. To change numbering formats,
   * use normalizeNumberedLists() instead.
   *
   * @returns Number of numbered lists updated
   * @example
   * ```typescript
   * const doc = await Document.load('document.docx');
   * const count = doc.applyStandardNumberedListFormatting();
   * console.log(`Standardized ${count} numbered lists`);
   * await doc.save('document-formatted.docx');
   * ```
   */
  applyStandardNumberedListFormatting(): number {
    const instances = this.numberingManager.getAllInstances();
    let count = 0;

    for (const instance of instances) {
      const abstractNumId = instance.getAbstractNumId();
      const abstractNum =
        this.numberingManager.getAbstractNumbering(abstractNumId);

      if (!abstractNum) continue;

      // Only process numbered lists (skip bullet lists)
      const level0 = abstractNum.getLevel(0);
      if (!level0 || level0.getFormat() === "bullet") continue;

      // Update all 9 levels (0-8) with standard formatting
      for (let levelIndex = 0; levelIndex < 9; levelIndex++) {
        const numLevel = abstractNum.getLevel(levelIndex);
        if (!numLevel) continue;

        // Set number font to Verdana 12pt
        numLevel.setFont("Verdana");
        numLevel.setFontSize(24); // 12pt = 24 half-points

        // Indentation: 0.5" per level (720 twips)
        // Level 0 = 720, Level 1 = 1440, Level 2 = 2160, etc.
        numLevel.setLeftIndent(720 * (levelIndex + 1));

        // Hanging indent: 0.25" (360 twips) for all levels
        numLevel.setHangingIndent(360);

        // Set alignment to left
        numLevel.setAlignment("left");
      }

      // Apply paragraph formatting to all paragraphs using this list
      this.applyFormattingToListParagraphs(instance.getNumId());
      count++;
    }

    return count;
  }

  /**
   * Applies formatting to all paragraphs that use a specific numbering instance
   * Sets font, spacing, and contextual spacing properties
   * @param numId The numbering instance ID
   * @private
   */
  private applyFormattingToListParagraphs(numId: number): void {
    const paragraphs = this.getAllParagraphs();

    for (const para of paragraphs) {
      const numbering = para.getNumbering();
      if (numbering?.numId === numId) {
        // Apply font to all runs in the paragraph
        const runs = para.getRuns();
        for (const run of runs) {
          run.setFont("Verdana", 12);
        }

        // Apply paragraph spacing
        para.setSpaceBefore(0); // 0pt before
        para.setSpaceAfter(60); // 3pt after (60 twips)
        para.setContextualSpacing(true); // No spacing between same-type paragraphs

        // Clear paragraph-level indentation so numbering definition indentation takes effect
        // Paragraph-level indentation overrides numbering indentation, so we need to remove it
        para.formatting.indentation = undefined;
      }
    }
  }

  /**
   * Checks if a paragraph is contained within a table cell
   * @param para The paragraph to check
   * @returns Object with inTable boolean and cell reference if found
   * @private
   */
  private isParagraphInTable(para: Paragraph): {
    inTable: boolean;
    cell?: TableCell;
  } {
    const allTables = this.getAllTables();

    for (const table of allTables) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          const cellParas = cell.getParagraphs();
          for (const cellPara of cellParas) {
            if (cellPara === para) {
              return { inTable: true, cell };
            }
          }
        }
      }
    }

    return { inTable: false };
  }

  /**
   * Wraps a paragraph in a 1x1 table and applies cell formatting
   * @param para The paragraph to wrap
   * @param options Formatting options for the table cell
   * @returns The created table
   * @private
   */
  private wrapParagraphInTable(
    para: Paragraph,
    options: {
      shading?: string;
      marginTop?: number;
      marginBottom?: number;
      marginLeft?: number;
      marginRight?: number;
      tableWidthPercent?: number;
    }
  ): Table {
    // Find the paragraph index in bodyElements
    const paraIndex = this.bodyElements.indexOf(para);
    if (paraIndex === -1) {
      throw new Error("Paragraph not found in document body elements");
    }

    // Create 1x1 table
    const table = new Table(1, 1);
    const cell = table.getCell(0, 0);

    if (!cell) {
      throw new Error("Failed to get cell from newly created table");
    }

    // Move paragraph to cell
    // Remove paragraph from document body
    this.bodyElements.splice(paraIndex, 1);

    // Add paragraph to cell
    cell.addParagraph(para);

    // Apply cell formatting
    if (options.shading) {
      cell.setShading({ fill: options.shading });
    }

    if (
      options.marginTop !== undefined ||
      options.marginBottom !== undefined ||
      options.marginLeft !== undefined ||
      options.marginRight !== undefined
    ) {
      cell.setMargins({
        top: options.marginTop ?? 100,
        bottom: options.marginBottom ?? 100,
        left: options.marginLeft ?? 100,
        right: options.marginRight ?? 100,
      });
    }

    // Set table width (percentage of page width)
    if (options.tableWidthPercent !== undefined) {
      table.setWidth(options.tableWidthPercent);
      table.setWidthType("pct");
    }

    // Insert table where paragraph was
    this.bodyElements.splice(paraIndex, 0, table);

    return table;
  }

  /**
   * Creates and applies custom styles to the document
   * Applies three custom styles: Header1, Header2, and Normal
   *
   * - Header1: 18pt black bold Verdana, left aligned, 0pt before / 12pt after
   * - Header2: 14pt black bold Verdana, left aligned, 6pt before/after, wrapped in gray table
   * - Normal: 12pt Verdana, left aligned, 3pt before/after
   *
   * @returns Object with counts of modified paragraphs
   */
  public applyCustomStylesToDocument(): {
    heading1: number;
    heading2: number;
    normal: number;
  } {
    const counts = { heading1: 0, heading2: 0, normal: 0 };

    // Create custom styles
    const header1Style = Style.create({
      styleId: "CustomHeader1",
      name: "Custom Header 1",
      type: "paragraph",
      basedOn: "Normal",
      runFormatting: {
        font: "Verdana",
        size: 18,
        bold: true,
        color: "000000",
      },
      paragraphFormatting: {
        alignment: "left",
        spacing: {
          before: 0, // 0pt before
          after: 240, // 12pt after (240 twips)
        },
      },
    });

    const header2Style = Style.create({
      styleId: "CustomHeader2",
      name: "Custom Header 2",
      type: "paragraph",
      basedOn: "Normal",
      runFormatting: {
        font: "Verdana",
        size: 14,
        bold: true,
        color: "000000",
      },
      paragraphFormatting: {
        alignment: "left",
        spacing: {
          before: 120, // 6pt before (120 twips)
          after: 120, // 6pt after (120 twips)
        },
      },
    });

    const normalStyle = Style.create({
      styleId: "CustomNormal",
      name: "Custom Normal",
      type: "paragraph",
      basedOn: "Normal",
      runFormatting: {
        font: "Verdana",
        size: 12,
        color: "000000",
      },
      paragraphFormatting: {
        alignment: "left",
        spacing: {
          before: 60, // 3pt before (60 twips)
          after: 60, // 3pt after (60 twips)
        },
      },
    });

    // Add styles to document
    this.addStyle(header1Style);
    this.addStyle(header2Style);
    this.addStyle(normalStyle);

    // Get all paragraphs (this will be modified as we wrap some in tables)
    // We need to work with a copy since we'll be modifying bodyElements
    const allParagraphs = this.getAllParagraphs();

    // Process each paragraph
    for (const para of allParagraphs) {
      const currentStyle = para.getStyle();

      // Match Heading1 / Heading 1
      if (currentStyle === "Heading1" || currentStyle === "Heading 1") {
        para.setStyle("CustomHeader1");
        counts.heading1++;
      }
      // Match Heading2 / Heading 2
      else if (currentStyle === "Heading2" || currentStyle === "Heading 2") {
        para.setStyle("CustomHeader2");

        // Check if paragraph is in a table
        const { inTable, cell } = this.isParagraphInTable(para);

        if (inTable && cell) {
          // Paragraph is already in a table - apply cell formatting
          cell.setShading({ fill: "BFBFBF" });
          cell.setMargins({ top: 0, bottom: 0, left: 101, right: 101 });

          // Set table width to 100%
          const table = this.getAllTables().find((t) => {
            for (const row of t.getRows()) {
              for (const c of row.getCells()) {
                if (c === cell) return true;
              }
            }
            return false;
          });
          if (table) {
            table.setWidth(5000); // 100% width
            table.setWidthType("pct");
          }
        } else {
          // Paragraph is not in a table - wrap it
          this.wrapParagraphInTable(para, {
            shading: "BFBFBF",
            marginTop: 0,
            marginBottom: 0,
            marginLeft: 101,
            marginRight: 101,
            tableWidthPercent: 5000, // 100% in Word's percentage units
          });
        }

        counts.heading2++;
      }
      // Match Normal (explicit "Normal" style or undefined/no style set)
      // In Word, paragraphs without an explicit pStyle default to "Normal"
      else if (currentStyle === "Normal" || currentStyle === undefined) {
        para.setStyle("CustomNormal");
        counts.normal++;
      }
    }

    return counts;
  }

  // Default style configurations for applyCustomFormattingToExistingStyles()
  private static readonly DEFAULT_HEADING1_CONFIG: StyleConfig = {
    run: {
      font: "Verdana",
      size: 18,
      bold: true,
      color: "000000",
    },
    paragraph: {
      alignment: "left",
      spacing: { before: 0, after: 240, line: 240, lineRule: "auto" },
    },
  };

  private static readonly DEFAULT_HEADING2_CONFIG: Heading2Config = {
    run: {
      font: "Verdana",
      size: 14,
      bold: true,
      color: "000000",
    },
    paragraph: {
      alignment: "left",
      spacing: { before: 120, after: 120, line: 240, lineRule: "auto" },
    },
    tableOptions: {
      shading: "BFBFBF",
      marginTop: 0,
      marginBottom: 0,
      marginLeft: 115,
      marginRight: 115,
      tableWidthPercent: 5000,
    },
  };

  private static readonly DEFAULT_HEADING3_CONFIG: StyleConfig = {
    run: {
      font: "Verdana",
      size: 12,
      bold: true,
      color: "000000",
    },
    paragraph: {
      alignment: "left",
      spacing: { before: 60, after: 60, line: 240, lineRule: "auto" },
    },
  };

  private static readonly DEFAULT_NORMAL_CONFIG: StyleConfig = {
    run: {
      font: "Verdana",
      size: 12,
      color: "000000",
    },
    paragraph: {
      alignment: "left",
      spacing: { before: 60, after: 60, line: 240, lineRule: "auto" },
    },
  };

  private static readonly DEFAULT_LIST_PARAGRAPH_CONFIG: StyleConfig = {
    run: {
      font: "Verdana",
      size: 12,
      color: "000000",
    },
    paragraph: {
      alignment: "left",
      spacing: { before: 0, after: 60, line: 240, lineRule: "auto" },
      indentation: { left: 360, hanging: 360 },
      contextualSpacing: true,
    },
  };

  /**
   * Modifies the existing Heading1, Heading2, Heading3, Normal, and List Paragraph style definitions
   * with custom formatting. This approach preserves the original style names while updating their formatting.
   *
   * **Key Feature**: Only properties explicitly provided in options will override current style values.
   * The method reads the ACTUAL current values from the style objects, not hardcoded defaults.
   * This allows you to change just one property (like font) while keeping all other existing values.
   *
   * Per ECMA-376 §17.7.2, direct formatting in document.xml ALWAYS overrides
   * style definitions in styles.xml. This method clears conflicting direct
   * formatting from paragraphs to allow style modifications to take effect.
   *
   * Fallback defaults (when no options provided OR style doesn't exist):
   * - Heading1: 18pt black bold Verdana, left aligned, 0pt before / 12pt after, single line spacing, no italic/underline
   * - Heading2: 14pt black bold Verdana, left aligned, 6pt before/after, single line spacing, wrapped in gray tables (0.08" margins), no italic/underline
   * - Heading3: 12pt black bold Verdana, left aligned, 3pt before/after, single line spacing, no table wrapping, no italic/underline
   * - Normal: 12pt Verdana, left aligned, 3pt before/after, single line spacing, no italic/underline
   * - List Paragraph: 12pt Verdana, left aligned, 0pt before / 3pt after, single line spacing, 0.25" bullet indent / 0.50" text indent, contextual spacing enabled, no italic/underline
   *
   * Heading2 table wrapping behavior:
   * - Empty Heading2 paragraphs are skipped (not wrapped in tables)
   * - Heading2 paragraphs already in tables have their cell formatted (shading, margins, width)
   * - Heading2 paragraphs not in tables are wrapped in new 1x1 tables
   * - Table appearance is configurable via options.heading2.tableOptions
   *
   * @param options - Optional custom formatting configuration for each style. Properties merge with current values.
   * @returns Object indicating which styles were successfully modified
   *
   * @example
   * // Use default formatting (applies Verdana defaults)
   * doc.applyCustomFormattingToExistingStyles();
   *
   * @example
   * // Just change font, keep all other existing style values
   * doc.applyCustomFormattingToExistingStyles({
   *   heading1: {
   *     run: { font: 'Arial' } // Only overrides font, keeps existing size/bold/color/etc.
   *   }
   * });
   *
   * @example
   * // Comprehensive custom formatting
   * doc.applyCustomFormattingToExistingStyles({
   *   heading1: {
   *     run: { font: 'Arial', size: 16, bold: true, color: '000000' },
   *     paragraph: { spacing: { before: 0, after: 200, line: 240, lineRule: 'auto' } }
   *   },
   *   heading2: {
   *     run: { font: 'Arial', size: 14, bold: true },
   *     paragraph: { spacing: { before: 100, after: 100 } },
   *     tableOptions: { shading: '808080', marginLeft: 150, marginRight: 150 }
   *   }
   * });
   */
  public applyCustomFormattingToExistingStyles(
    options?: ApplyCustomFormattingOptions
  ): {
    heading1: boolean;
    heading2: boolean;
    heading3: boolean;
    normal: boolean;
    listParagraph: boolean;
  } {
    const results = {
      heading1: false,
      heading2: false,
      heading3: false,
      normal: false,
      listParagraph: false,
    };

    // Get existing styles from StylesManager
    const heading1 = this.stylesManager.getStyle("Heading1");
    const heading2 = this.stylesManager.getStyle("Heading2");
    const heading3 = this.stylesManager.getStyle("Heading3");
    const normal = this.stylesManager.getStyle("Normal");
    const listParagraph = this.stylesManager.getStyle("ListParagraph");

    // Merge provided options with ACTUAL current style values (not hardcoded defaults)
    // This allows users to only specify properties they want to change
    const h1Config = {
      run: {
        ...(heading1?.getRunFormatting() ||
          Document.DEFAULT_HEADING1_CONFIG.run),
        ...options?.heading1?.run,
      },
      paragraph: {
        ...(heading1?.getParagraphFormatting() ||
          Document.DEFAULT_HEADING1_CONFIG.paragraph),
        ...options?.heading1?.paragraph,
      },
    };
    const h2Config = {
      run: {
        ...(heading2?.getRunFormatting() ||
          Document.DEFAULT_HEADING2_CONFIG.run),
        ...options?.heading2?.run,
      },
      paragraph: {
        ...(heading2?.getParagraphFormatting() ||
          Document.DEFAULT_HEADING2_CONFIG.paragraph),
        ...options?.heading2?.paragraph,
      },
      tableOptions: {
        ...Document.DEFAULT_HEADING2_CONFIG.tableOptions,
        ...options?.heading2?.tableOptions,
      },
    };
    const h3Config = {
      run: {
        ...(heading3?.getRunFormatting() ||
          Document.DEFAULT_HEADING3_CONFIG.run),
        ...options?.heading3?.run,
      },
      paragraph: {
        ...(heading3?.getParagraphFormatting() ||
          Document.DEFAULT_HEADING3_CONFIG.paragraph),
        ...options?.heading3?.paragraph,
      },
    };
    const normalConfig = {
      run: {
        ...(normal?.getRunFormatting() || Document.DEFAULT_NORMAL_CONFIG.run),
        ...options?.normal?.run,
      },
      paragraph: {
        ...(normal?.getParagraphFormatting() ||
          Document.DEFAULT_NORMAL_CONFIG.paragraph),
        ...options?.normal?.paragraph,
      },
    };
    const listParaConfig = {
      run: {
        ...(listParagraph?.getRunFormatting() ||
          Document.DEFAULT_LIST_PARAGRAPH_CONFIG.run),
        ...options?.listParagraph?.run,
      },
      paragraph: {
        ...(listParagraph?.getParagraphFormatting() ||
          Document.DEFAULT_LIST_PARAGRAPH_CONFIG.paragraph),
        ...options?.listParagraph?.paragraph,
      },
    };

    // Extract preserve blank lines option (defaults to true)
    const preserveBlankLines =
      options?.preserveBlankLinesAfterHeader2Tables ?? true;

    // Modify Heading1 definition
    if (heading1 && h1Config.run && h1Config.paragraph) {
      if (h1Config.run) heading1.setRunFormatting(h1Config.run);
      if (h1Config.paragraph)
        heading1.setParagraphFormatting(h1Config.paragraph);
      results.heading1 = true;
    }

    // Modify Heading2 definition
    if (heading2 && h2Config.run && h2Config.paragraph) {
      if (h2Config.run) heading2.setRunFormatting(h2Config.run);
      if (h2Config.paragraph)
        heading2.setParagraphFormatting(h2Config.paragraph);
      results.heading2 = true;
    }

    // Modify Heading3 definition
    if (heading3 && h3Config.run && h3Config.paragraph) {
      if (h3Config.run) heading3.setRunFormatting(h3Config.run);
      if (h3Config.paragraph)
        heading3.setParagraphFormatting(h3Config.paragraph);
      results.heading3 = true;
    }

    // Modify Normal definition
    if (normal && normalConfig.run && normalConfig.paragraph) {
      if (normalConfig.run) normal.setRunFormatting(normalConfig.run);
      if (normalConfig.paragraph)
        normal.setParagraphFormatting(normalConfig.paragraph);
      results.normal = true;
    }

    // Modify List Paragraph definition
    if (listParagraph && listParaConfig.run && listParaConfig.paragraph) {
      if (listParaConfig.run)
        listParagraph.setRunFormatting(listParaConfig.run);
      if (listParaConfig.paragraph)
        listParagraph.setParagraphFormatting(listParaConfig.paragraph);
      results.listParagraph = true;
    }

    // Extract preserve flags from configurations
    const h1Preserve = {
      bold: h1Config.run?.preserveBold ?? false,
      italic: h1Config.run?.preserveItalic ?? false,
      underline: h1Config.run?.preserveUnderline ?? false,
    };
    const h2Preserve = {
      bold: h2Config.run?.preserveBold ?? false,
      italic: h2Config.run?.preserveItalic ?? false,
      underline: h2Config.run?.preserveUnderline ?? false,
    };
    const h3Preserve = {
      bold: h3Config.run?.preserveBold ?? false,
      italic: h3Config.run?.preserveItalic ?? false,
      underline: h3Config.run?.preserveUnderline ?? false,
    };
    const normalPreserve = {
      bold: normalConfig.run?.preserveBold ?? true,
      italic: normalConfig.run?.preserveItalic ?? false,
      underline: normalConfig.run?.preserveUnderline ?? false,
    };
    const listParaPreserve = {
      bold: listParaConfig.run?.preserveBold ?? true,
      italic: listParaConfig.run?.preserveItalic ?? false,
      underline: listParaConfig.run?.preserveUnderline ?? false,
    };

    // Clear direct formatting from affected paragraphs and wrap Heading2 in tables
    // Use a Set to track processed paragraphs and prevent duplicate wrapping
    const processedParagraphs = new Set<Paragraph>();

    // Get all paragraphs ONCE before modifications to prevent processing duplicates
    const allParas = this.getAllParagraphs();

    for (const para of allParas) {
      // Skip if already processed
      if (processedParagraphs.has(para)) {
        continue;
      }

      const styleId = para.getStyle();

      // Process Heading1 paragraphs
      if (styleId === "Heading1" && heading1) {
        para.clearDirectFormattingConflicts(heading1);

        // Apply formatting to all runs, respecting preserve flags
        for (const run of para.getRuns()) {
          if (!h1Preserve.bold) {
            run.setBold(h1Config.run?.bold ?? false);
          }
          if (!h1Preserve.italic) {
            run.setItalic(h1Config.run?.italic ?? false);
          }
          if (!h1Preserve.underline) {
            run.setUnderline(h1Config.run?.underline ? "single" : false);
          }
        }

        // Update paragraph mark properties to match configuration
        if (para.formatting.paragraphMarkRunProperties) {
          const markProps = para.formatting.paragraphMarkRunProperties;
          if (
            !h1Preserve.bold &&
            h1Config.run?.bold === false &&
            markProps.bold
          ) {
            delete markProps.bold;
          }
          if (
            !h1Preserve.italic &&
            h1Config.run?.italic === false &&
            markProps.italic
          ) {
            delete markProps.italic;
          }
          if (
            !h1Preserve.underline &&
            h1Config.run?.underline === false &&
            markProps.underline
          ) {
            delete markProps.underline;
          }
        }

        processedParagraphs.add(para);
      }

      // Process Heading2 paragraphs
      else if (styleId === "Heading2" && heading2) {
        // Check if paragraph has actual text content (skip empty paragraphs)
        const hasContent = para
          .getRuns()
          .some((run) => run.getText().trim().length > 0);

        if (!hasContent) {
          // Skip empty Heading2 paragraphs - don't wrap them in tables
          processedParagraphs.add(para);
          continue;
        }

        // Clear direct formatting first
        para.clearDirectFormattingConflicts(heading2);

        // Apply formatting to all runs, respecting preserve flags
        for (const run of para.getRuns()) {
          if (!h2Preserve.bold) {
            run.setBold(h2Config.run?.bold ?? false);
          }
          if (!h2Preserve.italic) {
            run.setItalic(h2Config.run?.italic ?? false);
          }
          if (!h2Preserve.underline) {
            run.setUnderline(h2Config.run?.underline ? "single" : false);
          }
        }

        // Update paragraph mark properties to match configuration
        if (para.formatting.paragraphMarkRunProperties) {
          const markProps = para.formatting.paragraphMarkRunProperties;
          if (
            !h2Preserve.bold &&
            h2Config.run?.bold === false &&
            markProps.bold
          ) {
            delete markProps.bold;
          }
          if (
            !h2Preserve.italic &&
            h2Config.run?.italic === false &&
            markProps.italic
          ) {
            delete markProps.italic;
          }
          if (
            !h2Preserve.underline &&
            h2Config.run?.underline === false &&
            markProps.underline
          ) {
            delete markProps.underline;
          }
        }

        // Check if paragraph is in a table
        const { inTable, cell } = this.isParagraphInTable(para);

        if (inTable && cell) {
          // Paragraph is already in a table - apply cell formatting using config
          if (h2Config.tableOptions?.shading) {
            cell.setShading({ fill: h2Config.tableOptions.shading });
          }
          if (
            h2Config.tableOptions?.marginTop !== undefined ||
            h2Config.tableOptions?.marginBottom !== undefined ||
            h2Config.tableOptions?.marginLeft !== undefined ||
            h2Config.tableOptions?.marginRight !== undefined
          ) {
            cell.setMargins({
              top: h2Config.tableOptions.marginTop ?? 0,
              bottom: h2Config.tableOptions.marginBottom ?? 0,
              left: h2Config.tableOptions.marginLeft ?? 115,
              right: h2Config.tableOptions.marginRight ?? 115,
            });
          }

          // Set table width using config
          if (h2Config.tableOptions?.tableWidthPercent) {
            const table = this.getAllTables().find((t) => {
              for (const row of t.getRows()) {
                for (const c of row.getCells()) {
                  if (c === cell) return true;
                }
              }
              return false;
            });
            if (table) {
              table.setWidth(h2Config.tableOptions.tableWidthPercent);
              table.setWidthType("pct");
            }
          }
        } else {
          // Paragraph is not in a table - wrap it using config
          const table = this.wrapParagraphInTable(para, {
            shading: h2Config.tableOptions?.shading ?? "BFBFBF",
            marginTop: h2Config.tableOptions?.marginTop ?? 0,
            marginBottom: h2Config.tableOptions?.marginBottom ?? 0,
            marginLeft: h2Config.tableOptions?.marginLeft ?? 115,
            marginRight: h2Config.tableOptions?.marginRight ?? 115,
            tableWidthPercent: h2Config.tableOptions?.tableWidthPercent ?? 5000,
          });

          // Add blank paragraph after table for spacing (only if not already present)
          const tableIndex = this.bodyElements.indexOf(table);
          if (tableIndex !== -1) {
            // Check if the next element has any content (text, hyperlinks, images, etc.)
            const nextElement = this.bodyElements[tableIndex + 1];

            // Check if next element is truly blank (no content at all)
            const isNextElementBlank = (() => {
              if (!(nextElement instanceof Paragraph)) return false;

              const content = nextElement.getContent();
              if (!content || content.length === 0) return true;

              // Check if all content items are empty
              for (const item of content) {
                // Hyperlinks count as content
                if ((item as any).constructor.name === "Hyperlink") {
                  return false;
                }
                // Images count as content (check for Image class when implemented)
                // Runs with text count as content
                if ((item as any).getText) {
                  const text = (item as any).getText().trim();
                  if (text !== "") return false;
                }
              }

              return true; // All content is empty
            })();

            // Only add blank paragraph if next element is truly blank or doesn't exist
            if (!isNextElementBlank) {
              const blankPara = Paragraph.create();
              // Add explicit spacing to ensure visibility in Word (120 twips = 6pt)
              blankPara.setSpaceAfter(120);
              // Mark as preserved if option is enabled (defaults to true)
              if (preserveBlankLines) {
                blankPara.setPreserved(true);
              }
              this.bodyElements.splice(tableIndex + 1, 0, blankPara);
            }
          }
        }

        processedParagraphs.add(para);
      }

      // Process Heading3 paragraphs
      else if (styleId === "Heading3" && heading3) {
        para.clearDirectFormattingConflicts(heading3);

        // Apply formatting to all runs, respecting preserve flags
        for (const run of para.getRuns()) {
          if (!h3Preserve.bold) {
            run.setBold(h3Config.run?.bold ?? false);
          }
          if (!h3Preserve.italic) {
            run.setItalic(h3Config.run?.italic ?? false);
          }
          if (!h3Preserve.underline) {
            run.setUnderline(h3Config.run?.underline ? "single" : false);
          }
        }

        // Update paragraph mark properties to match configuration
        if (para.formatting.paragraphMarkRunProperties) {
          const markProps = para.formatting.paragraphMarkRunProperties;
          if (
            !h3Preserve.bold &&
            h3Config.run?.bold === false &&
            markProps.bold
          ) {
            delete markProps.bold;
          }
          if (
            !h3Preserve.italic &&
            h3Config.run?.italic === false &&
            markProps.italic
          ) {
            delete markProps.italic;
          }
          if (
            !h3Preserve.underline &&
            h3Config.run?.underline === false &&
            markProps.underline
          ) {
            delete markProps.underline;
          }
        }

        processedParagraphs.add(para);
      }

      // Process List Paragraph paragraphs
      else if (styleId === "ListParagraph" && listParagraph) {
        // Save formatting that should be preserved BEFORE clearing
        const preservedFormatting = para.getRuns().map((run) => {
          const fmt = run.getFormatting();
          return {
            run: run,
            bold: listParaPreserve.bold ? fmt.bold : undefined,
            italic: listParaPreserve.italic ? fmt.italic : undefined,
            underline: listParaPreserve.underline ? fmt.underline : undefined,
          };
        });

        para.clearDirectFormattingConflicts(listParagraph);

        // Restore preserved formatting AFTER clearing
        for (const saved of preservedFormatting) {
          if (saved.bold !== undefined) {
            saved.run.setBold(saved.bold);
          }
          if (saved.italic !== undefined) {
            saved.run.setItalic(saved.italic);
          }
          if (saved.underline !== undefined) {
            saved.run.setUnderline(saved.underline);
          }
        }

        // Apply formatting to all runs, respecting preserve flags
        for (const run of para.getRuns()) {
          if (!listParaPreserve.bold) {
            run.setBold(listParaConfig.run?.bold ?? false);
          }
          if (!listParaPreserve.italic) {
            run.setItalic(listParaConfig.run?.italic ?? false);
          }
          if (!listParaPreserve.underline) {
            run.setUnderline(listParaConfig.run?.underline ? "single" : false);
          }
        }

        // Update paragraph mark properties to match configuration
        if (para.formatting.paragraphMarkRunProperties) {
          const markProps = para.formatting.paragraphMarkRunProperties;
          if (
            !listParaPreserve.bold &&
            listParaConfig.run?.bold === false &&
            markProps.bold
          ) {
            delete markProps.bold;
          }
          if (
            !listParaPreserve.italic &&
            listParaConfig.run?.italic === false &&
            markProps.italic
          ) {
            delete markProps.italic;
          }
          if (
            !listParaPreserve.underline &&
            listParaConfig.run?.underline === false &&
            markProps.underline
          ) {
            delete markProps.underline;
          }
        }

        processedParagraphs.add(para);
      }

      // Process Normal paragraphs (including undefined style which defaults to Normal)
      else if ((styleId === "Normal" || styleId === undefined) && normal) {
        // Save formatting that should be preserved BEFORE clearing
        const preservedFormatting = para.getRuns().map((run) => {
          const fmt = run.getFormatting();
          return {
            run: run,
            bold: normalPreserve.bold ? fmt.bold : undefined,
            italic: normalPreserve.italic ? fmt.italic : undefined,
            underline: normalPreserve.underline ? fmt.underline : undefined,
          };
        });

        para.clearDirectFormattingConflicts(normal);

        // Restore preserved formatting AFTER clearing
        for (const saved of preservedFormatting) {
          if (saved.bold !== undefined) {
            saved.run.setBold(saved.bold);
          }
          if (saved.italic !== undefined) {
            saved.run.setItalic(saved.italic);
          }
          if (saved.underline !== undefined) {
            saved.run.setUnderline(saved.underline);
          }
        }

        // Apply formatting to all runs, respecting preserve flags
        for (const run of para.getRuns()) {
          if (!normalPreserve.bold) {
            run.setBold(normalConfig.run?.bold ?? false);
          }
          if (!normalPreserve.italic) {
            run.setItalic(normalConfig.run?.italic ?? false);
          }
          if (!normalPreserve.underline) {
            run.setUnderline(normalConfig.run?.underline ? "single" : false);
          }
        }

        // Update paragraph mark properties to match configuration
        if (para.formatting.paragraphMarkRunProperties) {
          const markProps = para.formatting.paragraphMarkRunProperties;
          if (
            !normalPreserve.bold &&
            normalConfig.run?.bold === false &&
            markProps.bold
          ) {
            delete markProps.bold;
          }
          if (
            !normalPreserve.italic &&
            normalConfig.run?.italic === false &&
            markProps.italic
          ) {
            delete markProps.italic;
          }
          if (
            !normalPreserve.underline &&
            normalConfig.run?.underline === false &&
            markProps.underline
          ) {
            delete markProps.underline;
          }
        }

        processedParagraphs.add(para);
      }
    }

    return results;
  }

  /**
   * Helper function to apply formatting from Style objects
   *
   * This is a convenience wrapper that accepts Style objects, extracts their properties,
   * and converts them to configuration format before calling applyCustomFormattingToExistingStyles().
   *
   * Style objects are matched by their styleId property:
   * - 'Heading1' → applies to Heading1 style
   * - 'Heading2' → applies to Heading2 style (with optional table options)
   * - 'Heading3' → applies to Heading3 style
   * - 'Normal' → applies to Normal style
   * - 'ListParagraph' → applies to List Paragraph style
   *
   * Other styleIds are ignored.
   *
   * @param styles - Variable number of Style objects to apply
   * @returns Object indicating which styles were successfully modified
   *
   * @example
   * // Create Style objects with your desired formatting
   * const h1 = new Style({
   *   styleId: 'Heading1',
   *   name: 'Heading 1',
   *   type: 'paragraph',
   *   runFormatting: { font: 'Arial', size: 16, bold: true },
   *   paragraphFormatting: { spacing: { before: 0, after: 200 } }
   * });
   *
   * const h2 = Style.createHeadingStyle(2);
   * h2.setRunFormatting({ font: 'Arial', size: 14, bold: true });
   * h2.setHeading2TableOptions({ shading: '808080', marginLeft: 150, marginRight: 150 });
   *
   * // Apply to document
   * doc.applyStylesFromObjects(h1, h2);
   *
   * @example
   * // Just modify one style
   * const myNormal = new Style({
   *   styleId: 'Normal',
   *   name: 'Normal',
   *   type: 'paragraph',
   *   runFormatting: { font: 'Times New Roman', size: 12 }
   * });
   * doc.applyStylesFromObjects(myNormal);
   */
  public applyStylesFromObjects(...styles: Style[]): {
    heading1: boolean;
    heading2: boolean;
    heading3: boolean;
    normal: boolean;
    listParagraph: boolean;
  } {
    // Convert Style objects to ApplyCustomFormattingOptions
    const options: ApplyCustomFormattingOptions = {};

    for (const style of styles) {
      const styleId = style.getStyleId();

      switch (styleId) {
        case "Heading1":
          options.heading1 = {
            run: style.getRunFormatting() as any,
            paragraph: style.getParagraphFormatting() as any,
          };
          break;

        case "Heading2":
          options.heading2 = {
            run: style.getRunFormatting() as any,
            paragraph: style.getParagraphFormatting() as any,
            tableOptions: style.getHeading2TableOptions(),
          };
          break;

        case "Heading3":
          options.heading3 = {
            run: style.getRunFormatting() as any,
            paragraph: style.getParagraphFormatting() as any,
          };
          break;

        case "Normal":
          options.normal = {
            run: style.getRunFormatting() as any,
            paragraph: style.getParagraphFormatting() as any,
          };
          break;

        case "ListParagraph":
          options.listParagraph = {
            run: style.getRunFormatting() as any,
            paragraph: style.getParagraphFormatting() as any,
          };
          break;

        default:
          // Ignore styles with other styleIds
          break;
      }
    }

    // Call existing method with converted options
    return this.applyCustomFormattingToExistingStyles(options);
  }

  /**
   * Removes extra blank paragraphs from the document while preserving marked paragraphs
   *
   * This method removes consecutive blank paragraphs, keeping only one blank line for spacing.
   * Paragraphs marked as "preserved" (via setPreserved(true)) will NOT be removed.
   *
   * A paragraph is considered blank if it:
   * - Has no text content (or only whitespace)
   * - Has no images, hyperlinks, or other non-text content
   * - Has no bookmarks or comments
   *
   * @param options Configuration options for removal
   * @returns Statistics about removed paragraphs
   *
   * @example
   * // Remove all extra blank paragraphs (except preserved ones)
   * const result = doc.removeExtraBlankParagraphs();
   * console.log(`Removed ${result.removed} blank paragraphs, preserved ${result.preserved}`);
   *
   * @example
   * // Keep one blank line between sections
   * const result = doc.removeExtraBlankParagraphs({ keepOne: true });
   *
   * @example
   * // Preserve Header 2 blank lines before removal
   * const result = doc.removeExtraBlankParagraphs({
   *   preserveHeader2BlankLines: true
   * });
   */
  public removeExtraBlankParagraphs(options?: {
    /** Whether to keep one blank paragraph between content blocks (default: true) */
    keepOne?: boolean;
    /**
     * Whether to mark blank lines after Header 2 tables as preserved before removing extra paragraphs.
     * When true, scans for 1x1 tables containing Header 2 paragraphs and marks any blank
     * paragraphs immediately after them as preserved (so they won't be removed).
     * @default false
     */
    preserveHeader2BlankLines?: boolean;
  }): {
    removed: number;
    preserved: number;
    total: number;
  } {
    const keepOne = options?.keepOne ?? true;
    const preserveHeader2BlankLines =
      options?.preserveHeader2BlankLines ?? false;
    let removed = 0;
    let preserved = 0;

    // Step 1: If requested, mark blank lines after Header 2 tables as preserved
    if (preserveHeader2BlankLines) {
      this.markHeader2BlankLinesAsPreserved();
    }

    // Track consecutive blank paragraphs
    const toRemove: Paragraph[] = [];
    let consecutiveBlanks: Paragraph[] = [];

    for (let i = 0; i < this.bodyElements.length; i++) {
      const element = this.bodyElements[i];

      // Only process paragraphs
      if (!(element instanceof Paragraph)) {
        // Not a paragraph - process any accumulated blanks
        if (consecutiveBlanks.length > 0) {
          this.processConsecutiveBlanks(consecutiveBlanks, keepOne, toRemove);
          consecutiveBlanks = [];
        }
        continue;
      }

      const para = element;

      // Check if paragraph is blank
      const isBlank = this.isParagraphBlank(para);

      if (isBlank) {
        consecutiveBlanks.push(para);
      } else {
        // Non-blank paragraph - process any accumulated blanks
        if (consecutiveBlanks.length > 0) {
          this.processConsecutiveBlanks(consecutiveBlanks, keepOne, toRemove);
          consecutiveBlanks = [];
        }
      }
    }

    // Process any remaining consecutive blanks at the end
    if (consecutiveBlanks.length > 0) {
      this.processConsecutiveBlanks(consecutiveBlanks, keepOne, toRemove);
    }

    // Count preserved paragraphs
    for (const para of toRemove) {
      if (para.isPreserved()) {
        preserved++;
      }
    }

    // Remove paragraphs that aren't preserved
    for (const para of toRemove) {
      if (!para.isPreserved()) {
        const index = this.bodyElements.indexOf(para);
        if (index !== -1) {
          this.bodyElements.splice(index, 1);
          removed++;
        }
      }
    }

    return {
      removed,
      preserved,
      total: toRemove.length,
    };
  }

  /**
   * Helper method to process consecutive blank paragraphs
   * @private
   */
  private processConsecutiveBlanks(
    blanks: Paragraph[],
    keepOne: boolean,
    toRemove: Paragraph[]
  ): void {
    if (blanks.length === 0) return;

    if (keepOne && blanks.length > 1) {
      // Keep the first one, remove the rest
      for (let i = 1; i < blanks.length; i++) {
        const blank = blanks[i];
        if (blank) {
          toRemove.push(blank);
        }
      }
    } else if (!keepOne) {
      // Remove all
      toRemove.push(...blanks);
    }
    // If keepOne is true and there's only 1 blank, don't remove it
  }

  /**
   * Marks blank lines after 1x1 Header 2 tables as preserved
   * @private
   */
  private markHeader2BlankLinesAsPreserved(): void {
    const tables = this.getAllTables();

    for (const table of tables) {
      const rowCount = table.getRowCount();
      const colCount = table.getColumnCount();

      // Check if it's a 1x1 table
      if (rowCount !== 1 || colCount !== 1) {
        continue;
      }

      // Get the cell and check if it contains a Header 2 paragraph
      const cell = table.getCell(0, 0);
      if (!cell) continue;

      const cellParas = cell.getParagraphs();
      let hasHeader2 = false;

      for (const para of cellParas) {
        const style = para.getStyle();
        if (
          style === "Heading2" ||
          style === "Heading 2" ||
          style === "CustomHeader2" ||
          style === "Header2"
        ) {
          hasHeader2 = true;
          break;
        }
      }

      if (!hasHeader2) continue;

      // Found a 1x1 table with Header 2 - mark next paragraph as preserved if it's blank
      const tableIndex = this.bodyElements.indexOf(table);
      if (tableIndex === -1) continue;

      const nextElement = this.bodyElements[tableIndex + 1];
      if (nextElement instanceof Paragraph) {
        if (this.isParagraphBlank(nextElement)) {
          nextElement.setPreserved(true);
        }
      }
    }
  }

  /**
   * Ensures that all 1x1 tables have a blank line after them with optional preserve flag.
   * This is useful for maintaining spacing after single-cell tables (e.g., Header 2 tables).
   *
   * The method:
   * 1. Finds all 1x1 tables in the document
   * 2. Checks if there's a blank paragraph immediately after each table
   * 3. If no blank paragraph exists, adds one with spacing and preserve flag
   * 4. If a blank paragraph exists, optionally marks it as preserved
   *
   * @param options Configuration options
   * @param options.spacingAfter Spacing after the blank paragraph in twips (default: 120 twips = 6pt)
   * @param options.markAsPreserved Whether to mark blank paragraphs as preserved (default: true)
   * @param options.style Style to apply to blank paragraphs (default: 'Normal')
   * @param options.filter Optional filter function to select which tables to process
   * @returns Statistics about the operation
   *
   * @example
   * // Add blank lines after all 1x1 tables with default settings
   * const result = doc.ensureBlankLinesAfter1x1Tables();
   * console.log(`Added ${result.blankLinesAdded} blank lines`);
   * console.log(`Marked ${result.existingLinesMarked} existing blank lines as preserved`);
   *
   * @example
   * // Custom spacing and preserve flag
   * doc.ensureBlankLinesAfter1x1Tables({
   *   spacingAfter: 240,  // 12pt spacing
   *   markAsPreserved: true
   * });
   *
   * @example
   * // Custom style for blank paragraphs
   * doc.ensureBlankLinesAfter1x1Tables({
   *   style: 'BodyText',  // Use BodyText instead of Normal
   *   spacingAfter: 120
   * });
   *
   * @example
   * // Only process tables with Header 2 paragraphs
   * doc.ensureBlankLinesAfter1x1Tables({
   *   filter: (table, index) => {
   *     const cell = table.getCell(0, 0);
   *     if (!cell) return false;
   *     return cell.getParagraphs().some(p => {
   *       const style = p.getStyle();
   *       return style === 'Heading2' || style === 'Heading 2';
   *     });
   *   }
   * });
   */
  public ensureBlankLinesAfter1x1Tables(options?: {
    spacingAfter?: number;
    markAsPreserved?: boolean;
    style?: string;
    filter?: (table: Table, index: number) => boolean;
  }): {
    tablesProcessed: number;
    blankLinesAdded: number;
    existingLinesMarked: number;
  } {
    const spacingAfter = options?.spacingAfter ?? 120;
    const markAsPreserved = options?.markAsPreserved ?? true;
    const style = options?.style ?? "Normal";
    const filter = options?.filter;

    let tablesProcessed = 0;
    let blankLinesAdded = 0;
    let existingLinesMarked = 0;

    const tables = this.getAllTables();

    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      if (!table) continue;

      const rowCount = table.getRowCount();
      const colCount = table.getColumnCount();

      // Check if it's a 1x1 table
      if (rowCount !== 1 || colCount !== 1) {
        continue;
      }

      // Apply filter if provided
      if (filter && !filter(table, i)) {
        continue;
      }

      tablesProcessed++;

      // Find table index in body elements
      const tableIndex = this.bodyElements.indexOf(table);
      if (tableIndex === -1) continue;

      // Check next element
      const nextElement = this.bodyElements[tableIndex + 1];

      if (nextElement instanceof Paragraph) {
        // Next element is a paragraph - check if it's blank
        if (this.isParagraphBlank(nextElement)) {
          // Blank paragraph exists - mark as preserved if requested
          if (markAsPreserved && !nextElement.isPreserved()) {
            nextElement.setPreserved(true);
            existingLinesMarked++;
          }
        } else {
          // Next paragraph has content - add blank paragraph between table and content
          const blankPara = Paragraph.create();
          blankPara.setStyle(style);
          blankPara.setSpaceAfter(spacingAfter);
          if (markAsPreserved) {
            blankPara.setPreserved(true);
          }
          this.bodyElements.splice(tableIndex + 1, 0, blankPara);
          blankLinesAdded++;
        }
      } else {
        // No paragraph after table (or it's another table/element) - add blank paragraph
        const blankPara = Paragraph.create();
        blankPara.setStyle(style);
        blankPara.setSpaceAfter(spacingAfter);
        if (markAsPreserved) {
          blankPara.setPreserved(true);
        }
        this.bodyElements.splice(tableIndex + 1, 0, blankPara);
        blankLinesAdded++;
      }
    }

    return {
      tablesProcessed,
      blankLinesAdded,
      existingLinesMarked,
    };
  }
  /**
   * Ensures that all tables (excluding 1x1 tables) have a blank line after them with optional preserve flag.
   * This is useful for maintaining spacing after multi-cell tables.
   *
   * The method:
   * 1. Finds all tables with more than one cell (not 1x1) in the document
   * 2. Checks if there's a blank paragraph immediately after each table
   * 3. If no blank paragraph exists, adds one with spacing and preserve flag
   * 4. If a blank paragraph exists, optionally marks it as preserved
   *
   * @param options Configuration options
   * @param options.spacingAfter Spacing after the blank paragraph in twips (default: 120 twips = 6pt)
   * @param options.markAsPreserved Whether to mark blank paragraphs as preserved (default: true)
   * @param options.style Style to apply to blank paragraphs (default: 'Normal')
   * @param options.filter Optional filter function to select which tables to process
   * @returns Statistics about the operation
   *
   * @example
   * // Add blank lines after all multi-cell tables with default settings
   * const result = doc.ensureBlankLinesAfterOtherTables();
   * console.log(`Added ${result.blankLinesAdded} blank lines`);
   * console.log(`Marked ${result.existingLinesMarked} existing blank lines as preserved`);
   *
   * @example
   * // Custom spacing and preserve flag
   * doc.ensureBlankLinesAfterOtherTables({
   *   spacingAfter: 240,  // 12pt spacing
   *   markAsPreserved: true
   * });
   *
   * @example
   * // Custom style for blank paragraphs
   * doc.ensureBlankLinesAfterOtherTables({
   *   style: 'BodyText',  // Use BodyText instead of Normal
   *   spacingAfter: 120
   * });
   *
   * @example
   * // Only process tables with more than 2 rows
   * doc.ensureBlankLinesAfterOtherTables({
   *   filter: (table, index) => table.getRowCount() > 2
   * });
   */
  public ensureBlankLinesAfterOtherTables(options?: {
    spacingAfter?: number;
    markAsPreserved?: boolean;
    style?: string;
    filter?: (table: Table, index: number) => boolean;
  }): {
    tablesProcessed: number;
    blankLinesAdded: number;
    existingLinesMarked: number;
  } {
    const spacingAfter = options?.spacingAfter ?? 120;
    const markAsPreserved = options?.markAsPreserved ?? true;
    const style = options?.style ?? "Normal";
    const filter = options?.filter;

    let tablesProcessed = 0;
    let blankLinesAdded = 0;
    let existingLinesMarked = 0;

    const tables = this.getAllTables();

    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      if (!table) continue;

      const rowCount = table.getRowCount();
      const colCount = table.getColumnCount();

      // Skip 1x1 tables (handled by ensureBlankLinesAfter1x1Tables)
      if (rowCount === 1 && colCount === 1) {
        continue;
      }

      // Apply filter if provided
      if (filter && !filter(table, i)) {
        continue;
      }

      tablesProcessed++;

      // Find table index in body elements
      const tableIndex = this.bodyElements.indexOf(table);
      if (tableIndex === -1) continue;

      // Check next element
      const nextElement = this.bodyElements[tableIndex + 1];

      if (nextElement instanceof Paragraph) {
        // Next element is a paragraph - check if it's blank
        if (this.isParagraphBlank(nextElement)) {
          // Blank paragraph exists - mark as preserved if requested
          if (markAsPreserved && !nextElement.isPreserved()) {
            nextElement.setPreserved(true);
            existingLinesMarked++;
          }
        } else {
          // Next paragraph has content - add blank paragraph between table and content
          const blankPara = Paragraph.create();
          blankPara.setStyle(style);
          blankPara.setSpaceAfter(spacingAfter);
          if (markAsPreserved) {
            blankPara.setPreserved(true);
          }
          this.bodyElements.splice(tableIndex + 1, 0, blankPara);
          blankLinesAdded++;
        }
      } else {
        // No paragraph after table (or it's another table/element) - add blank paragraph
        const blankPara = Paragraph.create();
        blankPara.setStyle(style);
        blankPara.setSpaceAfter(spacingAfter);
        if (markAsPreserved) {
          blankPara.setPreserved(true);
        }
        this.bodyElements.splice(tableIndex + 1, 0, blankPara);
        blankLinesAdded++;
      }
    }

    return {
      tablesProcessed,
      blankLinesAdded,
      existingLinesMarked,
    };
  }

  /**
   * Standardizes all bullet list symbols to be bold, 12pt, and black (#000000)
   *
   * This helper ensures consistent bullet formatting across all bullet lists in the document.
   * It modifies the numbering definitions (not individual paragraphs), preserving the actual
   * bullet symbols (•, ○, ▪, etc.) while standardizing their visual formatting.
   *
   * **Important**: This only affects the bullet symbol formatting, not the text content after it.
   * The actual bullet characters are preserved as they were originally defined.
   *
   * @param options Formatting options
   * @returns Statistics about lists updated
   *
   * @example
   * // Standardize all bullet symbols with defaults (Arial 12pt bold black)
   * const result = doc.standardizeBulletSymbols();
   * console.log(`Updated ${result.listsUpdated} bullet lists (${result.levelsModified} levels)`);
   *
   * @example
   * // Custom formatting for bullet symbols
   * const result = doc.standardizeBulletSymbols({
   *   bold: true,
   *   fontSize: 28,  // 14pt
   *   color: 'FF0000',  // Red
   *   font: 'Calibri'
   * });
   */
  public standardizeBulletSymbols(options?: {
    bold?: boolean;
    fontSize?: number;
    color?: string;
    font?: string;
  }): {
    listsUpdated: number;
    levelsModified: number;
  } {
    const {
      bold = true,
      fontSize = 24, // 12pt
      color = "000000",
      font = "Arial",
    } = options || {};

    let listsUpdated = 0;
    let levelsModified = 0;

    const instances = this.numberingManager.getAllInstances();

    for (const instance of instances) {
      const abstractNumId = instance.getAbstractNumId();
      const abstractNum =
        this.numberingManager.getAbstractNumbering(abstractNumId);

      if (!abstractNum) continue;

      // Only process bullet lists (skip numbered lists)
      const level0 = abstractNum.getLevel(0);
      if (!level0 || level0.getFormat() !== "bullet") continue;

      // Update all 9 levels (0-8)
      for (let levelIndex = 0; levelIndex < 9; levelIndex++) {
        const numLevel = abstractNum.getLevel(levelIndex);
        if (!numLevel) continue;

        numLevel.setFont(font);
        numLevel.setFontSize(fontSize);
        numLevel.setBold(bold);
        numLevel.setColor(color);

        levelsModified++;
      }

      listsUpdated++;
    }

    return { listsUpdated, levelsModified };
  }

  /**
   * Standardizes numbered list prefixes (1., a., i., etc.) to Verdana 12pt bold black
   *
   * This only affects the prefix/number formatting, not the text content after it.
   * It modifies the numbering definitions in the document while preserving the
   * numbering format type (decimal, roman, letter, etc.).
   *
   * **Important**: This updates the visual formatting of prefixes like "1.", "a.", "i."
   * but does not change the numbering type itself.
   *
   * @param options Formatting options
   * @returns Statistics about lists updated
   *
   * @example
   * // Standardize all numbered list prefixes with defaults (Verdana 12pt bold black)
   * const result = doc.standardizeNumberedListPrefixes();
   * console.log(`Updated ${result.listsUpdated} numbered lists (${result.levelsModified} levels)`);
   *
   * @example
   * // Custom formatting for numbered list prefixes
   * const result = doc.standardizeNumberedListPrefixes({
   *   bold: true,
   *   fontSize: 24,
   *   color: '000000',
   *   font: 'Verdana'
   * });
   */
  public standardizeNumberedListPrefixes(options?: {
    bold?: boolean;
    fontSize?: number;
    color?: string;
    font?: string;
  }): {
    listsUpdated: number;
    levelsModified: number;
  } {
    const {
      bold = true,
      fontSize = 24, // 12pt
      color = "000000",
      font = "Verdana",
    } = options || {};

    let listsUpdated = 0;
    let levelsModified = 0;

    const instances = this.numberingManager.getAllInstances();

    for (const instance of instances) {
      const abstractNumId = instance.getAbstractNumId();
      const abstractNum =
        this.numberingManager.getAbstractNumbering(abstractNumId);

      if (!abstractNum) continue;

      // Only process numbered lists (skip bullet lists)
      const level0 = abstractNum.getLevel(0);
      if (!level0 || level0.getFormat() === "bullet") continue;

      // Update all 9 levels (0-8)
      for (let levelIndex = 0; levelIndex < 9; levelIndex++) {
        const numLevel = abstractNum.getLevel(levelIndex);
        if (!numLevel) continue;

        numLevel.setFont(font);
        numLevel.setFontSize(fontSize);
        numLevel.setBold(bold);
        numLevel.setColor(color);

        levelsModified++;
      }

      listsUpdated++;
    }

    return { listsUpdated, levelsModified };
  }

  /**
   * Standardizes all hyperlinks in the document to Verdana 12pt blue (#0000FF) underline
   *
   * This applies consistent formatting to all hyperlinks throughout the document,
   * including those in tables. The method preserves the hyperlink URLs and text
   * while updating only the visual formatting.
   *
   * @param options Formatting options
   * @returns Number of hyperlinks updated
   *
   * @example
   * // Use default formatting (Verdana 12pt blue underline)
   * const count = doc.standardizeAllHyperlinks();
   * console.log(`Standardized ${count} hyperlinks`);
   *
   * @example
   * // Custom hyperlink formatting
   * const count = doc.standardizeAllHyperlinks({
   *   font: 'Arial',
   *   size: 11,
   *   color: 'FF0000',  // Red
   *   underline: true
   * });
   */
  public standardizeAllHyperlinks(options?: {
    font?: string;
    size?: number;
    color?: string;
    underline?: boolean;
  }): number {
    const {
      font = "Verdana",
      size = 12,
      color = "0000FF",
      underline = true,
    } = options || {};

    const hyperlinks = this.getHyperlinks();

    for (const { hyperlink } of hyperlinks) {
      hyperlink.setFormatting({
        font: font,
        size: size,
        color: color,
        underline: underline ? "single" : false,
      });
    }

    return hyperlinks.length;
  }

  /**
   * Cleans direct formatting from paragraphs that have a style applied.
   * Preserves the style reference while removing formatting overrides.
   *
   * @param options Optional list of style names to clean (default: all styles)
   * @returns Number of paragraphs cleaned
   * @example
   * ```typescript
   * // Clean all styled paragraphs
   * const count = doc.cleanFormatting();
   *
   * // Clean specific styles only
   * const count = doc.cleanFormatting(['Heading1', 'Heading2', 'Normal']);
   * ```
   */
  public cleanFormatting(styleNames?: string[]): number {
    let cleaned = 0;

    for (const para of this.bodyElements) {
      if (!(para instanceof Paragraph)) continue;
      if (para.isPreserved()) continue;

      const currentStyle = para.getStyle();
      if (!currentStyle) continue;

      // If style filter provided, only clean matching styles
      if (styleNames && !styleNames.includes(currentStyle)) continue;

      para.clearDirectFormatting();
      cleaned++;
    }

    return cleaned;
  }

  /**
   * Applies Heading 1 style to paragraphs with H1-like style names
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   * @example
   * ```typescript
   * // Simple usage
   * doc.applyH1();
   *
   * // With custom formatting
   * doc.applyH1({
   *   format: { font: 'Arial', size: 18, emphasis: ['bold'] }
   * });
   *
   * // Preserve specific properties
   * doc.applyH1({
   *   keepProperties: ['bold', 'color'],
   *   format: { font: 'Verdana' }
   * });
   * ```
   */
  public applyH1(options?: StyleApplyOptions): number {
    return this.applyStyleToMatching("Heading1", options, (style) =>
      /^(heading\s*1|header\s*1|h1)$/i.test(style)
    );
  }

  /**
   * Applies Heading 2 style to paragraphs with H2-like style names
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   * @example
   * ```typescript
   * doc.applyH2({
   *   format: { font: 'Verdana', size: 14, color: '000000' }
   * });
   * ```
   */
  public applyH2(options?: StyleApplyOptions): number {
    return this.applyStyleToMatching("Heading2", options, (style) =>
      /^(heading\s*2|header\s*2|h2)$/i.test(style)
    );
  }

  /**
   * Applies Heading 3 style to paragraphs with H3-like style names
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   * @example
   * ```typescript
   * doc.applyH3({
   *   format: { font: 'Verdana', size: 12, emphasis: ['bold'] }
   * });
   * ```
   */
  public applyH3(options?: StyleApplyOptions): number {
    return this.applyStyleToMatching("Heading3", options, (style) =>
      /^(heading\s*3|header\s*3|h3)$/i.test(style)
    );
  }

  /**
   * Applies Normal style to paragraphs without recognized styles
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   */
  public applyNormal(options?: StyleApplyOptions): number {
    const targets =
      options?.paragraphs ||
      this.getAllParagraphs().filter((p) => {
        const style = p.getStyle();
        return (
          !style ||
          !/^(heading|header|h\d|list|toc|tod|caution|table)/i.test(style)
        );
      });

    let count = 0;
    for (const para of targets) {
      if (para.isPreserved()) continue;
      para.setStyle("Normal");

      if (options?.keepProperties && options.keepProperties.length > 0) {
        this.clearFormattingExcept(para, options.keepProperties);
      } else {
        para.clearDirectFormatting();
      }

      if (options?.format) {
        this.applyFormatOptions(para, options.format);
      }

      count++;
    }
    return count;
  }

  /**
   * Applies list style to numbered lists
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   */
  public applyNumList(options?: StyleApplyOptions): number {
    return this.applyStyleToMatching("ListParagraph", options, (style) =>
      /^(list\s*number|numbered\s*list|list\s*paragraph)$/i.test(style)
    );
  }

  /**
   * Applies list style to bullet lists
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   */
  public applyBulletList(options?: StyleApplyOptions): number {
    return this.applyStyleToMatching("ListParagraph", options, (style) =>
      /^(list\s*bullet|bullet\s*list|list\s*paragraph)$/i.test(style)
    );
  }

  /**
   * Applies Table of Contents style
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   */
  public applyTOC(options?: StyleApplyOptions): number {
    return this.applyStyleToMatching("TOC", options, (style) =>
      /^(toc|table\s*of\s*contents|toc\s*heading)$/i.test(style)
    );
  }

  /**
   * Applies Top of Document style
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   */
  public applyTOD(options?: StyleApplyOptions): number {
    return this.applyStyleToMatching("TopOfDocument", options, (style) =>
      /^(tod|top\s*of\s*document|document\s*top)$/i.test(style)
    );
  }

  /**
   * Applies Caution style
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   */
  public applyCaution(options?: StyleApplyOptions): number {
    return this.applyStyleToMatching("Caution", options, (style) =>
      /^(caution|warning|important|alert)$/i.test(style)
    );
  }

  /**
   * Applies header style to table cell paragraphs (typically first row)
   * @param options Optional style application options
   * @returns Number of paragraphs updated
   */
  public applyCellHeader(options?: StyleApplyOptions): number {
    let count = 0;
    const tables = this.getAllTables();

    for (const table of tables) {
      const firstRow = table.getRow(0);
      if (!firstRow) continue;

      for (const cell of firstRow.getCells()) {
        for (const para of cell.getParagraphs()) {
          if (para.isPreserved()) continue;
          para.setStyle("TableHeader");

          if (options?.keepProperties && options.keepProperties.length > 0) {
            this.clearFormattingExcept(para, options.keepProperties);
          } else {
            para.clearDirectFormatting();
          }

          if (options?.format) {
            this.applyFormatOptions(para, options.format);
          }

          count++;
        }
      }
    }

    return count;
  }

  /**
   * Applies hyperlink style to hyperlinks
   * @returns Number of hyperlinks updated
   */
  public applyHyperlink(): number {
    let count = 0;
    const hyperlinks = this.getHyperlinks();

    for (const { hyperlink } of hyperlinks) {
      hyperlink.resetToStandardFormatting();
      count++;
    }

    return count;
  }

  /**
   * Helper method to apply formatting options to a paragraph
   * @private
   */
  private applyFormatOptions(para: Paragraph, options: FormatOptions): void {
    // Text formatting
    if (options.font || options.size || options.color || options.emphasis) {
      for (const run of para.getRuns()) {
        if (options.font) run.setFont(options.font);
        if (options.size) run.setSize(options.size);
        if (options.color) run.setColor(options.color);
        if (options.emphasis) {
          options.emphasis.forEach((emp) => {
            if (emp === "bold") run.setBold(true);
            if (emp === "italic") run.setItalic(true);
            if (emp === "underline") run.setUnderline("single");
          });
        }
      }
    }

    // Alignment
    if (options.alignment) {
      para.setAlignment(options.alignment);
    }

    // Spacing (convert points to twips: 1pt = 20 twips)
    if (options.spaceAbove !== undefined) {
      para.setSpaceBefore(options.spaceAbove * 20);
    }
    if (options.spaceBelow !== undefined) {
      para.setSpaceAfter(options.spaceBelow * 20);
    }
    if (options.lineSpacing !== undefined) {
      para.setLineSpacing(options.lineSpacing * 20);
    }

    // Indentation (convert inches to twips: 1in = 1440 twips)
    if (options.indentLeft !== undefined) {
      para.setLeftIndent(options.indentLeft * 1440);
    }
    if (options.indentRight !== undefined) {
      para.setRightIndent(options.indentRight * 1440);
    }
    if (options.indentFirst !== undefined) {
      para.setFirstLineIndent(options.indentFirst * 1440);
    }
    if (options.indentHanging !== undefined) {
      // Set hanging indent directly through formatting
      if (!para.formatting.indentation) {
        para.formatting.indentation = {};
      }
      para.formatting.indentation.hanging = options.indentHanging * 1440;
    }

    // Advanced options (only set if true)
    if (options.keepWithNext) {
      para.setKeepNext(true);
    }
    if (options.keepLines) {
      para.setKeepLines(true);
    }
  }

  /**
   * Helper method to selectively clear formatting while preserving specific properties
   * @private
   */
  private clearFormattingExcept(
    para: Paragraph,
    keepProperties: string[]
  ): void {
    // Save properties to keep
    const savedProps: any = {};
    const formatting = para.formatting;

    for (const prop of keepProperties) {
      if ((formatting as any)[prop] !== undefined) {
        savedProps[prop] = (formatting as any)[prop];
      }
    }

    // Clear all formatting
    para.clearDirectFormatting();

    // Restore saved properties
    for (const prop of keepProperties) {
      if (savedProps[prop] !== undefined) {
        (para.formatting as any)[prop] = savedProps[prop];
      }
    }

    // Handle run-level properties
    for (const run of para.getRuns()) {
      const runFormatting = run.getFormatting();
      const runSavedProps: any = {};
      for (const prop of keepProperties) {
        if ((runFormatting as any)[prop] !== undefined) {
          runSavedProps[prop] = (runFormatting as any)[prop];
        }
      }

      run.clearFormatting();

      // Restore saved properties using appropriate setters
      if (runSavedProps.bold !== undefined) run.setBold(runSavedProps.bold);
      if (runSavedProps.italic !== undefined)
        run.setItalic(runSavedProps.italic);
      if (runSavedProps.underline !== undefined)
        run.setUnderline(runSavedProps.underline);
      if (runSavedProps.color !== undefined) run.setColor(runSavedProps.color);
      if (runSavedProps.font !== undefined) run.setFont(runSavedProps.font);
      if (runSavedProps.size !== undefined) run.setSize(runSavedProps.size);
      if (runSavedProps.highlight !== undefined)
        run.setHighlight(runSavedProps.highlight);
      if (runSavedProps.strike !== undefined)
        run.setStrike(runSavedProps.strike);
      if (runSavedProps.subscript !== undefined)
        run.setSubscript(runSavedProps.subscript);
      if (runSavedProps.superscript !== undefined)
        run.setSuperscript(runSavedProps.superscript);
    }
  }

  /**
   * Helper method to apply style to matching paragraphs
   * @private
   */
  private applyStyleToMatching(
    targetStyle: string,
    options: StyleApplyOptions | undefined,
    matcher: (style: string) => boolean
  ): number {
    const targets =
      options?.paragraphs ||
      this.getAllParagraphs().filter((p) => {
        const style = p.getStyle();
        return style && matcher(style);
      });

    let count = 0;
    for (const para of targets) {
      if (para.isPreserved()) continue;

      // Apply style
      para.setStyle(targetStyle);

      // Handle formatting
      if (options?.keepProperties && options.keepProperties.length > 0) {
        // Clear formatting except specified properties
        this.clearFormattingExcept(para, options.keepProperties);
      } else {
        // Clear all formatting
        para.clearDirectFormatting();
      }

      // Apply custom formatting if provided
      if (options?.format) {
        this.applyFormatOptions(para, options.format);
      }

      count++;
    }
    return count;
  }

  /**
   * Checks if a paragraph is blank (no meaningful content)
   * @private
   */
  private isParagraphBlank(para: Paragraph): boolean {
    const content = para.getContent();

    // No content at all
    if (!content || content.length === 0) {
      return true;
    }

    // Check all content items
    for (const item of content) {
      // Hyperlinks count as content
      if ((item as any).constructor.name === "Hyperlink") {
        return false;
      }

      // Images/shapes count as content
      if ((item as any).constructor.name === "Shape") {
        return false;
      }

      // TextBox count as content
      if ((item as any).constructor.name === "TextBox") {
        return false;
      }

      // Fields count as content
      if ((item as any).constructor.name === "Field") {
        return false;
      }

      // Check runs for non-whitespace text
      if ((item as any).getText) {
        const text = (item as any).getText().trim();
        if (text !== "") {
          return false;
        }
      }
    }

    // Check for bookmarks
    if (
      para.getBookmarksStart().length > 0 ||
      para.getBookmarksEnd().length > 0
    ) {
      return false;
    }

    return true;
  }

  /**
   * Parses a TOC field instruction to extract which heading levels to include
   *
   * Handles field codes like:
   * - "TOC \o "1-3"" → [1, 2, 3]
   * - "TOC \t "Heading 2,2,"" → [2]
   * - "TOC \o "1-2" \t "Heading 3,3,"" → [1, 2, 3]
   *
   * @param instrText The TOC field instruction text
   * @returns Array of heading levels (1-9) to include
   */
  private parseTOCFieldInstruction(instrText: string): number[] {
    const levels = new Set<number>();

    // Parse \o "X-Y" switch (outline levels)
    const outlineMatch = instrText.match(/\\o\s+"(\d+)-(\d+)"/);
    if (outlineMatch && outlineMatch[1] && outlineMatch[2]) {
      const start = parseInt(outlineMatch[1], 10);
      const end = parseInt(outlineMatch[2], 10);
      for (let i = start; i <= end; i++) {
        if (i >= 1 && i <= 9) {
          levels.add(i);
        }
      }
    }

    // Parse \t "StyleName,Level," switches (custom styles)
    // Microsoft Word format: \t "Heading 2,2," (everything inside quotes)
    const styleMatches = instrText.matchAll(/\\t\s+"([^"]+)"/g);
    for (const match of styleMatches) {
      const content = match[1]; // e.g., "Heading 2,2,"
      if (!content) continue;

      // Split by comma: ["Heading 2", "2", ""]
      const parts = content
        .split(",")
        .map((p) => p.trim())
        .filter((p) => p);
      if (parts.length < 2) continue;

      const styleName = parts[0]; // "Heading 2"
      const levelStr = parts[1]; // "2"

      if (!styleName || !levelStr) continue;

      const level = parseInt(levelStr, 10);
      if (isNaN(level)) continue;

      // Extract level from Heading style names (e.g., "Heading 2" → 2)
      const headingMatch = styleName.match(/Heading\s*(\d+)/i);
      if (headingMatch && headingMatch[1]) {
        const headingLevel = parseInt(headingMatch[1], 10);
        if (headingLevel >= 1 && headingLevel <= 9) {
          levels.add(headingLevel);
        }
      } else if (level >= 1 && level <= 9) {
        // For custom styles, use the level number from the \t switch
        levels.add(level);
      }
    }

    // If no levels found, default to 1-3
    if (levels.size === 0) {
      return [1, 2, 3];
    }

    return Array.from(levels).sort((a, b) => a - b);
  }

  /**
   * Finds all headings in the document that match the specified levels
   *
   * @param levels Array of heading levels to include (e.g., [1, 2, 3])
   * @returns Array of heading information objects
   */
  /**
   * Find headings for TOC by parsing XML directly (searches body AND tables)
   * This is more reliable than using bodyElements as it searches inside table cells too
   */
  private findHeadingsForTOCFromXML(
    docXml: string,
    levels: number[]
  ): Array<{ level: number; text: string; bookmark: string }> {
    const headings: Array<{ level: number; text: string; bookmark: string }> =
      [];
    const levelSet = new Set(levels);

    try {
      // Parse document.xml to object structure
      const parsed = XMLParser.parseToObject(docXml, { trimValues: false });
      const document = parsed["w:document"];
      if (!document) {
        return headings;
      }

      const body = (document as any)["w:body"];
      if (!body) {
        return headings;
      }

      // Helper function to extract heading info from a parsed paragraph object
      const extractHeading = (para: any): void => {
        const pPr = para["w:pPr"];
        if (!pPr || !pPr["w:pStyle"]) {
          return;
        }

        const styleVal = pPr["w:pStyle"]["@_w:val"];
        if (!styleVal) {
          return;
        }

        // Check if style matches "HeadingN" format (exact match, case-insensitive)
        const headingMatch = styleVal.match(/^Heading(\d+)$/i);
        if (!headingMatch || !headingMatch[1]) {
          return;
        }

        const headingLevel = parseInt(headingMatch[1], 10);

        // Check if this level should be included in TOC
        if (!levelSet.has(headingLevel)) {
          return;
        }

        // Extract bookmark (look for bookmarks with "_heading" in name)
        let bookmark = "";
        const bookmarkStart = para["w:bookmarkStart"];
        if (bookmarkStart) {
          const bookmarkArray = Array.isArray(bookmarkStart)
            ? bookmarkStart
            : [bookmarkStart];
          for (const bm of bookmarkArray) {
            const bmName = bm["@_w:name"];
            if (bmName && bmName.toLowerCase().includes("_heading")) {
              bookmark = bmName;
              break;
            }
          }
        }

        // Extract text from runs
        let text = "";
        const runs = para["w:r"];
        if (runs) {
          const runArray = Array.isArray(runs) ? runs : [runs];
          for (const run of runArray) {
            const textElement = run["w:t"];
            if (textElement) {
              if (typeof textElement === "string") {
                text += textElement;
              } else if (textElement["#text"]) {
                text += textElement["#text"];
              }
            }
          }
        }

        // Only add if we have text
        text = text.trim();
        if (!text) {
          return;
        }

        // Generate bookmark if not found
        if (!bookmark) {
          bookmark = `_Toc${Date.now()}_${headings.length}`;
        }

        headings.push({
          level: headingLevel,
          text: text,
          bookmark: bookmark,
        });
      };

      // Search in direct paragraphs
      const paragraphs = body["w:p"];
      if (paragraphs) {
        const paraArray = Array.isArray(paragraphs) ? paragraphs : [paragraphs];
        for (const para of paraArray) {
          extractHeading(para);
        }
      }

      // Search in tables (this is critical - many documents have headings in tables)
      const tables = body["w:tbl"];
      if (tables) {
        const tableArray = Array.isArray(tables) ? tables : [tables];
        for (const table of tableArray) {
          const rows = table["w:tr"];
          if (!rows) continue;

          const rowArray = Array.isArray(rows) ? rows : [rows];
          for (const row of rowArray) {
            const cells = row["w:tc"];
            if (!cells) continue;

            const cellArray = Array.isArray(cells) ? cells : [cells];
            for (const cell of cellArray) {
              const cellParas = cell["w:p"];
              if (!cellParas) continue;

              const cellParaArray = Array.isArray(cellParas)
                ? cellParas
                : [cellParas];
              for (const para of cellParaArray) {
                extractHeading(para);
              }
            }
          }
        }
      }
    } catch (error) {
      defaultLogger.error(
        "Error parsing document.xml for headings:",
        error instanceof Error
          ? { message: error.message, stack: error.stack }
          : { error: String(error) }
      );
    }

    return headings;
  }

  /**
   * Legacy method - searches only bodyElements (doesn't search inside tables)
   * Kept for compatibility but not recommended
   * @deprecated Use findHeadingsForTOCFromXML instead
   */
  private findHeadingsForTOC(
    levels: number[]
  ): Array<{ level: number; text: string; bookmark: string }> {
    const headings: Array<{ level: number; text: string; bookmark: string }> =
      [];
    const levelSet = new Set(levels);

    // Iterate through body elements
    for (const element of this.bodyElements) {
      if (element instanceof Paragraph) {
        const para = element;
        const formatting = para.getFormatting();

        // Check if paragraph has a heading style (handle both "Heading1" and "Heading 1")
        if (formatting.style) {
          const styleMatch = formatting.style.match(/Heading\s*(\d+)/i);
          if (styleMatch && styleMatch[1]) {
            const headingLevel = parseInt(styleMatch[1], 10);

            // Check if this level should be included in TOC
            if (levelSet.has(headingLevel)) {
              const text = para.getText().trim();

              if (text) {
                // Create or get bookmark for this heading
                const bookmark =
                  this.bookmarkManager.createHeadingBookmark(text);

                headings.push({
                  level: headingLevel,
                  text: text,
                  bookmark: bookmark.getName(),
                });
              }
            }
          }
        }
      }
    }

    return headings;
  }

  /**
   * Generates TOC XML structure with populated entries
   *
   * Creates a complete SDT-wrapped TOC with:
   * - Complex field structure (begin/instruction/separate/entries/end)
   * - Pre-populated hyperlink entries for each heading
   * - Proper formatting (Verdana, blue, underlined)
   *
   * @param headings Array of heading information
   * @param originalInstrText Original TOC field instruction to preserve switches
   * @returns Complete TOC XML string
   */
  private generateTOCXML(
    headings: Array<{ level: number; text: string; bookmark: string }>,
    originalInstrText: string
  ): string {
    const sdtId = Math.floor(Math.random() * 2000000000) - 1000000000;

    let tocXml = "<w:sdt>";

    // SDT properties
    tocXml += "<w:sdtPr>";
    tocXml += `<w:id w:val="${sdtId}"/>`;
    tocXml += "<w:docPartObj>";
    tocXml += '<w:docPartGallery w:val="Table of Contents"/>';
    tocXml += '<w:docPartUnique w:val="1"/>';
    tocXml += "</w:docPartObj>";
    tocXml += "</w:sdtPr>";

    // SDT content
    tocXml += "<w:sdtContent>";

    // Calculate minimum level for relative indentation
    // If TOC shows only Header 2s, minLevel=2, so Header 2 gets 0" indent
    const minLevel =
      headings.length > 0 ? Math.min(...headings.map((h) => h.level)) : 1;

    // First paragraph: field begin + instruction + separator + first entry (if any)
    tocXml += "<w:p>";
    tocXml += "<w:pPr>";
    tocXml +=
      '<w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/>';

    // Add indentation for first entry relative to minimum level (0.25" per level)
    if (headings.length > 0 && headings[0]) {
      const firstIndent = (headings[0].level - minLevel) * 360; // 360 twips = 0.25 inches
      if (firstIndent > 0) {
        tocXml += `<w:ind w:left="${firstIndent}"/>`;
      }
    }

    tocXml += "</w:pPr>";

    // Field begin
    tocXml += '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';

    // Field instruction (preserve original switches)
    tocXml += "<w:r>";
    tocXml += `<w:instrText xml:space="preserve">${this.escapeXml(
      originalInstrText
    )}</w:instrText>`;
    tocXml += "</w:r>";

    // Field separator
    tocXml += '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';

    // First entry (if any)
    if (headings.length > 0 && headings[0]) {
      tocXml += this.buildTOCEntryXML(headings[0]);
    }

    tocXml += "</w:p>";

    // Remaining entries (each in its own paragraph)
    for (let i = 1; i < headings.length; i++) {
      const heading = headings[i];
      if (!heading) continue;

      tocXml += "<w:p>";
      tocXml += "<w:pPr>";
      tocXml +=
        '<w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/>';

      // Add indentation relative to minimum level (0.25" per level above minimum)
      const indent = (heading.level - minLevel) * 360;
      if (indent > 0) {
        tocXml += `<w:ind w:left="${indent}"/>`;
      }

      tocXml += "</w:pPr>";
      tocXml += this.buildTOCEntryXML(heading);
      tocXml += "</w:p>";
    }

    // Final paragraph with field end
    tocXml += "<w:p>";
    tocXml += "<w:pPr>";
    tocXml +=
      '<w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/>';
    tocXml += "</w:pPr>";
    tocXml += '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
    tocXml += "</w:p>";

    tocXml += "</w:sdtContent>";
    tocXml += "</w:sdt>";

    return tocXml;
  }

  /**
   * Builds XML for a single TOC entry with hyperlink
   *
   * @param heading Heading information
   * @returns XML string for the TOC entry
   */
  private buildTOCEntryXML(heading: {
    level: number;
    text: string;
    bookmark: string;
  }): string {
    const escapedText = this.escapeXml(heading.text);

    let xml = "";
    xml += `<w:hyperlink w:anchor="${this.escapeXml(heading.bookmark)}">`;
    xml += "<w:r>";
    xml += "<w:rPr>";
    xml +=
      '<w:rFonts w:ascii="Verdana" w:hAnsi="Verdana" w:cs="Verdana" w:eastAsia="Verdana"/>';
    xml += '<w:color w:val="0000FF"/>';
    xml += '<w:sz w:val="24"/>';
    xml += '<w:szCs w:val="24"/>';
    xml += '<w:u w:val="single"/>';
    xml += "</w:rPr>";
    xml += `<w:t xml:space="preserve">${escapedText}</w:t>`;
    xml += "</w:r>";
    xml += "</w:hyperlink>";

    return xml;
  }

  /**
   * Escapes XML special characters
   *
   * @param text Text to escape
   * @returns Escaped text
   */
  private escapeXml(text: string): string {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  /**
   * Populates TOCs in a saved file
   * Private helper used by save() when auto-populate is enabled
   *
   * @param filePath Path to the saved DOCX file
   * @returns Promise resolving to the number of TOCs populated
   * @private
   */
  private async populateTOCsInFile(filePath: string): Promise<number> {
    // Load the saved document
    const handler = new ZipHandler();
    await handler.load(filePath);

    // Get document.xml
    const docXml = handler.getFileAsString("word/document.xml");
    if (!docXml) {
      return 0;
    }

    // Populate all TOCs in the XML
    const modifiedXml = this.populateAllTOCsInXML(docXml);

    // Update and save if changes were made
    if (modifiedXml !== docXml) {
      handler.updateFile("word/document.xml", modifiedXml);
      await handler.save(filePath);

      // Count TOCs that were populated
      const tocRegex =
        /<w:sdt>[\s\S]*?<w:docPartGallery w:val="Table of Contents"[\s\S]*?<\/w:sdt>/g;
      const matches = Array.from(docXml.matchAll(tocRegex));
      return matches.length;
    }

    return 0;
  }

  /**
   * Populates all TOCs in document XML
   * Extracted from replaceTableOfContents for reuse
   *
   * @param docXml The document XML string
   * @returns Modified XML with populated TOCs
   * @private
   */
  private populateAllTOCsInXML(docXml: string): string {
    const tocRegex =
      /<w:sdt>[\s\S]*?<w:docPartGallery w:val="Table of Contents"[\s\S]*?<\/w:sdt>/g;
    const tocMatches = Array.from(docXml.matchAll(tocRegex));

    if (tocMatches.length === 0) return docXml;

    let modifiedXml = docXml;

    for (const match of tocMatches) {
      try {
        const tocXml = match[0];

        // Extract field instruction
        const instrMatch = tocXml.match(
          /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/
        );
        if (!instrMatch?.[1]) continue;

        // Decode XML entities
        let fieldInstruction = instrMatch[1]
          .replace(/&amp;/g, "&")
          .replace(/&lt;/g, "<")
          .replace(/&gt;/g, ">")
          .replace(/&quot;/g, '"')
          .replace(/&apos;/g, "'");

        // Parse levels and find headings
        const levels = this.parseTOCFieldInstruction(fieldInstruction);
        const headings = this.findHeadingsForTOCFromXML(docXml, levels);

        if (headings.length === 0) continue;

        // Generate populated TOC
        const newTocXml = this.generateTOCXML(headings, fieldInstruction);
        modifiedXml = modifiedXml.replace(tocXml, newTocXml);
      } catch (error) {
        // Skip this TOC on error
        this.logger.error(
          "Error populating TOC",
          error instanceof Error
            ? {
                message: error.message,
                stack: error.stack,
              }
            : { error: String(error) }
        );
        continue;
      }
    }

    return modifiedXml;
  }

  /**
   * Replaces all Table of Contents in a saved document file with pre-populated entries
   *
   * This helper function works with saved DOCX files:
   * 1. Loads the document file
   * 2. Finds all TOC elements in document.xml
   * 3. For each TOC, parses its field instruction to determine which heading levels to include
   * 4. Scans the document for all headings matching those levels
   * 5. Generates pre-populated TOC entries with working hyperlinks
   * 6. Replaces the TOC in the XML and saves the file
   *
   * The generated TOC maintains the complex field structure, so users can still
   * right-click "Update Field" in Word to refresh it.
   *
   * @param filePath Path to the DOCX file to process
   * @returns Promise resolving to the number of TOC elements that were replaced
   *
   * @example
   * // Save document first
   * await doc.save('output.docx');
   *
   * // Then replace TOCs with populated entries
   * const count = await doc.replaceTableOfContents('output.docx');
   * console.log(`Replaced ${count} TOC element(s) with populated entries`);
   */
  public async replaceTableOfContents(filePath: string): Promise<number> {
    // Reuse the new extracted logic
    return await this.populateTOCsInFile(filePath);
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
   * Removes a specific header from the document
   * Removes the header from HeaderFooterManager, RelationshipManager, and section references
   * Also removes the header XML file from the ZIP archive
   * @param type Header type to remove (default, first, even)
   * @returns This document for chaining
   */
  removeHeader(type: "default" | "first" | "even"): this {
    const sectionProps = this.section.getProperties();

    // Get the relationship ID from section properties
    const rId = sectionProps.headers?.[type];
    if (!rId) {
      return this; // No header of this type exists
    }

    // Get the header filename from relationship
    const rel = this.relationshipManager.getRelationship(rId);
    if (rel) {
      const headerPath = `word/${rel.getTarget()}`;

      // Remove the header XML file from ZIP
      if (this.zipHandler.hasFile(headerPath)) {
        this.zipHandler.removeFile(headerPath);
      }

      // Remove the relationship
      this.relationshipManager.removeRelationship(rId);
    }

    // Remove from section references
    if (sectionProps.headers) {
      delete sectionProps.headers[type];

      // If no more headers, remove the headers object
      if (Object.keys(sectionProps.headers).length === 0) {
        delete sectionProps.headers;
      }
    }

    // Note: We can't directly remove from HeaderFooterManager without access to the Header object
    // The manager will be rebuilt on next save() call

    return this;
  }

  /**
   * Removes a specific footer from the document
   * Removes the footer from HeaderFooterManager, RelationshipManager, and section references
   * Also removes the footer XML file from the ZIP archive
   * @param type Footer type to remove (default, first, even)
   * @returns This document for chaining
   */
  removeFooter(type: "default" | "first" | "even"): this {
    const sectionProps = this.section.getProperties();

    // Get the relationship ID from section properties
    const rId = sectionProps.footers?.[type];
    if (!rId) {
      return this; // No footer of this type exists
    }

    // Get the footer filename from relationship
    const rel = this.relationshipManager.getRelationship(rId);
    if (rel) {
      const footerPath = `word/${rel.getTarget()}`;

      // Remove the footer XML file from ZIP
      if (this.zipHandler.hasFile(footerPath)) {
        this.zipHandler.removeFile(footerPath);
      }

      // Remove the relationship
      this.relationshipManager.removeRelationship(rId);
    }

    // Remove from section references
    if (sectionProps.footers) {
      delete sectionProps.footers[type];

      // If no more footers, remove the footers object
      if (Object.keys(sectionProps.footers).length === 0) {
        delete sectionProps.footers;
      }
    }

    // Note: We can't directly remove from HeaderFooterManager without access to the Footer object
    // The manager will be rebuilt on next save() call

    return this;
  }

  /**
   * Removes all headers from the document
   * Removes all header relationships, section references, and header XML files from the ZIP archive
   * @returns This document for chaining
   */
  clearHeaders(): this {
    const sectionProps = this.section.getProperties();

    // Remove each header type
    if (sectionProps.headers) {
      const types = Object.keys(sectionProps.headers) as Array<
        "default" | "first" | "even"
      >;
      for (const type of types) {
        this.removeHeader(type);
      }
    }

    // Note: Don't call headerFooterManager.clear() as that would clear footers too
    // The manager will be rebuilt correctly during save based on section properties

    return this;
  }

  /**
   * Removes all footers from the document
   * Removes all footer relationships, section references, and footer XML files from the ZIP archive
   * @returns This document for chaining
   */
  clearFooters(): this {
    const sectionProps = this.section.getProperties();

    // Remove each footer type
    if (sectionProps.footers) {
      const types = Object.keys(sectionProps.footers) as Array<
        "default" | "first" | "even"
      >;
      for (const type of types) {
        this.removeFooter(type);
      }
    }

    // Note: Don't call headerFooterManager.clear() as that would clear headers too
    // The manager will be rebuilt correctly during save based on section properties

    return this;
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
   * Adds or retrieves a "_top" bookmark at the beginning of the document
   *
   * This is a convenience method that ensures a "_top" bookmark exists at the start of the document body.
   * The bookmark is placed in an empty paragraph at position 0, with the structure:
   * ```xml
   * <w:bookmarkStart w:id="0" w:name="_top"/>
   * <w:bookmarkEnd w:id="0"/>
   * ```
   *
   * This method is idempotent - calling it multiple times will not create duplicate bookmarks.
   * If the "_top" bookmark already exists, the existing bookmark is returned.
   *
   * @returns Object containing:
   *  - `bookmark`: The Bookmark instance for "_top"
   *  - `anchor`: The anchor name ("_top") for creating hyperlinks
   *  - `hyperlink`: Convenience function to create a hyperlink to this bookmark
   *
   * @example
   * ```typescript
   * // Add or retrieve the _top bookmark
   * const { bookmark, anchor, hyperlink } = doc.addTopBookmark();
   *
   * // Create a hyperlink to the top of the document
   * const link = hyperlink('Back to top');
   * paragraph.addHyperlink(link);
   *
   * // Or create it manually
   * const link2 = Hyperlink.createInternal(anchor, 'Go to top');
   * ```
   */
  addTopBookmark(): {
    bookmark: Bookmark;
    anchor: string;
    hyperlink: (text: string, formatting?: RunFormatting) => Hyperlink;
  } {
    const BOOKMARK_NAME = "_top";

    // Check if _top bookmark already exists
    let bookmark = this.getBookmark(BOOKMARK_NAME);

    if (!bookmark) {
      // Create new _top bookmark
      // Note: We use skipNormalization to preserve the exact name "_top"
      // and explicitly set id to 0 as per the XML structure requirement
      bookmark = new Bookmark({
        id: 0,
        name: BOOKMARK_NAME,
        skipNormalization: true,
      });

      // Register the bookmark with the manager
      this.bookmarkManager.register(bookmark);

      // Add bookmark to the first existing paragraph if document has content
      // This avoids creating a visible newline at the top of the document
      const paragraphs = this.getParagraphs();

      if (paragraphs.length > 0 && paragraphs[0]) {
        // Add bookmark to first existing paragraph
        paragraphs[0].addBookmark(bookmark);
      } else {
        // Fallback: Create empty paragraph if document is empty
        const topParagraph = new Paragraph();
        topParagraph.addBookmark(bookmark);
        this.bodyElements.unshift(topParagraph);
      }
    }

    // Return the bookmark information and a convenience function
    return {
      bookmark,
      anchor: BOOKMARK_NAME,
      hyperlink: (text: string, formatting?: RunFormatting) => {
        return Hyperlink.createInternal(BOOKMARK_NAME, text, formatting);
      },
    };
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
   * Adds a tracked field instruction deletion to a paragraph
   * Uses w:delInstrText instead of w:delText for field codes
   * @param paragraph - The paragraph to add the deletion to
   * @param author - Author who made the deletion
   * @param fieldInstruction - Deleted field instruction text (e.g., 'PAGE', 'DATE')
   * @param date - Optional date (defaults to now)
   * @returns The created revision
   * @example
   * // Track deletion of a PAGE field
   * const para = doc.createParagraph();
   * doc.trackFieldInstructionDeletion(para, 'Alice', 'PAGE \\* MERGEFORMAT');
   */
  trackFieldInstructionDeletion(
    paragraph: Paragraph,
    author: string,
    fieldInstruction: string,
    date?: Date
  ): Revision {
    const run = new Run(fieldInstruction);
    const revision = Revision.createFieldInstructionDeletion(author, run, date);
    this.revisionManager.register(revision);
    paragraph.addRevision(revision);
    return revision;
  }

  /**
   * Marks a paragraph mark as deleted (tracked change)
   *
   * When a paragraph mark is deleted, it indicates that the paragraph
   * was joined with the next paragraph. This creates a deletion marker
   * in the paragraph properties (w:pPr/w:rPr/w:del) per ECMA-376 Part 1 §17.13.5.14.
   *
   * @param paragraph - Paragraph whose mark is deleted
   * @param author - Author who deleted the paragraph mark
   * @param date - Optional date (defaults to now)
   * @returns The paragraph for chaining
   * @example
   * // Mark paragraph mark as deleted when joining paragraphs
   * const para = doc.createParagraph('First paragraph');
   * doc.trackParagraphMarkDeletion(para, 'Alice');
   * // In Word, this shows the ¶ symbol as deleted
   */
  trackParagraphMarkDeletion(
    paragraph: Paragraph,
    author: string,
    date?: Date
  ): Paragraph {
    const revisionId = this.revisionManager.getNextId();
    paragraph.markParagraphMarkAsDeleted(revisionId, author, date);
    return paragraph;
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
    propertyChanges: number;
    moves: number;
    tableCellChanges: number;
    authors: string[];
    nextId: number;
  } {
    return this.revisionManager.getStats();
  }

  /**
   * Enables track changes for this document
   * When enabled, the w:trackRevisions flag is added to settings.xml
   * @param options - Optional track changes settings
   */
  enableTrackChanges(options?: {
    trackFormatting?: boolean;
    showInsertionsAndDeletions?: boolean;
    showFormatting?: boolean;
    showInkAnnotations?: boolean;
  }): this {
    this.trackChangesEnabled = true;

    if (options) {
      if (options.trackFormatting !== undefined) {
        this.trackFormatting = options.trackFormatting;
      }
      if (options.showInsertionsAndDeletions !== undefined) {
        this.revisionViewSettings.showInsertionsAndDeletions =
          options.showInsertionsAndDeletions;
      }
      if (options.showFormatting !== undefined) {
        this.revisionViewSettings.showFormatting = options.showFormatting;
      }
      if (options.showInkAnnotations !== undefined) {
        this.revisionViewSettings.showInkAnnotations =
          options.showInkAnnotations;
      }
    }

    return this;
  }

  /**
   * Disables track changes for this document
   */
  disableTrackChanges(): this {
    this.trackChangesEnabled = false;
    return this;
  }

  /**
   * Checks if track changes is enabled
   * @returns True if track changes is enabled
   */
  isTrackChangesEnabled(): boolean {
    return this.trackChangesEnabled;
  }

  /**
   * Gets the track formatting setting
   * @returns True if formatting changes are tracked
   */
  isTrackFormattingEnabled(): boolean {
    return this.trackFormatting;
  }

  /**
   * Gets the revision view settings
   * @returns Revision view settings
   */
  getRevisionViewSettings(): {
    showInsertionsAndDeletions: boolean;
    showFormatting: boolean;
    showInkAnnotations: boolean;
  } {
    return { ...this.revisionViewSettings };
  }

  /**
   * Sets the RSID root for this document
   * RSID (Revision Save ID) identifies the first editing session
   * @param rsidRoot - 8-character hexadecimal RSID value
   */
  setRsidRoot(rsidRoot: string): this {
    // Validate RSID format (8 hex characters)
    if (!/^[0-9A-Fa-f]{8}$/.test(rsidRoot)) {
      throw new Error("RSID must be an 8-character hexadecimal value");
    }
    this.rsidRoot = rsidRoot.toUpperCase();
    this.rsids.add(this.rsidRoot);
    return this;
  }

  /**
   * Adds an RSID to the document
   * Each editing session gets a unique RSID
   * @param rsid - 8-character hexadecimal RSID value
   */
  addRsid(rsid: string): this {
    // Validate RSID format
    if (!/^[0-9A-Fa-f]{8}$/.test(rsid)) {
      throw new Error("RSID must be an 8-character hexadecimal value");
    }
    this.rsids.add(rsid.toUpperCase());
    return this;
  }

  /**
   * Generates a new random RSID and adds it to the document
   * @returns The generated RSID
   */
  generateRsid(): string {
    const rsid = Math.floor(Math.random() * 0xffffffff)
      .toString(16)
      .toUpperCase()
      .padStart(8, "0");
    this.rsids.add(rsid);
    return rsid;
  }

  /**
   * Gets the RSID root value
   * @returns RSID root or undefined if not set
   */
  getRsidRoot(): string | undefined {
    return this.rsidRoot;
  }

  /**
   * Gets all RSIDs in the document
   * @returns Array of RSID values
   */
  getRsids(): string[] {
    return Array.from(this.rsids);
  }

  /**
   * Protects the document with specified edit restrictions
   * @param protection - Document protection settings
   */
  protectDocument(protection: {
    edit: "readOnly" | "comments" | "trackedChanges" | "forms";
    enforcement?: boolean;
    password?: string;
    cryptProviderType?: string;
    cryptAlgorithmClass?: string;
    cryptAlgorithmType?: string;
    cryptAlgorithmSid?: number;
    cryptSpinCount?: number;
  }): this {
    this.documentProtection = {
      edit: protection.edit,
      enforcement: protection.enforcement ?? true,
      cryptProviderType: protection.cryptProviderType,
      cryptAlgorithmClass: protection.cryptAlgorithmClass,
      cryptAlgorithmType: protection.cryptAlgorithmType,
      cryptAlgorithmSid: protection.cryptAlgorithmSid,
      cryptSpinCount: protection.cryptSpinCount,
    };

    // If password provided, generate hash and salt
    if (protection.password) {
      // For now, use a simple hash. In production, use proper cryptographic functions
      const crypto = require("crypto");
      const salt = crypto.randomBytes(16).toString("base64");
      const hash = crypto
        .pbkdf2Sync(
          protection.password,
          salt,
          protection.cryptSpinCount || 100000,
          32,
          "sha512"
        )
        .toString("base64");

      this.documentProtection.hash = hash;
      this.documentProtection.salt = salt;
    }

    return this;
  }

  /**
   * Removes document protection
   */
  unprotectDocument(): this {
    this.documentProtection = undefined;
    return this;
  }

  /**
   * Checks if document is protected
   * @returns True if document has protection enabled
   */
  isProtected(): boolean {
    return this.documentProtection !== undefined;
  }

  /**
   * Gets document protection settings
   * @returns Document protection settings or undefined
   */
  getProtection(): typeof this.documentProtection {
    return this.documentProtection;
  }

  /**
   * Creates and registers a run properties change revision
   * @param author - Author who made the change
   * @param content - Content with changed formatting
   * @param previousProperties - Previous run properties
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createRunPropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    const revision = Revision.createRunPropertiesChange(
      author,
      content,
      previousProperties,
      date
    );
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a paragraph properties change revision
   * @param author - Author who made the change
   * @param content - Paragraph content
   * @param previousProperties - Previous paragraph properties
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createParagraphPropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    const revision = Revision.createParagraphPropertiesChange(
      author,
      content,
      previousProperties,
      date
    );
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a table properties change revision
   * @param author - Author who made the change
   * @param content - Table content
   * @param previousProperties - Previous table properties
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createTablePropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    const revision = Revision.createTablePropertiesChange(
      author,
      content,
      previousProperties,
      date
    );
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a moveFrom revision (source of moved content)
   * @param author - Author who moved the content
   * @param content - Content that was moved
   * @param moveId - Unique move operation ID (links moveFrom and moveTo)
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createMoveFrom(
    author: string,
    content: Run | Run[],
    moveId: string,
    date?: Date
  ): Revision {
    const revision = Revision.createMoveFrom(author, content, moveId, date);
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a moveTo revision (destination of moved content)
   * @param author - Author who moved the content
   * @param content - Content that was moved
   * @param moveId - Unique move operation ID (links moveFrom and moveTo)
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createMoveTo(
    author: string,
    content: Run | Run[],
    moveId: string,
    date?: Date
  ): Revision {
    const revision = Revision.createMoveTo(author, content, moveId, date);
    return this.revisionManager.register(revision);
  }

  /**
   * Creates a pair of moveFrom and moveTo revisions for moving content
   * @param author - Author who moved the content
   * @param content - Content that was moved
   * @param date - Optional date (defaults to now)
   * @returns Object with both moveFrom and moveTo revisions and range markers
   */
  trackMove(
    author: string,
    content: Run | Run[],
    date?: Date
  ): {
    moveFrom: Revision;
    moveTo: Revision;
    moveId: string;
    moveFromRangeStart: RangeMarker;
    moveFromRangeEnd: RangeMarker;
    moveToRangeStart: RangeMarker;
    moveToRangeEnd: RangeMarker;
  } {
    // Generate unique move ID and name
    const moveId = `move${Date.now()}_${Math.random()
      .toString(36)
      .substr(2, 9)}`;
    const moveName = `move${Date.now()}`;

    // Get unique IDs for range markers (use revision manager's next ID)
    const rangeIdStart = this.revisionManager.getStats().nextId;

    // Create range markers for moveFrom
    const moveFromRangeStart = RangeMarker.createMoveFromStart(
      rangeIdStart,
      moveName,
      author,
      date
    );
    const moveFromRangeEnd = RangeMarker.createMoveFromEnd(rangeIdStart);

    // Create range markers for moveTo
    const moveToRangeStart = RangeMarker.createMoveToStart(
      rangeIdStart,
      moveName,
      author,
      date
    );
    const moveToRangeEnd = RangeMarker.createMoveToEnd(rangeIdStart);

    // Create the actual move revisions
    const moveFrom = this.createMoveFrom(author, content, moveId, date);
    const moveTo = this.createMoveTo(author, content, moveId, date);

    return {
      moveFrom,
      moveTo,
      moveId,
      moveFromRangeStart,
      moveFromRangeEnd,
      moveToRangeStart,
      moveToRangeEnd,
    };
  }

  /**
   * Creates and registers a table cell insertion revision
   * @param author - Author who inserted the cell
   * @param content - Cell content
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createTableCellInsert(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    const revision = Revision.createTableCellInsert(author, content, date);
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a table cell deletion revision
   * @param author - Author who deleted the cell
   * @param content - Cell content
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createTableCellDelete(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    const revision = Revision.createTableCellDelete(author, content, date);
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a table cell merge revision
   * @param author - Author who merged cells
   * @param content - Merged cell content
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createTableCellMerge(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    const revision = Revision.createTableCellMerge(author, content, date);
    return this.revisionManager.register(revision);
  }

  /**
   * Creates and registers a numbering change revision
   * @param author - Author who changed the numbering
   * @param content - Content with changed numbering
   * @param previousProperties - Previous numbering properties
   * @param date - Optional date (defaults to now)
   * @returns The created and registered revision
   */
  createNumberingChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    const revision = Revision.createNumberingChange(
      author,
      content,
      previousProperties,
      date
    );
    return this.revisionManager.register(revision);
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
        content = file.content.toString("utf-8");
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
   * Advanced find and replace with regex support and track changes integration
   *
   * Performs global find and replace operations with support for:
   * - Regular expressions for complex pattern matching
   * - Case-sensitive and whole-word matching
   * - Track changes integration (creates revision objects)
   * - Replacement across paragraphs and table cells
   *
   * @param pattern - String or RegExp pattern to find
   * @param replacement - Replacement text
   * @param options - Search and tracking options
   * @returns Object with replacement count and optional revisions
   *
   * @example
   * ```typescript
   * // Simple text replacement
   * const result = doc.findAndReplaceAll('old text', 'new text');
   * console.log(`Made ${result.count} replacements`);
   *
   * // Regex replacement
   * const phoneResult = doc.findAndReplaceAll(
   *   /\d{3}-\d{4}/g,
   *   '***-****',
   *   { caseSensitive: true }
   * );
   *
   * // With track changes
   * const tracked = doc.findAndReplaceAll('error', 'correction', {
   *   trackChanges: true,
   *   author: 'John Doe'
   * });
   * console.log(`Created ${tracked.revisions?.length} revisions`);
   *
   * // Whole word replacement
   * doc.findAndReplaceAll('test', 'exam', { wholeWord: true });
   * ```
   */
  findAndReplaceAll(
    pattern: string | RegExp,
    replacement: string,
    options?: {
      caseSensitive?: boolean;
      wholeWord?: boolean;
      trackChanges?: boolean;
      author?: string;
    }
  ): { count: number; revisions?: Revision[] } {
    const {
      caseSensitive = false,
      wholeWord = false,
      trackChanges = false,
      author = "Unknown",
    } = options || {};

    let count = 0;
    const revisions: Revision[] = [];

    // Convert pattern to RegExp if it's a string
    let regex: RegExp;
    if (typeof pattern === "string") {
      // Escape special regex characters
      const escaped = pattern.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const boundaryPattern = wholeWord ? `\\b${escaped}\\b` : escaped;
      const flags = caseSensitive ? "g" : "gi";
      regex = new RegExp(boundaryPattern, flags);
    } else {
      // Use provided RegExp, ensure global flag
      const flags = pattern.flags.includes("g")
        ? pattern.flags
        : pattern.flags + "g";
      regex = new RegExp(pattern.source, flags);
    }

    // Process all runs in document
    const runs = this.getAllRuns();

    for (const run of runs) {
      const originalText = run.getText();
      const matches = originalText.match(regex);

      if (matches && matches.length > 0) {
        const newText = originalText.replace(regex, replacement);

        if (trackChanges) {
          // Create deletion revision for original text
          const deletionRun = new Run(originalText, run.getFormatting());
          const deletion = Revision.createDeletion(author, deletionRun);
          revisions.push(deletion);

          // Create insertion revision for new text
          const insertionRun = new Run(newText, run.getFormatting());
          const insertion = Revision.createInsertion(author, insertionRun);
          revisions.push(insertion);

          // Register revisions with the document
          this.revisionManager.register(deletion);
          this.revisionManager.register(insertion);
        }

        // Update the run text
        run.setText(newText);
        count += matches.length;
      }
    }

    return trackChanges ? { count, revisions } : { count };
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
   * Validates a paragraph before insertion
   * @param paragraph - The paragraph to validate
   * @throws Error if paragraph is invalid
   */
  private validateParagraph(paragraph: Paragraph): void {
    // Type validation
    if (!(paragraph instanceof Paragraph)) {
      throw new Error(
        "insertParagraphAt: parameter must be a Paragraph instance"
      );
    }

    // Check for duplicate paragraph IDs
    const paraId = paragraph.getFormatting().paraId;
    if (paraId) {
      const existingIds = this.bodyElements
        .filter((el): el is Paragraph => el instanceof Paragraph)
        .map((p) => p.getFormatting().paraId)
        .filter((id) => id === paraId);

      if (existingIds.length > 0) {
        throw new Error(
          `Duplicate paragraph ID detected: ${paraId}. Each paragraph must have a unique ID.`
        );
      }
    }

    // Warn about missing styles
    const style = paragraph.getFormatting().style;
    if (style && !this.stylesManager.hasStyle(style)) {
      defaultLogger.warn(
        `Style "${style}" not found in document. Paragraph will fall back to Normal style.`
      );
    }

    // Warn about missing numbering
    const numbering = paragraph.getFormatting().numbering;
    if (numbering && !this.numberingManager.hasInstance(numbering.numId)) {
      defaultLogger.warn(
        `Numbering ID ${numbering.numId} not found in document. List formatting will not display.`
      );
    }
  }

  /**
   * Validates a table before insertion
   * @param table - The table to validate
   * @throws Error if table is invalid
   */
  private validateTable(table: Table): void {
    // Type validation
    if (!(table instanceof Table)) {
      throw new Error("insertTableAt: parameter must be a Table instance");
    }

    // Content validation - table must have rows
    const rows = table.getRows();
    if (rows.length === 0) {
      throw new Error("insertTableAt: table must have at least one row");
    }

    // Check first row has cells (rows.length > 0 already checked above)
    const firstRow = rows[0];
    if (firstRow && firstRow.getCells().length === 0) {
      throw new Error("insertTableAt: table rows must have at least one cell");
    }

    // Warn about missing table styles
    const tableStyle = table.getFormatting().style;
    if (tableStyle && !this.stylesManager.hasStyle(tableStyle)) {
      defaultLogger.warn(
        `Table style "${tableStyle}" not found in document. Table will use default formatting.`
      );
    }
  }

  /**
   * Validates a TOC element before insertion
   * @param toc - The TOC to validate
   * @throws Error if TOC is invalid
   */
  private validateToc(toc: TableOfContentsElement): void {
    // Type validation
    if (!(toc instanceof TableOfContentsElement)) {
      throw new Error(
        "insertTocAt: parameter must be a TableOfContentsElement instance"
      );
    }

    // Check if document has heading styles for TOC to reference
    const hasHeadings = [
      "Heading1",
      "Heading2",
      "Heading3",
      "Heading4",
      "Heading5",
      "Heading6",
      "Heading7",
      "Heading8",
      "Heading9",
    ].some((style) => this.stylesManager.hasStyle(style));

    if (!hasHeadings) {
      defaultLogger.warn(
        "No heading styles found in document. Table of Contents may not display entries correctly."
      );
    }
  }

  /**
   * Normalizes and validates insertion index
   * @param index - The requested index
   * @returns Normalized index within valid bounds
   */
  private normalizeIndex(index: number): number {
    if (index < 0) {
      return 0;
    } else if (index > this.bodyElements.length) {
      return this.bodyElements.length;
    }
    return index;
  }

  /**
   * Inserts a paragraph at a specific position
   * @param index - The position to insert at (0-based)
   * @param paragraph - The paragraph to insert
   * @returns This document for chaining
   * @throws Error if paragraph is invalid or has duplicate IDs
   */
  insertParagraphAt(index: number, paragraph: Paragraph): this {
    // Validate the paragraph
    this.validateParagraph(paragraph);

    // Normalize index
    index = this.normalizeIndex(index);

    // Insert the paragraph
    this.bodyElements.splice(index, 0, paragraph);
    return this;
  }

  /**
   * Inserts a table at a specific position
   * @param index - The position to insert at (0-based)
   * @param table - The table to insert
   * @returns This document for chaining
   * @throws Error if table is invalid or malformed
   * @example
   * ```typescript
   * const table = new Table(2, 3);
   * doc.insertTableAt(5, table);  // Insert at position 5
   * ```
   */
  insertTableAt(index: number, table: Table): this {
    // Validate the table
    this.validateTable(table);

    // Normalize index
    index = this.normalizeIndex(index);

    // Insert the table
    this.bodyElements.splice(index, 0, table);
    return this;
  }

  /**
   * Inserts a Table of Contents at a specific position
   * @param index - The position to insert at (0-based)
   * @param toc - The TableOfContentsElement to insert
   * @returns This document for chaining
   * @throws Error if TOC is invalid
   * @example
   * ```typescript
   * const toc = TableOfContentsElement.createStandard();
   * doc.insertTocAt(0, toc);  // Insert at beginning
   * ```
   */
  insertTocAt(index: number, toc: TableOfContentsElement): this {
    // Validate the TOC
    this.validateToc(toc);

    // Normalize index
    index = this.normalizeIndex(index);

    // Insert the TOC
    this.bodyElements.splice(index, 0, toc);
    return this;
  }

  /**
   * Replaces a paragraph at a specific position
   * @param index - The position to replace at (0-based)
   * @param paragraph - The new paragraph
   * @returns True if replaced, false if index invalid
   * @throws Error if replacement paragraph is invalid or has duplicate IDs
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
        // Validate the replacement paragraph
        this.validateParagraph(paragraph);

        // Replace the paragraph
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
   * @throws Error if replacement table is invalid or malformed
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
        // Validate the replacement table
        this.validateTable(table);

        // Replace the table
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
   * Defragments and optimizes hyperlinks in the document
   * This merges fragmented hyperlinks with the same URL (common in Google Docs exports)
   * and optionally resets their formatting to standard style
   *
   * @param options - Defragmentation options
   * @returns Number of hyperlinks merged
   *
   * @example
   * ```typescript
   * // Basic defragmentation
   * const merged = doc.defragmentHyperlinks();
   * console.log(`Merged ${merged} fragmented hyperlinks`);
   *
   * // With formatting reset (fixes corrupted fonts like Caveat)
   * const fixed = doc.defragmentHyperlinks({ resetFormatting: true });
   * console.log(`Fixed ${fixed} hyperlinks with standard formatting`);
   * ```
   */
  defragmentHyperlinks(options?: {
    resetFormatting?: boolean;
    cleanupRelationships?: boolean;
  }): number {
    const { resetFormatting = false, cleanupRelationships = false } =
      options || {};
    let mergedCount = 0;

    // Get the DocumentParser instance to use its merging method
    const parser = new DocumentParser();

    // Process all paragraphs in the document
    for (const paragraph of this.getParagraphs()) {
      const originalContent = paragraph.getContent();

      // Call the enhanced mergeConsecutiveHyperlinks method
      (parser as any).mergeConsecutiveHyperlinks(paragraph, resetFormatting);

      const newContent = paragraph.getContent();

      // Count merges by comparing content length
      if (originalContent.length > newContent.length) {
        mergedCount += originalContent.length - newContent.length;
      }
    }

    // Also process paragraphs in tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          const cellParagraphs =
            cell instanceof TableCell ? cell.getParagraphs() : [];
          for (const para of cellParagraphs) {
            const originalContent = para.getContent();

            (parser as any).mergeConsecutiveHyperlinks(para, resetFormatting);

            const newContent = para.getContent();

            if (originalContent.length > newContent.length) {
              mergedCount += originalContent.length - newContent.length;
            }
          }
        }
      }
    }

    // Optionally clean up orphaned relationships
    if (cleanupRelationships && mergedCount > 0) {
      // Collect all referenced hyperlink relationship IDs
      const referencedIds = new Set<string>();

      // Collect IDs from all hyperlinks in the document
      const allHyperlinks = this.getHyperlinks();
      for (const { hyperlink } of allHyperlinks) {
        const relId = hyperlink.getRelationshipId();
        if (relId) {
          referencedIds.add(relId);
        }
      }

      // Remove orphaned hyperlink relationships
      const removedCount =
        this.relationshipManager.removeOrphanedHyperlinks(referencedIds);
      if (removedCount > 0) {
        defaultLogger.info(
          `Cleaned up ${removedCount} orphaned hyperlink relationship(s)`
        );
      }
    }

    return mergedCount;
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
   * Gets all runs in the document (flattened from all paragraphs)
   *
   * This method returns all Run objects from:
   * - All body paragraphs
   * - All paragraphs inside table cells
   *
   * Useful for bulk operations on text formatting across the entire document.
   *
   * @returns Array of all Run objects in the document
   *
   * @example
   * ```typescript
   * // Get all runs
   * const runs = doc.getAllRuns();
   * console.log(`Document has ${runs.length} text runs`);
   *
   * // Make all text bold
   * for (const run of doc.getAllRuns()) {
   *   run.setBold(true);
   * }
   * ```
   */
  getAllRuns(): Run[] {
    const runs: Run[] = [];

    // Get runs from all body paragraphs
    for (const paragraph of this.getParagraphs()) {
      runs.push(...paragraph.getRuns());
    }

    // Get runs from paragraphs inside table cells
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          const cellParagraphs =
            cell instanceof TableCell ? cell.getParagraphs() : [];
          for (const para of cellParagraphs) {
            runs.push(...para.getRuns());
          }
        }
      }
    }

    return runs;
  }

  /**
   * Removes a specific formatting type from all runs in the document
   *
   * This is a bulk operation that removes the specified formatting property
   * from ALL text runs in the document (including runs inside table cells).
   *
   * @param type - The formatting property to remove
   * @returns Number of runs that were modified
   *
   * @example
   * ```typescript
   * // Remove all bold formatting from document
   * const count = doc.removeFormattingFromAll('bold');
   * console.log(`Removed bold from ${count} runs`);
   *
   * // Remove all highlighting
   * doc.removeFormattingFromAll('highlight');
   *
   * // Remove all font color
   * doc.removeFormattingFromAll('color');
   *
   * // Remove underlines
   * doc.removeFormattingFromAll('underline');
   * ```
   */
  removeFormattingFromAll(
    type:
      | "bold"
      | "italic"
      | "underline"
      | "strike"
      | "dstrike"
      | "highlight"
      | "color"
      | "font"
      | "size"
      | "subscript"
      | "superscript"
      | "smallCaps"
      | "allCaps"
      | "outline"
      | "shadow"
      | "emboss"
      | "imprint"
  ): number {
    let modifiedCount = 0;

    // Get all runs in the document
    const runs = this.getAllRuns();

    for (const run of runs) {
      const formatting = run.getFormatting();

      // Check if the property exists before removing it
      if (type in formatting) {
        // Access the private formatting property to modify it
        // This is a valid pattern for bulk operations in the framework
        (run as any).formatting[type] = undefined;
        delete (run as any).formatting[type];

        modifiedCount++;
      }
    }

    return modifiedCount;
  }

  /**
   * Applies a formatting function to all hyperlinks in the document
   *
   * This is a bulk operation that calls the provided formatter function
   * for each hyperlink in the document (including hyperlinks inside table cells).
   *
   * The formatter function receives the hyperlink and its containing paragraph,
   * allowing for sophisticated conditional formatting based on context.
   *
   * @param formatter - Function to apply to each hyperlink
   * @returns Number of hyperlinks processed
   *
   * @example
   * ```typescript
   * // Make all hyperlinks red and bold
   * doc.updateAllHyperlinks((link) => {
   *   link.setFormatting({ color: 'FF0000', bold: true });
   * });
   *
   * // Remove underline from all hyperlinks
   * doc.updateAllHyperlinks((link) => {
   *   const fmt = link.getFormatting();
   *   delete fmt.underline;
   *   link.setFormatting(fmt);
   * });
   *
   * // Conditional formatting based on URL
   * doc.updateAllHyperlinks((link) => {
   *   const url = link.getUrl();
   *   if (url?.includes('internal')) {
   *     link.setFormatting({ color: '0000FF' }); // Blue for internal links
   *   } else if (url?.includes('external')) {
   *     link.setFormatting({ color: 'FF0000' }); // Red for external links
   *   }
   * });
   *
   * // Access paragraph context for advanced logic
   * doc.updateAllHyperlinks((link, para) => {
   *   const paraStyle = para.getFormatting().style;
   *   if (paraStyle === 'Heading1') {
   *     link.setFormatting({ bold: true, size: 16 });
   *   }
   * });
   * ```
   */
  updateAllHyperlinks(
    formatter: (hyperlink: Hyperlink, paragraph: Paragraph) => void
  ): number {
    // Get all hyperlinks with their containing paragraphs
    const hyperlinks = this.getHyperlinks();

    // Apply formatter to each hyperlink
    for (const { hyperlink, paragraph } of hyperlinks) {
      formatter(hyperlink, paragraph);
    }

    return hyperlinks.length;
  }

  /**
   * Normalizes spacing throughout the document
   *
   * Ensures consistent spacing by:
   * - Removing duplicate consecutive empty paragraphs
   * - Applying standard paragraph spacing (before/after)
   * - Standardizing line spacing
   * - Removing trailing spaces from runs
   *
   * @param rules - Normalization rules
   * @returns Object with counts of elements removed and normalized
   *
   * @example
   * ```typescript
   * // Remove duplicate empty paragraphs only
   * const result = doc.normalizeSpacing({
   *   removeDuplicateEmptyParagraphs: true
   * });
   * console.log(`Removed ${result.removed} empty paragraphs`);
   *
   * // Apply standard spacing
   * doc.normalizeSpacing({
   *   standardParagraphSpacing: { before: 0, after: 200 }, // 200 twips = 10pt
   *   standardLineSpacing: 240, // Single spacing
   *   removeTrailingSpaces: true
   * });
   *
   * // Full normalization
   * const stats = doc.normalizeSpacing({
   *   removeDuplicateEmptyParagraphs: true,
   *   standardParagraphSpacing: { after: 200 },
   *   removeTrailingSpaces: true
   * });
   * console.log(`Removed: ${stats.removed}, Normalized: ${stats.normalized}`);
   * ```
   */
  normalizeSpacing(
    rules: {
      removeDuplicateEmptyParagraphs?: boolean;
      standardParagraphSpacing?: { before?: number; after?: number };
      standardLineSpacing?: number;
      removeTrailingSpaces?: boolean;
    } = {}
  ): { removed: number; normalized: number } {
    const {
      removeDuplicateEmptyParagraphs = true,
      standardParagraphSpacing,
      standardLineSpacing,
      removeTrailingSpaces = true,
    } = rules;

    let removed = 0;
    let normalized = 0;

    // Remove duplicate empty paragraphs
    if (removeDuplicateEmptyParagraphs) {
      let lastWasEmpty = false;
      const toRemove: number[] = [];

      this.bodyElements.forEach((element, index) => {
        if (element instanceof Paragraph) {
          const isEmpty = element.getText().trim() === "";
          if (isEmpty && lastWasEmpty) {
            toRemove.push(index);
          }
          lastWasEmpty = isEmpty;
        } else {
          lastWasEmpty = false; // Reset for non-paragraph elements
        }
      });

      // Remove in reverse order to maintain indices
      toRemove.reverse().forEach((index) => {
        this.bodyElements.splice(index, 1);
        removed++;
      });
    }

    // Apply standard spacing to all paragraphs
    for (const para of this.getParagraphs()) {
      if (standardParagraphSpacing) {
        if (standardParagraphSpacing.before !== undefined) {
          para.setSpaceBefore(standardParagraphSpacing.before);
          normalized++;
        }
        if (standardParagraphSpacing.after !== undefined) {
          para.setSpaceAfter(standardParagraphSpacing.after);
          normalized++;
        }
      }

      if (standardLineSpacing !== undefined) {
        para.setLineSpacing(standardLineSpacing, "auto");
        normalized++;
      }

      // Remove trailing spaces from runs
      if (removeTrailingSpaces) {
        const runs = para.getRuns();
        if (runs.length > 0) {
          const lastRun = runs[runs.length - 1];
          if (lastRun) {
            const text = lastRun.getText();
            const trimmed = text.trimEnd();
            if (text !== trimmed) {
              lastRun.setText(trimmed);
              normalized++;
            }
          }
        }
      }
    }

    // Also process tables
    for (const table of this.getTables()) {
      for (const row of table.getRows()) {
        for (const cell of row.getCells()) {
          const cellParagraphs =
            cell instanceof TableCell ? cell.getParagraphs() : [];
          for (const para of cellParagraphs) {
            if (standardParagraphSpacing) {
              if (standardParagraphSpacing.before !== undefined) {
                para.setSpaceBefore(standardParagraphSpacing.before);
                normalized++;
              }
              if (standardParagraphSpacing.after !== undefined) {
                para.setSpaceAfter(standardParagraphSpacing.after);
                normalized++;
              }
            }

            if (standardLineSpacing !== undefined) {
              para.setLineSpacing(standardLineSpacing, "auto");
              normalized++;
            }

            if (removeTrailingSpaces) {
              const runs = para.getRuns();
              if (runs.length > 0) {
                const lastRun = runs[runs.length - 1];
                if (lastRun) {
                  const text = lastRun.getText();
                  const trimmed = text.trimEnd();
                  if (text !== trimmed) {
                    lastRun.setText(trimmed);
                    normalized++;
                  }
                }
              }
            }
          }
        }
      }
    }

    return { removed, normalized };
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
