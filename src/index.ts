/**
 * DocXML - DOCX Editing Framework
 * Main entry point
 */

// Main ZIP handler
export { ZipHandler } from './zip/ZipHandler';
export { ZipReader } from './zip/ZipReader';
export { ZipWriter } from './zip/ZipWriter';

// Types
export {
  ZipFile,
  FileMap,
  LoadOptions,
  SaveOptions,
  AddFileOptions,
  REQUIRED_DOCX_FILES,
  DOCX_PATHS,
} from './zip/types';

// Errors
export {
  DocxError,
  DocxNotFoundError,
  InvalidDocxError,
  CorruptedArchiveError,
  MissingRequiredFileError,
  FileOperationError,
} from './zip/errors';

// Utilities
export {
  validateDocxStructure,
  isBinaryFile,
  normalizePath,
  isValidZipBuffer,
  isTextContent,
  validateTwips,
  validateColor,
  validateHexColor,
  validateNumberingId,
  validateLevel,
  validateAlignment,
  validateFontSize,
  validateNonEmptyString,
  validatePercentage,
  validateEmus,
  detectXmlInText,
  cleanXmlFromText,
  validateRunText,
  TextValidationResult,
} from './utils/validation';

// Corruption Detection
export {
  detectCorruptionInDocument,
  detectCorruptionInText,
  suggestFix,
  looksCorrupted,
  CorruptionReport,
  CorruptionLocation,
  CorruptionType,
} from './utils/corruptionDetection';

// Unit conversions
export {
  STANDARD_DPI,
  UNITS,
  PAGE_SIZES,
  COMMON_MARGINS,
  twipsToPoints,
  twipsToInches,
  twipsToCm,
  twipsToEmus,
  emusToTwips,
  emusToInches,
  emusToCm,
  emusToPoints,
  emusToPixels,
  pointsToTwips,
  pointsToEmus,
  pointsToInches,
  pointsToCm,
  inchesToTwips,
  inchesToEmus,
  inchesToPoints,
  inchesToCm,
  inchesToPixels,
  cmToTwips,
  cmToEmus,
  cmToInches,
  cmToPoints,
  cmToPixels,
  pixelsToEmus,
  pixelsToInches,
  pixelsToTwips,
  pixelsToCm,
  pixelsToPoints,
} from './utils/units';

// Core classes
export { Document, DocumentProperties, DocumentOptions, DocumentPart } from './core/Document';
export { Relationship, RelationshipType, RelationshipProperties } from './core/Relationship';
export { RelationshipManager } from './core/RelationshipManager';
export { DocumentParser, ParseError } from './core/DocumentParser';
export { DocumentGenerator } from './core/DocumentGenerator';
export { DocumentValidator, SizeEstimate, MemoryOptions } from './core/DocumentValidator';

// Formatting classes
export { Style, StyleType, StyleProperties } from './formatting/Style';
export { StylesManager, ValidationResult } from './formatting/StylesManager';
export {
  NumberingLevel,
  NumberFormat,
  NumberAlignment,
  NumberingLevelProperties,
} from './formatting/NumberingLevel';
export {
  AbstractNumbering,
  AbstractNumberingProperties,
} from './formatting/AbstractNumbering';
export {
  NumberingInstance,
  NumberingInstanceProperties,
} from './formatting/NumberingInstance';
export { NumberingManager } from './formatting/NumberingManager';

// Document elements
export { Paragraph, ParagraphAlignment, ParagraphFormatting } from './elements/Paragraph';
export { Run, RunFormatting } from './elements/Run';
export {
  Section,
  PageOrientation,
  SectionType,
  PageNumberFormat,
  PageSize,
  Margins,
  Columns,
  PageNumbering,
  SectionProperties,
} from './elements/Section';
export {
  Table,
  TableAlignment,
  TableLayout,
  TableBorder,
  TableBorders,
  TableFormatting,
} from './elements/Table';
export { TableRow, RowFormatting } from './elements/TableRow';
export {
  TableCell,
  BorderStyle,
  CellBorder,
  CellBorders,
  CellShading,
  CellVerticalAlignment,
  CellFormatting,
} from './elements/TableCell';
export { Image, ImageFormat, ImageProperties } from './elements/Image';
export { ImageManager } from './elements/ImageManager';
export { ImageRun } from './elements/ImageRun';
export { FontManager, FontFormat, FontEntry } from './elements/FontManager';
export { Field, FieldType, FieldProperties } from './elements/Field';
export { Header, HeaderType, HeaderProperties } from './elements/Header';
export { Footer, FooterType, FooterProperties } from './elements/Footer';
export { HeaderFooterManager } from './elements/HeaderFooterManager';
export { Hyperlink, HyperlinkProperties } from './elements/Hyperlink';
export { TableOfContents, TOCProperties } from './elements/TableOfContents';
export { TableOfContentsElement } from './elements/TableOfContentsElement';
export { Bookmark, BookmarkProperties } from './elements/Bookmark';
export { BookmarkManager } from './elements/BookmarkManager';
export { StructuredDocumentTag, SDTProperties, SDTLockType, SDTContent } from './elements/StructuredDocumentTag';
export { Revision, RevisionType, RevisionProperties } from './elements/Revision';
export { RevisionManager } from './elements/RevisionManager';
export { Comment, CommentProperties } from './elements/Comment';
export { CommentManager } from './elements/CommentManager';
export { Footnote, FootnoteType, FootnoteProperties } from './elements/Footnote';
export { FootnoteManager } from './elements/FootnoteManager';
export { Endnote, EndnoteType, EndnoteProperties } from './elements/Endnote';
export { EndnoteManager } from './elements/EndnoteManager';

// XML Builder and Parser
export { XMLBuilder, XMLElement } from './xml/XMLBuilder';
export {
  XMLParser,
  ParseToObjectOptions,
  ParsedXMLValue,
  ParsedXMLObject,
} from './xml/XMLParser';

// Logging utilities
export {
  ILogger,
  LogLevel,
  LogEntry,
  ConsoleLogger,
  SilentLogger,
  CollectingLogger,
  defaultLogger,
  createScopedLogger,
} from './utils/logger';

// Error handling utilities
export { isError, toError, wrapError, getErrorMessage } from './utils/errorHandling';

// Constants
export { LIMITS } from './constants/limits';
