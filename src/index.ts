/**
 * DocXML - DOCX Editing Framework
 * Main entry point
 */

// =============================================================================
// PUBLIC API — Core Document Classes
// =============================================================================

export { Document, DocumentProperties, DocumentOptions, DocumentLoadOptions, DocumentPart } from './core/Document';

// =============================================================================
// PUBLIC API — Document Elements
// =============================================================================

export {
  Paragraph,
  ParagraphAlignment,
  ParagraphFormatting,
  ParagraphContent,
  FieldLike,
  isRun,
  isField,
  isSimpleField,
  isComplexField,
  isHyperlink,
  isRevision,
  isRangeMarker,
  isShape,
  isTextBox,
} from './elements/Paragraph';
export { Run, RunFormatting, ThemeColorValue } from './elements/Run';
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
  LineNumbering,
  LineNumberingRestart,
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
  CellBorder,
  CellBorders,
  CellShading,
  CellFormatting,
} from './elements/TableCell';
export { TableGridChange, GridColumn, TableGridChangeProperties } from './elements/TableGridChange';
export {
  Image,
  ImageFormat,
  ImageProperties,
  ImageBorder,
  ImageEffects,
  PresetGeometry,
  BlipCompressionState,
  PicLockAttribute,
  PicNonVisualProperties,
} from './elements/Image';
export { ImageRun } from './elements/ImageRun';
export { Shape, ShapeType, ShapeProperties, ShapeFill, ShapeOutline } from './elements/Shape';
export { TextBox, TextBoxProperties, TextBoxFill, TextBoxMargins } from './elements/TextBox';
export { Header, HeaderType, HeaderProperties } from './elements/Header';
export { Footer, FooterType, FooterProperties } from './elements/Footer';
export { Hyperlink, HyperlinkProperties } from './elements/Hyperlink';
export { Bookmark, BookmarkProperties } from './elements/Bookmark';
export { RangeMarker, RangeMarkerType, RangeMarkerProperties } from './elements/RangeMarker';
export { Comment, CommentProperties } from './elements/Comment';
export { Footnote, FootnoteType, FootnoteProperties } from './elements/Footnote';
export { Endnote, EndnoteType, EndnoteProperties } from './elements/Endnote';
export {
  Field,
  FieldType,
  FieldProperties,
  ComplexField,
  ComplexFieldProperties,
  FieldCharType,
  TOCFieldOptions,
  createTOCField,
} from './elements/Field';
export {
  createNestedIFMergeField,
  createMergeField,
  createRefField,
  createIFField,
  createNestedField,
  parseHyperlinkInstruction,
  buildHyperlinkInstruction,
  isHyperlinkInstruction,
  ParsedHyperlinkInstruction,
} from './elements/FieldHelpers';
export { StructuredDocumentTag, SDTProperties, SDTLockType, SDTContent, SDTPlaceholder, SDTDataBinding, ContentControlType } from './elements/StructuredDocumentTag';
export { TableOfContents, TOCProperties } from './elements/TableOfContents';
export { TableOfContentsElement } from './elements/TableOfContentsElement';
export { AlternateContent } from './elements/AlternateContent';
export { MathParagraph, MathExpression } from './elements/MathElement';
export { CustomXmlBlock } from './elements/CustomXml';
export { PreservedElement, PreservedElementContext } from './elements/PreservedElement';

// =============================================================================
// PUBLIC API — Track Changes / Revisions
// =============================================================================

export { Revision, RevisionType, RevisionProperties, FieldContext } from './elements/Revision';
export { RevisionContent, isRunContent, isHyperlinkContent } from './elements/RevisionContent';
export {
  RevisionLocation,
  RunPropertyChange,
  ParagraphPropertyChange,
  ParagraphFormattingPartial,
  ParagraphBorderDef,
  ParagraphBorders,
  ParagraphShading,
  TabStopDef,
  PropertyChangeBase,
  TablePropertyChange,
  TablePropertyChangeType,
  SectionPropertyChange,
  NumberingChange,
  AnyPropertyChange,
  isRunPropertyChange,
  isParagraphPropertyChange,
  isTablePropertyChange,
  isSectionPropertyChange,
  isNumberingChange,
} from './elements/PropertyChangeTypes';
export {
  acceptRevisionsInMemory,
  AcceptRevisionsOptions,
  AcceptRevisionsResult,
  paragraphHasRevisions,
  getRevisionsFromParagraph,
  countRevisionsByType,
  stripRevisionsFromXml,
} from './utils/InMemoryRevisionAcceptor';
export {
  MoveOperationHelper,
  MoveOperationOptions,
  MoveOperationResult,
} from './utils/MoveOperationHelper';
export {
  SelectiveRevisionAcceptor,
  SelectiveAcceptResult,
} from './utils/SelectiveRevisionAcceptor';
export {
  RevisionAwareProcessor,
  RevisionHandlingMode,
  RevisionProcessingOptions,
  SelectionCriteria,
  RevisionProcessingResult,
  ConflictInfo,
  ProcessingLogEntry,
} from './utils/RevisionAwareProcessor';
export {
  ChangelogGenerator,
  ChangeEntry,
  ChangeCategory,
  ChangeLocation,
  ChangelogOptions,
  ChangelogFormat,
  ConsolidatedChange,
  ChangelogSummary,
} from './utils/ChangelogGenerator';

// =============================================================================
// PUBLIC API — Formatting / Styles / Numbering
// =============================================================================

export { Style, StyleType, StyleProperties } from './formatting/Style';
export { StylesManager, ValidationResult, LatentStylesConfig, LatentStyleException } from './formatting/StylesManager';
export {
  NumberingLevel,
  NumberFormat,
  NumberAlignment,
  NumberingLevelProperties,
  WORD_NATIVE_BULLETS,
  WordNativeBullet,
} from './formatting/NumberingLevel';
export {
  AbstractNumbering,
  AbstractNumberingProperties,
} from './formatting/AbstractNumbering';
export {
  NumberingInstance,
  NumberingInstanceProperties,
} from './formatting/NumberingInstance';
export {
  NumberingManager,
  NumberingConsolidationOptions,
  NumberingConsolidationResult,
} from './formatting/NumberingManager';
export {
  StyleRunFormatting,
  StyleParagraphFormatting,
  Heading2TableOptions,
  StyleConfig,
  Heading2Config,
  ApplyCustomFormattingOptions,
} from './types/styleConfig';
export {
  FormatOptions,
  StyleApplyOptions,
  EmphasisType,
  ListPrefix,
} from './types/formatting';

// =============================================================================
// PUBLIC API — Managers
// =============================================================================

export { RevisionManager, RevisionCategory, RevisionSummary } from './elements/RevisionManager';
export { ImageManager } from './elements/ImageManager';
export { BookmarkManager } from './elements/BookmarkManager';
export { CommentManager } from './elements/CommentManager';
export { FootnoteManager } from './elements/FootnoteManager';
export { EndnoteManager } from './elements/EndnoteManager';
export { HeaderFooterManager } from './elements/HeaderFooterManager';
export { FontManager, FontFormat, FontEntry } from './elements/FontManager';
export { DrawingManager, DrawingElement, DrawingType, PreservedDrawing } from './managers/DrawingManager';

// =============================================================================
// PUBLIC API — Image Optimization
// =============================================================================

export type { ImageOptimizationResult } from './images/ImageOptimizer';

// =============================================================================
// TYPES — Common / Shared Type Definitions
// =============================================================================

export {
  ShadingPattern,
  BasicShadingPattern,
  BorderStyle,
  ExtendedBorderStyle,
  FullBorderStyle,
  BorderDefinition,
  FourSidedBorders,
  TableBorderDefinitions,
  HorizontalAlignment,
  VerticalAlignment,
  PageVerticalAlignment,
  CellVerticalAlignment,
  ParagraphAlignment as CommonParagraphAlignment,
  TableAlignment as CommonTableAlignment,
  RowJustification,
  TextVerticalAlignment,
  TabAlignment,
  PositionAnchor,
  HorizontalAnchor,
  VerticalAnchor,
  TextDirection,
  SectionTextDirection,
  WidthType,
  ShadingConfig,
  buildShadingAttributes,
  TabLeader,
  TabStop,
  isShadingPattern,
  isBorderStyle,
  isHorizontalAlignment,
  isVerticalAlignment,
  isParagraphAlignment,
  isWidthType,
  DEFAULT_BORDER,
  NO_BORDER,
} from './elements/CommonTypes';
export {
  ListCategory,
  NumberFormat as ListNumberFormat,
  BulletFormat,
  ListDetectionResult,
  ListAnalysis,
  ListNormalizationOptions,
  ListNormalizationReport,
  IndentationLevel,
} from './types/list-types';

export {
  CompatibilityMode,
  CompatibilityInfo,
  CompatSetting,
} from './types/compatibility-types';

export {
  DocumentProtection,
  RevisionViewSettings,
  TrackChangesSettings,
  WebSettingsInfo,
} from './types/settings-types';

// =============================================================================
// TYPES — Compatibility Upgrade
// =============================================================================

export { CompatibilityUpgrader, UpgradeReport } from './utils/CompatibilityUpgrader';
export {
  LEGACY_COMPAT_ELEMENTS,
  LEGACY_COMPAT_ELEMENT_NAMES,
  MODERN_COMPAT_SETTINGS,
  MS_WORD_COMPAT_URI,
} from './constants/legacyCompatFlags';

// =============================================================================
// UTILITIES — Unit Conversions
// =============================================================================

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
  pointsToHalfPoints,
  halfPointsToPoints,
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

// =============================================================================
// UTILITIES — Validation, Corruption Detection, Error Handling
// =============================================================================

export {
  validateDocxStructure,
  isBinaryFile,
  normalizePath,
  isValidZipBuffer,
  isTextContent,
  validateTwips,
  validateColor,
  validateHexColor,
  normalizeColor,
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
export {
  detectCorruptionInDocument,
  detectCorruptionInText,
  suggestFix,
  looksCorrupted,
  CorruptionReport,
  CorruptionLocation,
  CorruptionType,
} from './utils/corruptionDetection';
export { isError, toError, wrapError, getErrorMessage } from './utils/errorHandling';
export {
  REVISION_RULES,
  ValidationSeverity,
  ValidationIssue,
  ValidationRule,
  ValidationOptions,
  AutoFixOptions,
  ValidationResult as RevisionValidationResult,
  FixAction,
  AutoFixResult,
  createIssueFromRule,
  getRuleByCode,
  getRulesBySeverity,
  getAutoFixableRules,
  RevisionValidator,
  RevisionAutoFixer,
} from './validation';

// =============================================================================
// UTILITIES — Formatting, Parsing, Sanitization
// =============================================================================

export {
  mergeFormatting,
  cloneFormatting,
  hasFormatting,
  cleanFormatting,
  isEqualFormatting,
  applyDefaults,
} from './utils/formatting';
export {
  safeParseInt,
  parseOoxmlBoolean,
  isExplicitlySet,
  parseNumericAttribute,
  parseOnOffAttribute,
} from './utils/parsingHelpers';
export {
  removeInvalidXmlChars,
  findInvalidXmlChars,
  hasInvalidXmlChars,
  XML_CONTROL_CHARS,
} from './utils/xmlSanitization';

// =============================================================================
// UTILITIES — List Detection (kept for basic detection; normalization moved to consumer)
// =============================================================================

export {
  detectTypedPrefix,
  detectListType,
  inferLevelFromIndentation,
  getParagraphIndentation,
  validateListSequence,
  getListCategoryFromFormat,
  getLevelFromFormat,
  TYPED_LIST_PATTERNS,
  PATTERN_TO_CATEGORY,
  FORMAT_TO_LEVEL,
} from './utils/list-detection';

// =============================================================================
// UTILITIES — Logging
// =============================================================================

export {
  ILogger,
  LogLevel,
  LogEntry,
  ConsoleLogger,
  SilentLogger,
  CollectingLogger,
  defaultLogger,
  createScopedLogger,
  createComponentLogger,
  getGlobalLogger,
  setGlobalLogger,
  resetGlobalLogger,
} from './utils/logger';

// =============================================================================
// UTILITIES — Revision Walker
// =============================================================================

export { RevisionWalker, RevisionWalkerOptions } from './utils/RevisionWalker';
export { resolveCellShading } from './utils/ShadingResolver';
export { decodeCnfStyle, getActiveConditionalsInPriorityOrder } from './utils/cnfStyleDecoder';

// =============================================================================
// UTILITIES — Cleanup
// =============================================================================

export {
  CleanupHelper,
  CleanupOptions,
  CleanupReport,
} from './helpers/CleanupHelper';

// =============================================================================
// INTERNAL — ZIP Handling (advanced usage)
// =============================================================================

export { ZipHandler } from './zip/ZipHandler';
export { ZipReader } from './zip/ZipReader';
export { ZipWriter } from './zip/ZipWriter';
export {
  ZipFile,
  FileMap,
  LoadOptions,
  SaveOptions,
  AddFileOptions,
  SizeLimitOptions,
  DEFAULT_SIZE_LIMITS,
  REQUIRED_DOCX_FILES,
  DOCX_PATHS,
} from './zip/types';
export {
  DocxError,
  DocxNotFoundError,
  InvalidDocxError,
  CorruptedArchiveError,
  MissingRequiredFileError,
  FileOperationError,
} from './zip/errors';

// =============================================================================
// INTERNAL — XML Builder and Parser (advanced usage)
// =============================================================================

export { XMLBuilder, XMLElement } from './xml/XMLBuilder';
export {
  XMLParser,
  ParseToObjectOptions,
  ParsedXMLValue,
  ParsedXMLObject,
  DEFAULT_MAX_NESTING_DEPTH,
} from './xml/XMLParser';

// =============================================================================
// INTERNAL — Parser, Generator, Validator (advanced usage)
// =============================================================================

export { Relationship, RelationshipType, RelationshipProperties } from './core/Relationship';
export { RelationshipManager } from './core/RelationshipManager';
export { DocumentParser, ParseError } from './core/DocumentParser';
export { DocumentGenerator, IZipHandlerReader } from './core/DocumentGenerator';
export { DocumentValidator, SizeEstimate, MemoryOptions } from './core/DocumentValidator';
export { DocumentIdManager } from './core/DocumentIdManager';

// =============================================================================
// INTERNAL — Document Subsystem Classes (advanced usage)
// =============================================================================

export { DocumentContent, BodyElement } from './core/DocumentContent';

// =============================================================================
// INTERNAL — Constants
// =============================================================================

export { LIMITS } from './constants/limits';
