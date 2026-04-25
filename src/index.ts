/**
 * DocXML - DOCX Editing Framework
 * Main entry point
 */

// =============================================================================
// PUBLIC API — Core Document Classes
// =============================================================================

export {
  Document,
  DocumentProperties,
  DocumentOptions,
  DocumentLoadOptions,
  DocumentPart,
} from './core/Document.js';

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
} from './elements/Paragraph.js';
export { Run, RunFormatting, ThemeColorValue, LanguageConfig } from './elements/Run.js';
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
} from './elements/Section.js';
export {
  Table,
  TableAlignment,
  TableLayout,
  TableBorder,
  TableBorders,
  TableFormatting,
} from './elements/Table.js';
export { TableRow, RowFormatting } from './elements/TableRow.js';
export {
  TableCell,
  CellBorder,
  CellBorders,
  CellShading,
  CellFormatting,
} from './elements/TableCell.js';
export {
  TableGridChange,
  GridColumn,
  TableGridChangeProperties,
} from './elements/TableGridChange.js';
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
} from './elements/Image.js';
export { ImageRun } from './elements/ImageRun.js';
export { Shape, ShapeType, ShapeProperties, ShapeFill, ShapeOutline } from './elements/Shape.js';
export { TextBox, TextBoxProperties, TextBoxFill, TextBoxMargins } from './elements/TextBox.js';
export { Header, HeaderType, HeaderProperties } from './elements/Header.js';
export { Footer, FooterType, FooterProperties } from './elements/Footer.js';
export { Hyperlink, HyperlinkProperties } from './elements/Hyperlink.js';
export { Bookmark, BookmarkProperties } from './elements/Bookmark.js';
export { RangeMarker, RangeMarkerType, RangeMarkerProperties } from './elements/RangeMarker.js';
export { Comment, CommentProperties } from './elements/Comment.js';
export { Footnote, FootnoteType, FootnoteProperties } from './elements/Footnote.js';
export { Endnote, EndnoteType, EndnoteProperties } from './elements/Endnote.js';
export {
  Field,
  FieldType,
  FieldProperties,
  ComplexField,
  ComplexFieldProperties,
  FieldCharType,
  TOCFieldOptions,
  createTOCField,
} from './elements/Field.js';
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
} from './elements/FieldHelpers.js';
export {
  StructuredDocumentTag,
  SDTProperties,
  SDTLockType,
  SDTContent,
  SDTPlaceholder,
  SDTDataBinding,
  ContentControlType,
} from './elements/StructuredDocumentTag.js';
export { TableOfContents, TOCProperties } from './elements/TableOfContents.js';
export { TableOfContentsElement } from './elements/TableOfContentsElement.js';
export { AlternateContent } from './elements/AlternateContent.js';
export { MathParagraph, MathExpression } from './elements/MathElement.js';
export { CustomXmlBlock } from './elements/CustomXml.js';
export { PreservedElement, PreservedElementContext } from './elements/PreservedElement.js';

// =============================================================================
// PUBLIC API — Track Changes / Revisions
// =============================================================================

export { Revision, RevisionType, RevisionProperties, FieldContext } from './elements/Revision.js';
export { RevisionContent, isRunContent, isHyperlinkContent } from './elements/RevisionContent.js';
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
} from './elements/PropertyChangeTypes.js';
export {
  acceptRevisionsInMemory,
  AcceptRevisionsOptions,
  AcceptRevisionsResult,
  paragraphHasRevisions,
  getRevisionsFromParagraph,
  countRevisionsByType,
  stripRevisionsFromXml,
} from './processors/InMemoryRevisionAcceptor.js';
export {
  MoveOperationHelper,
  MoveOperationOptions,
  MoveOperationResult,
} from './processors/MoveOperationHelper.js';
export {
  SelectiveRevisionAcceptor,
  SelectiveAcceptResult,
} from './processors/SelectiveRevisionAcceptor.js';
export {
  RevisionAwareProcessor,
  RevisionHandlingMode,
  RevisionProcessingOptions,
  SelectionCriteria,
  RevisionProcessingResult,
  ConflictInfo,
  ProcessingLogEntry,
} from './processors/RevisionAwareProcessor.js';
export {
  ChangelogGenerator,
  ChangeEntry,
  ChangeCategory,
  ChangeLocation,
  ChangelogOptions,
  ChangelogFormat,
  ConsolidatedChange,
  ChangelogSummary,
} from './processors/ChangelogGenerator.js';

// =============================================================================
// PUBLIC API — Formatting / Styles / Numbering
// =============================================================================

export { Style, StyleType, StyleProperties } from './formatting/Style.js';
export {
  StylesManager,
  ValidationResult,
  LatentStylesConfig,
  LatentStyleException,
} from './formatting/StylesManager.js';
export {
  NumberingLevel,
  NumberFormat,
  NumberAlignment,
  NumberingLevelProperties,
  WORD_NATIVE_BULLETS,
  WordNativeBullet,
} from './formatting/NumberingLevel.js';
export { AbstractNumbering, AbstractNumberingProperties } from './formatting/AbstractNumbering.js';
export { NumberingInstance, NumberingInstanceProperties } from './formatting/NumberingInstance.js';
export {
  NumberingManager,
  NumberingConsolidationOptions,
  NumberingConsolidationResult,
} from './formatting/NumberingManager.js';
export {
  StyleRunFormatting,
  StyleParagraphFormatting,
  Heading2TableOptions,
  StyleConfig,
  Heading2Config,
  ApplyCustomFormattingOptions,
} from './types/styleConfig.js';
export { FormatOptions, StyleApplyOptions, EmphasisType, ListPrefix } from './types/formatting.js';

// =============================================================================
// PUBLIC API — Managers
// =============================================================================

export { RevisionManager, RevisionCategory, RevisionSummary } from './elements/RevisionManager.js';
export { ImageManager } from './elements/ImageManager.js';
export { BookmarkManager } from './elements/BookmarkManager.js';
export { CommentManager } from './elements/CommentManager.js';
export { FootnoteManager } from './elements/FootnoteManager.js';
export { EndnoteManager } from './elements/EndnoteManager.js';
export { HeaderFooterManager } from './elements/HeaderFooterManager.js';
export { FontManager, FontFormat, FontEntry } from './elements/FontManager.js';
export {
  DrawingManager,
  DrawingElement,
  DrawingType,
  PreservedDrawing,
} from './managers/DrawingManager.js';

// =============================================================================
// PUBLIC API — Image Optimization
// =============================================================================

export type { ImageOptimizationResult } from './images/ImageOptimizer.js';

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
} from './elements/CommonTypes.js';
export {
  ListCategory,
  NumberFormat as ListNumberFormat,
  BulletFormat,
  ListDetectionResult,
  ListAnalysis,
  ListNormalizationOptions,
  ListNormalizationReport,
  IndentationLevel,
} from './types/list-types.js';

export {
  CompatibilityMode,
  CompatibilityInfo,
  CompatSetting,
} from './types/compatibility-types.js';

export {
  DocumentProtection,
  RevisionViewSettings,
  TrackChangesSettings,
  WebSettingsInfo,
} from './types/settings-types.js';

// =============================================================================
// TYPES — Compatibility Upgrade
// =============================================================================

export { CompatibilityUpgrader, UpgradeReport } from './processors/CompatibilityUpgrader.js';
export {
  LEGACY_COMPAT_ELEMENTS,
  LEGACY_COMPAT_ELEMENT_NAMES,
  MODERN_COMPAT_SETTINGS,
  MS_WORD_COMPAT_URI,
} from './constants/legacyCompatFlags.js';

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
} from './utils/units.js';

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
  sanitizeHyperlinkUrl,
  SanitizeHyperlinkUrlResult,
  TextValidationResult,
} from './utils/validation.js';
export {
  detectCorruptionInDocument,
  detectCorruptionInText,
  suggestFix,
  looksCorrupted,
  CorruptionReport,
  CorruptionLocation,
  CorruptionType,
} from './utils/corruptionDetection.js';
export { isError, toError, wrapError, getErrorMessage } from './utils/errorHandling.js';
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
} from './validation/index.js';

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
} from './utils/formatting.js';
export {
  safeParseInt,
  parseOoxmlBoolean,
  isExplicitlySet,
  parseNumericAttribute,
  parseOnOffAttribute,
} from './utils/parsingHelpers.js';
export {
  removeInvalidXmlChars,
  findInvalidXmlChars,
  hasInvalidXmlChars,
  XML_CONTROL_CHARS,
} from './utils/xmlSanitization.js';

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
} from './utils/list-detection.js';

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
} from './utils/logger.js';

// =============================================================================
// UTILITIES — Revision Walker
// =============================================================================

export { RevisionWalker, RevisionWalkerOptions } from './processors/RevisionWalker.js';
export { resolveCellShading } from './processors/ShadingResolver.js';
export {
  decodeCnfStyle,
  getActiveConditionalsInPriorityOrder,
} from './processors/cnfStyleDecoder.js';

// =============================================================================
// UTILITIES — Cleanup
// =============================================================================

export { CleanupHelper, CleanupOptions, CleanupReport } from './helpers/CleanupHelper.js';

// =============================================================================
// PUBLIC — Type unions consumers see when iterating doc.getBodyElements()
// =============================================================================

export type { BodyElement } from './core/DocumentContent.js';
export {
  RelationshipType,
  type Relationship,
  type RelationshipProperties,
} from './core/Relationship.js';

// =============================================================================
// EXTENSIBILITY — Plugin Registries and Document Events
// =============================================================================

export {
  ElementRegistry,
  type ElementHandler,
  type ElementParseContext,
  type ElementSerializeContext,
} from './core/ElementRegistry.js';
export { RegisteredBodyElement } from './elements/RegisteredBodyElement.js';
export {
  ValidationRuleRegistry,
  type CustomValidationRule,
  type CustomValidationIssue,
  type CustomValidationSeverity,
} from './validation/ValidationRuleRegistry.js';
export type {
  DocumentEventMap,
  DocumentEventType,
  DocumentEventListener,
} from './core/DocumentEvents.js';

// Low-level building blocks — XMLBuilder, XMLParser, ZipHandler,
// ZipReader, ZipWriter, DocumentParser, DocumentGenerator,
// DocumentValidator, DocumentIdManager, DocumentContent,
// RelationshipManager, DocumentEventEmitter, plus zip/xml types and
// DocxError hierarchy — moved to the `docxmlater/internal` subpath in
// 11.0.0. Migrate via:
//   import { XMLParser, ZipHandler } from 'docxmlater/internal';
// The internal subpath has relaxed semver — pin an exact version.
