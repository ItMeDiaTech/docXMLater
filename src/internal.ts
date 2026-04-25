/**
 * docxmlater INTERNAL API
 *
 * Subpath export consumed via `import { ... } from 'docxmlater/internal'`.
 * Surfaces low-level building blocks (XMLBuilder, XMLParser, ZipHandler,
 * DocumentParser, DocumentGenerator, RelationshipManager,
 * DocumentValidator, DocumentIdManager) plus the plugin registries
 * (ElementRegistry, ValidationRuleRegistry).
 *
 * **Stability:** these exports follow a relaxed semver — minor version
 * bumps may change signatures. Consumers building plugins or tooling
 * should pin to an exact docxmlater version. The main API
 * (`from 'docxmlater'`) follows full semver.
 *
 * The same identifiers are also re-exported from `src/index.ts` for
 * backward compatibility, so existing call sites are not affected.
 */

// ZIP layer
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

// XML layer
export { XMLBuilder, XMLElement } from './xml/XMLBuilder';
export {
  XMLParser,
  ParseToObjectOptions,
  ParsedXMLValue,
  ParsedXMLObject,
  DEFAULT_MAX_NESTING_DEPTH,
} from './xml/XMLParser';

// Core orchestration
export { Relationship, RelationshipType, RelationshipProperties } from './core/Relationship';
export { RelationshipManager } from './core/RelationshipManager';
export { DocumentParser, ParseError } from './core/DocumentParser';
export { DocumentGenerator, IZipHandlerReader } from './core/DocumentGenerator';
export { DocumentValidator, SizeEstimate, MemoryOptions } from './core/DocumentValidator';
export { DocumentIdManager } from './core/DocumentIdManager';
export { DocumentContent, BodyElement } from './core/DocumentContent';

// Plugin registries — extension points for custom elements / validation
export {
  ElementRegistry,
  ElementHandler,
  ElementParseContext,
  ElementSerializeContext,
} from './core/ElementRegistry';
export {
  ValidationRuleRegistry,
  CustomValidationRule,
  CustomValidationIssue,
  CustomValidationSeverity,
} from './validation/ValidationRuleRegistry';

// Document events
export {
  DocumentEventEmitter,
  DocumentEventMap,
  DocumentEventType,
  DocumentEventListener,
} from './core/DocumentEvents';
