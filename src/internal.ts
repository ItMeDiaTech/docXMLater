/**
 * docxmlater INTERNAL API
 *
 * Subpath export consumed via `import { ... } from 'docxmlater/internal'`.
 * Surfaces low-level building blocks for advanced consumers writing
 * tooling that needs raw XML / ZIP / parser access:
 *   - ZipHandler, ZipReader, ZipWriter (+ types and DocxError hierarchy)
 *   - XMLBuilder, XMLParser (+ ParseToObjectOptions etc.)
 *   - DocumentParser, DocumentGenerator, DocumentValidator
 *   - DocumentIdManager, DocumentContent, RelationshipManager
 *   - DocumentEventEmitter (the emitter class itself)
 *   - LIMITS constants
 *
 * **Stability:** these exports follow a relaxed semver — minor version
 * bumps may change signatures. Pin an exact docxmlater version when
 * consuming the internal subpath. The main API (`from 'docxmlater'`)
 * follows full semver.
 *
 * Symbols moved out of the main API in 11.0.0 — see CHANGELOG.
 *
 * The plugin registries (`ElementRegistry`, `ValidationRuleRegistry`)
 * and event types (`DocumentEventMap`, `DocumentEventListener`,
 * `DocumentEventType`) are NOT re-exported here — they belong in the
 * main API as stable extension points.
 */

// ZIP layer
export { ZipHandler } from './zip/ZipHandler.js';
export { ZipReader } from './zip/ZipReader.js';
export { ZipWriter } from './zip/ZipWriter.js';
export {
  type ZipFile,
  type FileMap,
  type LoadOptions,
  type SaveOptions,
  type AddFileOptions,
  type SizeLimitOptions,
  DEFAULT_SIZE_LIMITS,
  REQUIRED_DOCX_FILES,
  DOCX_PATHS,
} from './zip/types.js';
export {
  DocxError,
  DocxNotFoundError,
  InvalidDocxError,
  CorruptedArchiveError,
  MissingRequiredFileError,
  FileOperationError,
} from './zip/errors.js';

// XML layer
export { XMLBuilder, type XMLElement } from './xml/XMLBuilder.js';
export {
  XMLParser,
  type ParseToObjectOptions,
  type ParsedXMLValue,
  type ParsedXMLObject,
  DEFAULT_MAX_NESTING_DEPTH,
} from './xml/XMLParser.js';

// Core orchestration
export { RelationshipManager } from './core/RelationshipManager.js';
export { DocumentParser, ParseError } from './core/DocumentParser.js';
export { DocumentGenerator, type IZipHandlerReader } from './core/DocumentGenerator.js';
export {
  DocumentValidator,
  type SizeEstimate,
  type MemoryOptions,
} from './core/DocumentValidator.js';
export { DocumentIdManager } from './core/DocumentIdManager.js';
export { DocumentContent } from './core/DocumentContent.js';

// Document events — the emitter class is internal; the type aliases
// (DocumentEventMap, DocumentEventListener, DocumentEventType) live in
// the main API.
export { DocumentEventEmitter } from './core/DocumentEvents.js';

// Constants
export { LIMITS } from './constants/limits.js';
