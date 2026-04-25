/**
 * Document-shaped types extracted out of `src/core/Document.ts` so that
 * `DocumentParser`, `DocumentGenerator`, and `DocumentValidator` can
 * import them without re-introducing a circular dependency back to the
 * Document class.
 *
 * `Document.ts` re-exports these for backward compatibility — existing
 * `import { DocumentProperties } from 'docxmlater'` continues to work.
 */

/** Document properties (core and extended). */
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
 * Document part representation. Any part inside the DOCX package
 * (XML, binary).
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
