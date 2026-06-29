/**
 * ZipReader - Handles reading ZIP archives (DOCX files)
 */

import JSZip from 'jszip';
import { promises as fs } from 'fs';
import { ZipFile, FileMap, LoadOptions, SizeLimitOptions, DEFAULT_SIZE_LIMITS } from './types.js';
import {
  DocxNotFoundError,
  InvalidDocxError,
  CorruptedArchiveError,
  FileOperationError,
  ResourceLimitError,
} from './errors.js';
import {
  validateDocxStructure,
  isBinaryFile,
  normalizePath,
  isValidZipBuffer,
} from '../utils/validation.js';

/**
 * Shape of JSZip's undocumented per-entry `_data` field, consulted only as an
 * early-reject hint — never as the authoritative size guard (measured decompressed
 * bytes are). Mirrors the commented-out `CompressedObject` interface in jszip's own
 * `node_modules/jszip/index.d.ts`.
 *
 * [TS-5] Verified against jszip 3.10.x (package.json range `^3.10.1`). A minor jszip
 * bump could rename or remove this field; if so the early-reject/ratio hints silently
 * degrade to no-ops while the measured-byte accounting in {@link extractFiles} keeps
 * enforcing the budgets. `tests/zip/JSZipInternalContract.test.ts` canaries the field
 * so such a regression surfaces on upgrade.
 */
interface JSZipObjectPrivate {
  _data?: {
    compressedSize?: number;
    uncompressedSize?: number;
  };
}

/**
 * Handles reading operations on ZIP archives
 */
export class ZipReader {
  private zip: JSZip | null = null;
  private files: FileMap = new Map();
  private loaded = false;

  /**
   * Loads a DOCX file from the filesystem
   * @param filePath - Path to the DOCX file
   * @param options - Load options
   */
  async loadFromFile(filePath: string, options: LoadOptions = {}): Promise<void> {
    try {
      // Check if file exists
      try {
        await fs.access(filePath);
      } catch {
        throw new DocxNotFoundError(filePath);
      }

      // Read file as buffer
      const buffer = await fs.readFile(filePath);
      await this.loadFromBuffer(buffer, options);
    } catch (error: unknown) {
      // Surface load-classification errors verbatim so callers can distinguish a
      // missing/oversized/hostile archive from a generic read failure.
      if (
        error instanceof DocxNotFoundError ||
        error instanceof ResourceLimitError ||
        error instanceof InvalidDocxError ||
        error instanceof CorruptedArchiveError
      ) {
        throw error;
      }
      const message = error instanceof Error ? error.message : String(error);
      throw new FileOperationError('read', message);
    }
  }

  /**
   * Loads a DOCX file from a buffer
   * @param buffer - Buffer containing the DOCX data
   * @param options - Load options
   */
  async loadFromBuffer(buffer: Buffer, options: LoadOptions = {}): Promise<void> {
    const { validate = true } = options;
    const limits: Required<SizeLimitOptions> = { ...DEFAULT_SIZE_LIMITS, ...options.sizeLimits };

    try {
      // Validate ZIP signature
      if (!isValidZipBuffer(buffer)) {
        throw new InvalidDocxError('File is not a valid ZIP archive');
      }

      // Load ZIP archive
      this.zip = await JSZip.loadAsync(buffer);

      // Enumerate non-directory entries once; reused for the count guard below and
      // for extraction, so the file list is not filtered twice.
      const filePaths = Object.keys(this.zip.files).filter((path) => !this.zip!.files[path]!.dir);

      // Reject entry-count amplification before decompressing anything.
      if (limits.maxEntryCount > 0 && filePaths.length > limits.maxEntryCount) {
        throw new ResourceLimitError(
          `archive entry count (${filePaths.length}) exceeds maxEntryCount (${limits.maxEntryCount})`
        );
      }

      // Extract all files (enforces uncompressed/ratio budgets while decompressing)
      await this.extractFiles(limits, filePaths);

      // Validate DOCX structure if requested
      if (validate) {
        this.validate();
      }

      this.loaded = true;
    } catch (error: unknown) {
      // A resource-limit breach must not be masked as a generic corruption error;
      // the cause (which budget was exceeded) is actionable for the caller.
      if (error instanceof InvalidDocxError || error instanceof ResourceLimitError) {
        throw error;
      }
      const message = error instanceof Error ? error.message : String(error);
      throw new CorruptedArchiveError(message);
    }
  }

  /**
   * Extracts all files from the ZIP archive into memory
   *
   * **Encoding Note:**
   * - Known-text parts (.xml, .rels, etc.) are extracted as UTF-8 strings using `async('string')`
   * - JSZip automatically decodes UTF-8 when extracting as 'string'
   * - Everything else is extracted as a Buffer to preserve exact bytes — a
   *   string decode is lossy (invalid UTF-8 becomes U+FFFD), which would
   *   corrupt embedded binary parts (OLE packages, fonts, metafiles)
   * - All text content is guaranteed to be valid UTF-8
   *
   * **Resource limits:** the running total of *actual* decompressed bytes is the
   * primary defense against high-ratio archives. JSZip's internal `_data` metadata
   * is consulted only as an early-reject hint (it is undocumented and may be absent),
   * never as the sole guard — every breach is re-checked against measured bytes.
   *
   * @param limits - Fully-resolved size limits to enforce while extracting
   * @param filePaths - Pre-filtered non-directory entry paths (computed by the caller)
   */
  private async extractFiles(
    limits: Required<SizeLimitOptions>,
    filePaths: string[]
  ): Promise<void> {
    if (!this.zip) {
      throw new Error('ZIP archive not loaded');
    }

    this.files.clear();

    const maxEntryBytes =
      limits.maxEntryUncompressedMB > 0 ? limits.maxEntryUncompressedMB * 1024 * 1024 : 0;
    const maxTotalBytes =
      limits.maxTotalUncompressedMB > 0 ? limits.maxTotalUncompressedMB * 1024 * 1024 : 0;
    // Only enforce the compression ratio on entries large enough for the ratio to be
    // meaningful, so a small but highly-compressible part (e.g. tiny repetitive XML)
    // cannot trip a false positive.
    const ratioFloorBytes = 1024 * 1024;

    let totalBytes = 0;

    // Extract each file
    for (const filePath of filePaths) {
      const normalizedPath = normalizePath(filePath);
      const zipObject = this.zip.files[filePath];

      if (!zipObject) {
        continue;
      }

      // Early-reject hint: JSZip exposes declared uncompressed/compressed sizes on a
      // private `_data` field. Cast once per entry (see {@link JSZipObjectPrivate}) and
      // use it only to avoid decompressing an obviously oversized entry; the
      // authoritative check below uses the measured byte count.
      const internalData = (zipObject as unknown as JSZipObjectPrivate)._data;
      const declaredSize = internalData?.uncompressedSize;
      if (typeof declaredSize === 'number' && declaredSize >= 0) {
        if (maxEntryBytes > 0 && declaredSize > maxEntryBytes) {
          throw new ResourceLimitError(
            `entry "${normalizedPath}" uncompressed size (${declaredSize} bytes) exceeds ` +
              `maxEntryUncompressedMB (${limits.maxEntryUncompressedMB}MB)`
          );
        }
        if (maxTotalBytes > 0 && totalBytes + declaredSize > maxTotalBytes) {
          throw new ResourceLimitError(
            `total uncompressed size would exceed maxTotalUncompressedMB ` +
              `(${limits.maxTotalUncompressedMB}MB)`
          );
        }
      }

      const isBinary = isBinaryFile(normalizedPath);

      // Extract content based on type
      // For known-text files: JSZip's async('string') automatically uses UTF-8 decoding
      // For everything else: async('nodebuffer') preserves exact bytes
      let content;
      if (isBinary) {
        content = await zipObject.async('nodebuffer');
      } else {
        // Known-text files are extracted as UTF-8 strings
        // JSZip automatically handles UTF-8 decoding for 'string' type
        content = await zipObject.async('string');
      }

      // Measure the *actual* decompressed byte count (UTF-8 for strings) — this is the
      // primary, authoritative resource accounting that does not trust archive metadata.
      const entryBytes = Buffer.isBuffer(content)
        ? content.length
        : Buffer.byteLength(content, 'utf8');

      if (maxEntryBytes > 0 && entryBytes > maxEntryBytes) {
        throw new ResourceLimitError(
          `entry "${normalizedPath}" uncompressed size (${entryBytes} bytes) exceeds ` +
            `maxEntryUncompressedMB (${limits.maxEntryUncompressedMB}MB)`
        );
      }

      totalBytes += entryBytes;
      if (maxTotalBytes > 0 && totalBytes > maxTotalBytes) {
        throw new ResourceLimitError(
          `total uncompressed size (${totalBytes} bytes) exceeds maxTotalUncompressedMB ` +
            `(${limits.maxTotalUncompressedMB}MB)`
        );
      }

      // Compression-ratio guard for sizable entries, when the archive reports a
      // compressed size we can divide by (reusing the single `_data` read above).
      const compressedSize = internalData?.compressedSize;
      if (
        limits.maxCompressionRatio > 0 &&
        entryBytes > ratioFloorBytes &&
        typeof compressedSize === 'number' &&
        compressedSize > 0
      ) {
        const ratio = entryBytes / compressedSize;
        if (ratio > limits.maxCompressionRatio) {
          throw new ResourceLimitError(
            `entry "${normalizedPath}" compression ratio (${ratio.toFixed(1)}:1) exceeds ` +
              `maxCompressionRatio (${limits.maxCompressionRatio}:1)`
          );
        }
      }

      // Get file metadata
      const date = zipObject.date;

      // Store file information
      this.files.set(normalizedPath, {
        path: normalizedPath,
        content,
        isBinary,
        size: isBinary ? (content as Buffer).length : (content as string).length,
        date,
      });
    }
  }

  /**
   * Validates the DOCX structure
   * @throws {MissingRequiredFileError} If required files are missing
   */
  private validate(): void {
    const filePaths = Array.from(this.files.keys());
    validateDocxStructure(filePaths);
  }

  /**
   * Gets a specific file from the archive
   * @param filePath - Path to the file within the archive
   * @returns The file data, or undefined if not found
   */
  getFile(filePath: string): ZipFile | undefined {
    this.ensureLoaded();
    const normalizedPath = normalizePath(filePath);
    return this.files.get(normalizedPath);
  }

  /**
   * Gets the content of a specific file as a string
   * @param filePath - Path to the file within the archive
   * @returns The file content as a UTF-8 string, or undefined if not found
   *
   * **Encoding Note:**
   * - Returns UTF-8 decoded string content
   * - For binary files, converts the Buffer to UTF-8 string
   * - Assumes all text content is UTF-8 encoded (per OpenXML standard)
   */
  getFileAsString(filePath: string): string | undefined {
    const file = this.getFile(filePath);
    if (!file) {
      return undefined;
    }

    // Check actual content type instead of flag (Issue #4)
    // Content is Buffer for binary files, string for text files
    if (Buffer.isBuffer(file.content)) {
      // Convert binary buffer to UTF-8 string
      return file.content.toString('utf8');
    }

    return file.content;
  }

  /**
   * Gets the content of a specific file as a buffer
   * @param filePath - Path to the file within the archive
   * @returns The file content as a Buffer, or undefined if not found
   *
   * **Encoding Note:**
   * - Returns Buffer with UTF-8 encoded content for text files
   * - For binary files, returns raw bytes
   * - String content is explicitly encoded as UTF-8
   */
  getFileAsBuffer(filePath: string): Buffer | undefined {
    const file = this.getFile(filePath);
    if (!file) {
      return undefined;
    }

    // Check actual content type instead of flag (Issue #4)
    // Content is Buffer for binary files, string for text files
    if (Buffer.isBuffer(file.content)) {
      return file.content;
    }

    // Encode string content as UTF-8 Buffer
    return Buffer.from(file.content, 'utf8');
  }

  /**
   * Gets all files from the archive
   * @returns Map of file paths to file data
   */
  getAllFiles(): FileMap {
    this.ensureLoaded();
    return new Map(this.files);
  }

  /**
   * Gets a list of all file paths in the archive
   * @returns Array of file paths
   */
  getFilePaths(): string[] {
    this.ensureLoaded();
    return Array.from(this.files.keys());
  }

  /**
   * Checks if a file exists in the archive
   * @param filePath - Path to check
   * @returns True if the file exists
   */
  hasFile(filePath: string): boolean {
    this.ensureLoaded();
    const normalizedPath = normalizePath(filePath);
    return this.files.has(normalizedPath);
  }

  /**
   * Gets files matching a pattern (simple glob)
   * @param pattern - Pattern to match (supports * wildcard)
   * @returns Array of matching files
   */
  getFilesByPattern(pattern: string): ZipFile[] {
    this.ensureLoaded();

    // Convert simple glob pattern to regex
    const regexPattern = pattern.replace(/\*/g, '.*').replace(/\?/g, '.');
    const regex = new RegExp(`^${regexPattern}$`);

    const matchingFiles: ZipFile[] = [];
    for (const [path, file] of this.files) {
      if (regex.test(path)) {
        matchingFiles.push(file);
      }
    }

    return matchingFiles;
  }

  /**
   * Ensures the archive is loaded before operations
   * @throws {Error} If archive is not loaded
   */
  private ensureLoaded(): void {
    if (!this.loaded) {
      throw new Error('Archive not loaded. Call loadFromFile() or loadFromBuffer() first.');
    }
  }

  /**
   * Checks if the archive is loaded
   * @returns True if loaded
   */
  isLoaded(): boolean {
    return this.loaded;
  }

  /**
   * Clears all loaded data
   */
  clear(): void {
    this.zip = null;
    this.files.clear();
    this.loaded = false;
  }
}
