/**
 * ZipReader - Handles reading ZIP archives (DOCX files)
 */

import JSZip from 'jszip';
import { promises as fs } from 'fs';
import { ZipFile, FileMap, LoadOptions } from './types';
import {
  DocxNotFoundError,
  InvalidDocxError,
  CorruptedArchiveError,
  FileOperationError,
} from './errors';
import {
  validateDocxStructure,
  isBinaryFile,
  normalizePath,
  isValidZipBuffer,
} from '../utils/validation';

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
    } catch (error) {
      if (error instanceof DocxNotFoundError) {
        throw error;
      }
      throw new FileOperationError('read', (error as Error).message);
    }
  }

  /**
   * Loads a DOCX file from a buffer
   * @param buffer - Buffer containing the DOCX data
   * @param options - Load options
   */
  async loadFromBuffer(buffer: Buffer, options: LoadOptions = {}): Promise<void> {
    const { validate = true } = options;

    try {
      // Validate ZIP signature
      if (!isValidZipBuffer(buffer)) {
        throw new InvalidDocxError('File is not a valid ZIP archive');
      }

      // Load ZIP archive
      this.zip = await JSZip.loadAsync(buffer);

      // Extract all files
      await this.extractFiles();

      // Validate DOCX structure if requested
      if (validate) {
        this.validate();
      }

      this.loaded = true;
    } catch (error) {
      if (error instanceof InvalidDocxError) {
        throw error;
      }
      throw new CorruptedArchiveError((error as Error).message);
    }
  }

  /**
   * Extracts all files from the ZIP archive into memory
   */
  private async extractFiles(): Promise<void> {
    if (!this.zip) {
      throw new Error('ZIP archive not loaded');
    }

    this.files.clear();

    // Get all file paths
    const filePaths = Object.keys(this.zip.files).filter(
      (path) => !this.zip!.files[path]!.dir
    );

    // Extract each file
    for (const filePath of filePaths) {
      const normalizedPath = normalizePath(filePath);
      const zipObject = this.zip.files[filePath];

      if (!zipObject) {
        continue;
      }

      const isBinary = isBinaryFile(normalizedPath);

      // Extract content based on type
      const content = isBinary
        ? await zipObject.async('nodebuffer')
        : await zipObject.async('string');

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
   * @returns The file content as a string, or undefined if not found
   */
  getFileAsString(filePath: string): string | undefined {
    const file = this.getFile(filePath);
    if (!file) {
      return undefined;
    }

    if (file.isBinary) {
      return (file.content as Buffer).toString('utf8');
    }

    return file.content as string;
  }

  /**
   * Gets the content of a specific file as a buffer
   * @param filePath - Path to the file within the archive
   * @returns The file content as a buffer, or undefined if not found
   */
  getFileAsBuffer(filePath: string): Buffer | undefined {
    const file = this.getFile(filePath);
    if (!file) {
      return undefined;
    }

    if (file.isBinary) {
      return file.content as Buffer;
    }

    return Buffer.from(file.content as string, 'utf8');
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
    const regexPattern = pattern
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.');
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
