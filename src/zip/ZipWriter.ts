/**
 * ZipWriter - Handles writing ZIP archives (DOCX files)
 */

import JSZip from 'jszip';
import { promises as fs } from 'fs';
import { ZipFile, FileMap, SaveOptions, AddFileOptions } from './types';
import {
  FileOperationError,
} from './errors';
import {
  validateDocxStructure,
  normalizePath,
} from '../utils/validation';

/**
 * Handles writing operations on ZIP archives
 */
export class ZipWriter {
  private zip: JSZip;
  private files: FileMap = new Map();

  constructor() {
    this.zip = new JSZip();
  }

  /**
   * Adds a file to the archive
   * @param filePath - Path where the file will be stored in the archive
   * @param content - File content (string or Buffer)
   * @param options - Options for adding the file
   */
  addFile(
    filePath: string,
    content: string | Buffer,
    options: AddFileOptions = {}
  ): void {
    const {
      binary = Buffer.isBuffer(content),
      compression = 6,
      date = new Date(),
    } = options;

    const normalizedPath = normalizePath(filePath);

    // Add to JSZip
    this.zip.file(normalizedPath, content, {
      binary,
      compression: compression > 0 ? 'DEFLATE' : 'STORE',
      compressionOptions: {
        level: compression,
      },
      date,
    });

    // Store in our file map
    this.files.set(normalizedPath, {
      path: normalizedPath,
      content,
      isBinary: binary,
      size: Buffer.isBuffer(content) ? content.length : content.length,
      date,
    });
  }

  /**
   * Adds multiple files to the archive
   * @param files - Map of file paths to contents
   * @param options - Options for adding files
   */
  addFiles(files: FileMap, options: AddFileOptions = {}): void {
    for (const [path, file] of files) {
      this.addFile(path, file.content, {
        ...options,
        binary: file.isBinary,
        date: file.date,
      });
    }
  }

  /**
   * Removes a file from the archive
   * @param filePath - Path to the file to remove
   * @returns True if the file was removed, false if it didn't exist
   */
  removeFile(filePath: string): boolean {
    const normalizedPath = normalizePath(filePath);

    // Remove from JSZip
    const zipFile = this.zip.file(normalizedPath);
    if (zipFile) {
      this.zip.remove(normalizedPath);
      this.files.delete(normalizedPath);
      return true;
    }

    return false;
  }

  /**
   * Checks if a file exists in the archive
   * @param filePath - Path to check
   * @returns True if the file exists
   */
  hasFile(filePath: string): boolean {
    const normalizedPath = normalizePath(filePath);
    return this.files.has(normalizedPath);
  }

  /**
   * Gets a file from the archive
   * @param filePath - Path to the file
   * @returns The file data, or undefined if not found
   */
  getFile(filePath: string): ZipFile | undefined {
    const normalizedPath = normalizePath(filePath);
    return this.files.get(normalizedPath);
  }

  /**
   * Gets all files in the archive
   * @returns Map of file paths to file data
   */
  getAllFiles(): FileMap {
    return new Map(this.files);
  }

  /**
   * Gets a list of all file paths in the archive
   * @returns Array of file paths
   */
  getFilePaths(): string[] {
    return Array.from(this.files.keys());
  }

  /**
   * Validates the DOCX structure before saving
   * @throws {MissingRequiredFileError} If required files are missing
   */
  validate(): void {
    const filePaths = this.getFilePaths();
    validateDocxStructure(filePaths);
  }

  /**
   * Generates the ZIP archive as a buffer
   * @param options - Save options
   * @returns Buffer containing the ZIP archive
   */
  async toBuffer(options: SaveOptions = {}): Promise<Buffer> {
    const { compression = 6, validate = true } = options;

    // Validate structure if requested
    if (validate) {
      this.validate();
    }

    try {
      // Generate ZIP with specified compression
      const buffer = await this.zip.generateAsync({
        type: 'nodebuffer',
        compression: compression > 0 ? 'DEFLATE' : 'STORE',
        compressionOptions: {
          level: compression,
        },
        // Use ZIP64 for large files
        streamFiles: true,
      });

      return buffer;
    } catch (error) {
      throw new FileOperationError('generate', (error as Error).message);
    }
  }

  /**
   * Saves the archive to a file
   * @param filePath - Path where the file will be saved
   * @param options - Save options
   */
  async saveToFile(filePath: string, options: SaveOptions = {}): Promise<void> {
    try {
      const buffer = await this.toBuffer(options);
      await fs.writeFile(filePath, buffer);
    } catch (error) {
      if (error instanceof FileOperationError) {
        throw error;
      }
      throw new FileOperationError('save', (error as Error).message);
    }
  }

  /**
   * Creates a new empty archive
   */
  clear(): void {
    this.zip = new JSZip();
    this.files.clear();
  }

  /**
   * Gets the number of files in the archive
   * @returns Number of files
   */
  getFileCount(): number {
    return this.files.size;
  }

  /**
   * Creates a clone of this writer with all its files
   * @returns A new ZipWriter instance with the same files
   */
  clone(): ZipWriter {
    const newWriter = new ZipWriter();
    newWriter.addFiles(this.files);
    return newWriter;
  }
}
