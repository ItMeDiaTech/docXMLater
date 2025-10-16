/**
 * ZipHandler - Main facade for ZIP archive operations
 * Provides a unified interface for reading and writing DOCX files
 */

import { ZipReader } from './ZipReader';
import { ZipWriter } from './ZipWriter';
import {
  ZipFile,
  FileMap,
  LoadOptions,
  SaveOptions,
  AddFileOptions,
} from './types';

/**
 * Main class for handling ZIP archives (DOCX files)
 * Combines reading and writing operations into a single interface
 */
export class ZipHandler {
  private reader: ZipReader;
  private writer: ZipWriter;
  private mode: 'read' | 'write' | 'modify' = 'write';

  constructor() {
    this.reader = new ZipReader();
    this.writer = new ZipWriter();
  }

  // ==================== LOADING ====================

  /**
   * Loads a DOCX file from the filesystem
   * @param filePath - Path to the DOCX file
   * @param options - Load options
   */
  async load(filePath: string, options: LoadOptions = {}): Promise<void> {
    // Check file size before loading
    const { promises: fs } = await import('fs');
    const stats = await fs.stat(filePath);
    const sizeMB = stats.size / (1024 * 1024);

    // Warn on large files
    const WARNING_SIZE_MB = 50;
    const ERROR_SIZE_MB = 150;

    if (sizeMB > ERROR_SIZE_MB) {
      throw new Error(
        `Document size (${sizeMB.toFixed(1)}MB) exceeds maximum supported size (${ERROR_SIZE_MB}MB). ` +
        `This would likely cause out-of-memory errors. Consider:\n` +
        `- Compressing/optimizing images\n` +
        `- Splitting into multiple documents\n` +
        `- Processing on a machine with more memory`
      );
    } else if (sizeMB > WARNING_SIZE_MB) {
      console.warn(
        `DocXML Warning: Large document detected (${sizeMB.toFixed(1)}MB). ` +
        `Loading may use significant memory. Consider optimizing document size.`
      );
    }

    await this.reader.loadFromFile(filePath, options);

    // Copy all files from reader to writer for modification
    const files = this.reader.getAllFiles();
    this.writer.clear();
    this.writer.addFiles(files);

    this.mode = 'modify';
  }

  /**
   * Loads a DOCX file from a buffer
   * @param buffer - Buffer containing the DOCX data
   * @param options - Load options
   */
  async loadFromBuffer(buffer: Buffer, options: LoadOptions = {}): Promise<void> {
    // Check buffer size before loading
    const sizeMB = buffer.length / (1024 * 1024);

    // Warn on large files
    const WARNING_SIZE_MB = 50;
    const ERROR_SIZE_MB = 150;

    if (sizeMB > ERROR_SIZE_MB) {
      throw new Error(
        `Document size (${sizeMB.toFixed(1)}MB) exceeds maximum supported size (${ERROR_SIZE_MB}MB). ` +
        `This would likely cause out-of-memory errors. Consider:\n` +
        `- Compressing/optimizing images\n` +
        `- Splitting into multiple documents\n` +
        `- Processing on a machine with more memory`
      );
    } else if (sizeMB > WARNING_SIZE_MB) {
      console.warn(
        `DocXML Warning: Large document detected (${sizeMB.toFixed(1)}MB). ` +
        `Loading may use significant memory. Consider optimizing document size.`
      );
    }

    await this.reader.loadFromBuffer(buffer, options);

    // Copy all files from reader to writer for modification
    const files = this.reader.getAllFiles();
    this.writer.clear();
    this.writer.addFiles(files);

    this.mode = 'modify';
  }

  // ==================== FILE OPERATIONS ====================

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
    this.writer.addFile(filePath, content, options);
  }

  /**
   * Adds multiple files to the archive
   * @param files - Map of file paths to contents
   * @param options - Options for adding files
   */
  addFiles(files: FileMap, options: AddFileOptions = {}): void {
    this.writer.addFiles(files, options);
  }

  /**
   * Updates an existing file in the archive
   * @param filePath - Path to the file to update
   * @param content - New content
   * @param options - Options for updating the file
   * @returns True if the file was updated, false if it didn't exist
   */
  updateFile(
    filePath: string,
    content: string | Buffer,
    options: AddFileOptions = {}
  ): boolean {
    if (!this.hasFile(filePath)) {
      return false;
    }
    this.addFile(filePath, content, options);
    return true;
  }

  /**
   * Removes a file from the archive
   * @param filePath - Path to the file to remove
   * @returns True if the file was removed, false if it didn't exist
   */
  removeFile(filePath: string): boolean {
    return this.writer.removeFile(filePath);
  }

  /**
   * Gets a specific file from the archive
   * @param filePath - Path to the file within the archive
   * @returns The file data, or undefined if not found
   */
  getFile(filePath: string): ZipFile | undefined {
    return this.writer.getFile(filePath);
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
    return this.writer.getAllFiles();
  }

  /**
   * Gets a list of all file paths in the archive
   * @returns Array of file paths
   */
  getFilePaths(): string[] {
    return this.writer.getFilePaths();
  }

  /**
   * Checks if a file exists in the archive
   * @param filePath - Path to check
   * @returns True if the file exists
   */
  hasFile(filePath: string): boolean {
    return this.writer.hasFile(filePath);
  }

  /**
   * Gets the number of files in the archive
   * @returns Number of files
   */
  getFileCount(): number {
    return this.writer.getFileCount();
  }

  // ==================== HELPER METHODS ====================

  /**
   * Renames a file in the archive
   * @param oldPath - Current path of the file
   * @param newPath - New path for the file
   * @returns True if the file was renamed, false if it didn't exist
   */
  renameFile(oldPath: string, newPath: string): boolean {
    const file = this.getFile(oldPath);
    if (!file) {
      return false;
    }
    this.addFile(newPath, file.content, {
      binary: file.isBinary,
      date: file.date,
    });
    this.removeFile(oldPath);
    return true;
  }

  /**
   * Copies a file within the archive
   * @param srcPath - Source file path
   * @param destPath - Destination file path
   * @returns True if the file was copied, false if source didn't exist
   */
  copyFile(srcPath: string, destPath: string): boolean {
    const file = this.getFile(srcPath);
    if (!file) {
      return false;
    }
    this.addFile(destPath, file.content, {
      binary: file.isBinary,
      date: file.date,
    });
    return true;
  }

  /**
   * Moves a file within the archive (copy and delete)
   * @param srcPath - Source file path
   * @param destPath - Destination file path
   * @returns True if the file was moved, false if source didn't exist
   */
  moveFile(srcPath: string, destPath: string): boolean {
    if (!this.copyFile(srcPath, destPath)) {
      return false;
    }
    this.removeFile(srcPath);
    return true;
  }

  /**
   * Checks if a file exists, throws if it doesn't
   * @param filePath - Path to check
   * @throws {Error} If file doesn't exist
   */
  existsOrThrow(filePath: string): void {
    if (!this.hasFile(filePath)) {
      throw new Error(`File not found in archive: ${filePath}`);
    }
  }

  /**
   * Removes multiple files from the archive
   * @param filePaths - Array of file paths to remove
   * @returns Number of files successfully removed
   */
  removeFiles(filePaths: string[]): number {
    let count = 0;
    for (const filePath of filePaths) {
      if (this.removeFile(filePath)) {
        count++;
      }
    }
    return count;
  }

  /**
   * Gets all files with a specific extension
   * @param extension - File extension (with or without leading dot)
   * @returns Array of files with the specified extension
   */
  getFilesByExtension(extension: string): ZipFile[] {
    const ext = extension.startsWith('.') ? extension : `.${extension}`;
    const files: ZipFile[] = [];
    for (const [path, file] of this.getAllFiles()) {
      if (path.toLowerCase().endsWith(ext.toLowerCase())) {
        files.push(file);
      }
    }
    return files;
  }

  /**
   * Gets the total uncompressed size of all files in the archive
   * @returns Total size in bytes
   */
  getTotalSize(): number {
    let totalSize = 0;
    for (const file of this.getAllFiles().values()) {
      totalSize += file.size;
    }
    return totalSize;
  }

  /**
   * Gets comprehensive statistics about the archive
   * @returns Statistics object
   */
  getStats(): {
    fileCount: number;
    totalSize: number;
    textFileCount: number;
    binaryFileCount: number;
    avgFileSize: number;
  } {
    let textCount = 0;
    let binaryCount = 0;
    const totalSize = this.getTotalSize();
    const fileCount = this.getFileCount();

    for (const file of this.getAllFiles().values()) {
      if (file.isBinary) {
        binaryCount++;
      } else {
        textCount++;
      }
    }

    return {
      fileCount,
      totalSize,
      textFileCount: textCount,
      binaryFileCount: binaryCount,
      avgFileSize: fileCount > 0 ? Math.round(totalSize / fileCount) : 0,
    };
  }

  /**
   * Checks if the archive is empty
   * @returns True if the archive has no files
   */
  isEmpty(): boolean {
    return this.getFileCount() === 0;
  }

  /**
   * Gets all text (non-binary) files from the archive
   * @returns Array of text files
   */
  getTextFiles(): ZipFile[] {
    const files: ZipFile[] = [];
    for (const file of this.getAllFiles().values()) {
      if (!file.isBinary) {
        files.push(file);
      }
    }
    return files;
  }

  /**
   * Gets all binary files from the archive
   * @returns Array of binary files
   */
  getBinaryFiles(): ZipFile[] {
    const files: ZipFile[] = [];
    for (const file of this.getAllFiles().values()) {
      if (file.isBinary) {
        files.push(file);
      }
    }
    return files;
  }

  /**
   * Gets all media files from the word/media/ directory
   * @returns Array of media files
   */
  getMediaFiles(): ZipFile[] {
    const files: ZipFile[] = [];
    for (const [path, file] of this.getAllFiles()) {
      if (path.startsWith('word/media/')) {
        files.push(file);
      }
    }
    return files;
  }

  /**
   * Exports a file from the archive to the filesystem
   * @param internalPath - Path of the file within the archive
   * @param outputPath - Path where the file will be saved
   */
  async exportFile(internalPath: string, outputPath: string): Promise<void> {
    const { promises: fs } = await import('fs');
    const content = this.getFileAsBuffer(internalPath);
    if (!content) {
      throw new Error(`File not found in archive: ${internalPath}`);
    }
    await fs.writeFile(outputPath, content);
  }

  /**
   * Imports a file from the filesystem into the archive
   * @param sourcePath - Path to the file on filesystem
   * @param internalPath - Path where the file will be stored in archive
   * @param options - Options for adding the file
   */
  async importFile(
    sourcePath: string,
    internalPath: string,
    options: AddFileOptions = {}
  ): Promise<void> {
    const { promises: fs } = await import('fs');
    const content = await fs.readFile(sourcePath);
    this.addFile(internalPath, content, {
      ...options,
      binary: options.binary !== undefined ? options.binary : true,
    });
  }

  // ==================== SAVING ====================

  /**
   * Saves the archive to a file
   * @param filePath - Path where the file will be saved
   * @param options - Save options
   */
  async save(filePath: string, options: SaveOptions = {}): Promise<void> {
    await this.writer.saveToFile(filePath, options);
  }

  /**
   * Generates the ZIP archive as a buffer
   * @param options - Save options
   * @returns Buffer containing the ZIP archive
   */
  async toBuffer(options: SaveOptions = {}): Promise<Buffer> {
    return await this.writer.toBuffer(options);
  }

  // ==================== VALIDATION ====================

  /**
   * Validates the DOCX structure
   * @throws {MissingRequiredFileError} If required files are missing
   */
  validate(): void {
    this.writer.validate();
  }

  // ==================== UTILITY ====================

  /**
   * Creates a new empty archive
   */
  clear(): void {
    this.reader.clear();
    this.writer.clear();
    this.mode = 'write';
  }

  /**
   * Gets the current mode of the handler
   * @returns Current mode ('read', 'write', or 'modify')
   */
  getMode(): 'read' | 'write' | 'modify' {
    return this.mode;
  }

  /**
   * Checks if the archive has been loaded
   * @returns True if loaded
   */
  isLoaded(): boolean {
    return this.reader.isLoaded();
  }

  /**
   * Creates a clone of this handler with all its files
   * @returns A new ZipHandler instance with the same files
   */
  clone(): ZipHandler {
    const newHandler = new ZipHandler();
    newHandler.writer = this.writer.clone();
    newHandler.mode = this.mode;
    return newHandler;
  }

  // ==================== CONVENIENCE METHODS ====================

  /**
   * Reads a DOCX file, modifies it, and saves it back
   * @param inputPath - Path to the input DOCX file
   * @param outputPath - Path to save the modified file
   * @param modifier - Function that modifies the handler
   * @param loadOptions - Options for loading
   * @param saveOptions - Options for saving
   */
  static async modify(
    inputPath: string,
    outputPath: string,
    modifier: (handler: ZipHandler) => void | Promise<void>,
    loadOptions: LoadOptions = {},
    saveOptions: SaveOptions = {}
  ): Promise<void> {
    const handler = new ZipHandler();
    await handler.load(inputPath, loadOptions);
    await modifier(handler);
    await handler.save(outputPath, saveOptions);
  }

  /**
   * Creates a new DOCX file with the provided files
   * @param outputPath - Path to save the new file
   * @param files - Map of file paths to contents
   * @param saveOptions - Options for saving
   */
  static async create(
    outputPath: string,
    files: FileMap,
    saveOptions: SaveOptions = {}
  ): Promise<void> {
    const handler = new ZipHandler();
    handler.addFiles(files);
    await handler.save(outputPath, saveOptions);
  }
}
