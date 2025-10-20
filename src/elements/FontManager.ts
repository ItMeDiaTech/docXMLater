/**
 * FontManager - Manages embedded fonts in DOCX documents
 * Handles font registration and Content_Types.xml updates
 */

/**
 * Font format types supported in DOCX
 */
export type FontFormat = 'ttf' | 'otf' | 'woff' | 'woff2';

/**
 * Font entry information
 */
export interface FontEntry {
  /** Font file name (e.g., 'arial.ttf') */
  filename: string;
  /** Font format */
  format: FontFormat;
  /** Font family name (e.g., 'Arial') */
  fontFamily: string;
  /** Font data (Buffer) */
  data: Buffer;
  /** Path in DOCX archive (e.g., 'word/fonts/arial.ttf') */
  path: string;
}

/**
 * FontManager handles embedded fonts in documents
 * Ensures fonts are properly registered in [Content_Types].xml
 */
export class FontManager {
  private fonts: Map<string, FontEntry> = new Map();
  private static fontCounter = 1;

  /**
   * Creates a new FontManager
   */
  constructor() {
    // Empty constructor
  }

  /**
   * Factory method to create a new FontManager
   */
  static create(): FontManager {
    return new FontManager();
  }

  /**
   * Adds a font to the document
   * @param fontFamily - Font family name (e.g., 'Arial', 'Times New Roman')
   * @param data - Font file data as Buffer
   * @param format - Font format (ttf, otf, woff, woff2)
   * @returns Font path in the archive
   */
  addFont(fontFamily: string, data: Buffer, format: FontFormat = 'ttf'): string {
    // Generate unique filename
    const sanitizedFamily = this.sanitizeFontName(fontFamily);
    const filename = `${sanitizedFamily}_${FontManager.fontCounter++}.${format}`;
    const path = `word/fonts/${filename}`;

    // Create font entry
    const entry: FontEntry = {
      filename,
      format,
      fontFamily,
      data,
      path,
    };

    // Store with path as key for easy lookup
    this.fonts.set(path, entry);

    return path;
  }

  /**
   * Adds a font from a file path
   * @param fontFamily - Font family name
   * @param filePath - Path to font file
   * @param format - Font format (optional, detected from extension)
   * @returns Font path in the archive
   */
  async addFontFromFile(
    fontFamily: string,
    filePath: string,
    format?: FontFormat
  ): Promise<string> {
    const fs = require('fs').promises;

    // Detect format from extension if not provided
    if (!format) {
      const ext = filePath.split('.').pop()?.toLowerCase();
      if (ext === 'ttf' || ext === 'otf' || ext === 'woff' || ext === 'woff2') {
        format = ext as FontFormat;
      } else {
        throw new Error(`Unable to detect font format from file: ${filePath}`);
      }
    }

    // Read font file
    const data = await fs.readFile(filePath);

    return this.addFont(fontFamily, data, format);
  }

  /**
   * Checks if a font exists
   * @param fontFamily - Font family name
   * @returns True if font is registered
   */
  hasFont(fontFamily: string): boolean {
    for (const entry of this.fonts.values()) {
      if (entry.fontFamily === fontFamily) {
        return true;
      }
    }
    return false;
  }

  /**
   * Gets all registered fonts
   * @returns Array of font entries
   */
  getAllFonts(): FontEntry[] {
    return Array.from(this.fonts.values());
  }

  /**
   * Gets font count
   * @returns Number of fonts
   */
  getCount(): number {
    return this.fonts.size;
  }

  /**
   * Gets MIME type for font format
   * @param format - Font format
   * @returns MIME type string
   */
  static getMimeType(format: FontFormat): string {
    const mimeTypes: Record<FontFormat, string> = {
      ttf: 'application/x-font-ttf',
      otf: 'application/x-font-opentype',
      woff: 'application/font-woff',
      woff2: 'font/woff2',
    };

    return mimeTypes[format] || 'application/octet-stream';
  }

  /**
   * Gets all unique font extensions
   * @returns Set of extensions (e.g., 'ttf', 'otf')
   */
  getExtensions(): Set<string> {
    const extensions = new Set<string>();
    for (const entry of this.fonts.values()) {
      extensions.add(entry.format);
    }
    return extensions;
  }

  /**
   * Generates Content_Types.xml entries for fonts
   * @returns Array of XML strings
   */
  generateContentTypeEntries(): string[] {
    const entries: string[] = [];
    const extensions = this.getExtensions();

    // Add Default entries for each extension
    for (const ext of extensions) {
      const mimeType = FontManager.getMimeType(ext as FontFormat);
      entries.push(`  <Default Extension="${ext}" ContentType="${mimeType}"/>`);
    }

    return entries;
  }

  /**
   * Sanitizes font name for use in filename
   * @param fontName - Font family name
   * @returns Sanitized name safe for filename
   */
  private sanitizeFontName(fontName: string): string {
    // Remove spaces and special characters
    return fontName
      .toLowerCase()
      .replace(/\s+/g, '_')
      .replace(/[^a-z0-9_-]/g, '')
      .substring(0, 50); // Limit length
  }

  /**
   * Clears all fonts
   */
  clear(): void {
    this.fonts.clear();
  }

  /**
   * Removes a specific font
   * @param path - Font path in archive
   * @returns True if font was removed
   */
  removeFont(path: string): boolean {
    return this.fonts.delete(path);
  }
}
