/**
 * Validation utilities for DOCX files
 */

import { REQUIRED_DOCX_FILES } from '../zip/types';
import { MissingRequiredFileError } from '../zip/errors';

/**
 * Validates that all required DOCX files are present
 * @param filePaths - Array of file paths in the archive
 * @throws {MissingRequiredFileError} If a required file is missing
 */
export function validateDocxStructure(filePaths: string[]): void {
  const fileSet = new Set(filePaths);

  for (const requiredFile of REQUIRED_DOCX_FILES) {
    if (!fileSet.has(requiredFile)) {
      throw new MissingRequiredFileError(requiredFile);
    }
  }
}

/**
 * Checks if a file path represents a binary file based on extension
 * @param filePath - The file path to check
 * @returns True if the file is likely binary
 */
export function isBinaryFile(filePath: string): boolean {
  const binaryExtensions = [
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.ico',
    '.emf', '.wmf', '.bin', '.dat', '.ttf', '.otf', '.woff',
  ];

  const extension = filePath.substring(filePath.lastIndexOf('.')).toLowerCase();
  return binaryExtensions.includes(extension);
}

/**
 * Normalizes a file path for consistent comparisons
 * Converts backslashes to forward slashes and removes leading slashes
 * @param path - The path to normalize
 * @returns Normalized path
 */
export function normalizePath(path: string): string {
  return path.replace(/\\/g, '/').replace(/^\/+/, '');
}

/**
 * Validates that a buffer contains a valid ZIP file signature
 * ZIP files start with the signature 'PK' (0x50 0x4B)
 * @param buffer - The buffer to validate
 * @returns True if the buffer appears to be a ZIP file
 */
export function isValidZipBuffer(buffer: Buffer): boolean {
  if (buffer.length < 4) {
    return false;
  }

  // Check for ZIP signature: PK\x03\x04 or PK\x05\x06 (for empty archives)
  return (
    (buffer[0] === 0x50 && buffer[1] === 0x4B) &&
    ((buffer[2] === 0x03 && buffer[3] === 0x04) ||
     (buffer[2] === 0x05 && buffer[3] === 0x06))
  );
}

/**
 * Checks if a string is valid UTF-8 text
 * @param content - The content to check
 * @returns True if the content is valid text
 */
export function isTextContent(content: Buffer | string): boolean {
  if (typeof content === 'string') {
    return true;
  }

  // Try to decode as UTF-8 and check for null bytes
  try {
    const text = content.toString('utf8');
    // Binary files often contain null bytes
    return !text.includes('\0');
  } catch {
    return false;
  }
}

/**
 * Validates a twips value (used for spacing, indentation, margins)
 * Twips: 1/20th of a point, 1440 twips = 1 inch
 * Reasonable range: -31680 to 31680 (±22 inches)
 * @param value - The twips value to validate
 * @param fieldName - Name of the field (for error messages)
 * @throws {Error} If the value is invalid
 */
export function validateTwips(value: number, fieldName: string = 'value'): void {
  if (!Number.isFinite(value)) {
    throw new Error(`${fieldName} must be a finite number, got ${value}`);
  }

  // Reasonable range: ±22 inches (31680 twips)
  const MIN_TWIPS = -31680;
  const MAX_TWIPS = 31680;

  if (value < MIN_TWIPS || value > MAX_TWIPS) {
    throw new Error(
      `${fieldName} out of range: ${value} twips (allowed: ${MIN_TWIPS} to ${MAX_TWIPS}, ±22 inches)`
    );
  }
}

/**
 * Validates a hexadecimal color value
 * Must be 6 characters (RRGGBB format)
 * @param color - The color hex string to validate (without #)
 * @param fieldName - Name of the field (for error messages)
 * @throws {Error} If the color is invalid
 */
export function validateColor(color: string, fieldName: string = 'color'): void {
  if (typeof color !== 'string') {
    throw new Error(`${fieldName} must be a string, got ${typeof color}`);
  }

  // Allow both with and without # prefix
  const cleanColor = color.startsWith('#') ? color.substring(1) : color;

  if (!/^[0-9A-Fa-f]{6}$/.test(cleanColor)) {
    throw new Error(
      `${fieldName} must be a 6-digit hex color (e.g., 'FF0000' or '#FF0000'), got '${color}'`
    );
  }
}

/**
 * Alias for validateColor for backwards compatibility
 */
export const validateHexColor = validateColor;

/**
 * Validates a numbering ID (must be non-negative integer)
 * @param numId - The numbering ID to validate
 * @param fieldName - Name of the field (for error messages)
 * @throws {Error} If the ID is invalid
 */
export function validateNumberingId(numId: number, fieldName: string = 'numbering ID'): void {
  if (!Number.isInteger(numId)) {
    throw new Error(`${fieldName} must be an integer, got ${numId}`);
  }

  if (numId < 0) {
    throw new Error(`${fieldName} must be non-negative, got ${numId}`);
  }

  // Word supports numbering IDs up to 2147483647
  const MAX_NUM_ID = 2147483647;
  if (numId > MAX_NUM_ID) {
    throw new Error(`${fieldName} exceeds maximum value ${MAX_NUM_ID}, got ${numId}`);
  }
}

/**
 * Validates a numbering level (0-8 for Word)
 * @param level - The level to validate
 * @param fieldName - Name of the field (for error messages)
 * @param maxLevel - Maximum allowed level (default 8)
 * @throws {Error} If the level is invalid
 */
export function validateLevel(
  level: number,
  fieldName: string = 'level',
  maxLevel: number = 8
): void {
  if (!Number.isInteger(level)) {
    throw new Error(`${fieldName} must be an integer, got ${level}`);
  }

  if (level < 0 || level > maxLevel) {
    throw new Error(`${fieldName} must be between 0 and ${maxLevel}, got ${level}`);
  }
}

/**
 * Validates an alignment value against allowed values
 * @param alignment - The alignment value to validate
 * @param allowed - Array of allowed alignment values
 * @param fieldName - Name of the field (for error messages)
 * @throws {Error} If the alignment is invalid
 */
export function validateAlignment(
  alignment: string,
  allowed: readonly string[],
  fieldName: string = 'alignment'
): void {
  if (typeof alignment !== 'string') {
    throw new Error(`${fieldName} must be a string, got ${typeof alignment}`);
  }

  if (!allowed.includes(alignment)) {
    throw new Error(
      `Invalid ${fieldName}: '${alignment}' (allowed: ${allowed.join(', ')})`
    );
  }
}

/**
 * Validates a font size (in half-points for Word)
 * Reasonable range: 2-1638 (1-819 points)
 * @param size - The font size in half-points to validate
 * @param fieldName - Name of the field (for error messages)
 * @throws {Error} If the size is invalid
 */
export function validateFontSize(size: number, fieldName: string = 'font size'): void {
  if (!Number.isFinite(size)) {
    throw new Error(`${fieldName} must be a finite number, got ${size}`);
  }

  if (!Number.isInteger(size)) {
    throw new Error(`${fieldName} must be an integer (in half-points), got ${size}`);
  }

  // Reasonable range: 2-1638 half-points (1-819 points)
  const MIN_SIZE = 2;
  const MAX_SIZE = 1638;

  if (size < MIN_SIZE || size > MAX_SIZE) {
    throw new Error(
      `${fieldName} out of range: ${size} half-points (allowed: ${MIN_SIZE}-${MAX_SIZE}, or ${MIN_SIZE / 2}-${MAX_SIZE / 2} points)`
    );
  }
}

/**
 * Validates that a string is not empty
 * @param value - The string to validate
 * @param fieldName - Name of the field (for error messages)
 * @throws {Error} If the string is empty or not a string
 */
export function validateNonEmptyString(value: string, fieldName: string = 'value'): void {
  if (typeof value !== 'string') {
    throw new Error(`${fieldName} must be a string, got ${typeof value}`);
  }

  if (value.trim().length === 0) {
    throw new Error(`${fieldName} cannot be empty`);
  }
}

/**
 * Validates a percentage value (0-100)
 * @param value - The percentage to validate
 * @param fieldName - Name of the field (for error messages)
 * @throws {Error} If the percentage is invalid
 */
export function validatePercentage(value: number, fieldName: string = 'percentage'): void {
  if (!Number.isFinite(value)) {
    throw new Error(`${fieldName} must be a finite number, got ${value}`);
  }

  if (value < 0 || value > 100) {
    throw new Error(`${fieldName} must be between 0 and 100, got ${value}`);
  }
}

/**
 * Validates EMUs (English Metric Units) value
 * Used for image dimensions: 914400 EMUs = 1 inch
 * Reasonable range: 0 to 50 million (about 55 inches)
 * @param value - The EMUs value to validate
 * @param fieldName - Name of the field (for error messages)
 * @throws {Error} If the value is invalid
 */
export function validateEmus(value: number, fieldName: string = 'EMUs'): void {
  if (!Number.isFinite(value)) {
    throw new Error(`${fieldName} must be a finite number, got ${value}`);
  }

  if (!Number.isInteger(value)) {
    throw new Error(`${fieldName} must be an integer, got ${value}`);
  }

  if (value < 0) {
    throw new Error(`${fieldName} must be non-negative, got ${value}`);
  }

  // Reasonable maximum: 50 million EMUs (about 55 inches)
  const MAX_EMUS = 50000000;
  if (value > MAX_EMUS) {
    throw new Error(
      `${fieldName} exceeds maximum ${MAX_EMUS} (about 55 inches), got ${value}`
    );
  }
}
