/**
 * Date formatting utilities for OOXML XML generation
 *
 * ECMA-376 requires dates in ISO 8601 format WITHOUT milliseconds:
 *   Valid:   "2024-01-01T12:00:00Z"
 *   Invalid: "2024-01-01T12:00:00.000Z" (Word rejects milliseconds in w:date attributes)
 *
 * JavaScript's Date.toISOString() always includes milliseconds (YYYY-MM-DDTHH:mm:ss.sssZ),
 * so we must strip the .sss portion for all w:date attributes in tracked changes.
 */

/**
 * Formats a Date to ISO 8601 without milliseconds for OOXML w:date attributes.
 *
 * @param date - Date to format
 * @returns ISO 8601 date string without milliseconds (e.g., "2024-01-01T12:00:00Z")
 */
export function formatDateForXml(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}
