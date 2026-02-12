/**
 * Text diff utility for character-level granular tracked changes.
 *
 * Uses a prefix/suffix diff algorithm to compute minimal edit operations
 * between two strings. This allows tracked changes to show only the
 * actual differences instead of marking entire runs as deleted/inserted.
 */

/**
 * Represents a segment of a diff result.
 */
export interface DiffSegment {
  /** Whether this segment is unchanged, deleted, or inserted */
  type: "equal" | "delete" | "insert";
  /** The text content of this segment */
  text: string;
}

/**
 * Computes minimal diff segments between two strings.
 *
 * Algorithm: Find common prefix, then common suffix from the remaining text,
 * then the middle portion is a delete (from old) + insert (from new).
 *
 * This handles the most common edit patterns optimally:
 * - Space removal: "word  word" → "word word" → [equal "word ", delete " ", equal "word"]
 * - Word replacement: "The quick fox" → "The slow fox" → [equal "The ", delete "quick", insert "slow", equal " fox"]
 * - Prefix change: "Hello World" → "Goodbye World" → [delete "Hello", insert "Goodbye", equal " World"]
 * - Suffix change: "Hello World" → "Hello Earth" → [equal "Hello ", delete "World", insert "Earth"]
 *
 * @param oldText - The original text
 * @param newText - The new text
 * @returns Array of diff segments
 */
export function diffText(oldText: string, newText: string): DiffSegment[] {
  // Identical strings — no changes
  if (oldText === newText) {
    return oldText.length > 0 ? [{ type: "equal", text: oldText }] : [];
  }

  // Empty old — entire new text is an insertion
  if (oldText.length === 0) {
    return newText.length > 0 ? [{ type: "insert", text: newText }] : [];
  }

  // Empty new — entire old text is a deletion
  if (newText.length === 0) {
    return [{ type: "delete", text: oldText }];
  }

  // Find common prefix length
  let prefixLen = 0;
  const minLen = Math.min(oldText.length, newText.length);
  while (prefixLen < minLen && oldText[prefixLen] === newText[prefixLen]) {
    prefixLen++;
  }

  // Find common suffix length (not overlapping with prefix)
  let suffixLen = 0;
  const maxSuffix = minLen - prefixLen;
  while (
    suffixLen < maxSuffix &&
    oldText[oldText.length - 1 - suffixLen] === newText[newText.length - 1 - suffixLen]
  ) {
    suffixLen++;
  }

  const segments: DiffSegment[] = [];

  // Common prefix
  if (prefixLen > 0) {
    segments.push({ type: "equal", text: oldText.slice(0, prefixLen) });
  }

  // Middle portion — what changed
  const oldMiddle = oldText.slice(prefixLen, oldText.length - suffixLen);
  const newMiddle = newText.slice(prefixLen, newText.length - suffixLen);

  if (oldMiddle.length > 0) {
    segments.push({ type: "delete", text: oldMiddle });
  }
  if (newMiddle.length > 0) {
    segments.push({ type: "insert", text: newMiddle });
  }

  // Common suffix
  if (suffixLen > 0) {
    segments.push({ type: "equal", text: oldText.slice(oldText.length - suffixLen) });
  }

  return segments;
}

/**
 * Checks whether a diff result has any unchanged (equal) portions.
 * If false, the entire text was replaced (no benefit from granular tracking).
 */
export function diffHasUnchangedParts(segments: DiffSegment[]): boolean {
  return segments.some(s => s.type === "equal");
}
