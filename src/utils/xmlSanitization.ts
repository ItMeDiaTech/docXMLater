/**
 * XML Sanitization Utilities
 *
 * Provides functions for validating and sanitizing text content per XML 1.0 specification.
 * Per XML 1.0, certain control characters are invalid and must be removed before
 * including text in XML documents.
 *
 * Valid characters in XML 1.0:
 * - 0x09 (tab), 0x0A (newline), 0x0D (carriage return)
 * - 0x20-0xD7FF, 0xE000-0xFFFD, 0x10000-0x10FFFF
 *
 * Invalid characters that must be removed:
 * - 0x00-0x08 (NULL through BACKSPACE)
 * - 0x0B-0x0C (VERTICAL TAB and FORM FEED)
 * - 0x0E-0x1F (SHIFT OUT through UNIT SEPARATOR)
 * - 0x7F (DELETE)
 * - 0xFFFE-0xFFFF (Unicode noncharacters excluded by the Char production)
 * - Unpaired surrogates (0xD800-0xDFFF outside a valid high/low pair)
 *
 * @module xmlSanitization
 */

import { getGlobalLogger } from './logger.js';

/**
 * Regular expression matching invalid XML 1.0 characters.
 * Matches: 0x00-0x08, 0x0B-0x0C, 0x0E-0x1F, 0x7F, 0xFFFE, 0xFFFF
 * Does NOT match valid chars: 0x09 (tab), 0x0A (newline), 0x0D (CR)
 */
const INVALID_XML_CHAR_REGEX = /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F￾￿]/g;

/**
 * Matches surrogate code units that do not form a valid high/low pair.
 * XML 1.0 Char excludes 0xD800-0xDFFF entirely; only properly paired
 * surrogates (which decode to a single 0x10000-0x10FFFF character) are
 * legal, so lone halves must be stripped or they corrupt the part at
 * UTF-8 encode time.
 */
const UNPAIRED_SURROGATE_REGEX =
  /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

/**
 * Removes invalid XML 1.0 characters from text.
 *
 * Per XML 1.0 spec, characters 0x00-0x08, 0x0B-0x0C, 0x0E-0x1F, 0x7F,
 * 0xFFFE, 0xFFFF, and unpaired surrogates are invalid and cannot appear
 * in XML documents. This function removes them.
 *
 * Valid control characters are preserved:
 * - Tab (0x09)
 * - Line Feed / Newline (0x0A)
 * - Carriage Return (0x0D)
 *
 * Properly paired surrogates (astral-plane characters such as emoji)
 * are preserved.
 *
 * @param text - Input text to sanitize
 * @param logWarning - If true, logs a warning when invalid chars are found (default: true)
 * @returns Sanitized text with invalid characters removed
 *
 * @example
 * ```typescript
 * // Remove NULL byte from text
 * const clean = removeInvalidXmlChars("Hello\x00World");
 * // Returns: "HelloWorld"
 *
 * // Tab and newline are preserved
 * const preserved = removeInvalidXmlChars("Hello\tWorld\n");
 * // Returns: "Hello\tWorld\n"
 * ```
 */
export function removeInvalidXmlChars(text: string, logWarning = true): string {
  // Reset regex lastIndex for global regexes
  INVALID_XML_CHAR_REGEX.lastIndex = 0;
  UNPAIRED_SURROGATE_REGEX.lastIndex = 0;

  if (logWarning && (INVALID_XML_CHAR_REGEX.test(text) || UNPAIRED_SURROGATE_REGEX.test(text))) {
    // Reset regex lastIndex after test
    INVALID_XML_CHAR_REGEX.lastIndex = 0;
    UNPAIRED_SURROGATE_REGEX.lastIndex = 0;

    const invalidChars = findInvalidXmlChars(text);
    const hexCodes = invalidChars
      .map((c) => `0x${c.toString(16).toUpperCase().padStart(2, '0')}`)
      .join(', ');
    getGlobalLogger().warn(`[XMLSanitization] Removing invalid XML characters: ${hexCodes}`);
  }

  // Reset regex lastIndex before replace
  INVALID_XML_CHAR_REGEX.lastIndex = 0;
  UNPAIRED_SURROGATE_REGEX.lastIndex = 0;
  return text.replace(INVALID_XML_CHAR_REGEX, '').replace(UNPAIRED_SURROGATE_REGEX, '');
}

/**
 * Finds all invalid XML 1.0 characters in text.
 *
 * Returns an array of unique character codes that are invalid per XML 1.0 spec,
 * including 0xFFFE/0xFFFF noncharacters and unpaired surrogate code units.
 * This is useful for diagnostics and error reporting.
 *
 * @param text - Text to scan for invalid characters
 * @returns Array of unique invalid character codes found, or empty array if text is valid
 *
 * @example
 * ```typescript
 * const invalid = findInvalidXmlChars("Hello\x00\x08World");
 * // Returns: [0, 8] - NULL and BACKSPACE codes
 *
 * const valid = findInvalidXmlChars("Hello\tWorld");
 * // Returns: [] - tab is valid
 * ```
 */
export function findInvalidXmlChars(text: string): number[] {
  const invalid: number[] = [];

  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);

    // Check if character is in invalid ranges
    let isInvalid =
      (code >= 0x00 && code <= 0x08) || // NULL through BACKSPACE
      (code >= 0x0b && code <= 0x0c) || // VERTICAL TAB and FORM FEED
      (code >= 0x0e && code <= 0x1f) || // SHIFT OUT through UNIT SEPARATOR
      code === 0x7f || // DELETE
      code === 0xfffe || // Noncharacter excluded by XML 1.0 Char
      code === 0xffff; // Noncharacter excluded by XML 1.0 Char

    if (!isInvalid && code >= 0xd800 && code <= 0xdbff) {
      // High surrogate is only valid when followed by a low surrogate
      const next = text.charCodeAt(i + 1); // NaN at end of string fails the range check
      isInvalid = !(next >= 0xdc00 && next <= 0xdfff);
    } else if (!isInvalid && code >= 0xdc00 && code <= 0xdfff) {
      // Low surrogate is only valid when preceded by a high surrogate
      const prev = text.charCodeAt(i - 1); // NaN at start of string fails the range check
      isInvalid = !(prev >= 0xd800 && prev <= 0xdbff);
    }

    if (isInvalid) {
      // Only add unique codes
      if (!invalid.includes(code)) {
        invalid.push(code);
      }
    }
  }

  return invalid;
}

/**
 * Checks if text contains any invalid XML 1.0 characters.
 *
 * This is a fast check that returns true/false without identifying specific characters.
 * Use `findInvalidXmlChars()` if you need to know which characters are invalid.
 *
 * @param text - Text to check
 * @returns true if text contains invalid characters, false otherwise
 *
 * @example
 * ```typescript
 * hasInvalidXmlChars("Hello\x00World");  // true - NULL byte
 * hasInvalidXmlChars("Hello\tWorld");    // false - tab is valid
 * hasInvalidXmlChars("Normal text");     // false
 * ```
 */
export function hasInvalidXmlChars(text: string): boolean {
  // Reset regex lastIndex for global regexes
  INVALID_XML_CHAR_REGEX.lastIndex = 0;
  UNPAIRED_SURROGATE_REGEX.lastIndex = 0;
  return INVALID_XML_CHAR_REGEX.test(text) || UNPAIRED_SURROGATE_REGEX.test(text);
}

/**
 * CT_RPr child element order per ECMA-376 Part 1 §17.3.2.28 (local names,
 * `w:` namespace). Used to repair raw-preserved run-property blocks whose
 * children are out of schema order — Word rejects, for example, a `w:b`
 * that appears after `w:color`/`w:sz` and reports the document as corrupt.
 */
const CT_RPR_CHILD_ORDER: readonly string[] = [
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
];

/**
 * Splits the inner content of an element into its top-level child elements,
 * respecting nesting so that a child's descendants are kept with it. Returns
 * the raw XML fragment for each direct child in document order. Non-element
 * text between children (whitespace) is dropped, which is safe for the
 * element-only content models this helper targets (e.g. `w:rPr`).
 */
function splitTopLevelElements(inner: string): string[] {
  const children: string[] = [];
  const tagRegex = /<([A-Za-z][\w:.-]*)([^>]*?)(\/?)>|<\/([A-Za-z][\w:.-]*)\s*>/g;
  let depth = 0;
  let start = -1;
  let match: RegExpExecArray | null;
  while ((match = tagRegex.exec(inner)) !== null) {
    const isClosing = match[4] !== undefined;
    const isSelfClosing = match[3] === '/';
    if (isClosing) {
      if (depth > 0) {
        depth--;
        if (depth === 0 && start !== -1) {
          children.push(inner.slice(start, tagRegex.lastIndex));
          start = -1;
        }
      }
      continue;
    }
    if (isSelfClosing) {
      if (depth === 0) {
        children.push(inner.slice(match.index, tagRegex.lastIndex));
      }
      continue;
    }
    // Opening tag of a paired element
    if (depth === 0) start = match.index;
    depth++;
  }
  return children;
}

/**
 * Reorders the direct children of every `<w:rPr>` block in a raw XML part to
 * the ECMA-376 CT_RPr sequence. Unknown/unmodeled children retain their
 * relative position at the end of the block. The transform is idempotent:
 * already-ordered run properties are returned unchanged.
 *
 * Some producers (and older versions of this framework) appended boolean
 * toggles such as `w:b`/`w:bCs` to the end of an existing `w:rPr`, leaving
 * them after `w:color`/`w:sz`. Word treats the out-of-order child as an
 * unexpected element and flags the whole document as corrupt. Numbering
 * definitions are written back verbatim for fidelity, so this repair runs on
 * the preserved part before it is saved.
 *
 * @param xml - Raw XML for a part containing `w:rPr` blocks (e.g. numbering.xml)
 * @returns XML with every `w:rPr` child sequence normalized to schema order
 */
export function reorderRunPropertyChildren(xml: string): string {
  if (!xml || !xml.includes('<w:rPr')) return xml;

  const orderIndex = (fragment: string): number => {
    const nameMatch = /^<w:(\w+)/.exec(fragment);
    const name = nameMatch?.[1];
    if (!name) return CT_RPR_CHILD_ORDER.length;
    const idx = CT_RPR_CHILD_ORDER.indexOf(name);
    return idx === -1 ? CT_RPR_CHILD_ORDER.length : idx;
  };

  // Only match non-empty <w:rPr>...</w:rPr> blocks; self-closing <w:rPr/> has
  // no children to reorder.
  return xml.replace(/<w:rPr(\s[^>]*)?>([\s\S]*?)<\/w:rPr>/g, (full, attrs, inner) => {
    if (!inner || !(inner as string).includes('<')) return full;
    const children = splitTopLevelElements(inner);
    if (children.length < 2) return full;

    const ordered = children
      .map((frag, i) => ({ frag, i, order: orderIndex(frag) }))
      // Stable sort: ties (including unknown children) keep original order.
      .sort((a, b) => a.order - b.order || a.i - b.i);

    if (ordered.every((c, i) => c.i === i)) return full; // already in order

    return `<w:rPr${attrs ?? ''}>${ordered.map((c) => c.frag).join('')}</w:rPr>`;
  });
}

/**
 * Normalizes malformed floating-drawing relative-size elements in a raw anchor
 * fragment. Per ECMA-376 (DrawingML wordprocessing 2010 extensions),
 * `wp14:sizeRelH` must contain a `wp14:pctWidth` child and `wp14:sizeRelV` a
 * `wp14:pctHeight` child. Some producers emit the percentage as bare text
 * content (`<wp14:sizeRelH relativeFrom="margin">0</wp14:sizeRelH>`), which
 * Word rejects as incomplete content. This wraps such text in the required
 * child element. Already-valid nested forms are left untouched because the
 * text-content pattern (`[^<]`) cannot match an element child.
 *
 * @param xml - Raw anchor-extras fragment captured during parsing
 * @returns Fragment with sizeRelH/sizeRelV normalized to schema-valid form
 */
export function normalizeAnchorSizeRel(xml: string): string {
  if (!xml) return xml;
  return xml
    .replace(
      /(<wp14:sizeRelH\b[^>]*>)([^<]+)(<\/wp14:sizeRelH>)/g,
      (_m, open: string, text: string, close: string) =>
        text.trim() === ''
          ? `${open}${text}${close}`
          : `${open}<wp14:pctWidth>${text.trim()}</wp14:pctWidth>${close}`
    )
    .replace(
      /(<wp14:sizeRelV\b[^>]*>)([^<]+)(<\/wp14:sizeRelV>)/g,
      (_m, open: string, text: string, close: string) =>
        text.trim() === ''
          ? `${open}${text}${close}`
          : `${open}<wp14:pctHeight>${text.trim()}</wp14:pctHeight>${close}`
    );
}

/**
 * Character code constants for documentation and testing.
 */
export const XML_CONTROL_CHARS = {
  /** NULL (0x00) - Invalid */
  NULL: 0x00,
  /** Start of Heading (0x01) - Invalid */
  SOH: 0x01,
  /** Start of Text (0x02) - Invalid */
  STX: 0x02,
  /** End of Text (0x03) - Invalid */
  ETX: 0x03,
  /** End of Transmission (0x04) - Invalid */
  EOT: 0x04,
  /** Enquiry (0x05) - Invalid */
  ENQ: 0x05,
  /** Acknowledge (0x06) - Invalid */
  ACK: 0x06,
  /** Bell (0x07) - Invalid */
  BEL: 0x07,
  /** Backspace (0x08) - Invalid */
  BS: 0x08,
  /** Horizontal Tab (0x09) - VALID */
  TAB: 0x09,
  /** Line Feed / Newline (0x0A) - VALID */
  LF: 0x0a,
  /** Vertical Tab (0x0B) - Invalid */
  VT: 0x0b,
  /** Form Feed (0x0C) - Invalid */
  FF: 0x0c,
  /** Carriage Return (0x0D) - VALID */
  CR: 0x0d,
  /** Shift Out (0x0E) - Invalid */
  SO: 0x0e,
  /** Unit Separator (0x1F) - Invalid */
  US: 0x1f,
  /** Delete (0x7F) - Invalid */
  DEL: 0x7f,
} as const;
