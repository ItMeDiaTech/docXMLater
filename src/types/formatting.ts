/**
 * Additional formatting types for enhanced document manipulation
 */

/**
 * Border style options
 */
export type BorderStyleType = 'single' | 'double' | 'dashed' | 'dotted' | 'triple' | 'none';

/**
 * Individual border configuration
 */
export interface BorderStyle {
  /** Border style */
  style: BorderStyleType;
  /** Border width in eighths of a point */
  width: number;
  /** Border color in hex format */
  color: string;
  /** Space between border and text in points */
  space?: number;
}

/**
 * Paragraph border configuration
 */
export interface ParagraphBorder {
  /** Top border */
  top?: BorderStyle;
  /** Bottom border */
  bottom?: BorderStyle;
  /** Left border */
  left?: BorderStyle;
  /** Right border */
  right?: BorderStyle;
  /** Apply border between paragraphs */
  between?: boolean;
}

/**
 * Shading patterns
 */
export type ShadingPattern =
  | 'clear'
  | 'solid'
  | 'pct5'
  | 'pct10'
  | 'pct20'
  | 'pct25'
  | 'pct30'
  | 'pct40'
  | 'pct50'
  | 'pct60'
  | 'pct70'
  | 'pct75'
  | 'pct80'
  | 'pct90'
  | 'horzStripe'
  | 'vertStripe'
  | 'diagStripe'
  | 'diagCross';

/**
 * Paragraph shading configuration
 */
export interface ParagraphShading {
  /** Fill color in hex format */
  fill: string;
  /** Shading pattern */
  pattern?: ShadingPattern;
  /** Pattern color in hex format */
  color?: string;
}

/**
 * Tab stop alignment
 */
export type TabAlignment = 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear';

/**
 * Tab leader character
 */
export type TabLeader = 'none' | 'dot' | 'dash' | 'underscore' | 'heavy' | 'middleDot';

/**
 * Tab stop configuration
 */
export interface TabStop {
  /** Position in twips from left margin */
  position: number;
  /** Tab alignment type */
  type: TabAlignment;
  /** Leader character */
  leader?: TabLeader;
}

/**
 * Text search options
 */
export interface FindOptions {
  /** Case sensitive search */
  caseSensitive?: boolean;
  /** Match whole words only */
  wholeWord?: boolean;
  /** Use regular expressions */
  useRegex?: boolean;
  /** Include headers and footers */
  includeHeadersFooters?: boolean;
  /** Include footnotes and endnotes */
  includeNotes?: boolean;
}

/**
 * Text replacement options
 */
export interface ReplaceOptions extends FindOptions {
  /** Maximum number of replacements (0 = unlimited) */
  maxReplacements?: number;
  /** Track changes for replacements */
  trackChanges?: boolean;
  /** Author for track changes */
  author?: string;
}

/**
 * Search result
 */
export interface SearchResult {
  /** The paragraph containing the match */
  paragraph: any; // Will be Paragraph type
  /** The run containing the match */
  run?: any; // Will be Run type
  /** Match start position in paragraph */
  startIndex: number;
  /** Match end position in paragraph */
  endIndex: number;
  /** The matched text */
  match: string;
}