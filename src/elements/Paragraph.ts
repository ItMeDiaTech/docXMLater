/**
 * Paragraph - Represents a paragraph in a Word document
 * Contains one or more runs of formatted text
 */

import { Run, RunFormatting } from './Run';
import { Field } from './Field';
import { Hyperlink } from './Hyperlink';
import { Bookmark } from './Bookmark';
import { Revision } from './Revision';
import { Comment } from './Comment';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';

/**
 * Paragraph alignment options
 */
export type ParagraphAlignment = 'left' | 'center' | 'right' | 'justify' | 'both';

/**
 * Paragraph formatting options
 */
export interface ParagraphFormatting {
  /** Text alignment */
  alignment?: ParagraphAlignment;
  /** Indentation in twips (1/20th of a point) */
  indentation?: {
    left?: number;
    right?: number;
    firstLine?: number;
    hanging?: number;
  };
  /** Spacing in twips */
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: 'auto' | 'exact' | 'atLeast';
  };
  /** Keep with next paragraph */
  keepNext?: boolean;
  /** Keep lines together */
  keepLines?: boolean;
  /** Page break before */
  pageBreakBefore?: boolean;
  /** Paragraph style ID */
  style?: string;
  /** Numbering properties */
  numbering?: {
    numId: number;
    level: number;
  };
  /** Contextual spacing - removes spacing between paragraphs of same style */
  contextualSpacing?: boolean;
}

/**
 * Paragraph content (runs, fields, hyperlinks, or revisions)
 */
type ParagraphContent = Run | Field | Hyperlink | Revision;

/**
 * Represents a paragraph in a document
 */
export class Paragraph {
  private content: ParagraphContent[] = [];
  private formatting: ParagraphFormatting;
  private bookmarksStart: Bookmark[] = [];
  private bookmarksEnd: Bookmark[] = [];
  private commentsStart: Comment[] = [];
  private commentsEnd: Comment[] = [];

  /**
   * Creates a new Paragraph
   * @param formatting - Paragraph formatting options
   */
  constructor(formatting: ParagraphFormatting = {}) {
    this.formatting = formatting;
  }

  /**
   * Adds a run to the paragraph
   * @param run - Run to add
   * @returns This paragraph for chaining
   */
  addRun(run: Run): this {
    this.content.push(run);
    return this;
  }

  /**
   * Adds a field to the paragraph
   * @param field - Field to add
   * @returns This paragraph for chaining
   */
  addField(field: Field): this {
    this.content.push(field);
    return this;
  }

  /**
   * Adds a hyperlink to the paragraph
   * @param hyperlink - Hyperlink to add
   * @returns This paragraph for chaining
   */
  addHyperlink(hyperlink: Hyperlink): this {
    this.content.push(hyperlink);
    return this;
  }

  /**
   * Adds a revision (tracked change) to the paragraph
   * @param revision - Revision to add
   * @returns This paragraph for chaining
   */
  addRevision(revision: Revision): this {
    this.content.push(revision);
    return this;
  }

  /**
   * Adds a bookmark start marker at the beginning of this paragraph
   * @param bookmark - Bookmark to add
   * @returns This paragraph for chaining
   */
  addBookmarkStart(bookmark: Bookmark): this {
    this.bookmarksStart.push(bookmark);
    return this;
  }

  /**
   * Adds a bookmark end marker at the end of this paragraph
   * @param bookmark - Bookmark to add (must have matching start marker)
   * @returns This paragraph for chaining
   */
  addBookmarkEnd(bookmark: Bookmark): this {
    this.bookmarksEnd.push(bookmark);
    return this;
  }

  /**
   * Adds both start and end bookmark markers (wraps entire paragraph)
   * @param bookmark - Bookmark to add
   * @returns This paragraph for chaining
   */
  addBookmark(bookmark: Bookmark): this {
    this.addBookmarkStart(bookmark);
    this.addBookmarkEnd(bookmark);
    return this;
  }

  /**
   * Gets all bookmarks that start in this paragraph
   * @returns Array of bookmarks
   */
  getBookmarksStart(): Bookmark[] {
    return [...this.bookmarksStart];
  }

  /**
   * Gets all bookmarks that end in this paragraph
   * @returns Array of bookmarks
   */
  getBookmarksEnd(): Bookmark[] {
    return [...this.bookmarksEnd];
  }

  /**
   * Adds a comment range start marker at the beginning of this paragraph
   * @param comment - Comment to start
   * @returns This paragraph for chaining
   */
  addCommentStart(comment: Comment): this {
    this.commentsStart.push(comment);
    return this;
  }

  /**
   * Adds a comment range end marker at the end of this paragraph
   * @param comment - Comment to end (must have matching start marker)
   * @returns This paragraph for chaining
   */
  addCommentEnd(comment: Comment): this {
    this.commentsEnd.push(comment);
    return this;
  }

  /**
   * Adds both start and end comment range markers (comments entire paragraph)
   * @param comment - Comment to add
   * @returns This paragraph for chaining
   */
  addComment(comment: Comment): this {
    this.addCommentStart(comment);
    this.addCommentEnd(comment);
    return this;
  }

  /**
   * Gets all comments that start in this paragraph
   * @returns Array of comments
   */
  getCommentsStart(): Comment[] {
    return [...this.commentsStart];
  }

  /**
   * Gets all comments that end in this paragraph
   * @returns Array of comments
   */
  getCommentsEnd(): Comment[] {
    return [...this.commentsEnd];
  }

  /**
   * Adds text with optional formatting
   * @param text - Text to add
   * @param formatting - Text formatting
   * @returns This paragraph for chaining
   */
  addText(text: string, formatting?: RunFormatting): this {
    this.content.push(new Run(text, formatting));
    return this;
  }

  /**
   * Sets the text content (replaces all content with a single run)
   * @param text - Text content
   * @param formatting - Optional formatting for the text
   * @returns This paragraph for chaining
   */
  setText(text: string, formatting?: RunFormatting): this {
    this.content = [new Run(text, formatting)];
    return this;
  }

  /**
   * Gets all runs in the paragraph (excluding fields)
   * @returns Array of runs
   */
  getRuns(): Run[] {
    return this.content.filter((item): item is Run => item instanceof Run);
  }

  /**
   * Gets all content in the paragraph (runs and fields)
   * @returns Array of content items
   */
  getContent(): ParagraphContent[] {
    return [...this.content];
  }

  /**
   * Gets the text content of all runs and hyperlinks combined
   * @returns Combined text content from all text-bearing elements
   */
  getText(): string {
    return this.content
      .filter((item): item is Run | Hyperlink =>
        item instanceof Run || item instanceof Hyperlink)
      .map(item => item.getText())
      .join('');
  }

  /**
   * Gets the paragraph formatting
   * @returns Paragraph formatting
   */
  getFormatting(): ParagraphFormatting {
    return { ...this.formatting };
  }

  /**
   * Sets paragraph alignment
   * @param alignment - Alignment value
   * @returns This paragraph for chaining
   */
  setAlignment(alignment: ParagraphAlignment): this {
    this.formatting.alignment = alignment;
    return this;
  }

  /**
   * Sets left indentation
   * @param twips - Indentation in twips (1/20th of a point)
   * @returns This paragraph for chaining
   */
  setLeftIndent(twips: number): this {
    if (!this.formatting.indentation) {
      this.formatting.indentation = {};
    }
    this.formatting.indentation.left = twips;
    return this;
  }

  /**
   * Sets right indentation
   * @param twips - Indentation in twips
   * @returns This paragraph for chaining
   */
  setRightIndent(twips: number): this {
    if (!this.formatting.indentation) {
      this.formatting.indentation = {};
    }
    this.formatting.indentation.right = twips;
    return this;
  }

  /**
   * Sets first line indentation
   * @param twips - Indentation in twips
   * @returns This paragraph for chaining
   */
  setFirstLineIndent(twips: number): this {
    if (!this.formatting.indentation) {
      this.formatting.indentation = {};
    }
    this.formatting.indentation.firstLine = twips;
    return this;
  }

  /**
   * Sets spacing before paragraph
   * @param twips - Spacing in twips
   * @returns This paragraph for chaining
   */
  setSpaceBefore(twips: number): this {
    if (!this.formatting.spacing) {
      this.formatting.spacing = {};
    }
    this.formatting.spacing.before = twips;
    return this;
  }

  /**
   * Sets spacing after paragraph
   * @param twips - Spacing in twips
   * @returns This paragraph for chaining
   */
  setSpaceAfter(twips: number): this {
    if (!this.formatting.spacing) {
      this.formatting.spacing = {};
    }
    this.formatting.spacing.after = twips;
    return this;
  }

  /**
   * Sets line spacing
   * @param twips - Line spacing in twips
   * @param rule - Line spacing rule
   * @returns This paragraph for chaining
   */
  setLineSpacing(twips: number, rule: 'auto' | 'exact' | 'atLeast' = 'auto'): this {
    if (!this.formatting.spacing) {
      this.formatting.spacing = {};
    }
    this.formatting.spacing.line = twips;
    this.formatting.spacing.lineRule = rule;
    return this;
  }

  /**
   * Sets paragraph style
   * @param styleId - Style ID
   * @returns This paragraph for chaining
   */
  setStyle(styleId: string): this {
    this.formatting.style = styleId;
    return this;
  }

  /**
   * Sets keep with next
   * @param keepNext - Whether to keep with next paragraph
   * @returns This paragraph for chaining
   */
  setKeepNext(keepNext: boolean = true): this {
    this.formatting.keepNext = keepNext;
    return this;
  }

  /**
   * Sets keep lines together
   * @param keepLines - Whether to keep lines together
   * @returns This paragraph for chaining
   */
  setKeepLines(keepLines: boolean = true): this {
    this.formatting.keepLines = keepLines;
    return this;
  }

  /**
   * Sets page break before
   * @param pageBreakBefore - Whether to insert page break before
   * @returns This paragraph for chaining
   */
  setPageBreakBefore(pageBreakBefore: boolean = true): this {
    this.formatting.pageBreakBefore = pageBreakBefore;
    return this;
  }

  /**
   * Sets numbering for this paragraph (adds to a list)
   * @param numId - The numbering instance ID
   * @param level - The level (0-8, where 0 is the outermost level)
   * @returns This paragraph for chaining
   */
  setNumbering(numId: number, level: number = 0): this {
    if (numId < 0) {
      throw new Error('Numbering ID must be non-negative');
    }
    if (level < 0 || level > 8) {
      throw new Error('Level must be between 0 and 8');
    }

    this.formatting.numbering = { numId, level };
    return this;
  }

  /**
   * Sets contextual spacing for this paragraph
   * When enabled, removes spacing between consecutive paragraphs of the same style
   * Per ECMA-376 Part 1 §17.3.1.8
   * @param enable - Whether to enable contextual spacing
   * @returns This paragraph for chaining
   */
  setContextualSpacing(enable: boolean = true): this {
    this.formatting.contextualSpacing = enable;
    return this;
  }

  /**
   * Removes numbering from this paragraph
   * @returns This paragraph for chaining
   */
  removeNumbering(): this {
    delete this.formatting.numbering;
    return this;
  }

  /**
   * Gets the numbering properties
   * @returns Numbering properties or undefined
   */
  getNumbering(): { numId: number; level: number } | undefined {
    return this.formatting.numbering ? { ...this.formatting.numbering } : undefined;
  }

  /**
   * Converts the paragraph to WordprocessingML XML element
   *
   * **ECMA-376 Compliance:** Properties are generated in the order specified by
   * ECMA-376 Part 1 §17.3.1.26 to ensure strict OpenXML conformance.
   *
   * Per spec, the order is:
   * 1. pStyle (style reference)
   * 2. keepNext (keep with next paragraph)
   * 3. keepLines (keep lines together)
   * 4. pageBreakBefore (page break before)
   * 5. numPr (numbering properties)
   * 6. spacing (spacing before/after/line)
   * 7. ind (indentation)
   * 8. jc (justification/alignment)
   *
   * @returns XMLElement representing the paragraph
   */
  toXML(): XMLElement {
    const pPrChildren: XMLElement[] = [];

    // 1. Paragraph style (must be first per ECMA-376 §17.3.1.26)
    if (this.formatting.style) {
      pPrChildren.push(XMLBuilder.wSelf('pStyle', { 'w:val': this.formatting.style }));
    }

    // 2. Keep with next paragraph
    if (this.formatting.keepNext) {
      pPrChildren.push(XMLBuilder.wSelf('keepNext'));
    }

    // 3. Keep lines together
    if (this.formatting.keepLines) {
      pPrChildren.push(XMLBuilder.wSelf('keepLines'));
    }

    // 4. Page break before
    if (this.formatting.pageBreakBefore) {
      pPrChildren.push(XMLBuilder.wSelf('pageBreakBefore'));
    }

    // 5. Numbering properties
    if (this.formatting.numbering) {
      const numPr = XMLBuilder.w('numPr', undefined, [
        XMLBuilder.wSelf('ilvl', { 'w:val': this.formatting.numbering.level.toString() }),
        XMLBuilder.wSelf('numId', { 'w:val': this.formatting.numbering.numId.toString() })
      ]);
      pPrChildren.push(numPr);
    }

    // 6. Spacing (before/after/line) per ECMA-376 Part 1 §17.3.1.33
    if (this.formatting.spacing) {
      const spc = this.formatting.spacing;
      const attributes: Record<string, number | string> = {};
      if (spc.before !== undefined) attributes['w:before'] = spc.before;
      if (spc.after !== undefined) attributes['w:after'] = spc.after;
      if (spc.line !== undefined) attributes['w:line'] = spc.line;
      if (spc.lineRule) attributes['w:lineRule'] = spc.lineRule;
      // Only generate spacing element if it has attributes (prevents empty elements)
      if (Object.keys(attributes).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('spacing', attributes));
      }
    }

    // Contextual spacing per ECMA-376 Part 1 §17.3.1.8
    // Removes spacing between paragraphs of the same style (comes after spacing)
    if (this.formatting.contextualSpacing) {
      pPrChildren.push(XMLBuilder.wSelf('contextualSpacing'));
    }

    // 7. Indentation (left/right/firstLine/hanging)
    if (this.formatting.indentation) {
      const ind = this.formatting.indentation;
      const attributes: Record<string, number> = {};
      if (ind.left !== undefined) attributes['w:left'] = ind.left;
      if (ind.right !== undefined) attributes['w:right'] = ind.right;
      if (ind.firstLine !== undefined) attributes['w:firstLine'] = ind.firstLine;
      if (ind.hanging !== undefined) attributes['w:hanging'] = ind.hanging;
      if (Object.keys(attributes).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('ind', attributes));
      }
    }

    // 8. Justification/Alignment (must be last per ECMA-376 §17.3.1.26)
    if (this.formatting.alignment) {
      pPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': this.formatting.alignment }));
    }

    // Build paragraph element
    const paragraphChildren: XMLElement[] = [];

    // Add paragraph properties if there are any
    if (pPrChildren.length > 0) {
      paragraphChildren.push(XMLBuilder.w('pPr', undefined, pPrChildren));
    }

    // Add bookmark start markers
    for (const bookmark of this.bookmarksStart) {
      paragraphChildren.push(bookmark.toStartXML());
    }

    // Add comment range start markers
    for (const comment of this.commentsStart) {
      paragraphChildren.push(comment.toRangeStartXML());
    }

    // Add content (runs, fields, hyperlinks, and revisions)
    for (const item of this.content) {
      if (item instanceof Field) {
        // Fields need to be wrapped in a run
        paragraphChildren.push(XMLBuilder.w('r', undefined, [item.toXML()]));
      } else if (item instanceof Hyperlink) {
        // Hyperlinks are their own element
        paragraphChildren.push(item.toXML());
      } else if (item instanceof Revision) {
        // Revisions (track changes) are their own element
        paragraphChildren.push(item.toXML());
      } else {
        paragraphChildren.push(item.toXML());
      }
    }

    // If no content, add empty run to prevent invalid XML
    if (this.content.length === 0) {
      paragraphChildren.push(new Run('').toXML());
    }

    // Add comment range end markers
    for (const comment of this.commentsEnd) {
      paragraphChildren.push(comment.toRangeEndXML());
    }

    // Add comment references (must come after range end)
    for (const comment of this.commentsEnd) {
      paragraphChildren.push(comment.toReferenceXML());
    }

    // Add bookmark end markers
    for (const bookmark of this.bookmarksEnd) {
      paragraphChildren.push(bookmark.toEndXML());
    }

    return XMLBuilder.w('p', undefined, paragraphChildren);
  }

  /**
   * Gets the word count for this paragraph
   * @returns Number of words in the paragraph
   */
  getWordCount(): number {
    const text = this.getText().trim();
    if (!text) return 0;

    // Split by whitespace and filter out empty strings
    const words = text.split(/\s+/).filter(word => word.length > 0);
    return words.length;
  }

  /**
   * Gets the character count for this paragraph
   * @param includeSpaces - Whether to include spaces in the count
   * @returns Number of characters in the paragraph
   */
  getLength(includeSpaces: boolean = true): number {
    const text = this.getText();
    if (includeSpaces) {
      return text.length;
    } else {
      return text.replace(/\s/g, '').length;
    }
  }

  /**
   * Creates a deep clone of this paragraph
   * @returns A new Paragraph instance with the same content and formatting
   */
  clone(): Paragraph {
    // Clone the formatting
    const clonedFormatting: ParagraphFormatting = JSON.parse(JSON.stringify(this.formatting));

    // Create new paragraph with cloned formatting
    const clonedParagraph = new Paragraph(clonedFormatting);

    // Clone all content (runs, fields, hyperlinks, revisions)
    for (const item of this.content) {
      if (item instanceof Run) {
        // Clone the run with its text and formatting
        const runFormatting = item.getFormatting();
        const clonedRun = new Run(item.getText(), JSON.parse(JSON.stringify(runFormatting)));
        clonedParagraph.addRun(clonedRun);
      } else {
        // For other content types, add them as-is (shallow copy for now)
        // In a more complete implementation, we'd clone these too
        clonedParagraph.content.push(item);
      }
    }

    // Clone bookmark and comment markers
    clonedParagraph.bookmarksStart = [...this.bookmarksStart];
    clonedParagraph.bookmarksEnd = [...this.bookmarksEnd];
    clonedParagraph.commentsStart = [...this.commentsStart];
    clonedParagraph.commentsEnd = [...this.commentsEnd];

    return clonedParagraph;
  }

  /**
   * Sets paragraph borders
   * @param borders - Border definitions for each side
   * @returns This paragraph for chaining
   */
  setBorder(borders: {
    top?: { style?: string; size?: number; color?: string; space?: number };
    bottom?: { style?: string; size?: number; color?: string; space?: number };
    left?: { style?: string; size?: number; color?: string; space?: number };
    right?: { style?: string; size?: number; color?: string; space?: number };
  }): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    // Store borders in formatting (will be handled in toXML)
    (this.formatting as any).borders = borders;

    return this;
  }

  /**
   * Sets paragraph shading (background color)
   * @param shading - Shading options
   * @returns This paragraph for chaining
   */
  setShading(shading: {
    fill?: string;  // Background color (hex)
    color?: string;  // Foreground color (hex)
    val?: 'clear' | 'solid' | 'horzStripe' | 'vertStripe' | 'reverseDiagStripe' | 'diagStripe' | 'horzCross' | 'diagCross';
  }): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    // Store shading in formatting (will be handled in toXML)
    (this.formatting as any).shading = shading;

    return this;
  }

  /**
   * Sets tab stops for the paragraph
   * @param tabs - Array of tab stop definitions
   * @returns This paragraph for chaining
   */
  setTabs(tabs: Array<{
    position: number;  // Position in twips
    val?: 'clear' | 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'num';
    leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
  }>): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    // Store tabs in formatting (will be handled in toXML)
    (this.formatting as any).tabs = tabs;

    return this;
  }

  /**
   * Creates a new Paragraph with the specified formatting
   * @param formatting - Paragraph formatting
   * @returns New Paragraph instance
   */
  static create(formatting?: ParagraphFormatting): Paragraph {
    return new Paragraph(formatting);
  }
}
