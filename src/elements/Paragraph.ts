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
 * Border style types for paragraph borders
 */
export type BorderStyle = 'single' | 'double' | 'dashed' | 'dotted' | 'thick' | 'none';

/**
 * Shading pattern types
 */
export type ShadingPattern = 'clear' | 'solid' | 'horzStripe' | 'vertStripe' |
  'reverseDiagStripe' | 'diagStripe' | 'horzCross' | 'diagCross';

/**
 * Tab stop alignment types
 */
export type TabAlignment = 'clear' | 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'num';

/**
 * Tab stop leader types
 */
export type TabLeader = 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';

/**
 * Single border definition
 */
export interface BorderDefinition {
  /** Border style */
  style?: BorderStyle;
  /** Border width in eighths of a point (1-96) */
  size?: number;
  /** Border color (hex without #) */
  color?: string;
  /** Space between border and text in points (0-31) */
  space?: number;
}

/**
 * Tab stop definition
 */
export interface TabStop {
  /** Position in twips */
  position: number;
  /** Alignment type */
  val?: TabAlignment;
  /** Leader character */
  leader?: TabLeader;
}

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
  /** Paragraph ID (Word 2010+) - required by modern Word for change tracking */
  paraId?: string;
  /** Paragraph borders (top, bottom, left, right, between, bar) */
  borders?: {
    top?: BorderDefinition;
    bottom?: BorderDefinition;
    left?: BorderDefinition;
    right?: BorderDefinition;
    between?: BorderDefinition;
    bar?: BorderDefinition;
  };
  /** Paragraph shading (background color and pattern) */
  shading?: {
    /** Background fill color (hex without #) */
    fill?: string;
    /** Foreground color for patterns (hex without #) */
    color?: string;
    /** Shading pattern type */
    val?: ShadingPattern;
  };
  /** Tab stops */
  tabs?: TabStop[];
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
  public formatting: ParagraphFormatting;
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
   * Creates a detached paragraph (not yet added to a document)
   * @param textOrFormatting - Optional text content or paragraph formatting
   * @param formatting - Optional paragraph formatting (only used if first param is text)
   * @returns New Paragraph instance
   * @example
   * // Create with text and formatting
   * const para1 = Paragraph.create('Hello World', { alignment: 'center' });
   *
   * // Create with just formatting
   * const para2 = Paragraph.create({ alignment: 'right' });
   *
   * // Create empty
   * const para3 = Paragraph.create();
   *
   * // Add to document later
   * doc.addParagraph(para1);
   */
  static create(textOrFormatting?: string | ParagraphFormatting, formatting?: ParagraphFormatting): Paragraph {
    // Handle overloaded parameters
    if (typeof textOrFormatting === 'string') {
      // First param is text
      const paragraph = new Paragraph(formatting);
      paragraph.addText(textOrFormatting);
      return paragraph;
    } else {
      // First param is formatting (or undefined)
      return new Paragraph(textOrFormatting);
    }
  }

  /**
   * Creates a detached paragraph with a specific style
   * @param text - Text content
   * @param styleId - Style ID (e.g., 'Heading1', 'Title')
   * @returns New Paragraph instance
   * @example
   * const heading = Paragraph.createWithStyle('Chapter 1', 'Heading1');
   * doc.addParagraph(heading);
   */
  static createWithStyle(text: string, styleId: string): Paragraph {
    const paragraph = new Paragraph({ style: styleId });
    paragraph.addText(text);
    return paragraph;
  }

  /**
   * Creates a detached empty paragraph
   * Useful for adding blank lines or spacing
   * @returns New empty Paragraph instance
   * @example
   * const blank = Paragraph.createEmpty();
   * doc.addParagraph(blank);
   */
  static createEmpty(): Paragraph {
    return new Paragraph();
  }

  /**
   * Creates a detached paragraph with formatted text
   * @param text - Text content
   * @param runFormatting - Run formatting (bold, italic, etc.)
   * @param paragraphFormatting - Paragraph formatting (alignment, spacing, etc.)
   * @returns New Paragraph instance
   * @example
   * const para = Paragraph.createFormatted(
   *   'Important Text',
   *   { bold: true, color: 'FF0000' },
   *   { alignment: 'center' }
   * );
   */
  static createFormatted(
    text: string,
    runFormatting?: RunFormatting,
    paragraphFormatting?: ParagraphFormatting
  ): Paragraph {
    const paragraph = new Paragraph(paragraphFormatting);
    paragraph.addText(text, runFormatting);
    return paragraph;
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
   * Gets the paragraph style ID
   * @returns Style ID or undefined if no style is set
   */
  getStyle(): string | undefined {
    return this.formatting.style;
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
   *
   * **Automatic Conflict Resolution:**
   * When setting keepNext to true, pageBreakBefore is automatically cleared to prevent
   * layout conflicts. The pageBreakBefore property creates contradictory constraints
   * (breaking the page vs. keeping content together) that cause massive whitespace
   * in Word as the layout engine tries to satisfy both. Keep properties take priority
   * as they represent explicit user intent to keep content together.
   *
   * @param keepNext - Whether to keep with next paragraph
   * @returns This paragraph for chaining
   */
  setKeepNext(keepNext: boolean = true): this {
    this.formatting.keepNext = keepNext;

    // Resolve property conflicts: keepNext contradicts pageBreakBefore
    if (keepNext) {
      this.formatting.pageBreakBefore = false;
    }

    return this;
  }

  /**
   * Sets keep lines together
   *
   * **Automatic Conflict Resolution:**
   * When setting keepLines to true, pageBreakBefore is automatically cleared to prevent
   * layout conflicts. The pageBreakBefore property creates contradictory constraints
   * (breaking the page vs. keeping lines together) that cause massive whitespace
   * in Word as the layout engine tries to satisfy both. Keep properties take priority
   * as they represent explicit user intent to keep content together.
   *
   * @param keepLines - Whether to keep lines together
   * @returns This paragraph for chaining
   */
  setKeepLines(keepLines: boolean = true): this {
    this.formatting.keepLines = keepLines;

    // Resolve property conflicts: keepLines contradicts pageBreakBefore
    if (keepLines) {
      this.formatting.pageBreakBefore = false;
    }

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
      pPrChildren.push(XMLBuilder.wSelf('contextualSpacing', { 'w:val': '1' }));
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

    // 8. Paragraph borders per ECMA-376 Part 1 §17.3.1.24
    if (this.formatting.borders) {
      const borderChildren: XMLElement[] = [];
      const borders = this.formatting.borders;

      // Helper function to create border element
      const createBorder = (borderType: string, border: BorderDefinition | undefined): XMLElement | null => {
        if (!border) return null;
        const attributes: Record<string, string | number> = {};
        if (border.style) attributes['w:val'] = border.style;
        if (border.size !== undefined) attributes['w:sz'] = border.size;
        if (border.color) attributes['w:color'] = border.color;
        if (border.space !== undefined) attributes['w:space'] = border.space;
        if (Object.keys(attributes).length > 0) {
          return XMLBuilder.wSelf(borderType, attributes);
        }
        return null;
      };

      // Add borders in order: top, left, bottom, right, between, bar
      const topBorder = createBorder('top', borders.top);
      if (topBorder) borderChildren.push(topBorder);

      const leftBorder = createBorder('left', borders.left);
      if (leftBorder) borderChildren.push(leftBorder);

      const bottomBorder = createBorder('bottom', borders.bottom);
      if (bottomBorder) borderChildren.push(bottomBorder);

      const rightBorder = createBorder('right', borders.right);
      if (rightBorder) borderChildren.push(rightBorder);

      const betweenBorder = createBorder('between', borders.between);
      if (betweenBorder) borderChildren.push(betweenBorder);

      const barBorder = createBorder('bar', borders.bar);
      if (barBorder) borderChildren.push(barBorder);

      if (borderChildren.length > 0) {
        pPrChildren.push(XMLBuilder.w('pBdr', undefined, borderChildren));
      }
    }

    // 9. Paragraph shading per ECMA-376 Part 1 §17.3.1.32
    if (this.formatting.shading) {
      const shd = this.formatting.shading;
      const attributes: Record<string, string> = {};
      if (shd.fill) attributes['w:fill'] = shd.fill;
      if (shd.color) attributes['w:color'] = shd.color;
      if (shd.val) attributes['w:val'] = shd.val;
      if (Object.keys(attributes).length > 0) {
        pPrChildren.push(XMLBuilder.wSelf('shd', attributes));
      }
    }

    // 10. Tab stops per ECMA-376 Part 1 §17.3.1.38
    if (this.formatting.tabs && this.formatting.tabs.length > 0) {
      const tabChildren: XMLElement[] = [];
      for (const tab of this.formatting.tabs) {
        const attributes: Record<string, string | number> = {};
        attributes['w:pos'] = tab.position;
        if (tab.val) attributes['w:val'] = tab.val;
        if (tab.leader) attributes['w:leader'] = tab.leader;
        tabChildren.push(XMLBuilder.wSelf('tab', attributes));
      }
      if (tabChildren.length > 0) {
        pPrChildren.push(XMLBuilder.w('tabs', undefined, tabChildren));
      }
    }

    // 11. Justification/Alignment (must be last per ECMA-376 §17.3.1.26)
    if (this.formatting.alignment) {
      // Map 'justify' to 'both' per ECMA-376 (Word uses 'both' for justified text)
      const alignmentValue = this.formatting.alignment === 'justify' ? 'both' : this.formatting.alignment;
      pPrChildren.push(XMLBuilder.wSelf('jc', { 'w:val': alignmentValue }));
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

    // Add paragraph-level attributes (Word 2010+ requires w14:paraId)
    const paragraphAttributes: Record<string, string> = {};
    if (this.formatting.paraId) {
      paragraphAttributes['w14:paraId'] = this.formatting.paraId;
    }

    return XMLBuilder.w('p', Object.keys(paragraphAttributes).length > 0 ? paragraphAttributes : undefined, paragraphChildren);
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
   * @example
   * ```typescript
   * para.setBorder({
   *   top: { style: 'single', size: 4, color: '000000', space: 1 },
   *   bottom: { style: 'single', size: 4, color: '000000', space: 1 }
   * });
   * ```
   */
  setBorder(borders: {
    top?: BorderDefinition;
    bottom?: BorderDefinition;
    left?: BorderDefinition;
    right?: BorderDefinition;
    between?: BorderDefinition;
    bar?: BorderDefinition;
  }): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    this.formatting.borders = borders;

    return this;
  }

  /**
   * Sets paragraph shading (background color and pattern)
   * @param shading - Shading options
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * // Solid background
   * para.setShading({ fill: 'FFFF00', val: 'solid' });
   *
   * // Pattern with colors
   * para.setShading({
   *   fill: 'FFFF00',
   *   color: '000000',
   *   val: 'diagStripe'
   * });
   * ```
   */
  setShading(shading: {
    fill?: string;
    color?: string;
    val?: ShadingPattern;
  }): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    this.formatting.shading = shading;

    return this;
  }

  /**
   * Sets tab stops for the paragraph
   * @param tabs - Array of tab stop definitions
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * para.setTabs([
   *   { position: 720, val: 'left' },
   *   { position: 1440, val: 'center', leader: 'dot' },
   *   { position: 2160, val: 'right' }
   * ]);
   * ```
   */
  setTabs(tabs: TabStop[]): this {
    if (!this.formatting) {
      this.formatting = {};
    }

    this.formatting.tabs = tabs;

    return this;
  }

  /**
   * Inserts a run at a specific position
   * @param index - Position to insert at (0-based)
   * @param run - Run to insert
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * const para = new Paragraph();
   * const run = new Run('Inserted', { bold: true });
   * para.insertRunAt(0, run);
   * ```
   */
  insertRunAt(index: number, run: Run): this {
    if (index < 0) index = 0;
    if (index > this.content.length) index = this.content.length;

    this.content.splice(index, 0, run);
    return this;
  }

  /**
   * Removes a run at a specific position
   * @param index - Position to remove (0-based)
   * @returns True if removed, false if index invalid
   * @example
   * ```typescript
   * para.removeRunAt(2);  // Remove third run
   * ```
   */
  removeRunAt(index: number): boolean {
    if (index >= 0 && index < this.content.length) {
      const item = this.content[index];
      if (item instanceof Run) {
        this.content.splice(index, 1);
        return true;
      }
    }
    return false;
  }

  /**
   * Replaces a run at a specific position
   * @param index - Position to replace (0-based)
   * @param run - New run
   * @returns True if replaced, false if index invalid or not a run
   * @example
   * ```typescript
   * const newRun = new Run('Replacement', { italic: true });
   * para.replaceRunAt(1, newRun);
   * ```
   */
  replaceRunAt(index: number, run: Run): boolean {
    if (index >= 0 && index < this.content.length) {
      const item = this.content[index];
      if (item instanceof Run) {
        this.content[index] = run;
        return true;
      }
    }
    return false;
  }

  /**
   * Finds text within the paragraph
   * @param text - Text to search for
   * @param options - Search options
   * @returns Array of run indices containing the text
   * @example
   * ```typescript
   * const indices = para.findText('important');
   * console.log(`Found in runs: ${indices.join(', ')}`);
   * ```
   */
  findText(text: string, options?: { caseSensitive?: boolean; wholeWord?: boolean }): number[] {
    const indices: number[] = [];
    const caseSensitive = options?.caseSensitive || false;
    const wholeWord = options?.wholeWord || false;

    const searchText = caseSensitive ? text : text.toLowerCase();

    for (let i = 0; i < this.content.length; i++) {
      const item = this.content[i];
      if (item instanceof Run) {
        const runText = caseSensitive ? item.getText() : item.getText().toLowerCase();

        if (wholeWord) {
          // Use word boundary regex
          const wordPattern = new RegExp(`\\b${searchText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`);
          if (wordPattern.test(runText)) {
            indices.push(i);
          }
        } else {
          // Simple substring search
          if (runText.includes(searchText)) {
            indices.push(i);
          }
        }
      }
    }

    return indices;
  }

  /**
   * Replaces text within paragraph runs
   * @param find - Text to find
   * @param replace - Replacement text
   * @param options - Search options
   * @returns Number of replacements made
   * @example
   * ```typescript
   * const count = para.replaceText('old', 'new');
   * console.log(`Replaced ${count} occurrences`);
   * ```
   */
  replaceText(find: string, replace: string, options?: { caseSensitive?: boolean; wholeWord?: boolean }): number {
    let replacementCount = 0;
    const caseSensitive = options?.caseSensitive || false;
    const wholeWord = options?.wholeWord || false;

    for (const item of this.content) {
      if (item instanceof Run) {
        const originalText = item.getText();
        let newText = originalText;

        if (wholeWord) {
          // Use word boundary regex
          const wordPattern = new RegExp(
            `\\b${find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`,
            caseSensitive ? 'g' : 'gi'
          );
          const matches = originalText.match(wordPattern);
          if (matches) {
            replacementCount += matches.length;
            newText = originalText.replace(wordPattern, replace);
          }
        } else {
          // Simple substring replacement
          const searchPattern = new RegExp(
            find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
            caseSensitive ? 'g' : 'gi'
          );
          const matches = originalText.match(searchPattern);
          if (matches) {
            replacementCount += matches.length;
            newText = originalText.replace(searchPattern, replace);
          }
        }

        if (newText !== originalText) {
          item.setText(newText);
        }
      }
    }

    return replacementCount;
  }

  /**
   * Merges another paragraph's content into this one
   * @param otherPara - Paragraph to merge
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * const para1 = new Paragraph().addText('Hello ');
   * const para2 = new Paragraph().addText('World');
   * para1.mergeWith(para2);  // para1 now contains "Hello World"
   * ```
   */
  mergeWith(otherPara: Paragraph): this {
    // Add all content from other paragraph
    for (const item of otherPara.content) {
      if (item instanceof Run) {
        this.content.push(item.clone());
      } else {
        this.content.push(item);
      }
    }

    // Merge bookmarks
    this.bookmarksStart.push(...otherPara.bookmarksStart);
    this.bookmarksEnd.push(...otherPara.bookmarksEnd);

    // Merge comments
    this.commentsStart.push(...otherPara.commentsStart);
    this.commentsEnd.push(...otherPara.commentsEnd);

    return this;
  }

  /**
   * Clears direct run formatting from all runs in this paragraph
   *
   * This is useful when applying styles, as direct run formatting takes precedence
   * over style formatting in Word. By clearing direct formatting, the style's
   * formatting can take effect.
   *
   * @param properties - Optional array of specific properties to clear.
   *                     If not specified, clears ALL direct formatting.
   *                     Valid properties: 'bold', 'italic', 'underline', 'strike',
   *                     'font', 'size', 'color', 'highlight', 'subscript', 'superscript',
   *                     'smallCaps', 'allCaps', 'dstrike'
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * // Clear all direct formatting
   * para.clearDirectRunFormatting();
   *
   * // Clear only font and color
   * para.clearDirectRunFormatting(['font', 'color']);
   *
   * // Apply style and clear all formatting
   * para.setStyle('Heading1').clearDirectRunFormatting();
   * ```
   */
  clearDirectRunFormatting(properties?: string[]): this {
    const runs = this.getRuns();

    for (const run of runs) {
      const formatting = run.getFormatting();

      if (properties && properties.length > 0) {
        // Clear only specified properties
        const newFormatting: RunFormatting = { ...formatting };
        for (const prop of properties) {
          if (prop in newFormatting) {
            delete (newFormatting as any)[prop];
          }
        }

        // Create new run with cleared formatting
        const text = run.getText();
        const newRun = new Run(text, newFormatting);

        // Replace in content array
        const index = this.content.indexOf(run);
        if (index !== -1) {
          this.content[index] = newRun;
        }
      } else {
        // Clear ALL direct formatting - replace with plain text run
        const text = run.getText();
        const newRun = new Run(text, {});

        // Replace in content array
        const index = this.content.indexOf(run);
        if (index !== -1) {
          this.content[index] = newRun;
        }
      }
    }

    return this;
  }

  /**
   * Applies a style to this paragraph and optionally clears direct run formatting
   *
   * In Word, direct run formatting takes precedence over style formatting. This method
   * provides a way to apply a style while ensuring it takes effect by clearing
   * conflicting direct formatting.
   *
   * @param styleId - Style ID to apply (e.g., 'Normal', 'Heading1')
   * @param clearProperties - Optional array of run properties to clear.
   *                          If not specified, does NOT clear any formatting (style overlay behavior).
   *                          Pass empty array [] to clear ALL formatting.
   * @returns This paragraph for chaining
   * @example
   * ```typescript
   * // Apply Heading1 and clear all direct formatting
   * para.applyStyleAndClearFormatting('Heading1', []);
   *
   * // Apply Normal and clear only font and color
   * para.applyStyleAndClearFormatting('Normal', ['font', 'color']);
   *
   * // Apply Title but keep existing run formatting (overlay style)
   * para.applyStyleAndClearFormatting('Title');
   * ```
   */
  applyStyleAndClearFormatting(styleId: string, clearProperties?: string[]): this {
    // Apply the style
    this.setStyle(styleId);

    // Clear direct formatting if requested
    if (clearProperties !== undefined) {
      this.clearDirectRunFormatting(clearProperties.length === 0 ? undefined : clearProperties);
    }

    return this;
  }

}
