/**
 * Comment - Represents a comment/annotation in a Word document
 *
 * Comments allow users to annotate specific text ranges in a document.
 * They include author, date, content, and can have replies.
 */

import { Run } from './Run.js';
import { XMLElement } from '../xml/XMLBuilder.js';
import { formatDateForXml } from '../utils/dateFormatting.js';

/**
 * Inline comment content preserved as raw XML (e.g. w:hyperlink wrappers
 * from a parsed comments.xml). The wrapper is re-emitted verbatim on
 * regeneration so its r:id target survives; the runs parsed from inside
 * it are kept alongside so text extraction still sees their text.
 */
export interface CommentPreservedContent {
  /** Original XML emitted verbatim when comments.xml is regenerated */
  rawXml: string;
  /** Runs parsed from inside the preserved element (text extraction only) */
  runs: Run[];
}

/** Ordered item of comment paragraph content */
export type CommentContentItem = Run | CommentPreservedContent;

/**
 * One paragraph of comment content. CT_Comment accepts block-level content
 * (ECMA-376 §17.13.4.2), so comments can span multiple paragraphs, each
 * with its own paragraph properties (typically the CommentText style).
 */
export interface CommentParagraph {
  /** Raw <w:pPr> XML preserved for round-trip */
  pPr?: string;
  /** Paragraph content in document order */
  content: CommentContentItem[];
}

/**
 * Comment properties
 */
export interface CommentProperties {
  /** Unique comment ID (assigned by CommentManager) */
  id?: number;
  /** Author who created the comment */
  author: string;
  /** Author's initials (optional) */
  initials?: string;
  /** Date when the comment was created */
  date?: Date;
  /** Comment content (text or runs) */
  content: string | Run | Run[];
  /** Parent comment ID (for replies) */
  parentId?: number;
  /** Whether the comment is resolved/done (w:done attribute per ECMA-376) */
  done?: boolean;
}

/**
 * Represents a comment/annotation in a document
 */
export class Comment {
  private id: number;
  private author: string;
  private initials: string;
  private date: Date;
  private paragraphs: CommentParagraph[];
  private parentId?: number;
  private done: boolean;
  private onModified: (() => void) | null = null;

  /**
   * Creates a new Comment
   * @param properties - Comment properties
   */
  constructor(properties: CommentProperties) {
    this.id = properties.id ?? 0; // Will be assigned by CommentManager
    this.author = properties.author;
    this.initials = properties.initials || this.generateInitials(properties.author);
    this.date = properties.date || new Date();
    this.parentId = properties.parentId;
    this.done = properties.done ?? false;

    // Convert content to runs in a single paragraph
    let runs: Run[];
    if (typeof properties.content === 'string') {
      runs = [new Run(properties.content)];
    } else if (Array.isArray(properties.content)) {
      runs = properties.content;
    } else {
      runs = [properties.content];
    }
    this.paragraphs = [{ content: runs }];
  }

  /**
   * Generates initials from author name
   * Examples: "John Doe" -> "JD", "Jane Smith" -> "JS"
   */
  private generateInitials(author: string): string {
    const words = author
      .trim()
      .split(/\s+/)
      .filter((w) => w.length > 0);
    if (words.length === 0) return 'U';
    if (words.length === 1) return (words[0] || 'U').substring(0, 2).toUpperCase();
    return (
      words
        .map((word) => word[0] || '')
        .join('')
        .toUpperCase()
        .substring(0, 3) || 'U'
    );
  }

  /**
   * Gets the comment ID
   */
  getId(): number {
    return this.id;
  }

  /**
   * Sets the comment ID (used by CommentManager)
   * @internal
   */
  setId(id: number): void {
    this.id = id;
  }

  /**
   * Registers a callback invoked whenever this comment is mutated.
   * Document wires this to its comments dirty flag so edits to parsed
   * comments trigger regeneration of comments.xml instead of being
   * silently discarded by the original-XML passthrough on save.
   * @internal
   */
  setModifiedCallback(callback: () => void): void {
    this.onModified = callback;
  }

  private notifyModified(): void {
    if (this.onModified) {
      this.onModified();
    }
  }

  /**
   * Gets the author
   */
  getAuthor(): string {
    return this.author;
  }

  /**
   * Sets the author
   */
  setAuthor(author: string): this {
    this.author = author;
    this.notifyModified();
    return this;
  }

  /**
   * Gets the author's initials
   */
  getInitials(): string {
    return this.initials;
  }

  /**
   * Sets the author's initials
   */
  setInitials(initials: string): this {
    this.initials = initials;
    this.notifyModified();
    return this;
  }

  /**
   * Gets the comment date
   */
  getDate(): Date {
    return this.date;
  }

  /**
   * Sets the comment date
   */
  setDate(date: Date): this {
    this.date = date;
    this.notifyModified();
    return this;
  }

  /**
   * Gets the parent comment ID (for replies)
   */
  getParentId(): number | undefined {
    return this.parentId;
  }

  /**
   * Checks if this is a reply to another comment
   */
  isReply(): boolean {
    return this.parentId !== undefined;
  }

  /**
   * Checks if the comment is resolved/done
   * Per ECMA-376 Part 1, Section 17.13.4.2 (w:done attribute)
   */
  isResolved(): boolean {
    return this.done;
  }

  /**
   * Marks the comment as resolved/done
   * Persists as w15:done="1" in word/commentsExtended.xml on save
   */
  resolve(): this {
    this.done = true;
    this.notifyModified();
    return this;
  }

  /**
   * Marks the comment as unresolved
   * Drops the w15:done flag from word/commentsExtended.xml on save
   */
  unresolve(): this {
    this.done = false;
    this.notifyModified();
    return this;
  }

  /**
   * Gets the runs in this comment (flattened across paragraphs,
   * including runs inside preserved inline content)
   */
  getRuns(): Run[] {
    const runs: Run[] = [];
    for (const paragraph of this.paragraphs) {
      for (const item of paragraph.content) {
        if (item instanceof Run) {
          runs.push(item);
        } else {
          runs.push(...item.runs);
        }
      }
    }
    return runs;
  }

  /**
   * Adds a run to this comment (appended to the last paragraph)
   */
  addRun(run: Run): this {
    let last = this.paragraphs[this.paragraphs.length - 1];
    if (!last) {
      last = { content: [] };
      this.paragraphs.push(last);
    }
    last.content.push(run);
    this.notifyModified();
    return this;
  }

  /**
   * Gets the comment content grouped by paragraph
   */
  getParagraphs(): CommentParagraph[] {
    return this.paragraphs.map((paragraph) => ({
      pPr: paragraph.pPr,
      content: [...paragraph.content],
    }));
  }

  /**
   * Replaces the comment content with parsed paragraphs.
   * Used by DocumentParser to preserve <w:p> boundaries, paragraph
   * properties, and hyperlink wrappers of comments loaded from an
   * existing comments.xml. Not a user mutation — no dirty notification.
   * @internal
   */
  setParagraphs(paragraphs: CommentParagraph[]): void {
    this.paragraphs = paragraphs;
  }

  /**
   * Gets the comment text (combines all runs)
   */
  getText(): string {
    return this.getRuns()
      .map((run) => run.getText())
      .join('');
  }

  /**
   * Gets the comment content (alias for getText for compatibility)
   */
  getContent(): string {
    return this.getText();
  }

  /**
   * Formats a date to ISO 8601 format for XML
   * Uses formatDateForXml() to strip milliseconds which Word does not accept.
   */
  private formatDate(date: Date): string {
    return formatDateForXml(date);
  }

  /**
   * Generates XML for the comment range start marker
   * This goes in word/document.xml at the start of the commented range
   */
  toRangeStartXML(): XMLElement {
    return {
      name: 'w:commentRangeStart',
      attributes: {
        'w:id': this.id.toString(),
      },
      selfClosing: true,
    };
  }

  /**
   * Generates XML for the comment range end marker
   * This goes in word/document.xml at the end of the commented range
   */
  toRangeEndXML(): XMLElement {
    return {
      name: 'w:commentRangeEnd',
      attributes: {
        'w:id': this.id.toString(),
      },
      selfClosing: true,
    };
  }

  /**
   * Generates XML for the comment reference
   * This goes in word/document.xml after the range end, inside a run
   */
  toReferenceXML(): XMLElement {
    return {
      name: 'w:r',
      children: [
        {
          name: 'w:commentReference',
          attributes: {
            'w:id': this.id.toString(),
          },
          selfClosing: true,
        },
      ],
    };
  }

  /**
   * Generates XML for the comment definition
   * This goes in word/comments.xml
   */
  toXML(): XMLElement {
    const attributes: Record<string, string> = {
      'w:id': this.id.toString(),
      'w:author': this.author,
      'w:date': this.formatDate(this.date),
      'w:initials': this.initials,
    };

    // Add parent ID for replies (w15 namespace per Office 2012 extensions)
    if (this.parentId !== undefined) {
      attributes['w15:parentId'] = this.parentId.toString();
    }

    // Add done attribute for resolved comments (per ECMA-376)
    if (this.done) {
      attributes['w:done'] = '1';
    }

    // One <w:p> per stored paragraph — collapsing into a single paragraph
    // concatenates words across the lost boundaries and drops the
    // CommentText paragraph style of parsed comments
    const commentParagraphs: XMLElement[] = this.paragraphs.map((paragraph) => {
      const children: XMLElement[] = [];
      if (paragraph.pPr) {
        children.push({ name: '__rawXml', rawXml: paragraph.pPr });
      }
      for (const item of paragraph.content) {
        if (item instanceof Run) {
          children.push(item.toXML());
        } else {
          children.push({ name: '__rawXml', rawXml: item.rawXml });
        }
      }
      return { name: 'w:p', children };
    });

    return {
      name: 'w:comment',
      attributes,
      children: commentParagraphs,
    };
  }

  /**
   * Creates a new comment
   * @param author - Comment author
   * @param content - Comment content (text or runs)
   * @param initials - Optional author initials
   * @returns New Comment instance
   */
  static create(author: string, content: string | Run | Run[], initials?: string): Comment {
    return new Comment({
      author,
      content,
      initials,
    });
  }

  /**
   * Creates a reply to an existing comment
   * @param parentId - ID of the parent comment
   * @param author - Reply author
   * @param content - Reply content (text or runs)
   * @param initials - Optional author initials
   * @returns New Comment instance (reply)
   */
  static createReply(
    parentId: number,
    author: string,
    content: string | Run | Run[],
    initials?: string
  ): Comment {
    return new Comment({
      author,
      content,
      initials,
      parentId,
    });
  }

  /**
   * Creates a comment with formatted content
   * @param author - Comment author
   * @param runs - Array of formatted runs
   * @param initials - Optional author initials
   * @returns New Comment instance
   */
  static createFormatted(author: string, runs: Run[], initials?: string): Comment {
    return new Comment({
      author,
      content: runs,
      initials,
    });
  }
}
