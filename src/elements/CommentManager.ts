/**
 * CommentManager - Manages comments in a document
 *
 * Tracks all comments, assigns unique IDs, handles replies,
 * and generates the comments.xml file.
 *
 * Per ECMA-376, comment IDs must be unique across ALL annotation types
 * in a document. Use setIdProvider() to connect to a centralized ID allocator.
 */

import { Comment } from './Comment.js';
import { Run } from './Run.js';
import { XMLBuilder } from '../xml/XMLBuilder.js';
import { formatDateForXml } from '../utils/dateFormatting.js';

/**
 * Type for the centralized ID provider callback.
 * Returns the next available annotation ID from a shared counter.
 */
export type IdProviderCallback = () => number;

/**
 * Type for callback to notify of existing IDs (for synchronization).
 * Called when registering existing comments to keep the central counter in sync.
 */
export type IdExistsCallback = (existingId: number) => void;

/**
 * Comment entry stored by the manager
 */
interface CommentEntry {
  comment: Comment;
  /** Comments that are replies to this comment */
  replies: Comment[];
}

/**
 * Manages document comments
 */
export class CommentManager {
  private comments = new Map<number, CommentEntry>();
  private nextId = 0;
  private idProvider: IdProviderCallback | null = null;
  private idExistsNotifier: IdExistsCallback | null = null;
  private modifiedNotifier: (() => void) | null = null;

  /**
   * Sets the centralized ID provider callback.
   * When set, IDs will be allocated from the centralized DocumentIdManager
   * instead of the local nextId counter.
   *
   * @param provider - Callback that returns the next available ID
   * @param existsNotifier - Optional callback to notify when existing IDs are found
   */
  setIdProvider(provider: IdProviderCallback, existsNotifier?: IdExistsCallback): void {
    this.idProvider = provider;
    this.idExistsNotifier = existsNotifier || null;
  }

  /**
   * Sets the callback invoked when a registered comment is mutated.
   * Document wires this to its comments dirty flag so instance-level edits
   * (resolve, setAuthor, addRun, ...) force comments.xml regeneration
   * instead of being silently discarded by the original-XML passthrough.
   *
   * @param notifier - Callback invoked on any comment mutation
   */
  setModifiedNotifier(notifier: () => void): void {
    this.modifiedNotifier = notifier;
    for (const entry of this.comments.values()) {
      entry.comment.setModifiedCallback(notifier);
    }
  }

  /**
   * Registers a comment with the manager
   * Assigns a unique ID
   * @param comment - Comment to register
   * @returns The registered comment (same instance)
   */
  register(comment: Comment): Comment {
    // Assign unique ID - use centralized provider if available
    const id = this.idProvider ? this.idProvider() : this.nextId++;
    comment.setId(id);

    if (this.modifiedNotifier) {
      comment.setModifiedCallback(this.modifiedNotifier);
    }

    // Store comment
    const entry: CommentEntry = {
      comment,
      replies: [],
    };
    this.comments.set(comment.getId(), entry);

    // If this is a reply, add it to the parent's replies array
    if (comment.isReply() && comment.getParentId() !== undefined) {
      const parentEntry = this.comments.get(comment.getParentId()!);
      if (parentEntry) {
        parentEntry.replies.push(comment);
      }
    }

    return comment;
  }

  /**
   * Registers an existing comment (from parsing) with its pre-assigned ID.
   * Unlike register(), this does NOT assign a new ID - the comment must already have one.
   * Used when loading comments from an existing document.
   * @param comment - Comment with ID already set
   */
  registerExisting(comment: Comment): void {
    const id = comment.getId();

    // Notify centralized ID manager if connected
    if (this.idExistsNotifier) {
      this.idExistsNotifier(id);
    }

    if (this.modifiedNotifier) {
      comment.setModifiedCallback(this.modifiedNotifier);
    }

    // Store comment
    const entry: CommentEntry = {
      comment,
      replies: [],
    };
    this.comments.set(id, entry);

    // Update local nextId if needed
    if (id >= this.nextId) {
      this.nextId = id + 1;
    }
  }

  /**
   * Links reply comments to their parents after all comments are parsed.
   * Must be called after all comments are registered via registerExisting().
   * This builds the reply arrays for each parent comment.
   */
  linkReplies(): void {
    // Clear existing replies first for idempotency (safe to call multiple times)
    for (const entry of this.comments.values()) {
      entry.replies = [];
    }
    for (const entry of this.comments.values()) {
      const comment = entry.comment;
      if (comment.isReply() && comment.getParentId() !== undefined) {
        const parentEntry = this.comments.get(comment.getParentId()!);
        if (parentEntry) {
          parentEntry.replies.push(comment);
        }
      }
    }
  }

  /**
   * Gets a comment by ID
   * @param id - Comment ID
   * @returns The comment, or undefined if not found
   */
  getComment(id: number): Comment | undefined {
    return this.comments.get(id)?.comment;
  }

  /**
   * Gets all comments (top-level only, not replies)
   * @returns Array of all top-level comments
   */
  getAllComments(): Comment[] {
    return Array.from(this.comments.values())
      .filter((entry) => !entry.comment.isReply())
      .map((entry) => entry.comment);
  }

  /**
   * Gets all comments including replies
   * @returns Array of all comments
   */
  getAllCommentsWithReplies(): Comment[] {
    return Array.from(this.comments.values()).map((entry) => entry.comment);
  }

  /**
   * Gets replies to a comment
   * @param commentId - ID of the parent comment
   * @returns Array of reply comments
   */
  getReplies(commentId: number): Comment[] {
    const entry = this.comments.get(commentId);
    return entry ? [...entry.replies] : [];
  }

  /**
   * Checks if a comment has replies
   * @param commentId - ID of the comment
   * @returns True if the comment has replies
   */
  hasReplies(commentId: number): boolean {
    const entry = this.comments.get(commentId);
    return entry ? entry.replies.length > 0 : false;
  }

  /**
   * Gets the number of comments (including replies)
   * @returns Number of comments
   */
  getCount(): number {
    return this.comments.size;
  }

  /**
   * Gets the number of top-level comments (excluding replies)
   * @returns Number of top-level comments
   */
  getTopLevelCount(): number {
    return this.getAllComments().length;
  }

  /**
   * Gets all unique authors who have made comments
   * @returns Array of unique author names
   */
  getAuthors(): string[] {
    const authorsSet = new Set<string>();
    for (const entry of this.comments.values()) {
      authorsSet.add(entry.comment.getAuthor());
    }
    return Array.from(authorsSet);
  }

  /**
   * Gets comments by author
   * @param author - Author name to filter by
   * @returns Array of comments by the specified author
   */
  getCommentsByAuthor(author: string): Comment[] {
    return Array.from(this.comments.values())
      .map((entry) => entry.comment)
      .filter((comment) => comment.getAuthor() === author);
  }

  /**
   * Gets comments within a date range
   * @param startDate - Start of date range
   * @param endDate - End of date range
   * @returns Array of comments within the date range
   */
  getCommentsByDateRange(startDate: Date, endDate: Date): Comment[] {
    return Array.from(this.comments.values())
      .map((entry) => entry.comment)
      .filter((comment) => {
        const commentDate = comment.getDate();
        return commentDate >= startDate && commentDate <= endDate;
      });
  }

  /**
   * Removes a comment
   * Also removes all replies to that comment
   * @param id - Comment ID
   * @returns True if the comment was removed
   */
  removeComment(id: number): boolean {
    const entry = this.comments.get(id);
    if (!entry) {
      return false;
    }

    // Remove all replies first
    for (const reply of entry.replies) {
      this.comments.delete(reply.getId());
    }

    // A reply is also tracked in its parent's replies array — drop it there
    // too so getReplies()/hasReplies()/getCommentThread() agree with the map
    const parentId = entry.comment.getParentId();
    if (parentId !== undefined) {
      const parentEntry = this.comments.get(parentId);
      if (parentEntry) {
        parentEntry.replies = parentEntry.replies.filter((reply) => reply !== entry.comment);
      }
    }

    // Remove the comment itself
    return this.comments.delete(id);
  }

  /**
   * Clears all comments
   */
  clear(): void {
    this.comments.clear();
    this.nextId = 0;
  }

  /**
   * Sets the next ID to be assigned.
   * Used when loading documents to avoid ID collisions with existing comments.
   * @param id - The next ID value to use
   */
  setNextId(id: number): void {
    this.nextId = id;
  }

  /**
   * Creates and registers a new comment
   * @param author - Comment author
   * @param content - Comment content (text or runs)
   * @param initials - Optional author initials
   * @returns The created and registered comment
   */
  createComment(author: string, content: string | Run | Run[], initials?: string): Comment {
    const comment = Comment.create(author, content, initials);
    return this.register(comment);
  }

  /**
   * Creates and registers a reply to an existing comment
   * @param parentCommentId - ID of the parent comment
   * @param author - Reply author
   * @param content - Reply content (text or runs)
   * @param initials - Optional author initials
   * @returns The created and registered reply
   * @throws Error if parent comment doesn't exist
   */
  createReply(
    parentCommentId: number,
    author: string,
    content: string | Run | Run[],
    initials?: string
  ): Comment {
    // Verify parent exists
    if (!this.comments.has(parentCommentId)) {
      throw new Error(
        `Cannot create reply: parent comment with ID ${parentCommentId} does not exist`
      );
    }

    const reply = Comment.createReply(parentCommentId, author, content, initials);
    return this.register(reply);
  }

  /**
   * Checks if there are any comments
   * @returns True if there are no comments
   */
  isEmpty(): boolean {
    return this.comments.size === 0;
  }

  /**
   * Gets a comment thread (comment and all its replies)
   * @param commentId - ID of the top-level comment
   * @returns Object with the comment and its replies
   */
  getCommentThread(commentId: number): { comment: Comment; replies: Comment[] } | undefined {
    const entry = this.comments.get(commentId);
    if (!entry || entry.comment.isReply()) {
      return undefined;
    }
    return {
      comment: entry.comment,
      replies: [...entry.replies],
    };
  }

  /**
   * Searches comments by text content
   * @param searchText - Text to search for (case-insensitive)
   * @returns Array of comments containing the search text
   */
  findCommentsByText(searchText: string): Comment[] {
    const lowerSearch = searchText.toLowerCase();
    return Array.from(this.comments.values())
      .map((entry) => entry.comment)
      .filter((comment) => comment.getText().toLowerCase().includes(lowerSearch));
  }

  /**
   * Gets all resolved/done comments
   * @returns Array of resolved comments
   */
  getResolvedComments(): Comment[] {
    return Array.from(this.comments.values())
      .map((entry) => entry.comment)
      .filter((comment) => comment.isResolved());
  }

  /**
   * Gets all unresolved comments
   * @returns Array of unresolved comments
   */
  getUnresolvedComments(): Comment[] {
    return Array.from(this.comments.values())
      .map((entry) => entry.comment)
      .filter((comment) => !comment.isResolved());
  }

  /**
   * Gets the most recent comments
   * @param count - Number of recent comments to return
   * @returns Array of most recent comments
   */
  getRecentComments(count: number): Comment[] {
    const allComments = this.getAllCommentsWithReplies();
    return allComments
      .sort((a, b) => b.getDate().getTime() - a.getDate().getTime())
      .slice(0, count);
  }

  /**
   * Gets statistics about comments
   * @returns Object with comment statistics
   */
  getStats(): {
    total: number;
    topLevel: number;
    replies: number;
    resolved: number;
    unresolved: number;
    authors: string[];
    nextId: number;
  } {
    const topLevel = this.getTopLevelCount();
    const resolved = this.getResolvedComments().length;
    return {
      total: this.comments.size,
      topLevel,
      replies: this.comments.size - topLevel,
      resolved,
      unresolved: this.comments.size - resolved,
      authors: this.getAuthors(),
      nextId: this.nextId,
    };
  }

  /**
   * Generates the word/comments.xml file content
   * @returns XML string for comments.xml
   */
  generateCommentsXml(): string {
    const comments = this.getAllCommentsWithReplies();

    if (comments.length === 0) {
      // Return minimal comments.xml
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:comments>`;
    }

    // Build XML manually for comments
    const hasReplies = comments.some((c) => c.isReply());
    const hasResolved = comments.some((c) => c.isResolved());
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    if (hasResolved) {
      xml += ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
    }
    if (hasReplies) {
      xml += ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
    }
    if (hasReplies || hasResolved) {
      const ignorable = [hasResolved ? 'w14' : '', hasReplies ? 'w15' : '']
        .filter((ns) => ns.length > 0)
        .join(' ');
      xml += ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="${ignorable}"`;
    }
    xml += '>\n';

    // Add each comment. CT_Comment declares no resolved-state attribute —
    // that lives in commentsExtended.xml keyed by the w14:paraId of each
    // comment's last paragraph, so emit those anchors when any comment is
    // resolved (see generateCommentsExtendedXml)
    for (const comment of comments) {
      xml += this.commentToXmlString(
        comment,
        hasResolved ? CommentManager.commentParaId(comment.getId()) : undefined
      );
    }

    xml += '</w:comments>';
    return xml;
  }

  /**
   * Derives the w14:paraId anchor for a comment's last paragraph.
   * Deterministic so comments.xml and commentsExtended.xml agree without
   * shared state. Values must be 8 hex digits, nonzero, and below
   * 0x80000000 per ST_LongHexNumber conventions for paraId.
   * @param commentId - Comment ID
   * @returns 8-character uppercase hex paraId
   */
  static commentParaId(commentId: number): string {
    return ((commentId + 1) & 0x7fffffff).toString(16).toUpperCase().padStart(8, '0');
  }

  /**
   * Generates the word/commentsExtended.xml content carrying resolved
   * state (w15:done) and reply threading (w15:paraIdParent).
   * @returns XML string, or null when no comment is resolved (the part is
   *   only regenerated when it carries information)
   */
  generateCommentsExtendedXml(): string | null {
    const comments = this.getAllCommentsWithReplies();
    if (!comments.some((c) => c.isResolved())) {
      return null;
    }

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<w15:commentsEx xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    xml +=
      ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w15">\n';
    for (const comment of comments) {
      xml += `  <w15:commentEx w15:paraId="${CommentManager.commentParaId(comment.getId())}"`;
      const parentId = comment.getParentId();
      if (parentId !== undefined && this.comments.has(parentId)) {
        xml += ` w15:paraIdParent="${CommentManager.commentParaId(parentId)}"`;
      }
      xml += ` w15:done="${comment.isResolved() ? '1' : '0'}"/>\n`;
    }
    xml += '</w15:commentsEx>';
    return xml;
  }

  /**
   * Converts a comment to XML string
   * @param comment - Comment to convert
   * @param paraId - w14:paraId to anchor on the last paragraph (links the
   *   comment to its commentsExtended.xml entry), omitted when no comment
   *   in the document is resolved
   * @returns XML string for the comment
   */
  private commentToXmlString(comment: Comment, paraId?: string): string {
    let xml = `  <w:comment w:id="${comment.getId()}"`;
    xml += ` w:author="${XMLBuilder.escapeXmlAttribute(comment.getAuthor())}"`;
    xml += ` w:date="${formatDateForXml(comment.getDate())}"`;
    xml += ` w:initials="${XMLBuilder.escapeXmlAttribute(comment.getInitials())}"`;

    if (comment.isReply() && comment.getParentId() !== undefined) {
      xml += ` w15:parentId="${comment.getParentId()}"`;
    }

    xml += '>\n';

    // One <w:p> per stored paragraph — collapsing into a single paragraph
    // concatenates words across the lost boundaries and drops paragraph
    // properties (CommentText pStyle) of parsed comments
    const paragraphs = comment.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i]!;
      // commentsExtended.xml entries reference the LAST paragraph's paraId
      const isLast = i === paragraphs.length - 1;
      xml += isLast && paraId ? `    <w:p w14:paraId="${paraId}">\n` : '    <w:p>\n';
      if (paragraph.pPr) {
        xml += `      ${paragraph.pPr}\n`;
      }
      for (const item of paragraph.content) {
        if (item instanceof Run) {
          xml += this.runToXmlString(item, 6);
        } else {
          // Preserved inline XML (e.g. w:hyperlink) re-emitted verbatim so
          // its r:id target keeps pointing at word/_rels/comments.xml.rels
          xml += `      ${item.rawXml}\n`;
        }
      }
      xml += '    </w:p>\n';
    }

    xml += '  </w:comment>\n';
    return xml;
  }

  /**
   * Converts a run to XML string
   * @param run - Run to convert
   * @param indent - Number of spaces for indentation
   * @returns XML string for the run
   */
  private runToXmlString(run: Run, indent: number): string {
    const spaces = ' '.repeat(indent);
    // Serialize through Run.toXML() so w:rPr formatting and non-text
    // content (w:annotationRef, tabs, breaks) survive regeneration
    return `${spaces}${XMLBuilder.elementToString(run.toXML())}\n`;
  }

  /**
   * Creates a new CommentManager
   * @returns New CommentManager instance
   */
  static create(): CommentManager {
    return new CommentManager();
  }
}
