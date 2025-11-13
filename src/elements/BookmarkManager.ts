/**
 * BookmarkManager - Manages bookmarks in a document
 *
 * Tracks all bookmarks, assigns unique IDs, and ensures name uniqueness.
 */

import { Bookmark } from './Bookmark';

/**
 * Manages document bookmarks
 */
export class BookmarkManager {
  private bookmarks: Map<string, Bookmark> = new Map();
  private nextId: number = 0;

  /**
   * Registers a bookmark with the manager
   * Assigns a unique ID and ensures name uniqueness
   * @param bookmark - Bookmark to register
   * @returns The registered bookmark (same instance)
   * @throws Error if a bookmark with the same name already exists
   */
  register(bookmark: Bookmark): Bookmark {
    const name = bookmark.getName();

    // Check for duplicate names
    if (this.bookmarks.has(name)) {
      throw new Error(
        `Bookmark with name "${name}" already exists. Bookmark names must be unique within a document.`
      );
    }

    // Assign unique ID
    bookmark.setId(this.nextId++);

    // Store bookmark
    this.bookmarks.set(name, bookmark);

    return bookmark;
  }

  /**
   * Gets a bookmark by name
   * @param name - Bookmark name
   * @returns The bookmark, or undefined if not found
   */
  getBookmark(name: string): Bookmark | undefined {
    return this.bookmarks.get(name);
  }

  /**
   * Checks if a bookmark exists
   * @param name - Bookmark name
   * @returns True if the bookmark exists
   */
  hasBookmark(name: string): boolean {
    return this.bookmarks.has(name);
  }

  /**
   * Gets all bookmarks
   * @returns Array of all bookmarks
   */
  getAllBookmarks(): Bookmark[] {
    return Array.from(this.bookmarks.values());
  }

  /**
   * Gets the number of bookmarks
   * @returns Number of bookmarks
   */
  getCount(): number {
    return this.bookmarks.size;
  }

  /**
   * Removes a bookmark
   * @param name - Bookmark name
   * @returns True if the bookmark was removed
   */
  removeBookmark(name: string): boolean {
    return this.bookmarks.delete(name);
  }

  /**
   * Clears all bookmarks
   */
  clear(): void {
    this.bookmarks.clear();
    this.nextId = 0;
  }

  /**
   * Gets a unique bookmark name by adding a suffix if needed
   * @param baseName - Base name for the bookmark
   * @returns A unique bookmark name
   */
  getUniqueName(baseName: string): string {
    if (!this.hasBookmark(baseName)) {
      return baseName;
    }

    // Try adding numbers until we find a unique name
    let counter = 1;
    let uniqueName = `${baseName}_${counter}`;

    while (this.hasBookmark(uniqueName)) {
      counter++;
      uniqueName = `${baseName}_${counter}`;

      // Safety limit
      if (counter > 1000) {
        throw new Error(
          `Could not generate unique bookmark name from base "${baseName}"`
        );
      }
    }

    return uniqueName;
  }

  /**
   * Creates and registers a new bookmark with a unique name
   * @param name - Desired bookmark name
   * @returns The created and registered bookmark
   */
  createBookmark(name: string): Bookmark {
    const uniqueName = this.getUniqueName(name);
    const bookmark = Bookmark.create(uniqueName);
    return this.register(bookmark);
  }

  /**
   * Creates and registers a bookmark for a heading
   * Automatically generates a unique name from the heading text
   * @param headingText - The text of the heading
   * @returns The created and registered bookmark
   */
  createHeadingBookmark(headingText: string): Bookmark {
    const bookmark = Bookmark.createForHeading(headingText);
    const uniqueName = this.getUniqueName(bookmark.getName());
    bookmark.setName(uniqueName);
    return this.register(bookmark);
  }

  /**
   * Gets statistics about bookmarks
   * @returns Object with bookmark statistics
   */
  getStats(): {
    total: number;
    nextId: number;
    names: string[];
  } {
    return {
      total: this.bookmarks.size,
      nextId: this.nextId,
      names: Array.from(this.bookmarks.keys()),
    };
  }

  /**
   * Creates a new BookmarkManager
   * @returns New BookmarkManager instance
   */
  static create(): BookmarkManager {
    return new BookmarkManager();
  }
}
