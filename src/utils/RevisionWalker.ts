/**
 * RevisionWalker - DOM-based tree walker for accepting tracked changes
 *
 * Replaces the fragile RegEx-based revision acceptance with a robust
 * DOM-based approach that properly handles nested elements and preserves
 * element ordering.
 *
 * @module RevisionWalker
 */

import { ParsedXMLObject } from '../xml/XMLParser';

/**
 * Options for controlling which revision types to accept
 */
export interface RevisionWalkerOptions {
  /** Keep content, remove w:ins wrapper (default: true) */
  acceptInsertions?: boolean;
  /** Remove w:del and content entirely (default: true) */
  acceptDeletions?: boolean;
  /** Handle w:moveFrom/w:moveTo (default: true) */
  acceptMoves?: boolean;
  /** Remove *Change elements (default: true) */
  acceptPropertyChanges?: boolean;
}

/**
 * Structure for tracking element order
 */
interface OrderedChildInfo {
  type: string;
  index: number;
}

/**
 * Revision element categories
 */
const REVISION_ELEMENTS = {
  /** Elements to unwrap (keep content, remove wrapper) */
  UNWRAP: ['w:ins', 'w:moveTo'],

  /** Elements to remove entirely (with content) */
  REMOVE: ['w:del', 'w:moveFrom'],

  /** Property change tracking elements */
  PROPERTY_CHANGES: [
    'w:rPrChange',
    'w:pPrChange',
    'w:tblPrChange',
    'w:tcPrChange',
    'w:trPrChange',
    'w:sectPrChange',
    'w:tblGridChange',
    'w:numberingChange',
    'w:tblPrExChange',
  ],

  /** Range marker elements */
  RANGE_MARKERS: [
    'w:moveFromRangeStart',
    'w:moveFromRangeEnd',
    'w:moveToRangeStart',
    'w:moveToRangeEnd',
    'w:customXmlInsRangeStart',
    'w:customXmlInsRangeEnd',
    'w:customXmlDelRangeStart',
    'w:customXmlDelRangeEnd',
    'w:customXmlMoveFromRangeStart',
    'w:customXmlMoveFromRangeEnd',
    'w:customXmlMoveToRangeStart',
    'w:customXmlMoveToRangeEnd',
  ],
};

/**
 * DOM-based tree walker for accepting Word document revisions
 *
 * This class processes a parsed XML object tree (from XMLParser.parseToObject())
 * and accepts all tracked changes by:
 * - Unwrapping insertions (w:ins, w:moveTo) - keeping content
 * - Removing deletions (w:del, w:moveFrom) - discarding content
 * - Removing property changes (*Change elements)
 * - Removing range markers
 *
 * Element order is preserved using the _orderedChildren metadata.
 */
export class RevisionWalker {
  /**
   * Process a parsed XML object tree and accept all revisions
   *
   * @param obj - Parsed XML object from XMLParser.parseToObject()
   * @param options - Options controlling which revisions to accept
   * @returns New object tree with revisions accepted
   *
   * @example
   * ```typescript
   * const parsed = XMLParser.parseToObject(documentXml);
   * const clean = RevisionWalker.processTree(parsed);
   * ```
   */
  static processTree(
    obj: ParsedXMLObject,
    options?: RevisionWalkerOptions
  ): ParsedXMLObject {
    const opts: Required<RevisionWalkerOptions> = {
      acceptInsertions: options?.acceptInsertions ?? true,
      acceptDeletions: options?.acceptDeletions ?? true,
      acceptMoves: options?.acceptMoves ?? true,
      acceptPropertyChanges: options?.acceptPropertyChanges ?? true,
    };

    // Deep clone the object to avoid mutating the original
    const clone = RevisionWalker.deepClone(obj);

    // Walk and transform the tree
    RevisionWalker.walkAndTransform(clone, opts);

    return clone;
  }

  /**
   * Deep clone an object
   */
  private static deepClone(obj: any): any {
    if (obj === null || typeof obj !== 'object') {
      return obj;
    }

    if (Array.isArray(obj)) {
      return obj.map((item) => RevisionWalker.deepClone(item));
    }

    const clone: any = {};
    for (const key of Object.keys(obj)) {
      clone[key] = RevisionWalker.deepClone(obj[key]);
    }
    return clone;
  }

  /**
   * Recursively walk and transform the object tree
   * Processes children first (depth-first) to handle nested revisions
   */
  private static walkAndTransform(
    obj: any,
    options: Required<RevisionWalkerOptions>
  ): void {
    if (obj === null || typeof obj !== 'object') {
      return;
    }

    // Get keys to process (excluding metadata keys)
    const keys = Object.keys(obj).filter(
      (k) => !k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren'
    );

    // First pass: recurse into children (depth-first)
    for (const key of keys) {
      const value = obj[key];
      if (Array.isArray(value)) {
        for (const item of value) {
          RevisionWalker.walkAndTransform(item, options);
        }
      } else if (typeof value === 'object' && value !== null) {
        RevisionWalker.walkAndTransform(value, options);
      }
    }

    // Second pass: process revision elements at this level
    // We need to iterate carefully because we're modifying the object
    RevisionWalker.processRevisions(obj, options);
  }

  /**
   * Process revision elements at the current level
   */
  private static processRevisions(
    parent: any,
    options: Required<RevisionWalkerOptions>
  ): void {
    if (!parent || typeof parent !== 'object') {
      return;
    }

    const keysToProcess = Object.keys(parent).filter(
      (k) => !k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren'
    );

    for (const key of keysToProcess) {
      // Check if this is a revision element
      if (RevisionWalker.shouldUnwrap(key, options)) {
        RevisionWalker.unwrapAllElements(parent, key);
      } else if (RevisionWalker.shouldRemove(key, options)) {
        RevisionWalker.removeAllElements(parent, key);
      }
    }
  }

  /**
   * Check if an element should be unwrapped (content kept)
   */
  private static shouldUnwrap(
    key: string,
    options: Required<RevisionWalkerOptions>
  ): boolean {
    if (REVISION_ELEMENTS.UNWRAP.includes(key)) {
      if (key === 'w:ins' && !options.acceptInsertions) return false;
      if (key === 'w:moveTo' && !options.acceptMoves) return false;
      return true;
    }
    return false;
  }

  /**
   * Check if an element should be removed (content discarded)
   */
  private static shouldRemove(
    key: string,
    options: Required<RevisionWalkerOptions>
  ): boolean {
    if (REVISION_ELEMENTS.REMOVE.includes(key)) {
      if (key === 'w:del' && !options.acceptDeletions) return false;
      if (key === 'w:moveFrom' && !options.acceptMoves) return false;
      return true;
    }
    if (REVISION_ELEMENTS.PROPERTY_CHANGES.includes(key)) {
      return options.acceptPropertyChanges;
    }
    if (REVISION_ELEMENTS.RANGE_MARKERS.includes(key)) {
      return true; // Always remove range markers
    }
    return false;
  }

  /**
   * Unwrap all elements of a given type, promoting their children to parent
   */
  private static unwrapAllElements(parent: any, key: string): void {
    const elements = parent[key];
    if (!elements) return;

    const elementArray = Array.isArray(elements) ? elements : [elements];

    // Collect all children from all elements to unwrap
    const allPromotedChildren: Array<{
      childKey: string;
      childValue: any;
      sourceOrderedChildren?: OrderedChildInfo[];
    }> = [];

    for (const element of elementArray) {
      if (!element || typeof element !== 'object') continue;

      // Extract children from this element
      const elementKeys = Object.keys(element).filter(
        (k) => !k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren'
      );

      for (const childKey of elementKeys) {
        allPromotedChildren.push({
          childKey,
          childValue: element[childKey],
          sourceOrderedChildren: element._orderedChildren,
        });
      }
    }

    // Update _orderedChildren before removing the element
    if (parent._orderedChildren) {
      parent._orderedChildren = RevisionWalker.updateOrderedChildrenForUnwrap(
        parent._orderedChildren,
        key,
        elementArray
      );
    }

    // Remove the revision wrapper
    delete parent[key];

    // Promote children to parent level
    for (const { childKey, childValue } of allPromotedChildren) {
      RevisionWalker.mergeIntoParent(parent, childKey, childValue);
    }
  }

  /**
   * Remove all elements of a given type (including their content)
   */
  private static removeAllElements(parent: any, key: string): void {
    if (!parent[key]) return;

    // Update _orderedChildren before removing
    if (parent._orderedChildren) {
      parent._orderedChildren = parent._orderedChildren.filter(
        (c: OrderedChildInfo) => c.type !== key
      );
      // Re-index remaining elements of same types
      RevisionWalker.reindexOrderedChildren(parent._orderedChildren);
    }

    // Remove the element
    delete parent[key];
  }

  /**
   * Merge a child value into the parent, handling arrays properly
   */
  private static mergeIntoParent(
    parent: any,
    childKey: string,
    childValue: any
  ): void {
    if (parent[childKey] === undefined) {
      // No existing value, just assign
      parent[childKey] = childValue;
    } else {
      // Existing value, need to merge
      const existing = parent[childKey];
      const incoming = Array.isArray(childValue) ? childValue : [childValue];

      if (Array.isArray(existing)) {
        parent[childKey] = [...existing, ...incoming];
      } else {
        parent[childKey] = [existing, ...incoming];
      }
    }
  }

  /**
   * Update _orderedChildren when unwrapping elements
   *
   * When we unwrap w:ins, we need to:
   * 1. Find where w:ins was in _orderedChildren
   * 2. Replace it with the children's order info
   * 3. Re-index all elements
   */
  private static updateOrderedChildrenForUnwrap(
    orderedChildren: OrderedChildInfo[],
    unwrappedType: string,
    unwrappedElements: any[]
  ): OrderedChildInfo[] {
    const result: OrderedChildInfo[] = [];
    let unwrappedIndex = 0;

    for (const entry of orderedChildren) {
      if (entry.type === unwrappedType) {
        // Get the element being unwrapped
        const element = unwrappedElements[unwrappedIndex];
        unwrappedIndex++;

        if (element && element._orderedChildren) {
          // Insert the children's ordered children here
          for (const childEntry of element._orderedChildren) {
            result.push({ type: childEntry.type, index: childEntry.index });
          }
        } else if (element && typeof element === 'object') {
          // No _orderedChildren, add children in object key order
          const childKeys = Object.keys(element).filter(
            (k) =>
              !k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren'
          );
          for (const childKey of childKeys) {
            const children = element[childKey];
            if (Array.isArray(children)) {
              for (let i = 0; i < children.length; i++) {
                result.push({ type: childKey, index: i });
              }
            } else {
              result.push({ type: childKey, index: 0 });
            }
          }
        }
      } else {
        result.push({ ...entry });
      }
    }

    // Re-index to fix indices after merging
    RevisionWalker.reindexOrderedChildren(result);

    return result;
  }

  /**
   * Re-index _orderedChildren to ensure indices are sequential per type
   */
  private static reindexOrderedChildren(
    orderedChildren: OrderedChildInfo[]
  ): void {
    const typeCounters: Map<string, number> = new Map();

    for (const entry of orderedChildren) {
      const currentIndex = typeCounters.get(entry.type) || 0;
      entry.index = currentIndex;
      typeCounters.set(entry.type, currentIndex + 1);
    }
  }

  /**
   * Check if an element is a revision-related element (any type)
   */
  static isRevisionElement(key: string): boolean {
    return (
      REVISION_ELEMENTS.UNWRAP.includes(key) ||
      REVISION_ELEMENTS.REMOVE.includes(key) ||
      REVISION_ELEMENTS.PROPERTY_CHANGES.includes(key) ||
      REVISION_ELEMENTS.RANGE_MARKERS.includes(key)
    );
  }

  /**
   * Get revision element categories (for external use/testing)
   */
  static getRevisionElementCategories(): typeof REVISION_ELEMENTS {
    return { ...REVISION_ELEMENTS };
  }
}
