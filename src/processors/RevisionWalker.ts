/**
 * RevisionWalker - DOM-based tree walker for accepting tracked changes
 *
 * Replaces the fragile RegEx-based revision acceptance with a robust
 * DOM-based approach that properly handles nested elements and preserves
 * element ordering.
 *
 * @module RevisionWalker
 */

import { ParsedXMLObject } from '../xml/XMLParser.js';

/**
 * Options for controlling which revision types to process
 */
export interface RevisionWalkerOptions {
  /**
   * Direction of processing (default: 'accept').
   *
   * - 'accept': keep inserted content, discard deleted content (the post-edit
   *   document).
   * - 'reject': discard inserted content, restore deleted content and previous
   *   formatting (the original pre-edit document — the exact inverse of accept).
   */
  mode?: 'accept' | 'reject';
  /**
   * Process insertions. In 'accept' mode w:ins is unwrapped (content kept); in
   * 'reject' mode w:ins is removed (content discarded). Inserted rows/tables are
   * removed under 'reject'. (default: true)
   */
  acceptInsertions?: boolean;
  /**
   * Process deletions. In 'accept' mode w:del is removed (content discarded); in
   * 'reject' mode w:del is unwrapped and w:delText restored to w:t. Deleted
   * rows/tables are removed under 'accept'. (default: true)
   */
  acceptDeletions?: boolean;
  /** Handle w:moveFrom/w:moveTo (default: true) */
  acceptMoves?: boolean;
  /**
   * Process *Change elements. In 'accept' mode they are removed (current
   * formatting kept); in 'reject' mode the previous formatting stored inside the
   * change is restored. (default: true)
   */
  acceptPropertyChanges?: boolean;
}

/** Action to take on a revision element at a given level. */
type RevisionAction = 'unwrap' | 'remove' | 'restore' | 'none';

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
 * Children that are tracked independently of a property-change snapshot and
 * must be carried over (not wiped) when rejecting that change restores the
 * previous properties. Keyed by the wrapper element; `position` is where they
 * sit in the wrapper's ECMA-376 child order relative to the restored snapshot:
 *
 * - CT_PPr (§17.3.1.26): the paragraph-mark run properties (`w:rPr`) and the
 *   section properties (`w:sectPr`) follow the base paragraph properties, and a
 *   `w:pPrChange` snapshot is a CT_PPrBase that contains neither.
 * - CT_SectPr (§17.6.17): the header/footer references PRECEDE the section base
 *   contents, and a `w:sectPrChange` snapshot is a CT_SectPrBase that contains
 *   neither — so they must be prepended or the section loses its headers/footers.
 */
const PRESERVED_ON_RESTORE: Record<string, { keys: string[]; position: 'before' | 'after' }> = {
  'w:pPr': { keys: ['w:rPr', 'w:sectPr'], position: 'after' },
  'w:sectPr': { keys: ['w:headerReference', 'w:footerReference'], position: 'before' },
};

/**
 * Deleted run-content elements and the live elements they restore to when a
 * deletion is rejected (§17.3.3.7 w:delText is structurally CT_Text, identical
 * to w:t; w:delInstrText likewise mirrors w:instrText).
 */
const DELETED_TEXT_RENAMES: Record<string, string> = {
  'w:delText': 'w:t',
  'w:delInstrText': 'w:instrText',
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
  static processTree(obj: ParsedXMLObject, options?: RevisionWalkerOptions): ParsedXMLObject {
    const opts: Required<RevisionWalkerOptions> = {
      mode: options?.mode ?? 'accept',
      acceptInsertions: options?.acceptInsertions ?? true,
      acceptDeletions: options?.acceptDeletions ?? true,
      acceptMoves: options?.acceptMoves ?? true,
      acceptPropertyChanges: options?.acceptPropertyChanges ?? true,
    };

    // Deep clone the object to avoid mutating the original
    const clone = RevisionWalker.deepClone(obj);

    // Walk and transform the tree
    RevisionWalker.walkAndTransform(clone, opts);

    // Reject restores deleted content by unwrapping w:del; the runs inside a
    // deletion carry their text in w:delText (and field codes in
    // w:delInstrText). Once unwrapped they are live content again, so convert
    // those back to the normal w:t / w:instrText elements. (§17.3.3.7)
    if (opts.mode === 'reject' && opts.acceptDeletions) {
      RevisionWalker.convertDeletedTextToNormal(clone);
    }

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
  private static walkAndTransform(obj: any, options: Required<RevisionWalkerOptions>): void {
    if (obj === null || typeof obj !== 'object') {
      return;
    }

    // Pre-recursion: row-level tracked deletions (CT_TrPr > w:del per
    // §17.13.5.14). Must run BEFORE the depth-first recursion — otherwise
    // processRevisions in trPr would strip the w:del marker before we can
    // detect it at the table level. The default `w:del` removal in
    // processRevisions only strips the marker itself, leaving a zombie
    // empty row; per spec, accepting the deletion removes the entire row.
    // Tables are checked from the parent first: a table whose rows are ALL
    // tracked-deleted must be removed entirely (a row-less w:tbl violates
    // §17.4.38's at-least-one-row requirement and Word's Accept All drops
    // the whole table), and only the parent can remove the w:tbl itself.
    // Reject is the mirror image: a row whose w:trPr carries a self-closing
    // w:ins marker was added by the edit and must be removed to revert to the
    // original (gated on acceptInsertions, since this undoes an insertion). A
    // table whose rows are ALL insert-marked is removed entirely, mirroring the
    // accept-side at-least-one-row handling.
    if (options.mode === 'reject') {
      if (options.acceptInsertions && obj['w:tbl']) {
        RevisionWalker.removeFullyMarkedTables(obj, 'w:ins');
      }
      if (options.acceptInsertions && obj['w:tr']) {
        RevisionWalker.filterMarkedRows(obj, 'w:ins');
      }
    } else {
      if (options.acceptDeletions && obj['w:tbl']) {
        RevisionWalker.removeFullyMarkedTables(obj, 'w:del');
      }
      if (options.acceptDeletions && obj['w:tr']) {
        RevisionWalker.filterMarkedRows(obj, 'w:del');
      }
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
   * Remove `<w:tr>` entries whose `<w:trPr>` contains a self-closing row-level
   * tracking marker. `marker` is `w:del` when accepting (drop deleted rows) or
   * `w:ins` when rejecting (drop inserted rows). Operates on the parsed-object
   * representation (arrays when multiple, single object otherwise).
   */
  private static filterMarkedRows(tbl: any, marker: 'w:del' | 'w:ins'): void {
    const rows = tbl['w:tr'];
    const isRowMarked = (row: any): boolean => {
      if (!row || typeof row !== 'object') return false;
      const trPr = row['w:trPr'];
      if (!trPr || typeof trPr !== 'object') return false;
      // The row-level marker is a CT_TrackChange (no children). When absent
      // the parser leaves no key; when present it's a self-closing tag
      // parsed as an object with only @_w:id/@_w:author/@_w:date keys.
      return !!trPr[marker];
    };

    if (Array.isArray(rows)) {
      const removedIndices = new Set<number>();
      for (let i = 0; i < rows.length; i++) {
        if (isRowMarked(rows[i])) removedIndices.add(i);
      }
      if (removedIndices.size === 0) return;
      const kept = rows.filter((_, i) => !removedIndices.has(i));
      if (kept.length === 0) {
        delete tbl['w:tr'];
      } else {
        tbl['w:tr'] = kept;
      }
      RevisionWalker.dropOrderedChildEntries(tbl, 'w:tr', removedIndices);
    } else if (isRowMarked(rows)) {
      delete tbl['w:tr'];
      RevisionWalker.dropOrderedChildEntries(tbl, 'w:tr', new Set([0]));
    }
  }

  /**
   * Remove `<w:tbl>` children whose rows are ALL marked with the given
   * row-level tracking marker (`w:del` when accepting, `w:ins` when rejecting).
   *
   * Runs at the table's parent because filterMarkedRows (invoked on the table
   * itself) has no reference back to the container holding the `w:tbl` key.
   * Skips tables that still carry potential row containers — the wrappers that
   * SURVIVE this direction (w:ins/w:moveTo when accepting deletions,
   * w:del/w:moveFrom when rejecting insertions), plus w:sdt / w:customXml —
   * since their rows only materialize as direct `w:tr` children later, when the
   * wrappers are unwrapped during recursion.
   */
  private static removeFullyMarkedTables(parent: any, marker: 'w:del' | 'w:ins'): void {
    const tables = parent['w:tbl'];
    // Wrappers whose rows survive in this direction must not be treated as
    // "rowless": accept keeps w:ins/w:moveTo content, reject keeps w:del/
    // w:moveFrom content.
    const survivingWrappers: string[] =
      marker === 'w:del' ? ['w:ins', 'w:moveTo'] : ['w:del', 'w:moveFrom'];
    const becomesRowless = (tbl: any): boolean => {
      // Tables with no rows at all are pre-existing structures, not the
      // result of a tracked change — leave them untouched.
      if (!tbl || typeof tbl !== 'object' || tbl['w:tr'] === undefined) return false;
      RevisionWalker.filterMarkedRows(tbl, marker);
      return (
        tbl['w:tr'] === undefined &&
        !survivingWrappers.some((w) => tbl[w]) &&
        !tbl['w:sdt'] &&
        !tbl['w:customXml']
      );
    };

    if (Array.isArray(tables)) {
      const removedIndices = new Set<number>();
      for (let i = 0; i < tables.length; i++) {
        if (becomesRowless(tables[i])) removedIndices.add(i);
      }
      if (removedIndices.size === 0) return;
      const kept = tables.filter((_: any, i: number) => !removedIndices.has(i));
      if (kept.length === 0) {
        delete parent['w:tbl'];
      } else {
        parent['w:tbl'] = kept;
      }
      RevisionWalker.dropOrderedChildEntries(parent, 'w:tbl', removedIndices);
    } else if (becomesRowless(tables)) {
      delete parent['w:tbl'];
      RevisionWalker.dropOrderedChildEntries(parent, 'w:tbl', new Set([0]));
    }
  }

  /**
   * Drop specific {type, index} entries from `_orderedChildren` and
   * re-index the survivors. The serializer maps each entry onto the
   * element array by index, so stale entries left after filtering would
   * shift surviving elements relative to inter-element siblings (e.g. a
   * w:bookmarkEnd between table rows) or silently drop the tail.
   */
  private static dropOrderedChildEntries(
    parent: any,
    type: string,
    removedIndices: Set<number>
  ): void {
    if (!parent._orderedChildren) return;
    parent._orderedChildren = parent._orderedChildren.filter(
      (c: OrderedChildInfo) => c.type !== type || !removedIndices.has(c.index)
    );
    RevisionWalker.reindexOrderedChildren(parent._orderedChildren);
  }

  /**
   * Process revision elements at the current level
   */
  private static processRevisions(parent: any, options: Required<RevisionWalkerOptions>): void {
    if (!parent || typeof parent !== 'object') {
      return;
    }

    const keysToProcess = Object.keys(parent).filter(
      (k) => !k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren'
    );

    for (const key of keysToProcess) {
      switch (RevisionWalker.getActionForKey(key, options)) {
        case 'unwrap':
          RevisionWalker.unwrapAllElements(parent, key);
          break;
        case 'remove':
          RevisionWalker.removeAllElements(parent, key);
          break;
        case 'restore':
          RevisionWalker.restorePropertyChange(parent, key);
          break;
        case 'none':
          break;
      }
    }
  }

  /**
   * Resolve what to do with a revision element of the given key, honouring the
   * processing direction. Accept and reject are exact inverses for content
   * revisions; property changes are removed on accept and restored on reject.
   */
  private static getActionForKey(
    key: string,
    options: Required<RevisionWalkerOptions>
  ): RevisionAction {
    const reject = options.mode === 'reject';

    // Content revisions are driven off REVISION_ELEMENTS so membership stays a
    // single source of truth (shared with isRevisionElement). On accept, UNWRAP
    // keeps content and REMOVE discards it; reject is the exact inverse.
    const isUnwrap = REVISION_ELEMENTS.UNWRAP.includes(key);
    const isRemove = REVISION_ELEMENTS.REMOVE.includes(key);
    if (isUnwrap || isRemove) {
      // Per-key gating: w:ins ↔ insertions, w:del ↔ deletions, w:moveTo/
      // w:moveFrom ↔ moves.
      const gate =
        key === 'w:ins'
          ? options.acceptInsertions
          : key === 'w:del'
            ? options.acceptDeletions
            : options.acceptMoves;
      if (!gate) return 'none';
      const keepContent = reject ? isRemove : isUnwrap;
      return keepContent ? 'unwrap' : 'remove';
    }
    if (REVISION_ELEMENTS.PROPERTY_CHANGES.includes(key)) {
      if (!options.acceptPropertyChanges) return 'none';
      // w:numberingChange (CT_TrackChangeNumbering) is a legacy attribute-only
      // marker with no embedded previous-properties element, so there is
      // nothing to restore — drop it in both directions.
      if (reject && key !== 'w:numberingChange') return 'restore';
      return 'remove';
    }
    if (REVISION_ELEMENTS.RANGE_MARKERS.includes(key)) {
      return 'remove'; // Always remove range markers
    }
    return 'none';
  }

  /**
   * Unwrap all elements of a given type, promoting their children to parent
   *
   * This is the most complex operation in RevisionWalker. When we unwrap a
   * revision element like w:ins, we need to:
   * 1. Extract the children from inside w:ins
   * 2. Insert them at the correct position in the parent's element arrays
   * 3. Update _orderedChildren to reflect the new structure
   *
   * The key challenge is maintaining correct element order. For example:
   *   Before: <w:tbl><w:tr>Row1</w:tr><w:ins><w:tr>Row2</w:tr></w:ins><w:tr>Row3</w:tr></w:tbl>
   *   After:  <w:tbl><w:tr>Row1</w:tr><w:tr>Row2</w:tr><w:tr>Row3</w:tr></w:tbl>
   */
  private static unwrapAllElements(parent: any, key: string): void {
    const elements = parent[key];
    if (!elements) return;

    const elementArray = Array.isArray(elements) ? elements : [elements];

    // Build a complete ordered list of child elements by walking _orderedChildren
    // This captures the intended order before we modify anything
    const orderedElements: { type: string; element: any }[] = [];

    if (parent._orderedChildren) {
      // Track how many of each type we've seen to get correct array index
      const typeCounters = new Map<string, number>();
      let unwrappedIndex = 0;

      for (const entry of parent._orderedChildren) {
        const { type } = entry;

        if (type === key) {
          // This is a revision element - extract its children in order
          const revElement = elementArray[unwrappedIndex];
          unwrappedIndex++;

          if (revElement && typeof revElement === 'object') {
            if (revElement._orderedChildren) {
              // Use the revision element's _orderedChildren for correct order
              const childCounters = new Map<string, number>();
              for (const childEntry of revElement._orderedChildren) {
                const childType = childEntry.type;
                const childIdx = childCounters.get(childType) || 0;
                childCounters.set(childType, childIdx + 1);

                const childElements = revElement[childType];
                if (childElements) {
                  const childArray = Array.isArray(childElements) ? childElements : [childElements];
                  if (childIdx < childArray.length) {
                    orderedElements.push({
                      type: childType,
                      element: childArray[childIdx],
                    });
                  }
                }
              }
            } else {
              // No _orderedChildren, extract children in object key order
              const childKeys = Object.keys(revElement).filter(
                (k) => !k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren'
              );
              for (const childKey of childKeys) {
                const childValue = revElement[childKey];
                const childArray = Array.isArray(childValue) ? childValue : [childValue];
                for (const child of childArray) {
                  orderedElements.push({ type: childKey, element: child });
                }
              }
            }
          }
        } else {
          // Regular element - get it from the parent
          const idx = typeCounters.get(type) || 0;
          typeCounters.set(type, idx + 1);

          const parentElements = parent[type];
          if (parentElements) {
            const parentArray = Array.isArray(parentElements) ? parentElements : [parentElements];
            if (idx < parentArray.length) {
              orderedElements.push({ type, element: parentArray[idx] });
            }
          }
        }
      }
    }

    // Remove the revision wrapper
    delete parent[key];

    // If we have ordered elements, rebuild the arrays in correct order
    if (orderedElements.length > 0) {
      // Group elements by type
      const rebuiltArrays = new Map<string, any[]>();

      for (const { type, element } of orderedElements) {
        if (!rebuiltArrays.has(type)) {
          rebuiltArrays.set(type, []);
        }
        rebuiltArrays.get(type)!.push(element);
      }

      // Update parent with rebuilt arrays
      for (const [type, elements] of rebuiltArrays) {
        if (elements.length === 1) {
          parent[type] = elements[0];
        } else {
          parent[type] = elements;
        }
      }

      // Rebuild _orderedChildren
      const newOrderedChildren: OrderedChildInfo[] = [];
      const typeCounters = new Map<string, number>();

      for (const { type } of orderedElements) {
        const idx = typeCounters.get(type) || 0;
        typeCounters.set(type, idx + 1);
        newOrderedChildren.push({ type, index: idx });
      }

      parent._orderedChildren = newOrderedChildren;
    } else {
      // Fallback for when there's no _orderedChildren - just merge
      for (const element of elementArray) {
        if (!element || typeof element !== 'object') continue;

        const elementKeys = Object.keys(element).filter(
          (k) => !k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren'
        );

        for (const childKey of elementKeys) {
          RevisionWalker.mergeIntoParent(parent, childKey, element[childKey]);
        }
      }
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
   * Restore the previous formatting recorded by a property-change element
   * (reject direction). A change such as `w:rPrChange` lives inside the
   * properties wrapper it tracks (`w:rPr`) and embeds the complete previous
   * wrapper as a child:
   *
   *   <w:rPr><w:b/><w:rPrChange><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>
   *
   * Rejecting it replaces the current direct formatting (`w:b`) with the
   * embedded previous formatting (`w:i`) and drops the change marker. Children
   * that are tracked independently of the snapshot — e.g. a paragraph's mark
   * run properties / section properties, or a section's header/footer
   * references (see {@link PRESERVED_ON_RESTORE}) — are carried over from the
   * current wrapper in their correct schema position rather than discarded.
   *
   * @param parent - the properties wrapper containing the change element
   * @param changeKey - the change element key, e.g. 'w:rPrChange'
   */
  private static restorePropertyChange(parent: any, changeKey: string): void {
    const changeRaw = parent[changeKey];
    if (!changeRaw) return;
    const changeEl = Array.isArray(changeRaw) ? changeRaw[0] : changeRaw;

    // 'w:rPrChange' -> 'w:rPr', 'w:tblGridChange' -> 'w:tblGrid', etc.
    const previousKey = changeKey.replace(/Change$/, '');
    const previousRaw =
      changeEl && typeof changeEl === 'object' ? changeEl[previousKey] : undefined;

    // Malformed change with no embedded previous-properties snapshot: there is
    // nothing to restore, so drop just the marker and leave the current
    // formatting intact rather than emptying the wrapper.
    if (previousRaw === undefined) {
      RevisionWalker.removeAllElements(parent, changeKey);
      return;
    }
    const previousEl = Array.isArray(previousRaw) ? previousRaw[0] : previousRaw;

    const preserve = PRESERVED_ON_RESTORE[previousKey];
    const preserveKeys = preserve?.keys ?? [];

    // Snapshot children become the restored formatting (minus any preserved
    // keys it might redundantly carry).
    const snapshotChildren =
      previousEl && typeof previousEl === 'object'
        ? RevisionWalker.getOrderedChildList(previousEl).filter(
            (c) => !preserveKeys.includes(c.type)
          )
        : [];

    // Preserved children are read from the CURRENT wrapper in their existing
    // document order (captured before the wrapper is cleared below).
    const preservedChildren = preserveKeys.length
      ? RevisionWalker.getOrderedChildList(parent).filter(
          (c) => preserveKeys.includes(c.type) && c.type !== changeKey
        )
      : [];

    const newChildren =
      preserve?.position === 'before'
        ? [...preservedChildren, ...snapshotChildren]
        : [...snapshotChildren, ...preservedChildren];

    RevisionWalker.rebuildChildren(parent, newChildren);
  }

  /**
   * Replace all non-metadata children of `parent` with the supplied ordered
   * `{type, element}` list, regrouping per-type arrays and rebuilding
   * `_orderedChildren` so the serializer emits them in the given order.
   * Attributes (`@_*`) and `#text` are left untouched.
   */
  private static rebuildChildren(parent: any, ordered: { type: string; element: any }[]): void {
    for (const k of Object.keys(parent)) {
      if (!k.startsWith('@_') && k !== '#text' && k !== '_orderedChildren') {
        delete parent[k];
      }
    }

    const grouped = new Map<string, any[]>();
    for (const { type, element } of ordered) {
      if (!grouped.has(type)) grouped.set(type, []);
      grouped.get(type)!.push(element);
    }
    for (const [type, elements] of grouped) {
      parent[type] = elements.length === 1 ? elements[0] : elements;
    }

    const orderedChildren: OrderedChildInfo[] = [];
    const counters = new Map<string, number>();
    for (const { type } of ordered) {
      const idx = counters.get(type) || 0;
      counters.set(type, idx + 1);
      orderedChildren.push({ type, index: idx });
    }
    if (orderedChildren.length > 0) {
      parent._orderedChildren = orderedChildren;
    } else {
      delete parent._orderedChildren;
    }
  }

  /**
   * Flatten an element's children into an ordered `{type, element}` list,
   * honouring `_orderedChildren` when present and falling back to key order.
   */
  private static getOrderedChildList(el: any): { type: string; element: any }[] {
    const result: { type: string; element: any }[] = [];
    const ordered = el._orderedChildren as OrderedChildInfo[] | undefined;

    if (Array.isArray(ordered) && ordered.length > 0) {
      const counters = new Map<string, number>();
      for (const entry of ordered) {
        const { type } = entry;
        const idx = counters.get(type) || 0;
        counters.set(type, idx + 1);
        const value = el[type];
        if (value === undefined) continue;
        const arr = Array.isArray(value) ? value : [value];
        if (idx < arr.length) result.push({ type, element: arr[idx] });
      }
      return result;
    }

    for (const key of Object.keys(el)) {
      if (key.startsWith('@_') || key === '#text' || key === '_orderedChildren') continue;
      const value = el[key];
      const arr = Array.isArray(value) ? value : [value];
      for (const element of arr) result.push({ type: key, element });
    }
    return result;
  }

  /**
   * Convert restored deleted text back to live text throughout the tree
   * (reject direction). After w:del unwrapping, the runs that were deleted
   * still hold their text in w:delText / w:delInstrText; rename those to the
   * normal w:t / w:instrText so the content is real again, fixing any
   * `_orderedChildren` type references as we go.
   */
  private static convertDeletedTextToNormal(obj: any): void {
    if (obj === null || typeof obj !== 'object') return;

    if (Array.isArray(obj)) {
      for (const item of obj) RevisionWalker.convertDeletedTextToNormal(item);
      return;
    }

    // When this node carries any deleted-text element, rebuild its children
    // from the document-ordered list with the deleted-text keys renamed. Going
    // through the ordered list keeps text in its original position even if a
    // run holds both a w:delText and a live w:t (CT_R permits interleaving).
    const hasDeletedText = Object.keys(obj).some((k) => k in DELETED_TEXT_RENAMES);
    if (hasDeletedText) {
      const ordered = RevisionWalker.getOrderedChildList(obj).map((child) => ({
        type: DELETED_TEXT_RENAMES[child.type] ?? child.type,
        element: child.element,
      }));
      RevisionWalker.rebuildChildren(obj, ordered);
    }

    for (const key of Object.keys(obj)) {
      if (key.startsWith('@_') || key === '#text' || key === '_orderedChildren') continue;
      RevisionWalker.convertDeletedTextToNormal(obj[key]);
    }
  }

  /**
   * Merge a child value into the parent, handling arrays properly
   */
  private static mergeIntoParent(parent: any, childKey: string, childValue: any): void {
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
   * Re-index _orderedChildren to ensure indices are sequential per type
   */
  private static reindexOrderedChildren(orderedChildren: OrderedChildInfo[]): void {
    const typeCounters = new Map<string, number>();

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
