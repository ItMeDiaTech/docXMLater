/**
 * Hyperlink - Represents a hyperlink in a Word document
 *
 * Hyperlinks can be external (to websites, files) or internal (to bookmarks within the document).
 * They are represented using the `<w:hyperlink>` element.
 *
 * ## Important: Relationship ID Requirement
 *
 * **External hyperlinks REQUIRE a relationship ID to be set before XML generation.**
 * Per ECMA-376 Part 1 §17.16.22, `<w:hyperlink>` elements with external targets must have
 * an `r:id` attribute that references a relationship in `word/_rels/document.xml.rels`.
 *
 * ### Correct Usage Pattern:
 *
 * ```typescript
 * // RECOMMENDED: Use Document.save() - automatically handles relationships
 * const doc = Document.create();
 * const para = doc.createParagraph();
 * para.addHyperlink(Hyperlink.createExternal('https://example.com', 'Link'));
 * await doc.save('document.docx'); // ✅ Relationships auto-registered
 * ```
 *
 * ### Manual Relationship Registration (Advanced):
 *
 * ```typescript
 * const link = Hyperlink.createExternal('https://example.com', 'Link');
 * const relationship = relationshipManager.addHyperlink('https://example.com');
 * link.setRelationshipId(relationship.getId());
 * link.toXML(); // ✅ Now valid
 * ```
 *
 * ### What NOT to Do:
 *
 * ```typescript
 * const link = Hyperlink.createExternal('https://example.com', 'Link');
 * link.toXML(); // ❌ ERROR: Missing relationship ID
 * ```
 *
 * ## Internal Hyperlinks
 *
 * Internal hyperlinks (bookmarks) do NOT require relationships:
 *
 * ```typescript
 * const link = Hyperlink.createInternal('Section1', 'Go to Section 1');
 * link.toXML(); // ✅ Valid - uses w:anchor attribute
 * ```
 *
 * @see {@link https://www.ecma-international.org/publications-and-standards/standards/ecma-376/ | ECMA-376 Part 1 §17.16.22}
 */

import { XMLElement } from '../xml/XMLBuilder.js';
import { Run, RunFormatting } from './Run.js';
import { Revision } from './Revision.js';
import { validateRunText, sanitizeHyperlinkUrl } from '../utils/validation.js';
import { defaultLogger } from '../utils/logger.js';

/**
 * Hyperlink properties
 */
export interface HyperlinkProperties {
  /** Hyperlink URL (for external links) */
  url?: string;
  /** Bookmark anchor (for internal links) */
  anchor?: string;
  /** Display text (optional for empty/invisible hyperlinks) */
  text?: string;
  /** Text formatting */
  formatting?: RunFormatting;
  /**
   * Use the supplied formatting exactly as given instead of layering it over
   * the default hyperlink styling. Set when the formatting comes from parsed
   * XML, so load→save does not inject direct formatting that overrides the
   * document's Hyperlink character style.
   */
  preserveFormatting?: boolean;
  /** Tooltip text */
  tooltip?: string;
  /** Relationship ID (set by Document when saving) */
  relationshipId?: string;
  /** Whether this is an empty/invisible hyperlink with no display text */
  isEmpty?: boolean;
  /** Target frame attribute (e.g., "_blank" for new window) */
  tgtFrame?: string;
  /** History tracking attribute */
  history?: string;
  /** Document location for within-document navigation in external files (ECMA-376 §17.16.22) */
  docLocation?: string;
}

/**
 * Represents a hyperlink
 */
export class Hyperlink {
  private url?: string;
  private anchor?: string;
  private text: string;
  /**
   * Display-text runs. Word splits hyperlink text into multiple runs when
   * formatting varies mid-link, so a single-run model would drop per-run
   * formatting on round-trip. Always holds at least one run.
   */
  private runs: Run[];
  private tooltip?: string;
  private relationshipId?: string;
  private formatting: RunFormatting;
  /** Whether this is an empty/invisible hyperlink with no display text */
  private _isEmpty = false;
  /** Target frame attribute (e.g., "_blank" for new window) */
  private tgtFrame?: string;
  /** History tracking attribute */
  private history?: string;
  /** Document location for within-document navigation in external files */
  private docLocation?: string;
  /** Tracking context for automatic change tracking */
  private trackingContext?: import('../tracking/TrackingContext.js').TrackingContext;
  /** Parent paragraph reference for automatic tracking */
  private _parentParagraph?: import('./Paragraph.js').Paragraph;
  /** Baseline formatting captured before first tracked formatting change */
  private _preTrackingFormatting?: RunFormatting;

  /**
   * Creates a new hyperlink
   *
   * **Note:** A hyperlink must have either a URL (external) or anchor (internal), but not both.
   * If both are provided, the URL takes precedence and a warning is logged.
   *
   * @param properties Hyperlink properties
   */
  constructor(properties: HyperlinkProperties) {
    this.url = properties.url;
    this.anchor = properties.anchor;
    this.tooltip = properties.tooltip;
    this.relationshipId = properties.relationshipId;
    this.tgtFrame = properties.tgtFrame;
    this.history = properties.history;
    this.docLocation = properties.docLocation;
    this._isEmpty = properties.isEmpty ?? false;

    // VALIDATION: Warn about hybrid links (url + anchor)
    if (this.url && this.anchor) {
      defaultLogger.warn(
        `DocXML Warning: Hyperlink has both URL ("${this.url}") and anchor ("${this.anchor}"). ` +
          `This is ambiguous per ECMA-376 spec. URL will take precedence. ` +
          `Use Hyperlink.createExternal() or Hyperlink.createInternal() to avoid ambiguity.`
      );
    }

    // Handle empty/invisible hyperlinks (no display text)
    if (this._isEmpty) {
      this.text = '';
      this.formatting = {};
      this.runs = [new Run('', {})];
      return;
    }

    // Text fallback: properties.text → url → 'Link'
    // NOTE: Do NOT use anchor (bookmark ID) as display text - it should only be used for navigation
    // Using bookmark IDs as visible text causes TOC corruption (Issue: TOC shows "HEADING=II.MNKE7E8NA385_" instead of proper headings)
    this.text = properties.text || this.url || 'Link';

    // Validate text for XML patterns
    // Default to auto-cleaning XML patterns unless explicitly disabled (matches Run behavior)
    const validation = validateRunText(this.text, {
      context: 'Hyperlink text',
      autoClean: properties.formatting?.cleanXmlFromText !== false,
      warnToConsole: true,
    });

    // Use cleaned text if available and cleaning was requested
    if (validation.cleanedText) {
      this.text = validation.cleanedText;
    }

    // Default hyperlink styling (Verdana 12pt blue underlined) is for
    // programmatically created links. Parsed links must keep exactly the
    // formatting found in the source XML — merging these defaults underneath
    // would add direct formatting that overrides the document's Hyperlink
    // character style on round-trip.
    this.formatting = properties.preserveFormatting
      ? { ...properties.formatting }
      : {
          font: 'Verdana',
          size: 12,
          color: '0000FF', // Standard hyperlink blue
          underline: 'single',
          ...properties.formatting,
        };

    this.runs = [new Run(this.text, this.formatting)];
  }

  /**
   * Sets the tracking context for automatic change tracking.
   * Called by Document when track changes is enabled.
   * @internal
   */
  _setTrackingContext(context: import('../tracking/TrackingContext.js').TrackingContext): void {
    this.trackingContext = context;
  }

  /**
   * Sets the parent paragraph reference for automatic tracking.
   * Called by Paragraph when hyperlink is added.
   * @internal
   */
  _setParentParagraph(paragraph: import('./Paragraph.js').Paragraph): void {
    this._parentParagraph = paragraph;
  }

  /**
   * Gets the parent paragraph reference.
   * @internal
   */
  _getParentParagraph(): import('./Paragraph.js').Paragraph | undefined {
    return this._parentParagraph;
  }

  /**
   * Applies an rPrChange to the inner Run, tracking the delta between
   * the baseline formatting and the current formatting.
   *
   * On first call, captures `previousFormatting` as the baseline.
   * Subsequent calls always compare against the same baseline so that
   * multiple sequential formatting changes produce a single merged rPrChange.
   *
   * @internal
   */
  private _applyFormattingRPrChange(previousFormatting: RunFormatting): void {
    // Capture baseline on first tracked formatting change
    if (!this._preTrackingFormatting) {
      this._preTrackingFormatting = { ...previousFormatting };
    }

    // Build previousProperties from the baseline (only changed keys)
    const baseline = this._preTrackingFormatting;
    const current = this.formatting;
    const previousProperties: Partial<RunFormatting> = {};
    let hasChanges = false;

    for (const key of Object.keys(baseline) as (keyof RunFormatting)[]) {
      if (baseline[key] !== current[key]) {
        (previousProperties as Record<string, unknown>)[key] = baseline[key];
        hasChanges = true;
      }
    }
    // Also check for keys in current that weren't in baseline (new properties)
    for (const key of Object.keys(current) as (keyof RunFormatting)[]) {
      if (!(key in baseline) && current[key] !== undefined) {
        // Property didn't exist before — previous value is undefined
        (previousProperties as Record<string, unknown>)[key] = undefined;
        hasChanges = true;
      }
    }

    if (!hasChanges) return;

    const revisionId = this.trackingContext!.getRevisionManager().consumeNextId();
    const author = this.trackingContext!.getAuthor();

    this.runs[0]!.setPropertyChangeRevision({
      id: revisionId,
      author,
      date: new Date(),
      previousProperties,
    });
  }

  /**
   * Gets the hyperlink URL
   */
  getUrl(): string | undefined {
    return this.url;
  }

  /**
   * Gets the complete URL including any anchor fragment.
   *
   * For external links that also have an anchor (e.g., internal bookmark within external page),
   * this returns the URL with the anchor appended as a fragment.
   * For internal-only links (anchor without URL), returns undefined.
   *
   * Note: As of v7.2.0, DocumentParser automatically combines external URLs with anchors
   * during parsing, so getUrl() typically returns the full URL. This method is provided
   * for cases where URL and anchor are set separately via the API.
   *
   * @returns The complete URL with fragment, or undefined for internal-only links
   *
   * @example
   * ```typescript
   * // External link with anchor fragment
   * const link = new Hyperlink({ url: 'https://example.com/', anchor: '!/view?id=123', text: 'Link' });
   * link.getUrl();      // 'https://example.com/'
   * link.getAnchor();   // '!/view?id=123'
   * link.getFullUrl();  // 'https://example.com/#!/view?id=123'
   *
   * // External link without anchor
   * const link2 = Hyperlink.createExternal('https://example.com/page', 'Link');
   * link2.getFullUrl(); // 'https://example.com/page'
   *
   * // Internal link (bookmark reference)
   * const link3 = Hyperlink.createInternal('Section1', 'Go to Section 1');
   * link3.getFullUrl(); // undefined
   * ```
   */
  getFullUrl(): string | undefined {
    if (this.url && this.anchor) {
      return this.url + '#' + this.anchor;
    }
    return this.url;
  }

  /**
   * Gets the anchor (for internal links)
   */
  getAnchor(): string | undefined {
    return this.anchor;
  }

  /**
   * Returns whether this is an empty/invisible hyperlink (has no display text).
   * Empty hyperlinks are self-closing elements in the XML.
   */
  isEmpty(): boolean {
    return this._isEmpty;
  }

  /**
   * Gets the target frame attribute (e.g., "_blank" for new window)
   */
  getTgtFrame(): string | undefined {
    return this.tgtFrame;
  }

  /**
   * Gets the history tracking attribute
   */
  getHistory(): string | undefined {
    return this.history;
  }

  /**
   * Sets the target frame attribute
   * @param tgtFrame Target frame (e.g., "_blank" for new window)
   */
  setTgtFrame(tgtFrame: string | undefined): this {
    this.tgtFrame = tgtFrame;
    return this;
  }

  /**
   * Sets the history tracking attribute
   * @param history History value (e.g., "1" to add to history)
   */
  setHistory(history: string | undefined): this {
    this.history = history;
    return this;
  }

  /**
   * Gets the document location attribute (ECMA-376 §17.16.22)
   */
  getDocLocation(): string | undefined {
    return this.docLocation;
  }

  /**
   * Sets the document location for within-document navigation in external files
   * @param docLocation Location string
   */
  setDocLocation(docLocation: string | undefined): this {
    this.docLocation = docLocation;
    return this;
  }

  /**
   * Gets the display text
   *
   * This method delegates to the internal runs to ensure the returned text
   * is always accurate and matches what will be in the generated XML,
   * per ECMA-376 Part 1 §17.16.22. Multi-run hyperlinks concatenate the
   * text of every run in document order.
   *
   * @returns The display text including any special characters (tabs, breaks, etc.)
   */
  getText(): string {
    return this.runs.map((run) => run.getText()).join('');
  }

  /**
   * Replaces the display text. Collapses to a single run (keeping the first
   * run's formatting) because the caller-supplied text supersedes any
   * per-run split the source document had.
   */
  private replaceRunText(text: string): void {
    const first = this.runs[0]!;
    first.setText(text);
    this.runs = [first];
  }

  /**
   * Sets the display text
   */
  setText(text: string): this {
    // Validate text for XML patterns
    // Default to auto-cleaning unless explicitly disabled (matches Run behavior)
    const validation = validateRunText(text, {
      context: 'Hyperlink.setText',
      autoClean: this.formatting.cleanXmlFromText !== false,
      warnToConsole: true,
    });

    // Use cleaned text if available.
    // `??` (not `||`) so a cleaned result of '' (input that was entirely XML
    // markup) is kept rather than falling back to the original markup-laden text.
    const cleanedText = validation.cleanedText ?? text;

    const previousValue = this.text;

    // Assigning display text clears the empty/invisible flag — otherwise
    // toXML() would keep emitting a childless w:hyperlink and silently drop
    // the text that was just set.
    this._isEmpty = false;

    // Skip if text unchanged
    if (previousValue === cleanedText) {
      return this;
    }

    // If tracking enabled AND has parent paragraph, create delete/insert revision pair
    if (this.trackingContext?.isEnabled() && this._parentParagraph) {
      const author = this.trackingContext.getAuthor();

      // Clone current state for deletion (before applying changes)
      const oldHyperlink = this.clone();

      // Apply the change to this hyperlink
      this.text = cleanedText;
      this.replaceRunText(cleanedText);

      // Create delete/insert revision pair
      const deletion = Revision.createDeletion(author, [oldHyperlink]);
      const insertion = Revision.createInsertion(author, [this]);

      // Replace this hyperlink with the revision pair in parent paragraph
      this._parentParagraph.replaceContent(this, [deletion, insertion]);

      // Clear parent reference since we're now inside a revision
      this._parentParagraph = undefined;

      // Clear baseline formatting (fresh start inside insertion)
      this._preTrackingFormatting = undefined;

      return this;
    }

    // Non-tracking path (original behavior)
    this.text = cleanedText;
    this.replaceRunText(cleanedText);
    return this;
  }

  /**
   * Sets the internal run directly (for advanced use cases like TOC parsing)
   * Used by DocumentParser to preserve run content (tabs, breaks, etc.)
   * @param run - The run to use for this hyperlink
   */
  setRun(run: Run): this {
    this.runs = [run];
    this.text = run.getText();
    return this;
  }

  /**
   * Replaces all runs of this hyperlink (multi-run display text)
   * Used by DocumentParser when a w:hyperlink holds multiple w:r children —
   * Word splits the display text into one run per formatting change, and
   * each run must keep its own rPr for round-trip fidelity.
   * @param runs - Runs in document order (ignored if empty)
   */
  setRuns(runs: Run[]): this {
    if (runs.length === 0) {
      return this;
    }
    this.runs = [...runs];
    this.text = this.getText();
    return this;
  }

  /**
   * Gets all runs carrying the display text
   * Single-run hyperlinks return a one-element array.
   */
  getRuns(): Run[] {
    return [...this.runs];
  }

  /**
   * Gets the tooltip
   */
  getTooltip(): string | undefined {
    return this.tooltip;
  }

  /**
   * Sets the tooltip
   */
  setTooltip(tooltip: string): this {
    const previousValue = this.tooltip;

    // Skip if tooltip unchanged
    if (previousValue === tooltip) {
      return this;
    }

    // If tracking enabled AND has parent paragraph, create delete/insert revision pair
    // Tooltip is a w:hyperlink attribute — no OOXML property-change element exists for it
    if (this.trackingContext?.isEnabled() && this._parentParagraph) {
      const author = this.trackingContext.getAuthor();

      // Clone current state for deletion (before applying changes)
      const oldHyperlink = this.clone();

      // Apply the change to this hyperlink
      this.tooltip = tooltip;

      // Create delete/insert revision pair
      const deletion = Revision.createDeletion(author, [oldHyperlink]);
      const insertion = Revision.createInsertion(author, [this]);

      // Replace this hyperlink with the revision pair in parent paragraph
      this._parentParagraph.replaceContent(this, [deletion, insertion]);

      // Clear parent reference since we're now inside a revision
      this._parentParagraph = undefined;

      // Clear baseline formatting (fresh start inside insertion)
      this._preTrackingFormatting = undefined;

      return this;
    }

    // Non-tracking path
    this.tooltip = tooltip;
    return this;
  }

  /**
   * Gets the relationship ID
   */
  getRelationshipId(): string | undefined {
    return this.relationshipId;
  }

  /**
   * Sets the relationship ID (called by Document during save)
   */
  setRelationshipId(id: string): this {
    this.relationshipId = id;
    return this;
  }

  /**
   * Sets or updates the hyperlink URL
   *
   * When URL is updated, we mark that the relationship needs updating.
   * The actual relationship update happens during Document.save() to ensure
   * proper coordination with the RelationshipManager.
   *
   * **Important:** This method maintains the relationship ID but flags it for update.
   * The RelationshipManager will update the existing relationship's target URL
   * during save, preventing orphaned relationships per ECMA-376 §17.16.22.
   *
   * @param url - The new URL (or undefined to clear)
   * @returns This hyperlink for chaining
   * @throws {Error} If clearing URL would create empty hyperlink (no URL and no anchor)
   *
   * @example
   * ```typescript
   * const link = Hyperlink.createExternal('https://old.com', 'Link');
   * link.setUrl('https://new.com');  // Marks for relationship update
   * await doc.save('updated.docx');  // Updates relationship target
   * ```
   */
  setUrl(url: string | undefined): this {
    // Validate that clearing URL doesn't create empty hyperlink
    if (!url && !this.anchor) {
      throw new Error(
        `Cannot set URL to undefined: Hyperlink "${this.getText()}" has no anchor. ` +
          `Clearing the URL would create an invalid hyperlink per ECMA-376 §17.16.22. ` +
          `Either provide a new URL or delete the hyperlink entirely.`
      );
    }

    // Save old URL before updating (for text fallback logic)
    const oldUrl = this.url;

    // Skip if URL unchanged (optimization)
    if (oldUrl === url) {
      return this;
    }

    // If tracking enabled AND has parent paragraph, create revision pair
    // OOXML has no w:hyperlinkChange element - Word tracks hyperlink changes as delete/insert pairs
    if (this.trackingContext?.isEnabled() && this._parentParagraph) {
      const author = this.trackingContext.getAuthor();

      // Clone current state for deletion (before applying changes)
      const oldHyperlink = this.clone();

      // Apply the change to this hyperlink
      this.url = url;
      this.relationshipId = undefined;
      if (this.getText() === oldUrl) {
        this.text = url || this.anchor || 'Link';
        this.replaceRunText(this.text);
      }

      // Create delete/insert revision pair
      const deletion = Revision.createDeletion(author, [oldHyperlink]);
      const insertion = Revision.createInsertion(author, [this]);

      // Replace this hyperlink with the revision pair in parent paragraph
      this._parentParagraph.replaceContent(this, [deletion, insertion]);

      // Clear parent reference since we're now inside a revision
      this._parentParagraph = undefined;

      return this;
    }

    // Non-tracking path (original behavior)
    this.url = url;

    // Clear the relationship ID so it will be re-registered during save
    // This ensures the relationship target is updated to point to the new URL
    this.relationshipId = undefined;

    // Update text ONLY if it was auto-generated from the old URL
    // This preserves user-provided text (even if it's "Link")
    // Use getText() to ensure we check the actual current text, not stale cache
    if (this.getText() === oldUrl) {
      this.text = url || this.anchor || 'Link';
      this.replaceRunText(this.text);
    }

    return this;
  }

  /**
   * Sets the anchor (for internal links)
   * @param anchor Bookmark name to link to
   * @returns This hyperlink for chaining
   * @throws {Error} If clearing anchor would create empty hyperlink (no URL and no anchor)
   * @example
   * ```typescript
   * const link = Hyperlink.createInternal('OldBookmark', 'Go there');
   * link.setAnchor('NewBookmark');  // Update internal link target
   * ```
   */
  setAnchor(anchor: string | undefined): this {
    // Validate that clearing anchor doesn't create empty hyperlink
    if (!anchor && !this.url) {
      throw new Error(
        `Cannot set anchor to undefined: Hyperlink "${this.getText()}" has no URL. ` +
          `Clearing the anchor would create an invalid hyperlink per ECMA-376 §17.16.22. ` +
          `Either provide a new anchor or delete the hyperlink entirely.`
      );
    }

    // Save old anchor before updating
    const oldAnchor = this.anchor;

    // Skip if anchor unchanged (optimization)
    if (oldAnchor === anchor) {
      return this;
    }

    // If tracking enabled AND has parent paragraph, create revision pair
    // OOXML has no w:hyperlinkChange element - Word tracks hyperlink changes as delete/insert pairs
    if (this.trackingContext?.isEnabled() && this._parentParagraph) {
      const author = this.trackingContext.getAuthor();

      // Clone current state for deletion (before applying changes)
      const oldHyperlink = this.clone();

      // Apply the change to this hyperlink
      this.anchor = anchor;
      if (anchor && this.url) {
        defaultLogger.warn(
          `DocXML Warning: Setting anchor "${anchor}" on hyperlink that has URL "${this.url}". ` +
            `Clearing URL to make this an internal link. Use separate hyperlinks for external and internal links.`
        );
        this.url = undefined;
        this.relationshipId = undefined;
      }
      if (this.getText() === oldAnchor) {
        this.text = anchor || this.url || 'Link';
        this.replaceRunText(this.text);
      }

      // Create delete/insert revision pair
      const deletion = Revision.createDeletion(author, [oldHyperlink]);
      const insertion = Revision.createInsertion(author, [this]);

      // Replace this hyperlink with the revision pair in parent paragraph
      this._parentParagraph.replaceContent(this, [deletion, insertion]);

      // Clear parent reference since we're now inside a revision
      this._parentParagraph = undefined;

      return this;
    }

    // Non-tracking path (original behavior)
    this.anchor = anchor;

    // If converting from external to internal, clear URL and relationship
    if (anchor && this.url) {
      defaultLogger.warn(
        `DocXML Warning: Setting anchor "${anchor}" on hyperlink that has URL "${this.url}". ` +
          `Clearing URL to make this an internal link. Use separate hyperlinks for external and internal links.`
      );
      this.url = undefined;
      this.relationshipId = undefined;
    }

    // Update text ONLY if it was auto-generated from the old anchor
    // Use getText() to ensure we check the actual current text, not stale cache
    if (this.getText() === oldAnchor) {
      this.text = anchor || this.url || 'Link';
      this.replaceRunText(this.text);
    }

    return this;
  }

  /**
   * Gets the run
   * Multi-run hyperlinks return the first run — use getRuns() for all of them.
   */
  getRun(): Run {
    return this.runs[0]!;
  }

  /**
   * Sets run formatting
   *
   * @param formatting - The formatting to apply
   * @param options - Optional settings
   * @param options.replace - If true, replaces ALL existing formatting instead of merging.
   *                          Use this when you want to clear inherited styles like characterStyle.
   *
   * @example
   * ```typescript
   * // Merge mode (default): adds/updates properties while preserving others
   * hyperlink.setFormatting({ bold: true });
   *
   * // Replace mode: clears all existing formatting and applies only the new properties
   * hyperlink.setFormatting({ font: "Verdana", size: 12 }, { replace: true });
   * ```
   */
  setFormatting(formatting: RunFormatting, options?: { replace?: boolean }): this {
    // Update stored formatting
    const previousFormatting = { ...this.formatting };
    if (options?.replace) {
      // Replace mode: new formatting replaces ALL existing properties
      this.formatting = { ...formatting };
    } else {
      // Merge mode (default, backwards-compatible): merge with existing
      this.formatting = { ...this.formatting, ...formatting };
    }
    // Rebuild every run, layering the patch over each run's own rPr so
    // multi-run links keep per-run distinctions the caller did not touch.
    // Replace mode discards per-run rPr by design.
    this.runs = this.runs.map((run) =>
      Run.createFromContent(
        run.getContent(),
        options?.replace ? { ...formatting } : { ...run.getFormatting(), ...formatting }
      )
    );
    this.text = this.getText(); // Keep cache in sync
    if (this.trackingContext?.isEnabled()) {
      this._applyFormattingRPrChange(previousFormatting);
    }
    return this;
  }

  /**
   * Layers a formatting patch over every display-text run. Content is
   * rebuilt via Run.createFromContent so tabs/breaks survive, and each
   * run keeps its other rPr values — stamping the hyperlink-level
   * formatting over them would erase per-run formatting in multi-run links.
   */
  private patchRunFormatting(patch: RunFormatting): void {
    this.runs = this.runs.map((run) =>
      Run.createFromContent(run.getContent(), { ...run.getFormatting(), ...patch })
    );
  }

  /**
   * Gets run formatting (returns this hyperlink for fluent API)
   * @returns This hyperlink for method chaining
   *
   * @example
   * ```typescript
   * hyperlink.getFormatting().setColor('0563C1').setUnderline('single');
   * ```
   */
  getFormatting(): this {
    return this;
  }

  /**
   * Gets the raw formatting object (for direct access)
   * @returns RunFormatting object
   */
  getRawFormatting(): RunFormatting {
    return this.formatting;
  }

  // ============================================================================
  // Individual Formatting Getters
  // ============================================================================

  /**
   * Gets the text color
   * @returns Color hex string or undefined
   */
  getColor(): string | undefined {
    return this.formatting.color;
  }

  /**
   * Gets the underline style
   * @returns Underline style or undefined
   */
  getUnderline(): string | boolean | undefined {
    return this.formatting.underline;
  }

  /**
   * Gets whether the hyperlink is bold
   * @returns True if bold, false otherwise
   */
  getBold(): boolean {
    return this.formatting.bold ?? false;
  }

  /**
   * Gets whether the hyperlink is italic
   * @returns True if italic, false otherwise
   */
  getItalic(): boolean {
    return this.formatting.italic ?? false;
  }

  /**
   * Gets the font family
   * @returns Font name or undefined
   */
  getFont(): string | undefined {
    return this.formatting.font;
  }

  /**
   * Gets the font size
   * @returns Font size in points or undefined
   */
  getSize(): number | undefined {
    return this.formatting.size;
  }

  /**
   * Sets text color
   * @param color Color in hex format (e.g., '0563C1')
   * @returns This hyperlink for chaining
   */
  setColor(color: string): this {
    const previousFormatting = { ...this.formatting };
    this.formatting.color = color;
    this.patchRunFormatting({ color });
    if (this.trackingContext?.isEnabled() && previousFormatting.color !== color) {
      this._applyFormattingRPrChange(previousFormatting);
    }
    return this;
  }

  /**
   * Sets underline style
   * @param underline Underline style ('single', 'double', etc.)
   * @returns This hyperlink for chaining
   */
  setUnderline(underline: boolean | 'single' | 'double' | 'dotted' | 'thick' | 'dash'): this {
    const previousFormatting = { ...this.formatting };
    this.formatting.underline = underline;
    this.patchRunFormatting({ underline });
    if (this.trackingContext?.isEnabled() && previousFormatting.underline !== underline) {
      this._applyFormattingRPrChange(previousFormatting);
    }
    return this;
  }

  /**
   * Sets bold formatting
   * @param bold Bold state (default: true)
   * @returns This hyperlink for chaining
   */
  setBold(bold = true): this {
    const previousFormatting = { ...this.formatting };
    this.formatting.bold = bold;
    this.patchRunFormatting({ bold });
    if (this.trackingContext?.isEnabled() && previousFormatting.bold !== bold) {
      this._applyFormattingRPrChange(previousFormatting);
    }
    return this;
  }

  /**
   * Sets italic formatting
   * @param italic Italic state (default: true)
   * @returns This hyperlink for chaining
   */
  setItalic(italic = true): this {
    const previousFormatting = { ...this.formatting };
    this.formatting.italic = italic;
    this.patchRunFormatting({ italic });
    if (this.trackingContext?.isEnabled() && previousFormatting.italic !== italic) {
      this._applyFormattingRPrChange(previousFormatting);
    }
    return this;
  }

  /**
   * Sets font family
   * @param font Font name (e.g., 'Arial', 'Verdana')
   * @returns This hyperlink for chaining
   */
  setFont(font: string): this {
    const previousFormatting = { ...this.formatting };
    this.formatting.font = font;
    this.patchRunFormatting({ font });
    if (this.trackingContext?.isEnabled() && previousFormatting.font !== font) {
      this._applyFormattingRPrChange(previousFormatting);
    }
    return this;
  }

  /**
   * Sets font size
   * @param size Font size in points (e.g., 12, 14)
   * @returns This hyperlink for chaining
   */
  setSize(size: number): this {
    const previousFormatting = { ...this.formatting };
    this.formatting.size = size;
    this.patchRunFormatting({ size });
    if (this.trackingContext?.isEnabled() && previousFormatting.size !== size) {
      this._applyFormattingRPrChange(previousFormatting);
    }
    return this;
  }

  /**
   * Validates the hyperlink URL and optionally fixes common issues
   *
   * Performs validation and fixing of hyperlink URLs including:
   * - Checking URL accessibility (HTTP HEAD request for external links)
   * - Fixing common URL issues (missing protocol, double slashes, spaces)
   * - Validating internal bookmark references
   * - Detecting broken links
   *
   * **Note:** This method is async due to network requests for accessibility checks.
   *
   * @param options - Validation options
   * @returns Promise with validation results
   *
   * @example
   * ```typescript
   * // Basic URL fixing without network check
   * const result = await link.validateAndFix({
   *   fixCommonIssues: true,
   *   checkAccessibility: false
   * });
   * console.log(`Fixed: ${result.fixed.join(', ')}`);
   *
   * // Full validation with accessibility check
   * const validation = await link.validateAndFix({
   *   checkAccessibility: true,
   *   timeout: 5000
   * });
   * if (!validation.valid) {
   *   console.log(`Issues: ${validation.issues.join(', ')}`);
   * }
   *
   * // Batch validate all hyperlinks in document
   * for (const { hyperlink } of doc.getHyperlinks()) {
   *   const result = await hyperlink.validateAndFix();
   *   if (result.fixed.length > 0) {
   *     console.log(`Fixed ${hyperlink.getUrl()}: ${result.fixed.join(', ')}`);
   *   }
   * }
   * ```
   */
  async validateAndFix(options?: {
    checkAccessibility?: boolean;
    fixCommonIssues?: boolean;
    timeout?: number;
    bookmarkManager?: { hasBookmark(name: string): boolean };
  }): Promise<{
    valid: boolean;
    issues: string[];
    fixed: string[];
    originalUrl?: string;
    fixedUrl?: string;
  }> {
    const {
      checkAccessibility = false,
      fixCommonIssues = true,
      timeout = 5000,
      bookmarkManager,
    } = options || {};

    const issues: string[] = [];
    const fixed: string[] = [];
    let fixedUrl = this.url;
    const originalUrl = this.url;

    // Internal link validation (bookmarks)
    if (this.anchor) {
      if (bookmarkManager) {
        const bookmarkExists = bookmarkManager.hasBookmark(this.anchor);
        if (!bookmarkExists) {
          issues.push(`Internal bookmark "${this.anchor}" not found`);
        }
      }
      return {
        valid: issues.length === 0,
        issues,
        fixed,
        originalUrl,
      };
    }

    // External link validation
    if (!this.url) {
      issues.push('No URL or anchor specified');
      return { valid: false, issues, fixed, originalUrl };
    }

    // Fix common issues
    if (fixCommonIssues && fixedUrl) {
      // Fix 0: Strip browser extension URL prefixes
      const sanitizeResult = sanitizeHyperlinkUrl(fixedUrl);
      if (sanitizeResult) {
        fixedUrl = sanitizeResult.url;
        fixed.push(...sanitizeResult.fixes);
      }

      // Fix 1: Add missing protocol
      if (!/^[a-z]+:\/\//i.exec(fixedUrl)) {
        fixedUrl = 'https://' + fixedUrl;
        fixed.push('Added missing protocol (https://)');
      }

      // Fix 2: Fix double slashes (except after protocol)
      const protocolMatch = /^([a-z]+:\/\/)/i.exec(fixedUrl);
      if (protocolMatch?.[1]) {
        const protocol = protocolMatch[1];
        const rest = fixedUrl.substring(protocol.length);
        const fixedRest = rest.replace(/\/\//g, '/');
        if (rest !== fixedRest) {
          fixedUrl = protocol + fixedRest;
          fixed.push('Fixed double slashes');
        }
      }

      // Fix 3: Encode spaces
      if (fixedUrl.includes(' ')) {
        fixedUrl = fixedUrl.replace(/ /g, '%20');
        fixed.push('Encoded spaces as %20');
      }

      // Fix 4: Remove trailing slashes for non-root URLs
      if (/^https?:\/\/[^/]+\/.+\/$/.exec(fixedUrl)) {
        fixedUrl = fixedUrl.replace(/\/$/, '');
        fixed.push('Removed trailing slash');
      }

      // Fix 5: Fix common typos
      // Compare against the value immediately before this replace (not this.url)
      // so 'Upgraded HTTP to HTTPS' is reported only when the http→https
      // rewrite actually fired — an earlier fix (space encoding, double-slash
      // collapse) on an already-https URL must not trigger this message.
      const beforeUpgrade = fixedUrl;
      fixedUrl = fixedUrl.replace(/^http:\/\//i, 'https://'); // Prefer HTTPS
      if (fixedUrl !== beforeUpgrade) {
        fixed.push('Upgraded HTTP to HTTPS');
      }

      // Update URL if fixes were applied
      if (fixedUrl !== this.url) {
        this.setUrl(fixedUrl);
      }
    }

    // Check accessibility (HTTP HEAD request)
    if (checkAccessibility && fixedUrl?.match(/^https?:\/\//i)) {
      // Check if fetch is available (Node.js 18+ or browser)
      if (typeof fetch === 'undefined') {
        issues.push('Network validation unavailable: fetch API not supported in this environment');
      } else {
        try {
          // Use fetch with AbortController for timeout
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), timeout);

          const response = await fetch(fixedUrl, {
            method: 'HEAD',
            signal: controller.signal,
            redirect: 'follow',
          });

          clearTimeout(timeoutId);

          if (!response.ok) {
            issues.push(`HTTP ${response.status}: ${response.statusText || 'Error'}`);
          }
        } catch (error: unknown) {
          // Type guard for error objects with name and message properties
          const isErrorWithName = (err: unknown): err is { name: string } => {
            return typeof err === 'object' && err !== null && 'name' in err;
          };
          const isErrorWithMessage = (err: unknown): err is { message: string } => {
            return typeof err === 'object' && err !== null && 'message' in err;
          };

          if (isErrorWithName(error) && error.name === 'AbortError') {
            issues.push(`Timeout after ${timeout}ms`);
          } else if (isErrorWithMessage(error) && error.message?.includes('fetch')) {
            issues.push(`Unreachable: ${error.message}`);
          } else if (isErrorWithMessage(error)) {
            issues.push(`Network error: ${error.message}`);
          } else {
            issues.push('Network error: Unknown error');
          }
        }
      }
    }

    return {
      valid: issues.length === 0,
      issues,
      fixed,
      originalUrl,
      fixedUrl: fixedUrl !== originalUrl ? fixedUrl : undefined,
    };
  }

  /**
   * Resets hyperlink formatting to standard style (Calibri, blue, underline)
   * This is useful for fixing corrupted hyperlinks from Google Docs or other sources
   * @returns this for method chaining
   */
  resetToStandardFormatting(): this {
    const standardFormatting: RunFormatting = {
      font: 'Verdana',
      color: '0000FF', // Standard hyperlink blue
      underline: 'single',
      // Clear any other formatting that might be causing issues
      bold: false,
      italic: false,
      strike: false,
    };

    this.setFormatting(standardFormatting);
    return this;
  }

  /**
   * Checks if this is an external link
   */
  isExternal(): boolean {
    return this.url !== undefined;
  }

  /**
   * Checks if this is an internal link (anchor)
   */
  isInternal(): boolean {
    return this.anchor !== undefined;
  }

  /**
   * Creates a deep copy of this hyperlink
   *
   * This is useful for preserving the original state before modifications,
   * particularly when creating tracked changes (revisions) where both the
   * old and new states need to be preserved.
   *
   * @returns A new Hyperlink instance with the same properties
   *
   * @example
   * ```typescript
   * // Clone before modifying for tracked changes
   * const originalLink = hyperlink.clone();
   * hyperlink.setUrl('https://new-url.com');
   * hyperlink.setText('New Text');
   *
   * // Now originalLink has old URL/text, hyperlink has new
   * const deletion = Revision.createDeletion(author, [originalLink]);
   * const insertion = Revision.createInsertion(author, [hyperlink]);
   * ```
   */
  clone(): Hyperlink {
    const cloned = new Hyperlink({
      url: this.url,
      anchor: this.anchor,
      text: this.text,
      tooltip: this.tooltip,
      relationshipId: this.relationshipId,
      formatting: { ...this.formatting },
      // this.formatting is already fully resolved — re-merging constructor
      // defaults would resurrect styling the original no longer carries
      preserveFormatting: true,
      tgtFrame: this.tgtFrame,
      history: this.history,
      docLocation: this.docLocation,
      // Carry the empty/invisible flag so a clone of a self-closing hyperlink
      // serializes self-closing too (matters for the deletion half of tracked
      // changes, which snapshots the source via clone()).
      isEmpty: this._isEmpty,
    });

    // Copy every run with its content and formatting (multi-run links keep
    // one cloned run per source run)
    cloned.runs = this.runs.map((run) => {
      const copy = run.clone();

      // Preserve rPrChange from the original run (formatting tracked changes)
      // — Run.clone() copies content/formatting but not the revision
      const existingRPrChange = run.getPropertyChangeRevision();
      if (existingRPrChange) {
        copy.setPropertyChangeRevision({ ...existingRPrChange });
      }
      return copy;
    });
    cloned.text = cloned.getText();

    return cloned;
  }

  /**
   * Generates XML for the hyperlink
   *
   * **CRITICAL:** For external links, relationshipId MUST be set before calling toXML().
   * This happens automatically when saving via Document.save(), but manual usage requires
   * registering the hyperlink with RelationshipManager first.
   *
   * @throws {Error} If external link (has url) is missing relationshipId
   * @throws {Error} If hyperlink has neither url nor anchor (empty hyperlink)
   */
  toXML(): XMLElement {
    // VALIDATION: Hyperlink must have url OR anchor (unless it's an empty hyperlink with relationshipId)
    if (!this.url && !this.anchor && !this.relationshipId) {
      throw new Error(
        'CRITICAL: Hyperlink must have either a URL (external link), anchor (internal link), or relationshipId. ' +
          'Cannot generate valid XML for hyperlink without destination.'
      );
    }

    // VALIDATION: External links MUST have relationship ID
    // Per ECMA-376 Part 1 §17.16.22, <w:hyperlink> with external target requires r:id attribute
    if (this.url && !this.relationshipId) {
      throw new Error(
        `CRITICAL: External hyperlink to "${this.url}" is missing relationship ID. ` +
          `This would create an invalid OpenXML document per ECMA-376 §17.16.22. ` +
          `Solution: Use Document.save() which automatically registers relationships, ` +
          `or manually call relationshipManager.addHyperlink(url) and set the relationship ID.`
      );
    }

    const attributes: Record<string, string> = {};

    // External link - add relationship ID
    if (this.relationshipId) {
      attributes['r:id'] = this.relationshipId;
    }

    // Internal link - uses anchor
    if (this.anchor) {
      attributes['w:anchor'] = this.anchor;
    }

    // Tooltip - explicitly escape attribute value for safety
    // XMLBuilder will handle escaping, but we document this for clarity
    if (this.tooltip) {
      // Note: XMLBuilder.elementToString() will escape this via escapeXmlAttribute()
      // when generating the actual XML string. We store the raw value here.
      attributes['w:tooltip'] = this.tooltip;
    }

    // Target frame attribute (e.g., "_blank" for new window)
    if (this.tgtFrame) {
      attributes['w:tgtFrame'] = this.tgtFrame;
    }

    // History tracking attribute (CT_OnOff per ECMA-376 §17.16.22).
    // Use `!== undefined` rather than a truthy check so explicit "0" /
    // "false" / "off" overrides survive round-trip — the parser
    // normalises to "1"/"0", but external callers may still set any
    // ST_OnOff literal.
    if (this.history !== undefined) {
      attributes['w:history'] = this.history;
    }

    // Document location attribute (ECMA-376 §17.16.22)
    if (this.docLocation) {
      attributes['w:docLocation'] = this.docLocation;
    }

    // Empty/invisible hyperlinks have no children (self-closing element)
    if (this._isEmpty) {
      return {
        name: 'w:hyperlink',
        attributes,
        children: [],
      };
    }

    // One <w:r> child per run, each with its own <w:rPr> — collapsing to a
    // single run would drop per-run formatting in multi-run display text
    return {
      name: 'w:hyperlink',
      attributes,
      children: this.runs.map((run) => run.toXML()),
    };
  }

  /**
   * Creates an external hyperlink
   * @param url The URL
   * @param text Display text
   * @param formatting Optional formatting
   */
  static createExternal(url: string, text: string, formatting?: RunFormatting): Hyperlink {
    return new Hyperlink({ url, text, formatting });
  }

  /**
   * Creates an internal hyperlink (to a bookmark)
   * @param anchor Bookmark name
   * @param text Display text
   * @param formatting Optional formatting
   */
  static createInternal(anchor: string, text: string, formatting?: RunFormatting): Hyperlink {
    return new Hyperlink({ anchor, text, formatting });
  }

  /**
   * Creates a web link (convenience method for URLs)
   * @param url The URL
   * @param text Display text (defaults to URL)
   * @param formatting Optional formatting
   */
  static createWebLink(url: string, text?: string, formatting?: RunFormatting): Hyperlink {
    return new Hyperlink({
      url,
      text: text || url,
      formatting,
    });
  }

  /**
   * Creates an email link
   * @param email Email address
   * @param text Display text (defaults to email)
   * @param formatting Optional formatting
   */
  static createEmail(email: string, text?: string, formatting?: RunFormatting): Hyperlink {
    return new Hyperlink({
      url: `mailto:${email}`,
      text: text || email,
      formatting,
    });
  }

  /**
   * Creates a hyperlink with properties
   * @param properties Hyperlink properties
   */
  static create(properties: HyperlinkProperties): Hyperlink {
    return new Hyperlink(properties);
  }
}
