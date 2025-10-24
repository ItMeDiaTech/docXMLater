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

import { XMLElement } from '../xml/XMLBuilder';
import { Run, RunFormatting } from './Run';
import { validateRunText } from '../utils/validation';

/**
 * Hyperlink properties
 */
export interface HyperlinkProperties {
  /** Hyperlink URL (for external links) */
  url?: string;
  /** Bookmark anchor (for internal links) */
  anchor?: string;
  /** Display text */
  text: string;
  /** Text formatting */
  formatting?: RunFormatting;
  /** Tooltip text */
  tooltip?: string;
  /** Relationship ID (set by Document when saving) */
  relationshipId?: string;
}

/**
 * Represents a hyperlink
 */
export class Hyperlink {
  private url?: string;
  private anchor?: string;
  private text: string;
  private run: Run;
  private tooltip?: string;
  private relationshipId?: string;


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

    // VALIDATION: Warn about hybrid links (url + anchor)
    if (this.url && this.anchor) {
      console.warn(
        `DocXML Warning: Hyperlink has both URL ("${this.url}") and anchor ("${this.anchor}"). ` +
        `This is ambiguous per ECMA-376 spec. URL will take precedence. ` +
        `Use Hyperlink.createExternal() or Hyperlink.createInternal() to avoid ambiguity.`
      );
    }

    // Improved text fallback: url → anchor → provided text → 'Link'
    // This provides better UX when text is empty
    this.text = properties.text || this.url || this.anchor || 'Link';

    // Validate text for XML patterns
    const validation = validateRunText(this.text, {
      context: 'Hyperlink text',
      autoClean: properties.formatting?.cleanXmlFromText || false,
      warnToConsole: true,
    });

    // Use cleaned text if available and cleaning was requested
    if (validation.cleanedText) {
      this.text = validation.cleanedText;
    }

    // Create run with default hyperlink styling (blue, underlined)
    const formatting: RunFormatting = {
      color: '0563C1', // Word's default hyperlink blue
      underline: 'single',
      ...properties.formatting,
    };

    this.run = new Run(this.text, formatting);
  }

  /**
   * Gets the hyperlink URL
   */
  getUrl(): string | undefined {
    return this.url;
  }

  /**
   * Gets the anchor (for internal links)
   */
  getAnchor(): string | undefined {
    return this.anchor;
  }

  /**
   * Gets the display text
   */
  getText(): string {
    return this.text;
  }

  /**
   * Sets the display text
   */
  setText(text: string): this {
    // Validate text for XML patterns (warning only, Run handles cleaning)
    validateRunText(text, {
      context: 'Hyperlink.setText',
      warnToConsole: true,
    });

    this.text = text;
    this.run.setText(text); // Run.setText also validates
    return this;
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
        `Cannot set URL to undefined: Hyperlink "${this.text}" has no anchor. ` +
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

    // Update URL
    this.url = url;

    // Clear the relationship ID so it will be re-registered during save
    // This ensures the relationship target is updated to point to the new URL
    this.relationshipId = undefined;

    // Update text ONLY if it was auto-generated from the old URL
    // This preserves user-provided text (even if it's "Link")
    if (this.text === oldUrl) {
      this.text = url || this.anchor || 'Link';
      this.run.setText(this.text);
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
        `Cannot set anchor to undefined: Hyperlink "${this.text}" has no URL. ` +
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

    // Update anchor
    this.anchor = anchor;

    // If converting from external to internal, clear URL and relationship
    if (anchor && this.url) {
      console.warn(
        `DocXML Warning: Setting anchor "${anchor}" on hyperlink that has URL "${this.url}". ` +
        `Clearing URL to make this an internal link. Use separate hyperlinks for external and internal links.`
      );
      this.url = undefined;
      this.relationshipId = undefined;
    }

    // Update text ONLY if it was auto-generated from the old anchor
    if (this.text === oldAnchor) {
      this.text = anchor || this.url || 'Link';
      this.run.setText(this.text);
    }

    return this;
  }

  /**
   * Gets the run
   */
  getRun(): Run {
    return this.run;
  }

  /**
   * Sets run formatting
   */
  setFormatting(formatting: RunFormatting): this {
    const currentFormatting = this.run.getFormatting();
    this.run = new Run(this.text, { ...currentFormatting, ...formatting });
    return this;
  }

  /**
   * Gets run formatting
   */
  getFormatting(): RunFormatting {
    return this.run.getFormatting();
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
    // VALIDATION: Hyperlink must have url OR anchor
    if (!this.url && !this.anchor) {
      throw new Error(
        'CRITICAL: Hyperlink must have either a URL (external link) or anchor (internal link). ' +
        'Cannot generate valid XML for empty hyperlink.'
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
    if (this.url && this.relationshipId) {
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

    // Generate run XML
    const runXml = this.run.toXML();

    return {
      name: 'w:hyperlink',
      attributes,
      children: [runXml],
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
