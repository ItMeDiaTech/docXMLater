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
    this.text = text;
    this.run.setText(text);
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
   * When URL is updated, the relationship ID is cleared to force re-registration.
   * This ensures the relationship points to the new URL when Document.save() is called.
   *
   * **Important:** This clears the relationship ID, which will be automatically
   * re-registered when save() or toBuffer() is called. The new relationship will
   * point to the updated URL, ensuring OpenXML compliance per ECMA-376.
   *
   * @param url - The new URL (or undefined to clear)
   * @returns This hyperlink for chaining
   *
   * @example
   * ```typescript
   * const link = Hyperlink.createExternal('https://old.com', 'Link');
   * link.setUrl('https://new.com');  // Clears relationshipId
   * await doc.save('updated.docx');  // Re-registers with new URL
   * ```
   */
  setUrl(url: string | undefined): this {
    // Save old URL before updating (for text fallback logic)
    const oldUrl = this.url;

    // Update URL
    this.url = url;

    // Clear relationship ID when URL changes
    // This forces Document.save() to re-register the hyperlink
    // with the new URL, preventing orphaned relationships
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
