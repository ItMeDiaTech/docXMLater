/**
 * Hyperlink - Represents a hyperlink in a Word document
 *
 * Hyperlinks can be external (to websites, files) or internal (to bookmarks within the document).
 * They are represented using the <w:hyperlink> element with a relationship ID.
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
   * @param properties Hyperlink properties
   */
  constructor(properties: HyperlinkProperties) {
    this.url = properties.url;
    this.anchor = properties.anchor;
    this.text = properties.text;
    this.tooltip = properties.tooltip;
    this.relationshipId = properties.relationshipId;

    // Create run with default hyperlink styling (blue, underlined)
    const formatting: RunFormatting = {
      color: '0563C1', // Word's default hyperlink blue
      underline: 'single',
      ...properties.formatting,
    };

    this.run = new Run(properties.text, formatting);
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
   * For external links, requires relationshipId to be set
   */
  toXML(): XMLElement {
    const attributes: Record<string, string> = {};

    // External link - requires relationship ID
    if (this.url && this.relationshipId) {
      attributes['r:id'] = this.relationshipId;
    }

    // Internal link - uses anchor
    if (this.anchor) {
      attributes['w:anchor'] = this.anchor;
    }

    // Tooltip
    if (this.tooltip) {
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
