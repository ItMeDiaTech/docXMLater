/**
 * TableOfContents - Represents a Table of Contents in a Word document
 *
 * A TOC is a special field that automatically generates a list of headings
 * based on heading styles (Heading1, Heading2, etc.) in the document.
 */

import { XMLElement } from '../xml/XMLBuilder';

/**
 * TOC properties
 */
export interface TOCProperties {
  /** Title for the TOC (default: "Table of Contents") */
  title?: string;
  /** Heading levels to include (1-9, default: 1-3) */
  levels?: number;
  /** Whether to show page numbers (default: true) */
  showPageNumbers?: boolean;
  /** Whether to right-align page numbers (default: true) */
  rightAlignPageNumbers?: boolean;
  /** Whether to use hyperlinks instead of page numbers (default: false) */
  useHyperlinks?: boolean;
  /** Custom TOC style (default: built-in TOC style) */
  style?: string;
  /** Tab leader character (default: dot) */
  tabLeader?: 'dot' | 'hyphen' | 'underscore' | 'none';
  /** Custom field switches */
  fieldSwitches?: string;
}

/**
 * Represents a Table of Contents
 */
export class TableOfContents {
  private title: string;
  private levels: number;
  private showPageNumbers: boolean;
  private useHyperlinks: boolean;
  private tabLeader: 'dot' | 'hyphen' | 'underscore' | 'none';
  private fieldSwitches?: string;

  /**
   * Creates a new Table of Contents
   * @param properties TOC properties
   */
  constructor(properties: TOCProperties = {}) {
    this.title = properties.title || 'Table of Contents';
    this.levels = properties.levels || 3;
    this.showPageNumbers = properties.showPageNumbers !== false;
    this.useHyperlinks = properties.useHyperlinks || false;
    this.tabLeader = properties.tabLeader || 'dot';
    this.fieldSwitches = properties.fieldSwitches;

    // Note: rightAlignPageNumbers and style are stored in properties but not used in current implementation
    // These can be used for future enhancements
  }

  /**
   * Gets the TOC title
   */
  getTitle(): string {
    return this.title;
  }

  /**
   * Sets the TOC title
   */
  setTitle(title: string): this {
    this.title = title;
    return this;
  }

  /**
   * Gets the number of heading levels to include
   */
  getLevels(): number {
    return this.levels;
  }

  /**
   * Sets the number of heading levels to include (1-9)
   */
  setLevels(levels: number): this {
    if (levels < 1 || levels > 9) {
      throw new Error('TOC levels must be between 1 and 9');
    }
    this.levels = levels;
    return this;
  }

  /**
   * Gets whether page numbers are shown
   */
  getShowPageNumbers(): boolean {
    return this.showPageNumbers;
  }

  /**
   * Sets whether to show page numbers
   */
  setShowPageNumbers(show: boolean): this {
    this.showPageNumbers = show;
    return this;
  }

  /**
   * Gets whether to use hyperlinks
   */
  getUseHyperlinks(): boolean {
    return this.useHyperlinks;
  }

  /**
   * Sets whether to use hyperlinks instead of page numbers
   */
  setUseHyperlinks(use: boolean): this {
    this.useHyperlinks = use;
    return this;
  }

  /**
   * Builds the TOC field instruction string
   */
  private buildFieldInstruction(): string {
    let instruction = 'TOC';

    // Add heading levels switch
    instruction += ` \\o "1-${this.levels}"`;

    // Add hyperlinks switch if enabled
    if (this.useHyperlinks) {
      instruction += ' \\h';
    }

    // Add page number switches
    if (!this.showPageNumbers) {
      instruction += ' \\n';
    }

    // Add tab leader switch
    if (this.tabLeader !== 'dot') {
      const leaderMap = {
        hyphen: 'h',
        underscore: 'u',
        none: 'n',
      };
      instruction += ` \\p "${leaderMap[this.tabLeader as keyof typeof leaderMap]}"`;
    }

    // Add custom field switches
    if (this.fieldSwitches) {
      instruction += ` ${this.fieldSwitches}`;
    }

    // Add MERGEFORMAT to preserve formatting
    instruction += ' \\* MERGEFORMAT';

    return instruction;
  }

  /**
   * Generates XML for the TOC field
   *
   * A TOC in Word is represented as:
   * 1. A paragraph with the title (styled as TOC Heading)
   * 2. A complex field (fldChar) with the TOC instruction
   * 3. Placeholder entries (updated by Word when opening)
   */
  toXML(): XMLElement[] {
    const elements: XMLElement[] = [];

    // 1. Title paragraph
    if (this.title) {
      elements.push({
        name: 'w:p',
        children: [
          {
            name: 'w:pPr',
            children: [
              {
                name: 'w:pStyle',
                attributes: { 'w:val': 'TOCHeading' },
                selfClosing: true,
              },
            ],
          },
          {
            name: 'w:r',
            children: [
              {
                name: 'w:t',
                children: [this.title],
              },
            ],
          },
        ],
      });
    }

    // 2. TOC field paragraph
    const tocParagraph: XMLElement = {
      name: 'w:p',
      children: [],
    };

    // Begin field
    tocParagraph.children!.push({
      name: 'w:r',
      children: [
        {
          name: 'w:fldChar',
          attributes: { 'w:fldCharType': 'begin' },
          selfClosing: true,
        },
      ],
    });

    // Field instruction
    tocParagraph.children!.push({
      name: 'w:r',
      children: [
        {
          name: 'w:instrText',
          attributes: { 'xml:space': 'preserve' },
          children: [this.buildFieldInstruction()],
        },
      ],
    });

    // Separate field
    tocParagraph.children!.push({
      name: 'w:r',
      children: [
        {
          name: 'w:fldChar',
          attributes: { 'w:fldCharType': 'separate' },
          selfClosing: true,
        },
      ],
    });

    // Placeholder text (Word will replace this)
    tocParagraph.children!.push({
      name: 'w:r',
      children: [
        {
          name: 'w:rPr',
          children: [
            {
              name: 'w:noProof',
              selfClosing: true,
            },
          ],
        },
        {
          name: 'w:t',
          children: ['Right-click to update field.'],
        },
      ],
    });

    // End field
    tocParagraph.children!.push({
      name: 'w:r',
      children: [
        {
          name: 'w:fldChar',
          attributes: { 'w:fldCharType': 'end' },
          selfClosing: true,
        },
      ],
    });

    elements.push(tocParagraph);

    return elements;
  }

  /**
   * Creates a standard TOC with 3 levels
   */
  static createStandard(title?: string): TableOfContents {
    return new TableOfContents({
      title: title || 'Table of Contents',
      levels: 3,
      showPageNumbers: true,
      rightAlignPageNumbers: true,
    });
  }

  /**
   * Creates a simple TOC with 2 levels
   */
  static createSimple(title?: string): TableOfContents {
    return new TableOfContents({
      title: title || 'Contents',
      levels: 2,
      showPageNumbers: true,
      rightAlignPageNumbers: true,
    });
  }

  /**
   * Creates a detailed TOC with 4 levels
   */
  static createDetailed(title?: string): TableOfContents {
    return new TableOfContents({
      title: title || 'Table of Contents',
      levels: 4,
      showPageNumbers: true,
      rightAlignPageNumbers: true,
    });
  }

  /**
   * Creates a hyperlinked TOC (for web documents)
   */
  static createHyperlinked(title?: string): TableOfContents {
    return new TableOfContents({
      title: title || 'Contents',
      levels: 3,
      showPageNumbers: false,
      useHyperlinks: true,
    });
  }

  /**
   * Creates a TOC with custom properties
   */
  static create(properties?: TOCProperties): TableOfContents {
    return new TableOfContents(properties);
  }
}
