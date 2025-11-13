/**
 * Run - Represents a run of text with uniform formatting
 * A run is the smallest unit of text formatting in a Word document
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { validateRunText, normalizeColor } from '../utils/validation';
import { logSerialization, logTextDirection } from '../utils/diagnostics';
import { defaultLogger } from '../utils/logger';
import { deepClone } from '../utils/deepClone';

/**
 * Run content element types
 * Per ECMA-376 Part 1 §17.3.3 EG_RunInnerContent, runs can contain multiple types of content
 */
export type RunContentType =
  | 'text'            // <w:t> - Regular text
  | 'tab'             // <w:tab/> - Tab character (used in TOC entries)
  | 'break'           // <w:br/> - Line/page/column break
  | 'carriageReturn'  // <w:cr/> - Carriage return
  | 'softHyphen'      // <w:softHyphen/> - Optional hyphen
  | 'noBreakHyphen';  // <w:noBreakHyphen/> - Non-breaking hyphen

/**
 * Break type for <w:br> elements
 * Per ECMA-376 Part 1 §17.18.3
 */
export type BreakType = 'page' | 'column' | 'textWrapping';

/**
 * Run content element
 * Represents a single content element within a run (text, tab, break, etc.)
 */
export interface RunContent {
  /** Type of content element */
  type: RunContentType;
  /** Text value (only for 'text' type) */
  value?: string;
  /** Break type (only for 'break' type) */
  breakType?: BreakType;
}

/**
 * Border style for text
 */
export type TextBorderStyle = 'single' | 'double' | 'dashed' | 'dotted' | 'thick' | 'wave' | 'dashDotStroked' | 'threeDEmboss' | 'threeDEngrave';

/**
 * Text border definition
 */
export interface TextBorder {
  /** Border style */
  style?: TextBorderStyle;
  /** Border size in eighths of a point */
  size?: number;
  /** Border color in hex format (without #) */
  color?: string;
  /** Space between border and text in points */
  space?: number;
}

/**
 * Shading pattern for text background
 */
export type ShadingPattern = 'clear' | 'solid' | 'horzStripe' | 'vertStripe' | 'reverseDiagStripe' | 'diagStripe' | 'horzCross' | 'diagCross' | 'thinHorzStripe' | 'thinVertStripe' | 'thinReverseDiagStripe' | 'thinDiagStripe' | 'thinHorzCross' | 'thinDiagCross' | 'pct5' | 'pct10' | 'pct12' | 'pct15' | 'pct20' | 'pct25' | 'pct30' | 'pct35' | 'pct37' | 'pct40' | 'pct45' | 'pct50' | 'pct55' | 'pct60' | 'pct62' | 'pct65' | 'pct70' | 'pct75' | 'pct80' | 'pct85' | 'pct87' | 'pct90' | 'pct95';

/**
 * Character shading definition
 */
export interface CharacterShading {
  /** Background fill color in hex format (without #) */
  fill?: string;
  /** Foreground pattern color in hex format (without #) */
  color?: string;
  /** Shading pattern */
  val?: ShadingPattern;
}

/**
 * East Asian typography layout options
 */
export interface EastAsianLayout {
  /** Layout ID for specific Asian typography */
  id?: number;
  /** Vertical text layout */
  vert?: boolean;
  /** Compress vertical text */
  vertCompress?: boolean;
  /** Combine characters into single character space */
  combine?: boolean;
  /** Bracket characters for combined text */
  combineBrackets?: 'none' | 'round' | 'square' | 'angle' | 'curly';
}

/**
 * Emphasis mark type - decorative marks above/below text
 */
export type EmphasisMark = 'dot' | 'comma' | 'circle' | 'underDot';

/**
 * Text formatting options for a run
 */
export interface RunFormatting {
  /** Character style reference - links to a character style definition */
  characterStyle?: string;
  /** Text border - draws a border around the text */
  border?: TextBorder;
  /** Character shading - background color/pattern for text */
  shading?: CharacterShading;
  /** Emphasis mark - decorative mark above/below text (e.g., dot, comma) */
  emphasis?: EmphasisMark;
  /** Bold text */
  bold?: boolean;
  /** Italic text */
  italic?: boolean;
  /** Bold text for complex scripts (RTL languages like Arabic, Hebrew) */
  complexScriptBold?: boolean;
  /** Italic text for complex scripts (RTL languages like Arabic, Hebrew) */
  complexScriptItalic?: boolean;
  /** Character spacing (letter spacing) in twips (1/20th of a point) */
  characterSpacing?: number;
  /** Horizontal scaling percentage (e.g., 200 = 200% width, 50 = 50% width) */
  scaling?: number;
  /** Vertical position in half-points (positive = raised, negative = lowered) */
  position?: number;
  /** Kerning threshold in half-points (font size at which kerning starts) */
  kerning?: number;
  /** Language code (e.g., 'en-US', 'fr-FR', 'es-ES') */
  language?: string;
  /** Underline text */
  underline?: boolean | 'single' | 'double' | 'thick' | 'dotted' | 'dash';
  /** Strikethrough text */
  strike?: boolean;
  /** Double strikethrough */
  dstrike?: boolean;
  /** Subscript */
  subscript?: boolean;
  /** Superscript */
  superscript?: boolean;
  /** Font name */
  font?: string;
  /** Font size in points (half-points for Word) */
  size?: number;
  /** Text color in hex format (without #) */
  color?: string;
  /** Highlight color */
  highlight?: 'yellow' | 'green' | 'cyan' | 'magenta' | 'blue' | 'red' | 'darkBlue' | 'darkCyan' | 'darkGreen' | 'darkMagenta' | 'darkRed' | 'darkYellow' | 'darkGray' | 'lightGray' | 'black' | 'white';
  /** Small caps */
  smallCaps?: boolean;
  /** All caps */
  allCaps?: boolean;
  /** Outline text effect - displays text with an outline */
  outline?: boolean;
  /** Shadow text effect - displays text with a shadow */
  shadow?: boolean;
  /** Emboss text effect - displays text with a 3D embossed appearance */
  emboss?: boolean;
  /** Imprint/engrave text effect - displays text with a pressed-in appearance */
  imprint?: boolean;
  /** Right-to-left text direction (for languages like Arabic, Hebrew) */
  rtl?: boolean;
  /** Hidden/vanish text (not displayed but present in document) */
  vanish?: boolean;
  /** No proofing - skip spell check and grammar check for this text */
  noProof?: boolean;
  /** Snap to grid - align text to document grid */
  snapToGrid?: boolean;
  /** Special vanish - hidden text for specific scenarios (like TOC entries) */
  specVanish?: boolean;
  /** Text effect/animation type */
  effect?: 'none' | 'lights' | 'blinkBackground' | 'sparkleText' | 'marchingBlackAnts' | 'marchingRedAnts' | 'shimmer' | 'antsBlack' | 'antsRed';
  /** Fit text to width in twips (1/20th of a point) */
  fitText?: number;
  /** East Asian typography layout options */
  eastAsianLayout?: EastAsianLayout;
  /**
   * Automatically clean XML-like patterns from text content.
   * When true (default), removes XML tags like <w:t> from text to prevent display issues.
   * Set to false to disable auto-cleaning (useful for debugging).
   * Default: true (auto-clean enabled by default for defensive data handling)
   */
  cleanXmlFromText?: boolean;
}

/**
 * Represents a run of formatted text
 */
export class Run {
  private content: RunContent[];
  private formatting: RunFormatting;

  /**
   * Creates a new Run
   * @param text - The text content
   * @param formatting - Formatting options
   */
  constructor(text: string, formatting: RunFormatting = {}) {
    // Warn about undefined/null text to help catch data quality issues
    if (text === undefined || text === null) {
      defaultLogger.warn(
        `DocXML Text Validation Warning [Run constructor]:\n` +
        `  - Received ${text === undefined ? 'undefined' : 'null'} text value\n` +
        `  - Converting to empty string for Word compatibility`
      );
    }

    // Default to auto-cleaning XML patterns unless explicitly disabled
    const shouldClean = formatting.cleanXmlFromText !== false;

    // Validate text for XML patterns
    const validation = validateRunText(text, {
      context: 'Run constructor',
      autoClean: shouldClean,
      warnToConsole: true,  // Enable warnings to help catch data quality issues
    });

    // Use cleaned text if available and cleaning was requested
    const cleanedText = validation.cleanedText || text;

    // Convert undefined/null to empty string for consistent XML generation
    const normalizedText = cleanedText ?? '';

    // Parse text to extract special characters into separate content elements
    this.content = this.parseTextWithSpecialCharacters(normalizedText);

    // Remove cleanXmlFromText from formatting as it's not a display property
    const { cleanXmlFromText, ...displayFormatting } = formatting;
    this.formatting = displayFormatting;
  }

  /**
   * Parses text containing special characters and converts them to content elements
   * @param text Text that may contain tabs, newlines, non-breaking hyphens, etc.
   * @returns Array of content elements
   * @private
   */
  private parseTextWithSpecialCharacters(text: string): RunContent[] {
    if (!text) {
      return [{ type: 'text', value: '' }];
    }

    const content: RunContent[] = [];
    let currentText = '';

    for (let i = 0; i < text.length; i++) {
      const char = text[i];

      switch (char) {
        case '\t':
          // Add accumulated text before tab
          if (currentText) {
            content.push({ type: 'text', value: currentText });
            currentText = '';
          }
          // Add tab element
          content.push({ type: 'tab' });
          break;

        case '\n':
          // Add accumulated text before newline
          if (currentText) {
            content.push({ type: 'text', value: currentText });
            currentText = '';
          }
          // Add break element
          content.push({ type: 'break' });
          break;

        case '\u2011': // Non-breaking hyphen
          // Add accumulated text before non-breaking hyphen
          if (currentText) {
            content.push({ type: 'text', value: currentText });
            currentText = '';
          }
          // Add non-breaking hyphen element
          content.push({ type: 'noBreakHyphen' });
          break;

        case '\r': // Carriage return
          // Add accumulated text before carriage return
          if (currentText) {
            content.push({ type: 'text', value: currentText });
            currentText = '';
          }
          // Add carriage return element
          content.push({ type: 'carriageReturn' });
          break;

        case '\u00AD': // Soft hyphen
          // Add accumulated text before soft hyphen
          if (currentText) {
            content.push({ type: 'text', value: currentText });
            currentText = '';
          }
          // Add soft hyphen element
          content.push({ type: 'softHyphen' });
          break;

        default:
          // Accumulate regular text
          currentText += char;
          break;
      }
    }

    // Add any remaining text
    if (currentText) {
      content.push({ type: 'text', value: currentText });
    }

    // If no content was added (empty string), add empty text element
    if (content.length === 0) {
      content.push({ type: 'text', value: '' });
    }

    return content;
  }

  /**
   * Gets the text content (concatenates all content elements)
   * Converts special characters back to their string representation
   */
  getText(): string {
    return this.content
      .map(c => {
        switch (c.type) {
          case 'text':
            return c.value || '';
          case 'tab':
            return '\t';
          case 'break':
            return '\n';
          case 'carriageReturn':
            return '\r';
          case 'softHyphen':
            return '\u00AD';
          case 'noBreakHyphen':
            return '\u2011';
          default:
            return '';
        }
      })
      .join('');
  }

  /**
   * Sets the text content (replaces all content with single text element)
   * Backward-compatible method that preserves existing API
   * @param text - New text content
   */
  setText(text: string): void {
    // Warn about undefined/null text to help catch data quality issues
    if (text === undefined || text === null) {
      defaultLogger.warn(
        `DocXML Text Validation Warning [Run.setText]:\n` +
        `  - Received ${text === undefined ? 'undefined' : 'null'} text value\n` +
        `  - Converting to empty string for Word compatibility`
      );
    }

    // Respect original cleanXmlFromText setting (Issue #9 fix)
    // This ensures consistent behavior with constructor
    const shouldClean = this.formatting.cleanXmlFromText !== false;

    // Validate text for XML patterns
    const validation = validateRunText(text, {
      context: 'Run.setText',
      autoClean: shouldClean,
      warnToConsole: true,  // Enable warnings to help catch data quality issues
    });

    // Use cleaned text if available and cleaning was requested
    const cleanedText = validation.cleanedText || text;

    // Convert undefined/null to empty string for consistent XML generation
    const normalizedText = cleanedText ?? '';

    // Parse text to extract special characters into separate content elements
    this.content = this.parseTextWithSpecialCharacters(normalizedText);
  }

  /**
   * Gets the formatting
   */
  getFormatting(): RunFormatting {
    return { ...this.formatting };
  }

  /**
   * Sets character style reference
   * Per ECMA-376 Part 1 §17.3.2.36
   * @param styleId - Character style ID to apply
   * @returns This run for chaining
   */
  setCharacterStyle(styleId: string): this {
    this.formatting.characterStyle = styleId;
    return this;
  }

  /**
   * Sets text border
   * Per ECMA-376 Part 1 §17.3.2.5
   * @param border - Border definition
   * @returns This run for chaining
   */
  setBorder(border: TextBorder): this {
    this.formatting.border = border;
    return this;
  }

  /**
   * Sets character shading (background)
   * Per ECMA-376 Part 1 §17.3.2.32
   * @param shading - Shading definition
   * @returns This run for chaining
   */
  setShading(shading: CharacterShading): this {
    this.formatting.shading = shading;
    return this;
  }

  /**
   * Sets emphasis mark
   * Per ECMA-376 Part 1 §17.3.2.13
   * @param emphasis - Emphasis mark type ('dot', 'comma', 'circle', 'underDot')
   * @returns This run for chaining
   */
  setEmphasis(emphasis: EmphasisMark): this {
    this.formatting.emphasis = emphasis;
    return this;
  }

  /**
   * Sets bold formatting
   * @param bold - Whether text is bold
   */
  setBold(bold: boolean = true): this {
    this.formatting.bold = bold;
    return this;
  }

  /**
   * Sets italic formatting
   * @param italic - Whether text is italic
   */
  setItalic(italic: boolean = true): this {
    this.formatting.italic = italic;
    return this;
  }

  /**
   * Sets bold formatting for complex scripts (RTL languages)
   * Per ECMA-376 Part 1 §17.3.2.3
   * @param bold - Whether text is bold for complex scripts
   */
  setComplexScriptBold(bold: boolean = true): this {
    this.formatting.complexScriptBold = bold;
    return this;
  }

  /**
   * Sets italic formatting for complex scripts (RTL languages)
   * Per ECMA-376 Part 1 §17.3.2.17
   * @param italic - Whether text is italic for complex scripts
   */
  setComplexScriptItalic(italic: boolean = true): this {
    this.formatting.complexScriptItalic = italic;
    return this;
  }

  /**
   * Sets character spacing (letter spacing)
   * Per ECMA-376 Part 1 §17.3.2.33
   * @param spacing - Spacing in twips (1/20th of a point). Positive values expand, negative values condense.
   */
  setCharacterSpacing(spacing: number): this {
    this.formatting.characterSpacing = spacing;
    return this;
  }

  /**
   * Sets horizontal text scaling
   * Per ECMA-376 Part 1 §17.3.2.43
   * @param scaling - Scaling percentage (e.g., 200 = 200% width, 50 = 50% width). Default is 100.
   */
  setScaling(scaling: number): this {
    this.formatting.scaling = scaling;
    return this;
  }

  /**
   * Sets vertical text position
   * Per ECMA-376 Part 1 §17.3.2.31
   * @param position - Position in half-points. Positive values raise text, negative values lower it.
   */
  setPosition(position: number): this {
    this.formatting.position = position;
    return this;
  }

  /**
   * Sets kerning threshold
   * Per ECMA-376 Part 1 §17.3.2.20
   * @param kerning - Font size in half-points at which kerning starts. 0 disables kerning.
   */
  setKerning(kerning: number): this {
    this.formatting.kerning = kerning;
    return this;
  }

  /**
   * Sets language
   * Per ECMA-376 Part 1 §17.3.2.20
   * @param language - Language code (e.g., 'en-US', 'fr-FR', 'es-ES')
   */
  setLanguage(language: string): this {
    this.formatting.language = language;
    return this;
  }

  /**
   * Sets underline formatting
   * @param underline - Underline style or boolean
   */
  setUnderline(underline: RunFormatting['underline'] = true): this {
    this.formatting.underline = underline;
    return this;
  }

  /**
   * Sets strikethrough formatting
   * @param strike - Whether text has strikethrough
   */
  setStrike(strike: boolean = true): this {
    this.formatting.strike = strike;
    return this;
  }

  /**
   * Sets subscript formatting
   * @param subscript - Whether text is subscript
   */
  setSubscript(subscript: boolean = true): this {
    this.formatting.subscript = subscript;
    if (subscript) {
      this.formatting.superscript = false;
    }
    return this;
  }

  /**
   * Sets superscript formatting
   * @param superscript - Whether text is superscript
   */
  setSuperscript(superscript: boolean = true): this {
    this.formatting.superscript = superscript;
    if (superscript) {
      this.formatting.subscript = false;
    }
    return this;
  }

  /**
   * Sets font
   * @param font - Font name
   * @param size - Font size in points (optional)
   */
  setFont(font: string, size?: number): this {
    this.formatting.font = font;
    if (size !== undefined) {
      this.formatting.size = size;
    }
    return this;
  }

  /**
   * Sets font size
   * @param size - Font size in points
   */
  setSize(size: number): this {
    this.formatting.size = size;
    return this;
  }

  /**
   * Sets text color with normalization to uppercase hex
   * @param color - Color in hex format (with or without #)
   * @throws Error if color format is invalid
   */
  setColor(color: string): this {
    this.formatting.color = normalizeColor(color);
    return this;
  }


  /**
   * Sets highlight color
   * @param highlight - Highlight color
   */
  setHighlight(highlight: RunFormatting['highlight']): this {
    this.formatting.highlight = highlight;
    return this;
  }

  /**
   * Sets small caps
   * @param smallCaps - Whether text is in small caps
   */
  setSmallCaps(smallCaps: boolean = true): this {
    this.formatting.smallCaps = smallCaps;
    return this;
  }

  /**
   * Sets all caps
   * @param allCaps - Whether text is in all caps
   */
  setAllCaps(allCaps: boolean = true): this {
    this.formatting.allCaps = allCaps;
    return this;
  }

  /**
   * Sets outline text effect
   * @param outline - Whether to apply outline effect (default: true)
   * @returns This run for method chaining
   */
  setOutline(outline: boolean = true): this {
    this.formatting.outline = outline;
    return this;
  }

  /**
   * Sets shadow text effect
   * @param shadow - Whether to apply shadow effect (default: true)
   * @returns This run for method chaining
   */
  setShadow(shadow: boolean = true): this {
    this.formatting.shadow = shadow;
    return this;
  }

  /**
   * Sets emboss text effect
   * @param emboss - Whether to apply emboss effect (default: true)
   * @returns This run for method chaining
   */
  setEmboss(emboss: boolean = true): this {
    this.formatting.emboss = emboss;
    return this;
  }

  /**
   * Sets imprint/engrave text effect
   * @param imprint - Whether to apply imprint effect (default: true)
   * @returns This run for method chaining
   */
  setImprint(imprint: boolean = true): this {
    this.formatting.imprint = imprint;
    return this;
  }

  /**
   * Sets right-to-left text direction
   * @param rtl - Whether text is RTL (default: true)
   * @returns This run for method chaining
   */
  setRTL(rtl: boolean = true): this {
    this.formatting.rtl = rtl;
    return this;
  }

  /**
   * Sets hidden/vanish text
   * @param vanish - Whether text is hidden (default: true)
   * @returns This run for method chaining
   */
  setVanish(vanish: boolean = true): this {
    this.formatting.vanish = vanish;
    return this;
  }

  /**
   * Sets no proofing (skip spell/grammar check)
   * @param noProof - Whether to skip proofing (default: true)
   * @returns This run for method chaining
   */
  setNoProof(noProof: boolean = true): this {
    this.formatting.noProof = noProof;
    return this;
  }

  /**
   * Sets snap to grid alignment
   * @param snapToGrid - Whether to snap to grid (default: true)
   * @returns This run for method chaining
   */
  setSnapToGrid(snapToGrid: boolean = true): this {
    this.formatting.snapToGrid = snapToGrid;
    return this;
  }

  /**
   * Sets special vanish (hidden for specific scenarios like TOC)
   * @param specVanish - Whether to apply special vanish (default: true)
   * @returns This run for method chaining
   */
  setSpecVanish(specVanish: boolean = true): this {
    this.formatting.specVanish = specVanish;
    return this;
  }

  /**
   * Sets text effect/animation
   * @param effect - Effect type (e.g., 'shimmer', 'sparkleText')
   * @returns This run for method chaining
   */
  setEffect(effect: 'none' | 'lights' | 'blinkBackground' | 'sparkleText' | 'marchingBlackAnts' | 'marchingRedAnts' | 'shimmer' | 'antsBlack' | 'antsRed'): this {
    this.formatting.effect = effect;
    return this;
  }

  /**
   * Sets fit text to width
   * @param width - Width in twips (1/20th of a point)
   * @returns This run for method chaining
   */
  setFitText(width: number): this {
    this.formatting.fitText = width;
    return this;
  }

  /**
   * Sets East Asian typography layout
   * @param layout - East Asian layout options
   * @returns This run for method chaining
   */
  setEastAsianLayout(layout: EastAsianLayout): this {
    this.formatting.eastAsianLayout = layout;
    return this;
  }

  /**
   * Converts the run to WordprocessingML XML element
   *
   * **ECMA-376 Compliance:** Properties are generated in the order specified by
   * ECMA-376 Part 1 §17.3.2.28 to ensure strict OpenXML conformance.
   *
   * Per spec, the order is:
   * 1. rFonts (font family)
   * 2. b (bold)
   * 3. i (italic)
   * 4. caps/smallCaps (capitalization)
   * 5. strike/dstrike (strikethrough)
   * 6. u (underline)
   * 7. sz/szCs (font size)
   * 8. color (text color)
   * 9. highlight (highlight color)
   * 10. vertAlign (subscript/superscript)
   *
   * @returns XMLElement representing the run
   */
  toXML(): XMLElement {
    // Get text for diagnostic logging (backward compatibility)
    const text = this.getText();

    // Diagnostic logging before serialization
    logSerialization(`Serializing run: "${text}"`, { rtl: this.formatting.rtl || false });
    if (this.formatting.rtl) {
      logTextDirection(`Run with RTL being serialized: "${text}"`);
    }

    // Build the run element
    const runChildren: XMLElement[] = [];

    // Add run properties using the static helper
    const rPr = Run.generateRunPropertiesXML(this.formatting);
    if (rPr) {
      runChildren.push(rPr);
    }

    // Add run content elements (text, tabs, breaks, etc.) in order
    for (const contentElement of this.content) {
      switch (contentElement.type) {
        case 'text':
          // Always generate <w:t> element, even for empty strings
          // This ensures proper Word compatibility and round-trip preservation
          runChildren.push(XMLBuilder.w('t', {
            'xml:space': 'preserve',
          }, [contentElement.value || '']));
          break;

        case 'tab':
          runChildren.push(XMLBuilder.wSelf('tab'));
          break;

        case 'break':
          {
            const attrs: Record<string, string> = {};
            if (contentElement.breakType) {
              attrs['w:type'] = contentElement.breakType;
            }
            runChildren.push(XMLBuilder.wSelf('br', Object.keys(attrs).length > 0 ? attrs : undefined));
          }
          break;

        case 'carriageReturn':
          runChildren.push(XMLBuilder.wSelf('cr'));
          break;

        case 'softHyphen':
          runChildren.push(XMLBuilder.wSelf('softHyphen'));
          break;

        case 'noBreakHyphen':
          runChildren.push(XMLBuilder.wSelf('noBreakHyphen'));
          break;
      }
    }

    return XMLBuilder.w('r', undefined, runChildren);
  }

  /**
   * Checks if the run has non-empty text content
   * @returns True if the run has text with length > 0
   */
  hasText(): boolean {
    const text = this.getText();
    return text.length > 0;
  }

  /**
   * Checks if the run has any formatting applied
   * @returns True if any formatting properties are set
   */
  hasFormatting(): boolean {
    return Object.keys(this.formatting).length > 0;
  }

  /**
   * Checks if the run is valid (has either text or formatting)
   * An empty run with no formatting is considered invalid
   * @returns True if the run has text or formatting
   */
  isValid(): boolean {
    return this.hasText() || this.hasFormatting();
  }

  /**
   * Gets the run content elements (text, tabs, breaks, etc.)
   * @returns Array of run content elements
   */
  getContent(): RunContent[] {
    return [...this.content];
  }

  /**
   * Adds a tab character to the run
   * Used in TOC entries to separate heading text from page numbers
   * @returns This run for method chaining
   */
  addTab(): this {
    this.content.push({ type: 'tab' });
    return this;
  }

  /**
   * Adds a line, page, or column break to the run
   * @param breakType - Type of break ('page' | 'column' | 'textWrapping')
   * @returns This run for method chaining
   */
  addBreak(breakType?: BreakType): this {
    this.content.push({ type: 'break', breakType });
    return this;
  }

  /**
   * Appends text to the run (adds new text element)
   * @param text - Text to append
   * @returns This run for method chaining
   */
  appendText(text: string): this {
    if (text) {
      this.content.push({ type: 'text', value: text });
    }
    return this;
  }

  /**
   * Adds a carriage return to the run
   * @returns This run for method chaining
   */
  addCarriageReturn(): this {
    this.content.push({ type: 'carriageReturn' });
    return this;
  }

  /**
   * Creates a Run from an array of content elements
   * Factory method for advanced use cases (used by DocumentParser)
   * @param content - Array of run content elements
   * @param formatting - Run formatting options
   * @returns New Run instance
   */
  static createFromContent(content: RunContent[], formatting: RunFormatting = {}): Run {
    const run = Object.create(Run.prototype) as Run;
    run.content = content;
    const { cleanXmlFromText, ...displayFormatting } = formatting;
    run.formatting = displayFormatting;
    return run;
  }

  /**
   * Generates run properties XML (<w:rPr>) from RunFormatting
   * This is a static helper used by both Run and Paragraph (for paragraph mark properties)
   *
   * Per ECMA-376 Part 1 §17.3.2.28, properties must be in specific order for strict compliance
   *
   * @param formatting - Run formatting options
   * @returns XMLElement representing <w:rPr> or null if no formatting
   */
  static generateRunPropertiesXML(formatting: RunFormatting): XMLElement | null {
    const rPrChildren: XMLElement[] = [];

    // 1. Character style reference (must be absolutely first per ECMA-376 §17.3.2.36)
    if (formatting.characterStyle) {
      rPrChildren.push(XMLBuilder.wSelf('rStyle', {
        'w:val': formatting.characterStyle,
      }));
    }

    // 2. Font family (must be second per ECMA-376 §17.3.2.28)
    if (formatting.font) {
      rPrChildren.push(XMLBuilder.wSelf('rFonts', {
        'w:ascii': formatting.font,
        'w:hAnsi': formatting.font,
        'w:cs': formatting.font,
      }));
    }

    // 2.5. Text border (w:bdr) per ECMA-376 Part 1 §17.3.2.5
    if (formatting.border) {
      const bdrAttrs: Record<string, string | number> = {};
      if (formatting.border.style) bdrAttrs['w:val'] = formatting.border.style;
      if (formatting.border.size !== undefined) bdrAttrs['w:sz'] = formatting.border.size;
      if (formatting.border.color) bdrAttrs['w:color'] = formatting.border.color;
      if (formatting.border.space !== undefined) bdrAttrs['w:space'] = formatting.border.space;

      if (Object.keys(bdrAttrs).length > 0) {
        rPrChildren.push(XMLBuilder.wSelf('bdr', bdrAttrs));
      }
    }

    // 3. Bold
    if (formatting.bold) {
      rPrChildren.push(XMLBuilder.wSelf('b', { 'w:val': '1' }));
    }

    // 3.5. Bold for complex scripts (w:bCs) per ECMA-376 Part 1 §17.3.2.3
    if (formatting.complexScriptBold) {
      rPrChildren.push(XMLBuilder.wSelf('bCs', { 'w:val': '1' }));
    }

    // 4. Italic
    if (formatting.italic) {
      rPrChildren.push(XMLBuilder.wSelf('i', { 'w:val': '1' }));
    }

    // 4.5. Italic for complex scripts (w:iCs) per ECMA-376 Part 1 §17.3.2.17
    if (formatting.complexScriptItalic) {
      rPrChildren.push(XMLBuilder.wSelf('iCs', { 'w:val': '1' }));
    }

    // 5. Capitalization (caps/smallCaps)
    if (formatting.allCaps) {
      rPrChildren.push(XMLBuilder.wSelf('caps', { 'w:val': '1' }));
    }
    if (formatting.smallCaps) {
      rPrChildren.push(XMLBuilder.wSelf('smallCaps', { 'w:val': '1' }));
    }

    // 6. Character shading (w:shd) per ECMA-376 Part 1 §17.3.2.32
    if (formatting.shading) {
      const shdAttrs: Record<string, string> = {};
      if (formatting.shading.val) shdAttrs['w:val'] = formatting.shading.val;
      if (formatting.shading.fill) shdAttrs['w:fill'] = formatting.shading.fill;
      if (formatting.shading.color) shdAttrs['w:color'] = formatting.shading.color;

      if (Object.keys(shdAttrs).length > 0) {
        rPrChildren.push(XMLBuilder.wSelf('shd', shdAttrs));
      }
    }

    // 6.5. Emphasis marks (w:em) per ECMA-376 Part 1 §17.3.2.13
    if (formatting.emphasis) {
      rPrChildren.push(XMLBuilder.wSelf('em', { 'w:val': formatting.emphasis }));
    }

    // 6.6. Outline text effect (w:outline) per ECMA-376 Part 1 §17.3.2.23
    if (formatting.outline) {
      rPrChildren.push(XMLBuilder.wSelf('outline', { 'w:val': '1' }));
    }

    // 6.7. Shadow text effect (w:shadow) per ECMA-376 Part 1 §17.3.2.32
    if (formatting.shadow) {
      rPrChildren.push(XMLBuilder.wSelf('shadow', { 'w:val': '1' }));
    }

    // 6.8. Emboss text effect (w:emboss) per ECMA-376 Part 1 §17.3.2.13
    if (formatting.emboss) {
      rPrChildren.push(XMLBuilder.wSelf('emboss', { 'w:val': '1' }));
    }

    // 6.9. Imprint/engrave text effect (w:imprint) per ECMA-376 Part 1 §17.3.2.18
    if (formatting.imprint) {
      rPrChildren.push(XMLBuilder.wSelf('imprint', { 'w:val': '1' }));
    }

    // 6.10. No proofing (w:noProof) per ECMA-376 Part 1 §17.3.2.21
    if (formatting.noProof) {
      rPrChildren.push(XMLBuilder.wSelf('noProof', { 'w:val': '1' }));
    }

    // 6.11. Snap to grid (w:snapToGrid) per ECMA-376 Part 1 §17.3.2.35
    if (formatting.snapToGrid) {
      rPrChildren.push(XMLBuilder.wSelf('snapToGrid', { 'w:val': '1' }));
    }

    // 6.12. Vanish/hidden (w:vanish) per ECMA-376 Part 1 §17.3.2.42
    if (formatting.vanish) {
      rPrChildren.push(XMLBuilder.wSelf('vanish', { 'w:val': '1' }));
    }

    // 6.12.5. Special vanish (w:specVanish) per ECMA-376 Part 1 §17.3.2.36
    if (formatting.specVanish) {
      rPrChildren.push(XMLBuilder.wSelf('specVanish', { 'w:val': '1' }));
    }

    // 6.13. RTL text (w:rtl) per ECMA-376 Part 1 §17.3.2.30
    // FIX: Must include w:val="1" to explicitly enable RTL, otherwise Word interprets empty tag incorrectly
    if (formatting.rtl) {
      rPrChildren.push(XMLBuilder.wSelf('rtl', { 'w:val': '1' }));
    }

    // 7. Strikethrough
    if (formatting.strike) {
      rPrChildren.push(XMLBuilder.wSelf('strike', { 'w:val': '1' }));
    }
    if (formatting.dstrike) {
      rPrChildren.push(XMLBuilder.wSelf('dstrike', { 'w:val': '1' }));
    }

    // 8. Underline
    if (formatting.underline) {
      const underlineValue = typeof formatting.underline === 'string'
        ? formatting.underline
        : 'single';
      rPrChildren.push(XMLBuilder.wSelf('u', { 'w:val': underlineValue }));
    }

    // 8.5. Character spacing (w:spacing) per ECMA-376 Part 1 §17.3.2.33
    if (formatting.characterSpacing !== undefined) {
      rPrChildren.push(XMLBuilder.wSelf('spacing', { 'w:val': formatting.characterSpacing }));
    }

    // 8.6. Horizontal scaling (w:w) per ECMA-376 Part 1 §17.3.2.43
    if (formatting.scaling !== undefined) {
      rPrChildren.push(XMLBuilder.wSelf('w', { 'w:val': formatting.scaling }));
    }

    // 8.7. Vertical position (w:position) per ECMA-376 Part 1 §17.3.2.31
    if (formatting.position !== undefined) {
      rPrChildren.push(XMLBuilder.wSelf('position', { 'w:val': formatting.position }));
    }

    // 8.8. Kerning (w:kern) per ECMA-376 Part 1 §17.3.2.20
    if (formatting.kerning !== undefined && formatting.kerning !== null) {
      rPrChildren.push(XMLBuilder.wSelf('kern', { 'w:val': formatting.kerning }));
    }

    // 8.9. Language (w:lang) per ECMA-376 Part 1 §17.3.2.20
    if (formatting.language) {
      rPrChildren.push(XMLBuilder.wSelf('lang', { 'w:val': formatting.language }));
    }

    // 8.9.5. East Asian layout (w:eastAsianLayout) per ECMA-376 Part 1 §17.3.2.10
    if (formatting.eastAsianLayout) {
      const layout = formatting.eastAsianLayout;
      const attrs: Record<string, string | number> = {};
      if (layout.id !== undefined) attrs['w:id'] = layout.id;
      if (layout.vert) attrs['w:vert'] = '1';
      if (layout.vertCompress) attrs['w:vertCompress'] = '1';
      if (layout.combine) attrs['w:combine'] = '1';
      if (layout.combineBrackets) attrs['w:combineBrackets'] = layout.combineBrackets;

      if (Object.keys(attrs).length > 0) {
        rPrChildren.push(XMLBuilder.wSelf('eastAsianLayout', attrs));
      }
    }

    // 8.10. Fit text to width (w:fitText) per ECMA-376 Part 1 §17.3.2.15
    if (formatting.fitText !== undefined) {
      rPrChildren.push(XMLBuilder.wSelf('fitText', { 'w:val': formatting.fitText }));
    }

    // 8.11. Text effect/animation (w:effect) per ECMA-376 Part 1 §17.3.2.12
    if (formatting.effect) {
      rPrChildren.push(XMLBuilder.wSelf('effect', { 'w:val': formatting.effect }));
    }

    // 9. Font size
    if (formatting.size !== undefined) {
      // Word uses half-points (size * 2)
      const halfPoints = formatting.size * 2;
      rPrChildren.push(XMLBuilder.wSelf('sz', { 'w:val': halfPoints }));
      rPrChildren.push(XMLBuilder.wSelf('szCs', { 'w:val': halfPoints }));
    }

    // 10. Text color
    if (formatting.color) {
      rPrChildren.push(XMLBuilder.wSelf('color', { 'w:val': formatting.color }));
    }

    // 11. Highlight color
    if (formatting.highlight) {
      rPrChildren.push(XMLBuilder.wSelf('highlight', { 'w:val': formatting.highlight }));
    }

    // 12. Vertical alignment (subscript/superscript) - must be last
    if (formatting.subscript) {
      rPrChildren.push(XMLBuilder.wSelf('vertAlign', { 'w:val': 'subscript' }));
    }
    if (formatting.superscript) {
      rPrChildren.push(XMLBuilder.wSelf('vertAlign', { 'w:val': 'superscript' }));
    }

    // Return null if no properties (prevents empty <w:rPr/> elements)
    if (rPrChildren.length === 0) {
      return null;
    }

    return XMLBuilder.w('rPr', undefined, rPrChildren);
  }

  /**
   * Creates a new Run with the specified text and formatting
   * @param text - Text content
   * @param formatting - Formatting options
   * @returns New Run instance
   */
  static create(text: string, formatting?: RunFormatting): Run {
    return new Run(text, formatting);
  }

  /**
   * Creates a deep clone of this run
   * @returns New Run instance with copied text and formatting
   * @example
   * ```typescript
   * const original = new Run('Hello', { bold: true });
   * const copy = original.clone();
   * copy.setText('World');  // Original unchanged
   * ```
   */
  clone(): Run {
    // Deep copy content and formatting to avoid shared references
    const clonedContent: RunContent[] = deepClone(this.content);
    const clonedFormatting: RunFormatting = deepClone(this.formatting);
    return Run.createFromContent(clonedContent, clonedFormatting);
  }

  /**
   * Inserts text at a specific position
   * NOTE: This flattens run content (loses tabs/breaks). Use with caution.
   * @param index - Position to insert at (0-based)
   * @param text - Text to insert
   * @returns This run for chaining
   * @example
   * ```typescript
   * const run = new Run('Hello World');
   * run.insertText(6, 'Beautiful ');  // "Hello Beautiful World"
   * ```
   */
  insertText(index: number, text: string): this {
    const currentText = this.getText();
    if (index < 0) index = 0;
    if (index > currentText.length) index = currentText.length;

    const newText = currentText.slice(0, index) + text + currentText.slice(index);
    this.setText(newText);
    return this;
  }

  /**
   * Replaces text in a specific range
   * NOTE: This flattens run content (loses tabs/breaks). Use with caution.
   * @param start - Start position (0-based, inclusive)
   * @param end - End position (0-based, exclusive)
   * @param text - Replacement text
   * @returns This run for chaining
   * @example
   * ```typescript
   * const run = new Run('Hello World');
   * run.replaceText(0, 5, 'Hi');  // "Hi World"
   * ```
   */
  replaceText(start: number, end: number, text: string): this {
    const currentText = this.getText();
    if (start < 0) start = 0;
    if (end > currentText.length) end = currentText.length;
    if (start > end) [start, end] = [end, start];  // Swap if reversed

    const newText = currentText.slice(0, start) + text + currentText.slice(end);
    this.setText(newText);
    return this;
  }

  /**
   * Clears run formatting properties that conflict with a style definition.
   * Uses smart clearing: only removes properties that DIFFER from the style.
   * Preserves properties not defined in the style (e.g., if style doesn't define bold, keeps run's bold).
   *
   * This is critical for allowing style definitions to apply properly, as direct formatting
   * in document.xml ALWAYS overrides style definitions in styles.xml per ECMA-376 §17.7.2.
   *
   * @param styleRunFormatting - Run formatting from the style definition to compare against
   * @returns This run for method chaining
   * @example
   * ```typescript
   * // Style says: black, 14pt Verdana
   * // Run has: red, 12pt Arial, bold
   * run.clearFormattingConflicts({
   *   color: '000000',
   *   size: 14,
   *   font: 'Verdana'
   * });
   * // Result: Run keeps bold (not in style), but color/size/font are cleared
   * ```
   */
  clearFormattingConflicts(styleRunFormatting: RunFormatting): this {
    // For each property in run's formatting, check if it conflicts with style
    const conflictingProperties: (keyof RunFormatting)[] = [];

    // Compare each property
    for (const key in this.formatting) {
      const propKey = key as keyof RunFormatting;

      // Skip if style doesn't define this property (preserve run's property)
      if (styleRunFormatting[propKey] === undefined) {
        continue;
      }

      // If style defines this property AND run's value differs, it's a conflict
      if (this.formatting[propKey] !== styleRunFormatting[propKey]) {
        conflictingProperties.push(propKey);
      }
    }

    // Clear conflicting properties
    for (const prop of conflictingProperties) {
      delete this.formatting[prop];
    }

    return this;
  }

  /**
   * Clears all formatting from this run
   *
   * Removes all direct formatting properties, leaving only the text content.
   * This is useful for applying clean styles without formatting overrides.
   *
   * @returns This run for chaining
   * @example
   * ```typescript
   * run.clearFormatting();
   * ```
   */
  clearFormatting(): this {
    this.formatting = {};
    return this;
  }
}
