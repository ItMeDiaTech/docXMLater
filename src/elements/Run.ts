/**
 * Run - Represents a run of text with uniform formatting
 * A run is the smallest unit of text formatting in a Word document
 */

import { deepClone } from "../utils/deepClone";
import { formatDateForXml } from "../utils/dateFormatting";
import { logSerialization, logTextDirection } from "../utils/diagnostics";
import { defaultLogger } from "../utils/logger";
import { normalizeColor, validateRunText } from "../utils/validation";
import { pointsToHalfPoints } from "../utils/units";
import { diffText, diffHasUnchangedParts } from "../utils/textDiff";
import { getActiveConditionalsInPriorityOrder } from "../utils/cnfStyleDecoder";
import { XMLBuilder, XMLElement } from "../xml/XMLBuilder";
import {
  ShadingPattern as CommonShadingPattern,
  ShadingConfig,
  buildShadingAttributes,
  ExtendedBorderStyle,
  BorderDefinition,
} from "./CommonTypes";
import type { RunPropertyChange } from "./PropertyChangeTypes";
// Type-only import to avoid circular dependency (Revision imports Run)
import type { Revision as RevisionType } from "./Revision";

/**
 * Run content element types
 * Per ECMA-376 Part 1 §17.3.3 EG_RunInnerContent, runs can contain multiple types of content
 */
export type RunContentType =
  | "text" // <w:t> - Regular text
  | "tab" // <w:tab/> - Tab character (used in TOC entries)
  | "break" // <w:br/> - Line/page/column break
  | "carriageReturn" // <w:cr/> - Carriage return
  | "softHyphen" // <w:softHyphen/> - Optional hyphen
  | "noBreakHyphen" // <w:noBreakHyphen/> - Non-breaking hyphen
  | "instructionText" // <w:instrText> - Field instruction text
  | "fieldChar" // <w:fldChar/> - Field character markers
  | "vml" // <w:pict> - VML/legacy graphics (preserved as raw XML)
  | "lastRenderedPageBreak" // <w:lastRenderedPageBreak/> - Last rendered page break position
  | "separator" // <w:separator/> - Footnote/endnote separator line
  | "continuationSeparator" // <w:continuationSeparator/> - Continuation separator line
  | "pageNumber" // <w:pgNum/> - Page number field
  | "annotationRef" // <w:annotationRef/> - Annotation reference marker
  | "dayShort" // <w:dayShort/> - Short date day field
  | "dayLong" // <w:dayLong/> - Long date day field
  | "monthShort" // <w:monthShort/> - Short month field
  | "monthLong" // <w:monthLong/> - Long month field
  | "yearShort" // <w:yearShort/> - Short year field
  | "yearLong" // <w:yearLong/> - Long year field
  | "symbol" // <w:sym/> - Symbol character with font and char code
  | "positionTab" // <w:ptab/> - Absolute position tab
  | "embeddedObject" // <w:object> - Embedded OLE object (preserved as raw XML)
  | "footnoteReference" // <w:footnoteReference/> - Footnote reference marker
  | "endnoteReference"; // <w:endnoteReference/> - Endnote reference marker

/**
 * Break type for <w:br> elements
 * Per ECMA-376 Part 1 §17.18.3
 */
export type BreakType = "page" | "column" | "textWrapping";

/**
 * Run content element
 * Represents a single content element within a run (text, tab, break, etc.)
 */
export interface RunContent {
  /** Type of content element */
  type: RunContentType;
  /** Text value (for 'text' and 'instructionText' types) */
  value?: string;
  /** Break type (only for 'break' type) */
  breakType?: BreakType;
  /** Field character subtype (only for 'fieldChar' type) */
  fieldCharType?: "begin" | "separate" | "end";
  /** Whether the field char is marked dirty */
  fieldCharDirty?: boolean;
  /** Whether the field char is locked */
  fieldCharLocked?: boolean;
  /** Form field data (only for 'fieldChar' type with fieldCharType='begin') per ECMA-376 Part 1 §17.16.17 */
  formFieldData?: FormFieldData;
  /** Raw XML content (for 'vml' type - preserves w:pict elements as-is) */
  rawXml?: string;
  /**
   * Whether this content came from a deleted section (w:delText or w:delInstrText)
   * Per ECMA-376 Part 1 §22.1.2.26-27, deleted content uses special elements
   * This flag helps with proper serialization back to w:delText/w:delInstrText
   */
  isDeleted?: boolean;
  /** Symbol font name (only for 'symbol' type, w:sym w:font) */
  symbolFont?: string;
  /** Symbol character code (only for 'symbol' type, w:sym w:char) */
  symbolChar?: string;
  /** Position tab alignment (only for 'positionTab' type, w:ptab w:alignment) */
  ptabAlignment?: string;
  /** Position tab relative-to (only for 'positionTab' type, w:ptab w:relativeTo) */
  ptabRelativeTo?: string;
  /** Position tab leader character (only for 'positionTab' type, w:ptab w:leader) */
  ptabLeader?: string;
  /** Footnote ID (only for 'footnoteReference' type, w:footnoteReference w:id) */
  footnoteId?: number;
  /** Endnote ID (only for 'endnoteReference' type, w:endnoteReference w:id) */
  endnoteId?: number;
}

/**
 * Border style for text
 * @see CommonTypes.ExtendedBorderStyle for the canonical definition
 */
export type TextBorderStyle = ExtendedBorderStyle;

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
 * @see CommonTypes.ShadingPattern for the canonical definition
 */
export type ShadingPattern = CommonShadingPattern;

/**
 * Character shading definition
 * @see ShadingConfig in CommonTypes.ts for the canonical definition
 */
export type CharacterShading = ShadingConfig;

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
  combineBrackets?: "none" | "round" | "square" | "angle" | "curly";
}

/**
 * Form field text input data per ECMA-376 Part 1 §17.16.33
 */
export interface FormFieldTextInput {
  type: 'textInput';
  /** Input type per ECMA-376 §17.16.34 (e.g., 'regular', 'number', 'date', 'currentDate', 'currentTime', 'calculated') */
  inputType?: string;
  /** Default text value */
  defaultValue?: string;
  /** Maximum length (0 = unlimited) */
  maxLength?: number;
  /** Text format (e.g., 'UPPERCASE', 'Lowercase', 'First capital') */
  format?: string;
}

/**
 * Form field checkbox data per ECMA-376 Part 1 §17.16.7
 */
export interface FormFieldCheckBox {
  type: 'checkBox';
  /** Default state */
  defaultChecked?: boolean;
  /** Current checked state */
  checked?: boolean;
  /** Size (auto or specific in half-points) */
  size?: number | 'auto';
}

/**
 * Form field dropdown list data per ECMA-376 Part 1 §17.16.12
 */
export interface FormFieldDropDownList {
  type: 'dropDownList';
  /** Selected item index (0-based) */
  result?: number;
  /** Default item index */
  defaultResult?: number;
  /** List entries */
  listEntries?: string[];
}

/**
 * Form field data per ECMA-376 Part 1 §17.16.17
 */
export interface FormFieldData {
  /** Field name */
  name?: string;
  /** Whether the field is enabled */
  enabled?: boolean;
  /** Calculate on exit */
  calcOnExit?: boolean;
  /** Help text */
  helpText?: string;
  /** Status bar text */
  statusText?: string;
  /** Entry macro name */
  entryMacro?: string;
  /** Exit macro name */
  exitMacro?: string;
  /** Field-specific data */
  fieldType?: FormFieldTextInput | FormFieldCheckBox | FormFieldDropDownList;
}

/**
 * Emphasis mark type - decorative marks above/below text
 */
export type EmphasisMark = "dot" | "comma" | "circle" | "underDot";

/**
 * Theme color values per ECMA-376 Part 1 Section 17.18.96 (ST_ThemeColor)
 * These reference colors defined in the document's theme (theme1.xml)
 */
export type ThemeColorValue =
  | "dark1"
  | "light1"
  | "dark2"
  | "light2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hyperlink"
  | "followedHyperlink"
  | "background1"
  | "text1"
  | "background2"
  | "text2";

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
  /** Underline text. Use "none" to explicitly override style underline. */
  underline?: boolean | "single" | "double" | "thick" | "dotted" | "dash" | "none";
  /** Underline color in hex format (without #) per ECMA-376 Part 1 §17.3.2.40 */
  underlineColor?: string;
  /** Underline theme color reference per ECMA-376 Part 1 §17.3.2.40 */
  underlineThemeColor?: ThemeColorValue;
  /** Underline theme tint (0-255) per ECMA-376 Part 1 §17.3.2.40 */
  underlineThemeTint?: number;
  /** Underline theme shade (0-255) per ECMA-376 Part 1 §17.3.2.40 */
  underlineThemeShade?: number;
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
  /** Font size for complex scripts (RTL languages) in points. If not set, uses size. */
  sizeCs?: number;
  /** Text color in hex format (without #) */
  color?: string;
  /**
   * Theme color reference for text color per ECMA-376 Part 1 Section 17.3.2.6
   * When set, the color is derived from the document's theme rather than a fixed hex value
   */
  themeColor?: ThemeColorValue;
  /**
   * Theme color tint (0-255, where 0=no tint, 255=full tint toward white)
   * Applied to themeColor to create lighter variations
   */
  themeTint?: number;
  /**
   * Theme color shade (0-255, where 0=no shade, 255=full shade toward black)
   * Applied to themeColor to create darker variations
   */
  themeShade?: number;
  /** Highlight color */
  highlight?:
    | "yellow"
    | "green"
    | "cyan"
    | "magenta"
    | "blue"
    | "red"
    | "darkBlue"
    | "darkCyan"
    | "darkGreen"
    | "darkMagenta"
    | "darkRed"
    | "darkYellow"
    | "darkGray"
    | "lightGray"
    | "black"
    | "white";
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
  effect?:
    | "none"
    | "lights"
    | "blinkBackground"
    | "sparkleText"
    | "marchingBlackAnts"
    | "marchingRedAnts"
    | "shimmer"
    | "antsBlack"
    | "antsRed";
  /** Fit text to width in twips (1/20th of a point) */
  fitText?: number;
  /** East Asian typography layout options */
  eastAsianLayout?: EastAsianLayout;
  /** Complex script formatting flag (w:cs) per ECMA-376 Part 1 §17.3.2.7 */
  complexScript?: boolean;
  /** Web hidden - hide text in web layout view (w:webHidden) per ECMA-376 Part 1 §17.3.2.44 */
  webHidden?: boolean;
  /** High ANSI font (w:rFonts w:hAnsi) - font for characters in the high ANSI range */
  fontHAnsi?: string;
  /** East Asian font (w:rFonts w:eastAsia) - font for East Asian characters */
  fontEastAsia?: string;
  /** Complex script font (w:rFonts w:cs) - font for complex script (RTL) characters */
  fontCs?: string;
  /** Font hint (w:rFonts w:hint) - hint for font selection: 'default' | 'eastAsia' | 'cs' */
  fontHint?: string;
  /** ASCII theme font reference (w:rFonts w:asciiTheme) per ECMA-376 Part 1 §17.3.2.26 */
  fontAsciiTheme?: string;
  /** High ANSI theme font reference (w:rFonts w:hAnsiTheme) per ECMA-376 Part 1 §17.3.2.26 */
  fontHAnsiTheme?: string;
  /** East Asian theme font reference (w:rFonts w:eastAsiaTheme) per ECMA-376 Part 1 §17.3.2.26 */
  fontEastAsiaTheme?: string;
  /** Complex script theme font reference (w:rFonts w:cstheme) per ECMA-376 Part 1 §17.3.2.26 */
  fontCsTheme?: string;
  /**
   * Raw w14: namespace elements from rPr (Word 2010+ text effects)
   * Stored as raw XML strings for passthrough round-trip fidelity.
   * Includes: w14:textOutline, w14:shadow, w14:reflection, w14:glow,
   * w14:ligatures, w14:numForm, w14:numSpacing, w14:cntxtAlts, w14:stylisticSets
   */
  rawW14Properties?: string[];
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
  private trackingContext?: import('../tracking/TrackingContext').TrackingContext;
  private propertyChangeRevision?: RunPropertyChange;
  /** Parent paragraph reference for automatic tracking */
  private _parentParagraph?: import('./Paragraph').Paragraph;

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
          `  - Received ${
            text === undefined ? "undefined" : "null"
          } text value\n` +
          `  - Converting to empty string for Word compatibility`
      );
    }

    // Default to auto-cleaning XML patterns unless explicitly disabled
    const shouldClean = formatting.cleanXmlFromText !== false;

    // Validate text for XML patterns
    const validation = validateRunText(text, {
      context: "Run constructor",
      autoClean: shouldClean,
      warnToConsole: true, // Enable warnings to help catch data quality issues
    });

    // Use cleaned text if available and cleaning was requested
    const cleanedText = validation.cleanedText || text;

    // Convert undefined/null to empty string for consistent XML generation
    const normalizedText = cleanedText ?? "";

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
      return [{ type: "text", value: "" }];
    }

    const content: RunContent[] = [];
    let currentText = "";

    for (let i = 0; i < text.length; i++) {
      const char = text[i];

      switch (char) {
        case "\t":
          // Add accumulated text before tab
          if (currentText) {
            content.push({ type: "text", value: currentText });
            currentText = "";
          }
          // Add tab element
          content.push({ type: "tab" });
          break;

        case "\n":
          // Add accumulated text before newline
          if (currentText) {
            content.push({ type: "text", value: currentText });
            currentText = "";
          }
          // Add break element
          content.push({ type: "break" });
          break;

        case "\u2011": // Non-breaking hyphen
          // Add accumulated text before non-breaking hyphen
          if (currentText) {
            content.push({ type: "text", value: currentText });
            currentText = "";
          }
          // Add non-breaking hyphen element
          content.push({ type: "noBreakHyphen" });
          break;

        case "\r": // Carriage return
          // Add accumulated text before carriage return
          if (currentText) {
            content.push({ type: "text", value: currentText });
            currentText = "";
          }
          // Add carriage return element
          content.push({ type: "carriageReturn" });
          break;

        case "\u00AD": // Soft hyphen
          // Add accumulated text before soft hyphen
          if (currentText) {
            content.push({ type: "text", value: currentText });
            currentText = "";
          }
          // Add soft hyphen element
          content.push({ type: "softHyphen" });
          break;

        default:
          // Accumulate regular text
          currentText += char;
          break;
      }
    }

    // Add any remaining text
    if (currentText) {
      content.push({ type: "text", value: currentText });
    }

    // If no content was added (empty string), add empty text element
    if (content.length === 0) {
      content.push({ type: "text", value: "" });
    }

    return content;
  }

  /**
   * Gets the text content from the run
   *
   * Concatenates all content elements and converts special characters
   * (tabs, breaks, etc.) back to their string representation.
   *
   * @returns The complete text string including tabs (\t) and line breaks (\n)
   *
   * @example
   * ```typescript
   * const run = new Run('Hello\tWorld');
   * console.log(run.getText()); // "Hello\tWorld"
   * ```
   */
  getText(): string {
    return this.content
      .map((c) => {
        switch (c.type) {
          case "text":
            return c.value || "";
          case "tab":
            return "\t";
          case "break":
            return "\n";
          case "carriageReturn":
            return "\r";
          case "softHyphen":
            return "\u00AD";
          case "noBreakHyphen":
            return "\u2011";
          case "instructionText":
            return "";
          case "fieldChar":
            return "";
          default:
            return "";
        }
      })
      .join("");
  }

  /**
   * Sets the text content of the run
   *
   * Replaces all existing content with new text. Special characters like
   * tabs (\t) and newlines (\n) are automatically converted to their
   * corresponding XML elements.
   *
   * @remarks
   * When change tracking is enabled and this run has a parent paragraph,
   * calling `setText()` will replace this run in the paragraph's content
   * array with multiple new elements (unchanged runs + revision wrappers).
   * After this call, the original run reference may no longer be present
   * in the paragraph — callers should re-query the paragraph content
   * rather than continuing to use the original run reference.
   *
   * @param text - The new text content
   *
   * @example
   * ```typescript
   * const run = new Run('Old text');
   * run.setText('New text');
   * run.setText('Text with\ttab'); // Tab is preserved
   * ```
   */
  setText(text: string): void {
    // Capture old text and content for tracking before any changes
    const oldText = this.getText();
    const oldContent = [...this.content];

    // Warn about undefined/null text to help catch data quality issues
    if (text === undefined || text === null) {
      defaultLogger.warn(
        `DocXML Text Validation Warning [Run.setText]:\n` +
          `  - Received ${
            text === undefined ? "undefined" : "null"
          } text value\n` +
          `  - Converting to empty string for Word compatibility`
      );
    }

    // Respect original cleanXmlFromText setting (Issue #9 fix)
    // This ensures consistent behavior with constructor
    const shouldClean = this.formatting.cleanXmlFromText !== false;

    // Validate text for XML patterns
    const validation = validateRunText(text, {
      context: "Run.setText",
      autoClean: shouldClean,
      warnToConsole: true, // Enable warnings to help catch data quality issues
    });

    // Use cleaned text if available and cleaning was requested
    const cleanedText = validation.cleanedText || text;

    // Convert undefined/null to empty string for consistent XML generation
    const normalizedText = cleanedText ?? "";

    // Check if current content is all instructionText - preserve the type
    // This is critical for TOC field instructions and other field codes
    // Without this, calling setText() on a run with instrText would convert it to w:t
    // which causes field instruction codes to appear as visible text in Word
    const isAllInstructionText = this.content.length > 0 &&
      this.content.every(c => c.type === "instructionText");

    if (isAllInstructionText) {
      // Preserve instructionText type - just update the value
      this.content = [{ type: "instructionText", value: normalizedText }];
    } else {
      // Normal behavior - parse text to extract special characters into separate content elements
      this.content = this.parseTextWithSpecialCharacters(normalizedText);
    }

    // Track text change if tracking is enabled and text actually changed
    if (this.trackingContext?.isEnabled() && this._parentParagraph && oldText !== normalizedText && oldText) {
      // Check if original content had non-text types that getText() can't faithfully represent
      // (VML, fieldChar, symbol, embeddedObject) — if so, fall back to whole-run replacement
      const hasNonTextContent = oldContent.some(c =>
        c.type === 'vml' || c.type === 'fieldChar' || c.type === 'symbol' || c.type === 'embeddedObject'
      );

      const segments = hasNonTextContent ? [] : diffText(oldText, normalizedText);
      const useGranular = !hasNonTextContent && diffHasUnchangedParts(segments);

      if (useGranular) {
        // Fine-grained tracking: split into unchanged runs + delete/insert revisions
        const revManager = this.trackingContext.getRevisionManager();
        const now = new Date();
        const newContent: (Run | RevisionType)[] = [];

        for (const seg of segments) {
          if (seg.type === 'equal') {
            // Create a new run with the same formatting for the unchanged portion
            const equalRun = this.clone();
            equalRun.content = this.parseTextWithSpecialCharacters(seg.text);
            equalRun._setTrackingContext(this.trackingContext);
            equalRun._setParentParagraph(this._parentParagraph);
            newContent.push(equalRun);
          } else if (seg.type === 'delete') {
            const delRun = this.clone();
            delRun.content = this.parseTextWithSpecialCharacters(seg.text);
            const delRev = this.trackingContext.createDeletion(delRun, now);
            revManager.register(delRev);
            newContent.push(delRev);
          } else if (seg.type === 'insert') {
            const insRun = this.clone();
            insRun.content = this.parseTextWithSpecialCharacters(seg.text);
            const insRev = this.trackingContext.createInsertion(insRun, now);
            revManager.register(insRev);
            newContent.push(insRev);
          }
        }

        this._parentParagraph.replaceContent(this, newContent);
      } else {
        // Whole-run fallback: complete replacement (no shared text, or non-text content)
        const revManager = this.trackingContext.getRevisionManager();
        const now = new Date();

        const deleteRun = this.clone();
        deleteRun.content = this.parseTextWithSpecialCharacters(oldText);

        const deleteRev = this.trackingContext.createDeletion(deleteRun, now);
        revManager.register(deleteRev);

        const insertRev = this.trackingContext.createInsertion(this, now);
        revManager.register(insertRev);

        this._parentParagraph.replaceContent(this, [deleteRev, insertRev]);
      }
    }
  }

  /**
   * Gets a copy of the run formatting
   *
   * Returns a copy of all formatting properties including font, size,
   * bold, italic, color, and all other formatting attributes.
   *
   * @returns Copy of the run formatting object
   *
   * @example
   * ```typescript
   * const formatting = run.getFormatting();
   * console.log(`Font: ${formatting.font}, Size: ${formatting.size}pt`);
   * if (formatting.bold) {
   *   console.log('Text is bold');
   * }
   * ```
   */
  getFormatting(): RunFormatting {
    return { ...this.formatting };
  }

  /**
   * Gets the bold formatting value
   * @returns True if bold, false otherwise
   */
  getBold(): boolean {
    return this.formatting.bold ?? false;
  }

  /**
   * Gets the italic formatting value
   * @returns True if italic, false otherwise
   */
  getItalic(): boolean {
    return this.formatting.italic ?? false;
  }

  /**
   * Gets the underline style
   * @returns Underline style (string, boolean, or undefined)
   */
  getUnderline(): boolean | "none" | "single" | "double" | "dotted" | "thick" | "dash" | undefined {
    return this.formatting.underline;
  }

  /**
   * Gets the strikethrough formatting value
   * @returns True if strikethrough, false otherwise
   */
  getStrike(): boolean {
    return this.formatting.strike ?? false;
  }

  /**
   * Gets the font family name
   * @returns Font name or undefined if not set
   */
  getFont(): string | undefined {
    return this.formatting.font;
  }

  /**
   * Gets the font size in half-points
   * @returns Size in half-points or undefined if not set
   */
  getSize(): number | undefined {
    return this.formatting.size;
  }

  /**
   * Gets the text color as hex string
   * @returns Color hex string or undefined if not set
   */
  getColor(): string | undefined {
    return this.formatting.color;
  }

  /**
   * Gets the highlight color
   * @returns Highlight color name or undefined if not set
   */
  getHighlight(): string | undefined {
    return this.formatting.highlight;
  }

  /**
   * Gets the subscript formatting value
   * @returns True if subscript, false otherwise
   */
  getSubscript(): boolean {
    return this.formatting.subscript ?? false;
  }

  /**
   * Gets the superscript formatting value
   * @returns True if superscript, false otherwise
   */
  getSuperscript(): boolean {
    return this.formatting.superscript ?? false;
  }

  /**
   * Gets whether the run is right-to-left text
   * @returns True if RTL, false otherwise
   */
  isRTL(): boolean {
    return this.formatting.rtl ?? false;
  }

  /**
   * Gets the small caps formatting value
   * @returns True if small caps, false otherwise
   */
  getSmallCaps(): boolean {
    return this.formatting.smallCaps ?? false;
  }

  /**
   * Gets the all caps formatting value
   * @returns True if all caps, false otherwise
   */
  getAllCaps(): boolean {
    return this.formatting.allCaps ?? false;
  }

  /**
   * Gets effective bold formatting, resolving from:
   * 1. Direct formatting on this run
   * 2. Table conditional formatting (if in a table cell)
   *
   * @returns True if text should render bold, undefined if not determinable
   *
   * @example
   * ```typescript
   * // Check effective bold (considers table conditional formatting)
   * const isBold = run.getEffectiveBold();
   * if (isBold) {
   *   console.log('Text appears bold (direct or via table style)');
   * }
   * ```
   */
  getEffectiveBold(): boolean | undefined {
    // Direct formatting takes highest precedence
    if (this.formatting.bold !== undefined) {
      return this.formatting.bold;
    }

    // Check table conditional formatting
    return this.getConditionalFormattingProperty("bold");
  }

  /**
   * Gets effective italic formatting, resolving from:
   * 1. Direct formatting on this run
   * 2. Table conditional formatting (if in a table cell)
   *
   * @returns True if text should render italic, undefined if not determinable
   */
  getEffectiveItalic(): boolean | undefined {
    if (this.formatting.italic !== undefined) {
      return this.formatting.italic;
    }
    return this.getConditionalFormattingProperty("italic");
  }

  /**
   * Gets effective color formatting, resolving from:
   * 1. Direct formatting on this run
   * 2. Table conditional formatting (if in a table cell)
   *
   * @returns Hex color string or undefined
   */
  getEffectiveColor(): string | undefined {
    if (this.formatting.color !== undefined) {
      return this.formatting.color;
    }
    return this.getConditionalFormattingProperty("color");
  }

  /**
   * Gets effective font formatting, resolving from:
   * 1. Direct formatting on this run
   * 2. Table conditional formatting (if in a table cell)
   *
   * @returns Font name or undefined
   */
  getEffectiveFont(): string | undefined {
    if (this.formatting.font !== undefined) {
      return this.formatting.font;
    }
    return this.getConditionalFormattingProperty("font");
  }

  /**
   * Gets effective font size formatting, resolving from:
   * 1. Direct formatting on this run
   * 2. Table conditional formatting (if in a table cell)
   *
   * @returns Font size in points or undefined
   */
  getEffectiveSize(): number | undefined {
    if (this.formatting.size !== undefined) {
      return this.formatting.size;
    }
    return this.getConditionalFormattingProperty("size");
  }

  /**
   * Checks if this run has the "Hyperlink" character style applied.
   *
   * This is useful when applying bulk style changes to a document,
   * to avoid overwriting hyperlink formatting (which would make
   * hyperlinks appear as plain text).
   *
   * @returns True if the run has "Hyperlink" character style
   *
   * @example
   * ```typescript
   * // Skip hyperlink-styled runs when applying color changes
   * for (const para of doc.getParagraphs()) {
   *   for (const run of para.getRuns()) {
   *     if (run.isHyperlinkStyled()) {
   *       continue; // Preserve hyperlink formatting
   *     }
   *     run.setColor('000000');
   *   }
   * }
   * ```
   */
  isHyperlinkStyled(): boolean {
    return this.formatting.characterStyle === "Hyperlink";
  }

  /**
   * Gets a specific property from table conditional formatting.
   * Traverses the parent chain to find conditional formatting from table styles.
   * @internal
   */
  private getConditionalFormattingProperty<K extends keyof RunFormatting>(
    property: K
  ): RunFormatting[K] | undefined {
    // Need parent paragraph to access cell context
    const para = this._parentParagraph;
    if (!para) return undefined;

    // Check if paragraph is in a table cell
    const cell = para._getParentCell();
    if (!cell) return undefined;

    // Get the cnfStyle (which conditionals apply)
    const cnfStyle = para.getTableConditionalStyle();
    if (!cnfStyle) return undefined;

    // Decode cnfStyle to find active conditionals
    const activeConditionals = getActiveConditionalsInPriorityOrder(cnfStyle);

    // Try explicit table style first
    const tableStyleId = cell.getTableStyleId();
    const stylesManager = para._getStylesManager();

    if (tableStyleId && stylesManager) {
      const style = stylesManager.getStyle(tableStyleId);
      if (style) {
        const tableStyleProps = style.getProperties().tableStyle;
        if (tableStyleProps?.conditionalFormatting) {
          // Check active conditionals in priority order (corners > edges > banding)
          for (const conditionalType of activeConditionals) {
            const conditional = tableStyleProps.conditionalFormatting.find(
              (cf) => cf.type === conditionalType
            );
            if (conditional?.runFormatting?.[property] !== undefined) {
              return conditional.runFormatting[property];
            }
          }

          // Fallback: check wholeTable conditional (applies to all cells)
          const wholeTable = tableStyleProps.conditionalFormatting.find(
            (cf) => cf.type === "wholeTable"
          );
          if (wholeTable?.runFormatting?.[property] !== undefined) {
            return wholeTable.runFormatting[property];
          }
        }
      }
    }

    // Apply Word's default formatting when no explicit style
    // Per Word behavior: firstRow, firstCol, lastRow, lastCol get bold by default
    if (property === "bold" && activeConditionals.length > 0) {
      const table = cell._getParentRow()?._getParentTable();
      if (table) {
        const tblLookFlags = table.getTblLookFlags();

        // Check if cell's conditional matches enabled tblLook flags
        for (const conditionalType of activeConditionals) {
          if (conditionalType === "firstRow" && tblLookFlags.firstRow)
            return true as RunFormatting[K];
          if (conditionalType === "lastRow" && tblLookFlags.lastRow)
            return true as RunFormatting[K];
          if (conditionalType === "firstCol" && tblLookFlags.firstColumn)
            return true as RunFormatting[K];
          if (conditionalType === "lastCol" && tblLookFlags.lastColumn)
            return true as RunFormatting[K];
        }
      }
    }

    return undefined;
  }

  /**
   * Sets character style reference
   * Per ECMA-376 Part 1 §17.3.2.36
   * @param styleId - Character style ID to apply
   * @returns This run for chaining
   */
  setCharacterStyle(styleId: string): this {
    const previousValue = this.formatting.characterStyle;
    this.formatting.characterStyle = styleId;
    if (this.trackingContext?.isEnabled() && previousValue !== styleId) {
      this.trackingContext.trackRunPropertyChange(this, 'characterStyle', previousValue, styleId);
    }
    return this;
  }

  /**
   * Sets text border
   * Per ECMA-376 Part 1 §17.3.2.5
   * @param border - Border definition
   * @returns This run for chaining
   */
  setBorder(border: TextBorder): this {
    const previousValue = this.formatting.border;
    this.formatting.border = border;
    if (this.trackingContext?.isEnabled() && previousValue !== border) {
      this.trackingContext.trackRunPropertyChange(this, 'border', previousValue, border);
    }
    return this;
  }

  /**
   * Sets character shading (background)
   * Per ECMA-376 Part 1 §17.3.2.32
   * @param shading - Shading definition
   * @returns This run for chaining
   */
  setShading(shading: CharacterShading): this {
    const previousValue = this.formatting.shading;
    this.formatting.shading = shading;
    if (this.trackingContext?.isEnabled() && previousValue !== shading) {
      this.trackingContext.trackRunPropertyChange(this, 'shading', previousValue, shading);
    }
    return this;
  }

  /**
   * Sets emphasis mark
   * Per ECMA-376 Part 1 §17.3.2.13
   * @param emphasis - Emphasis mark type ('dot', 'comma', 'circle', 'underDot')
   * @returns This run for chaining
   */
  setEmphasis(emphasis: EmphasisMark): this {
    const previousValue = this.formatting.emphasis;
    this.formatting.emphasis = emphasis;
    if (this.trackingContext?.isEnabled() && previousValue !== emphasis) {
      this.trackingContext.trackRunPropertyChange(this, 'emphasis', previousValue, emphasis);
    }
    return this;
  }

  /**
   * Sets the tracking context for automatic change tracking.
   * Called by Document when track changes is enabled.
   * @internal
   */
  _setTrackingContext(context: import('../tracking/TrackingContext').TrackingContext): void {
    this.trackingContext = context;
  }

  /**
   * Sets the parent paragraph reference for this run
   * @internal Used for content tracking
   */
  _setParentParagraph(paragraph: import('./Paragraph').Paragraph): void {
    this._parentParagraph = paragraph;
  }

  /**
   * Gets the parent paragraph reference
   * @internal Used for content tracking
   */
  _getParentParagraph(): import('./Paragraph').Paragraph | undefined {
    return this._parentParagraph;
  }

  /**
   * Gets the property change revision for this run (if any)
   *
   * Property change revisions (w:rPrChange) track formatting changes to runs.
   * They contain the PREVIOUS properties before the change was made.
   *
   * @returns Property change revision or undefined if not set
   *
   * @example
   * ```typescript
   * const propChange = run.getPropertyChangeRevision();
   * if (propChange) {
   *   console.log(`Changed by ${propChange.author} on ${propChange.date}`);
   *   console.log(`Previous: ${JSON.stringify(propChange.previousProperties)}`);
   * }
   * ```
   */
  getPropertyChangeRevision(): RunPropertyChange | undefined {
    return this.propertyChangeRevision;
  }

  /**
   * Sets the property change revision for this run
   *
   * Property change revisions (w:rPrChange) are stored inside the run properties
   * element (w:rPr) and track what the previous formatting was before a change.
   * This is used for round-trip preservation of tracked changes.
   *
   * @param propChange - Property change revision data
   * @returns This run for method chaining
   *
   * @example
   * ```typescript
   * run.setPropertyChangeRevision({
   *   id: 1,
   *   author: 'John Doe',
   *   date: new Date(),
   *   previousProperties: { bold: true }
   * });
   * ```
   */
  setPropertyChangeRevision(propChange: RunPropertyChange): this {
    this.propertyChangeRevision = propChange;
    return this;
  }

  /**
   * Clears the property change revision for this run
   *
   * @returns This run for method chaining
   */
  clearPropertyChangeRevision(): this {
    this.propertyChangeRevision = undefined;
    return this;
  }

  /**
   * Checks if this run has a property change revision
   *
   * @returns True if a property change revision is set
   */
  hasPropertyChangeRevision(): boolean {
    return this.propertyChangeRevision !== undefined;
  }

  /**
   * Sets bold formatting
   *
   * Makes the text bold or removes bold formatting.
   *
   * @param bold - If true, applies bold; if false, removes bold (default: true)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setBold();        // Apply bold
   * run.setBold(true);    // Apply bold
   * run.setBold(false);   // Remove bold
   * ```
   */
  setBold(bold = true): this {
    const previousValue = this.formatting.bold;
    this.formatting.bold = bold;
    if (this.trackingContext?.isEnabled() && previousValue !== bold) {
      this.trackingContext.trackRunPropertyChange(this, 'bold', previousValue, bold);
    }
    return this;
  }

  /**
   * Sets italic formatting
   *
   * Makes the text italic or removes italic formatting.
   *
   * @param italic - If true, applies italic; if false, removes italic (default: true)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setItalic();       // Apply italic
   * run.setItalic(true);   // Apply italic
   * run.setItalic(false);  // Remove italic
   * ```
   */
  setItalic(italic = true): this {
    const previousValue = this.formatting.italic;
    this.formatting.italic = italic;
    if (this.trackingContext?.isEnabled() && previousValue !== italic) {
      this.trackingContext.trackRunPropertyChange(this, 'italic', previousValue, italic);
    }
    return this;
  }

  /**
   * Sets bold formatting for complex scripts (RTL languages)
   * Per ECMA-376 Part 1 §17.3.2.3
   * @param bold - Whether text is bold for complex scripts
   */
  setComplexScriptBold(bold = true): this {
    const previousValue = this.formatting.complexScriptBold;
    this.formatting.complexScriptBold = bold;
    if (this.trackingContext?.isEnabled() && previousValue !== bold) {
      this.trackingContext.trackRunPropertyChange(this, 'complexScriptBold', previousValue, bold);
    }
    return this;
  }

  /**
   * Sets italic formatting for complex scripts (RTL languages)
   * Per ECMA-376 Part 1 §17.3.2.17
   * @param italic - Whether text is italic for complex scripts
   */
  setComplexScriptItalic(italic = true): this {
    const previousValue = this.formatting.complexScriptItalic;
    this.formatting.complexScriptItalic = italic;
    if (this.trackingContext?.isEnabled() && previousValue !== italic) {
      this.trackingContext.trackRunPropertyChange(this, 'complexScriptItalic', previousValue, italic);
    }
    return this;
  }

  /**
   * Sets character spacing (letter spacing)
   * Per ECMA-376 Part 1 §17.3.2.33
   * @param spacing - Spacing in twips (1/20th of a point). Positive values expand, negative values condense.
   */
  setCharacterSpacing(spacing: number): this {
    const previousValue = this.formatting.characterSpacing;
    this.formatting.characterSpacing = spacing;
    if (this.trackingContext?.isEnabled() && previousValue !== spacing) {
      this.trackingContext.trackRunPropertyChange(this, 'characterSpacing', previousValue, spacing);
    }
    return this;
  }

  /**
   * Sets horizontal text scaling
   * Per ECMA-376 Part 1 §17.3.2.43
   * @param scaling - Scaling percentage (e.g., 200 = 200% width, 50 = 50% width). Default is 100.
   */
  setScaling(scaling: number): this {
    const previousValue = this.formatting.scaling;
    this.formatting.scaling = scaling;
    if (this.trackingContext?.isEnabled() && previousValue !== scaling) {
      this.trackingContext.trackRunPropertyChange(this, 'scaling', previousValue, scaling);
    }
    return this;
  }

  /**
   * Sets vertical text position
   * Per ECMA-376 Part 1 §17.3.2.31
   * @param position - Position in half-points. Positive values raise text, negative values lower it.
   */
  setPosition(position: number): this {
    const previousValue = this.formatting.position;
    this.formatting.position = position;
    if (this.trackingContext?.isEnabled() && previousValue !== position) {
      this.trackingContext.trackRunPropertyChange(this, 'position', previousValue, position);
    }
    return this;
  }

  /**
   * Sets kerning threshold
   * Per ECMA-376 Part 1 §17.3.2.20
   * @param kerning - Font size in half-points at which kerning starts. 0 disables kerning.
   */
  setKerning(kerning: number): this {
    const previousValue = this.formatting.kerning;
    this.formatting.kerning = kerning;
    if (this.trackingContext?.isEnabled() && previousValue !== kerning) {
      this.trackingContext.trackRunPropertyChange(this, 'kerning', previousValue, kerning);
    }
    return this;
  }

  /**
   * Sets language
   * Per ECMA-376 Part 1 §17.3.2.20
   * @param language - Language code (e.g., 'en-US', 'fr-FR', 'es-ES')
   */
  setLanguage(language: string): this {
    const previousValue = this.formatting.language;
    this.formatting.language = language;
    if (this.trackingContext?.isEnabled() && previousValue !== language) {
      this.trackingContext.trackRunPropertyChange(this, 'language', previousValue, language);
    }
    return this;
  }

  /**
   * Sets underline formatting
   *
   * Applies various underline styles or removes underlining.
   *
   * @param underline - Underline style or boolean (default: true = 'single')
   *   - true or 'single': Single underline
   *   - 'double': Double underline
   *   - 'thick': Thick underline
   *   - 'dotted': Dotted underline
   *   - 'dash': Dashed underline
   *   - false: No underline
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setUnderline();          // Single underline
   * run.setUnderline('double');  // Double underline
   * run.setUnderline(false);     // Remove underline
   * ```
   */
  setUnderline(underline: RunFormatting["underline"] = true): this {
    const previousValue = this.formatting.underline;
    this.formatting.underline = underline;
    if (this.trackingContext?.isEnabled() && previousValue !== underline) {
      this.trackingContext.trackRunPropertyChange(this, 'underline', previousValue, underline);
    }
    return this;
  }

  /**
   * Sets underline color per ECMA-376 Part 1 §17.3.2.40
   * @param color - Color in hex format (without #)
   * @returns This run for method chaining
   */
  setUnderlineColor(color: string): this {
    this.formatting.underlineColor = normalizeColor(color);
    return this;
  }

  /**
   * Sets underline theme color per ECMA-376 Part 1 §17.3.2.40
   * @param themeColor - Theme color reference
   * @param themeTint - Optional tint (0-255)
   * @param themeShade - Optional shade (0-255)
   * @returns This run for method chaining
   */
  setUnderlineThemeColor(themeColor: ThemeColorValue, themeTint?: number, themeShade?: number): this {
    this.formatting.underlineThemeColor = themeColor;
    if (themeTint !== undefined) this.formatting.underlineThemeTint = themeTint;
    if (themeShade !== undefined) this.formatting.underlineThemeShade = themeShade;
    return this;
  }

  /**
   * Sets strikethrough formatting
   *
   * Adds or removes a line through the text.
   *
   * @param strike - If true, applies strikethrough; if false, removes it (default: true)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setStrike();       // Apply strikethrough
   * run.setStrike(false);  // Remove strikethrough
   * ```
   */
  setStrike(strike = true): this {
    const previousValue = this.formatting.strike;
    this.formatting.strike = strike;
    if (this.trackingContext?.isEnabled() && previousValue !== strike) {
      this.trackingContext.trackRunPropertyChange(this, 'strike', previousValue, strike);
    }
    return this;
  }

  /**
   * Sets subscript formatting
   *
   * Lowers the text below the baseline (e.g., H₂O).
   * Automatically removes superscript if set.
   *
   * @param subscript - If true, applies subscript; if false, removes it (default: true)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setText('H₂O');
   * run.setSubscript();  // Format as subscript
   * ```
   */
  setSubscript(subscript = true): this {
    const previousValue = this.formatting.subscript;
    this.formatting.subscript = subscript;
    if (subscript) {
      this.formatting.superscript = false;
    }
    if (this.trackingContext?.isEnabled() && previousValue !== subscript) {
      this.trackingContext.trackRunPropertyChange(this, 'subscript', previousValue, subscript);
    }
    return this;
  }

  /**
   * Sets superscript formatting
   *
   * Raises the text above the baseline (e.g., x²).
   * Automatically removes subscript if set.
   *
   * @param superscript - If true, applies superscript; if false, removes it (default: true)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setText('x²');
   * run.setSuperscript();  // Format as superscript
   * ```
   */
  setSuperscript(superscript = true): this {
    const previousValue = this.formatting.superscript;
    this.formatting.superscript = superscript;
    if (superscript) {
      this.formatting.subscript = false;
    }
    if (this.trackingContext?.isEnabled() && previousValue !== superscript) {
      this.trackingContext.trackRunPropertyChange(this, 'superscript', previousValue, superscript);
    }
    return this;
  }

  /**
   * Sets font family and optionally size
   *
   * Changes the font family (typeface) and optionally the font size.
   *
   * @param font - Font family name (e.g., 'Arial', 'Times New Roman', 'Verdana')
   * @param size - Optional font size in points
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setFont('Arial');           // Change font only
   * run.setFont('Verdana', 14);     // Change font and size
   * ```
   */
  setFont(font: string, size?: number): this {
    const previousFont = this.formatting.font;
    const previousSize = this.formatting.size;
    this.formatting.font = font;
    if (size !== undefined) {
      this.formatting.size = size;
    }
    if (this.trackingContext?.isEnabled()) {
      if (previousFont !== font) {
        this.trackingContext.trackRunPropertyChange(this, 'font', previousFont, font);
      }
      if (size !== undefined && previousSize !== size) {
        this.trackingContext.trackRunPropertyChange(this, 'size', previousSize, size);
      }
    }
    return this;
  }

  /**
   * Sets font size
   *
   * Changes the size of the text in points.
   *
   * @param size - Font size in points (e.g., 12 for 12pt text)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setSize(12);   // 12pt text
   * run.setSize(18);   // 18pt text
   * ```
   */
  setSize(size: number): this {
    const previousValue = this.formatting.size;
    this.formatting.size = size;
    if (this.trackingContext?.isEnabled() && previousValue !== size) {
      this.trackingContext.trackRunPropertyChange(this, 'size', previousValue, size);
    }
    return this;
  }

  /**
   * Sets font size for complex scripts (RTL languages like Arabic, Hebrew)
   *
   * Sets the font size used for complex script text (w:szCs element).
   * If not set, the regular size is used for both regular and complex script text.
   *
   * @param size - Font size in points for complex scripts
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setSize(12).setSizeCs(14);   // 12pt for regular, 14pt for complex scripts
   * ```
   */
  setSizeCs(size: number): this {
    const previousValue = this.formatting.sizeCs;
    this.formatting.sizeCs = size;
    if (this.trackingContext?.isEnabled() && previousValue !== size) {
      this.trackingContext.trackRunPropertyChange(this, 'sizeCs', previousValue, size);
    }
    return this;
  }

  /**
   * Sets text color
   *
   * Sets the foreground (text) color using hexadecimal format.
   * Color is automatically normalized to uppercase without the # prefix.
   *
   * @param color - Color in hex format (with or without # prefix)
   * @returns This run instance for method chaining
   *
   * @throws Error if color format is invalid (not 6 hex characters)
   *
   * @example
   * ```typescript
   * run.setColor('FF0000');   // Red text
   * run.setColor('#0000FF');  // Blue text (# is removed)
   * run.setColor('00FF00');   // Green text
   * ```
   */
  setColor(color: string): this {
    const previousValue = this.formatting.color;
    const normalizedColor = normalizeColor(color);
    this.formatting.color = normalizedColor;
    if (this.trackingContext?.isEnabled() && previousValue !== normalizedColor) {
      this.trackingContext.trackRunPropertyChange(this, 'color', previousValue, normalizedColor);
    }
    return this;
  }

  /**
   * Sets theme color reference for text
   *
   * Uses a color from the document's theme instead of a fixed hex value.
   * Theme colors automatically update when the document theme changes.
   *
   * @param themeColor - Theme color reference (e.g., 'accent1', 'dark1', 'hyperlink')
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setThemeColor('accent1');    // Use theme accent color 1
   * run.setThemeColor('hyperlink');  // Use theme hyperlink color
   * ```
   */
  setThemeColor(themeColor: ThemeColorValue): this {
    const previousValue = this.formatting.themeColor;
    this.formatting.themeColor = themeColor;
    if (this.trackingContext?.isEnabled() && previousValue !== themeColor) {
      this.trackingContext.trackRunPropertyChange(this, 'themeColor', previousValue, themeColor);
    }
    return this;
  }

  /**
   * Sets theme color tint for lighter variations
   *
   * Applied to the theme color to create a lighter shade.
   * The tint value is a percentage where 0 = no change and 255 = white.
   *
   * @param themeTint - Tint value (0-255, toward white)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setThemeColor('accent1').setThemeTint(128);  // 50% tint toward white
   * ```
   */
  setThemeTint(themeTint: number): this {
    const previousValue = this.formatting.themeTint;
    this.formatting.themeTint = themeTint;
    if (this.trackingContext?.isEnabled() && previousValue !== themeTint) {
      this.trackingContext.trackRunPropertyChange(this, 'themeTint', previousValue, themeTint);
    }
    return this;
  }

  /**
   * Sets theme color shade for darker variations
   *
   * Applied to the theme color to create a darker shade.
   * The shade value is a percentage where 0 = no change and 255 = black.
   *
   * @param themeShade - Shade value (0-255, toward black)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setThemeColor('accent1').setThemeShade(128);  // 50% shade toward black
   * ```
   */
  setThemeShade(themeShade: number): this {
    const previousValue = this.formatting.themeShade;
    this.formatting.themeShade = themeShade;
    if (this.trackingContext?.isEnabled() && previousValue !== themeShade) {
      this.trackingContext.trackRunPropertyChange(this, 'themeShade', previousValue, themeShade);
    }
    return this;
  }

  /**
   * Sets highlight (background) color
   *
   * Applies a background highlight color to the text, similar to using
   * a highlighter marker.
   *
   * @param highlight - Highlight color name
   *   - Standard colors: 'yellow', 'green', 'cyan', 'magenta', 'blue', 'red'
   *   - Dark variants: 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow'
   *   - Grayscale: 'darkGray', 'lightGray', 'black', 'white'
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setHighlight('yellow');     // Yellow highlight
   * run.setHighlight('darkGreen');  // Dark green highlight
   * ```
   */
  setHighlight(highlight: RunFormatting["highlight"]): this {
    const previousValue = this.formatting.highlight;
    this.formatting.highlight = highlight;
    if (this.trackingContext?.isEnabled() && previousValue !== highlight) {
      this.trackingContext.trackRunPropertyChange(this, 'highlight', previousValue, highlight);
    }
    return this;
  }

  /**
   * Sets small caps formatting
   *
   * Formats lowercase letters as smaller versions of capital letters.
   *
   * @param smallCaps - If true, applies small caps; if false, removes it (default: true)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setText('Small Caps Text');
   * run.setSmallCaps();  // SMALL CAPS TEXT
   * ```
   */
  setSmallCaps(smallCaps = true): this {
    const previousValue = this.formatting.smallCaps;
    this.formatting.smallCaps = smallCaps;
    if (this.trackingContext?.isEnabled() && previousValue !== smallCaps) {
      this.trackingContext.trackRunPropertyChange(this, 'smallCaps', previousValue, smallCaps);
    }
    return this;
  }

  /**
   * Sets all caps formatting
   *
   * Displays all text in capital letters regardless of original case.
   *
   * @param allCaps - If true, applies all caps; if false, removes it (default: true)
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.setText('All Caps Text');
   * run.setAllCaps();  // ALL CAPS TEXT
   * ```
   */
  setAllCaps(allCaps = true): this {
    const previousValue = this.formatting.allCaps;
    this.formatting.allCaps = allCaps;
    if (this.trackingContext?.isEnabled() && previousValue !== allCaps) {
      this.trackingContext.trackRunPropertyChange(this, 'allCaps', previousValue, allCaps);
    }
    return this;
  }

  /**
   * Sets outline text effect
   * @param outline - Whether to apply outline effect (default: true)
   * @returns This run for method chaining
   */
  setOutline(outline = true): this {
    const previousValue = this.formatting.outline;
    this.formatting.outline = outline;
    if (this.trackingContext?.isEnabled() && previousValue !== outline) {
      this.trackingContext.trackRunPropertyChange(this, 'outline', previousValue, outline);
    }
    return this;
  }

  /**
   * Sets shadow text effect
   * @param shadow - Whether to apply shadow effect (default: true)
   * @returns This run for method chaining
   */
  setShadow(shadow = true): this {
    const previousValue = this.formatting.shadow;
    this.formatting.shadow = shadow;
    if (this.trackingContext?.isEnabled() && previousValue !== shadow) {
      this.trackingContext.trackRunPropertyChange(this, 'shadow', previousValue, shadow);
    }
    return this;
  }

  /**
   * Sets emboss text effect
   * @param emboss - Whether to apply emboss effect (default: true)
   * @returns This run for method chaining
   */
  setEmboss(emboss = true): this {
    const previousValue = this.formatting.emboss;
    this.formatting.emboss = emboss;
    if (this.trackingContext?.isEnabled() && previousValue !== emboss) {
      this.trackingContext.trackRunPropertyChange(this, 'emboss', previousValue, emboss);
    }
    return this;
  }

  /**
   * Sets imprint/engrave text effect
   * @param imprint - Whether to apply imprint effect (default: true)
   * @returns This run for method chaining
   */
  setImprint(imprint = true): this {
    const previousValue = this.formatting.imprint;
    this.formatting.imprint = imprint;
    if (this.trackingContext?.isEnabled() && previousValue !== imprint) {
      this.trackingContext.trackRunPropertyChange(this, 'imprint', previousValue, imprint);
    }
    return this;
  }

  /**
   * Sets right-to-left text direction
   * @param rtl - Whether text is RTL (default: true)
   * @returns This run for method chaining
   */
  setRTL(rtl = true): this {
    const previousValue = this.formatting.rtl;
    this.formatting.rtl = rtl;
    if (this.trackingContext?.isEnabled() && previousValue !== rtl) {
      this.trackingContext.trackRunPropertyChange(this, 'rtl', previousValue, rtl);
    }
    return this;
  }

  /**
   * Sets hidden/vanish text
   * @param vanish - Whether text is hidden (default: true)
   * @returns This run for method chaining
   */
  setVanish(vanish = true): this {
    const previousValue = this.formatting.vanish;
    this.formatting.vanish = vanish;
    if (this.trackingContext?.isEnabled() && previousValue !== vanish) {
      this.trackingContext.trackRunPropertyChange(this, 'vanish', previousValue, vanish);
    }
    return this;
  }

  /**
   * Sets no proofing (skip spell/grammar check)
   * @param noProof - Whether to skip proofing (default: true)
   * @returns This run for method chaining
   */
  setNoProof(noProof = true): this {
    const previousValue = this.formatting.noProof;
    this.formatting.noProof = noProof;
    if (this.trackingContext?.isEnabled() && previousValue !== noProof) {
      this.trackingContext.trackRunPropertyChange(this, 'noProof', previousValue, noProof);
    }
    return this;
  }

  /**
   * Sets snap to grid alignment
   * @param snapToGrid - Whether to snap to grid (default: true)
   * @returns This run for method chaining
   */
  setSnapToGrid(snapToGrid = true): this {
    const previousValue = this.formatting.snapToGrid;
    this.formatting.snapToGrid = snapToGrid;
    if (this.trackingContext?.isEnabled() && previousValue !== snapToGrid) {
      this.trackingContext.trackRunPropertyChange(this, 'snapToGrid', previousValue, snapToGrid);
    }
    return this;
  }

  /**
   * Sets special vanish (hidden for specific scenarios like TOC)
   * @param specVanish - Whether to apply special vanish (default: true)
   * @returns This run for method chaining
   */
  setSpecVanish(specVanish = true): this {
    const previousValue = this.formatting.specVanish;
    this.formatting.specVanish = specVanish;
    if (this.trackingContext?.isEnabled() && previousValue !== specVanish) {
      this.trackingContext.trackRunPropertyChange(this, 'specVanish', previousValue, specVanish);
    }
    return this;
  }

  /**
   * Sets complex script formatting flag (w:cs)
   * @param complexScript - Whether complex script formatting applies (default: true)
   * @returns This run for method chaining
   */
  setComplexScript(complexScript = true): this {
    const previousValue = this.formatting.complexScript;
    this.formatting.complexScript = complexScript;
    if (this.trackingContext?.isEnabled() && previousValue !== complexScript) {
      this.trackingContext.trackRunPropertyChange(this, 'complexScript', previousValue, complexScript);
    }
    return this;
  }

  /**
   * Sets web hidden flag (w:webHidden) - hide text in web layout view
   * @param webHidden - Whether text is hidden in web view (default: true)
   * @returns This run for method chaining
   */
  setWebHidden(webHidden = true): this {
    const previousValue = this.formatting.webHidden;
    this.formatting.webHidden = webHidden;
    if (this.trackingContext?.isEnabled() && previousValue !== webHidden) {
      this.trackingContext.trackRunPropertyChange(this, 'webHidden', previousValue, webHidden);
    }
    return this;
  }

  /**
   * Sets the High ANSI font (w:rFonts w:hAnsi)
   * @param font - Font name for high ANSI characters
   * @returns This run for method chaining
   */
  setFontHAnsi(font: string): this {
    const previousValue = this.formatting.fontHAnsi;
    this.formatting.fontHAnsi = font;
    if (this.trackingContext?.isEnabled() && previousValue !== font) {
      this.trackingContext.trackRunPropertyChange(this, 'fontHAnsi', previousValue, font);
    }
    return this;
  }

  /**
   * Sets the East Asian font (w:rFonts w:eastAsia)
   * @param font - Font name for East Asian characters
   * @returns This run for method chaining
   */
  setFontEastAsia(font: string): this {
    const previousValue = this.formatting.fontEastAsia;
    this.formatting.fontEastAsia = font;
    if (this.trackingContext?.isEnabled() && previousValue !== font) {
      this.trackingContext.trackRunPropertyChange(this, 'fontEastAsia', previousValue, font);
    }
    return this;
  }

  /**
   * Sets the complex script font (w:rFonts w:cs)
   * @param font - Font name for complex script (RTL) characters
   * @returns This run for method chaining
   */
  setFontCs(font: string): this {
    const previousValue = this.formatting.fontCs;
    this.formatting.fontCs = font;
    if (this.trackingContext?.isEnabled() && previousValue !== font) {
      this.trackingContext.trackRunPropertyChange(this, 'fontCs', previousValue, font);
    }
    return this;
  }

  /**
   * Sets the font hint (w:rFonts w:hint)
   * @param hint - Font selection hint: 'default', 'eastAsia', or 'cs'
   * @returns This run for method chaining
   */
  setFontHint(hint: string): this {
    const previousValue = this.formatting.fontHint;
    this.formatting.fontHint = hint;
    if (this.trackingContext?.isEnabled() && previousValue !== hint) {
      this.trackingContext.trackRunPropertyChange(this, 'fontHint', previousValue, hint);
    }
    return this;
  }

  /**
   * Sets the ASCII theme font reference (w:rFonts w:asciiTheme)
   * @param theme - Theme font name (e.g., 'majorHAnsi', 'minorHAnsi')
   * @returns This run for method chaining
   */
  setFontAsciiTheme(theme: string): this {
    this.formatting.fontAsciiTheme = theme;
    return this;
  }

  /**
   * Sets the High ANSI theme font reference (w:rFonts w:hAnsiTheme)
   * @param theme - Theme font name
   * @returns This run for method chaining
   */
  setFontHAnsiTheme(theme: string): this {
    this.formatting.fontHAnsiTheme = theme;
    return this;
  }

  /**
   * Sets the East Asian theme font reference (w:rFonts w:eastAsiaTheme)
   * @param theme - Theme font name
   * @returns This run for method chaining
   */
  setFontEastAsiaTheme(theme: string): this {
    this.formatting.fontEastAsiaTheme = theme;
    return this;
  }

  /**
   * Sets the complex script theme font reference (w:rFonts w:cstheme)
   * @param theme - Theme font name
   * @returns This run for method chaining
   */
  setFontCsTheme(theme: string): this {
    this.formatting.fontCsTheme = theme;
    return this;
  }

  /**
   * Adds a raw w14: property XML string for passthrough (Word 2010+ text effects).
   *
   * **Warning:** The input XML is embedded directly in the output without sanitization.
   * Only pass trusted w14 namespace elements (e.g., w14:textOutline, w14:ligatures).
   * Do not pass user-controlled or untrusted input to this method.
   *
   * @param rawXml - The raw XML string for a w14: element (e.g., '<w14:textOutline ...>...</w14:textOutline>')
   * @returns This run for method chaining
   */
  addRawW14Property(rawXml: string): this {
    if (!this.formatting.rawW14Properties) {
      this.formatting.rawW14Properties = [];
    }
    this.formatting.rawW14Properties.push(rawXml);
    return this;
  }

  /**
   * Sets text effect/animation
   * @param effect - Effect type (e.g., 'shimmer', 'sparkleText')
   * @returns This run for method chaining
   */
  setEffect(
    effect:
      | "none"
      | "lights"
      | "blinkBackground"
      | "sparkleText"
      | "marchingBlackAnts"
      | "marchingRedAnts"
      | "shimmer"
      | "antsBlack"
      | "antsRed"
  ): this {
    const previousValue = this.formatting.effect;
    this.formatting.effect = effect;
    if (this.trackingContext?.isEnabled() && previousValue !== effect) {
      this.trackingContext.trackRunPropertyChange(this, 'effect', previousValue, effect);
    }
    return this;
  }

  /**
   * Sets fit text to width
   * @param width - Width in twips (1/20th of a point)
   * @returns This run for method chaining
   */
  setFitText(width: number): this {
    const previousValue = this.formatting.fitText;
    this.formatting.fitText = width;
    if (this.trackingContext?.isEnabled() && previousValue !== width) {
      this.trackingContext.trackRunPropertyChange(this, 'fitText', previousValue, width);
    }
    return this;
  }

  /**
   * Sets East Asian typography layout
   * @param layout - East Asian layout options
   * @returns This run for method chaining
   */
  setEastAsianLayout(layout: EastAsianLayout): this {
    const previousValue = this.formatting.eastAsianLayout;
    this.formatting.eastAsianLayout = layout;
    if (this.trackingContext?.isEnabled() && previousValue !== layout) {
      this.trackingContext.trackRunPropertyChange(this, 'eastAsianLayout', previousValue, layout);
    }
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
    logSerialization(`Serializing run: "${text}"`, {
      rtl: this.formatting.rtl || false,
    });
    if (this.formatting.rtl) {
      logTextDirection(`Run with RTL being serialized: "${text}"`);
    }

    // Build the run element
    const runChildren: XMLElement[] = [];

    // Add run properties using the static helper
    const rPr = Run.generateRunPropertiesXML(this.formatting);
    if (rPr || this.propertyChangeRevision) {
      // If we have a property change revision, we need to add w:rPrChange to the w:rPr element
      // Per ECMA-376, w:rPrChange must come after all other run properties
      const rPrElement = rPr || { name: 'w:rPr', attributes: {}, children: [] };

      if (this.propertyChangeRevision) {
        // Reuse generateRunPropertiesXML for full ECMA-376 property coverage
        // This ensures all 30+ run properties are serialized in correct schema order,
        // rather than the subset previously handled by manual serialization.
        const prevRPr = this.propertyChangeRevision.previousProperties
          ? Run.generateRunPropertiesXML(this.propertyChangeRevision.previousProperties as RunFormatting)
          : null;
        const rPrChangeChildren: XMLElement[] = prevRPr ? [prevRPr] : [];

        // Create w:rPrChange element with attributes
        const rPrChange: XMLElement = {
          name: 'w:rPrChange',
          attributes: {
            'w:id': this.propertyChangeRevision.id.toString(),
            'w:author': this.propertyChangeRevision.author,
            'w:date': formatDateForXml(this.propertyChangeRevision.date),
          },
          children: rPrChangeChildren,
        };

        // Add rPrChange to rPr children (must come last per ECMA-376)
        if (rPrElement.children) {
          rPrElement.children.push(rPrChange);
        } else {
          rPrElement.children = [rPrChange];
        }
      }

      runChildren.push(rPrElement);
    }

    // Add run content elements (text, tabs, breaks, etc.) in order
    for (const contentElement of this.content) {
      switch (contentElement.type) {
        case "text":
          // Always generate <w:t> element, even for empty strings
          // This ensures proper Word compatibility and round-trip preservation
          runChildren.push(
            XMLBuilder.w(
              "t",
              {
                "xml:space": "preserve",
              },
              [contentElement.value || ""]
            )
          );
          break;

        case "tab":
          runChildren.push(XMLBuilder.wSelf("tab"));
          break;

        case "break":
          {
            const attrs: Record<string, string> = {};
            if (contentElement.breakType) {
              attrs["w:type"] = contentElement.breakType;
            }
            runChildren.push(
              XMLBuilder.wSelf(
                "br",
                Object.keys(attrs).length > 0 ? attrs : undefined
              )
            );
          }
          break;

        case "carriageReturn":
          runChildren.push(XMLBuilder.wSelf("cr"));
          break;

        case "softHyphen":
          runChildren.push(XMLBuilder.wSelf("softHyphen"));
          break;

        case "noBreakHyphen":
          runChildren.push(XMLBuilder.wSelf("noBreakHyphen"));
          break;

        case "instructionText":
          runChildren.push(
            XMLBuilder.w("instrText", { "xml:space": "preserve" }, [
              contentElement.value || "",
            ])
          );
          break;

        case "fieldChar": {
          if (!contentElement.fieldCharType) {
            break;
          }
          const fldCharAttrs: Record<string, string> = {
            "w:fldCharType": contentElement.fieldCharType,
          };
          if (contentElement.fieldCharDirty !== undefined) {
            fldCharAttrs["w:dirty"] = contentElement.fieldCharDirty ? "1" : "0";
          }
          if (contentElement.fieldCharLocked !== undefined) {
            fldCharAttrs["w:fldLock"] = contentElement.fieldCharLocked ? "1" : "0";
          }
          // Generate ffData for begin field chars with form field data
          if (contentElement.formFieldData && contentElement.fieldCharType === "begin") {
            const ffDataChildren: (string | XMLElement)[] = [];
            const ffd = contentElement.formFieldData;
            if (ffd.name !== undefined) {
              ffDataChildren.push(XMLBuilder.wSelf("name", { "w:val": ffd.name }));
            }
            if (ffd.enabled !== undefined) {
              if (ffd.enabled) {
                ffDataChildren.push(XMLBuilder.wSelf("enabled"));
              } else {
                ffDataChildren.push(XMLBuilder.wSelf("enabled", { "w:val": "0" }));
              }
            }
            if (ffd.calcOnExit !== undefined) {
              ffDataChildren.push(XMLBuilder.wSelf("calcOnExit", { "w:val": ffd.calcOnExit ? "1" : "0" }));
            }
            if (ffd.helpText) {
              ffDataChildren.push(XMLBuilder.wSelf("helpText", { "w:type": "text", "w:val": ffd.helpText }));
            }
            if (ffd.statusText) {
              ffDataChildren.push(XMLBuilder.wSelf("statusText", { "w:type": "text", "w:val": ffd.statusText }));
            }
            if (ffd.entryMacro) {
              ffDataChildren.push(XMLBuilder.wSelf("entryMacro", { "w:val": ffd.entryMacro }));
            }
            if (ffd.exitMacro) {
              ffDataChildren.push(XMLBuilder.wSelf("exitMacro", { "w:val": ffd.exitMacro }));
            }
            if (ffd.fieldType) {
              switch (ffd.fieldType.type) {
                case 'textInput': {
                  const tiChildren: (string | XMLElement)[] = [];
                  if (ffd.fieldType.inputType) {
                    tiChildren.push(XMLBuilder.wSelf("type", { "w:val": ffd.fieldType.inputType }));
                  }
                  if (ffd.fieldType.defaultValue !== undefined) {
                    tiChildren.push(XMLBuilder.wSelf("default", { "w:val": ffd.fieldType.defaultValue }));
                  }
                  if (ffd.fieldType.maxLength !== undefined) {
                    tiChildren.push(XMLBuilder.wSelf("maxLength", { "w:val": ffd.fieldType.maxLength.toString() }));
                  }
                  if (ffd.fieldType.format) {
                    tiChildren.push(XMLBuilder.wSelf("format", { "w:val": ffd.fieldType.format }));
                  }
                  ffDataChildren.push(XMLBuilder.w("textInput", {}, tiChildren));
                  break;
                }
                case 'checkBox': {
                  const cbChildren: (string | XMLElement)[] = [];
                  if (ffd.fieldType.size !== undefined && ffd.fieldType.size !== 'auto') {
                    cbChildren.push(XMLBuilder.wSelf("size", { "w:val": ffd.fieldType.size.toString() }));
                  } else {
                    cbChildren.push(XMLBuilder.wSelf("sizeAuto"));
                  }
                  if (ffd.fieldType.defaultChecked !== undefined) {
                    cbChildren.push(XMLBuilder.wSelf("default", { "w:val": ffd.fieldType.defaultChecked ? "1" : "0" }));
                  }
                  if (ffd.fieldType.checked !== undefined) {
                    cbChildren.push(XMLBuilder.wSelf("checked", { "w:val": ffd.fieldType.checked ? "1" : "0" }));
                  }
                  ffDataChildren.push(XMLBuilder.w("checkBox", {}, cbChildren));
                  break;
                }
                case 'dropDownList': {
                  const ddChildren: (string | XMLElement)[] = [];
                  if (ffd.fieldType.result !== undefined) {
                    ddChildren.push(XMLBuilder.wSelf("result", { "w:val": ffd.fieldType.result.toString() }));
                  }
                  if (ffd.fieldType.defaultResult !== undefined) {
                    ddChildren.push(XMLBuilder.wSelf("default", { "w:val": ffd.fieldType.defaultResult.toString() }));
                  }
                  if (ffd.fieldType.listEntries) {
                    for (const entry of ffd.fieldType.listEntries) {
                      ddChildren.push(XMLBuilder.wSelf("listEntry", { "w:val": entry }));
                    }
                  }
                  ffDataChildren.push(XMLBuilder.w("ddList", {}, ddChildren));
                  break;
                }
              }
            }
            runChildren.push(
              XMLBuilder.w("fldChar", fldCharAttrs, [XMLBuilder.w("ffData", {}, ffDataChildren)])
            );
          } else {
            runChildren.push(
              XMLBuilder.wSelf(
                "fldChar",
                Object.keys(fldCharAttrs).length > 0 ? fldCharAttrs : undefined
              )
            );
          }
          break;
        }

        // Simple marker elements (self-closing, no attributes)
        case "lastRenderedPageBreak":
          runChildren.push(XMLBuilder.wSelf("lastRenderedPageBreak"));
          break;

        case "separator":
          runChildren.push(XMLBuilder.wSelf("separator"));
          break;

        case "continuationSeparator":
          runChildren.push(XMLBuilder.wSelf("continuationSeparator"));
          break;

        case "pageNumber":
          runChildren.push(XMLBuilder.wSelf("pgNum"));
          break;

        case "annotationRef":
          runChildren.push(XMLBuilder.wSelf("annotationRef"));
          break;

        case "dayShort":
          runChildren.push(XMLBuilder.wSelf("dayShort"));
          break;

        case "dayLong":
          runChildren.push(XMLBuilder.wSelf("dayLong"));
          break;

        case "monthShort":
          runChildren.push(XMLBuilder.wSelf("monthShort"));
          break;

        case "monthLong":
          runChildren.push(XMLBuilder.wSelf("monthLong"));
          break;

        case "yearShort":
          runChildren.push(XMLBuilder.wSelf("yearShort"));
          break;

        case "yearLong":
          runChildren.push(XMLBuilder.wSelf("yearLong"));
          break;

        // Symbol character (w:sym) per ECMA-376 Part 1 §17.3.3.30
        case "symbol": {
          const symAttrs: Record<string, string> = {};
          if (contentElement.symbolFont) symAttrs["w:font"] = contentElement.symbolFont;
          if (contentElement.symbolChar) symAttrs["w:char"] = contentElement.symbolChar;
          runChildren.push(XMLBuilder.wSelf("sym", symAttrs));
          break;
        }

        // Absolute position tab (w:ptab) per ECMA-376 Part 1 §17.3.3.23
        case "positionTab": {
          const ptabAttrs: Record<string, string> = {};
          if (contentElement.ptabAlignment) ptabAttrs["w:alignment"] = contentElement.ptabAlignment;
          if (contentElement.ptabRelativeTo) ptabAttrs["w:relativeTo"] = contentElement.ptabRelativeTo;
          if (contentElement.ptabLeader) ptabAttrs["w:leader"] = contentElement.ptabLeader;
          runChildren.push(XMLBuilder.wSelf("ptab", ptabAttrs));
          break;
        }

        // Footnote reference (w:footnoteReference) per ECMA-376 Part 1 §17.11.13
        case "footnoteReference": {
          const fnAttrs: Record<string, string | number> = {};
          if (contentElement.footnoteId !== undefined) fnAttrs["w:id"] = contentElement.footnoteId;
          runChildren.push(XMLBuilder.wSelf("footnoteReference", fnAttrs));
          break;
        }

        // Endnote reference (w:endnoteReference) per ECMA-376 Part 1 §17.11.2
        case "endnoteReference": {
          const enAttrs: Record<string, string | number> = {};
          if (contentElement.endnoteId !== undefined) enAttrs["w:id"] = contentElement.endnoteId;
          runChildren.push(XMLBuilder.wSelf("endnoteReference", enAttrs));
          break;
        }

        // Embedded OLE object (w:object) - preserved as raw XML
        case "embeddedObject":
          if (contentElement.rawXml) {
            runChildren.push({
              name: "__rawXml",
              rawXml: contentElement.rawXml,
            });
          }
          break;

        case "vml":
          // VML graphics (w:pict) - include raw XML without modification
          // This preserves legacy Word graphics like icons and symbols
          if (contentElement.rawXml) {
            // Use special __rawXml element name for passthrough (no wrapper element)
            runChildren.push({
              name: "__rawXml",
              rawXml: contentElement.rawXml,
            });
          }
          break;
      }
    }

    return XMLBuilder.w("r", undefined, runChildren);
  }

  /**
   * Checks if the run contains non-empty text
   *
   * @returns True if the run has text content with length > 0
   *
   * @example
   * ```typescript
   * const run1 = new Run('Hello');
   * const run2 = new Run('');
   * console.log(run1.hasText()); // true
   * console.log(run2.hasText()); // false
   * ```
   */
  hasText(): boolean {
    const text = this.getText();
    return text.length > 0;
  }

  /**
   * Checks if the run has any formatting properties
   *
   * @returns True if any formatting properties (bold, italic, font, etc.) are set
   *
   * @example
   * ```typescript
   * const run1 = new Run('Text', { bold: true });
   * const run2 = new Run('Text');
   * console.log(run1.hasFormatting()); // true
   * console.log(run2.hasFormatting()); // false
   * ```
   */
  hasFormatting(): boolean {
    return Object.keys(this.formatting).length > 0;
  }

  /**
   * Checks if the run is valid
   *
   * A run is considered valid if it has either text content or formatting.
   * An empty run with no formatting is invalid.
   *
   * @returns True if the run has text content or formatting properties
   *
   * @example
   * ```typescript
   * const run1 = new Run('Text');
   * const run2 = new Run('', { bold: true });
   * const run3 = new Run('');
   * console.log(run1.isValid()); // true (has text)
   * console.log(run2.isValid()); // true (has formatting)
   * console.log(run3.isValid()); // false (empty and unformatted)
   * ```
   */
  isValid(): boolean {
    return this.hasText() || this.hasFormatting();
  }

  /**
   * Gets the run content elements
   *
   * Returns the internal content structure including text elements,
   * tabs, breaks, and other special characters.
   *
   * @returns Array of RunContent elements
   *
   * @example
   * ```typescript
   * const content = run.getContent();
   * for (const element of content) {
   *   console.log(`Type: ${element.type}, Value: ${element.value || 'N/A'}`);
   * }
   * ```
   */
  getContent(): RunContent[] {
    return [...this.content];
  }

  /**
   * Adds a tab character to the run
   *
   * Inserts a tab character, commonly used in TOC entries to separate
   * heading text from page numbers, or for general text alignment.
   *
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * const run = new Run('Heading');
   * run.addTab();
   * run.appendText('Page 5'); // "Heading\tPage 5"
   * ```
   */
  addTab(): this {
    this.content.push({ type: "tab" });
    return this;
  }

  /**
   * Adds a line, page, or column break to the run
   *
   * Inserts a break element for line breaks, page breaks, or column breaks.
   *
   * @param breakType - Type of break (default: line break if not specified)
   *   - undefined or 'textWrapping': Line break (like pressing Enter within a paragraph)
   *   - 'page': Page break
   *   - 'column': Column break
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * run.addBreak();          // Line break
   * run.addBreak('page');    // Page break
   * run.addBreak('column');  // Column break
   * ```
   */
  addBreak(breakType?: BreakType): this {
    this.content.push({ type: "break", breakType });
    return this;
  }

  /**
   * Appends text to the run
   *
   * Adds additional text as a new content element without replacing
   * existing content. Useful for building text incrementally.
   *
   * @param text - Text to append
   * @returns This run instance for method chaining
   *
   * @example
   * ```typescript
   * const run = new Run('Hello');
   * run.appendText(' ');
   * run.appendText('World');  // "Hello World"
   * ```
   */
  appendText(text: string): this {
    if (text) {
      this.content.push({ type: "text", value: text });
    }
    return this;
  }

  /**
   * Adds a carriage return to the run
   * @returns This run for method chaining
   */
  addCarriageReturn(): this {
    this.content.push({ type: "carriageReturn" });
    return this;
  }

  /**
   * Creates a Run from an array of content elements
   * Factory method for advanced use cases (used by DocumentParser)
   * @param content - Array of run content elements
   * @param formatting - Run formatting options
   * @returns New Run instance
   */
  static createFromContent(
    content: RunContent[],
    formatting: RunFormatting = {}
  ): Run {
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
  static generateRunPropertiesXML(
    formatting: RunFormatting
  ): XMLElement | null {
    const rPrChildren: XMLElement[] = [];

    // ECMA-376 Part 1 §17.3.2.28 CT_RPr element ordering
    // Elements MUST appear in this exact sequence for Word compatibility

    // 1. w:rStyle — Character style reference
    if (formatting.characterStyle) {
      rPrChildren.push(
        XMLBuilder.wSelf("rStyle", {
          "w:val": formatting.characterStyle,
        })
      );
    }

    // 2. w:rFonts — Font family
    if (formatting.font || formatting.fontHAnsi || formatting.fontEastAsia || formatting.fontCs || formatting.fontHint ||
        formatting.fontAsciiTheme || formatting.fontHAnsiTheme || formatting.fontEastAsiaTheme || formatting.fontCsTheme) {
      const rFontsAttrs: Record<string, string> = {};
      if (formatting.font) rFontsAttrs["w:ascii"] = formatting.font;
      rFontsAttrs["w:hAnsi"] = formatting.fontHAnsi || formatting.font || "";
      if (formatting.fontEastAsia) rFontsAttrs["w:eastAsia"] = formatting.fontEastAsia;
      rFontsAttrs["w:cs"] = formatting.fontCs || formatting.font || "";
      if (formatting.fontHint) rFontsAttrs["w:hint"] = formatting.fontHint;
      // Theme font references per ECMA-376 Part 1 §17.3.2.26
      if (formatting.fontAsciiTheme) rFontsAttrs["w:asciiTheme"] = formatting.fontAsciiTheme;
      if (formatting.fontHAnsiTheme) rFontsAttrs["w:hAnsiTheme"] = formatting.fontHAnsiTheme;
      if (formatting.fontEastAsiaTheme) rFontsAttrs["w:eastAsiaTheme"] = formatting.fontEastAsiaTheme;
      if (formatting.fontCsTheme) rFontsAttrs["w:cstheme"] = formatting.fontCsTheme;

      // Remove empty string values (only include attributes that have actual values)
      for (const key of Object.keys(rFontsAttrs)) {
        if (rFontsAttrs[key] === "") delete rFontsAttrs[key];
      }

      if (Object.keys(rFontsAttrs).length > 0) {
        rPrChildren.push(XMLBuilder.wSelf("rFonts", rFontsAttrs));
      }
    }

    // 3. w:b — Bold
    if (formatting.bold) {
      rPrChildren.push(XMLBuilder.wSelf("b", { "w:val": "1" }));
    }

    // 4. w:bCs — Bold complex script
    if (formatting.complexScriptBold) {
      rPrChildren.push(XMLBuilder.wSelf("bCs", { "w:val": "1" }));
    }

    // 5. w:i — Italic
    if (formatting.italic) {
      rPrChildren.push(XMLBuilder.wSelf("i", { "w:val": "1" }));
    }

    // 6. w:iCs — Italic complex script
    if (formatting.complexScriptItalic) {
      rPrChildren.push(XMLBuilder.wSelf("iCs", { "w:val": "1" }));
    }

    // 7. w:caps — All caps
    if (formatting.allCaps) {
      rPrChildren.push(XMLBuilder.wSelf("caps", { "w:val": "1" }));
    }

    // 8. w:smallCaps — Small caps
    if (formatting.smallCaps) {
      rPrChildren.push(XMLBuilder.wSelf("smallCaps", { "w:val": "1" }));
    }

    // 9. w:strike — Single strikethrough
    if (formatting.strike) {
      rPrChildren.push(XMLBuilder.wSelf("strike", { "w:val": "1" }));
    }

    // 10. w:dstrike — Double strikethrough
    if (formatting.dstrike) {
      rPrChildren.push(XMLBuilder.wSelf("dstrike", { "w:val": "1" }));
    }

    // 11. w:outline — Outline text effect
    if (formatting.outline) {
      rPrChildren.push(XMLBuilder.wSelf("outline", { "w:val": "1" }));
    }

    // 12. w:shadow — Shadow text effect
    if (formatting.shadow) {
      rPrChildren.push(XMLBuilder.wSelf("shadow", { "w:val": "1" }));
    }

    // 13. w:emboss — Emboss text effect
    if (formatting.emboss) {
      rPrChildren.push(XMLBuilder.wSelf("emboss", { "w:val": "1" }));
    }

    // 14. w:imprint — Imprint/engrave text effect
    if (formatting.imprint) {
      rPrChildren.push(XMLBuilder.wSelf("imprint", { "w:val": "1" }));
    }

    // 15. w:noProof — No proofing
    if (formatting.noProof) {
      rPrChildren.push(XMLBuilder.wSelf("noProof", { "w:val": "1" }));
    }

    // 16. w:snapToGrid — Snap to grid
    if (formatting.snapToGrid) {
      rPrChildren.push(XMLBuilder.wSelf("snapToGrid", { "w:val": "1" }));
    }

    // 17. w:vanish — Hidden text
    if (formatting.vanish) {
      rPrChildren.push(XMLBuilder.wSelf("vanish", { "w:val": "1" }));
    }

    // 18. w:webHidden — Web hidden
    if (formatting.webHidden) {
      rPrChildren.push(XMLBuilder.wSelf("webHidden", { "w:val": "1" }));
    }

    // 19. w:color — Text color
    // Supports both hex colors and theme color references
    // w:val is REQUIRED per ECMA-376 — defaults to "000000" when only themeColor is set
    if (formatting.color || formatting.themeColor) {
      const colorAttrs: Record<string, string> = {};

      colorAttrs["w:val"] = formatting.color || "000000";
      if (formatting.themeColor) {
        colorAttrs["w:themeColor"] = formatting.themeColor;
      }
      if (formatting.themeTint !== undefined) {
        colorAttrs["w:themeTint"] = formatting.themeTint
          .toString(16)
          .toUpperCase()
          .padStart(2, "0");
      }
      if (formatting.themeShade !== undefined) {
        colorAttrs["w:themeShade"] = formatting.themeShade
          .toString(16)
          .toUpperCase()
          .padStart(2, "0");
      }

      rPrChildren.push(XMLBuilder.wSelf("color", colorAttrs));
    }

    // 20. w:spacing — Character spacing
    if (formatting.characterSpacing !== undefined) {
      rPrChildren.push(
        XMLBuilder.wSelf("spacing", { "w:val": formatting.characterSpacing })
      );
    }

    // 21. w:w — Horizontal scaling
    if (formatting.scaling !== undefined) {
      rPrChildren.push(XMLBuilder.wSelf("w", { "w:val": formatting.scaling }));
    }

    // 22. w:kern — Kerning
    if (formatting.kerning !== undefined && formatting.kerning !== null) {
      rPrChildren.push(
        XMLBuilder.wSelf("kern", { "w:val": formatting.kerning })
      );
    }

    // 23. w:position — Vertical position
    if (formatting.position !== undefined) {
      rPrChildren.push(
        XMLBuilder.wSelf("position", { "w:val": formatting.position })
      );
    }

    // 24/25. w:sz / w:szCs — Font size / Font size complex script
    if (formatting.size !== undefined) {
      const halfPoints = pointsToHalfPoints(formatting.size);
      rPrChildren.push(XMLBuilder.wSelf("sz", { "w:val": halfPoints }));
      const csHalfPoints = formatting.sizeCs !== undefined ? pointsToHalfPoints(formatting.sizeCs) : halfPoints;
      rPrChildren.push(XMLBuilder.wSelf("szCs", { "w:val": csHalfPoints }));
    } else if (formatting.sizeCs !== undefined) {
      const csHalfPoints = pointsToHalfPoints(formatting.sizeCs);
      rPrChildren.push(XMLBuilder.wSelf("szCs", { "w:val": csHalfPoints }));
    }

    // 26. w:highlight — Highlight color
    if (formatting.highlight) {
      rPrChildren.push(
        XMLBuilder.wSelf("highlight", { "w:val": formatting.highlight })
      );
    }

    // 27. w:u — Underline
    // When a character style is applied (e.g., Hyperlink) and underline is explicitly false,
    // we need to output <w:u w:val="none"/> to prevent the style's underline from being inherited.
    if (formatting.underline) {
      const underlineValue =
        typeof formatting.underline === "string"
          ? formatting.underline
          : "single";
      const uAttrs: Record<string, string | number> = { "w:val": underlineValue };
      if (formatting.underlineColor) uAttrs["w:color"] = formatting.underlineColor;
      if (formatting.underlineThemeColor) uAttrs["w:themeColor"] = formatting.underlineThemeColor;
      if (formatting.underlineThemeTint !== undefined) uAttrs["w:themeTint"] = formatting.underlineThemeTint.toString(16).toUpperCase().padStart(2, '0');
      if (formatting.underlineThemeShade !== undefined) uAttrs["w:themeShade"] = formatting.underlineThemeShade.toString(16).toUpperCase().padStart(2, '0');
      rPrChildren.push(XMLBuilder.wSelf("u", uAttrs));
    } else if (formatting.underline === false && formatting.characterStyle) {
      rPrChildren.push(XMLBuilder.wSelf("u", { "w:val": "none" }));
    }

    // 28. w:effect — Text effect/animation
    if (formatting.effect) {
      rPrChildren.push(
        XMLBuilder.wSelf("effect", { "w:val": formatting.effect })
      );
    }

    // 29. w:bdr — Text border
    if (formatting.border) {
      const bdrAttrs: Record<string, string | number> = {};
      if (formatting.border.style) bdrAttrs["w:val"] = formatting.border.style;
      if (formatting.border.size !== undefined)
        bdrAttrs["w:sz"] = formatting.border.size;
      if (formatting.border.color)
        bdrAttrs["w:color"] = formatting.border.color;
      if (formatting.border.space !== undefined)
        bdrAttrs["w:space"] = formatting.border.space;

      if (Object.keys(bdrAttrs).length > 0) {
        rPrChildren.push(XMLBuilder.wSelf("bdr", bdrAttrs));
      }
    }

    // 30. w:shd — Character shading
    if (formatting.shading) {
      const shdAttrs = buildShadingAttributes(formatting.shading);
      if (Object.keys(shdAttrs).length > 0) {
        rPrChildren.push(XMLBuilder.wSelf("shd", shdAttrs));
      }
    }

    // 31. w:fitText — Fit text to width
    if (formatting.fitText !== undefined) {
      rPrChildren.push(
        XMLBuilder.wSelf("fitText", { "w:val": formatting.fitText })
      );
    }

    // 32. w:vertAlign — Subscript/superscript
    if (formatting.subscript) {
      rPrChildren.push(XMLBuilder.wSelf("vertAlign", { "w:val": "subscript" }));
    }
    if (formatting.superscript) {
      rPrChildren.push(
        XMLBuilder.wSelf("vertAlign", { "w:val": "superscript" })
      );
    }

    // 33. w:rtl — Right-to-left text
    if (formatting.rtl) {
      rPrChildren.push(XMLBuilder.wSelf("rtl", { "w:val": "1" }));
    }

    // 34. w:cs — Complex script flag
    if (formatting.complexScript) {
      rPrChildren.push(XMLBuilder.wSelf("cs", { "w:val": "1" }));
    }

    // 35. w:em — Emphasis marks
    if (formatting.emphasis) {
      rPrChildren.push(
        XMLBuilder.wSelf("em", { "w:val": formatting.emphasis })
      );
    }

    // 36. w:lang — Language
    if (formatting.language) {
      rPrChildren.push(
        XMLBuilder.wSelf("lang", { "w:val": formatting.language })
      );
    }

    // 37. w:eastAsianLayout — East Asian layout
    if (formatting.eastAsianLayout) {
      const layout = formatting.eastAsianLayout;
      const attrs: Record<string, string | number> = {};
      if (layout.id !== undefined) attrs["w:id"] = layout.id;
      if (layout.vert) attrs["w:vert"] = "1";
      if (layout.vertCompress) attrs["w:vertCompress"] = "1";
      if (layout.combine) attrs["w:combine"] = "1";
      if (layout.combineBrackets)
        attrs["w:combineBrackets"] = layout.combineBrackets;

      if (Object.keys(attrs).length > 0) {
        rPrChildren.push(XMLBuilder.wSelf("eastAsianLayout", attrs));
      }
    }

    // 38. w:specVanish — Special vanish
    if (formatting.specVanish) {
      rPrChildren.push(XMLBuilder.wSelf("specVanish", { "w:val": "1" }));
    }

    // 39. w:oMath — (not currently generated)

    // 40. Raw w14: namespace elements (Word 2010+ text effects, after all schema elements)
    if (formatting.rawW14Properties && formatting.rawW14Properties.length > 0) {
      for (const rawXml of formatting.rawW14Properties) {
        rPrChildren.push({ name: "__rawXml", rawXml } as XMLElement);
      }
    }

    // Return null if no properties (prevents empty <w:rPr/> elements)
    if (rPrChildren.length === 0) {
      return null;
    }

    return XMLBuilder.w("rPr", undefined, rPrChildren);
  }

  /**
   * Creates a new Run instance
   *
   * Factory method for creating a Run with text and optional formatting.
   *
   * @param text - The text content
   * @param formatting - Optional formatting to apply
   * @returns New Run instance
   *
   * @example
   * ```typescript
   * const run = Run.create('Hello World');
   * const boldRun = Run.create('Bold Text', { bold: true });
   * ```
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

    const newText =
      currentText.slice(0, index) + text + currentText.slice(index);
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
    if (start > end) [start, end] = [end, start]; // Swap if reversed

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
   * Clears run formatting properties that MATCH a style definition.
   * The inverse of clearFormattingConflicts: removes properties whose values
   * are identical to the style, so the run inherits those values from the style.
   * Preserves properties that differ from the style (direct overrides) and
   * properties not defined in the style.
   *
   * This enables style inheritance: when direct formatting matches the style,
   * removing it allows future style definition changes to propagate automatically.
   *
   * @param styleRunFormatting - Run formatting from the style definition to compare against
   * @returns This run for method chaining
   * @example
   * ```typescript
   * // Style says: black, 12pt Verdana
   * // Run has: black, 12pt Verdana, bold
   * run.clearMatchingFormatting({
   *   color: '000000',
   *   size: 12,
   *   font: 'Verdana'
   * });
   * // Result: color/size/font cleared (inherit from style), bold kept (not in style)
   * ```
   */
  clearMatchingFormatting(styleRunFormatting: Partial<RunFormatting>): this {
    const matchingProperties: (keyof RunFormatting)[] = [];

    for (const key in this.formatting) {
      const propKey = key as keyof RunFormatting;

      // Skip if style doesn't define this property (preserve run's property)
      if (styleRunFormatting[propKey] === undefined) {
        continue;
      }

      // If run's value matches the style, it's redundant direct formatting
      if (this.formatting[propKey] === styleRunFormatting[propKey]) {
        matchingProperties.push(propKey);
      }
    }

    // Clear matching properties so run inherits from style
    for (const prop of matchingProperties) {
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
