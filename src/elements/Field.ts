/**
 * Field - Represents a dynamic field in a Word document
 *
 * Fields are used for dynamic content like page numbers, dates, document properties, etc.
 * They are represented using the <w:fldSimple> element with field codes.
 */

import { XMLElement } from '../xml/XMLBuilder';
import { RunFormatting, FormFieldData } from './Run';
import { ParsedHyperlinkInstruction, parseHyperlinkInstruction, isHyperlinkInstruction } from './FieldHelpers';
import type { Revision } from './Revision';
import { pointsToHalfPoints } from '../utils/units';

/**
 * Common field types
 */
export type FieldType =
  | 'PAGE'           // Current page number
  | 'NUMPAGES'       // Total number of pages
  | 'DATE'           // Current date
  | 'TIME'           // Current time
  | 'AUTHOR'         // Document author
  | 'TITLE'          // Document title
  | 'FILENAME'       // Document filename
  | 'FILENAMEWITHPATH' // Document filename with path
  | 'SUBJECT'        // Document subject
  | 'KEYWORDS'       // Document keywords
  | 'CREATEDATE'     // Document creation date
  | 'SAVEDATE'       // Last save date
  | 'PRINTDATE'      // Last print date
  | 'SECTIONPAGES'   // Pages in current section
  | 'SECTION'        // Current section number
  | 'REF'            // Cross-reference to bookmark
  | 'HYPERLINK'      // Hyperlink field
  | 'SEQ'            // Sequence numbering
  | 'TC'             // Table of contents entry
  | 'XE'             // Index entry
  | 'IF'             // Conditional field
  | 'MERGEFIELD'     // Mail merge field
  | 'INCLUDE'        // Include text from external file
  | 'INCLUDETEXT'    // Include text from external file (alias)
  | 'CUSTOM';        // Custom field type for unknown/specialized fields

/**
 * Field properties
 */
export interface FieldProperties {
  /** Field type */
  type: FieldType;
  /** Field instruction (e.g., 'PAGE \* MERGEFORMAT') */
  instruction?: string;
  /** Format switches (e.g., '\\* MERGEFORMAT') */
  format?: string;
  /** Date/time format (e.g., 'MMMM d, yyyy') */
  dateFormat?: string;
  /** Preserve formatting during updates */
  preserveFormatting?: boolean;
  /** Run formatting for field result */
  formatting?: RunFormatting;
}

/**
 * Represents a dynamic field
 */
export class Field {
  private type: FieldType;
  private instruction: string;
  private formatting?: RunFormatting;

  /**
   * Creates a new field
   * @param properties Field properties
   */
  constructor(properties: FieldProperties) {
    this.type = properties.type;
    this.formatting = properties.formatting;

    // Build field instruction
    if (properties.instruction) {
      this.instruction = properties.instruction;
    } else {
      this.instruction = this.buildInstruction(properties);
    }
  }

  /**
   * Builds field instruction from properties
   */
  private buildInstruction(properties: FieldProperties): string {
    let instruction = properties.type;

    // Add date format for date/time fields
    if (properties.dateFormat && this.isDateField(properties.type)) {
      instruction += ` \\@ "${properties.dateFormat}"`;
    }

    // Add format switch
    if (properties.format) {
      instruction += ` ${properties.format}`;
    } else if (properties.preserveFormatting !== false) {
      // Add MERGEFORMAT by default to preserve formatting
      instruction += ' \\* MERGEFORMAT';
    }

    return instruction;
  }

  /**
   * Checks if field type is a date field
   */
  private isDateField(type: FieldType): boolean {
    return ['DATE', 'TIME', 'CREATEDATE', 'SAVEDATE', 'PRINTDATE'].includes(type);
  }

  /**
   * Gets the field type
   */
  getType(): FieldType {
    return this.type;
  }

  /**
   * Gets the field instruction
   */
  getInstruction(): string {
    return this.instruction;
  }

  /**
   * Sets run formatting for the field
   */
  setFormatting(formatting: RunFormatting): this {
    this.formatting = formatting;
    return this;
  }

  /**
   * Gets run formatting
   */
  getFormatting(): RunFormatting | undefined {
    return this.formatting;
  }

  /**
   * Checks if this field is a HYPERLINK field
   * @returns True if the field type is HYPERLINK or instruction starts with HYPERLINK
   */
  isHyperlinkField(): boolean {
    return this.type === 'HYPERLINK' ||
           this.instruction.trim().toUpperCase().startsWith('HYPERLINK');
  }

  /**
   * Sets text color for the field
   * @param color Color in hex format (e.g., '0000FF')
   * @returns This field for chaining
   */
  setColor(color: string): this {
    if (!this.formatting) {
      this.formatting = {};
    }
    this.formatting.color = color.replace('#', '');
    return this;
  }

  /**
   * Generates XML for the field.
   * Per ECMA-376, w:fldSimple is a paragraph-level element that CONTAINS w:r children.
   * Structure: <w:fldSimple w:instr="..."><w:r><w:rPr/><w:t>...</w:t></w:r></w:fldSimple>
   * The fldSimple element should be added directly to paragraph children (not wrapped in w:r).
   */
  toXML(): XMLElement {
    // Build the inner run with optional formatting
    const runChildren: XMLElement[] = [];
    if (this.formatting) {
      runChildren.push(this.createRunProperties());
    }
    runChildren.push({
      name: 'w:t',
      children: [this.getPlaceholderText()],
    });

    return {
      name: 'w:fldSimple',
      attributes: {
        'w:instr': this.instruction,
      },
      children: [{
        name: 'w:r',
        children: runChildren,
      }],
    };
  }

  /**
   * Gets placeholder text for the field
   */
  private getPlaceholderText(): string {
    switch (this.type) {
      case 'PAGE':
        return '1';
      case 'NUMPAGES':
        return '1';
      case 'SECTIONPAGES':
        return '1';
      case 'SECTION':
        return '1';
      case 'DATE':
        return new Date().toLocaleDateString();
      case 'TIME':
        return new Date().toLocaleTimeString();
      case 'CREATEDATE':
      case 'SAVEDATE':
      case 'PRINTDATE':
        return new Date().toLocaleDateString();
      case 'FILENAME':
        return 'Document';
      case 'FILENAMEWITHPATH':
        return 'C:\\Document.docx';
      case 'AUTHOR':
        return 'Author';
      case 'TITLE':
        return 'Title';
      case 'SUBJECT':
        return 'Subject';
      case 'KEYWORDS':
        return 'Keywords';
      case 'REF':
        return '1'; // Reference typically shows page number or heading text
      case 'HYPERLINK':
        return 'Link'; // Hyperlink displays the link text
      case 'SEQ':
        return '1'; // Sequence shows the current number
      case 'TC':
        return ''; // TC fields are hidden
      case 'XE':
        return ''; // XE fields are hidden
      default:
        return '';
    }
  }

  /**
   * Creates run properties XML
   */
  private createRunProperties(): XMLElement {
    const children: XMLElement[] = [];

    if (!this.formatting) {
      return { name: 'w:rPr', children };
    }

    // Per ECMA-376 CT_RPr ordering:
    // rFonts → b → i → strike → color → sz/szCs → highlight → u

    if (this.formatting.font) {
      children.push({
        name: 'w:rFonts',
        attributes: {
          'w:ascii': this.formatting.font,
          'w:hAnsi': this.formatting.font,
          'w:cs': this.formatting.font,
        },
        selfClosing: true,
      });
    }

    if (this.formatting.bold) {
      children.push({ name: 'w:b', selfClosing: true });
    }

    if (this.formatting.italic) {
      children.push({ name: 'w:i', selfClosing: true });
    }

    if (this.formatting.strike) {
      children.push({ name: 'w:strike', selfClosing: true });
    }

    if (this.formatting.color) {
      const color = this.formatting.color.replace('#', '');
      children.push({
        name: 'w:color',
        attributes: { 'w:val': color },
        selfClosing: true,
      });
    }

    if (this.formatting.size) {
      const sizeValue = pointsToHalfPoints(this.formatting.size).toString();
      children.push({
        name: 'w:sz',
        attributes: { 'w:val': sizeValue },
        selfClosing: true,
      });
      children.push({
        name: 'w:szCs',
        attributes: { 'w:val': sizeValue },
        selfClosing: true,
      });
    }

    if (this.formatting.highlight) {
      children.push({
        name: 'w:highlight',
        attributes: { 'w:val': this.formatting.highlight },
        selfClosing: true,
      });
    }

    if (this.formatting.underline) {
      const val = typeof this.formatting.underline === 'string'
        ? this.formatting.underline
        : 'single';
      children.push({
        name: 'w:u',
        attributes: { 'w:val': val },
        selfClosing: true,
      });
    }

    return { name: 'w:rPr', children };
  }

  /**
   * Creates a page number field
   * @param formatting Optional run formatting
   */
  static createPageNumber(formatting?: RunFormatting): Field {
    return new Field({
      type: 'PAGE',
      formatting,
    });
  }

  /**
   * Creates a total pages field
   * @param formatting Optional run formatting
   */
  static createTotalPages(formatting?: RunFormatting): Field {
    return new Field({
      type: 'NUMPAGES',
      formatting,
    });
  }

  /**
   * Creates a date field
   * @param format Date format (e.g., 'MMMM d, yyyy')
   * @param formatting Optional run formatting
   */
  static createDate(format?: string, formatting?: RunFormatting): Field {
    return new Field({
      type: 'DATE',
      dateFormat: format,
      formatting,
    });
  }

  /**
   * Creates a time field
   * @param format Time format
   * @param formatting Optional run formatting
   */
  static createTime(format?: string, formatting?: RunFormatting): Field {
    return new Field({
      type: 'TIME',
      dateFormat: format,
      formatting,
    });
  }

  /**
   * Creates a filename field
   * @param includePath Whether to include full path
   * @param formatting Optional run formatting
   */
  static createFilename(includePath = false, formatting?: RunFormatting): Field {
    return new Field({
      type: includePath ? 'FILENAMEWITHPATH' : 'FILENAME',
      formatting,
    });
  }

  /**
   * Creates an author field
   * @param formatting Optional run formatting
   */
  static createAuthor(formatting?: RunFormatting): Field {
    return new Field({
      type: 'AUTHOR',
      formatting,
    });
  }

  /**
   * Creates a title field
   * @param formatting Optional run formatting
   */
  static createTitle(formatting?: RunFormatting): Field {
    return new Field({
      type: 'TITLE',
      formatting,
    });
  }

  /**
   * Creates a section pages field (pages in current section)
   * @param formatting Optional run formatting
   */
  static createSectionPages(formatting?: RunFormatting): Field {
    return new Field({
      type: 'SECTIONPAGES',
      formatting,
    });
  }

  /**
   * Creates a cross-reference field
   * @param bookmark Bookmark name to reference
   * @param format Reference format (\h for hyperlink, \p for page number, etc.)
   * @param formatting Optional run formatting
   */
  static createRef(bookmark: string, format?: string, formatting?: RunFormatting): Field {
    const formatSwitch = format || '\\h'; // Default to hyperlink format
    const instruction = `REF ${bookmark} ${formatSwitch} \\* MERGEFORMAT`;

    return new Field({
      type: 'REF',
      instruction,
      formatting,
    });
  }

  /**
   * Creates a hyperlink field
   * @param url The URL to link to
   * @param displayText The text to display
   * @param tooltip Optional tooltip text
   * @param formatting Optional run formatting
   */
  static createHyperlink(
    url: string,
    displayText: string = url,
    tooltip?: string,
    formatting?: RunFormatting
  ): Field {
    let instruction = `HYPERLINK "${url}"`;

    if (tooltip) {
      instruction += ` \\o "${tooltip}"`;
    }

    instruction += ' \\* MERGEFORMAT';

    return new Field({
      type: 'HYPERLINK',
      instruction,
      formatting,
    });
  }

  /**
   * Creates a sequence numbering field
   * @param identifier Sequence identifier (e.g., 'Figure', 'Table')
   * @param format Number format (\* ARABIC, \* ROMAN, etc.)
   * @param formatting Optional run formatting
   */
  static createSeq(
    identifier: string,
    format?: string,
    formatting?: RunFormatting
  ): Field {
    let instruction = `SEQ ${identifier}`;

    if (format) {
      instruction += ` ${format}`;
    } else {
      instruction += ' \\* ARABIC'; // Default to Arabic numerals
    }

    instruction += ' \\* MERGEFORMAT';

    return new Field({
      type: 'SEQ',
      instruction,
      formatting,
    });
  }

  /**
   * Creates a table of contents entry field (TC)
   * @param text Entry text
   * @param level TOC level (1-9)
   * @param formatting Optional run formatting
   */
  static createTCEntry(
    text: string,
    level = 1,
    formatting?: RunFormatting
  ): Field {
    if (level < 1 || level > 9) {
      throw new Error('TC level must be between 1 and 9');
    }

    const instruction = `TC "${text}" \\f C \\l ${level}`;

    return new Field({
      type: 'TC',
      instruction,
      formatting,
    });
  }

  /**
   * Creates an index entry field (XE)
   * @param text Entry text
   * @param subEntry Optional sub-entry text
   * @param formatting Optional run formatting
   */
  static createXEEntry(
    text: string,
    subEntry?: string,
    formatting?: RunFormatting
  ): Field {
    let instruction = `XE "${text}"`;

    if (subEntry) {
      instruction += `:${subEntry}`;
    }

    return new Field({
      type: 'XE',
      instruction,
      formatting,
    });
  }

  /**
   * Creates a custom field with instruction
   * @param instruction Field instruction code
   * @param formatting Optional run formatting
   */
  static createCustom(instruction: string, formatting?: RunFormatting): Field {
    return new Field({
      type: 'PAGE', // Placeholder type
      instruction,
      formatting,
    });
  }

  /**
   * Creates a field from properties
   * @param properties Field properties
   */
  static create(properties: FieldProperties): Field {
    return new Field(properties);
  }
}

/**
 * Field character type for complex fields
 */
export type FieldCharType = 'begin' | 'separate' | 'end';

/**
 * Complex field properties
 * Complex fields use begin/separate/end structure instead of fldSimple
 */
export interface ComplexFieldProperties {
  /** Field instruction (e.g., " TOC \\o \"1-3\" \\h \\z \\u ") */
  instruction: string;

  /** Current field result text (optional) */
  result?: string;

  /** Run formatting for instruction */
  instructionFormatting?: RunFormatting;

  /** Run formatting for result */
  resultFormatting?: RunFormatting;

  /** Nested fields to include within this field */
  nestedFields?: ComplexField[];

  /** Custom XML content for result section */
  resultContent?: XMLElement[];

  /** Whether field spans multiple paragraphs */
  multiParagraph?: boolean;

  /** Parsed HYPERLINK instruction components (auto-populated if instruction is HYPERLINK) */
  parsedHyperlink?: ParsedHyperlinkInstruction;

  /**
   * Whether the field has a result section (w:fldSep was present during parsing)
   * Per ECMA-376, fields without results skip the separator element.
   * This flag distinguishes between:
   * - hasResult=true, result="" → Field had separator but result was empty
   * - hasResult=false, result="" → Field never had a result section (empty field)
   */
  hasResult?: boolean;

  /** Form field data (w:ffData) from begin field char per ECMA-376 §17.16.17 */
  formFieldData?: FormFieldData;
}

/**
 * Represents a complex field (begin/separate/end structure)
 * Used for TOC, cross-references, and other advanced fields
 *
 * Structure:
 * <w:r><w:fldChar w:fldCharType="begin"/></w:r>
 * <w:r><w:instrText>INSTRUCTION</w:instrText></w:r>
 * <w:r><w:fldChar w:fldCharType="separate"/></w:r>
 * <w:r><w:t>RESULT</w:t></w:r>
 * <w:r><w:fldChar w:fldCharType="end"/></w:r>
 */
export class ComplexField {
  private instruction: string;
  private result?: string;
  private instructionFormatting?: RunFormatting;
  private resultFormatting?: RunFormatting;
  private nestedFields: ComplexField[];
  private resultContent: XMLElement[];
  private multiParagraph: boolean;
  private parsedHyperlink?: ParsedHyperlinkInstruction;
  /** Revisions that wrap the result section (for tracked changes in field content) */
  private resultRevisions: Revision[] = [];
  /**
   * Whether the field has a result section (w:fldSep was present during parsing)
   * Per ECMA-376, fields without results skip the separator element.
   */
  private _hasResultSection = false;
  private _formFieldData?: FormFieldData;

  /**
   * Creates a new complex field
   * @param properties Complex field properties
   */
  constructor(properties: ComplexFieldProperties) {
    this.instruction = properties.instruction;
    this.result = properties.result;
    this.instructionFormatting = properties.instructionFormatting;
    this.resultFormatting = properties.resultFormatting;
    this.nestedFields = properties.nestedFields || [];
    this.resultContent = properties.resultContent || [];
    this.multiParagraph = properties.multiParagraph || false;
    this._hasResultSection = properties.hasResult ?? false;
    this._formFieldData = properties.formFieldData;

    // Auto-parse HYPERLINK instruction if provided or detected
    if (properties.parsedHyperlink) {
      this.parsedHyperlink = properties.parsedHyperlink;
    } else if (isHyperlinkInstruction(this.instruction)) {
      this.parsedHyperlink = parseHyperlinkInstruction(this.instruction) || undefined;
    }
  }

  /**
   * Gets the field instruction
   */
  getInstruction(): string {
    return this.instruction;
  }

  /**
   * Sets the field instruction
   */
  setInstruction(instruction: string): this {
    this.instruction = instruction;
    return this;
  }

  /**
   * Gets form field data (w:ffData) if present
   */
  getFormFieldData(): FormFieldData | undefined {
    return this._formFieldData;
  }

  /**
   * Gets the field result text
   */
  getResult(): string | undefined {
    return this.result;
  }

  /**
   * Sets the field result text
   */
  setResult(result: string): this {
    this.result = result;
    return this;
  }

  /**
   * Checks if the field has a result section
   *
   * Per ECMA-376, complex fields may or may not have a result section.
   * Fields without results (like TOC markers or empty PAGE fields) skip
   * the w:fldSep (separator) element entirely.
   *
   * This method distinguishes between:
   * - `hasResultSection() === true && getResult() === ""` → Field had separator but result was empty
   * - `hasResultSection() === false && getResult() === undefined` → Field never had a result section
   *
   * @returns True if the field has a result section (w:fldSep was present)
   *
   * @example
   * ```typescript
   * const field = paragraph.getFields()[0];
   * if (field instanceof ComplexField) {
   *   if (field.hasResultSection()) {
   *     console.log('Field result:', field.getResult());
   *   } else {
   *     console.log('Field has no result section (empty field)');
   *   }
   * }
   * ```
   */
  hasResultSection(): boolean {
    return this._hasResultSection;
  }

  /**
   * Sets instruction formatting
   */
  setInstructionFormatting(formatting: RunFormatting): this {
    this.instructionFormatting = formatting;
    return this;
  }

  /**
   * Sets result formatting
   */
  setResultFormatting(formatting: RunFormatting): this {
    this.resultFormatting = formatting;
    return this;
  }

  /**
   * Gets the parsed HYPERLINK instruction components
   * Returns undefined if this is not a HYPERLINK field
   */
  getParsedHyperlink(): ParsedHyperlinkInstruction | undefined {
    return this.parsedHyperlink;
  }

  /**
   * Checks if this field is a HYPERLINK field
   */
  isHyperlinkField(): boolean {
    return this.parsedHyperlink !== undefined;
  }

  /**
   * Gets the full URL for HYPERLINK fields (combining base URL and anchor)
   * Returns undefined if not a HYPERLINK field
   */
  getHyperlinkUrl(): string | undefined {
    return this.parsedHyperlink?.fullUrl;
  }

  /**
   * Adds a nested field within this field
   * Nested fields appear between instruction and separator
   */
  addNestedField(field: ComplexField): this {
    this.nestedFields.push(field);
    return this;
  }

  /**
   * Gets all nested fields
   */
  getNestedFields(): ComplexField[] {
    return [...this.nestedFields];
  }

  /**
   * Removes a nested field at the specified index
   * @param index - Index of the nested field to remove (0-based)
   * @returns True if removed, false if index out of bounds
   *
   * @example
   * ```typescript
   * const field = new ComplexField({ instruction: 'TOC' });
   * field.addNestedField(nested1);
   * field.addNestedField(nested2);
   * field.removeNestedField(0); // Removes nested1
   * ```
   */
  removeNestedField(index: number): boolean {
    if (index < 0 || index >= this.nestedFields.length) {
      return false;
    }
    this.nestedFields.splice(index, 1);
    return true;
  }

  /**
   * Gets the count of nested fields
   * @returns Number of nested fields
   */
  getNestedFieldCount(): number {
    return this.nestedFields.length;
  }

  /**
   * Clears all nested fields
   * @returns This field for chaining
   */
  clearNestedFields(): this {
    this.nestedFields = [];
    return this;
  }

  /**
   * Updates the field instruction
   * @param instruction - New field instruction (e.g., 'TOC \\o "1-3"')
   * @returns This field for chaining
   *
   * @example
   * ```typescript
   * const field = new ComplexField({ instruction: 'DATE' });
   * field.updateInstruction('DATE \\@ "yyyy-MM-dd"');
   * ```
   */
  updateInstruction(instruction: string): this {
    this.instruction = instruction;
    return this;
  }

  /**
   * Adds custom XML content to the result section
   */
  addResultContent(content: XMLElement): this {
    this.resultContent.push(content);
    return this;
  }

  /**
   * Gets result content XML elements
   */
  getResultContent(): XMLElement[] {
    return [...this.resultContent];
  }

  /**
   * Sets whether this field spans multiple paragraphs
   */
  setMultiParagraph(multiParagraph: boolean): this {
    this.multiParagraph = multiParagraph;
    return this;
  }

  /**
   * Gets whether this field spans multiple paragraphs
   */
  isMultiParagraph(): boolean {
    return this.multiParagraph;
  }

  /**
   * Sets revisions that wrap the result section
   * These are tracked changes (w:ins, w:del) that need to wrap the result AND end marker
   * @param revisions Array of Revision objects
   */
  setResultRevisions(revisions: Revision[]): this {
    this.resultRevisions = revisions;
    return this;
  }

  /**
   * Adds a revision to the result section
   * @param revision Revision to add
   */
  addResultRevision(revision: Revision): this {
    this.resultRevisions.push(revision);
    return this;
  }

  /**
   * Gets revisions that wrap the result section
   */
  getResultRevisions(): Revision[] {
    return [...this.resultRevisions];
  }

  /**
   * Checks if this field has revisions in the result section
   */
  hasResultRevisions(): boolean {
    return this.resultRevisions.length > 0;
  }

  /**
   * Generates XML for the complex field
   * Returns array of run elements (begin, instr, sep, result, end)
   */
  toXML(): XMLElement[] {
    const runs: XMLElement[] = [];

    // 1. Begin marker run
    const beginFldChar: XMLElement = this._formFieldData
      ? this.buildFldCharWithFfData()
      : {
          name: 'w:fldChar',
          attributes: { 'w:fldCharType': 'begin' },
          selfClosing: true,
        };
    runs.push({
      name: 'w:r',
      children: [beginFldChar],
    });

    // 2. Instruction run
    const instrChildren: XMLElement[] = [];
    if (this.instructionFormatting) {
      instrChildren.push(this.createRunProperties(this.instructionFormatting));
    }
    instrChildren.push({
      name: 'w:instrText',
      attributes: { 'xml:space': 'preserve' },
      children: [this.instruction],
    });
    runs.push({
      name: 'w:r',
      children: instrChildren,
    });

    // 2a. Nested fields (if any)
    for (const nestedField of this.nestedFields) {
      runs.push(...nestedField.toXML());
    }

    // 3. Separator run
    runs.push({
      name: 'w:r',
      children: [
        {
          name: 'w:fldChar',
          attributes: { 'w:fldCharType': 'separate' },
          selfClosing: true,
        },
      ],
    });

    // 4. Result content (prioritize custom XML content, then simple text)
    // Design note: For INCLUDEPICTURE fields, the parser stores the w:drawing
    // content in resultContent so it survives the parser→generator round-trip.
    // When flattenFieldCodes() is active, the full field structure is emitted
    // here, then _postProcessDocumentXml() strips the INCLUDEPICTURE scaffolding
    // (fldChar/instrText runs) from the final XML while preserving the drawing.
    // Non-INCLUDEPICTURE fields emit their complete structure unchanged.
    if (this.resultContent.length > 0) {
      // Use custom XML content
      runs.push(...this.resultContent);
    } else if (this.result) {
      // Use simple text result
      const resultChildren: XMLElement[] = [];
      if (this.resultFormatting) {
        resultChildren.push(this.createRunProperties(this.resultFormatting));
      }
      resultChildren.push({
        name: 'w:t',
        attributes: { 'xml:space': 'preserve' },
        children: [this.result],
      });
      runs.push({
        name: 'w:r',
        children: resultChildren,
      });
    }

    // 4a. Result revisions (tracked changes within the result section)
    // These MUST appear between the separator and end marker per ECMA-376
    // The revisions contain the actual field result content wrapped in w:ins or w:del
    for (const revision of this.resultRevisions) {
      const revisionXml = revision.toXML();
      if (revisionXml) {
        runs.push(revisionXml);
      }
    }

    // 5. End marker run
    runs.push({
      name: 'w:r',
      children: [
        {
          name: 'w:fldChar',
          attributes: { 'w:fldCharType': 'end' },
          selfClosing: true,
        },
      ],
    });

    return runs;
  }

  /**
   * Creates run properties XML from formatting
   */
  private createRunProperties(formatting: RunFormatting): XMLElement {
    const children: XMLElement[] = [];

    if (formatting.bold) {
      children.push({ name: 'w:b', selfClosing: true });
    }

    if (formatting.italic) {
      children.push({ name: 'w:i', selfClosing: true });
    }

    if (formatting.underline) {
      const val =
        typeof formatting.underline === 'string'
          ? formatting.underline
          : 'single';
      children.push({
        name: 'w:u',
        attributes: { 'w:val': val },
        selfClosing: true,
      });
    }

    if (formatting.strike) {
      children.push({ name: 'w:strike', selfClosing: true });
    }

    if (formatting.font) {
      children.push({
        name: 'w:rFonts',
        attributes: {
          'w:ascii': formatting.font,
          'w:hAnsi': formatting.font,
          'w:cs': formatting.font,
        },
        selfClosing: true,
      });
    }

    if (formatting.size) {
      const sizeValue = pointsToHalfPoints(formatting.size).toString();
      children.push({
        name: 'w:sz',
        attributes: { 'w:val': sizeValue },
        selfClosing: true,
      });
      children.push({
        name: 'w:szCs',
        attributes: { 'w:val': sizeValue },
        selfClosing: true,
      });
    }

    if (formatting.color) {
      const color = formatting.color.replace('#', '');
      children.push({
        name: 'w:color',
        attributes: { 'w:val': color },
        selfClosing: true,
      });
    }

    if (formatting.highlight) {
      children.push({
        name: 'w:highlight',
        attributes: { 'w:val': formatting.highlight },
        selfClosing: true,
      });
    }

    return { name: 'w:rPr', children };
  }

  /**
   * Builds a w:fldChar begin element with w:ffData child
   */
  private buildFldCharWithFfData(): XMLElement {
    const ffd = this._formFieldData!;
    const ffDataChildren: (string | XMLElement)[] = [];

    if (ffd.name) {
      ffDataChildren.push({ name: 'w:name', attributes: { 'w:val': ffd.name }, selfClosing: true });
    }
    if (ffd.enabled !== undefined) {
      if (ffd.enabled) {
        ffDataChildren.push({ name: 'w:enabled', selfClosing: true });
      } else {
        ffDataChildren.push({ name: 'w:enabled', attributes: { 'w:val': '0' }, selfClosing: true });
      }
    }
    if (ffd.calcOnExit !== undefined) {
      ffDataChildren.push({ name: 'w:calcOnExit', attributes: { 'w:val': ffd.calcOnExit ? '1' : '0' }, selfClosing: true });
    }
    if (ffd.helpText) {
      ffDataChildren.push({ name: 'w:helpText', attributes: { 'w:type': 'text', 'w:val': ffd.helpText }, selfClosing: true });
    }
    if (ffd.statusText) {
      ffDataChildren.push({ name: 'w:statusText', attributes: { 'w:type': 'text', 'w:val': ffd.statusText }, selfClosing: true });
    }
    if (ffd.entryMacro) {
      ffDataChildren.push({ name: 'w:entryMacro', attributes: { 'w:val': ffd.entryMacro }, selfClosing: true });
    }
    if (ffd.exitMacro) {
      ffDataChildren.push({ name: 'w:exitMacro', attributes: { 'w:val': ffd.exitMacro }, selfClosing: true });
    }

    if (ffd.fieldType) {
      switch (ffd.fieldType.type) {
        case 'textInput': {
          const tiChildren: (string | XMLElement)[] = [];
          if (ffd.fieldType.inputType) tiChildren.push({ name: 'w:type', attributes: { 'w:val': ffd.fieldType.inputType }, selfClosing: true });
          if (ffd.fieldType.defaultValue) tiChildren.push({ name: 'w:default', attributes: { 'w:val': ffd.fieldType.defaultValue }, selfClosing: true });
          if (ffd.fieldType.maxLength !== undefined) tiChildren.push({ name: 'w:maxLength', attributes: { 'w:val': String(ffd.fieldType.maxLength) }, selfClosing: true });
          if (ffd.fieldType.format) tiChildren.push({ name: 'w:format', attributes: { 'w:val': ffd.fieldType.format }, selfClosing: true });
          ffDataChildren.push({ name: 'w:textInput', children: tiChildren });
          break;
        }
        case 'checkBox': {
          const cbChildren: (string | XMLElement)[] = [];
          if (ffd.fieldType.size === 'auto') {
            cbChildren.push({ name: 'w:sizeAuto', selfClosing: true });
          } else if (ffd.fieldType.size !== undefined) {
            cbChildren.push({ name: 'w:size', attributes: { 'w:val': String(ffd.fieldType.size) }, selfClosing: true });
          }
          if (ffd.fieldType.defaultChecked !== undefined) cbChildren.push({ name: 'w:default', attributes: { 'w:val': ffd.fieldType.defaultChecked ? '1' : '0' }, selfClosing: true });
          if (ffd.fieldType.checked !== undefined) cbChildren.push({ name: 'w:checked', attributes: { 'w:val': ffd.fieldType.checked ? '1' : '0' }, selfClosing: true });
          ffDataChildren.push({ name: 'w:checkBox', children: cbChildren });
          break;
        }
        case 'dropDownList': {
          const ddChildren: (string | XMLElement)[] = [];
          if (ffd.fieldType.result !== undefined) ddChildren.push({ name: 'w:result', attributes: { 'w:val': String(ffd.fieldType.result) }, selfClosing: true });
          if (ffd.fieldType.defaultResult !== undefined) ddChildren.push({ name: 'w:default', attributes: { 'w:val': String(ffd.fieldType.defaultResult) }, selfClosing: true });
          if (ffd.fieldType.listEntries) {
            for (const entry of ffd.fieldType.listEntries) {
              ddChildren.push({ name: 'w:listEntry', attributes: { 'w:val': entry }, selfClosing: true });
            }
          }
          ffDataChildren.push({ name: 'w:ddList', children: ddChildren });
          break;
        }
      }
    }

    return {
      name: 'w:fldChar',
      attributes: { 'w:fldCharType': 'begin' },
      children: [{ name: 'w:ffData', children: ffDataChildren }],
    };
  }
}

/**
 * TOC field options
 */
export interface TOCFieldOptions {
  /** Heading levels to include (e.g., "1-3" for levels 1-3) */
  levels?: string;

  /** Make entries hyperlinks (\h switch) */
  hyperlinks?: boolean;

  /** Hide tab leaders and page numbers in Web Layout (\z switch) */
  hideInWebLayout?: boolean;

  /** Use outline levels (\u switch) */
  useOutlineLevels?: boolean;

  /** Omit page numbers (\n switch) */
  omitPageNumbers?: boolean;

  /** Custom styles to use (\t switch) */
  customStyles?: string;
}

/**
 * Creates a TOC (Table of Contents) complex field
 * Generates proper field instruction with switches
 *
 * @param options TOC field options
 * @returns ComplexField configured for TOC
 *
 * @example
 * const toc = createTOCField({ levels: '1-3', hyperlinks: true });
 * // Generates: TOC \o "1-3" \h \z \u
 */
export function createTOCField(options: TOCFieldOptions = {}): ComplexField {
  // Build instruction string
  let instruction = ' TOC';

  // Add outline levels switch
  if (options.levels !== undefined) {
    instruction += ` \\o "${options.levels}"`;
  } else {
    instruction += ' \\o "1-3"'; // Default: levels 1-3
  }

  // Add hyperlinks switch
  if (options.hyperlinks !== false) {
    instruction += ' \\h';
  }

  // Add hide in web layout switch
  if (options.hideInWebLayout !== false) {
    instruction += ' \\z';
  }

  // Add use outline levels switch
  if (options.useOutlineLevels !== false) {
    instruction += ' \\u';
  }

  // Add omit page numbers switch
  if (options.omitPageNumbers) {
    instruction += ' \\n';
  }

  // Add custom styles switch
  if (options.customStyles) {
    instruction += ` \\t "${options.customStyles}"`;
  }

  instruction += ' '; // Trailing space per Microsoft convention

  return new ComplexField({
    instruction,
    result: 'Table of Contents', // Placeholder result
  });
}
