/**
 * Field - Represents a dynamic field in a Word document
 *
 * Fields are used for dynamic content like page numbers, dates, document properties, etc.
 * They are represented using the <w:fldSimple> element with field codes.
 */

import { XMLElement } from '../xml/XMLBuilder';
import { RunFormatting } from './Run';

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
  | 'SECTION';       // Current section number

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
   * Generates XML for the field
   * Uses fldSimple for simplicity
   */
  toXML(): XMLElement {
    const children: XMLElement[] = [];

    // Add run properties if formatting is specified
    if (this.formatting) {
      children.push(this.createRunProperties());
    }

    // Add field text (placeholder)
    children.push({
      name: 'w:t',
      children: [this.getPlaceholderText()],
    });

    return {
      name: 'w:fldSimple',
      attributes: {
        'w:instr': this.instruction,
      },
      children,
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

    if (this.formatting.bold) {
      children.push({ name: 'w:b', selfClosing: true });
    }

    if (this.formatting.italic) {
      children.push({ name: 'w:i', selfClosing: true });
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

    if (this.formatting.strike) {
      children.push({ name: 'w:strike', selfClosing: true });
    }

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

    if (this.formatting.size) {
      const sizeValue = (this.formatting.size * 2).toString();
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

    if (this.formatting.color) {
      const color = this.formatting.color.replace('#', '');
      children.push({
        name: 'w:color',
        attributes: { 'w:val': color },
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
  static createFilename(includePath: boolean = false, formatting?: RunFormatting): Field {
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

  /**
   * Creates a new complex field
   * @param properties Complex field properties
   */
  constructor(properties: ComplexFieldProperties) {
    this.instruction = properties.instruction;
    this.result = properties.result;
    this.instructionFormatting = properties.instructionFormatting;
    this.resultFormatting = properties.resultFormatting;
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
   * Generates XML for the complex field
   * Returns array of run elements (begin, instr, sep, result, end)
   */
  toXML(): XMLElement[] {
    const runs: XMLElement[] = [];

    // 1. Begin marker run
    runs.push({
      name: 'w:r',
      children: [
        {
          name: 'w:fldChar',
          attributes: { 'w:fldCharType': 'begin' },
          selfClosing: true,
        },
      ],
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

    // 4. Result run (optional)
    if (this.result) {
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
      const sizeValue = (formatting.size * 2).toString();
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
