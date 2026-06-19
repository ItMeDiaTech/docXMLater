/**
 * Revision - Represents tracked changes in a Word document
 *
 * Track changes allow tracking of insertions, deletions, and modifications
 * to document content, showing who made changes and when.
 */

import { Run } from './Run.js';
import type { RunFormatting } from './Run.js';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder.js';
import type { RevisionLocation } from './PropertyChangeTypes.js';
import type { RevisionContent } from './RevisionContent.js';
import { isRunContent, isHyperlinkContent } from './RevisionContent.js';
import { formatDateForXml } from '../utils/dateFormatting.js';

/**
 * Revision type - All OpenXML WordprocessingML revision types
 */
export type RevisionType =
  // Content changes
  | 'insert' // w:ins - Inserted content
  | 'delete' // w:del - Deleted content
  // Property changes
  | 'runPropertiesChange' // w:rPrChange - Run formatting change (bold, italic, font, etc.)
  | 'paragraphPropertiesChange' // w:pPrChange - Paragraph formatting change
  | 'tablePropertiesChange' // w:tblPrChange - Table formatting change
  | 'tableExceptionPropertiesChange' // w:tblPrExChange - Table exception properties change
  | 'tableRowPropertiesChange' // w:trPrChange - Table row properties change
  | 'tableCellPropertiesChange' // w:tcPrChange - Table cell properties change
  | 'sectionPropertiesChange' // w:sectPrChange - Section properties change
  // Move operations
  | 'moveFrom' // w:moveFrom - Content moved from this location
  | 'moveTo' // w:moveTo - Content moved to this location
  // Table operations
  | 'tableCellInsert' // w:cellIns - Table cell inserted
  | 'tableCellDelete' // w:cellDel - Table cell deleted
  | 'tableCellMerge' // w:cellMerge - Table cells merged
  // Numbering
  | 'numberingChange' // w:numberingChange - List numbering changed
  // Hyperlink changes
  | 'hyperlinkChange' // Hyperlink URL, text, or formatting change
  // Rich content changes (new tracking types)
  | 'imageChange' // Image insertion, deletion, or property change
  | 'fieldChange' // Field insertion, deletion, or value change
  | 'commentChange' // Comment insertion, deletion, or content change
  | 'bookmarkChange' // Bookmark creation, deletion, or range change
  | 'contentControlChange'; // Content control insertion, deletion, or property change

/**
 * Field context for revisions inside complex fields
 * Provides information about the parent field when a revision appears in a field result
 */
export interface FieldContext {
  /** Reference to the parent ComplexField (if revision is in field result) */
  field?: import('./Field.js').ComplexField;
  /** Field instruction (e.g., "HYPERLINK", "TOC", "MERGEFIELD") */
  instruction?: string;
  /** Position within field: 'instruction' or 'result' */
  position: 'instruction' | 'result';
}

/**
 * Revision properties
 */
export interface RevisionProperties {
  /** Unique revision ID (assigned by RevisionManager) */
  id?: number;
  /** Author who made the change */
  author: string;
  /** Date when the change was made */
  date?: Date;
  /** Type of revision */
  type: RevisionType;
  /** Content affected by the revision (Run, Hyperlink, or arrays thereof) */
  content: RevisionContent | RevisionContent[];
  /** Previous properties (for property change revisions) */
  previousProperties?: Record<string, any>;
  /** New properties (for property change revisions) */
  newProperties?: Record<string, any>;
  /** Move ID (for moveFrom/moveTo operations) */
  moveId?: string;
  /** Destination location (for moveFrom) or source location (for moveTo) */
  moveLocation?: string;
  /** Location of this revision within the document structure */
  location?: RevisionLocation;
  /** Field context if revision is inside a complex field */
  fieldContext?: FieldContext;
}

/**
 * Represents a tracked change (revision) in a document
 */
export class Revision {
  private id: number;
  private author: string;
  private date: Date;
  private type: RevisionType;
  private content: RevisionContent[];
  private previousProperties?: Record<string, any>;
  private newProperties?: Record<string, any>;
  private moveId?: string;
  private moveLocation?: string;
  private isFieldInstruction = false;
  private location?: RevisionLocation;
  private fieldContext?: FieldContext;

  /**
   * Creates a new Revision
   * @param properties - Revision properties
   */
  constructor(properties: RevisionProperties) {
    this.id = properties.id ?? 0; // Will be assigned by RevisionManager
    this.author = properties.author;
    this.date = properties.date || new Date();
    this.type = properties.type;
    this.content = Array.isArray(properties.content) ? properties.content : [properties.content];
    this.previousProperties = properties.previousProperties;
    this.newProperties = properties.newProperties;
    this.moveId = properties.moveId;
    this.moveLocation = properties.moveLocation;
    this.location = properties.location;
    this.fieldContext = properties.fieldContext;
  }

  /**
   * Gets the revision ID
   */
  getId(): number {
    return this.id;
  }

  /**
   * Sets the revision ID (used by RevisionManager)
   * @internal
   */
  setId(id: number): void {
    this.id = id;
  }

  /**
   * Gets the author
   */
  getAuthor(): string {
    return this.author;
  }

  /**
   * Sets the author
   */
  setAuthor(author: string): this {
    this.author = author;
    return this;
  }

  /**
   * Gets the revision date
   */
  getDate(): Date {
    return this.date;
  }

  /**
   * Sets the revision date
   */
  setDate(date: Date): this {
    this.date = date;
    return this;
  }

  /**
   * Gets the revision type
   */
  getType(): RevisionType {
    return this.type;
  }

  /**
   * Gets all content items in this revision
   * @returns Array of RevisionContent (Run or Hyperlink objects)
   */
  getContent(): RevisionContent[] {
    return [...this.content];
  }

  /**
   * Gets only the Run objects from this revision (backward compatible)
   * @returns Array of Run objects
   */
  getRuns(): Run[] {
    return this.content.filter((item): item is Run => isRunContent(item));
  }

  /**
   * Gets only the Hyperlink objects from this revision
   * @returns Array of Hyperlink objects
   */
  getHyperlinks(): import('./Hyperlink.js').Hyperlink[] {
    return this.content.filter((item): item is import('./Hyperlink.js').Hyperlink =>
      isHyperlinkContent(item)
    );
  }

  /**
   * Adds a run to this revision
   */
  addRun(run: Run): this {
    this.content.push(run);
    return this;
  }

  /**
   * Adds a hyperlink to this revision
   */
  addHyperlink(hyperlink: import('./Hyperlink.js').Hyperlink): this {
    this.content.push(hyperlink);
    return this;
  }

  /**
   * Adds content (Run or Hyperlink) to this revision
   */
  addContent(item: RevisionContent): this {
    this.content.push(item);
    return this;
  }

  /**
   * Gets the combined text content from all Runs and Hyperlinks in this revision.
   * This is used by isParagraphBlank() to detect text inside revision elements.
   * @returns Combined text string from all content items
   */
  getText(): string {
    return this.content
      .filter(
        (item): item is Run | import('./Hyperlink.js').Hyperlink =>
          isRunContent(item) || isHyperlinkContent(item)
      )
      .map((item) => item.getText())
      .join('');
  }

  /**
   * Gets the previous properties (for property change revisions)
   */
  getPreviousProperties(): Record<string, any> | undefined {
    return this.previousProperties;
  }

  /**
   * Gets the new properties (for property change revisions)
   */
  getNewProperties(): Record<string, any> | undefined {
    return this.newProperties;
  }

  /**
   * Gets the move ID (for moveFrom/moveTo operations)
   */
  getMoveId(): string | undefined {
    return this.moveId;
  }

  /**
   * Gets the move location
   */
  getMoveLocation(): string | undefined {
    return this.moveLocation;
  }

  /**
   * Gets the location of this revision within the document
   * @returns Location information or undefined if not set
   */
  getLocation(): RevisionLocation | undefined {
    return this.location;
  }

  /**
   * Sets the location of this revision within the document
   * @param location - Location information
   * @returns This revision for chaining
   */
  setLocation(location: RevisionLocation): this {
    this.location = location;
    return this;
  }

  /**
   * Gets the field context if this revision is inside a complex field
   * @returns Field context information or undefined if not inside a field
   */
  getFieldContext(): FieldContext | undefined {
    return this.fieldContext;
  }

  /**
   * Sets the field context for this revision
   * @param context - Field context information
   * @returns This revision for chaining
   */
  setFieldContext(context: FieldContext): this {
    this.fieldContext = context;
    return this;
  }

  /**
   * Checks if this revision is inside a complex field
   * @returns True if the revision is inside a field result or instruction section
   */
  isInsideField(): boolean {
    return this.fieldContext !== undefined;
  }

  /**
   * Checks if this revision is inside a field result section
   * @returns True if the revision is in the result section of a complex field
   */
  isInsideFieldResult(): boolean {
    return this.fieldContext?.position === 'result';
  }

  /**
   * Checks if this revision is inside a field instruction section
   * @returns True if the revision is in the instruction section of a complex field
   */
  isInsideFieldInstruction(): boolean {
    return this.fieldContext?.position === 'instruction';
  }

  /**
   * Gets the parent field if this revision is inside a complex field
   * @returns The parent ComplexField or undefined
   */
  getParentField(): import('./Field.js').ComplexField | undefined {
    return this.fieldContext?.field;
  }

  /**
   * Marks this revision as a field instruction deletion
   * When true, uses w:delInstrText instead of w:delText
   */
  setAsFieldInstruction(): this {
    this.isFieldInstruction = true;
    return this;
  }

  /**
   * Checks if this is a field instruction deletion
   */
  isFieldInstructionDeletion(): boolean {
    return this.isFieldInstruction;
  }

  /**
   * Formats a date to ISO 8601 format for XML
   * Per ECMA-376, revision dates must be in ISO 8601 format (e.g., "2024-01-01T12:00:00Z")
   * Uses formatDateForXml() to strip milliseconds which Word does not accept.
   * @param date - Date to format
   * @returns ISO 8601 formatted date string without milliseconds
   */
  private formatDate(date: Date): string {
    return formatDateForXml(date);
  }

  /**
   * Gets the XML element name for this revision type
   * Maps internal revision types to OOXML WordprocessingML element names per ECMA-376
   *
   * Mappings:
   * - insert → w:ins (inserted content)
   * - delete → w:del (deleted content)
   * - runPropertiesChange → w:rPrChange (run formatting change)
   * - paragraphPropertiesChange → w:pPrChange (paragraph formatting change)
   * - tablePropertiesChange → w:tblPrChange (table formatting change)
   * - tableRowPropertiesChange → w:trPrChange (table row properties change)
   * - tableCellPropertiesChange → w:tcPrChange (table cell properties change)
   * - sectionPropertiesChange → w:sectPrChange (section properties change)
   * - moveFrom → w:moveFrom (source location of moved content)
   * - moveTo → w:moveTo (destination location of moved content)
   * - tableCellInsert → w:cellIns (inserted table cell)
   * - tableCellDelete → w:cellDel (deleted table cell)
   * - tableCellMerge → w:cellMerge (merged table cells)
   * - numberingChange → w:numberingChange (list numbering changed)
   *
   * @returns OOXML element name (e.g., "w:ins", "w:del")
   */
  private getElementName(): string {
    switch (this.type) {
      case 'insert':
        return 'w:ins';
      case 'delete':
        return 'w:del';
      case 'runPropertiesChange':
        return 'w:rPrChange';
      case 'paragraphPropertiesChange':
        return 'w:pPrChange';
      case 'tablePropertiesChange':
        return 'w:tblPrChange';
      case 'tableExceptionPropertiesChange':
        return 'w:tblPrExChange';
      case 'tableRowPropertiesChange':
        return 'w:trPrChange';
      case 'tableCellPropertiesChange':
        return 'w:tcPrChange';
      case 'sectionPropertiesChange':
        return 'w:sectPrChange';
      case 'moveFrom':
        return 'w:moveFrom';
      case 'moveTo':
        return 'w:moveTo';
      case 'tableCellInsert':
        return 'w:cellIns';
      case 'tableCellDelete':
        return 'w:cellDel';
      case 'tableCellMerge':
        return 'w:cellMerge';
      case 'numberingChange':
        return 'w:numberingChange';
      // Internal tracking types - no OOXML element equivalent
      // These are used for changelog generation and internal tracking,
      // not for XML serialization. OOXML tracks these as insert/delete pairs.
      case 'hyperlinkChange':
      case 'imageChange':
      case 'fieldChange':
      case 'commentChange':
      case 'bookmarkChange':
      case 'contentControlChange':
        throw new Error(
          `Revision type '${this.type}' is an internal tracking type and cannot be serialized to OOXML XML. ` +
            `OOXML does not have a native element for this type. ` +
            `Use insert/delete revision pairs for tracking changes to ${this.type.replace('Change', '')}s.`
        );
      default:
        // TypeScript exhaustiveness check - this should never be reached
        // if all RevisionType values are handled above
        const _exhaustiveCheck: never = this.type;
        throw new Error(`Unknown revision type: ${_exhaustiveCheck}`);
    }
  }

  /**
   * Generates XML for this revision per OOXML WordprocessingML specification (ECMA-376)
   *
   * **XML Structure:**
   *
   * Content revisions (w:ins, w:del, w:moveFrom, w:moveTo):
   * ```xml
   * <w:ins w:id="0" w:author="Author Name" w:date="2024-01-01T12:00:00Z">
   *   <w:r>
   *     <w:t>Inserted text</w:t>
   *   </w:r>
   * </w:ins>
   * ```
   *
   * Deletion revisions use w:delText instead of w:t:
   * ```xml
   * <w:del w:id="1" w:author="Author Name" w:date="2024-01-01T12:00:00Z">
   *   <w:r>
   *     <w:delText>Deleted text</w:delText>
   *   </w:r>
   * </w:del>
   * ```
   *
   * Property change revisions (w:rPrChange, w:pPrChange, etc.):
   * ```xml
   * <w:rPrChange w:id="2" w:author="Author Name" w:date="2024-01-01T12:00:00Z">
   *   <w:rPr>
   *     <w:b/>  <!-- Previous bold setting -->
   *     <w:sz w:val="24"/>  <!-- Previous font size -->
   *   </w:rPr>
   * </w:rPrChange>
   * ```
   *
   * **Required Attributes (per ECMA-376):**
   * - w:id: Unique revision identifier (ST_DecimalNumber) - REQUIRED
   * - w:author: Author who made the change (ST_String) - REQUIRED
   * - w:date: When the change was made (ST_DateTime, ISO 8601) - OPTIONAL
   *
   * **Move Operations:**
   * w:moveFrom/w:moveTo are CT_RunTrackChange and carry only w:id/w:author/w:date.
   * The moveId field stays in-memory pairing metadata (getMovePair/validateMovePairs);
   * on disk, source/destination linkage is expressed solely via the w:name attribute
   * on w:moveFromRangeStart/w:moveToRangeStart (see RangeMarker.toXML()).
   *
   * **Content vs Property Changes:**
   * - Content revisions (insert/delete/move): Contain w:r elements with text runs
   * - Property revisions (rPrChange/pPrChange): Contain previous property elements (w:rPr, w:pPr)
   *
   * **Hyperlink Content:**
   * w:hyperlink is not a valid child of CT_RunTrackChange, so hyperlink content is
   * serialized with inverted nesting (<w:hyperlink><w:ins>...</w:ins></w:hyperlink>).
   * Mixed Run + Hyperlink content produces sibling elements (one revision element per
   * run group, one hyperlink-wrapped revision element per link), returned as a single
   * raw-XML fragment element that XMLBuilder emits verbatim.
   *
   * @returns XMLElement representing the revision in OOXML format, or null for internal-only types
   * @see ECMA-376 Part 1 §17.13.5 (Revision Identifiers for Paragraph Content)
   * @see ECMA-376 Part 1 §17.13.5.15 (Inserted Paragraph)
   * @see ECMA-376 Part 1 §17.13.5.14 (Deleted Paragraph)
   */
  toXML(): XMLElement | null {
    // Internal tracking types have no OOXML equivalent and cannot be serialized
    // They are used for changelog generation and internal tracking only
    const INTERNAL_TRACKING_TYPES: RevisionType[] = [
      'hyperlinkChange',
      'imageChange',
      'fieldChange',
      'commentChange',
      'bookmarkChange',
      'contentControlChange',
    ];

    if (INTERNAL_TRACKING_TYPES.includes(this.type)) {
      // Return null for internal types - callers should skip these
      return null;
    }

    const attributes: Record<string, string> = {
      'w:id': this.id.toString(),
      'w:author': this.author,
      'w:date': this.formatDate(this.date),
    };

    // moveId is intentionally NOT serialized: ECMA-376 CT_RunTrackChange declares no
    // such attribute, and an undeclared w: attribute fails Open XML schema validation.
    // On-disk move pairing is carried by w:name on the move range start markers.

    const elementName = this.getElementName();
    const children: XMLElement[] = [];

    // Handle different revision types
    if (this.isPropertyChangeType()) {
      // Property change revisions contain the previous properties
      if (this.previousProperties) {
        children.push(this.createPropertiesElement());
      }
    }

    // Per ECMA-376, w:hyperlink is NOT a valid child of w:ins/w:del (CT_RunTrackChange);
    // the nesting must be inverted so w:ins/w:del sits inside w:hyperlink. Content is
    // therefore split into sibling segments: each consecutive run group gets its own
    // revision element and each hyperlink wraps its own revision element, preserving
    // the link target (r:id/w:anchor) instead of downgrading it to plain runs.
    if (!this.isPropertyChangeType()) {
      const segments: XMLElement[] = [];
      // CT_TrackChange requires a document-unique w:id, so sibling segments
      // cannot share the revision's id. Extra segments derive ids from a high
      // base (deterministic, stable across saves) that cannot collide with the
      // small sequential ids RevisionManager issues.
      const segmentAttributes = (segmentIndex: number): Record<string, string> =>
        segmentIndex === 0
          ? attributes
          : {
              ...attributes,
              'w:id': String(500000000 + (this.id % 1000000) * 1000 + segmentIndex),
            };
      let pendingRuns: XMLElement[] = [];
      const flushRuns = (): void => {
        if (pendingRuns.length > 0) {
          segments.push({
            name: elementName,
            attributes: segmentAttributes(segments.length),
            children: pendingRuns,
          });
          pendingRuns = [];
        }
      };

      for (const item of this.content) {
        if (isHyperlinkContent(item)) {
          flushRuns();
          segments.push(
            this.createHyperlinkWrappedRevisionXml(
              item,
              elementName,
              segmentAttributes(segments.length)
            )
          );
        } else if (isRunContent(item)) {
          if (this.type === 'delete' || this.type === 'moveFrom') {
            pendingRuns.push(this.createDeletedRunXml(item));
          } else {
            pendingRuns.push(item.toXML());
          }
        }
      }
      flushRuns();

      if (segments.length === 0) {
        return { name: elementName, attributes, children: [] };
      }
      if (segments.length === 1) {
        return segments[0]!;
      }
      // Multiple sibling elements cannot share one XML root; emit them as a
      // pre-serialized fragment that XMLBuilder passes through verbatim.
      return {
        name: '__rawXml',
        rawXml: segments.map((segment) => XMLBuilder.elementToString(segment)).join(''),
      };
    }

    // Property change revisions keep their content runs alongside the
    // previous-properties element
    for (const item of this.content) {
      if (isHyperlinkContent(item)) {
        const hyperlinkXml = item.toXML();
        if (hyperlinkXml.children) {
          for (const child of hyperlinkXml.children) {
            if (typeof child === 'object' && child.name === 'w:r') {
              children.push(child);
            }
          }
        }
      } else if (isRunContent(item)) {
        children.push(item.toXML());
      }
    }

    return {
      name: elementName,
      attributes,
      children,
    };
  }

  /**
   * Checks if this is a property change revision type
   *
   * Property change revisions track formatting changes, not content changes.
   * They contain previous property elements (w:rPr, w:pPr, etc.) instead of text runs.
   *
   * **Property Change Types:**
   * - runPropertiesChange: Run formatting (bold, italic, font, color, etc.)
   * - paragraphPropertiesChange: Paragraph formatting (alignment, spacing, indentation, etc.)
   * - tablePropertiesChange: Table formatting
   * - tableRowPropertiesChange: Table row properties
   * - tableCellPropertiesChange: Table cell properties
   * - sectionPropertiesChange: Section properties (page size, margins, etc.)
   * - numberingChange: List numbering properties
   *
   * **Content Change Types (NOT property changes):**
   * - insert: Added text
   * - delete: Removed text
   * - moveFrom: Moved text source
   * - moveTo: Moved text destination
   * - tableCellInsert: Added table cell
   * - tableCellDelete: Removed table cell
   * - tableCellMerge: Merged table cells
   *
   * @returns true if this revision tracks a property/formatting change, false otherwise
   */
  private isPropertyChangeType(): boolean {
    return [
      'runPropertiesChange',
      'paragraphPropertiesChange',
      'tablePropertiesChange',
      'tableExceptionPropertiesChange',
      'tableRowPropertiesChange',
      'tableCellPropertiesChange',
      'sectionPropertiesChange',
      'numberingChange',
    ].includes(this.type);
  }

  /**
   * Creates XML element for previous properties in property change revisions
   *
   * **Purpose:**
   * Property change revisions (w:rPrChange, w:pPrChange, etc.) must contain a child element
   * with the PREVIOUS state of the properties before the change. This allows Word to show
   * what changed and enables accepting/rejecting the change.
   *
   * **Structure:**
   * ```xml
   * <w:rPrChange w:id="0" w:author="Author" w:date="...">
   *   <w:rPr>
   *     <!-- Previous run properties -->
   *     <w:b/>  <!-- Was bold -->
   *     <w:sz w:val="24"/>  <!-- Was 12pt (24 half-points) -->
   *   </w:rPr>
   * </w:rPrChange>
   * ```
   *
   * **Property Element Mapping:**
   * - runPropertiesChange → w:rPr (run properties)
   * - paragraphPropertiesChange → w:pPr (paragraph properties)
   * - tablePropertiesChange → w:tblPr (table properties)
   * - tableRowPropertiesChange → w:trPr (table row properties)
   * - tableCellPropertiesChange → w:tcPr (table cell properties)
   * - sectionPropertiesChange → w:sectPr (section properties)
   * - numberingChange → w:numPr (numbering properties)
   *
   * **Implementation:**
   * This method converts the previousProperties object into OOXML elements.
   * - runPropertiesChange delegates to Run.generateRunPropertiesXML
   * - Other types translate API-style keys to their schema local names
   *   (e.g., alignment → w:jc), serialize known object values (spacing,
   *   indentation, borders, shading, table widths), and order children
   *   per the schema sequence of the containing property element
   *
   * @returns XMLElement containing previous properties (w:rPr, w:pPr, etc.)
   * @see ECMA-376 Part 1 §17.13.5.31 (Run Properties Change)
   * @see ECMA-376 Part 1 §17.13.5.29 (Paragraph Properties Change)
   */
  private createPropertiesElement(): XMLElement {
    // For runPropertiesChange, delegate to Run.generateRunPropertiesXML for correct
    // ECMA-376 element names and ordering (e.g., bold→w:b, font→w:rFonts, size→w:sz)
    if (this.type === 'runPropertiesChange' && this.previousProperties) {
      const rPr = Run.generateRunPropertiesXML(this.previousProperties as RunFormatting);
      return rPr || { name: 'w:rPr', attributes: {}, children: [] };
    }

    // The property element name depends on the revision type
    let propElementName = 'w:rPr';

    switch (this.type) {
      case 'runPropertiesChange':
        propElementName = 'w:rPr';
        break;
      case 'paragraphPropertiesChange':
        propElementName = 'w:pPr';
        break;
      case 'tablePropertiesChange':
        propElementName = 'w:tblPr';
        break;
      case 'tableExceptionPropertiesChange':
        propElementName = 'w:tblPrEx';
        break;
      case 'tableRowPropertiesChange':
        propElementName = 'w:trPr';
        break;
      case 'tableCellPropertiesChange':
        propElementName = 'w:tcPr';
        break;
      case 'sectionPropertiesChange':
        propElementName = 'w:sectPr';
        break;
      case 'numberingChange':
        propElementName = 'w:numPr';
        break;
    }

    // Build property children from previousProperties, translating API-style
    // keys to their OOXML local names (keys already given as local names pass
    // through unchanged)
    const keyMap = Revision.PROPERTY_KEY_TO_ELEMENT[this.type] ?? {};
    const propChildren: XMLElement[] = [];
    if (this.previousProperties) {
      for (const [key, value] of Object.entries(this.previousProperties)) {
        const child = this.createPreviousPropertyXml(`w:${keyMap[key] ?? key}`, value);
        if (child) {
          propChildren.push(child);
        }
      }
    }

    // CT_PPrBase, CT_TblPrBase, CT_TrPr, CT_TcPr and CT_SectPr are xsd:sequence
    // types, so children must follow the declared order to stay schema-valid
    const order = Revision.PROPERTY_CHILD_ORDER[propElementName];
    if (order) {
      const rank = (name: string): number => {
        const index = order.indexOf(name);
        return index === -1 ? order.length : index;
      };
      propChildren.sort((a, b) => rank(a.name) - rank(b.name));
    }

    return {
      name: propElementName,
      attributes: {},
      children: propChildren,
    };
  }

  /**
   * API-key → OOXML local-name translations for property-change snapshots.
   * previousProperties accepts the same key names as the element formatting
   * APIs (e.g. ParagraphFormatting.alignment); the schema requires the local
   * element names (w:jc), so keys are translated before serialization.
   */
  private static readonly PROPERTY_KEY_TO_ELEMENT: Partial<
    Record<RevisionType, Record<string, string>>
  > = {
    paragraphPropertiesChange: {
      alignment: 'jc',
      style: 'pStyle',
      styleId: 'pStyle',
      indentation: 'ind',
      numbering: 'numPr',
      outlineLevel: 'outlineLvl',
      shading: 'shd',
      borders: 'pBdr',
    },
    tablePropertiesChange: {
      alignment: 'jc',
      style: 'tblStyle',
      styleId: 'tblStyle',
      width: 'tblW',
      indent: 'tblInd',
      borders: 'tblBorders',
      shading: 'shd',
      layout: 'tblLayout',
      cellMargins: 'tblCellMar',
      cellSpacing: 'tblCellSpacing',
    },
    tableExceptionPropertiesChange: {
      alignment: 'jc',
      width: 'tblW',
      indent: 'tblInd',
      borders: 'tblBorders',
      shading: 'shd',
      layout: 'tblLayout',
      cellMargins: 'tblCellMar',
      cellSpacing: 'tblCellSpacing',
    },
    tableRowPropertiesChange: {
      alignment: 'jc',
      height: 'trHeight',
      isHeader: 'tblHeader',
      cellSpacing: 'tblCellSpacing',
    },
    tableCellPropertiesChange: {
      width: 'tcW',
      borders: 'tcBorders',
      shading: 'shd',
      verticalAlignment: 'vAlign',
      margins: 'tcMar',
    },
    sectionPropertiesChange: {
      pageSize: 'pgSz',
      margins: 'pgMar',
      columns: 'cols',
      pageNumbering: 'pgNumType',
    },
  };

  /**
   * Schema child order for property snapshot containers (xsd:sequence order
   * from the corresponding CT_* types). Unknown children sort after known ones.
   */
  private static readonly PROPERTY_CHILD_ORDER: Record<string, readonly string[]> = {
    'w:pPr': [
      'w:pStyle',
      'w:keepNext',
      'w:keepLines',
      'w:pageBreakBefore',
      'w:framePr',
      'w:widowControl',
      'w:numPr',
      'w:suppressLineNumbers',
      'w:pBdr',
      'w:shd',
      'w:tabs',
      'w:suppressAutoHyphens',
      'w:kinsoku',
      'w:wordWrap',
      'w:overflowPunct',
      'w:topLinePunct',
      'w:autoSpaceDE',
      'w:autoSpaceDN',
      'w:bidi',
      'w:adjustRightInd',
      'w:snapToGrid',
      'w:spacing',
      'w:ind',
      'w:contextualSpacing',
      'w:mirrorIndents',
      'w:suppressOverlap',
      'w:jc',
      'w:textDirection',
      'w:textAlignment',
      'w:textboxTightWrap',
      'w:outlineLvl',
    ],
    'w:tblPr': [
      'w:tblStyle',
      'w:tblpPr',
      'w:tblOverlap',
      'w:bidiVisual',
      'w:tblStyleRowBandSize',
      'w:tblStyleColBandSize',
      'w:tblW',
      'w:jc',
      'w:tblCellSpacing',
      'w:tblInd',
      'w:tblBorders',
      'w:shd',
      'w:tblLayout',
      'w:tblCellMar',
      'w:tblLook',
      'w:tblCaption',
      'w:tblDescription',
    ],
    'w:tblPrEx': [
      'w:tblW',
      'w:jc',
      'w:tblCellSpacing',
      'w:tblInd',
      'w:tblBorders',
      'w:shd',
      'w:tblLayout',
      'w:tblCellMar',
      'w:tblLook',
    ],
    'w:trPr': [
      'w:cnfStyle',
      'w:divId',
      'w:gridBefore',
      'w:gridAfter',
      'w:wBefore',
      'w:wAfter',
      'w:cantSplit',
      'w:trHeight',
      'w:tblHeader',
      'w:tblCellSpacing',
      'w:jc',
      'w:hidden',
    ],
    'w:tcPr': [
      'w:cnfStyle',
      'w:tcW',
      'w:gridSpan',
      'w:hMerge',
      'w:vMerge',
      'w:tcBorders',
      'w:shd',
      'w:noWrap',
      'w:tcMar',
      'w:textDirection',
      'w:tcFitText',
      'w:vAlign',
      'w:hideMark',
    ],
    'w:sectPr': [
      'w:headerReference',
      'w:footerReference',
      'w:footnotePr',
      'w:endnotePr',
      'w:type',
      'w:pgSz',
      'w:pgMar',
      'w:paperSrc',
      'w:pgBorders',
      'w:lnNumType',
      'w:pgNumType',
      'w:cols',
      'w:formProt',
      'w:vAlign',
      'w:noEndnote',
      'w:titlePg',
      'w:textDirection',
      'w:bidi',
      'w:rtlGutter',
      'w:docGrid',
    ],
    'w:numPr': ['w:ilvl', 'w:numId', 'w:numberingChange', 'w:ins'],
  };

  /**
   * Serializes one previous-property entry to its OOXML element.
   * Booleans follow the CT_OnOff convention (true = on, false = explicit off
   * via w:val="0") so the snapshot keeps the same tri-state information the
   * main property serializers emit.
   */
  private createPreviousPropertyXml(name: string, value: unknown): XMLElement | null {
    if (typeof value === 'boolean') {
      return { name, attributes: value ? {} : { 'w:val': '0' }, children: [] };
    }
    if (typeof value === 'string' || typeof value === 'number') {
      return { name, attributes: { 'w:val': value.toString() }, children: [] };
    }
    if (value && typeof value === 'object') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any -- previousProperties is Record<string, any>
      return this.createObjectPropertyXml(name, value as Record<string, any>);
    }
    return null;
  }

  /**
   * Serializes object-valued previous properties (complex OOXML structures).
   * Only shapes with a known schema mapping are emitted; unknown objects are
   * skipped because their attribute layout cannot be inferred safely.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- previousProperties is Record<string, any>
  private createObjectPropertyXml(name: string, value: Record<string, any>): XMLElement | null {
    switch (name) {
      case 'w:spacing':
        return {
          name,
          attributes: XMLBuilder.buildAttributes({
            'w:before': value.before,
            'w:beforeLines': value.beforeLines,
            'w:beforeAutospacing':
              value.beforeAutospacing === undefined
                ? undefined
                : value.beforeAutospacing
                  ? '1'
                  : '0',
            'w:after': value.after,
            'w:afterLines': value.afterLines,
            'w:afterAutospacing':
              value.afterAutospacing === undefined ? undefined : value.afterAutospacing ? '1' : '0',
            'w:line': value.line,
            'w:lineRule': value.lineRule,
          }),
          children: [],
        };
      case 'w:ind':
        return {
          name,
          attributes: XMLBuilder.buildAttributes({
            'w:start': value.start,
            'w:end': value.end,
            'w:left': value.left,
            'w:leftChars': value.leftChars,
            'w:right': value.right,
            'w:rightChars': value.rightChars,
            'w:firstLine': value.firstLine,
            'w:firstLineChars': value.firstLineChars,
            'w:hanging': value.hanging,
            'w:hangingChars': value.hangingChars,
          }),
          children: [],
        };
      case 'w:numPr': {
        const numPrChildren: XMLElement[] = [];
        const level = value.level ?? value.ilvl;
        const numId = value.numId ?? value.id;
        if (level !== undefined) {
          numPrChildren.push({
            name: 'w:ilvl',
            attributes: { 'w:val': String(level) },
            children: [],
          });
        }
        if (numId !== undefined) {
          numPrChildren.push({
            name: 'w:numId',
            attributes: { 'w:val': String(numId) },
            children: [],
          });
        }
        return numPrChildren.length > 0 ? { name, attributes: {}, children: numPrChildren } : null;
      }
      case 'w:shd':
        return XMLBuilder.createShading(value);
      case 'w:pBdr':
      case 'w:tblBorders':
      case 'w:tcBorders': {
        const sides =
          name === 'w:pBdr'
            ? ['top', 'left', 'bottom', 'right', 'between', 'bar']
            : ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
        const borderChildren: XMLElement[] = [];
        for (const side of sides) {
          if (value[side]) {
            borderChildren.push(XMLBuilder.createBorder(side, value[side]));
          }
        }
        return borderChildren.length > 0
          ? { name, attributes: {}, children: borderChildren }
          : null;
      }
      case 'w:tblW':
      case 'w:tcW':
      case 'w:tblInd':
      case 'w:tblCellSpacing':
        return {
          name,
          attributes: XMLBuilder.buildAttributes({
            'w:w': value.w ?? value.width ?? value.value,
            'w:type': value.type ?? 'dxa',
          }),
          children: [],
        };
      case 'w:trHeight':
        return {
          name,
          attributes: XMLBuilder.buildAttributes({
            'w:val': value.value ?? value.val ?? value.height,
            'w:hRule': value.rule ?? value.hRule,
          }),
          children: [],
        };
      case 'w:tcMar':
      case 'w:tblCellMar':
        return XMLBuilder.createMargins(name.slice(2), value);
      case 'w:pgSz':
        return {
          name,
          attributes: XMLBuilder.buildAttributes({
            'w:w': value.width ?? value.w,
            'w:h': value.height ?? value.h,
            'w:orient': value.orientation ?? value.orient,
          }),
          children: [],
        };
      case 'w:pgMar':
        return {
          name,
          attributes: XMLBuilder.buildAttributes({
            'w:top': value.top,
            'w:right': value.right,
            'w:bottom': value.bottom,
            'w:left': value.left,
            'w:header': value.header,
            'w:footer': value.footer,
            'w:gutter': value.gutter,
          }),
          children: [],
        };
      default:
        // Unknown object shape — no schema mapping to apply
        return null;
    }
  }

  /**
   * Creates XML for a deleted run (uses w:delText or w:delInstrText instead of w:t)
   *
   * **OOXML Requirement:**
   * Per ECMA-376, deleted text must use w:delText element instead of w:t element.
   * For deleted field instructions, w:delInstrText must be used instead.
   * This is required for proper rendering in Microsoft Word's Track Changes mode.
   *
   * **Transformation:**
   * ```xml
   * <!-- Normal run (NOT in deletion) -->
   * <w:r>
   *   <w:rPr><w:b/></w:rPr>
   *   <w:t>Text</w:t>
   * </w:r>
   *
   * <!-- Deleted run (inside w:del) -->
   * <w:r>
   *   <w:rPr><w:b/></w:rPr>
   *   <w:delText>Text</w:delText>
   * </w:r>
   *
   * <!-- Deleted field instruction (inside w:del) -->
   * <w:r>
   *   <w:delInstrText>MERGEFIELD Name</w:delInstrText>
   * </w:r>
   * ```
   *
   * **Why This Matters:**
   * - w:delText tells Word to render with strikethrough in Track Changes mode
   * - w:delInstrText is specifically for deleted field codes
   * - w:t would render as normal text even inside w:del element
   * - Word will reject documents with w:t inside deletions as malformed
   *
   * **Implementation:**
   * This method gets the run's normal XML and replaces all w:t elements with w:delText
   * or w:delInstrText (for field instructions) while preserving all other properties
   * (formatting, spacing attributes, etc.)
   *
   * @param run - Run containing deleted text
   * @returns XMLElement with w:delText or w:delInstrText instead of w:t
   * @see ECMA-376 Part 1 §17.13.5.14 (Deleted Run Content)
   * @see ECMA-376 Part 1 §22.1.2.27 (w:delText element)
   * @see ECMA-376 Part 1 §22.1.2.26 (w:delInstrText element)
   */
  private createDeletedRunXml(run: Run): XMLElement {
    // Get the regular run XML
    const runXml = run.toXML();

    // Determine which element to use for deleted text
    // w:delInstrText for field instructions, w:delText for regular text
    const deletedTextElement = this.isFieldInstruction ? 'w:delInstrText' : 'w:delText';

    // We need to replace text elements with their deleted counterparts:
    // - w:t -> w:delText (or w:delInstrText if isFieldInstruction)
    // - w:instrText -> w:delInstrText (always, regardless of isFieldInstruction flag)
    if (runXml.children) {
      const modifiedChildren = runXml.children.map((child) => {
        if (typeof child === 'object') {
          if (child.name === 'w:t') {
            // Replace w:t with appropriate deleted text element
            return {
              ...child,
              name: deletedTextElement,
            };
          } else if (child.name === 'w:instrText') {
            // Replace w:instrText with w:delInstrText
            // Per ECMA-376 §22.1.2.26, deleted field instructions must use w:delInstrText
            return {
              ...child,
              name: 'w:delInstrText',
            };
          }
        }
        return child;
      });

      return {
        ...runXml,
        children: modifiedChildren,
      };
    }

    return runXml;
  }

  /**
   * Creates XML for a deleted hyperlink (transforms nested runs to use w:delText)
   *
   * **OOXML Requirement:**
   * Per ECMA-376, when a hyperlink is inside a w:del element, its nested runs must
   * use w:delText instead of w:t. This transforms the hyperlink's internal run
   * content to comply with Word's Track Changes requirements.
   *
   * **XML Structure:**
   * ```xml
   * <w:del w:id="1" w:author="Author" w:date="...">
   *   <w:hyperlink r:id="rId1">
   *     <w:r>
   *       <w:delText>Link text</w:delText>  <!-- Transformed from w:t -->
   *     </w:r>
   *   </w:hyperlink>
   * </w:del>
   * ```
   *
   * @param hyperlink - Hyperlink containing deleted content
   * @returns XMLElement with nested runs transformed to use w:delText
   */
  /**
   * Creates a hyperlink-wrapped revision XML element.
   * Per ECMA-376, w:hyperlink is NOT a valid child of w:ins/w:del.
   * Instead, w:ins/w:del must be inside w:hyperlink:
   * <w:hyperlink r:id="rId1"><w:ins ...><w:r>...</w:r></w:ins></w:hyperlink>
   */
  private createHyperlinkWrappedRevisionXml(
    hyperlink: import('./Hyperlink.js').Hyperlink,
    revisionElementName: string,
    revisionAttributes: Record<string, string>
  ): XMLElement {
    const hyperlinkXml = hyperlink.toXML();

    // Extract runs from the hyperlink
    const runs: XMLElement[] = [];
    if (hyperlinkXml.children) {
      for (const child of hyperlinkXml.children) {
        if (typeof child === 'object' && child.name === 'w:r') {
          if (this.type === 'delete' || this.type === 'moveFrom') {
            runs.push(this.convertRunXmlToDeleted(child));
          } else {
            runs.push(child);
          }
        }
      }
    }

    // Build: <w:hyperlink ...><w:ins/w:del ...><w:r>...</w:r></w:ins/w:del></w:hyperlink>
    const revisionElement: XMLElement = {
      name: revisionElementName,
      attributes: revisionAttributes,
      children: runs,
    };

    return {
      name: hyperlinkXml.name,
      attributes: hyperlinkXml.attributes,
      children: [revisionElement],
    };
  }

  /**
   * Converts a run XMLElement to use deleted text elements
   */
  private convertRunXmlToDeleted(runXml: XMLElement): XMLElement {
    const deletedTextElement = this.isFieldInstruction ? 'w:delInstrText' : 'w:delText';

    if (!runXml.children) return runXml;

    const modifiedChildren = runXml.children.map((child) => {
      if (typeof child === 'object' && child.name === 'w:t') {
        return {
          ...child,
          name: deletedTextElement,
        };
      }
      return child;
    });

    return {
      ...runXml,
      children: modifiedChildren,
    };
  }

  /**
   * Creates an insertion revision
   * @param author - Author who made the insertion
   * @param content - Inserted content (Run, Hyperlink, or arrays thereof)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createInsertion(
    author: string,
    content: RevisionContent | RevisionContent[],
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'insert',
      content,
      date,
    });
  }

  /**
   * Creates a deletion revision
   * @param author - Author who made the deletion
   * @param content - Deleted content (Run, Hyperlink, or arrays thereof)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createDeletion(
    author: string,
    content: RevisionContent | RevisionContent[],
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'delete',
      content,
      date,
    });
  }

  /**
   * Creates a field instruction deletion revision
   * Uses w:delInstrText instead of w:delText for field codes
   * @param author - Author who made the deletion
   * @param content - Deleted field instruction content (Run or array of Runs)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createFieldInstructionDeletion(
    author: string,
    content: Run | Run[],
    date?: Date
  ): Revision {
    const revision = new Revision({
      author,
      type: 'delete',
      content,
      date,
    });
    revision.setAsFieldInstruction();
    return revision;
  }

  /**
   * Creates a revision from text
   * Convenience method that creates a Run from the text
   * @param type - Revision type
   * @param author - Author who made the change
   * @param text - Text content
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static fromText(type: RevisionType, author: string, text: string, date?: Date): Revision {
    const run = new Run(text);
    return new Revision({
      author,
      type,
      content: run,
      date,
    });
  }

  /**
   * Creates a run properties change revision
   * @param author - Author who made the change
   * @param content - Content with changed formatting
   * @param previousProperties - Previous run properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createRunPropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'runPropertiesChange',
      content,
      previousProperties,
      date,
    });
  }

  /**
   * Creates a paragraph properties change revision
   * @param author - Author who made the change
   * @param content - Paragraph content
   * @param previousProperties - Previous paragraph properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createParagraphPropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'paragraphPropertiesChange',
      content,
      previousProperties,
      date,
    });
  }

  /**
   * Creates a table properties change revision
   * @param author - Author who made the change
   * @param content - Table content
   * @param previousProperties - Previous table properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTablePropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'tablePropertiesChange',
      content,
      previousProperties,
      date,
    });
  }

  /**
   * Creates a table exception properties change revision
   * Tracks changes to table properties that override style defaults
   * @param author - Author who made the change
   * @param content - Table content
   * @param previousProperties - Previous table exception properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTableExceptionPropertiesChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'tableExceptionPropertiesChange',
      content,
      previousProperties,
      date,
    });
  }

  /**
   * Creates a moveFrom revision (source of moved content)
   * @param author - Author who moved the content
   * @param content - Content that was moved
   * @param moveId - Unique move operation ID (links moveFrom and moveTo)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createMoveFrom(
    author: string,
    content: Run | Run[],
    moveId: string,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'moveFrom',
      content,
      moveId,
      date,
    });
  }

  /**
   * Creates a moveTo revision (destination of moved content)
   * @param author - Author who moved the content
   * @param content - Content that was moved
   * @param moveId - Unique move operation ID (links moveFrom and moveTo)
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createMoveTo(author: string, content: Run | Run[], moveId: string, date?: Date): Revision {
    return new Revision({
      author,
      type: 'moveTo',
      content,
      moveId,
      date,
    });
  }

  /**
   * Creates a table cell insertion revision
   * @param author - Author who inserted the cell
   * @param content - Cell content
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTableCellInsert(author: string, content: Run | Run[], date?: Date): Revision {
    return new Revision({
      author,
      type: 'tableCellInsert',
      content,
      date,
    });
  }

  /**
   * Creates a table cell deletion revision
   * @param author - Author who deleted the cell
   * @param content - Cell content
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTableCellDelete(author: string, content: Run | Run[], date?: Date): Revision {
    return new Revision({
      author,
      type: 'tableCellDelete',
      content,
      date,
    });
  }

  /**
   * Creates a table cell merge revision
   * @param author - Author who merged cells
   * @param content - Merged cell content
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createTableCellMerge(author: string, content: Run | Run[], date?: Date): Revision {
    return new Revision({
      author,
      type: 'tableCellMerge',
      content,
      date,
    });
  }

  /**
   * Creates a numbering change revision
   * @param author - Author who changed the numbering
   * @param content - Content with changed numbering
   * @param previousProperties - Previous numbering properties
   * @param date - Optional date (defaults to now)
   * @returns New Revision instance
   */
  static createNumberingChange(
    author: string,
    content: Run | Run[],
    previousProperties: Record<string, any>,
    date?: Date
  ): Revision {
    return new Revision({
      author,
      type: 'numberingChange',
      content,
      previousProperties,
      date,
    });
  }
}
