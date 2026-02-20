/**
 * Image - Represents an embedded image in a Word document
 *
 * Images use DrawingML (a:) and WordprocessingML Drawing (wp:) namespaces
 * for proper positioning and formatting in Word documents.
 */

import { promises as fs } from 'fs';
import { defaultLogger } from '../utils/logger';
import { inchesToEmus, UNITS } from '../utils/units';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';

/**
 * Supported image formats
 */
export type ImageFormat = 'png' | 'jpeg' | 'jpg' | 'gif' | 'bmp' | 'tiff' | 'svg' | 'emf' | 'wmf';

/**
 * Preset geometry shape type (ECMA-376 §20.1.9.18)
 */
export type PresetGeometry = 'rect' | 'roundRect' | 'ellipse' | string;

/**
 * Blip compression state (ECMA-376 §20.1.8.15)
 */
export type BlipCompressionState = 'none' | 'print' | 'email' | 'hqprint' | 'screen';

/**
 * Picture lock attribute names (ECMA-376 §20.1.2.2.31)
 */
export type PicLockAttribute = 'noChangeAspect' | 'noChangeArrowheads' | 'noSelect' | 'noMove'
  | 'noResize' | 'noEditPoints' | 'noAdjustHandles' | 'noRot' | 'noChangeShapeType'
  | 'noCrop' | 'noGrp';

/**
 * Non-visual picture properties (ECMA-376 §19.3.1.12)
 */
export interface PicNonVisualProperties {
  id: string;
  name: string;
  descr: string;
}

/**
 * Image border definition (full a:ln support per ECMA-376)
 */
export interface ImageBorder {
  /** Border width in points */
  width: number;
  /** Line cap style */
  cap?: 'flat' | 'rnd' | 'sq';
  /** Compound line type */
  compound?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
  /** Alignment relative to shape */
  alignment?: 'ctr' | 'in';
  /** Fill specification */
  fill?: {
    type: 'srgbClr' | 'schemeClr';
    value: string;
    modifiers?: Array<{ name: string; val: string }>;
  };
  /** Raw XML for non-solid fills (gradFill, pattFill, etc.) */
  rawFillXml?: string;
  /** Preset dash pattern */
  dashPattern?: string;
  /** Line join style */
  join?: 'round' | 'bevel' | 'miter';
  /** Miter limit (percentage * 1000) */
  miterLimit?: number;
  /** Head end decoration */
  headEnd?: { type?: string; width?: string; length?: string };
  /** Tail end decoration */
  tailEnd?: { type?: string; width?: string; length?: string };
}

/**
 * Image extent (dimensions)
 */
export interface ImageExtent {
  /** Width in EMUs */
  width: number;
  /** Height in EMUs */
  height: number;
}

/**
 * Effect extent (additional space for shadows, reflections, glows)
 * Specifies additional space to add to each edge to prevent clipping of effects
 */
export interface EffectExtent {
  /** Left extent in EMUs */
  left: number;
  /** Top extent in EMUs */
  top: number;
  /** Right extent in EMUs */
  right: number;
  /** Bottom extent in EMUs */
  bottom: number;
}

/**
 * Text wrapping type
 */
export type WrapType = 'square' | 'tight' | 'through' | 'topAndBottom' | 'none';

/**
 * Text wrapping side
 */
export type WrapSide = 'bothSides' | 'left' | 'right' | 'largest';

/**
 * Text wrap settings
 */
export interface TextWrapSettings {
  /** Wrap type */
  type: WrapType;
  /** Which side(s) to wrap text */
  side?: WrapSide;
  /** Distance from top in EMUs */
  distanceTop?: number;
  /** Distance from bottom in EMUs */
  distanceBottom?: number;
  /** Distance from left in EMUs */
  distanceLeft?: number;
  /** Distance from right in EMUs */
  distanceRight?: number;
}

/**
 * Position anchor type (what to position relative to)
 */
export type PositionAnchor = 'page' | 'margin' | 'column' | 'character' | 'paragraph';

/**
 * Horizontal alignment options
 */
export type HorizontalAlignment = 'left' | 'center' | 'right' | 'inside' | 'outside';

/**
 * Vertical alignment options
 */
export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'inside' | 'outside';

/**
 * Image position configuration
 */
export interface ImagePosition {
  /** Horizontal positioning */
  horizontal: {
    /** Anchor point */
    anchor: PositionAnchor;
    /** Offset in EMUs (absolute positioning) */
    offset?: number;
    /** Alignment (relative positioning) */
    alignment?: HorizontalAlignment;
  };
  /** Vertical positioning */
  vertical: {
    /** Anchor point */
    anchor: PositionAnchor;
    /** Offset in EMUs (absolute positioning) */
    offset?: number;
    /** Alignment (relative positioning) */
    alignment?: VerticalAlignment;
  };
}

/**
 * Image anchor configuration (floating images)
 */
export interface ImageAnchor {
  /** Position behind text */
  behindDoc: boolean;
  /** Lock anchor (prevent movement) */
  locked: boolean;
  /** Layout in table cell */
  layoutInCell: boolean;
  /** Allow overlap with other objects */
  allowOverlap: boolean;
  /** Z-order (higher = in front) */
  relativeHeight: number;
  /** Use simple positioning (wp:simplePos coordinates) */
  simplePos?: boolean;
  /** Distance from text - top (EMUs) */
  distT?: number;
  /** Distance from text - bottom (EMUs) */
  distB?: number;
  /** Distance from text - left (EMUs) */
  distL?: number;
  /** Distance from text - right (EMUs) */
  distR?: number;
}

/**
 * Image crop settings (percentage-based)
 */
export interface ImageCrop {
  /** Left crop percentage (0-100) */
  left: number;
  /** Top crop percentage (0-100) */
  top: number;
  /** Right crop percentage (0-100) */
  right: number;
  /** Bottom crop percentage (0-100) */
  bottom: number;
}

/**
 * Image visual effects
 */
export interface ImageEffects {
  /** Brightness adjustment (-100 to +100) */
  brightness?: number;
  /** Contrast adjustment (-100 to +100) */
  contrast?: number;
  /** Convert to grayscale */
  grayscale?: boolean;
  /** Transparency (0-100, percentage) via a:alphaModFix */
  transparency?: number;
}

/**
 * Image properties
 */
export interface ImageProperties {
  /** Image source (file path or buffer) */
  source: string | Buffer;
  /** Image width in EMUs (optional - will auto-detect) */
  width?: number;
  /** Image height in EMUs (optional - will auto-detect) */
  height?: number;
  /** Maintain aspect ratio when resizing */
  maintainAspectRatio?: boolean;
  /** Alt text / description */
  description?: string;
  /** Image name/title */
  name?: string;
  /** Image title (wp:docPr title attribute for accessibility) */
  title?: string;
  /** Relationship ID (will be set by ImageManager) */
  relationshipId?: string;
  /** Effect extent (space for shadows/glows) */
  effectExtent?: EffectExtent;
  /** Text wrapping configuration */
  wrap?: TextWrapSettings;
  /** Position configuration (floating images) */
  position?: ImagePosition;
  /** Anchor configuration (floating images) */
  anchor?: ImageAnchor;
  /** Crop settings */
  crop?: ImageCrop;
  /** Visual effects */
  effects?: ImageEffects;
  /** Border settings */
  border?: ImageBorder | { width: number };
  /** Rotation angle in degrees (0-360) */
  rotation?: number;
  /** Horizontal flip (ECMA-376 §20.1.7.6) */
  flipH?: boolean;
  /** Vertical flip (ECMA-376 §20.1.7.6) */
  flipV?: boolean;
  /** Preset geometry shape (ECMA-376 §20.1.9.18) */
  presetGeometry?: PresetGeometry;
  /** Blip compression state (ECMA-376 §20.1.8.15) */
  compressionState?: BlipCompressionState;
  /** Black-and-white mode (ECMA-376 §20.1.2.2.35) */
  bwMode?: string;
  /** Inline distance from text - top (EMUs, ECMA-376 §20.4.2.8) */
  inlineDistT?: number;
  /** Inline distance from text - bottom (EMUs) */
  inlineDistB?: number;
  /** Inline distance from text - left (EMUs) */
  inlineDistL?: number;
  /** Inline distance from text - right (EMUs) */
  inlineDistR?: number;
  /** Whether aspect ratio lock is enabled (ECMA-376 §20.4.2.4) */
  noChangeAspect?: boolean;
  /** Hidden attribute on docPr (ECMA-376 §20.4.2.3) */
  hidden?: boolean;
  /** BlipFill DPI override (ECMA-376 §20.1.8.14) */
  blipFillDpi?: number;
  /** BlipFill rotate with shape flag (ECMA-376 §20.1.8.14) */
  blipFillRotWithShape?: boolean;
  /** Picture locks (ECMA-376 §20.1.2.2.31) */
  picLocks?: Partial<Record<PicLockAttribute, boolean>>;
  /** Non-visual picture properties (ECMA-376 §19.3.1.12) */
  picNonVisualProps?: PicNonVisualProperties;
  /** Whether image is linked (r:link) vs embedded (r:embed) */
  isLinked?: boolean;
  /** SVG relationship ID for Word 365 dual-relationship approach */
  svgRelationshipId?: string;
}

/**
 * Image validation result
 */
export interface ValidationResult {
  valid: boolean;
  error?: string;
}

export class Image {
  private source: string | Buffer;
  private width: number;
  private height: number;
  private description: string;
  private name: string;
  private title?: string;
  private relationshipId?: string;
  private imageData?: Buffer;
  private extension: string;
  private docPrId: number = 1;
  private dpi: number = 96;  // Default DPI

  // Advanced image properties
  private effectExtent?: EffectExtent;
  private wrap?: TextWrapSettings;
  private position?: ImagePosition;
  private anchor?: ImageAnchor;
  private crop?: ImageCrop;
  private effects?: ImageEffects;
  private rotation: number = 0;
  private flipH: boolean = false;
  private flipV: boolean = false;
  private border?: ImageBorder;

  // Group A: Simple attribute preservation (ECMA-376 compliance)
  private presetGeometry: PresetGeometry = 'rect';
  private compressionState: BlipCompressionState = 'none';
  private bwMode: string = 'auto';
  private inlineDistT: number = 0;
  private inlineDistB: number = 0;
  private inlineDistL: number = 0;
  private inlineDistR: number = 0;
  private noChangeAspect: boolean = true;
  private hidden: boolean = false;
  private blipFillDpi?: number;
  private blipFillRotWithShape?: boolean;
  private picLocks: Partial<Record<PicLockAttribute, boolean>> = {
    noChangeAspect: true,
    noChangeArrowheads: true,
  };
  private picNonVisualProps: PicNonVisualProperties = { id: '0', name: '', descr: '' };
  private isLinked: boolean = false;
  private svgRelationshipId?: string;

  // Group B: Raw XML passthrough for complex subtrees
  private _rawPassthrough: Map<string, string> = new Map();

  /**
   * Creates a new image from file path (async factory)
   * @param path File path
   * @param properties Additional properties
   * @returns Promise<Image>
   */
  static async fromFile(path: string, properties: Partial<ImageProperties> = {}): Promise<Image> {
    const image = new Image({ source: path, ...properties });
    await image.loadImageDataForDimensions();
    return image;
  }

  /**
   * Creates a new image from buffer (async factory)
   * Supports both modern and legacy API signatures
   *
   * @param buffer Image buffer
   * @param mimeTypeOrProperties MIME type string ('png', 'jpeg', etc.) or properties object
   * @param width Optional width in EMUs (legacy API)
   * @param height Optional height in EMUs (legacy API)
   * @returns Promise<Image>
   *
   * @example
   * // Modern API (recommended)
   * const img = await Image.fromBuffer(buffer, { mimeType: 'png', width: 914400, height: 914400 });
   *
   * // Legacy API (still supported)
   * const img = await Image.fromBuffer(buffer, 'png', 914400, 914400);
   */
  static async fromBuffer(
    buffer: Buffer,
    mimeTypeOrProperties?: string | Partial<ImageProperties>,
    width?: number,
    height?: number
  ): Promise<Image> {
    let properties: Partial<ImageProperties>;

    // Detect API signature
    if (typeof mimeTypeOrProperties === 'string') {
      // Legacy 4-parameter signature: fromBuffer(buffer, 'png', 914400, 914400)
      // Note: mimeType is ignored - extension is auto-detected from buffer
      properties = {
        width: width,
        height: height
      };
    } else {
      // Modern API: fromBuffer(buffer, { width: 914400, height: 914400 })
      properties = mimeTypeOrProperties || {};
    }

    const image = new Image({ source: buffer, ...properties });
    await image.loadImageDataForDimensions();
    return image;
  }

  /**
   * Unified create method for images (async factory)
   * @param properties Image properties including source (path or buffer)
   * @returns Promise<Image>
   */
  static async create(properties: ImageProperties): Promise<Image> {
    if (Buffer.isBuffer(properties.source)) {
      return Image.fromBuffer(properties.source, properties);
    } else if (typeof properties.source === 'string') {
      return Image.fromFile(properties.source, properties);
    } else {
      throw new Error('Invalid source: must be file path or Buffer');
    }
  }

  /**
   * Private constructor
   * @param properties Image properties
   */
  private constructor(properties: ImageProperties) {
    this.source = properties.source;
    this.description = properties.description || 'Image';
    this.name = properties.name || 'image';
    this.title = properties.title;
    this.relationshipId = properties.relationshipId;

    // Detect image extension
    this.extension = this.detectExtension();

    // Set default dimensions (6 inches x 4 inches) if not provided
    this.width = properties.width || inchesToEmus(6);
    this.height = properties.height || inchesToEmus(4);

    // Initialize advanced properties
    this.effectExtent = properties.effectExtent;
    this.wrap = properties.wrap;
    this.position = properties.position;
    this.anchor = properties.anchor;
    this.crop = properties.crop;
    this.effects = properties.effects;
    // Border: accept both legacy { width } and full ImageBorder
    if (properties.border) {
      this.border = properties.border as ImageBorder;
    }
    // Apply rotation if provided (normalize to 0-360)
    if (properties.rotation !== undefined && properties.rotation !== 0) {
      this.rotation = ((properties.rotation % 360) + 360) % 360;
    }
    // Apply flip attributes (ECMA-376 §20.1.7.6)
    this.flipH = properties.flipH || false;
    this.flipV = properties.flipV || false;

    // Group A: Simple attribute preservation
    if (properties.presetGeometry !== undefined) this.presetGeometry = properties.presetGeometry;
    if (properties.compressionState !== undefined) this.compressionState = properties.compressionState;
    if (properties.bwMode !== undefined) this.bwMode = properties.bwMode;
    if (properties.inlineDistT !== undefined) this.inlineDistT = properties.inlineDistT;
    if (properties.inlineDistB !== undefined) this.inlineDistB = properties.inlineDistB;
    if (properties.inlineDistL !== undefined) this.inlineDistL = properties.inlineDistL;
    if (properties.inlineDistR !== undefined) this.inlineDistR = properties.inlineDistR;
    if (properties.noChangeAspect !== undefined) this.noChangeAspect = properties.noChangeAspect;
    if (properties.hidden !== undefined) this.hidden = properties.hidden;
    if (properties.blipFillDpi !== undefined) this.blipFillDpi = properties.blipFillDpi;
    if (properties.blipFillRotWithShape !== undefined) this.blipFillRotWithShape = properties.blipFillRotWithShape;
    if (properties.picLocks !== undefined) this.picLocks = properties.picLocks;
    if (properties.picNonVisualProps !== undefined) this.picNonVisualProps = properties.picNonVisualProps;
    if (properties.isLinked !== undefined) this.isLinked = properties.isLinked;
    if (properties.svgRelationshipId !== undefined) this.svgRelationshipId = properties.svgRelationshipId;

    // Set default DPI
    this.dpi = 96;
  }

  /**
   * Loads image data temporarily for dimension detection only
   * Data is released after detection to save memory
   * @private
   */
  private async loadImageDataForDimensions(): Promise<void> {
    let tempData: Buffer | undefined;

    try {
      if (Buffer.isBuffer(this.source)) {
        tempData = this.source;
      } else if (typeof this.source === 'string') {
        await fs.access(this.source);
        tempData = await fs.readFile(this.source);
      }

      if (tempData) {
        this.imageData = tempData; // Temporarily store

        // Only auto-detect dimensions if they weren't explicitly provided
        // This preserves wp:extent values from parsed documents
        const defaultWidth = inchesToEmus(6);
        const defaultHeight = inchesToEmus(4);
        const hasExplicitDimensions = this.width !== defaultWidth || this.height !== defaultHeight;

        if (!hasExplicitDimensions) {
          const dimensions = this.detectDimensions();
          if (dimensions) {
            this.dpi = this.detectDPI() || 96;
            const emuPerInch = 914400;
            const pixelsPerInch = this.dpi;
            this.width = Math.round((dimensions.width / pixelsPerInch) * emuPerInch);
            this.height = Math.round((dimensions.height / pixelsPerInch) * emuPerInch);
          }
        }

        if (typeof this.source === 'string') {
          this.imageData = undefined; // Release
        }
      }
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      defaultLogger.error(`Failed to load image for dimensions: ${message}`);
      throw new Error(`Image loading failed: ${message}`);
    }
  }

  /**
   * Ensures image data is loaded (lazy loading)
   */
  async ensureDataLoaded(): Promise<void> {
    if (this.imageData) return;

    try {
      if (Buffer.isBuffer(this.source)) {
        this.imageData = this.source;
      } else if (typeof this.source === 'string') {
        await fs.access(this.source);
        this.imageData = await fs.readFile(this.source);
      } else {
        throw new Error('Invalid image source');
      }
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      defaultLogger.error(`Failed to load image data: ${message}`);
      throw new Error(`Image data loading failed: ${message}`);
    }
  }

  /**
   * Releases image data from memory
   */
  releaseData(): void {
    if (typeof this.source === 'string') {
      this.imageData = undefined;
    }
  }

  /**
   * Validates the image data integrity
   */
  validateImageData(): ValidationResult {
    // Skip validation for linked images (no data in package)
    if (this.isLinked) {
      return { valid: true };
    }

    if (!this.imageData || this.imageData.length === 0) {
      return { valid: false, error: 'Empty image data' };
    }

    // Skip signature validation for text-based formats (SVG)
    if (this.extension === 'svg') {
      return { valid: true };
    }

    const signatures: Record<string, number[]> = {
      png: [0x89, 0x50, 0x4E, 0x47],
      jpg: [0xFF, 0xD8],
      jpeg: [0xFF, 0xD8],
      gif: [0x47, 0x49, 0x46],
      bmp: [0x42, 0x4D],
      tiff: [0x49, 0x49, 0x2A, 0x00],
      tif: [0x49, 0x49, 0x2A, 0x00]
    };

    // EMF: check for ENHMETAHEADER signature at offset 40
    if (this.extension === 'emf') {
      if (this.imageData.length >= 44 &&
          this.imageData[40] === 0x20 && this.imageData[41] === 0x45 &&
          this.imageData[42] === 0x4D && this.imageData[43] === 0x46) {
        return { valid: true };
      }
      return { valid: false, error: 'Invalid EMF signature' };
    }

    // WMF: check for placeable or standard header
    if (this.extension === 'wmf') {
      if (this.imageData.length >= 4) {
        // Placeable WMF
        if (this.imageData[0] === 0xD7 && this.imageData[1] === 0xCD &&
            this.imageData[2] === 0xC6 && this.imageData[3] === 0x9A) {
          return { valid: true };
        }
        // Standard WMF
        if (this.imageData[0] === 0x01 && this.imageData[1] === 0x00 &&
            this.imageData[2] === 0x09 && this.imageData[3] === 0x00) {
          return { valid: true };
        }
      }
      return { valid: false, error: 'Invalid WMF signature' };
    }

    const sig = signatures[this.extension];
    if (sig) {
      for (let i = 0; i < sig.length; i++) {
        if (this.imageData[i] !== sig[i]) {
          return { valid: false, error: `Invalid ${this.extension.toUpperCase()} signature` };
        }
      }
    }

    return { valid: true };
  }

  /**
   * Detects image extension from source (path or buffer)
   */
  private detectExtension(): string {
    // Try path-based detection first
    if (typeof this.source === 'string') {
      const match = this.source.match(/\.([a-z]+)$/i);
      if (match && match[1]) {
        return match[1].toLowerCase();
      }
    }

    // Buffer-based detection using magic bytes
    if (Buffer.isBuffer(this.source) && this.source.length >= 4) {
      const buf = this.source;
      // PNG
      if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) return 'png';
      // JPEG
      if (buf[0] === 0xFF && buf[1] === 0xD8) return 'jpeg';
      // GIF
      if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) return 'gif';
      // BMP
      if (buf[0] === 0x42 && buf[1] === 0x4D) return 'bmp';
      // TIFF LE
      if (buf[0] === 0x49 && buf[1] === 0x49 && buf[2] === 0x2A && buf[3] === 0x00) return 'tiff';
      // TIFF BE
      if (buf[0] === 0x4D && buf[1] === 0x4D && buf[2] === 0x00 && buf[3] === 0x2A) return 'tiff';
      // EMF: byte 0 = 0x01,0x00,0x00,0x00 AND ' EMF' at offset 40
      if (buf.length >= 44 && buf[0] === 0x01 && buf[1] === 0x00 && buf[2] === 0x00 && buf[3] === 0x00 &&
          buf[40] === 0x20 && buf[41] === 0x45 && buf[42] === 0x4D && buf[43] === 0x46) return 'emf';
      // WMF placeable
      if (buf[0] === 0xD7 && buf[1] === 0xCD && buf[2] === 0xC6 && buf[3] === 0x9A) return 'wmf';
      // WMF standard
      if (buf[0] === 0x01 && buf[1] === 0x00 && buf[2] === 0x09 && buf[3] === 0x00) return 'wmf';
      // SVG: starts with '<' or UTF-8 BOM + '<'
      if (buf[0] === 0x3C || (buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF && buf.length > 3 && buf[3] === 0x3C)) return 'svg';
    }

    return 'png';
  }

  /**
   * Attempts to detect image dimensions from buffer
   */
  private detectDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 24) return null;

    if (this.imageData[0] === 0x89 && this.imageData[1] === 0x50 && this.imageData[2] === 0x4e && this.imageData[3] === 0x47) {
      return this.detectPngDimensions();
    }
    if (this.imageData[0] === 0xff && this.imageData[1] === 0xd8) {
      return this.detectJpegDimensions();
    }
    if (this.imageData[0] === 0x47 && this.imageData[1] === 0x49 && this.imageData[2] === 0x46) {
      return this.detectGifDimensions();
    }
    if (this.imageData[0] === 0x42 && this.imageData[1] === 0x4d) {
      return this.detectBmpDimensions();
    }
    if ((this.imageData[0] === 0x49 && this.imageData[1] === 0x49 && this.imageData[2] === 0x2a) ||
        (this.imageData[0] === 0x4d && this.imageData[1] === 0x4d && this.imageData[2] === 0x00)) {
      return this.detectTiffDimensions();
    }
    // EMF: ENHMETAHEADER has ' EMF' at offset 40
    if (this.imageData.length >= 44 &&
        this.imageData[40] === 0x20 && this.imageData[41] === 0x45 &&
        this.imageData[42] === 0x4D && this.imageData[43] === 0x46) {
      return this.detectEmfDimensions();
    }
    // WMF placeable
    if (this.imageData[0] === 0xD7 && this.imageData[1] === 0xCD &&
        this.imageData[2] === 0xC6 && this.imageData[3] === 0x9A) {
      return this.detectWmfDimensions();
    }
    // SVG (text-based)
    if (this.imageData[0] === 0x3C ||
        (this.imageData[0] === 0xEF && this.imageData[1] === 0xBB && this.imageData[2] === 0xBF)) {
      return this.detectSvgDimensions();
    }
    return null;
  }

  // Dimension detection helpers (as before, keeping them the same)

  private detectPngDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 24) return null;
    const width = this.imageData.readUInt32BE(16);
    const height = this.imageData.readUInt32BE(20);
    return { width, height };
  }

  private detectGifDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 10) return null;
    const width = this.imageData.readUInt16LE(6);
    const height = this.imageData.readUInt16LE(8);
    if (width > 0 && height > 0) return { width, height };
    return null;
  }

  private detectBmpDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 26) return null;
    const width = this.imageData.readInt32LE(18);
    const height = Math.abs(this.imageData.readInt32LE(22));
    if (width > 0 && height > 0) return { width, height };
    return null;
  }

  private detectTiffDimensions(): { width: number; height: number } | null {
    // Implementation as before
    if (!this.imageData || this.imageData.length < 14) return null;
    const isLittleEndian = this.imageData[0] === 0x49;
    const ifdOffset = isLittleEndian ? this.imageData.readUInt32LE(4) : this.imageData.readUInt32BE(4);
    if (ifdOffset + 14 > this.imageData.length) return null;
    const numEntries = isLittleEndian ? this.imageData.readUInt16LE(ifdOffset) : this.imageData.readUInt16BE(ifdOffset);
    let width = 0;
    let height = 0;
    for (let i = 0; i < numEntries; i++) {
      const entryOffset = ifdOffset + 2 + i * 12;
      if (entryOffset + 12 > this.imageData.length) break;
      const tag = isLittleEndian ? this.imageData.readUInt16LE(entryOffset) : this.imageData.readUInt16BE(entryOffset);
      const value = isLittleEndian ? this.imageData.readUInt32LE(entryOffset + 8) : this.imageData.readUInt32BE(entryOffset + 8);
      if (tag === 256) width = value;
      if (tag === 257) height = value;
      if (width > 0 && height > 0) break;
    }
    if (width > 0 && height > 0) return { width, height };
    return null;
  }

  private detectJpegDimensions(): { width: number; height: number } | null {
    // Implementation as before
    if (!this.imageData || this.imageData.length < 12) return null;
    let offset = 2;
    while (offset < this.imageData.length - 1) {
      if (this.imageData[offset] !== 0xff) break;
  const marker = this.imageData[offset + 1];
  if (marker === undefined) break;
  if (marker === 0x00 || marker === 0xff) {
    offset++;
    continue;
  }
  const isSOF = (marker >= 0xc0 && marker <= 0xcf) && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
      if (isSOF) {
        if (offset + 9 > this.imageData.length) break;
        const height = this.imageData.readUInt16BE(offset + 5);
        const width = this.imageData.readUInt16BE(offset + 7);
        if (width > 0 && height > 0) return { width, height };
      }
      if (marker === 0xda || marker === 0xd9) break;
      const segmentLength = this.imageData.readUInt16BE(offset + 2);
      if (segmentLength < 2 || offset + 2 + segmentLength > this.imageData.length) break;
      offset += 2 + segmentLength;
    }
    return null;
  }

  /**
   * Detects SVG dimensions from width/height attributes or viewBox
   */
  private detectSvgDimensions(): { width: number; height: number } | null {
    if (!this.imageData) return null;
    try {
      const svgText = this.imageData.toString('utf-8').substring(0, 2000);
      // Try width/height attributes on <svg> element
      const widthMatch = svgText.match(/<svg[^>]*\bwidth\s*=\s*["']?(\d+(?:\.\d+)?)/i);
      const heightMatch = svgText.match(/<svg[^>]*\bheight\s*=\s*["']?(\d+(?:\.\d+)?)/i);
      if (widthMatch?.[1] && heightMatch?.[1]) {
        return { width: Math.round(parseFloat(widthMatch[1])), height: Math.round(parseFloat(heightMatch[1])) };
      }
      // Try viewBox attribute
      const viewBoxMatch = svgText.match(/<svg[^>]*\bviewBox\s*=\s*["']?\s*[\d.]+\s+[\d.]+\s+([\d.]+)\s+([\d.]+)/i);
      if (viewBoxMatch?.[1] && viewBoxMatch?.[2]) {
        return { width: Math.round(parseFloat(viewBoxMatch[1])), height: Math.round(parseFloat(viewBoxMatch[2])) };
      }
    } catch {
      // SVG parsing failed
    }
    return null;
  }

  /**
   * Detects EMF dimensions from ENHMETAHEADER rclFrame (offsets 24-36, 0.01mm units)
   */
  private detectEmfDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 40) return null;
    try {
      // rclFrame: left(24), top(28), right(32), bottom(36) in 0.01mm units
      const left = this.imageData.readInt32LE(24);
      const top = this.imageData.readInt32LE(28);
      const right = this.imageData.readInt32LE(32);
      const bottom = this.imageData.readInt32LE(36);
      const widthMm = (right - left) / 100;
      const heightMm = (bottom - top) / 100;
      // Convert mm to pixels at 96 DPI (1 inch = 25.4mm)
      const width = Math.round((widthMm / 25.4) * 96);
      const height = Math.round((heightMm / 25.4) * 96);
      if (width > 0 && height > 0) return { width, height };
    } catch {
      // EMF header parsing failed
    }
    return null;
  }

  /**
   * Detects WMF dimensions from placeable WMF header bounding box (offsets 6-14)
   */
  private detectWmfDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 22) return null;
    try {
      // Placeable WMF header: left(6), top(8), right(10), bottom(12), inch(14)
      const left = this.imageData.readInt16LE(6);
      const top = this.imageData.readInt16LE(8);
      const right = this.imageData.readInt16LE(10);
      const bottom = this.imageData.readInt16LE(12);
      const unitsPerInch = this.imageData.readUInt16LE(14);
      if (unitsPerInch > 0) {
        const width = Math.round(((right - left) / unitsPerInch) * 96);
        const height = Math.round(((bottom - top) / unitsPerInch) * 96);
        if (width > 0 && height > 0) return { width, height };
      }
    } catch {
      // WMF header parsing failed
    }
    return null;
  }

  /**
   * Gets the image data buffer asynchronously
   */
  async getImageDataAsync(): Promise<Buffer> {
    await this.ensureDataLoaded();
    if (!this.imageData) throw new Error('Failed to load image data');
    return this.imageData;
  }

  /**
   * Gets the image data buffer synchronously
   */
  getImageData(): Buffer {
    if (!this.imageData) throw new Error('Image data not loaded. Call ensureDataLoaded first.');
    return this.imageData;
  }

  getExtension(): string {
    return this.extension;
  }

  getDPI(): number {
    return this.dpi;
  }

  getWidth(): number {
    return this.width;
  }

  getHeight(): number {
    return this.height;
  }

  getImageDataSafe(): Buffer | null {
    return this.imageData ?? null;
  }

  setWidth(width: number, maintainAspectRatio: boolean = true): this {
    if (maintainAspectRatio && this.height > 0) {
      const ratio = this.height / this.width;
      this.height = Math.round(width * ratio);
    }
    this.width = width;
    return this;
  }

  setHeight(height: number, maintainAspectRatio: boolean = true): this {
    if (maintainAspectRatio && this.width > 0) {
      const ratio = this.width / this.height;
      this.width = Math.round(height * ratio);
    }
    this.height = height;
    return this;
  }

  setSize(width: number, height: number): this {
    this.width = width;
    this.height = height;
    return this;
  }

  async updateImageData(newSource: string | Buffer): Promise<void> {
    this.source = newSource;
    this.imageData = undefined;
    await this.loadImageDataForDimensions();
    this.extension = this.detectExtension();
    this.dpi = this.detectDPI() || 96;
  }

  setRelationshipId(relationshipId: string): this {
    this.relationshipId = relationshipId;
    return this;
  }

  getRelationshipId(): string | undefined {
    return this.relationshipId;
  }

  setDocPrId(id: number): this {
    this.docPrId = id;
    return this;
  }

  setAltText(altText: string): this {
    this.description = altText;
    return this;
  }

  getAltText(): string {
    return this.description;
  }

  setTitle(title: string): this {
    this.title = title;
    return this;
  }

  getTitle(): string | undefined {
    return this.title;
  }

  rotate(degrees: number): this {
    this.rotation = ((degrees % 360) + 360) % 360;
    if (this.rotation === 90 || this.rotation === 270) {
      [this.width, this.height] = [this.height, this.width];
    }
    return this;
  }

  getRotation(): number {
    return this.rotation;
  }

  setFlipH(flip: boolean): this {
    this.flipH = flip;
    return this;
  }

  getFlipH(): boolean {
    return this.flipH;
  }

  setFlipV(flip: boolean): this {
    this.flipV = flip;
    return this;
  }

  getFlipV(): boolean {
    return this.flipV;
  }

  // --- Group A: Simple attribute getters/setters ---

  getPresetGeometry(): PresetGeometry { return this.presetGeometry; }
  setPresetGeometry(geom: PresetGeometry): this { this.presetGeometry = geom; return this; }

  getCompressionState(): BlipCompressionState { return this.compressionState; }
  setCompressionState(state: BlipCompressionState): this { this.compressionState = state; return this; }

  getBwMode(): string { return this.bwMode; }
  setBwMode(mode: string): this { this.bwMode = mode; return this; }

  getInlineDistT(): number { return this.inlineDistT; }
  getInlineDistB(): number { return this.inlineDistB; }
  getInlineDistL(): number { return this.inlineDistL; }
  getInlineDistR(): number { return this.inlineDistR; }
  setInlineDist(distT: number, distB: number, distL: number, distR: number): this {
    this.inlineDistT = distT;
    this.inlineDistB = distB;
    this.inlineDistL = distL;
    this.inlineDistR = distR;
    return this;
  }

  getNoChangeAspect(): boolean { return this.noChangeAspect; }
  setNoChangeAspect(val: boolean): this { this.noChangeAspect = val; return this; }

  getHidden(): boolean { return this.hidden; }
  setHidden(val: boolean): this { this.hidden = val; return this; }

  getBlipFillDpi(): number | undefined { return this.blipFillDpi; }
  setBlipFillDpi(dpi: number | undefined): this { this.blipFillDpi = dpi; return this; }

  getBlipFillRotWithShape(): boolean | undefined { return this.blipFillRotWithShape; }
  setBlipFillRotWithShape(val: boolean | undefined): this { this.blipFillRotWithShape = val; return this; }

  getPicLocks(): Partial<Record<PicLockAttribute, boolean>> { return { ...this.picLocks }; }
  setPicLocks(locks: Partial<Record<PicLockAttribute, boolean>>): this { this.picLocks = locks; return this; }

  getPicNonVisualProps(): PicNonVisualProperties { return { ...this.picNonVisualProps }; }
  setPicNonVisualProps(props: PicNonVisualProperties): this { this.picNonVisualProps = props; return this; }

  getIsLinked(): boolean { return this.isLinked; }
  setIsLinked(val: boolean): this { this.isLinked = val; return this; }

  getSvgRelationshipId(): string | undefined { return this.svgRelationshipId; }
  setSvgRelationshipId(id: string | undefined): this { this.svgRelationshipId = id; return this; }

  // --- Group B: Raw passthrough storage ---

  /** @internal */
  _setRawPassthrough(slot: string, xml: string): void {
    this._rawPassthrough.set(slot, xml);
  }

  /** @internal */
  _getRawPassthrough(slot: string): string | undefined {
    return this._rawPassthrough.get(slot);
  }

  /** @internal */
  _hasRawPassthrough(slot: string): boolean {
    return this._rawPassthrough.has(slot);
  }

  // --- Group C: Enhanced border ---

  getBorder(): ImageBorder | undefined { return this.border; }

  setEffectExtent(left: number, top: number, right: number, bottom: number): this {
    this.effectExtent = { left, top, right, bottom };
    return this;
  }

  getEffectExtent(): EffectExtent | undefined {
    return this.effectExtent;
  }

  setWrap(type: WrapType, side?: WrapSide, distances?: { top?: number; bottom?: number; left?: number; right?: number }): this {
    this.wrap = {
      type,
      side,
      distanceTop: distances?.top,
      distanceBottom: distances?.bottom,
      distanceLeft: distances?.left,
      distanceRight: distances?.right,
    };
    return this;
  }

  getWrap(): TextWrapSettings | undefined {
    return this.wrap;
  }

  /**
   * Validates a position offset value
   * @param offset - Offset value in EMUs
   * @param axis - 'horizontal' or 'vertical' for error messages
   * @throws {Error} If offset exceeds maximum reasonable value
   * @private
   */
  private validatePositionOffset(offset: number | undefined, axis: string): void {
    if (offset === undefined) return;

    // Maximum reasonable offset: 50 inches = 45,720,000 EMUs
    const MAX_OFFSET_EMUS = 45720000;
    if (Math.abs(offset) > MAX_OFFSET_EMUS) {
      throw new Error(
        `Invalid ${axis} position offset: ${offset} EMUs exceeds maximum of ${MAX_OFFSET_EMUS} EMUs (50 inches).`
      );
    }
  }

  /**
   * Sets the position for a floating image
   *
   * Position can be specified using either:
   * - Absolute offset (in EMUs from the anchor point)
   * - Relative alignment (left, center, right / top, center, bottom)
   *
   * @param horizontal - Horizontal positioning configuration
   * @param vertical - Vertical positioning configuration
   * @returns This image for chaining
   * @throws {Error} If offset values exceed maximum
   *
   * @example
   * ```typescript
   * // Absolute positioning (100,000 EMUs from page edges)
   * image.setPosition(
   *   { anchor: 'page', offset: 100000 },
   *   { anchor: 'page', offset: 100000 }
   * );
   *
   * // Relative alignment (centered on page)
   * image.setPosition(
   *   { anchor: 'page', alignment: 'center' },
   *   { anchor: 'page', alignment: 'center' }
   * );
   * ```
   */
  setPosition(horizontal: ImagePosition['horizontal'], vertical: ImagePosition['vertical']): this {
    // Validate offset values
    this.validatePositionOffset(horizontal.offset, 'horizontal');
    this.validatePositionOffset(vertical.offset, 'vertical');

    this.position = { horizontal, vertical };
    return this;
  }

  getPosition(): ImagePosition | undefined {
    return this.position;
  }

  /**
   * Validates the current image position configuration
   *
   * Checks for common configuration issues:
   * - Missing anchor when offset is used
   * - Conflicting offset and alignment values
   * - Invalid combinations
   *
   * @returns Validation result with details
   *
   * @example
   * ```typescript
   * const result = image.validatePosition();
   * if (!result.isValid) {
   *   console.log(result.warnings); // Array of warning messages
   * }
   * ```
   */
  validatePosition(): {
    isValid: boolean;
    warnings: string[];
  } {
    const warnings: string[] = [];

    if (!this.position) {
      return { isValid: true, warnings };
    }

    // Check if both offset and alignment are specified (unusual but not invalid)
    if (this.position.horizontal.offset !== undefined && this.position.horizontal.alignment) {
      warnings.push(
        'Horizontal position has both offset and alignment. Word will use alignment and ignore offset.'
      );
    }

    if (this.position.vertical.offset !== undefined && this.position.vertical.alignment) {
      warnings.push(
        'Vertical position has both offset and alignment. Word will use alignment and ignore offset.'
      );
    }

    // Check for floating image without anchor settings
    if (this.position && !this.anchor) {
      warnings.push(
        'Position is set but anchor is not. Consider setting anchor properties for proper floating behavior.'
      );
    }

    return {
      isValid: warnings.length === 0,
      warnings,
    };
  }

  setAnchor(options: ImageAnchor): this {
    this.anchor = options;
    return this;
  }

  getAnchor(): ImageAnchor | undefined {
    return this.anchor;
  }

  setCrop(left: number, top: number, right: number, bottom: number): this {
    const clamp = (val: number) => Math.max(0, Math.min(100, val));
    this.crop = { left: clamp(left), top: clamp(top), right: clamp(right), bottom: clamp(bottom) };
    return this;
  }

  getCrop(): ImageCrop | undefined {
    return this.crop;
  }

  setEffects(options: ImageEffects): this {
    const clamp = (val?: number) => val !== undefined ? Math.max(-100, Math.min(100, val)) : undefined;
    this.effects = { brightness: clamp(options.brightness), contrast: clamp(options.contrast), grayscale: options.grayscale, transparency: options.transparency !== undefined ? Math.max(0, Math.min(100, options.transparency)) : undefined };
    return this;
  }

  getEffects(): ImageEffects | undefined {
    return this.effects;
  }

  private detectDPI(): number | undefined {
    if (!this.imageData) return undefined;

    try {
      if (this.extension === 'png') {
        const physIndex = this.imageData.indexOf(Buffer.from([0x70, 0x48, 0x59, 0x73]));
        if (physIndex !== -1 && physIndex + 12 < this.imageData.length) {
          const xPixelsPerMeter = this.imageData.readUInt32BE(physIndex + 4);
          const yPixelsPerMeter = this.imageData.readUInt32BE(physIndex + 8);
          const unit = this.imageData[physIndex + 12];
          if (unit === 1) {
            const dpiX = Math.round(xPixelsPerMeter * 0.0254);
            const dpiY = Math.round(yPixelsPerMeter * 0.0254);
            return Math.min(dpiX, dpiY);
          }
        }
      } else if (this.extension === 'jpg' || this.extension === 'jpeg') {
        let offset = 2;
        while (offset < this.imageData.length) {
          if (this.imageData[offset] !== 0xFF) break;
          const marker = this.imageData[offset + 1];
          if (marker === 0xE0) {
            const length = this.imageData.readUInt16BE(offset + 2);
            if (length >= 16 && this.imageData.slice(offset + 4, offset + 9).toString('ascii') === 'JFIF\0') {
              const units = this.imageData[offset + 11];
              const xDensity = this.imageData.readUInt16BE(offset + 12);
              const yDensity = this.imageData.readUInt16BE(offset + 14);
              if (units === 1) return Math.min(xDensity, yDensity);
              if (units === 2) return Math.min(Math.round(xDensity * 2.54), Math.round(yDensity * 2.54));
            }
            offset += 2 + length;
            continue;
          }
          offset += 2 + this.imageData.readUInt16BE(offset + 2);
        }
      }
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      defaultLogger.warn(`DPI detection failed: ${message}`);
    }
    return undefined;
  }

  isFloating(): boolean {
    return !!this.anchor || !!this.position;
  }

  floatTopLeft(marginTop: number = 0, marginLeft: number = 0): this {
    this.setPosition(
      { anchor: 'page', offset: marginLeft },
      { anchor: 'page', offset: marginTop }
    );
    this.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: true,
      relativeHeight: 251658240
    });
    this.setWrap('square', 'bothSides');
    return this;
  }

  floatTopRight(marginTop: number = 0, marginRight: number = 0): this {
    this.setPosition(
      { anchor: 'page', alignment: 'right', offset: -marginRight },
      { anchor: 'page', offset: marginTop }
    );
    this.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: true,
      relativeHeight: 251658240
    });
    this.setWrap('square', 'bothSides');
    return this;
  }

  floatCenter(): this {
    this.setPosition(
      { anchor: 'page', alignment: 'center' },
      { anchor: 'page', alignment: 'center' }
    );
    this.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: true,
      relativeHeight: 251658240
    });
    this.setWrap('square', 'bothSides');
    return this;
  }

  setBehindText(behind: boolean = true): this {
    if (this.anchor) {
      this.anchor.behindDoc = behind;
    } else {
      this.setAnchor({
        behindDoc: behind,
        locked: false,
        layoutInCell: true,
        allowOverlap: true,
        relativeHeight: 251658240
      });
    }
    return this;
  }

  /**
   * Applies a border around the image
   * @param thicknessOrOptions Border thickness in points (number) or full ImageBorder options
   * @returns This image for chaining
   *
   * Note: effectExtent is set to accommodate the border width so it renders
   * properly without being clipped. The border is drawn centered on the image
   * edge, so half the border width extends outside the image bounds.
   */
  setBorder(thicknessOrOptions: number | ImageBorder = 2): this {
    if (typeof thicknessOrOptions === 'number') {
      this.border = { width: thicknessOrOptions };
    } else {
      this.border = thicknessOrOptions;
    }

    // Calculate space needed for border (half-width on each side)
    // Border is drawn centered on the edge
    const borderEmu = this.border.width * UNITS.EMUS_PER_POINT;
    const halfBorderEmu = Math.ceil(borderEmu / 2);

    // Ensure effectExtent has at least enough space for the border
    if (!this.effectExtent) {
      this.effectExtent = { left: 0, top: 0, right: 0, bottom: 0 };
    }
    this.effectExtent.left = Math.max(this.effectExtent.left, halfBorderEmu);
    this.effectExtent.top = Math.max(this.effectExtent.top, halfBorderEmu);
    this.effectExtent.right = Math.max(this.effectExtent.right, halfBorderEmu);
    this.effectExtent.bottom = Math.max(this.effectExtent.bottom, halfBorderEmu);

    return this;
  }

  /**
   * Removes the border from the image
   * @returns This image for chaining
   */
  removeBorder(): this {
    this.border = undefined;
    return this;
  }

  /**
   * @deprecated Use setBorder() instead. This method will be removed in a future version.
   * Applies a 2-point black border around the image.
   * @returns This image for chaining
   */
  applyTwoPixelBlackBorder(): this {
    return this.setBorder(2);
  }

  toXML(): XMLElement {
    const isFloating = this.isFloating();

    // Common elements - must include wp: namespace prefix
    const extent = XMLBuilder.wp('extent', { cx: this.width.toString(), cy: this.height.toString() });

    // --- Build blip element with effects ---
    const blipChildren: XMLElement[] = [];

    // Add luminance effect for brightness/contrast (per ECMA-376 §20.1.8.43)
    if (this.effects?.brightness !== undefined || this.effects?.contrast !== undefined) {
      const lumAttrs: Record<string, string> = {};
      if (this.effects.brightness !== undefined) {
        lumAttrs.bright = Math.round(this.effects.brightness * 1000).toString();
      }
      if (this.effects.contrast !== undefined) {
        lumAttrs.contrast = Math.round(this.effects.contrast * 1000).toString();
      }
      blipChildren.push(XMLBuilder.aSelf('lum', lumAttrs));
    }

    // Add grayscale effect (per ECMA-376 §20.1.8.37)
    if (this.effects?.grayscale) {
      blipChildren.push(XMLBuilder.aSelf('grayscl'));
    }

    // Add transparency effect via a:alphaModFix (ECMA-376 §20.1.8.4)
    if (this.effects?.transparency !== undefined && this.effects.transparency > 0) {
      // transparency is 0-100%, alphaModFix amt is in 1/1000ths of percent
      // e.g., 50% transparency = 50000 amt (= 50% opacity)
      const amt = Math.round((100 - this.effects.transparency) * 1000);
      blipChildren.push(XMLBuilder.aSelf('alphaModFix', { amt: amt.toString() }));
    }

    // Group B: Inject raw blip effects passthrough (a:clrChange, a:duotone, etc.)
    if (this._rawPassthrough.has('blip-effects')) {
      blipChildren.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('blip-effects')! } as XMLElement);
    }

    // Group B: Inject raw blip extLst passthrough (must come last per schema)
    if (this._rawPassthrough.has('blip-extLst')) {
      blipChildren.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('blip-extLst')! } as XMLElement);
    } else if (this.svgRelationshipId) {
      // SVG dual-relationship: add asvg:svgBlip reference in extLst
      blipChildren.push({
        name: '__rawXml',
        rawXml: `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${this.svgRelationshipId}"/></a:ext></a:extLst>`,
      } as XMLElement);
    }

    // Build blip attributes: r:embed or r:link, cstate
    const blipAttrs: Record<string, string | undefined> = {
      cstate: this.compressionState,
    };
    if (this.isLinked) {
      blipAttrs['r:link'] = this.relationshipId;
    } else {
      blipAttrs['r:embed'] = this.relationshipId;
    }

    const blip = blipChildren.length > 0
      ? XMLBuilder.a('blip', blipAttrs, blipChildren)
      : XMLBuilder.a('blip', blipAttrs);

    // --- Build transform (a:xfrm) ---
    const xfrmAttrs: Record<string, string> | undefined = (() => {
      const attrs: Record<string, string> = {};
      if (this.rotation > 0) attrs.rot = Math.round(this.rotation * 60000).toString();
      if (this.flipH) attrs.flipH = '1';
      if (this.flipV) attrs.flipV = '1';
      return Object.keys(attrs).length > 0 ? attrs : undefined;
    })();
    const xfrm = XMLBuilder.a('xfrm', xfrmAttrs, [
      XMLBuilder.a('off', { x: '0', y: '0' }),
      XMLBuilder.a('ext', { cx: this.width.toString(), cy: this.height.toString() })
    ]);

    // --- Build shape properties (pic:spPr) ---
    const spPrChildren: XMLElement[] = [xfrm];

    // Geometry: use passthrough for custGeom or prstGeom with avLst, otherwise default
    if (this._rawPassthrough.has('geometry')) {
      spPrChildren.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('geometry')! } as XMLElement);
    } else {
      spPrChildren.push(XMLBuilder.a('prstGeom', { prst: this.presetGeometry }, [
        XMLBuilder.a('avLst')
      ]));
    }

    // Border (a:ln) - full model (Group C)
    if (this.border) {
      // Add noFill element before the border line (required by Word)
      spPrChildren.push(XMLBuilder.a('noFill'));

      const ptToEmu = 12700;
      const widthEmu = this.border.width * ptToEmu;
      const lnAttrs: Record<string, string> = { w: widthEmu.toString() };
      if (this.border.cap) lnAttrs.cap = this.border.cap;
      if (this.border.compound) lnAttrs.cmpd = this.border.compound;
      if (this.border.alignment) lnAttrs.algn = this.border.alignment;

      const lnChildren: XMLElement[] = [];

      // Fill
      if (this.border.rawFillXml) {
        lnChildren.push({ name: '__rawXml', rawXml: this.border.rawFillXml } as XMLElement);
      } else if (this.border.fill) {
        const colorChildren: XMLElement[] = [];
        if (this.border.fill.modifiers) {
          for (const mod of this.border.fill.modifiers) {
            colorChildren.push(XMLBuilder.aSelf(mod.name, { val: mod.val }));
          }
        }
        const colorEl = colorChildren.length > 0
          ? XMLBuilder.a(this.border.fill.type, { val: this.border.fill.value }, colorChildren)
          : XMLBuilder.a(this.border.fill.type, { val: this.border.fill.value });
        lnChildren.push(XMLBuilder.a('solidFill', undefined, [colorEl]));
      } else {
        // Default: scheme color tx1 (backward compat)
        lnChildren.push(XMLBuilder.a('solidFill', undefined, [
          XMLBuilder.a('schemeClr', { val: 'tx1' })
        ]));
      }

      // Dash pattern
      if (this.border.dashPattern) {
        lnChildren.push(XMLBuilder.aSelf('prstDash', { val: this.border.dashPattern }));
      }

      // Join
      if (this.border.join === 'round') {
        lnChildren.push(XMLBuilder.aSelf('round'));
      } else if (this.border.join === 'bevel') {
        lnChildren.push(XMLBuilder.aSelf('bevel'));
      } else if (this.border.join === 'miter') {
        const miterAttrs: Record<string, string> = {};
        if (this.border.miterLimit) miterAttrs.lim = this.border.miterLimit.toString();
        lnChildren.push(XMLBuilder.aSelf('miter', miterAttrs));
      }

      // Head/tail end
      if (this.border.headEnd) {
        lnChildren.push(XMLBuilder.aSelf('headEnd', this.border.headEnd as Record<string, string>));
      }
      if (this.border.tailEnd) {
        lnChildren.push(XMLBuilder.aSelf('tailEnd', this.border.tailEnd as Record<string, string>));
      }

      spPrChildren.push(XMLBuilder.a('ln', lnAttrs, lnChildren));
    }

    // Group B: Inject raw spPr effects passthrough (effectLst, scene3d, sp3d, etc.)
    if (this._rawPassthrough.has('spPr-effects')) {
      spPrChildren.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('spPr-effects')! } as XMLElement);
    }

    // --- Build pic:cNvPr with passthrough ---
    const cNvPrChildren: XMLElement[] = [];
    if (this._rawPassthrough.has('cNvPr-extra')) {
      cNvPrChildren.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('cNvPr-extra')! } as XMLElement);
    }
    const cNvPr = cNvPrChildren.length > 0
      ? XMLBuilder.pic('cNvPr', {
          id: this.picNonVisualProps.id,
          name: this.picNonVisualProps.name,
          descr: this.picNonVisualProps.descr,
        }, cNvPrChildren)
      : XMLBuilder.pic('cNvPr', {
          id: this.picNonVisualProps.id,
          name: this.picNonVisualProps.name,
          descr: this.picNonVisualProps.descr,
        });

    // --- Build picLocks from map ---
    const picLocksAttrs: Record<string, string> = {};
    for (const [key, val] of Object.entries(this.picLocks)) {
      if (val) picLocksAttrs[key] = '1';
    }

    // --- Build blipFill ---
    const blipFillAttrs: Record<string, string> = {};
    if (this.blipFillDpi !== undefined) blipFillAttrs.dpi = this.blipFillDpi.toString();
    if (this.blipFillRotWithShape !== undefined) blipFillAttrs.rotWithShape = this.blipFillRotWithShape ? '1' : '0';

    const blipFillChildren: XMLElement[] = [blip];
    // Crop values are stored as percentages (0-100), serialized as per-mille (0-100000)
    if (this.crop) {
      blipFillChildren.push(XMLBuilder.a('srcRect', {
        l: Math.round(this.crop.left * 1000).toString(),
        t: Math.round(this.crop.top * 1000).toString(),
        r: Math.round(this.crop.right * 1000).toString(),
        b: Math.round(this.crop.bottom * 1000).toString(),
      }));
    }
    // Group B: Use tile passthrough instead of stretch when present
    if (this._rawPassthrough.has('blipFill-extra')) {
      blipFillChildren.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('blipFill-extra')! } as XMLElement);
    } else {
      blipFillChildren.push(XMLBuilder.a('stretch', undefined, [XMLBuilder.a('fillRect')]));
    }

    const blipFillAttrsObj = Object.keys(blipFillAttrs).length > 0 ? blipFillAttrs : undefined;

    const graphicData = XMLBuilder.a('graphicData', { uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture' }, [
      XMLBuilder.pic('pic', undefined, [
        XMLBuilder.pic('nvPicPr', undefined, [
          cNvPr,
          XMLBuilder.pic('cNvPicPr', undefined, [
            XMLBuilder.a('picLocks', Object.keys(picLocksAttrs).length > 0 ? picLocksAttrs : undefined)
          ])
        ]),
        XMLBuilder.pic('blipFill', blipFillAttrsObj, blipFillChildren),
        XMLBuilder.pic('spPr', { bwMode: this.bwMode }, spPrChildren)
      ])
    ]);

    const graphic = XMLBuilder.a('graphic', undefined, [graphicData]);

    // --- Build docPr element (shared between inline and floating) ---
    const buildDocPr = (idVal: string | number): XMLElement => {
      const attrs: Record<string, any> = { id: idVal, name: this.name, descr: this.description };
      if (this.title) attrs.title = this.title;
      if (this.hidden) attrs.hidden = '1';
      const children: XMLElement[] = [];
      if (this._rawPassthrough.has('docPr-extra')) {
        children.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('docPr-extra')! } as XMLElement);
      }
      return children.length > 0
        ? XMLBuilder.wp('docPr', attrs, children)
        : XMLBuilder.wp('docPr', attrs);
    };

    // --- Build cNvGraphicFramePr element (shared between inline and floating) ---
    const buildCNvGraphicFramePr = (): XMLElement => {
      return XMLBuilder.wp('cNvGraphicFramePr', undefined, [
        XMLBuilder.a('graphicFrameLocks', {
          'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          noChangeAspect: this.noChangeAspect ? '1' : '0',
        })
      ]);
    };

    if (isFloating) {
      // Floating image (anchor)
      const positionHChildren: XMLElement[] = [];
      if (this.position?.horizontal.alignment) {
        positionHChildren.push(XMLBuilder.wp('align', undefined, [this.position.horizontal.alignment]));
      } else {
        positionHChildren.push(XMLBuilder.wp('posOffset', undefined, [(this.position?.horizontal.offset || 0).toString()]));
      }
      const positionH = XMLBuilder.wp('positionH', { relativeFrom: this.position?.horizontal.anchor || 'page' }, positionHChildren);

      const positionVChildren: XMLElement[] = [];
      if (this.position?.vertical.alignment) {
        positionVChildren.push(XMLBuilder.wp('align', undefined, [this.position.vertical.alignment]));
      } else {
        positionVChildren.push(XMLBuilder.wp('posOffset', undefined, [(this.position?.vertical.offset || 0).toString()]));
      }
      const positionV = XMLBuilder.wp('positionV', { relativeFrom: this.position?.vertical.anchor || 'page' }, positionVChildren);

      // Effect extent for floating images (required by Word)
      const floatEffectExt = this.effectExtent || { left: 0, top: 0, right: 0, bottom: 0 };
      const effectExtentElement = XMLBuilder.wp('effectExtent', {
        t: floatEffectExt.top.toString(),
        r: floatEffectExt.right.toString(),
        b: floatEffectExt.bottom.toString(),
        l: floatEffectExt.left.toString()
      });

      const anchorChildren: XMLElement[] = [
        positionH,
        positionV,
        extent,
        effectExtentElement
      ];

      // Wrap element with optional polygon passthrough
      if (this.wrap) {
        const wrapAttrs: Record<string, any> = {};
        if (this.wrap.distanceTop !== undefined) wrapAttrs.distT = this.wrap.distanceTop;
        if (this.wrap.distanceBottom !== undefined) wrapAttrs.distB = this.wrap.distanceBottom;
        if (this.wrap.distanceLeft !== undefined) wrapAttrs.distL = this.wrap.distanceLeft;
        if (this.wrap.distanceRight !== undefined) wrapAttrs.distR = this.wrap.distanceRight;
        if (this.wrap.side) wrapAttrs.wrapText = this.wrap.side;

        let wrapElementName: string;
        switch (this.wrap.type) {
          case 'square': wrapElementName = 'wrapSquare'; break;
          case 'tight': wrapElementName = 'wrapTight'; break;
          case 'through': wrapElementName = 'wrapThrough'; break;
          case 'topAndBottom': wrapElementName = 'wrapTopAndBottom'; break;
          case 'none': wrapElementName = 'wrapNone'; break;
          default: wrapElementName = 'wrapSquare';
        }

        // Group B: Include wrap polygon passthrough as children
        const wrapChildren: XMLElement[] = [];
        if (this._rawPassthrough.has('wrap-polygon')) {
          wrapChildren.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('wrap-polygon')! } as XMLElement);
        }

        anchorChildren.push(
          wrapChildren.length > 0
            ? XMLBuilder.wp(wrapElementName, wrapAttrs, wrapChildren)
            : XMLBuilder.wp(wrapElementName, wrapAttrs)
        );
      }

      anchorChildren.push(buildDocPr(this.docPrId));
      anchorChildren.push(buildCNvGraphicFramePr());
      anchorChildren.push(graphic);

      // Group B: Inject anchor extras (wp14:sizeRelH, wp14:sizeRelV)
      if (this._rawPassthrough.has('anchor-extra')) {
        anchorChildren.push({ name: '__rawXml', rawXml: this._rawPassthrough.get('anchor-extra')! } as XMLElement);
      }

      // Build anchor attributes including simplePos and distance from text
      const anchorAttrs: Record<string, any> = {
        behindDoc: this.anchor?.behindDoc ? 1 : 0,
        locked: this.anchor?.locked ? 1 : 0,
        layoutInCell: this.anchor?.layoutInCell ? 1 : 0,
        allowOverlap: this.anchor?.allowOverlap ? 1 : 0,
        relativeHeight: this.anchor?.relativeHeight,
        simplePos: this.anchor?.simplePos ? '1' : '0',
        distT: (this.anchor?.distT ?? 0).toString(),
        distB: (this.anchor?.distB ?? 0).toString(),
        distL: (this.anchor?.distL ?? 0).toString(),
        distR: (this.anchor?.distR ?? 0).toString(),
      };
      if (this.hidden) anchorAttrs.hidden = '1';

      // Add wp:simplePos child element (required by ECMA-376 even when simplePos="0")
      anchorChildren.unshift(XMLBuilder.wp('simplePos', { x: '0', y: '0' }));

      return XMLBuilder.w('drawing', undefined, [
        XMLBuilder.wp('anchor', anchorAttrs, anchorChildren)
      ]);
    } else {
      // Inline image
      const effectExt = this.effectExtent || { left: 0, top: 0, right: 0, bottom: 0 };

      return XMLBuilder.w('drawing', undefined, [
        XMLBuilder.wp('inline', {
          distT: this.inlineDistT.toString(),
          distB: this.inlineDistB.toString(),
          distL: this.inlineDistL.toString(),
          distR: this.inlineDistR.toString(),
        }, [
          extent,
          XMLBuilder.wp('effectExtent', {
            t: effectExt.top.toString(),
            r: effectExt.right.toString(),
            b: effectExt.bottom.toString(),
            l: effectExt.left.toString()
          }),
          buildDocPr(this.docPrId.toString()),
          buildCNvGraphicFramePr(),
          graphic
        ])
      ]);
    }
  }
}
