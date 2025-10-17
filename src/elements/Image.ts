/**
 * Image - Represents an embedded image in a Word document
 *
 * Images use DrawingML (a:) and WordprocessingML Drawing (wp:) namespaces
 * for proper positioning and formatting in Word documents.
 */

import { promises as fs } from 'fs';
import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { inchesToEmus } from '../utils/units';

/**
 * Supported image formats
 */
export type ImageFormat = 'png' | 'jpeg' | 'jpg' | 'gif' | 'bmp' | 'tiff';

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
  /** Relationship ID (will be set by ImageManager) */
  relationshipId?: string;
}

/**
 * Represents an embedded image
 */
export class Image {
  private source: string | Buffer;
  private width: number;
  private height: number;
  private description: string;
  private name: string;
  private relationshipId?: string;
  private imageData?: Buffer;
  private extension: string;
  private docPrId: number = 1;

  /**
   * Creates a new image
   * @param properties Image properties
   * @private Use static factory methods instead (fromFile, fromBuffer, create)
   */
  private constructor(properties: ImageProperties) {
    this.source = properties.source;
    this.description = properties.description || 'Image';
    this.name = properties.name || 'image';
    this.relationshipId = properties.relationshipId;

    // Detect image extension
    this.extension = this.detectExtension();

    // Set default dimensions (6 inches x 4 inches) if not provided
    // Will be overridden if we can detect actual dimensions
    this.width = properties.width || inchesToEmus(6);
    this.height = properties.height || inchesToEmus(4);

    // Note: Dimension detection now happens in factory methods
    // This keeps constructor synchronous
  }

  /**
   * Loads image data temporarily for dimension detection only
   * Data is released after detection to save memory
   * @private Use this internally during initialization
   */
  private async loadImageDataForDimensions(): Promise<void> {
    let tempData: Buffer | undefined;

    if (Buffer.isBuffer(this.source)) {
      tempData = this.source;
    } else if (typeof this.source === 'string') {
      try {
        // Check if file exists asynchronously
        await fs.access(this.source);
        // Read file asynchronously
        tempData = await fs.readFile(this.source);
      } catch (error) {
        // File doesn't exist or can't be read
        // Store error for later retrieval, don't throw during init
        throw new Error(`Could not read image file: ${this.source}`);
      }
    }

    // Try to detect dimensions from image data
    if (tempData) {
      this.imageData = tempData; // Temporarily store for detection
      const dimensions = this.detectDimensions();
      if (dimensions) {
        this.width = dimensions.width;
        this.height = dimensions.height;
      }

      // Release data immediately after dimension detection
      // Data will be reloaded during save phase if needed
      if (typeof this.source === 'string') {
        // Only release if loaded from file (buffer sources are kept)
        this.imageData = undefined;
      }
    }
  }

  /**
   * Ensures image data is loaded (lazy loading)
   * Call this before accessing image data
   */
  async ensureDataLoaded(): Promise<void> {
    if (this.imageData) {
      return; // Already loaded
    }

    if (Buffer.isBuffer(this.source)) {
      this.imageData = this.source;
    } else if (typeof this.source === 'string') {
      try {
        // Check if file exists asynchronously
        await fs.access(this.source);
        // Read file asynchronously
        this.imageData = await fs.readFile(this.source);
      } catch (error) {
        throw new Error(
          `Failed to load image from ${this.source}: ${error instanceof Error ? error.message : error}`
        );
      }
    }
  }

  /**
   * Releases image data from memory
   * Use this after saving to free memory with large images
   */
  releaseData(): void {
    // Only release if loaded from file path (keep buffer sources)
    if (typeof this.source === 'string') {
      this.imageData = undefined;
    }
  }

  /**
   * Detects image extension from source
   */
  private detectExtension(): string {
    if (typeof this.source === 'string') {
      const match = this.source.match(/\.([a-z]+)$/i);
      if (match && match[1]) {
        return match[1].toLowerCase();
      }
    }
    // Default to png if can't detect
    return 'png';
  }

  /**
   * Attempts to detect image dimensions from buffer
   * Supports PNG, JPEG, GIF, BMP, and TIFF formats
   */
  private detectDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 24) {
      return null;
    }

    try {
      // PNG detection (signature: 89 50 4E 47)
      if (
        this.imageData[0] === 0x89 &&
        this.imageData[1] === 0x50 &&
        this.imageData[2] === 0x4e &&
        this.imageData[3] === 0x47
      ) {
        return this.detectPngDimensions();
      }

      // JPEG detection (signature: FF D8)
      if (this.imageData[0] === 0xff && this.imageData[1] === 0xd8) {
        return this.detectJpegDimensions();
      }

      // GIF detection (signature: "GIF87a" or "GIF89a")
      if (
        this.imageData[0] === 0x47 && // 'G'
        this.imageData[1] === 0x49 && // 'I'
        this.imageData[2] === 0x46    // 'F'
      ) {
        return this.detectGifDimensions();
      }

      // BMP detection (signature: "BM")
      if (this.imageData[0] === 0x42 && this.imageData[1] === 0x4d) {
        return this.detectBmpDimensions();
      }

      // TIFF detection (little-endian: "II*\0" or big-endian: "MM\0*")
      if (
        (this.imageData[0] === 0x49 && this.imageData[1] === 0x49 && this.imageData[2] === 0x2a) || // Little-endian
        (this.imageData[0] === 0x4d && this.imageData[1] === 0x4d && this.imageData[2] === 0x00)    // Big-endian
      ) {
        return this.detectTiffDimensions();
      }
    } catch (error) {
      // Dimension detection failed - will use defaults
    }

    return null;
  }

  /**
   * Detects dimensions from PNG image data
   */
  private detectPngDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 24) {
      return null;
    }

    try {
      const width = this.imageData.readUInt32BE(16);
      const height = this.imageData.readUInt32BE(20);

      // Convert pixels to EMUs (assuming 96 DPI)
      return {
        width: Math.round((width / 96) * 914400),
        height: Math.round((height / 96) * 914400),
      };
    } catch (error) {
      return null;
    }
  }

  /**
   * Detects dimensions from GIF image data
   * GIF format: width and height are at bytes 6-7 (width) and 8-9 (height), little-endian
   */
  private detectGifDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 10) {
      return null;
    }

    try {
      // GIF dimensions are stored as little-endian 16-bit integers
      const width = this.imageData.readUInt16LE(6);
      const height = this.imageData.readUInt16LE(8);

      // Validate dimensions
      if (width > 0 && height > 0 && width < 65535 && height < 65535) {
        // Convert pixels to EMUs (assuming 96 DPI)
        return {
          width: Math.round((width / 96) * 914400),
          height: Math.round((height / 96) * 914400),
        };
      }
    } catch (error) {
      // Detection failed
    }

    return null;
  }

  /**
   * Detects dimensions from BMP image data
   * BMP format: width at bytes 18-21, height at bytes 22-25, little-endian
   */
  private detectBmpDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 26) {
      return null;
    }

    try {
      // BMP dimensions are stored as little-endian 32-bit integers
      const width = this.imageData.readInt32LE(18);
      const height = Math.abs(this.imageData.readInt32LE(22)); // Height can be negative (top-down)

      // Validate dimensions
      if (width > 0 && height > 0 && width < 65535 && height < 65535) {
        // Convert pixels to EMUs (assuming 96 DPI)
        return {
          width: Math.round((width / 96) * 914400),
          height: Math.round((height / 96) * 914400),
        };
      }
    } catch (error) {
      // Detection failed
    }

    return null;
  }

  /**
   * Detects dimensions from TIFF image data
   * TIFF format is more complex - reads IFD (Image File Directory) entries
   */
  private detectTiffDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 14) {
      return null;
    }

    try {
      // Determine byte order
      const isLittleEndian = this.imageData[0] === 0x49; // 'II' = little-endian

      // Read IFD offset (bytes 4-7)
      const ifdOffset = isLittleEndian
        ? this.imageData.readUInt32LE(4)
        : this.imageData.readUInt32BE(4);

      if (ifdOffset + 14 > this.imageData.length) {
        return null;
      }

      // Read number of directory entries
      const numEntries = isLittleEndian
        ? this.imageData.readUInt16LE(ifdOffset)
        : this.imageData.readUInt16BE(ifdOffset);

      let width = 0;
      let height = 0;

      // Read IFD entries
      for (let i = 0; i < numEntries; i++) {
        const entryOffset = ifdOffset + 2 + i * 12;

        if (entryOffset + 12 > this.imageData.length) {
          break;
        }

        // Read tag ID
        const tag = isLittleEndian
          ? this.imageData.readUInt16LE(entryOffset)
          : this.imageData.readUInt16BE(entryOffset);

        // Read value
        const value = isLittleEndian
          ? this.imageData.readUInt32LE(entryOffset + 8)
          : this.imageData.readUInt32BE(entryOffset + 8);

        // Tag 256 (0x100) = ImageWidth
        // Tag 257 (0x101) = ImageHeight
        if (tag === 256 || tag === 0x100) {
          width = value;
        } else if (tag === 257 || tag === 0x101) {
          height = value;
        }

        // Stop if we have both dimensions
        if (width > 0 && height > 0) {
          break;
        }
      }

      // Validate dimensions
      if (width > 0 && height > 0 && width < 65535 && height < 65535) {
        // Convert pixels to EMUs (assuming 96 DPI)
        return {
          width: Math.round((width / 96) * 914400),
          height: Math.round((height / 96) * 914400),
        };
      }
    } catch (error) {
      // Detection failed
    }

    return null;
  }

  /**
   * Detects dimensions from JPEG image data
   * Handles baseline, progressive, and Exif JPEGs by parsing JPEG markers
   * @returns Dimensions in EMUs or null if detection fails
   */
  private detectJpegDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 12) {
      return null;
    }

    try {
      // Verify JPEG signature (0xFFD8)
      if (this.imageData[0] !== 0xff || this.imageData[1] !== 0xd8) {
        return null;
      }

      let offset = 2;

      // Parse JPEG markers
      while (offset < this.imageData.length - 1) {
        // All markers start with 0xFF
        if (this.imageData[offset] !== 0xff) {
          // Invalid marker - might be corrupted
          break;
        }

        const marker = this.imageData[offset + 1];

        // Ensure marker is defined
        if (marker === undefined) {
          break;
        }

        // Skip padding bytes (0xFF followed by 0x00 or 0xFF)
        if (marker === 0x00 || marker === 0xff) {
          offset++;
          continue;
        }

        // SOF (Start of Frame) markers contain dimension information
        // SOF0 (0xC0): Baseline DCT
        // SOF1 (0xC1): Extended sequential DCT
        // SOF2 (0xC2): Progressive DCT
        // SOF3 (0xC3): Lossless (sequential)
        // SOF5-7, SOF9-11, SOF13-15: Other SOF types
        const isSOF =
          (marker >= 0xc0 && marker <= 0xc3) ||
          (marker >= 0xc5 && marker <= 0xc7) ||
          (marker >= 0xc9 && marker <= 0xcb) ||
          (marker >= 0xcd && marker <= 0xcf);

        if (isSOF) {
          // SOF marker structure:
          // - 2 bytes: marker (0xFF, SOF marker)
          // - 2 bytes: length (including length itself)
          // - 1 byte: precision (bits per sample)
          // - 2 bytes: height
          // - 2 bytes: width

          if (offset + 9 > this.imageData.length) {
            // Not enough data
            break;
          }

          // Read height and width (big-endian)
          const height = this.imageData.readUInt16BE(offset + 5);
          const width = this.imageData.readUInt16BE(offset + 7);

          // Validate dimensions (sanity check)
          if (width > 0 && height > 0 && width < 65535 && height < 65535) {
            // Convert pixels to EMUs (assuming 96 DPI)
            // 1 inch = 96 pixels at 96 DPI
            // 1 inch = 914400 EMUs
            // Therefore: EMUs = pixels * (914400 / 96) = pixels * 9525
            return {
              width: Math.round((width / 96) * 914400),
              height: Math.round((height / 96) * 914400),
            };
          }
        }

        // SOS (Start of Scan) - marks the beginning of image data
        // After this, we won't find SOF markers
        if (marker === 0xda) {
          break;
        }

        // EOI (End of Image) - end of JPEG
        if (marker === 0xd9) {
          break;
        }

        // Standalone markers (no length field)
        // RST (Restart) markers: 0xD0-0xD7
        // SOI (Start of Image): 0xD8
        // EOI (End of Image): 0xD9
        // TEM (Temporary): 0x01
        const standaloneMarker =
          (marker >= 0xd0 && marker <= 0xd9) || marker === 0x01;

        if (standaloneMarker) {
          // Move to next marker (2 bytes)
          offset += 2;
          continue;
        }

        // For all other markers, read the length and skip the segment
        if (offset + 3 > this.imageData.length) {
          // Not enough data to read length
          break;
        }

        const segmentLength = this.imageData.readUInt16BE(offset + 2);

        // Validate segment length (sanity check)
        if (segmentLength < 2 || offset + 2 + segmentLength > this.imageData.length) {
          // Invalid segment length
          break;
        }

        // Move to next marker (2 bytes marker + segment length)
        offset += 2 + segmentLength;
      }
    } catch (error) {
      // Dimension detection failed - return null to use defaults
      // Silent failure is acceptable here - we'll use default dimensions
    }

    return null;
  }

  /**
   * Gets the image data buffer asynchronously
   * This is the preferred method for loading image data
   * @returns Promise<Buffer> containing the image data
   */
  async getImageDataAsync(): Promise<Buffer> {
    await this.ensureDataLoaded();

    if (this.imageData) {
      return this.imageData;
    }

    // Should not reach here after ensureDataLoaded()
    throw new Error('Failed to load image data');
  }

  /**
   * Gets the image data buffer synchronously
   * IMPORTANT: Only use this after calling ensureDataLoaded() or when the ImageManager
   * has already loaded all images via loadAllImageData()
   * @returns Buffer containing the image data
   * @throws {Error} If image data has not been loaded yet
   */
  getImageData(): Buffer {
    if (!this.imageData) {
      throw new Error(
        'Image data not loaded. ' +
        'Call await image.ensureDataLoaded() or await imageManager.loadAllImageData() first.'
      );
    }
    return this.imageData;
  }

  /**
   * Gets the image extension
   */
  getExtension(): string {
    return this.extension;
  }

  /**
   * Gets the image width in EMUs
   */
  getWidth(): number {
    return this.width;
  }

  /**
   * Gets the image height in EMUs
   */
  getHeight(): number {
    return this.height;
  }

  /**
   * Sets the image width in EMUs
   * @param width Width in EMUs
   * @param maintainAspectRatio Whether to adjust height proportionally
   */
  setWidth(width: number, maintainAspectRatio: boolean = true): this {
    if (maintainAspectRatio) {
      const ratio = this.height / this.width;
      this.height = Math.round(width * ratio);
    }
    this.width = width;
    return this;
  }

  /**
   * Sets the image height in EMUs
   * @param height Height in EMUs
   * @param maintainAspectRatio Whether to adjust width proportionally
   */
  setHeight(height: number, maintainAspectRatio: boolean = true): this {
    if (maintainAspectRatio) {
      const ratio = this.width / this.height;
      this.width = Math.round(height * ratio);
    }
    this.height = height;
    return this;
  }

  /**
   * Sets both width and height in EMUs
   * @param width Width in EMUs
   * @param height Height in EMUs
   */
  setSize(width: number, height: number): this {
    this.width = width;
    this.height = height;
    return this;
  }

  /**
   * Sets the relationship ID (used by ImageManager)
   * @param relationshipId Relationship ID
   */
  setRelationshipId(relationshipId: string): this {
    this.relationshipId = relationshipId;
    return this;
  }

  /**
   * Gets the relationship ID
   */
  getRelationshipId(): string | undefined {
    return this.relationshipId;
  }

  /**
   * Sets the docPr ID (drawing object ID)
   * @param id Document property ID
   */
  setDocPrId(id: number): this {
    this.docPrId = id;
    return this;
  }

  /**
   * Sets the alternative text (alt text) for the image
   * This is important for accessibility
   * @param altText Alternative text description
   * @returns This image for chaining
   */
  setAltText(altText: string): this {
    this.description = altText;
    return this;
  }

  /**
   * Gets the alternative text (alt text) for the image
   * @returns Alternative text description
   */
  getAltText(): string {
    return this.description;
  }

  /**
   * Sets image rotation in degrees
   * Note: This stores the rotation angle but doesn't actually rotate the image data
   * The rotation is applied via DrawingML transform in the XML
   * @param degrees Rotation angle in degrees (0-360)
   * @returns This image for chaining
   */
  rotate(degrees: number): this {
    // Normalize degrees to 0-360 range
    const normalizedDegrees = ((degrees % 360) + 360) % 360;

    // Store rotation for use in XML generation
    (this as any).rotation = normalizedDegrees;

    // If rotating 90 or 270 degrees, swap dimensions
    if (normalizedDegrees === 90 || normalizedDegrees === 270) {
      const temp = this.width;
      this.width = this.height;
      this.height = temp;
    }

    return this;
  }

  /**
   * Gets the rotation angle in degrees
   * @returns Rotation angle (0-360)
   */
  getRotation(): number {
    return (this as any).rotation || 0;
  }

  /**
   * Generates DrawingML XML for the image
   * This creates an inline image in the document
   */
  toXML(): XMLElement {
    if (!this.relationshipId) {
      throw new Error('Image must have a relationship ID before generating XML');
    }

    // Create the drawing structure
    const drawing = XMLBuilder.w('drawing', undefined, [
      this.createInline(),
    ]);

    return drawing;
  }

  /**
   * Creates the wp:inline element
   */
  private createInline(): XMLElement {
    const children: XMLElement[] = [];

    // Extent (size)
    children.push({
      name: 'wp:extent',
      attributes: {
        cx: this.width.toString(),
        cy: this.height.toString(),
      },
      selfClosing: true,
    });

    // Effect extent (for shadows, etc.)
    children.push({
      name: 'wp:effectExtent',
      attributes: {
        l: '0',
        t: '0',
        r: '0',
        b: '0',
      },
      selfClosing: true,
    });

    // Document properties
    children.push({
      name: 'wp:docPr',
      attributes: {
        id: this.docPrId.toString(),
        name: this.name,
        descr: this.description,
      },
      selfClosing: true,
    });

    // Non-visual picture properties
    children.push({
      name: 'wp:cNvGraphicFramePr',
      children: [
        {
          name: 'a:graphicFrameLocks',
          attributes: {
            'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            noChangeAspect: '1',
          },
          selfClosing: true,
        },
      ],
    });

    // Graphic (the actual image)
    children.push(this.createGraphic());

    return {
      name: 'wp:inline',
      attributes: {
        distT: '0',
        distB: '0',
        distL: '0',
        distR: '0',
      },
      children,
    };
  }

  /**
   * Creates the a:graphic element
   */
  private createGraphic(): XMLElement {
    return {
      name: 'a:graphic',
      attributes: {
        'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      },
      children: [
        {
          name: 'a:graphicData',
          attributes: {
            uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
          },
          children: [this.createPicture()],
        },
      ],
    };
  }

  /**
   * Creates the pic:pic element
   */
  private createPicture(): XMLElement {
    return {
      name: 'pic:pic',
      attributes: {
        'xmlns:pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
      },
      children: [
        // Non-visual picture properties
        {
          name: 'pic:nvPicPr',
          children: [
            {
              name: 'pic:cNvPr',
              attributes: {
                id: this.docPrId.toString(),
                name: this.name,
                descr: this.description,
              },
              selfClosing: true,
            },
            {
              name: 'pic:cNvPicPr',
              selfClosing: true,
            },
          ],
        },
        // Blip fill (reference to image)
        {
          name: 'pic:blipFill',
          children: [
            {
              name: 'a:blip',
              attributes: {
                'r:embed': this.relationshipId!,
              },
              selfClosing: true,
            },
            {
              name: 'a:stretch',
              children: [
                {
                  name: 'a:fillRect',
                  selfClosing: true,
                },
              ],
            },
          ],
        },
        // Shape properties (size and position)
        {
          name: 'pic:spPr',
          children: [
            {
              name: 'a:xfrm',
              attributes: this.getRotation() > 0 ? { rot: (this.getRotation() * 60000).toString() } : undefined,
              children: [
                {
                  name: 'a:off',
                  attributes: {
                    x: '0',
                    y: '0',
                  },
                  selfClosing: true,
                },
                {
                  name: 'a:ext',
                  attributes: {
                    cx: this.width.toString(),
                    cy: this.height.toString(),
                  },
                  selfClosing: true,
                },
              ],
            },
            {
              name: 'a:prstGeom',
              attributes: {
                prst: 'rect',
              },
              children: [
                {
                  name: 'a:avLst',
                  selfClosing: true,
                },
              ],
            },
          ],
        },
      ],
    };
  }

  /**
   * Creates an image from a file path
   * Async method that loads and detects dimensions
   * @param filePath Path to image file
   * @param width Optional width in EMUs (overrides detection)
   * @param height Optional height in EMUs (overrides detection)
   */
  static async fromFile(filePath: string, width?: number, height?: number): Promise<Image> {
    const image = new Image({
      source: filePath,
      width,
      height,
      name: filePath.split(/[/\\]/).pop() || 'image',
    });

    // Load dimensions if not provided
    if (!width || !height) {
      try {
        await image.loadImageDataForDimensions();
      } catch (error) {
        // Dimension detection failed, use defaults
        // Error will be thrown later if file truly doesn't exist
      }
    }

    return image;
  }

  /**
   * Creates an image from a buffer
   * Detects dimensions from buffer data
   * @param buffer Image buffer
   * @param extension Image file extension
   * @param width Optional width in EMUs (overrides detection)
   * @param height Optional height in EMUs (overrides detection)
   */
  static async fromBuffer(
    buffer: Buffer,
    extension: string,
    width?: number,
    height?: number
  ): Promise<Image> {
    const image = new Image({
      source: buffer,
      width,
      height,
      name: `image.${extension}`,
    });
    image.extension = extension;

    // Load dimensions if not provided
    if (!width || !height) {
      await image.loadImageDataForDimensions();
    }

    return image;
  }

  /**
   * Factory method for creating an image
   * Async method that loads and detects dimensions
   * @param properties Image properties
   */
  static async create(properties: ImageProperties): Promise<Image> {
    const image = new Image(properties);

    // Load dimensions if not provided and source is a file
    if ((!properties.width || !properties.height) && typeof properties.source === 'string') {
      try {
        await image.loadImageDataForDimensions();
      } catch (error) {
        // Dimension detection failed, use defaults
      }
    } else if ((!properties.width || !properties.height) && Buffer.isBuffer(properties.source)) {
      await image.loadImageDataForDimensions();
    }

    return image;
  }
}
