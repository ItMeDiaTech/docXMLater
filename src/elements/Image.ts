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
   * Basic detection for PNG and JPEG
   */
  private detectDimensions(): { width: number; height: number } | null {
    if (!this.imageData || this.imageData.length < 24) {
      return null;
    }

    try {
      // PNG detection
      if (
        this.imageData[0] === 0x89 &&
        this.imageData[1] === 0x50 &&
        this.imageData[2] === 0x4e &&
        this.imageData[3] === 0x47
      ) {
        const width = this.imageData.readUInt32BE(16);
        const height = this.imageData.readUInt32BE(20);
        // Convert pixels to EMUs (assuming 96 DPI)
        return {
          width: Math.round((width / 96) * 914400),
          height: Math.round((height / 96) * 914400),
        };
      }

      // JPEG detection
      if (this.imageData[0] === 0xff && this.imageData[1] === 0xd8) {
        const jpegDims = this.detectJpegDimensions();
        if (jpegDims) {
          return jpegDims;
        }
      }
    } catch (error) {
      // Dimension detection failed
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
   * Gets the image data buffer
   * Ensures data is loaded before returning
   * @deprecated Use async pattern: await image.ensureDataLoaded() then access imageData
   */
  getImageData(): Buffer {
    if (!this.imageData) {
      throw new Error('Image data not loaded. Call await ensureDataLoaded() first.');
    }
    return this.imageData;
  }

  /**
   * Gets the image data buffer asynchronously
   * Preferred method - ensures data is loaded before returning
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
