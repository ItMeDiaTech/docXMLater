/**
 * ImageManager - Manages images in a document
 *
 * Handles image tracking, unique filename generation, and coordination
 * with the RelationshipManager for image relationships.
 */

import { Image } from './Image';

/**
 * Image file entry
 */
interface ImageEntry {
  /** The Image object */
  image: Image;
  /** Filename in media folder (e.g., 'image1.png') */
  filename: string;
  /** Relationship ID */
  relationshipId: string;
  /** Document property ID (for DrawingML) */
  docPrId: number;
}

/**
 * Manages all images in a document
 */
export class ImageManager {
  private images: Map<Image, ImageEntry>;
  private nextImageNumber: number;
  private nextDocPrId: number;

  /**
   * Creates a new image manager
   */
  constructor() {
    this.images = new Map();
    this.nextImageNumber = 1;
    this.nextDocPrId = 1;
  }

  /**
   * Registers an image with the manager
   * @param image The image to register
   * @param relationshipId The relationship ID for this image
   * @returns The filename assigned to this image
   */
  registerImage(image: Image, relationshipId: string): string {
    // Check if already registered
    const existing = this.images.get(image);
    if (existing) {
      return existing.filename;
    }

    // Generate unique filename
    const extension = image.getExtension();
    const filename = `image${this.nextImageNumber++}.${extension}`;

    // Assign docPr ID
    const docPrId = this.nextDocPrId++;
    image.setDocPrId(docPrId);

    // Set relationship ID
    image.setRelationshipId(relationshipId);

    // Store entry
    const entry: ImageEntry = {
      image,
      filename,
      relationshipId,
      docPrId,
    };

    this.images.set(image, entry);

    return filename;
  }

  /**
   * Gets the filename for an image
   * @param image The image
   * @returns The filename, or undefined if not registered
   */
  getFilename(image: Image): string | undefined {
    return this.images.get(image)?.filename;
  }

  /**
   * Gets the relationship ID for an image
   * @param image The image
   * @returns The relationship ID, or undefined if not registered
   */
  getRelationshipId(image: Image): string | undefined {
    return this.images.get(image)?.relationshipId;
  }

  /**
   * Gets all registered images
   * @returns Array of image entries
   */
  getAllImages(): ImageEntry[] {
    return Array.from(this.images.values());
  }

  /**
   * Gets the number of images
   * @returns Number of registered images
   */
  getImageCount(): number {
    return this.images.size;
  }

  /**
   * Checks if an image is registered
   * @param image The image
   * @returns True if registered
   */
  hasImage(image: Image): boolean {
    return this.images.has(image);
  }

  /**
   * Gets the MIME type for an image extension
   * @param extension File extension (without dot)
   * @returns MIME type
   */
  static getMimeType(extension: string): string {
    const mimeTypes: Record<string, string> = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      tiff: 'image/tiff',
      tif: 'image/tiff',
    };

    return mimeTypes[extension.toLowerCase()] || 'image/png';
  }

  /**
   * Loads data for all images
   * Call this before saving to ensure all image data is available
   */
  async loadAllImageData(): Promise<void> {
    // Load all images in parallel for better performance
    const loadPromises = Array.from(this.images.values()).map(entry =>
      entry.image.ensureDataLoaded()
    );
    await Promise.all(loadPromises);
  }

  /**
   * Releases data for all images
   * Call this after saving to free memory
   */
  releaseAllImageData(): void {
    for (const entry of this.images.values()) {
      entry.image.releaseData();
    }
  }

  /**
   * Gets the total size of all loaded image data
   * Only counts images that are currently loaded in memory
   * @returns Total size in bytes
   */
  getTotalSize(): number {
    let totalSize = 0;
    for (const entry of this.images.values()) {
      try {
        const data = entry.image.getImageData();
        totalSize += data.length;
      } catch {
        // Image not loaded - skip (don't count unloaded images)
      }
    }
    return totalSize;
  }

  /**
   * Gets the total size of all images (loads them if needed)
   * Use this for accurate size estimation before saving
   * @returns Total size in bytes
   */
  async getTotalSizeAsync(): Promise<number> {
    let totalSize = 0;
    for (const entry of this.images.values()) {
      try {
        const data = await entry.image.getImageDataAsync();
        totalSize += data.length;
      } catch {
        // Image loading failed - skip
      }
    }
    return totalSize;
  }

  /**
   * Gets statistics about images
   * @returns Object with image statistics
   */
  getStats(): {
    count: number;
    totalSize: number;
    averageSize: number;
  } {
    const count = this.getImageCount();
    const totalSize = this.getTotalSize();
    return {
      count,
      totalSize,
      averageSize: count > 0 ? Math.round(totalSize / count) : 0,
    };
  }

  /**
   * Clears all images
   */
  clear(): this {
    this.images.clear();
    this.nextImageNumber = 1;
    this.nextDocPrId = 1;
    return this;
  }

  /**
   * Creates a new image manager
   */
  static create(): ImageManager {
    return new ImageManager();
  }
}
