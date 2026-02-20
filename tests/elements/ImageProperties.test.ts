/**
 * Tests for Image Properties (Phase 4.4)
 *
 * Tests advanced image formatting properties including:
 * - Effect extent (shadows, glows)
 * - Text wrapping
 * - Positioning (floating images)
 * - Anchor configuration
 * - Cropping
 * - Visual effects
 */

import { Document } from '../../src/core/Document';
import { Image } from '../../src/elements/Image';
import { ImageRun } from '../../src/elements/ImageRun';
import { Table } from '../../src/elements/Table';
import { ImageManager } from '../../src/elements/ImageManager';
import { promises as fs } from 'fs';
import { join } from 'path';

// Test image path
const TEST_IMAGE_PATH = join(__dirname, '..', 'fixtures', 'test-image.png');
const OUTPUT_DIR = join(__dirname, '..', 'output');

/**
 * Helper to create a test image buffer (1x1 PNG)
 */
function createTestImageBuffer(): Buffer {
  // 1x1 transparent PNG
  return Buffer.from([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
    0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
    0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
    0x42, 0x60, 0x82,
  ]);
}

describe('Image Properties - Effect Extent', () => {
  it('should set and get effect extent for shadows', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    // Set effect extent for shadow
    image.setEffectExtent(25400, 25400, 25400, 25400); // 0.25 inches on all sides

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    // Save and reload
    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-effect-extent.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images.length).toBe(1);
    const extent = images[0]?.image.getEffectExtent();
    expect(extent).toBeDefined();
    expect(extent!.left).toBe(25400);
    expect(extent!.top).toBe(25400);
    expect(extent!.right).toBe(25400);
    expect(extent!.bottom).toBe(25400);
  });

  it('should handle zero effect extent', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setEffectExtent(0, 0, 0, 0);

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    const extent = images[0]?.image.getEffectExtent();
    expect(extent).toBeDefined();
    expect(extent!.left).toBe(0);
  });
});

describe('Image Properties - Text Wrapping', () => {
  it('should set square wrap with both sides', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setWrap('square', 'bothSides', {
      top: 10000,
      bottom: 10000,
      left: 10000,
      right: 10000,
    });

    // Make it floating
    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });

    image.setPosition(
      { anchor: 'page', offset: 914400 },
      { anchor: 'page', offset: 914400 }
    );

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-wrap-square.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    const wrap = images[0]?.image.getWrap();
    expect(wrap).toBeDefined();
    expect(wrap!.type).toBe('square');
    expect(wrap!.side).toBe('bothSides');
    expect(wrap!.distanceTop).toBe(10000);
  });

  it('should set tight wrap with left side', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setWrap('tight', 'left');
    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });
    image.setPosition(
      { anchor: 'page', alignment: 'left' },
      { anchor: 'page', alignment: 'top' }
    );

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-wrap-tight.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    const wrap = images[0]?.image.getWrap();
    expect(wrap!.type).toBe('tight');
    expect(wrap!.side).toBe('left');
  });

  it('should set top and bottom wrap', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setWrap('topAndBottom');
    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });
    image.setPosition(
      { anchor: 'page', alignment: 'center' },
      { anchor: 'page', alignment: 'center' }
    );

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-wrap-topbottom.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images[0]?.image.getWrap()!.type).toBe('topAndBottom');
  });
});

describe('Image Properties - Positioning', () => {
  it('should set absolute positioning with offset', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 1828800, 1828800); // 2 inches

    image.setPosition(
      { anchor: 'page', offset: 1828800 }, // 2 inches from left
      { anchor: 'page', offset: 1828800 }  // 2 inches from top
    );

    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-position-absolute.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    const position = images[0]?.image.getPosition();
    expect(position).toBeDefined();
    expect(position!.horizontal.anchor).toBe('page');
    expect(position!.horizontal.offset).toBe(1828800);
    expect(position!.vertical.offset).toBe(1828800);
  });

  it('should set relative positioning with alignment', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setPosition(
      { anchor: 'margin', alignment: 'center' },
      { anchor: 'margin', alignment: 'center' }
    );

    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-position-relative.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    const position = images[0]?.image.getPosition();
    expect(position!.horizontal.alignment).toBe('center');
    expect(position!.vertical.alignment).toBe('center');
    expect(position!.horizontal.anchor).toBe('margin');
  });

  it('should anchor to column', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setPosition(
      { anchor: 'column', alignment: 'right' },
      { anchor: 'paragraph', alignment: 'top' }
    );

    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-position-column.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images[0]?.image.getPosition()!.horizontal.anchor).toBe('column');
    expect(images[0]?.image.getPosition()!.vertical.anchor).toBe('paragraph');
  });
});

describe('Image Properties - Anchor Configuration', () => {
  it('should set floating image behind text', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setAnchor({
      behindDoc: true,
      locked: false,
      layoutInCell: true,
      allowOverlap: true,
      relativeHeight: 251658240,
    });

    image.setPosition(
      { anchor: 'page', alignment: 'center' },
      { anchor: 'page', alignment: 'center' }
    );

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-anchor-behind.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    const anchor = images[0]?.image.getAnchor();
    expect(anchor).toBeDefined();
    expect(anchor!.behindDoc).toBe(true);
    expect(anchor!.allowOverlap).toBe(true);
  });

  it('should set locked floating image', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setAnchor({
      behindDoc: false,
      locked: true,
      layoutInCell: false,
      allowOverlap: false,
      relativeHeight: 500000000,
    });

    image.setPosition(
      { anchor: 'page', offset: 914400 },
      { anchor: 'page', offset: 914400 }
    );

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-anchor-locked.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images[0]?.image.getAnchor()!.locked).toBe(true);
    expect(images[0]?.image.getAnchor()!.layoutInCell).toBe(false);
  });
});

describe('Image Properties - Cropping', () => {
  it('should set image crop on all sides', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 1828800, 1828800);

    // Crop 10% from each side
    image.setCrop(10, 10, 10, 10);

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-crop.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    const crop = images[0]?.image.getCrop();
    expect(crop).toBeDefined();
    expect(crop!.left).toBe(10);
    expect(crop!.top).toBe(10);
    expect(crop!.right).toBe(10);
    expect(crop!.bottom).toBe(10);
  });

  it('should clamp crop values to 0-100', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    // Try to set invalid crop values
    image.setCrop(-10, 150, 50, 75);

    const crop = image.getCrop();
    expect(crop!.left).toBe(0);   // Clamped to 0
    expect(crop!.top).toBe(100);  // Clamped to 100
    expect(crop!.right).toBe(50);
    expect(crop!.bottom).toBe(75);
  });
});

describe('Image Properties - Visual Effects', () => {
  it('should set brightness and contrast', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setEffects({
      brightness: 25,
      contrast: -15,
    });

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-effects-brightness.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    const effects = images[0]?.image.getEffects();
    expect(effects).toBeDefined();
    expect(effects!.brightness).toBe(25);
    expect(effects!.contrast).toBe(-15);
  });

  it('should set grayscale effect', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setEffects({
      grayscale: true,
    });

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-effects-grayscale.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images[0]?.image.getEffects()!.grayscale).toBe(true);
  });
});

describe('Image Properties - Combined Properties', () => {
  it('should handle multiple properties together', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 1828800, 1828800);

    // Set all properties
    image.setEffectExtent(25400, 25400, 25400, 25400);
    image.setWrap('square', 'bothSides', { top: 10000, bottom: 10000, left: 10000, right: 10000 });
    image.setPosition(
      { anchor: 'page', offset: 914400 },
      { anchor: 'page', offset: 914400 }
    );
    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });
    image.setCrop(5, 5, 5, 5);
    image.setEffects({ brightness: 10, contrast: 5 });

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-combined.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images[0]?.image.getEffectExtent()).toBeDefined();
    expect(images[0]?.image.getWrap()).toBeDefined();
    expect(images[0]?.image.getPosition()).toBeDefined();
    expect(images[0]?.image.getAnchor()).toBeDefined();
    expect(images[0]?.image.getCrop()).toBeDefined();
    expect(images[0]?.image.getEffects()).toBeDefined();
  });

  it('should preserve all properties through multiple save/load cycles', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setEffectExtent(12700, 12700, 12700, 12700);
    image.setCrop(15, 15, 15, 15);

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    // First cycle
    const buffer1 = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer1);

    // Second cycle
    const buffer2 = await doc2.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-multicycle.docx'), buffer2);
    const doc3 = await Document.loadFromBuffer(buffer2);

    const images = doc3.getImages();
    expect(images[0]?.image.getEffectExtent()!.left).toBe(12700);
    expect(images[0]?.image.getCrop()!.left).toBe(15);
  });
});

describe('Image Properties - Inline vs Floating', () => {
  it('should correctly identify inline images', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    // No anchor or position = inline
    expect(image.isFloating()).toBe(false);

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);

    expect(doc2.getImages()[0]?.image.isFloating()).toBe(false);
  });

  it('should correctly identify floating images', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });

    image.setPosition(
      { anchor: 'page', alignment: 'center' },
      { anchor: 'page', alignment: 'center' }
    );

    expect(image.isFloating()).toBe(true);

    // Add image to document (registers it with ImageManager)
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);

    expect(doc2.getImages()[0]?.image.isFloating()).toBe(true);
  });
});

describe('Image Properties - Rotation', () => {
  it('should preserve rotation through save/load cycle', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    // Set rotation to 90 degrees
    image.rotate(90);
    expect(image.getRotation()).toBe(90);

    // Add image to document
    doc.addImage(image);

    // Save and reload
    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-rotation.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images.length).toBe(1);
    expect(images[0]?.image.getRotation()).toBe(90);
  });

  it('should preserve 180 degree rotation', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.rotate(180);

    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images[0]?.image.getRotation()).toBe(180);
  });

  it('should preserve 270 degree rotation', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    image.rotate(270);

    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images[0]?.image.getRotation()).toBe(270);
  });

  it('should handle zero rotation (no attribute)', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);

    // Don't set rotation - should remain 0
    expect(image.getRotation()).toBe(0);

    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images[0]?.image.getRotation()).toBe(0);
  });
});

describe('Image Properties - Edge Cases', () => {
  it('should handle images with fractional rotation', async () => {
    // Test that fractional rotation values (e.g., 45 degrees) work correctly
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 457200, 457200);

    // Set a non-90-degree rotation
    image.rotate(45);
    expect(image.getRotation()).toBe(45);

    doc.addImage(image);

    // Save and reload
    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-image-rotation-45.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    // Rotation should be preserved
    expect(images.length).toBe(1);
    expect(images[0]?.image.getRotation()).toBe(45);
  });

  it('should handle multiple images with different properties', async () => {
    const doc = Document.create();

    // Image 1: inline with effect extent
    const image1 = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image1.setEffectExtent(25400, 25400, 25400, 25400);
    doc.addImage(image1);

    // Image 2: rotated
    const image2 = await Image.fromBuffer(createTestImageBuffer(), 'png', 457200, 457200);
    image2.rotate(45);
    doc.addImage(image2);

    // Image 3: floating with wrap
    const image3 = await Image.fromBuffer(createTestImageBuffer(), 'png', 685800, 685800);
    image3.setWrap('square', 'bothSides');
    image3.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });
    image3.setPosition(
      { anchor: 'page', offset: 914400 },
      { anchor: 'page', offset: 914400 }
    );
    doc.addImage(image3);

    // Save and reload
    const buffer = await doc.toBuffer();
    await fs.writeFile(join(OUTPUT_DIR, 'test-multiple-images.docx'), buffer);

    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();

    expect(images.length).toBe(3);

    // Verify image 1 effect extent
    const extent1 = images[0]?.image.getEffectExtent();
    expect(extent1?.left).toBe(25400);

    // Verify image 2 rotation
    expect(images[1]?.image.getRotation()).toBe(45);

    // Verify image 3 wrap
    const wrap3 = images[2]?.image.getWrap();
    expect(wrap3?.type).toBe('square');
  });
});

describe('Image Properties - Flip (flipH/flipV)', () => {
  it('should default flipH and flipV to false', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    expect(image.getFlipH()).toBe(false);
    expect(image.getFlipV()).toBe(false);
  });

  it('should set and get flipH', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setFlipH(true);
    expect(image.getFlipH()).toBe(true);
    image.setFlipH(false);
    expect(image.getFlipH()).toBe(false);
  });

  it('should set and get flipV', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setFlipV(true);
    expect(image.getFlipV()).toBe(true);
  });

  it('should support fluent API for flip setters', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    const result = image.setFlipH(true).setFlipV(true);
    expect(result).toBe(image);
    expect(image.getFlipH()).toBe(true);
    expect(image.getFlipV()).toBe(true);
  });

  it('should include flipH in XML output when true', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setFlipH(true);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    const images = reloaded.getImages();
    expect(images.length).toBe(1);
    expect(images[0]?.image.getFlipH()).toBe(true);
    expect(images[0]?.image.getFlipV()).toBe(false);
    reloaded.dispose();
  });

  it('should include flipV in XML output when true', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setFlipV(true);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    const images = reloaded.getImages();
    expect(images.length).toBe(1);
    expect(images[0]?.image.getFlipH()).toBe(false);
    expect(images[0]?.image.getFlipV()).toBe(true);
    reloaded.dispose();
  });

  it('should preserve both flipH and flipV during round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setFlipH(true);
    image.setFlipV(true);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    const images = reloaded.getImages();
    expect(images.length).toBe(1);
    expect(images[0]?.image.getFlipH()).toBe(true);
    expect(images[0]?.image.getFlipV()).toBe(true);
    reloaded.dispose();
  });

  it('should not add flip attributes to XML when false', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    // No flip set — defaults to false
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    const images = reloaded.getImages();
    expect(images.length).toBe(1);
    expect(images[0]?.image.getFlipH()).toBe(false);
    expect(images[0]?.image.getFlipV()).toBe(false);
    reloaded.dispose();
  });

  it('should preserve flipH with rotation during round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.rotate(45);
    image.setFlipH(true);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    const images = reloaded.getImages();
    expect(images.length).toBe(1);
    expect(images[0]?.image.getRotation()).toBe(45);
    expect(images[0]?.image.getFlipH()).toBe(true);
    expect(images[0]?.image.getFlipV()).toBe(false);
    reloaded.dispose();
  });

  it('should initialize flipH/flipV from ImageProperties', async () => {
    const image = await Image.create({
      source: createTestImageBuffer(),
      width: 914400,
      height: 914400,
      flipH: true,
      flipV: true,
    });
    expect(image.getFlipH()).toBe(true);
    expect(image.getFlipV()).toBe(true);
  });
});

// ============================================================================
// Group A: Simple Attribute Preservation
// ============================================================================

describe('Image Properties - Group A: Simple Attributes', () => {
  it('should preserve presetGeometry through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setPresetGeometry('ellipse');
    expect(image.getPresetGeometry()).toBe('ellipse');
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const images = doc2.getImages();
    expect(images[0]?.image.getPresetGeometry()).toBe('ellipse');
    doc2.dispose();
  });

  it('should preserve compressionState through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setCompressionState('print');
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    expect(doc2.getImages()[0]?.image.getCompressionState()).toBe('print');
    doc2.dispose();
  });

  it('should preserve bwMode through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setBwMode('clr');
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    expect(doc2.getImages()[0]?.image.getBwMode()).toBe('clr');
    doc2.dispose();
  });

  it('should preserve inline dist attributes through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setInlineDist(1000, 2000, 3000, 4000);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const img = doc2.getImages()[0]!.image;
    expect(img.getInlineDistT()).toBe(1000);
    expect(img.getInlineDistB()).toBe(2000);
    expect(img.getInlineDistL()).toBe(3000);
    expect(img.getInlineDistR()).toBe(4000);
    doc2.dispose();
  });

  it('should preserve picNonVisualProps through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setPicNonVisualProps({ id: '5', name: 'test-pic', descr: 'My pic' });
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const props = doc2.getImages()[0]!.image.getPicNonVisualProps();
    expect(props.id).toBe('5');
    expect(props.name).toBe('test-pic');
    expect(props.descr).toBe('My pic');
    doc2.dispose();
  });

  it('should preserve picLocks through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setPicLocks({ noChangeAspect: true, noCrop: true, noRot: true });
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const locks = doc2.getImages()[0]!.image.getPicLocks();
    expect(locks.noChangeAspect).toBe(true);
    expect(locks.noCrop).toBe(true);
    expect(locks.noRot).toBe(true);
    expect(locks.noMove).toBeUndefined();
    doc2.dispose();
  });

  it('should preserve hidden attribute through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setHidden(true);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    expect(doc2.getImages()[0]?.image.getHidden()).toBe(true);
    doc2.dispose();
  });

  it('should preserve blipFillDpi through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setBlipFillDpi(150);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    expect(doc2.getImages()[0]?.image.getBlipFillDpi()).toBe(150);
    doc2.dispose();
  });

  it('should preserve blipFillRotWithShape through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setBlipFillRotWithShape(true);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    expect(doc2.getImages()[0]?.image.getBlipFillRotWithShape()).toBe(true);
    doc2.dispose();
  });

  it('should preserve transparency effect through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setEffects({ transparency: 50 });
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    expect(doc2.getImages()[0]?.image.getEffects()?.transparency).toBe(50);
    doc2.dispose();
  });

  it('should preserve noChangeAspect=false through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setNoChangeAspect(false);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    // noChangeAspect false means the attribute is "0"
    expect(doc2.getImages()[0]?.image.getNoChangeAspect()).toBe(false);
    doc2.dispose();
  });

  it('should default presetGeometry to rect', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    expect(image.getPresetGeometry()).toBe('rect');
  });

  it('should default compressionState to none', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    expect(image.getCompressionState()).toBe('none');
  });

  it('should default bwMode to auto', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    expect(image.getBwMode()).toBe('auto');
  });

  it('should initialize all Group A properties from ImageProperties', async () => {
    const image = await Image.create({
      source: createTestImageBuffer(),
      width: 914400,
      height: 914400,
      presetGeometry: 'roundRect',
      compressionState: 'email',
      bwMode: 'gray',
      inlineDistT: 100,
      inlineDistB: 200,
      inlineDistL: 300,
      inlineDistR: 400,
      noChangeAspect: false,
      hidden: true,
      blipFillDpi: 300,
      blipFillRotWithShape: false,
      picLocks: { noChangeAspect: true, noGrp: true },
      picNonVisualProps: { id: '10', name: 'myPic', descr: 'desc' },
    });
    expect(image.getPresetGeometry()).toBe('roundRect');
    expect(image.getCompressionState()).toBe('email');
    expect(image.getBwMode()).toBe('gray');
    expect(image.getInlineDistT()).toBe(100);
    expect(image.getNoChangeAspect()).toBe(false);
    expect(image.getHidden()).toBe(true);
    expect(image.getBlipFillDpi()).toBe(300);
    expect(image.getBlipFillRotWithShape()).toBe(false);
    expect(image.getPicLocks().noGrp).toBe(true);
    expect(image.getPicNonVisualProps().id).toBe('10');
  });
});

// ============================================================================
// Group B: Raw XML Passthrough
// ============================================================================

describe('Image Properties - Group B: Raw Passthrough', () => {
  it('should store and retrieve raw passthrough slots', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image._setRawPassthrough('blip-effects', '<a:clrChange/>');
    expect(image._getRawPassthrough('blip-effects')).toBe('<a:clrChange/>');
    expect(image._hasRawPassthrough('blip-effects')).toBe(true);
    expect(image._hasRawPassthrough('nonexistent')).toBe(false);
  });

  it('should include blip-effects passthrough in XML output', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image._setRawPassthrough('blip-effects', '<a:biLevel thresh="50000"/>');
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    // The biLevel element should be preserved as raw passthrough
    const img = doc2.getImages()[0]!.image;
    expect(img._getRawPassthrough('blip-effects')).toContain('a:biLevel');
    doc2.dispose();
  });

  it('should include spPr-effects passthrough in XML output', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image._setRawPassthrough('spPr-effects', '<a:effectLst><a:outerShdw dist="50000" dir="5400000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst>');
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const img = doc2.getImages()[0]!.image;
    expect(img._getRawPassthrough('spPr-effects')).toContain('a:effectLst');
    doc2.dispose();
  });
});

// ============================================================================
// Group C: Enhanced Border Model
// ============================================================================

describe('Image Properties - Group C: Enhanced Border', () => {
  it('should accept number for backward-compatible setBorder()', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setBorder(3);
    const border = image.getBorder();
    expect(border).toBeDefined();
    expect(border!.width).toBe(3);
  });

  it('should accept full ImageBorder object', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setBorder({
      width: 2,
      cap: 'rnd',
      compound: 'dbl',
      fill: { type: 'srgbClr', value: 'FF0000' },
      dashPattern: 'dash',
      join: 'round',
    });
    const border = image.getBorder();
    expect(border!.cap).toBe('rnd');
    expect(border!.compound).toBe('dbl');
    expect(border!.fill?.value).toBe('FF0000');
    expect(border!.dashPattern).toBe('dash');
    expect(border!.join).toBe('round');
  });

  it('should preserve basic border through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setBorder(4);
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const border = doc2.getImages()[0]?.image.getBorder();
    expect(border).toBeDefined();
    expect(border!.width).toBe(4);
    doc2.dispose();
  });

  it('should preserve enhanced border attributes through round-trip', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setBorder({
      width: 3,
      cap: 'sq',
      dashPattern: 'lgDash',
      join: 'bevel',
      fill: { type: 'srgbClr', value: 'FF0000' },
    });
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const border = doc2.getImages()[0]?.image.getBorder();
    expect(border).toBeDefined();
    expect(border!.width).toBe(3);
    expect(border!.cap).toBe('sq');
    expect(border!.dashPattern).toBe('lgDash');
    expect(border!.join).toBe('bevel');
    expect(border!.fill?.type).toBe('srgbClr');
    expect(border!.fill?.value).toBe('FF0000');
    doc2.dispose();
  });

  it('should preserve border with scheme color and modifiers', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setBorder({
      width: 2,
      fill: {
        type: 'schemeClr',
        value: 'accent1',
        modifiers: [{ name: 'lumMod', val: '75000' }],
      },
    });
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    const border = doc2.getImages()[0]?.image.getBorder();
    expect(border!.fill?.type).toBe('schemeClr');
    expect(border!.fill?.value).toBe('accent1');
    expect(border!.fill?.modifiers).toBeDefined();
    expect(border!.fill?.modifiers![0]?.name).toBe('lumMod');
    expect(border!.fill?.modifiers![0]?.val).toBe('75000');
    doc2.dispose();
  });
});

// ============================================================================
// Group D: Image Format Support
// ============================================================================

describe('Image Properties - Group D: Format Detection', () => {
  it('should detect PNG format from buffer', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer());
    expect(image.getExtension()).toBe('png');
  });

  it('should detect SVG format from buffer', async () => {
    const svgBuffer = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100"/></svg>');
    const image = await Image.fromBuffer(svgBuffer);
    expect(image.getExtension()).toBe('svg');
  });

  it('should detect EMF format from buffer', async () => {
    // Create minimal EMF buffer with signature at offset 40
    const emfBuffer = Buffer.alloc(48);
    emfBuffer[0] = 0x01; // Record type
    // ' EMF' at offset 40
    emfBuffer[40] = 0x20;
    emfBuffer[41] = 0x45;
    emfBuffer[42] = 0x4D;
    emfBuffer[43] = 0x46;
    const image = await Image.fromBuffer(emfBuffer);
    expect(image.getExtension()).toBe('emf');
  });

  it('should detect WMF placeable format from buffer', async () => {
    const wmfBuffer = Buffer.alloc(22);
    wmfBuffer[0] = 0xD7;
    wmfBuffer[1] = 0xCD;
    wmfBuffer[2] = 0xC6;
    wmfBuffer[3] = 0x9A;
    const image = await Image.fromBuffer(wmfBuffer);
    expect(image.getExtension()).toBe('wmf');
  });

  it('should validate SVG image data', async () => {
    const svgBuffer = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"></svg>');
    const image = await Image.fromBuffer(svgBuffer);
    await image.ensureDataLoaded();
    expect(image.validateImageData().valid).toBe(true);
  });

  it('should validate EMF image data', async () => {
    const emfBuffer = Buffer.alloc(48);
    emfBuffer[0] = 0x01;
    emfBuffer[40] = 0x20;
    emfBuffer[41] = 0x45;
    emfBuffer[42] = 0x4D;
    emfBuffer[43] = 0x46;
    const image = await Image.fromBuffer(emfBuffer);
    await image.ensureDataLoaded();
    expect(image.validateImageData().valid).toBe(true);
  });

  it('should validate WMF image data', async () => {
    const wmfBuffer = Buffer.alloc(22);
    wmfBuffer[0] = 0xD7;
    wmfBuffer[1] = 0xCD;
    wmfBuffer[2] = 0xC6;
    wmfBuffer[3] = 0x9A;
    const image = await Image.fromBuffer(wmfBuffer);
    await image.ensureDataLoaded();
    expect(image.validateImageData().valid).toBe(true);
  });

  it('should detect SVG dimensions from width/height attributes', async () => {
    const svgBuffer = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="200" height="150"><rect/></svg>');
    const image = await Image.fromBuffer(svgBuffer);
    // SVG dimensions are detected and converted to EMUs
    // If no explicit dimensions passed, it auto-detects
    expect(image.getWidth()).toBeGreaterThan(0);
    expect(image.getHeight()).toBeGreaterThan(0);
  });
});

describe('Image Properties - MIME Types', () => {
  it('should return correct MIME type for SVG', () => {
    expect(ImageManager.getMimeType('svg')).toBe('image/svg+xml');
  });

  it('should return correct MIME type for EMF', () => {
    expect(ImageManager.getMimeType('emf')).toBe('image/x-emf');
  });

  it('should return correct MIME type for WMF', () => {
    expect(ImageManager.getMimeType('wmf')).toBe('image/x-wmf');
  });
});

describe('Image Properties - Floating Image cNvGraphicFramePr', () => {
  it('should include cNvGraphicFramePr for floating images', async () => {
    const doc = Document.create();
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });
    image.setPosition(
      { anchor: 'page', alignment: 'center' },
      { anchor: 'page', alignment: 'center' }
    );
    doc.addImage(image);

    const buffer = await doc.toBuffer();
    const doc2 = await Document.loadFromBuffer(buffer);
    // If cNvGraphicFramePr is properly generated, noChangeAspect should be parsed
    const img = doc2.getImages()[0]!.image;
    expect(img.getNoChangeAspect()).toBe(true);
    doc2.dispose();
  });
});
