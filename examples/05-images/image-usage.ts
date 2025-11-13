/**
 * Image Usage Examples
 *
 * This example demonstrates how to add images to Word documents using DocXML.
 * Images are embedded using DrawingML and can be loaded from files or buffers.
 */

import { Document, Image, inchesToEmus } from '../../src';
import * as fs from 'fs';
import * as path from 'path';

// Create output directory if it doesn't exist
const outputDir = __dirname;
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

/**
 * Example 1: Create a simple 1x1 pixel PNG image programmatically
 * This is useful for testing without requiring external image files
 */
function createSimplePNG(): Buffer {
  // Minimal PNG file (1x1 pixel, red)
  // PNG signature + IHDR + IDAT + IEND chunks
  const pngData = Buffer.from([
    // PNG signature
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    // IHDR chunk
    0x00, 0x00, 0x00, 0x0d, // Length: 13 bytes
    0x49, 0x48, 0x44, 0x52, // "IHDR"
    0x00, 0x00, 0x00, 0x01, // Width: 1
    0x00, 0x00, 0x00, 0x01, // Height: 1
    0x08, 0x02, 0x00, 0x00, 0x00, // Bit depth, color type, etc.
    0x90, 0x77, 0x53, 0xde, // CRC
    // IDAT chunk (compressed image data)
    0x00, 0x00, 0x00, 0x0c, // Length: 12 bytes
    0x49, 0x44, 0x41, 0x54, // "IDAT"
    0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00,
    0x03, 0x01, 0x01, 0x00,
    0x18, 0xdd, 0x8d, 0xb4, // CRC
    // IEND chunk
    0x00, 0x00, 0x00, 0x00, // Length: 0 bytes
    0x49, 0x45, 0x4e, 0x44, // "IEND"
    0xae, 0x42, 0x60, 0x82  // CRC
  ]);

  return pngData;
}

/**
 * Example 1: Add a simple image from buffer
 */
async function example1_SimpleImage() {
  console.log('Example 1: Adding a simple image from buffer...');

  const doc = Document.create({
    properties: {
      title: 'Simple Image Example',
      creator: 'DocXML Examples',
    },
  });

  // Add title
  doc.createParagraph()
    .setAlignment('center')
    .setSpaceBefore(240)
    .setSpaceAfter(240)
    .addText('Simple Image Example', { bold: true, size: 18 });

  // Add description
  doc.createParagraph()
    .setSpaceAfter(240)
    .addText('This document contains a simple image loaded from a buffer.');

  // Create a simple PNG image
  const imageBuffer = createSimplePNG();

  // Create image from buffer
  const image = Image.fromBuffer(
    imageBuffer,
    'png',
    inchesToEmus(2), // 2 inches wide
    inchesToEmus(2)  // 2 inches tall
  );

  // Add image to document
  doc.addImage(image);

  // Add caption
  doc.createParagraph()
    .setAlignment('center')
    .setSpaceAfter(240)
    .addText('Figure 1: Simple test image', { italic: true, size: 10 });

  // Save document
  const outputPath = path.join(outputDir, 'example1-simple-image.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 2: Add multiple images with different sizes
 */
async function example2_MultipleImages() {
  console.log('Example 2: Adding multiple images with different sizes...');

  const doc = Document.create({
    properties: {
      title: 'Multiple Images Example',
      creator: 'DocXML Examples',
    },
  });

  // Add title
  doc.createParagraph()
    .setAlignment('center')
    .setSpaceBefore(240)
    .setSpaceAfter(480)
    .addText('Multiple Images Example', { bold: true, size: 20 });

  const imageBuffer = createSimplePNG();

  // Add images of different sizes
  const sizes = [
    { width: 1, height: 1, label: 'Small (1x1 inch)' },
    { width: 2, height: 2, label: 'Medium (2x2 inches)' },
    { width: 3, height: 2, label: 'Large (3x2 inches)' },
  ];

  for (const size of sizes) {
    // Add section header
    doc.createParagraph()
      .setSpaceBefore(240)
      .setSpaceAfter(120)
      .addText(size.label, { bold: true, size: 14 });

    // Create and add image
    const image = Image.fromBuffer(
      imageBuffer,
      'png',
      inchesToEmus(size.width),
      inchesToEmus(size.height)
    );

    doc.addImage(image);

    // Add spacing
    doc.createParagraph()
      .setSpaceAfter(240)
      .addText(''); // Empty paragraph for spacing
  }

  // Save document
  const outputPath = path.join(outputDir, 'example2-multiple-images.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 3: Image with text formatting around it
 */
async function example3_ImageWithText() {
  console.log('Example 3: Image mixed with text content...');

  const doc = Document.create({
    properties: {
      title: 'Image with Text Example',
      creator: 'DocXML Examples',
    },
  });

  // Add title
  doc.createParagraph()
    .setAlignment('center')
    .setSpaceBefore(240)
    .setSpaceAfter(480)
    .addText('Document with Mixed Content', { bold: true, size: 20 });

  // Add introduction
  doc.createParagraph()
    .setAlignment('justify')
    .setSpaceAfter(240)
    .addText(
      'Images can be seamlessly integrated into documents alongside text content. ' +
      'This example demonstrates how images work within the document flow.'
    );

  // Add section heading
  doc.createParagraph()
    .setSpaceBefore(240)
    .setSpaceAfter(120)
    .addText('Visual Content', { bold: true, size: 14 });

  // Add some text before image
  doc.createParagraph()
    .setSpaceAfter(120)
    .addText('Below is an embedded image:');

  // Add image
  const imageBuffer = createSimplePNG();
  const image = Image.create({
    source: imageBuffer,
    width: inchesToEmus(2.5),
    height: inchesToEmus(2.5),
    name: 'Sample Image',
    description: 'A sample test image for demonstration',
  });

  doc.addImage(image);

  // Add caption
  doc.createParagraph()
    .setAlignment('center')
    .setSpaceAfter(240)
    .addText('Figure 1: Sample demonstration image', { italic: true, size: 10 });

  // Add text after image
  doc.createParagraph()
    .setSpaceBefore(240)
    .setAlignment('justify')
    .addText(
      'As you can see, images are embedded directly into the document and can be ' +
      'combined with various text formatting options, paragraph styles, and other elements.'
    );

  // Save document
  const outputPath = path.join(outputDir, 'example3-image-with-text.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 4: Using Image class methods
 */
async function example4_ImageMethods() {
  console.log('Example 4: Using Image class methods...');

  const doc = Document.create({
    properties: {
      title: 'Image Methods Example',
      creator: 'DocXML Examples',
    },
  });

  // Add title
  doc.createParagraph()
    .setAlignment('center')
    .setSpaceBefore(240)
    .setSpaceAfter(480)
    .addText('Image Class Methods', { bold: true, size: 20 });

  const imageBuffer = createSimplePNG();

  // Example: Creating image and modifying size
  doc.createParagraph()
    .setSpaceAfter(120)
    .addText('Image created with default size then resized:', { bold: true });

  const image1 = Image.fromBuffer(imageBuffer, 'png');
  image1.setWidth(inchesToEmus(2), true); // Maintain aspect ratio
  doc.addImage(image1);

  doc.createParagraph()
    .setAlignment('center')
    .setSpaceAfter(480)
    .addText('Width set to 2 inches (aspect ratio maintained)', { italic: true, size: 10 });

  // Example: Set specific dimensions
  doc.createParagraph()
    .setSpaceAfter(120)
    .addText('Image with specific dimensions:', { bold: true });

  const image2 = Image.fromBuffer(imageBuffer, 'png');
  image2.setSize(inchesToEmus(3), inchesToEmus(1.5)); // 3x1.5 inches
  doc.addImage(image2);

  doc.createParagraph()
    .setAlignment('center')
    .setSpaceAfter(480)
    .addText('Dimensions: 3 x 1.5 inches', { italic: true, size: 10 });

  // Example: Height adjustment
  doc.createParagraph()
    .setSpaceAfter(120)
    .addText('Image with height set:', { bold: true });

  const image3 = Image.fromBuffer(imageBuffer, 'png');
  image3.setHeight(inchesToEmus(2.5), true); // Maintain aspect ratio
  doc.addImage(image3);

  doc.createParagraph()
    .setAlignment('center')
    .addText('Height set to 2.5 inches (aspect ratio maintained)', { italic: true, size: 10 });

  // Save document
  const outputPath = path.join(outputDir, 'example4-image-methods.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log('Running Image Examples...\n');

  try {
    await example1_SimpleImage();
    await example2_MultipleImages();
    await example3_ImageWithText();
    await example4_ImageMethods();

    console.log('\n✓ All examples completed successfully!');
    console.log(`\nOutput files saved to: ${outputDir}`);
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run examples if executed directly
if (require.main === module) {
  main();
}

export {
  example1_SimpleImage,
  example2_MultipleImages,
  example3_ImageWithText,
  example4_ImageMethods,
};
