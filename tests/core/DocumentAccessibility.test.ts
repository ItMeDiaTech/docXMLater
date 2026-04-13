import { Document, Image } from '../../src';

// Minimal valid PNG buffer (1x1 pixel)
function createPngBuffer(): Buffer {
  // PNG signature + minimal IHDR + IDAT + IEND
  return Buffer.from(
    '89504e470d0a1a0a' + // PNG signature
      '0000000d49484452' + // IHDR chunk length + type
      '00000001000000010800000000' + // 1x1, 8-bit grayscale
      '1f15c489' + // IHDR CRC
      '0000000a4944415478' + // IDAT chunk
      '9c626000000002000198e195' + // compressed data
      '0000000049454e44ae426082', // IEND
    'hex'
  );
}

describe('Document.findImagesWithoutAltText', () => {
  let doc: Document;

  afterEach(() => {
    doc?.dispose();
  });

  it('should return empty array when no images exist', () => {
    doc = Document.create();
    doc.createParagraph().addText('No images here');
    expect(doc.findImagesWithoutAltText()).toEqual([]);
  });

  it('should find images with default alt text', async () => {
    doc = Document.create();
    const image = await Image.fromBuffer(createPngBuffer(), {
      width: 914400,
      height: 914400,
    });
    doc.addImage(image);

    const missing = doc.findImagesWithoutAltText();
    expect(missing.length).toBe(1);
  });

  it('should not find images with custom alt text', async () => {
    doc = Document.create();
    const image = await Image.fromBuffer(createPngBuffer(), {
      width: 914400,
      height: 914400,
      description: 'A photo of a landscape',
    });
    doc.addImage(image);

    const missing = doc.findImagesWithoutAltText();
    expect(missing.length).toBe(0);
  });

  it('should find multiple images missing alt text', async () => {
    doc = Document.create();

    const img1 = await Image.fromBuffer(createPngBuffer(), {
      width: 914400,
      height: 914400,
    });
    doc.addImage(img1);

    const img2 = await Image.fromBuffer(createPngBuffer(), {
      width: 914400,
      height: 914400,
      description: 'Has alt text',
    });
    doc.addImage(img2);

    const img3 = await Image.fromBuffer(createPngBuffer(), {
      width: 914400,
      height: 914400,
    });
    doc.addImage(img3);

    const missing = doc.findImagesWithoutAltText();
    expect(missing.length).toBe(2);
  });

  it('should allow fixing found images', async () => {
    doc = Document.create();
    const image = await Image.fromBuffer(createPngBuffer(), {
      width: 914400,
      height: 914400,
    });
    doc.addImage(image);

    const missing = doc.findImagesWithoutAltText();
    expect(missing.length).toBe(1);

    for (const img of missing) {
      img.setAltText('Fixed alt text');
    }

    expect(doc.findImagesWithoutAltText().length).toBe(0);
  });
});
