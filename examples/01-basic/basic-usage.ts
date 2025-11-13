/**
 * Basic usage examples for DocXML
 */

import { ZipHandler, DOCX_PATHS } from '../src';

// Example 1: Create a simple DOCX file
async function createSimpleDocx() {
  const handler = new ZipHandler();

  // Add minimal required files
  handler.addFile(DOCX_PATHS.CONTENT_TYPES, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

  handler.addFile(DOCX_PATHS.RELS, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

  handler.addFile(DOCX_PATHS.DOCUMENT, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World! This is a DOCX created with DocXML.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`);

  await handler.save('example-simple.docx');
  console.log('✓ Created example-simple.docx');
}

// Example 2: Read and inspect a DOCX file
async function readDocx(filePath: string) {
  const handler = new ZipHandler();
  await handler.load(filePath);

  console.log(`\nInspecting: ${filePath}`);
  console.log('Files in archive:');

  const paths = handler.getFilePaths();
  paths.forEach(path => {
    const file = handler.getFile(path);
    console.log(`  - ${path} (${file?.size} bytes, ${file?.isBinary ? 'binary' : 'text'})`);
  });

  // Get document content
  const docXml = handler.getFileAsString(DOCX_PATHS.DOCUMENT);
  console.log('\nDocument XML preview:');
  console.log(docXml?.substring(0, 200) + '...');
}

// Example 3: Modify an existing DOCX
async function modifyDocx() {
  await ZipHandler.modify(
    'example-simple.docx',
    'example-modified.docx',
    (handler) => {
      let xml = handler.getFileAsString(DOCX_PATHS.DOCUMENT) || '';

      // Replace text
      xml = xml.replace('Hello, World!', 'Hello, Modified World!');

      // Update the file
      handler.updateFile(DOCX_PATHS.DOCUMENT, xml);
    }
  );

  console.log('✓ Created example-modified.docx');
}

// Example 4: Add an image (placeholder - would need actual image file)
async function addImageExample() {
  const handler = new ZipHandler();
  await handler.load('example-simple.docx');

  // In a real scenario, you would read an actual image file
  const fakeImageData = Buffer.from('fake-image-data');

  handler.addFile('word/media/image1.png', fakeImageData, { binary: true });

  await handler.save('example-with-image.docx');
  console.log('✓ Created example-with-image.docx (with placeholder image)');
}

// Run examples
async function runExamples() {
  try {
    console.log('=== DocXML Examples ===\n');

    // Example 1
    console.log('1. Creating a simple DOCX...');
    await createSimpleDocx();

    // Example 2
    console.log('\n2. Reading the DOCX...');
    await readDocx('example-simple.docx');

    // Example 3
    console.log('\n3. Modifying the DOCX...');
    await modifyDocx();

    // Example 4
    console.log('\n4. Adding an image...');
    await addImageExample();

    console.log('\n=== All examples completed successfully! ===');
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  runExamples();
}
