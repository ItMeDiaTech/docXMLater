/**
 * Font Reading Example
 * Demonstrates how to read/extract fonts from existing DOCX files
 */

import { FontManager } from '../../src/index';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, 'output');

/**
 * Example 1: Parse Font Extensions from Content_Types.xml
 * Shows how to detect which font formats are registered
 */
function example1_ParseFontExtensions() {
  console.log('\n=== Example 1: Parse Font Extensions ===');

  // Sample Content_Types.xml with fonts
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="ttf" ContentType="application/x-font-ttf"/>
  <Default Extension="otf" ContentType="application/x-font-opentype"/>
  <Default Extension="woff" ContentType="application/font-woff"/>
  <Default Extension="png" ContentType="image/png"/>
</Types>`;

  // Parse to find font extensions
  const fontExtensions = FontManager.parseFontExtensionsFromContentTypes(contentTypesXml);

  console.log(`Found ${fontExtensions.length} font format(s):`);
  fontExtensions.forEach(ext => {
    const mimeType = FontManager.getMimeType(ext as any);
    console.log(`  ${ext.toUpperCase()}: ${mimeType}`);
  });

  console.log('\n✓ Successfully parsed font extensions from Content_Types.xml');
}

/**
 * Example 2: Load Fonts from Archive
 * Shows how FontManager can load fonts from a ZIP archive
 */
function example2_LoadFontsFromArchive() {
  console.log('\n=== Example 2: Load Fonts from Archive ===');

  const fontManager = FontManager.create();

  // Simulate ZIP archive contents
  const zipFiles = new Map<string, Buffer | string>();

  // Add some font files
  const sampleFont1 = Buffer.from('TTF_DATA_1');
  const sampleFont2 = Buffer.from('OTF_DATA_2');
  const sampleFont3 = Buffer.from('WOFF_DATA_3');

  zipFiles.set('word/fonts/roboto_1.ttf', sampleFont1);
  zipFiles.set('word/fonts/open_sans_2.otf', sampleFont2);
  zipFiles.set('word/fonts/lato_3.woff', sampleFont3);
  zipFiles.set('word/document.xml', '<xml>...</xml>'); // Non-font file

  // Sample Content_Types.xml that registers these formats
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="ttf" ContentType="application/x-font-ttf"/>
  <Default Extension="otf" ContentType="application/x-font-opentype"/>
  <Default Extension="woff" ContentType="application/font-woff"/>
</Types>`;

  // Load fonts from archive
  fontManager.loadFontsFromArchive(zipFiles, contentTypesXml);

  console.log(`Loaded ${fontManager.getCount()} fonts from archive:`);
  const fonts = fontManager.getAllFonts();
  fonts.forEach((font, index) => {
    console.log(`  ${index + 1}. ${font.fontFamily} (${font.format.toUpperCase()})`);
    console.log(`     Path: ${font.path}`);
    console.log(`     Size: ${font.data.length} bytes`);
  });

  console.log('\n✓ Successfully loaded fonts from archive');
}

/**
 * Example 3: Get Font by Family Name
 * Shows how to retrieve specific fonts
 */
function example3_GetFontByFamily() {
  console.log('\n=== Example 3: Get Font by Family Name ===');

  const fontManager = FontManager.create();

  // Add some fonts
  fontManager.addFont('Arial', Buffer.from('ARIAL_DATA'), 'ttf');
  fontManager.addFont('Times New Roman', Buffer.from('TIMES_DATA'), 'otf');
  fontManager.addFont('Roboto', Buffer.from('ROBOTO_DATA'), 'woff');

  // Get specific font
  console.log('Looking up fonts by family name:');

  const arial = fontManager.getFontByFamily('Arial');
  if (arial) {
    console.log(`  ✓ Found Arial: ${arial.path} (${arial.format})`);
  }

  const times = fontManager.getFontByFamily('Times New Roman');
  if (times) {
    console.log(`  ✓ Found Times New Roman: ${times.path} (${times.format})`);
  }

  const helvetica = fontManager.getFontByFamily('Helvetica');
  if (!helvetica) {
    console.log(`  ✗ Helvetica not found (as expected)`);
  }

  console.log('\n✓ Font lookup working correctly');
}

/**
 * Example 4: Get Font by Path
 * Shows how to retrieve fonts by their archive path
 */
function example4_GetFontByPath() {
  console.log('\n=== Example 4: Get Font by Path ===');

  const fontManager = FontManager.create();

  // Add fonts
  const arialPath = fontManager.addFont('Arial', Buffer.from('DATA'), 'ttf');
  const robotoPath = fontManager.addFont('Roboto', Buffer.from('DATA'), 'otf');

  console.log('Looking up fonts by path:');

  const arialFont = fontManager.getFontByPath(arialPath);
  if (arialFont) {
    console.log(`  ✓ ${arialPath} → ${arialFont.fontFamily}`);
  }

  const robotoFont = fontManager.getFontByPath(robotoPath);
  if (robotoFont) {
    console.log(`  ✓ ${robotoPath} → ${robotoFont.fontFamily}`);
  }

  const invalidFont = fontManager.getFontByPath('word/fonts/nonexistent.ttf');
  if (!invalidFont) {
    console.log(`  ✗ Invalid path returns undefined (as expected)`);
  }

  console.log('\n✓ Path-based lookup working correctly');
}

/**
 * Example 5: Detect Font Format
 * Shows format detection from extensions
 */
function example5_DetectFontFormat() {
  console.log('\n=== Example 5: Detect Font Format ===');

  const extensions = ['ttf', 'otf', 'woff', 'woff2', '.TTF', 'docx'];

  console.log('Format detection:');
  extensions.forEach(ext => {
    const format = FontManager.detectFormatFromExtension(ext);
    if (format) {
      console.log(`  ${ext.padEnd(8)} → ${format}`);
    } else {
      console.log(`  ${ext.padEnd(8)} → (not a font format)`);
    }
  });

  console.log('\n✓ Format detection working correctly');
}

/**
 * Example 6: Extract Font from Loaded Document
 * Shows typical workflow for extracting fonts from a document
 */
async function example6_ExtractFontsWorkflow() {
  console.log('\n=== Example 6: Extract Fonts Workflow ===');

  console.log('Typical workflow for extracting fonts:');
  console.log('');
  console.log('1. Load DOCX file');
  console.log('   const doc = await Document.load("document.docx")');
  console.log('');
  console.log('2. Access FontManager (when implemented in Document class)');
  console.log('   const fontManager = doc.getFontManager()');
  console.log('');
  console.log('3. Get all fonts');
  console.log('   const fonts = fontManager.getAllFonts()');
  console.log('');
  console.log('4. Access specific font');
  console.log('   const customFont = fontManager.getFontByFamily("Custom Font")');
  console.log('');
  console.log('5. Extract font data');
  console.log('   if (customFont) {');
  console.log('     fs.writeFileSync("extracted.ttf", customFont.data)');
  console.log('   }');

  console.log('\n✓ Workflow documented');
}

/**
 * Example 7: Check if Document Has Fonts
 * Shows how to check for embedded fonts
 */
function example7_CheckForFonts() {
  console.log('\n=== Example 7: Check for Fonts ===');

  // Document without fonts
  const emptyFontManager = FontManager.create();
  console.log(`Document 1 has fonts: ${emptyFontManager.getCount() > 0}`);
  console.log(`  Font count: ${emptyFontManager.getCount()}`);

  // Document with fonts
  const populatedFontManager = FontManager.create();
  populatedFontManager.addFont('Arial', Buffer.from('DATA'), 'ttf');
  populatedFontManager.addFont('Roboto', Buffer.from('DATA'), 'otf');
  console.log(`\nDocument 2 has fonts: ${populatedFontManager.getCount() > 0}`);
  console.log(`  Font count: ${populatedFontManager.getCount()}`);
  console.log(`  Font families: ${populatedFontManager.getAllFonts().map(f => f.fontFamily).join(', ')}`);

  // Check specific font
  console.log(`\n  Has Arial: ${populatedFontManager.hasFont('Arial')}`);
  console.log(`  Has Times: ${populatedFontManager.hasFont('Times')}`);

  console.log('\n✓ Font checking working correctly');
}

/**
 * Example 8: List All Fonts in Archive
 * Comprehensive font inventory
 */
function example8_FontInventory() {
  console.log('\n=== Example 8: Font Inventory ===');

  const fontManager = FontManager.create();

  // Add various fonts
  const fonts = [
    { family: 'Arial', format: 'ttf' as const, data: Buffer.from('DATA1') },
    { family: 'Times New Roman', format: 'otf' as const, data: Buffer.from('DATA2') },
    { family: 'Roboto', format: 'woff' as const, data: Buffer.from('DATA3') },
    { family: 'Open Sans', format: 'woff2' as const, data: Buffer.from('DATA4') },
  ];

  fonts.forEach(({ family, format, data }) => {
    fontManager.addFont(family, data, format);
  });

  // Generate inventory report
  console.log('Font Inventory Report:');
  console.log('='.repeat(60));

  const allFonts = fontManager.getAllFonts();
  allFonts.forEach((font, index) => {
    console.log(`\nFont #${index + 1}:`);
    console.log(`  Family:   ${font.fontFamily}`);
    console.log(`  Format:   ${font.format.toUpperCase()}`);
    console.log(`  Path:     ${font.path}`);
    console.log(`  Filename: ${font.filename}`);
    console.log(`  Size:     ${font.data.length} bytes`);
    console.log(`  MIME:     ${FontManager.getMimeType(font.format)}`);
  });

  console.log('\n' + '='.repeat(60));
  console.log(`Total Fonts: ${fontManager.getCount()}`);
  console.log(`Extensions: ${Array.from(fontManager.getExtensions()).join(', ')}`);

  console.log('\n✓ Complete font inventory generated');
}

/**
 * Main execution
 */
async function main() {
  console.log('docXMLater - Font Reading Examples\n');
  console.log('Demonstrates how to read/extract fonts from DOCX files\n');

  // Create output directory
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  }

  try {
    example1_ParseFontExtensions();
    example2_LoadFontsFromArchive();
    example3_GetFontByFamily();
    example4_GetFontByPath();
    example5_DetectFontFormat();
    await example6_ExtractFontsWorkflow();
    example7_CheckForFonts();
    example8_FontInventory();

    console.log('\n✓ All font reading examples completed successfully!');
    console.log('\nKey Points:');
    console.log('  1. parseFontExtensionsFromContentTypes() - Parse Content_Types.xml');
    console.log('  2. loadFontsFromArchive() - Load fonts from ZIP');
    console.log('  3. getFontByFamily() - Get font by name');
    console.log('  4. getFontByPath() - Get font by archive path');
    console.log('  5. getAllFonts() - Get complete font list');
    console.log('  6. getCount() - Check if document has fonts');
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  main();
}

export {
  example1_ParseFontExtensions,
  example2_LoadFontsFromArchive,
  example3_GetFontByFamily,
  example4_GetFontByPath,
  example5_DetectFontFormat,
  example6_ExtractFontsWorkflow,
  example7_CheckForFonts,
  example8_FontInventory,
};
