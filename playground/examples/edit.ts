/**
 * Example: load an existing document, modify it, save it back.
 *
 * This is the workflow most other DOCX libraries cannot do without losing
 * formatting or breaking the file. docxmlater preserves the original XML
 * and only regenerates the parts you actually touched.
 *
 * Run with: npm run edit
 */

import { Document } from 'docxmlater';
import { readFileSync, writeFileSync, existsSync } from 'node:fs';

async function main() {
  // For convenience, create a starting file if one doesn't exist yet.
  if (!existsSync('input.docx')) {
    const seed = Document.create();
    seed.createParagraph().addText('Acme Corporation - draft proposal', {
      bold: true,
      fontSize: 20,
    });
    seed
      .createParagraph()
      .addText('Pricing as of 2025-01-01. Please replace [CUSTOMER] with the buyer name.');
    seed.createParagraph().addText('Total contract value: $00,000.00 (placeholder).');
    writeFileSync('input.docx', await seed.toBuffer());
    seed.dispose();
    console.log('Generated a sample input.docx for this example.');
  }

  // Load the existing file. All formatting, styles, and metadata are preserved.
  const doc = await Document.loadFromBuffer(readFileSync('input.docx'));

  // Replace placeholders. replaceText works across run boundaries.
  doc.replaceText('[CUSTOMER]', 'Globex Inc.');
  doc.replaceText('$00,000.00 (placeholder)', '$248,500.00');
  doc.replaceText('draft proposal', 'final proposal');

  // Apply formatting to all matching text in one call.
  doc.findAndHighlight('Globex Inc.', 'yellow');

  // Append a fresh paragraph at the end.
  doc.createParagraph().addText('Signed and approved on behalf of Acme.', {
    italic: true,
  });

  const buffer = await doc.toBuffer();
  writeFileSync('output.docx', buffer);
  doc.dispose();

  console.log('Wrote output.docx');
  console.log('Open output.docx in Word to see the modified result.');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
