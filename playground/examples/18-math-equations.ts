/**
 * 18 Math Equations: round-trip OMML math content.
 *
 * Authoring math equations programmatically requires the OMML schema, which
 * is verbose. The most common workflow is to author the equation in Word
 * once, then load and re-save the document with docxmlater. The math markup
 * is preserved verbatim.
 *
 * Run with: npm run 18-math-equations
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Math Equations').setStyle('Title');

  doc.createParagraph('Round-trip support').setStyle('Heading1');
  doc.createParagraph(
    'Equations authored in Word are stored as Office MathML (OMML) inside the document XML. docxmlater parses, preserves, and re-serializes this markup unchanged on save.'
  );

  doc.createParagraph('Workflow').setStyle('Heading1');
  doc.createParagraph('1. Open Word and insert an equation through Insert > Equation.');
  doc.createParagraph('2. Save the .docx and load it with Document.load(path).');
  doc.createParagraph('3. Modify any other part of the document programmatically.');
  doc.createParagraph('4. Save with doc.save(path). The math equation is preserved exactly.');

  doc.createParagraph('Authoring inline').setStyle('Heading1');
  doc.createParagraph(
    'For programmatic equation authoring, build the OMML XML directly and inject it as a preserved XML element. See the agent_docs notes in the main repo for examples.'
  );

  writeFileSync('18-math-equations.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 18-math-equations.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
