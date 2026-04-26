/**
 * 20 Compatibility Mode: detect a legacy Word version and upgrade the
 * document to Word 2013+ format. Equivalent to File -> Info -> Convert in Word.
 *
 * Run with: npm run 20-compatibility-mode
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  // Build a fresh document. New documents target Word 2013+ by default,
  // so this example creates one, then inspects the compatibility mode.
  const doc = Document.create();

  doc.createParagraph('Compatibility Mode').setStyle('Title');
  doc.createParagraph(
    'Older Word documents target legacy compatibility modes (Word 2003, 2007, 2010). docxmlater can detect the mode and upgrade legacy flags to the modern Word 2013+ baseline.'
  );

  doc.createParagraph('Inspection').setStyle('Heading1');
  doc.createParagraph(
    `Detected mode: ${doc.getCompatibilityMode()}    (11 = Word 2003, 12 = Word 2007, 14 = Word 2010, 15 = Word 2013+)`
  );
  doc.createParagraph(`Is legacy mode: ${doc.isCompatibilityMode() ? 'yes' : 'no'}`);

  // For loaded legacy documents, you would call doc.upgradeToModernFormat().
  // The call returns a report of what was changed:
  doc.createParagraph('Upgrade workflow').setStyle('Heading1');
  doc.createParagraph('For a real legacy document loaded via Document.load(path), call:');
  doc.createParagraph('const report = doc.upgradeToModernFormat();', {
    font: 'Courier New',
    size: 10,
  });
  doc.createParagraph(
    'The returned report lists removedFlags and addedSettings so you know exactly what changed.'
  );

  console.log(`Detected compatibility mode: ${doc.getCompatibilityMode()}`);
  console.log(`Is legacy: ${doc.isCompatibilityMode()}`);

  writeFileSync('20-compatibility-mode.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 20-compatibility-mode.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
