/**
 * 13 Logging: configure docxmlater's internal log level and observe output.
 *
 * Set DOCXMLATER_LOG_LEVEL=debug|info|warn|error before running. The
 * library logs lifecycle events (load, save, validation) at info level
 * and detailed parse/generate steps at debug.
 *
 * Run with: DOCXMLATER_LOG_LEVEL=debug npm run 13-logging
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const level = process.env.DOCXMLATER_LOG_LEVEL ?? '(unset)';
  console.log(`Current DOCXMLATER_LOG_LEVEL: ${level}`);
  console.log('Set this env var to "debug" for the most verbose output.');
  console.log('');

  const doc = Document.create();
  doc.createParagraph('Logging Demo').setStyle('Title');
  doc.createParagraph(
    'Watch the terminal output as this document is built and saved. ' +
      'Set DOCXMLATER_LOG_LEVEL=debug to see internal lifecycle events.'
  );

  writeFileSync('13-logging.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 13-logging.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
