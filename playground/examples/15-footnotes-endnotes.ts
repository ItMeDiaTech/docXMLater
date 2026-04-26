/**
 * 15 Footnotes and Endnotes: register notes on the document and inspect
 * the managers that own them.
 *
 * In v11.x footnote and endnote *content* is created via
 * doc.createFootnote() / doc.createEndnote(). Inserting the reference glyph
 * into a specific paragraph requires the lower-level Run API and is most
 * commonly used when round-tripping documents that already contain notes.
 *
 * Run with: npm run 15-footnotes-endnotes
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Notes Demonstration').setStyle('Title');

  doc.createParagraph(
    'docxmlater supports both footnotes (per-page) and endnotes (end of document) with full round-trip fidelity. The example below registers two notes and confirms they are tracked by the document.'
  );

  doc.createParagraph('Footnote and Endnote Registration').setStyle('Heading1');

  const footnote = doc.createFootnote('Source: internal research data, Q3 2025.');
  const endnote = doc.createEndnote('See the methodology appendix for details on data collection.');

  console.log(`Created footnote ID: ${footnote.getId()}`);
  console.log(`Created endnote ID:  ${endnote.getId()}`);
  console.log(`Footnotes in manager: ${doc.getFootnoteManager().getAllFootnotes().length}`);
  console.log(`Endnotes in manager:  ${doc.getEndnoteManager().getAllEndnotes().length}`);

  doc.createParagraph(
    `This document has ${doc.getFootnoteManager().getAllFootnotes().length} footnote(s) and ${doc.getEndnoteManager().getAllEndnotes().length} endnote(s) registered.`
  );

  doc.createParagraph('Round-trip workflow').setStyle('Heading1');
  doc.createParagraph(
    'For documents authored in Word that already contain notes, Document.load() preserves them automatically. You can read, modify, or remove individual notes through getFootnoteManager() and getEndnoteManager().'
  );

  writeFileSync('15-footnotes-endnotes.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 15-footnotes-endnotes.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
