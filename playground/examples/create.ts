/**
 * Example: create a new document from scratch.
 *
 * Run with: npm run create
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph().addText('Project Proposal', {
    bold: true,
    fontSize: 32,
    color: '1F4E79',
  });

  doc.createParagraph().addText('Prepared for the leadership team. Confidential.', {
    italic: true,
    color: '7F7F7F',
  });

  doc.addHeading('Overview', 1);
  doc
    .createParagraph()
    .addText(
      'This proposal outlines the scope, timeline, and budget for the Q3 platform initiative.'
    );

  doc.addHeading('Budget', 2);
  const table = doc.createTable(4, 3);
  const header = table.getRow(0);
  header.getCell(0).createParagraph().addText('Item', { bold: true });
  header.getCell(1).createParagraph().addText('Quantity', { bold: true });
  header.getCell(2).createParagraph().addText('Cost', { bold: true });

  const rows = [
    ['Engineering', '4 FTE', '$480,000'],
    ['Design', '1 FTE', '$120,000'],
    ['Infrastructure', '-', '$45,000'],
  ];
  rows.forEach(([item, qty, cost], i) => {
    const row = table.getRow(i + 1);
    row.getCell(0).createParagraph().addText(item);
    row.getCell(1).createParagraph().addText(qty);
    row.getCell(2).createParagraph().addText(cost);
  });

  table.setBorders({
    top: { style: 'single', size: 4, color: '000000' },
    bottom: { style: 'single', size: 4, color: '000000' },
    left: { style: 'single', size: 4, color: '000000' },
    right: { style: 'single', size: 4, color: '000000' },
    insideH: { style: 'single', size: 4, color: '000000' },
    insideV: { style: 'single', size: 4, color: '000000' },
  });

  const buffer = await doc.toBuffer();
  writeFileSync('proposal.docx', buffer);
  doc.dispose();

  console.log('Wrote proposal.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
