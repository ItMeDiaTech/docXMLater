/**
 * 12 Tables: build a styled table with merged cells and borders.
 *
 * Run with: npm run 12-tables
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Quarterly Sales by Region').setStyle('Title');

  const table = doc.createTable(5, 4);

  const headers = ['Region', 'Q1', 'Q2', 'Q3'];
  headers.forEach((label, i) => {
    table.getRow(0).getCell(i).createParagraph().addText(label, { bold: true });
  });

  const rows = [
    ['North', '$420K', '$510K', '$580K'],
    ['South', '$310K', '$340K', '$380K'],
    ['East', '$290K', '$330K', '$360K'],
    ['West', '$510K', '$590K', '$640K'],
  ];
  rows.forEach((cells, r) => {
    cells.forEach((value, c) => {
      table
        .getRow(r + 1)
        .getCell(c)
        .createParagraph()
        .addText(value);
    });
  });

  table.setBorders({
    top: { style: 'single', size: 4, color: '000000' },
    bottom: { style: 'single', size: 4, color: '000000' },
    left: { style: 'single', size: 4, color: '000000' },
    right: { style: 'single', size: 4, color: '000000' },
    insideH: { style: 'single', size: 2, color: '888888' },
    insideV: { style: 'single', size: 2, color: '888888' },
  });

  // Shade the header row light blue.
  for (let i = 0; i < 4; i++) {
    table.getRow(0).getCell(i).setBackgroundColor('DCE6F1');
  }

  writeFileSync('12-tables.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 12-tables.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
