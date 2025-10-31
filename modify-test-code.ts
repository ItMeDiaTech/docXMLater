/**
 * Script to modify Test_Code.docx with specific formatting requirements
 *
 * Requirements:
 * - Styles: Heading1 (18pt Verdana bold), Heading2 (14pt Verdana bold),
 *   Normal (12pt Verdana), List Paragraph (12pt Verdana + contextual spacing)
 * - Header 2s: Wrap in 1x1 tables with #BFBFBF shading, autofit, no top/bottom padding
 * - Hyperlinks: "Top of Document" → "Top of the Document", all color #0000FF
 * - Lists: 0.5" indent per level, alternating solid/open circle bullets
 * - Tables: Top row #D9D9D9 shading, centered bold 12pt, autofit
 */

import { Document, Paragraph, Table, TableRow, TableCell, Hyperlink, Style, NumberingLevel, TableOfContentsElement, StructuredDocumentTag } from './src';

// Type for body elements (not exported from framework)
type BodyElement = Paragraph | Table | TableOfContentsElement | StructuredDocumentTag;

async function modifyTestCode() {
  console.log('Loading Test_Code.docx...');

  // Phase 1: Load document
  const doc = await Document.load('./Test_Code.docx');

  // Verify doc is valid
  if (!doc || typeof doc.getBodyElements !== 'function') {
    throw new Error('Failed to load document - document object is invalid');
  }

  const stylesManager = doc.getStylesManager();
  const numberingManager = doc.getNumberingManager();
  const bodyElements = doc.getBodyElements();

  console.log(`Loaded document with ${bodyElements.length} body elements`);

  // Phase 2: Modify style definitions
  console.log('\nPhase 2: Modifying style definitions...');

  // Heading1: 18pt black bold Verdana, left aligned, 0pt before, 12pt after
  // Note: Spacing must be in twips (1pt = 20 twips)
  let heading1 = stylesManager.getStyle('Heading1');
  if (heading1) {
    heading1.setRunFormatting({
      font: 'Verdana',
      size: 18,
      color: '000000',
      bold: true
    });
    heading1.setParagraphFormatting({
      alignment: 'left',
      spacing: { before: 0, after: 240 } // 0pt, 12pt = 240 twips
    });
    console.log('✓ Updated Heading1 style');
  }

  // Heading2: 14pt black bold Verdana, left aligned, 6pt before/after
  let heading2 = stylesManager.getStyle('Heading2');
  if (heading2) {
    heading2.setRunFormatting({
      font: 'Verdana',
      size: 14,
      color: '000000',
      bold: true
    });
    heading2.setParagraphFormatting({
      alignment: 'left',
      spacing: { before: 120, after: 120 } // 6pt = 120 twips
    });
    console.log('✓ Updated Heading2 style');
  }

  // Normal: 12pt black Verdana, left aligned, 3pt before/after
  let normal = stylesManager.getStyle('Normal');
  if (normal) {
    normal.setRunFormatting({
      font: 'Verdana',
      size: 12,
      color: '000000'
    });
    normal.setParagraphFormatting({
      alignment: 'left',
      spacing: { before: 60, after: 60 } // 3pt = 60 twips
    });
    console.log('✓ Updated Normal style');
  }

  // List Paragraph: 12pt black Verdana, left aligned, 3pt before/after, contextual spacing
  let listPara = stylesManager.getStyle('ListParagraph');
  if (listPara) {
    listPara.setRunFormatting({
      font: 'Verdana',
      size: 12,
      color: '000000'
    });
    listPara.setParagraphFormatting({
      alignment: 'left',
      spacing: { before: 60, after: 60 }, // 3pt = 60 twips
      contextualSpacing: true
    });
    console.log('✓ Updated ListParagraph style');
  }

  // Phase 3: Wrap all Header 2 paragraphs in 1x1 tables
  console.log('\nPhase 3: Wrapping Header 2 paragraphs in tables...');

  const newBodyElements: BodyElement[] = [];
  let header2Count = 0;

  for (const element of bodyElements) {
    if (element instanceof Paragraph && element.getStyle() === 'Heading2') {
      // Create 1x1 table
      const table = new Table(1, 1);

      // Get the cell
      const cell = table.getRow(0)!.getCell(0)!;

      // Set cell shading #BFBFBF
      cell.setShading({ fill: 'BFBFBF' });

      // Set cell margins: top=0, bottom=0, preserve left/right
      // Get current margins first
      const currentMargins = cell.getFormatting().margins || {};
      cell.setMargins({
        top: 0,
        bottom: 0,
        left: currentMargins.left || 100,  // Keep existing or default
        right: currentMargins.right || 100
      });

      // Add the Header 2 paragraph to the cell
      cell.addParagraph(element);

      // Set table to autofit (100% width)
      table.setLayout('auto');
      table.setWidth(5000).setWidthType('pct'); // 100% width (5000 = 100.00%)

      newBodyElements.push(table);
      header2Count++;
      console.log(`  ✓ Wrapped Header 2 #${header2Count} in table`);
    } else {
      newBodyElements.push(element);
    }
  }

  // Replace body elements: clear and re-add transformed elements
  console.log('Clearing body elements...');
  doc.clearParagraphs();

  console.log(`Re-adding ${newBodyElements.length} elements...`);
  for (let i = 0; i < newBodyElements.length; i++) {
    const element = newBodyElements[i];
    if (element instanceof Paragraph) {
      doc.addParagraph(element);
    } else if (element instanceof Table) {
      doc.addTable(element);
    } else if (element instanceof TableOfContentsElement) {
      doc.addTableOfContents(element);
    } else if (element instanceof StructuredDocumentTag) {
      // Use type assertion since addBodyElement might not be available yet
      (doc as any).addBodyElement(element);
    }
  }
  console.log(`✓ Wrapped ${header2Count} Header 2 paragraphs in tables`);

  // Phase 4: Update hyperlinks
  console.log('\nPhase 4: Updating hyperlinks...');

  let hyperlinkCount = 0;
  let modifiedTextCount = 0;

  // Iterate through all elements to find hyperlinks
  for (const element of doc.getBodyElements()) {
    if (element instanceof Paragraph) {
      const content = element.getContent();

      for (let i = 0; i < content.length; i++) {
        const item = content[i];

        if (item instanceof Hyperlink) {
          hyperlinkCount++;
          const text = item.getText();
          const url = item.getUrl();
          const anchor = item.getAnchor();

          // Check if text needs to be changed
          let newText = text;
          if (text === 'Top of Document') {
            newText = 'Top of the Document';
            modifiedTextCount++;
          }

          // Create new hyperlink with #0000FF color
          let newLink: Hyperlink;
          if (url) {
            newLink = Hyperlink.createExternal(url, newText, {
              color: '0000FF',
              underline: 'single'
            });
          } else if (anchor) {
            newLink = Hyperlink.createInternal(anchor, newText, {
              color: '0000FF',
              underline: 'single'
            });
          } else {
            continue; // Skip if no URL or anchor
          }

          // Replace the hyperlink in the content array
          content[i] = newLink;
        }
      }
    } else if (element instanceof Table) {
      // Check hyperlinks in table cells
      for (const row of element.getRows()) {
        for (const cell of row.getCells()) {
          for (const para of cell.getParagraphs()) {
            const content = para.getContent();

            for (let i = 0; i < content.length; i++) {
              const item = content[i];

              if (item instanceof Hyperlink) {
                hyperlinkCount++;
                const text = item.getText();
                const url = item.getUrl();
                const anchor = item.getAnchor();

                let newText = text;
                if (text === 'Top of Document') {
                  newText = 'Top of the Document';
                  modifiedTextCount++;
                }

                let newLink: Hyperlink;
                if (url) {
                  newLink = Hyperlink.createExternal(url, newText, {
                    color: '0000FF',
                    underline: 'single'
                  });
                } else if (anchor) {
                  newLink = Hyperlink.createInternal(anchor, newText, {
                    color: '0000FF',
                    underline: 'single'
                  });
                } else {
                  continue;
                }

                content[i] = newLink;
              }
            }
          }
        }
      }
    }
  }

  console.log(`✓ Updated ${hyperlinkCount} hyperlinks to color #0000FF`);
  console.log(`✓ Changed ${modifiedTextCount} "Top of Document" → "Top of the Document"`);

  // Phase 5: Modify list numbering
  console.log('\nPhase 5: Modifying list numbering...');

  const abstractNums = numberingManager.getAllAbstractNumberings();

  for (const abstractNum of abstractNums) {
    // Modify levels 0-8
    for (let level = 0; level <= 8; level++) {
      const existingLevel = abstractNum.getLevel(level);
      if (existingLevel) {
        // Calculate indentation: 0.5" per level (720 twips per 0.5")
        const leftIndent = (level + 1) * 720;  // 720, 1440, 2160, 2880, etc.
        const hanging = 360;  // Standard hanging indent

        // Alternate between solid (●) and open (○) circle
        const bulletChar = level % 2 === 0 ? '●' : '○';

        // Create new level with modified properties
        const newLevel = new NumberingLevel({
          level: level,
          format: 'bullet',
          text: bulletChar,
          alignment: 'left',
          leftIndent: leftIndent,
          hangingIndent: hanging,
          font: 'Symbol'
        });

        // Replace the level (addLevel replaces if same index)
        abstractNum.addLevel(newLevel);
      }
    }
  }

  console.log(`✓ Updated ${abstractNums.length} numbering definitions`);
  console.log('  - Indentation: 0.5" per level');
  console.log('  - Bullets: alternating ● (solid) and ○ (open)');

  // Phase 6: Format multi-cell tables
  console.log('\nPhase 6: Formatting multi-cell tables...');

  let multiCellTableCount = 0;

  for (const element of doc.getBodyElements()) {
    if (element instanceof Table) {
      const rows = element.getRows();
      const cols = rows[0]?.getCells().length || 0;

      // Skip 1x1 tables (Header 2 wrappers)
      if (rows.length === 1 && cols === 1) {
        continue;
      }

      multiCellTableCount++;

      // Format first row
      const firstRow = rows[0];
      if (firstRow) {
        for (const cell of firstRow.getCells()) {
          // Set shading #D9D9D9
          cell.setShading({ fill: 'D9D9D9' });

          // Set cell margins: no top/bottom padding
          const currentMargins = cell.getFormatting().margins || {};
          cell.setMargins({
            top: 0,
            bottom: 0,
            left: currentMargins.left || 100,
            right: currentMargins.right || 100
          });

          // Format paragraphs in cell
          for (const para of cell.getParagraphs()) {
            para.setAlignment('center');
            para.setSpaceBefore(60); // 3pt = 60 twips
            para.setSpaceAfter(60);  // 3pt = 60 twips

            // Make all runs bold, Verdana 12pt, black
            for (const run of para.getRuns()) {
              run.setBold(true);
              run.setFont('Verdana', 12);
              run.setColor('000000');
            }
          }
        }
      }

      // Set table to autofit
      element.setLayout('auto');
      element.setWidth(5000).setWidthType('pct'); // 100% width (5000 = 100.00%)
    }
  }

  console.log(`✓ Formatted ${multiCellTableCount} multi-cell tables`);
  console.log('  - Top row: #D9D9D9 shading, centered, bold, 12pt Verdana');
  console.log('  - All tables: autofit to window');

  // Phase 7: Save
  console.log('\nPhase 7: Saving modified document...');
  await doc.save('./Test_Code_Modified.docx');
  console.log('✓ Saved as Test_Code_Modified.docx');

  console.log('\n=== Modification Complete ===');
  console.log('Summary:');
  console.log(`  - Modified 4 style definitions`);
  console.log(`  - Wrapped ${header2Count} Header 2 paragraphs in tables`);
  console.log(`  - Updated ${hyperlinkCount} hyperlinks`);
  console.log(`  - Modified ${abstractNums.length} numbering definitions`);
  console.log(`  - Formatted ${multiCellTableCount} multi-cell tables`);

  console.log('\n=== Helper Functions for Future Implementation ===');
  console.log('The following helper functions would simplify this task:');
  console.log('1. Document.wrapParagraphInTable(paragraph, tableOptions)');
  console.log('   - Would simplify wrapping Header 2s in tables');
  console.log('2. Hyperlink.setColor(color: string)');
  console.log('   - Direct color modification instead of recreating hyperlinks');
  console.log('3. Document.getElementsByStyle(styleId: string)');
  console.log('   - Query all paragraphs by style for bulk operations');
  console.log('4. Table.isMultiCell(): boolean');
  console.log('   - Detect if table is larger than 1x1');
  console.log('5. NumberingLevel.setBulletCharacter(char: string)');
  console.log('   - More intuitive than setFormat() + setText()');
  console.log('6. Paragraph.replaceHyperlinks(callback)');
  console.log('   - Simplify hyperlink modifications in bulk');
}

// Run the script
modifyTestCode().catch(error => {
  console.error('Error modifying document:', error);
  process.exit(1);
});
