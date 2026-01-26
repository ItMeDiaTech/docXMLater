/**
 * Test what extractHyperlinks returns for Original_16.docx
 * Check if field codes with revisions are being returned as Hyperlink objects
 */

import { Document, Hyperlink, Revision, ComplexField } from './src/index';

async function main() {
  console.log('='.repeat(60));
  console.log('TEST: EXTRACT HYPERLINKS FROM ORIGINAL_16.DOCX');
  console.log('='.repeat(60));

  // Load document
  const doc = await Document.load('Original_16.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });

  const paragraphs = doc.getAllParagraphs();
  console.log(`\nTotal paragraphs: ${paragraphs.length}`);

  // Find paragraphs containing 046762 or 075115
  let found046762 = false;
  let found075115 = false;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;
    const text = para.getText();
    const content = para.getContent();

    const has046762 = text.includes('046762');
    const has075115 = text.includes('075115');

    if (!has046762 && !has075115) continue;

    console.log('\n' + '='.repeat(50));
    console.log(`Paragraph ${i}: ${has046762 ? '046762' : ''}${has075115 ? '075115' : ''}`);
    console.log('='.repeat(50));
    console.log(`  Text: "${text.substring(0, 80)}..."`);
    console.log(`  Content items: ${content.length}`);

    // Analyze each content item
    for (let j = 0; j < content.length; j++) {
      const item = content[j];
      if (!item) continue;

      const typeName = item.constructor.name;
      console.log(`  [${j}] ${typeName}`);

      if (item instanceof Hyperlink) {
        console.log(`       URL: ${item.getUrl()}`);
        console.log(`       Text: ${item.getText()}`);
        console.log(`       *** This is a Hyperlink object ***`);
        if (has046762) found046762 = true;
        if (has075115) found075115 = true;
      } else if (item instanceof Revision) {
        const rev = item as any;
        console.log(`       Type: ${rev.getType()}`);
        console.log(`       ID: ${rev.getId()}`);
        const revContent = rev.getContent();
        console.log(`       Content: ${revContent.length} items`);

        // Check if contains Hyperlink
        for (const child of revContent) {
          if (child instanceof Hyperlink) {
            console.log(`         *** Contains Hyperlink: ${child.getUrl()} ***`);
            if (has046762) found046762 = true;
            if (has075115) found075115 = true;
          }
        }
      } else if (item instanceof ComplexField) {
        const field = item as any;
        console.log(`       isHyperlinkField: ${field.isHyperlinkField ? field.isHyperlinkField() : 'N/A'}`);
        console.log(`       *** This is a ComplexField ***`);
        if (has046762) found046762 = true;
        if (has075115) found075115 = true;
      } else if (typeName === 'Run') {
        const run = item as any;
        const runContent = run.getContent();
        const types = runContent.map((c: any) => c.type);
        const fldChar = runContent.find((c: any) => c.type === 'fieldChar');
        if (fldChar) {
          console.log(`       fieldChar: ${fldChar.fieldCharType}`);
        } else if (types.includes('instructionText')) {
          console.log(`       Has instructionText`);
        } else {
          console.log(`       Types: ${types.join(', ')}`);
        }
      }
    }
  }

  console.log('\n' + '='.repeat(60));
  console.log('SUMMARY');
  console.log('='.repeat(60));
  console.log(`046762 found as Hyperlink/ComplexField: ${found046762}`);
  console.log(`075115 found as Hyperlink/ComplexField: ${found075115}`);

  // Now test what extractHyperlinks would return
  console.log('\n' + '='.repeat(60));
  console.log('SIMULATING EXTRACTHYPERLINKS');
  console.log('='.repeat(60));

  const hyperlinks: any[] = [];
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;
    const content = para.getContent();

    for (const item of content) {
      // Case 1: Direct Hyperlink
      if (item && typeof (item as any).getUrl === 'function') {
        hyperlinks.push({
          type: 'direct',
          url: (item as any).getUrl(),
          text: (item as any).getText?.() || '',
          paragraphIndex: i
        });
      }
      // Case 2: Hyperlinks inside Revision
      if (item && typeof (item as any).getContent === 'function') {
        const revContent = (item as any).getContent();
        if (Array.isArray(revContent)) {
          for (const inner of revContent) {
            if (inner && typeof inner.getUrl === 'function') {
              hyperlinks.push({
                type: 'inRevision',
                url: inner.getUrl(),
                text: inner.getText?.() || '',
                paragraphIndex: i
              });
            }
          }
        }
      }
    }
  }

  console.log(`\nFound ${hyperlinks.length} hyperlinks total`);

  // Check if 046762 or 075115 are in the extracted hyperlinks
  const has046762InList = hyperlinks.some(h => h.text.includes('046762') || h.url?.includes('046762'));
  const has075115InList = hyperlinks.some(h => h.text.includes('075115') || h.url?.includes('075115'));

  console.log(`\n046762 in extractHyperlinks result: ${has046762InList}`);
  console.log(`075115 in extractHyperlinks result: ${has075115InList}`);

  // Show those hyperlinks if found
  const matches = hyperlinks.filter(h =>
    h.text.includes('046762') || h.url?.includes('046762') ||
    h.text.includes('075115') || h.url?.includes('075115')
  );

  if (matches.length > 0) {
    console.log(`\nMatching hyperlinks:`);
    for (const m of matches) {
      console.log(`  - [${m.type}] Text: "${m.text.substring(0, 50)}..."`);
      console.log(`    URL: ${m.url?.substring(0, 60)}...`);
    }
  }

  doc.dispose();

  console.log('\n' + '='.repeat(60));
  console.log('TEST COMPLETE');
  console.log('='.repeat(60));
}

main().catch(console.error);
