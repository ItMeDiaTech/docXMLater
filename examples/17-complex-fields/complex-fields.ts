/**
 * Example: Complex Field Codes (HYPERLINK, REF, IF, MERGEFIELD).
 *
 * Demonstrates the two ways docxmlater represents fields:
 *
 * 1. Simple fields via `Field` (`<w:fldSimple>`) — single-instruction,
 *    self-contained: `{ FIELDNAME args \\* MERGEFORMAT }`. Ideal for
 *    PAGE, NUMPAGES, DATE, TIME.
 *
 * 2. Complex fields via `ComplexField` (`<w:fldChar>` begin/separate/end
 *    with intermediate runs) — required for nested fields, calculated
 *    values, and multi-paragraph TOC fields.
 *
 * Run: `npx ts-node examples/17-complex-fields/complex-fields.ts`
 */

import { Document, Paragraph, Field, ComplexField } from '../../src';

async function main() {
  const doc = Document.create();

  // Heading
  doc.createParagraph('Complex Field Codes').setStyle('Heading1');

  // 1. Simple PAGE / NUMPAGES (footer-style "Page X of Y")
  doc.createParagraph('Page numbering:');
  const pageOf = new Paragraph();
  pageOf.addText('Page ');
  pageOf.addContent(new Field({ type: 'PAGE' }));
  pageOf.addText(' of ');
  pageOf.addContent(new Field({ type: 'NUMPAGES' }));
  doc.addParagraph(pageOf);

  // 2. DATE field with custom format
  doc.createParagraph('Inserted on:');
  const dateLine = new Paragraph();
  dateLine.addContent(new Field({ type: 'DATE', dateFormat: 'MMMM d, yyyy' }));
  doc.addParagraph(dateLine);

  // 3. HYPERLINK as a complex field with display text
  // Equivalent to: {HYPERLINK "https://docxmlater.dev"}docxmlater{}
  doc.createParagraph('Project link:');
  const hyperlinkPara = new Paragraph();
  hyperlinkPara.addContent(
    new ComplexField({
      instruction: ' HYPERLINK "https://docxmlater.dev" ',
      result: 'docxmlater',
    })
  );
  doc.addParagraph(hyperlinkPara);

  // 4. MERGEFIELD with default placeholder text
  doc.createParagraph('Mail-merge placeholder:');
  const mergePara = new Paragraph();
  mergePara.addContent(
    new ComplexField({
      instruction: ' MERGEFIELD CustomerName ',
      result: '«CustomerName»',
    })
  );
  doc.addParagraph(mergePara);

  // 5. IF field — conditional content
  doc.createParagraph('Conditional content:');
  const ifPara = new Paragraph();
  ifPara.addContent(
    new ComplexField({
      instruction: ' IF { MERGEFIELD Total } > 1000 "VIP" "Standard" ',
      result: 'Standard',
    })
  );
  doc.addParagraph(ifPara);

  // 6. REF field referencing a bookmark
  doc.createParagraph('Cross-reference:');
  const refPara = new Paragraph();
  refPara.addContent(
    new ComplexField({
      instruction: ' REF SectionTitle \\h ',
      result: 'See Section 1',
    })
  );
  doc.addParagraph(refPara);

  // Save
  await doc.save('examples/17-complex-fields/output.docx');
  doc.dispose();
  // eslint-disable-next-line no-console
  console.log('Wrote examples/17-complex-fields/output.docx');
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exit(1);
});
