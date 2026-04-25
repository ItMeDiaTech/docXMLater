/**
 * Example: Math equations (Office MathML, m:oMath / m:oMathPara).
 *
 * docxmlater preserves equations as raw XML passthrough via
 * `MathParagraph` (block-level `m:oMathPara`) and `MathExpression`
 * (inline `m:oMath`). Word renders them through its built-in equation
 * engine — docxmlater does not have an equation editor; it just keeps
 * the markup byte-for-byte safe across round-trips.
 *
 * This example demonstrates block-level math paragraphs added to the
 * document body. Inline math (m:oMath inside a w:p) is created during
 * parsing and survives round-trip; programmatically inserting an inline
 * MathExpression into a Paragraph requires the element-registry route
 * (see ElementRegistry plugin docs).
 *
 * Run: `npx ts-node examples/18-math-equations/math-equations.ts`
 */

import { Document, MathParagraph } from '../../src';

const QUADRATIC_FORMULA = `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:oMath>
    <m:r><m:t>x = </m:t></m:r>
    <m:f>
      <m:fPr><m:ctrlPr/></m:fPr>
      <m:num>
        <m:r><m:t>-b ± </m:t></m:r>
        <m:rad>
          <m:radPr><m:degHide m:val="1"/><m:ctrlPr/></m:radPr>
          <m:deg/>
          <m:e><m:r><m:t>b² - 4ac</m:t></m:r></m:e>
        </m:rad>
      </m:num>
      <m:den><m:r><m:t>2a</m:t></m:r></m:den>
    </m:f>
  </m:oMath>
</m:oMathPara>`;

const PYTHAGOREAN = `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <m:oMath>
    <m:sSup>
      <m:sSupPr><m:ctrlPr/></m:sSupPr>
      <m:e><m:r><m:t>a</m:t></m:r></m:e>
      <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
    </m:sSup>
    <m:r><m:t> + </m:t></m:r>
    <m:sSup>
      <m:sSupPr><m:ctrlPr/></m:sSupPr>
      <m:e><m:r><m:t>b</m:t></m:r></m:e>
      <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
    </m:sSup>
    <m:r><m:t> = </m:t></m:r>
    <m:sSup>
      <m:sSupPr><m:ctrlPr/></m:sSupPr>
      <m:e><m:r><m:t>c</m:t></m:r></m:e>
      <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
    </m:sSup>
  </m:oMath>
</m:oMathPara>`;

async function main() {
  const doc = Document.create();

  doc.createParagraph('Math Equations').setStyle('Heading1');

  doc.createParagraph('The quadratic formula:');
  doc.addBodyElement(new MathParagraph(QUADRATIC_FORMULA));

  doc.createParagraph('The Pythagorean theorem:');
  doc.addBodyElement(new MathParagraph(PYTHAGOREAN));

  await doc.save('examples/18-math-equations/output.docx');
  doc.dispose();
  // eslint-disable-next-line no-console
  console.log('Wrote examples/18-math-equations/output.docx');
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exit(1);
});
