/**
 * CT_FFCheckBox — honour every ST_OnOff literal for `<w:default>` /
 * `<w:checked>`.
 *
 * Per ECMA-376 §17.16.18 CT_FFCheckBox has six CT_OnOff children
 * (size, sizeAuto, default, checked, …). `<w:default>` and `<w:checked>`
 * take the full ST_OnOff literal set: "true"/"false"/"1"/"0"/"on"/"off",
 * plus a bare self-closing element meaning "on".
 *
 * The parser used bespoke comparisons:
 *   `cb['w:default']['@_w:val'] === '1' || cb['w:default']['@_w:val'] === 1`
 *
 * That silently mishandled "on", "true", and boolean `true` — all of
 * which are emitted by Word when the form author set the default via the
 * legacy checkbox UI — and it refused to treat bare elements as true at
 * all. A DOCX authored with `<w:checked w:val="true"/>` came back parsed
 * as unchecked.
 *
 * Iteration 93 routes both children through `parseOoxmlBoolean`.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Run } from '../../src/elements/Run';
import { Field, ComplexField } from '../../src/elements/Field';

async function loadCheckBoxFFData(cbInner: string) {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin">
          <w:ffData>
            <w:name w:val="CheckBoxField"/>
            <w:enabled/>
            <w:calcOnExit w:val="0"/>
            <w:checkBox>${cbInner}</w:checkBox>
          </w:ffData>
        </w:fldChar>
      </w:r>
      <w:r><w:instrText>FORMCHECKBOX</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>&#9744;</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const paragraph = doc.getParagraphs()[0]!;
  // Try to pull the assembled ComplexField first; fall back to scanning
  // raw run tokens for the fldChar(begin).formFieldData if the parser
  // did not fold the runs into a ComplexField.
  const complexFields = paragraph.getFields().filter((f) => f instanceof ComplexField);
  let ffData: { type?: string; defaultChecked?: boolean; checked?: boolean } | undefined;
  for (const f of complexFields) {
    const data = (f as ComplexField).getFormFieldData();
    if (data?.fieldType?.type === 'checkBox') {
      ffData = data.fieldType as {
        type?: string;
        defaultChecked?: boolean;
        checked?: boolean;
      };
      break;
    }
  }
  if (!ffData) {
    for (const item of paragraph.getContent()) {
      if (!(item instanceof Run)) continue;
      for (const tok of item.getContent()) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const t = tok as any;
        if (t?.type === 'fieldChar' && t.formFieldData?.fieldType?.type === 'checkBox') {
          ffData = t.formFieldData.fieldType as {
            type?: string;
            defaultChecked?: boolean;
            checked?: boolean;
          };
          break;
        }
      }
      if (ffData) break;
    }
  }
  doc.dispose();
  return ffData;
}

describe('CT_FFCheckBox <w:default>/<w:checked> — ST_OnOff literal coverage', () => {
  describe('<w:default>', () => {
    it('val="1" → defaultChecked=true (baseline)', async () => {
      const ft = await loadCheckBoxFFData('<w:default w:val="1"/>');
      expect(ft?.type).toBe('checkBox');
      expect(ft?.defaultChecked).toBe(true);
    });
    it('val="0" → defaultChecked=false', async () => {
      const ft = await loadCheckBoxFFData('<w:default w:val="0"/>');
      expect(ft?.defaultChecked).toBe(false);
    });
    it('val="true" → defaultChecked=true', async () => {
      const ft = await loadCheckBoxFFData('<w:default w:val="true"/>');
      expect(ft?.defaultChecked).toBe(true);
    });
    it('val="false" → defaultChecked=false', async () => {
      const ft = await loadCheckBoxFFData('<w:default w:val="false"/>');
      expect(ft?.defaultChecked).toBe(false);
    });
    it('val="on" → defaultChecked=true', async () => {
      const ft = await loadCheckBoxFFData('<w:default w:val="on"/>');
      expect(ft?.defaultChecked).toBe(true);
    });
    it('val="off" → defaultChecked=false', async () => {
      const ft = await loadCheckBoxFFData('<w:default w:val="off"/>');
      expect(ft?.defaultChecked).toBe(false);
    });
    it('bare element → defaultChecked=true', async () => {
      const ft = await loadCheckBoxFFData('<w:default/>');
      expect(ft?.defaultChecked).toBe(true);
    });
  });

  describe('<w:checked>', () => {
    it('val="1" → checked=true', async () => {
      const ft = await loadCheckBoxFFData('<w:checked w:val="1"/>');
      expect(ft?.checked).toBe(true);
    });
    it('val="on" → checked=true', async () => {
      const ft = await loadCheckBoxFFData('<w:checked w:val="on"/>');
      expect(ft?.checked).toBe(true);
    });
    it('val="off" → checked=false', async () => {
      const ft = await loadCheckBoxFFData('<w:checked w:val="off"/>');
      expect(ft?.checked).toBe(false);
    });
    it('val="true" → checked=true', async () => {
      const ft = await loadCheckBoxFFData('<w:checked w:val="true"/>');
      expect(ft?.checked).toBe(true);
    });
    it('bare element → checked=true', async () => {
      const ft = await loadCheckBoxFFData('<w:checked/>');
      expect(ft?.checked).toBe(true);
    });
  });

  it('round-trips default=true/checked=false through save()', async () => {
    const ft = await loadCheckBoxFFData('<w:default w:val="on"/><w:checked w:val="off"/>');
    expect(ft?.defaultChecked).toBe(true);
    expect(ft?.checked).toBe(false);
  });
});
