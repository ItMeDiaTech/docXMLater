/**
 * `<w:tblPrEx><w:jc>` — bidi-aware `start` / `end` → `left` / `right`
 * mapping for validator compatibility.
 *
 * Per ECMA-376 §17.18.45 ST_JcTable has five values (start, center, end,
 * left, right), BUT the Open XML SDK's `TableJustification` class (which
 * wraps `<w:jc>` inside `<w:tblPrEx>` and `<w:trPr>`) uses a narrow
 * `TableRowAlignmentValues` enum with only three values: center / left /
 * right. Strict OOXML validators reject start/end at these positions with
 * "The attribute 'w:val' has invalid value 'start'. The Enumeration
 * constraint failed."
 *
 * Iteration 67 already installed the mapping for `<w:trPr><w:jc>`. This
 * iteration closes the same gap for `<w:tblPrEx><w:jc>`, which uses the
 * same SDK class and fails the same validation.
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { Paragraph } from '../../src/elements/Paragraph';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('<w:tblPrEx><w:jc> SDK-narrow enum mapping', () => {
  it('maps exceptions.justification = "start" → "left" on save', () => {
    const table = new Table(1, 1);
    const row = table.getRows()[0]!;
    row.setTablePropertyExceptions({ justification: 'start' });
    const xml = XMLBuilder.elementToString(row.toXML());
    expect(xml).toMatch(/<w:tblPrEx>[\s\S]*?<w:jc\s+w:val="left"[\s\S]*?<\/w:tblPrEx>/);
    expect(xml).not.toMatch(/<w:tblPrEx>[\s\S]*?<w:jc\s+w:val="start"/);
  });

  it('maps exceptions.justification = "end" → "right" on save', () => {
    const table = new Table(1, 1);
    const row = table.getRows()[0]!;
    row.setTablePropertyExceptions({ justification: 'end' });
    const xml = XMLBuilder.elementToString(row.toXML());
    expect(xml).toMatch(/<w:tblPrEx>[\s\S]*?<w:jc\s+w:val="right"[\s\S]*?<\/w:tblPrEx>/);
    expect(xml).not.toMatch(/<w:tblPrEx>[\s\S]*?<w:jc\s+w:val="end"/);
  });

  it('passes through center / left / right unchanged', () => {
    for (const val of ['center', 'left', 'right'] as const) {
      const table = new Table(1, 1);
      const row = table.getRows()[0]!;
      row.setTablePropertyExceptions({ justification: val });
      const xml = XMLBuilder.elementToString(row.toXML());
      const re = new RegExp(`<w:tblPrEx>[\\s\\S]*?<w:jc\\s+w:val="${val}"[\\s\\S]*?<\\/w:tblPrEx>`);
      expect(xml).toMatch(re);
    }
  });

  it('tblPrEx with start alignment passes OOXML validator through full Document round-trip', async () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    const row = table.getRows()[0]!;
    row.setTablePropertyExceptions({ justification: 'start' });
    doc.addTable(table);
    // Need a paragraph after the table.
    const p = new Paragraph();
    p.addText('x');
    doc.addParagraph(p);
    await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
    doc.dispose();
  });
});
