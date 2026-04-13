/**
 * Tests for Paragraph.toJSON(), Paragraph.fromJSON(), and Table.mapColumn()
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';

// ============================================================================
// Paragraph.toJSON()
// ============================================================================

describe('Paragraph.toJSON()', () => {
  it('serializes plain text paragraph', () => {
    const para = new Paragraph().addText('Hello World');
    const json = para.toJSON();

    expect(json.text).toBe('Hello World');
    expect(json.runs).toHaveLength(1);
    expect(json.runs[0]!.text).toBe('Hello World');
  });

  it('serializes multiple runs with formatting', () => {
    const para = new Paragraph();
    para.addText('Bold', { bold: true });
    para.addText(' and ');
    para.addText('Italic', { italic: true });

    const json = para.toJSON();

    expect(json.text).toBe('Bold and Italic');
    expect(json.runs).toHaveLength(3);
    expect(json.runs[0]!.formatting.bold).toBe(true);
    expect(json.runs[2]!.formatting.italic).toBe(true);
  });

  it('includes style', () => {
    const para = new Paragraph().addText('Title');
    para.setStyle('Heading1');

    expect(para.toJSON().style).toBe('Heading1');
  });

  it('includes alignment', () => {
    const para = new Paragraph().addText('Centered');
    para.setAlignment('center');

    expect(para.toJSON().alignment).toBe('center');
  });

  it('includes indentation', () => {
    const para = new Paragraph().addText('Indented');
    para.setLeftIndent(720);
    para.setFirstLineIndent(360);

    const json = para.toJSON();
    expect(json.indentation?.left).toBe(720);
    expect(json.indentation?.firstLine).toBe(360);
  });

  it('includes spacing', () => {
    const para = new Paragraph().addText('Spaced');
    para.setSpaceBefore(240);
    para.setSpaceAfter(120);

    const json = para.toJSON();
    expect(json.spacing?.before).toBe(240);
    expect(json.spacing?.after).toBe(120);
  });

  it('omits undefined properties', () => {
    const para = new Paragraph().addText('Simple');
    const json = para.toJSON();

    expect(json.style).toBeUndefined();
    expect(json.alignment).toBeUndefined();
    expect(json.numbering).toBeUndefined();
    expect(json.indentation).toBeUndefined();
    expect(json.spacing).toBeUndefined();
  });

  it('handles empty paragraph', () => {
    const para = new Paragraph();
    const json = para.toJSON();

    expect(json.text).toBe('');
    expect(json.runs).toHaveLength(0);
  });

  it('is JSON-serializable', () => {
    const para = new Paragraph();
    para.addText('Test', { bold: true, color: 'FF0000' });
    para.setStyle('Heading2');

    const serialized = JSON.stringify(para.toJSON());
    const parsed = JSON.parse(serialized);

    expect(parsed.text).toBe('Test');
    expect(parsed.style).toBe('Heading2');
    expect(parsed.runs[0].formatting.bold).toBe(true);
  });

  it('preserves complex run formatting', () => {
    const para = new Paragraph();
    para.addText('Styled', {
      bold: true,
      italic: true,
      underline: 'double',
      font: 'Arial',
      size: 14,
      color: '0000FF',
    });

    const json = para.toJSON();
    const fmt = json.runs[0]!.formatting;
    expect(fmt.bold).toBe(true);
    expect(fmt.italic).toBe(true);
    expect(fmt.underline).toBe('double');
    expect(fmt.font).toBe('Arial');
    expect(fmt.size).toBe(14);
    expect(fmt.color).toBe('0000FF');
  });
});

// ============================================================================
// Paragraph.fromJSON()
// ============================================================================

describe('Paragraph.fromJSON()', () => {
  it('creates paragraph from simple JSON', () => {
    const para = Paragraph.fromJSON({
      text: 'Hello World',
      runs: [{ text: 'Hello World' }],
    });

    expect(para.getText()).toBe('Hello World');
  });

  it('creates paragraph with formatted runs', () => {
    const para = Paragraph.fromJSON({
      runs: [{ text: 'Bold', formatting: { bold: true } }, { text: ' Normal' }],
    });

    const runs = para.getRuns();
    expect(runs).toHaveLength(2);
    expect(runs[0]!.getFormatting().bold).toBe(true);
    expect(runs[1]!.getText()).toBe(' Normal');
  });

  it('applies style', () => {
    const para = Paragraph.fromJSON({ text: 'Title', style: 'Heading1', runs: [] });

    expect(para.getStyle()).toBe('Heading1');
  });

  it('applies alignment', () => {
    const para = Paragraph.fromJSON({ text: 'Centered', alignment: 'center', runs: [] });

    expect(para.getAlignment()).toBe('center');
  });

  it('applies indentation', () => {
    const para = Paragraph.fromJSON({
      runs: [{ text: 'Indented' }],
      indentation: { left: 720, right: 360 },
    });

    expect(para.getFormatting().indentation?.left).toBe(720);
    expect(para.getFormatting().indentation?.right).toBe(360);
  });

  it('applies spacing', () => {
    const para = Paragraph.fromJSON({
      runs: [{ text: 'Spaced' }],
      spacing: { before: 240, after: 120 },
    });

    expect(para.getFormatting().spacing?.before).toBe(240);
    expect(para.getFormatting().spacing?.after).toBe(120);
  });

  it('falls back to text when no runs provided', () => {
    const para = Paragraph.fromJSON({ text: 'Fallback text' });

    expect(para.getText()).toBe('Fallback text');
  });

  it('handles empty input', () => {
    const para = Paragraph.fromJSON({});
    expect(para.getText()).toBe('');
  });

  it('round-trips through toJSON/fromJSON', () => {
    const original = new Paragraph();
    original.addText('Bold', { bold: true });
    original.addText(' and ');
    original.addText('Red', { color: 'FF0000' });
    original.setStyle('Heading2');
    original.setAlignment('center');
    original.setSpaceBefore(240);

    const json = original.toJSON();
    const restored = Paragraph.fromJSON(json);

    expect(restored.getText()).toBe(original.getText());
    expect(restored.getStyle()).toBe('Heading2');
    expect(restored.getAlignment()).toBe('center');
    expect(restored.getFormatting().spacing?.before).toBe(240);

    const runs = restored.getRuns();
    expect(runs[0]!.getFormatting().bold).toBe(true);
    expect(runs[2]!.getFormatting().color).toBe('FF0000');
  });
});

// ============================================================================
// Table.mapColumn()
// ============================================================================

describe('Table.mapColumn()', () => {
  it('transforms text in a column', () => {
    const table = Table.fromArray([
      ['Name', 'City'],
      ['alice', 'new york'],
      ['bob', 'london'],
    ]);

    table.mapColumn(0, (text) => text.toUpperCase());

    expect(table.getCell(0, 0)!.getText()).toBe('NAME');
    expect(table.getCell(1, 0)!.getText()).toBe('ALICE');
    expect(table.getCell(2, 0)!.getText()).toBe('BOB');
    // Column 1 unchanged
    expect(table.getCell(1, 1)!.getText()).toBe('new york');
  });

  it('provides row index to transform function', () => {
    const table = Table.fromArray([['Price'], ['100'], ['200']]);

    // Skip header row (index 0), format data rows
    table.mapColumn(0, (text, rowIndex) => (rowIndex === 0 ? text : `$${Number(text).toFixed(2)}`));

    expect(table.getCell(0, 0)!.getText()).toBe('Price');
    expect(table.getCell(1, 0)!.getText()).toBe('$100.00');
    expect(table.getCell(2, 0)!.getText()).toBe('$200.00');
  });

  it('returns this for chaining', () => {
    const table = Table.fromArray([['A'], ['B']]);
    const result = table.mapColumn(0, (t) => t.toLowerCase());

    expect(result).toBe(table);
  });

  it('chains multiple column transforms', () => {
    const table = Table.fromArray([
      ['name', 'score'],
      ['alice', '95'],
      ['bob', '87'],
    ]);

    table
      .mapColumn(0, (t) => t.charAt(0).toUpperCase() + t.slice(1))
      .mapColumn(1, (t, r) => (r === 0 ? t : `${t}%`));

    expect(table.toArray()).toEqual([
      ['Name', 'score'],
      ['Alice', '95%'],
      ['Bob', '87%'],
    ]);
  });

  it('handles empty cells', () => {
    const table = Table.fromArray([
      ['', 'B'],
      ['C', ''],
    ]);

    table.mapColumn(0, (text) => text || 'N/A');

    expect(table.getCell(0, 0)!.getText()).toBe('N/A');
    expect(table.getCell(1, 0)!.getText()).toBe('C');
  });

  it('does not modify cells when transform returns same text', () => {
    const table = Table.fromArray([['Keep'], ['Same']]);

    table.mapColumn(0, (text) => text);

    expect(table.getCell(0, 0)!.getText()).toBe('Keep');
    expect(table.getCell(1, 0)!.getText()).toBe('Same');
  });

  it('handles out-of-range column index gracefully', () => {
    const table = Table.fromArray([['A', 'B']]);

    // Should not throw
    table.mapColumn(99, (t) => t.toUpperCase());

    expect(table.getCell(0, 0)!.getText()).toBe('A');
  });

  it('practical: add prefix to ID column', () => {
    const table = Table.fromArray([
      ['ID', 'Name'],
      ['1', 'Alice'],
      ['2', 'Bob'],
      ['3', 'Charlie'],
    ]);

    table.mapColumn(0, (text, row) => (row === 0 ? text : `EMP-${text.padStart(4, '0')}`));

    expect(table.toArray()).toEqual([
      ['ID', 'Name'],
      ['EMP-0001', 'Alice'],
      ['EMP-0002', 'Bob'],
      ['EMP-0003', 'Charlie'],
    ]);
  });

  it('practical: calculate totals column', () => {
    const table = Table.fromArray([
      ['Item', 'Price', 'Qty', 'Total'],
      ['Widget', '10', '5', ''],
      ['Gadget', '25', '2', ''],
    ]);

    table.mapColumn(3, (_, row) => {
      if (row === 0) return 'Total';
      const price = Number(table.getCell(row, 1)!.getText());
      const qty = Number(table.getCell(row, 2)!.getText());
      return String(price * qty);
    });

    expect(table.getCell(1, 3)!.getText()).toBe('50');
    expect(table.getCell(2, 3)!.getText()).toBe('50');
  });
});
