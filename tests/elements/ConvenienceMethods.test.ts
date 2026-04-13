/**
 * Tests for TableCell.setBackgroundColor/getBackgroundColor,
 * Paragraph.addLineBreak/addColumnBreak, and Run.getPlainText
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';
import { TableCell } from '../../src/elements/TableCell';

// ============================================================================
// TableCell.setBackgroundColor / getBackgroundColor
// ============================================================================

describe('TableCell.setBackgroundColor()', () => {
  it('sets a background color on the cell', () => {
    const cell = new TableCell();
    cell.setBackgroundColor('FF0000');

    expect(cell.getShading()?.fill).toBe('FF0000');
    expect(cell.getShading()?.pattern).toBe('clear');
  });

  it('getBackgroundColor returns the fill color', () => {
    const cell = new TableCell();
    cell.setBackgroundColor('00FF00');

    expect(cell.getBackgroundColor()).toBe('00FF00');
  });

  it('getBackgroundColor returns undefined when not set', () => {
    const cell = new TableCell();
    expect(cell.getBackgroundColor()).toBeUndefined();
  });

  it('returns this for chaining', () => {
    const cell = new TableCell();
    const result = cell.setBackgroundColor('FFFF00');

    expect(result).toBe(cell);
  });

  it('overwrites previous background color', () => {
    const cell = new TableCell();
    cell.setBackgroundColor('FF0000');
    cell.setBackgroundColor('0000FF');

    expect(cell.getBackgroundColor()).toBe('0000FF');
  });

  it('works within a table context', () => {
    const table = new Table(2, 3);
    table.getCell(0, 0)!.setBackgroundColor('F2F2F2');
    table.getCell(0, 1)!.setBackgroundColor('E0E0E0');
    table.getCell(0, 2)!.setBackgroundColor('D0D0D0');

    expect(table.getCell(0, 0)!.getBackgroundColor()).toBe('F2F2F2');
    expect(table.getCell(0, 1)!.getBackgroundColor()).toBe('E0E0E0');
    expect(table.getCell(0, 2)!.getBackgroundColor()).toBe('D0D0D0');
    expect(table.getCell(1, 0)!.getBackgroundColor()).toBeUndefined();
  });

  it('generates valid XML with background color', () => {
    const cell = new TableCell();
    cell.createParagraph('Colored');
    cell.setBackgroundColor('FFFF00');

    const xml = cell.toXML();
    expect(xml.name).toBe('w:tc');
  });

  it('enables table zebra-striping pattern', () => {
    const table = new Table(4, 2);
    for (let r = 0; r < 4; r++) {
      for (let c = 0; c < 2; c++) {
        table.getCell(r, c)!.createParagraph(`R${r}C${c}`);
        if (r % 2 === 1) {
          table.getCell(r, c)!.setBackgroundColor('F0F0F0');
        }
      }
    }

    expect(table.getCell(0, 0)!.getBackgroundColor()).toBeUndefined();
    expect(table.getCell(1, 0)!.getBackgroundColor()).toBe('F0F0F0');
    expect(table.getCell(2, 0)!.getBackgroundColor()).toBeUndefined();
    expect(table.getCell(3, 0)!.getBackgroundColor()).toBe('F0F0F0');
  });
});

// ============================================================================
// Paragraph.addLineBreak / addColumnBreak
// ============================================================================

describe('Paragraph.addLineBreak()', () => {
  it('adds a line break to the paragraph', () => {
    const para = new Paragraph();
    para.addText('Line 1');
    para.addLineBreak();
    para.addText('Line 2');

    const text = para.getText();
    expect(text).toContain('Line 1');
    expect(text).toContain('Line 2');
    // getText maps breaks to \n
    expect(text).toBe('Line 1\nLine 2');
  });

  it('returns this for chaining', () => {
    const para = new Paragraph();
    const result = para.addLineBreak();

    expect(result).toBe(para);
  });

  it('can chain text and breaks fluently', () => {
    const para = new Paragraph()
      .addText('First')
      .addLineBreak()
      .addText('Second')
      .addLineBreak()
      .addText('Third');

    expect(para.getText()).toBe('First\nSecond\nThird');
  });

  it('adds a run with break content', () => {
    const para = new Paragraph();
    para.addLineBreak();

    const runs = para.getRuns();
    expect(runs.length).toBeGreaterThanOrEqual(1);

    // Find the run with a break
    const breakRun = runs.find((r) => {
      const content = r.getContent();
      return content.some((c) => c.type === 'break');
    });
    expect(breakRun).toBeDefined();
  });

  it('generates valid XML', () => {
    const para = new Paragraph();
    para.addText('Before');
    para.addLineBreak();
    para.addText('After');

    const xml = para.toXML();
    expect(xml.name).toBe('w:p');
  });
});

describe('Paragraph.addColumnBreak()', () => {
  it('adds a column break', () => {
    const para = new Paragraph();
    para.addText('Before column break');
    para.addColumnBreak();

    const runs = para.getRuns();
    const breakRun = runs.find((r) => {
      const content = r.getContent();
      return content.some((c) => c.type === 'break' && c.breakType === 'column');
    });
    expect(breakRun).toBeDefined();
  });

  it('returns this for chaining', () => {
    const para = new Paragraph();
    expect(para.addColumnBreak()).toBe(para);
  });
});

// ============================================================================
// Run.getPlainText
// ============================================================================

describe('Run.getPlainText()', () => {
  it('returns text content only', () => {
    const run = new Run('Hello World');

    expect(run.getPlainText()).toBe('Hello World');
  });

  it('excludes tab characters', () => {
    const run = new Run('');
    run.appendText('Name');
    run.addTab();
    run.appendText('Value');

    expect(run.getText()).toBe('Name\tValue');
    expect(run.getPlainText()).toBe('NameValue');
  });

  it('excludes line breaks', () => {
    const run = new Run('');
    run.appendText('Line 1');
    run.addBreak();
    run.appendText('Line 2');

    expect(run.getText()).toBe('Line 1\nLine 2');
    expect(run.getPlainText()).toBe('Line 1Line 2');
  });

  it('excludes carriage returns', () => {
    const run = new Run('');
    run.appendText('Before');
    run.addCarriageReturn();
    run.appendText('After');

    expect(run.getText()).toBe('Before\rAfter');
    expect(run.getPlainText()).toBe('BeforeAfter');
  });

  it('returns empty string for empty run', () => {
    const run = new Run('');
    expect(run.getPlainText()).toBe('');
  });

  it('returns empty string for run with only special chars', () => {
    const run = new Run('');
    run.addTab();
    run.addBreak();
    run.addCarriageReturn();

    expect(run.getPlainText()).toBe('');
    expect(run.getText()).toBe('\t\n\r'); // getText still returns them
  });

  it('handles multiple text segments', () => {
    const run = new Run('');
    run.appendText('A');
    run.appendText('B');
    run.appendText('C');

    expect(run.getPlainText()).toBe('ABC');
  });

  it('matches getText for runs without special characters', () => {
    const run = new Run('Simple text without specials');

    expect(run.getPlainText()).toBe(run.getText());
  });

  it('useful for word counting', () => {
    const run = new Run('');
    run.appendText('Hello');
    run.addTab();
    run.appendText('World');
    run.addBreak();
    run.appendText('Test');

    const plain = run.getPlainText();
    // Plain text is "HelloWorldTest" — no whitespace-like separators
    expect(plain).toBe('HelloWorldTest');

    // getText gives "Hello\tWorld\nTest" — includes separators
    const words = run.getText().trim().split(/\s+/);
    expect(words).toHaveLength(3);
  });
});

// ============================================================================
// Integration tests
// ============================================================================

describe('combined convenience usage', () => {
  it('builds a formatted table with background colors and breaks', () => {
    const table = Table.fromArray([
      ['Header 1', 'Header 2'],
      ['Data 1', 'Data 2'],
      ['Data 3', 'Data 4'],
    ]);

    // Color header row
    const headerRow = table.getRow(0)!;
    for (const cell of headerRow.getCells()) {
      cell.setBackgroundColor('4472C4');
    }

    // Color alternate data rows
    table
      .getRow(2)!
      .getCells()
      .forEach((c) => c.setBackgroundColor('D9E2F3'));

    expect(table.getCell(0, 0)!.getBackgroundColor()).toBe('4472C4');
    expect(table.getCell(1, 0)!.getBackgroundColor()).toBeUndefined();
    expect(table.getCell(2, 0)!.getBackgroundColor()).toBe('D9E2F3');
  });

  it('creates multi-line paragraph content', () => {
    const para = new Paragraph()
      .addText('Company Name', { bold: true })
      .addLineBreak()
      .addText('123 Main Street')
      .addLineBreak()
      .addText('City, State 12345');

    expect(para.getText()).toBe('Company Name\n123 Main Street\nCity, State 12345');

    // All runs should be part of the same paragraph
    const runs = para.getRuns();
    expect(runs.length).toBeGreaterThanOrEqual(3);
  });
});
