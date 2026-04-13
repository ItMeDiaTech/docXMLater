import { Document, Paragraph, Table, TableRow, TableCell } from '../../src';

describe('Document.toPlainText', () => {
  let doc: Document;

  afterEach(() => {
    doc?.dispose();
  });

  it('should return empty string for empty document', () => {
    doc = Document.create();
    expect(doc.toPlainText()).toBe('');
  });

  it('should extract text from paragraphs', () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('First paragraph'));
    doc.addParagraph(new Paragraph().addText('Second paragraph'));
    doc.addParagraph(new Paragraph().addText('Third paragraph'));

    expect(doc.toPlainText()).toBe('First paragraph\nSecond paragraph\nThird paragraph');
  });

  it('should use custom separator', () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('Hello'));
    doc.addParagraph(new Paragraph().addText('World'));

    expect(doc.toPlainText(' ')).toBe('Hello World');
    expect(doc.toPlainText(' | ')).toBe('Hello | World');
    expect(doc.toPlainText('')).toBe('HelloWorld');
  });

  it('should include text from table cells', () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('Before table'));

    const table = new Table(2, 2);
    const rows = table.getRows();
    rows[0]!.getCells()[0]!.addParagraph(new Paragraph().addText('Cell A1'));
    rows[0]!.getCells()[1]!.addParagraph(new Paragraph().addText('Cell B1'));
    rows[1]!.getCells()[0]!.addParagraph(new Paragraph().addText('Cell A2'));
    rows[1]!.getCells()[1]!.addParagraph(new Paragraph().addText('Cell B2'));
    doc.addTable(table);

    doc.addParagraph(new Paragraph().addText('After table'));

    const text = doc.toPlainText();
    expect(text).toContain('Before table');
    expect(text).toContain('Cell A1');
    expect(text).toContain('Cell B1');
    expect(text).toContain('Cell A2');
    expect(text).toContain('Cell B2');
    expect(text).toContain('After table');
  });

  it('should handle paragraphs with multiple runs', () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('Bold text', { bold: true });
    para.addText(' and normal text');
    doc.addParagraph(para);

    expect(doc.toPlainText()).toBe('Bold text and normal text');
  });

  it('should handle empty paragraphs', () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('First'));
    doc.addParagraph(new Paragraph()); // empty
    doc.addParagraph(new Paragraph().addText('Third'));

    expect(doc.toPlainText()).toBe('First\n\nThird');
  });

  it('should round-trip: create, save, load, extract text', async () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('Round-trip test'));
    doc.addParagraph(new Paragraph().addText('Second line'));

    const buffer = await doc.toBuffer();
    doc.dispose();

    doc = await Document.loadFromBuffer(buffer);
    const text = doc.toPlainText();
    expect(text).toContain('Round-trip test');
    expect(text).toContain('Second line');
  });
});
