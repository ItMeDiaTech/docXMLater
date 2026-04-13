import { Document, Paragraph, Table } from '../../src';

describe('Document.toJSON', () => {
  let doc: Document;

  afterEach(() => {
    doc?.dispose();
  });

  it('should return document structure for empty document', () => {
    doc = Document.create();
    const json = doc.toJSON();

    expect(json.properties).toBeDefined();
    expect(json.stats.paragraphs).toBe(0);
    expect(json.stats.tables).toBe(0);
    expect(json.stats.images).toBe(0);
    expect(json.stats.headings).toBe(0);
    expect(json.headings).toEqual([]);
    expect(json.body).toEqual([]);
  });

  it('should include paragraph content', () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('Hello world'));
    doc.addParagraph(new Paragraph().addText('Second paragraph'));

    const json = doc.toJSON();
    expect(json.stats.paragraphs).toBe(2);
    expect(json.body.length).toBe(2);
    expect(json.body[0]!.type).toBe('paragraph');
    expect(json.body[0]!.text).toBe('Hello world');
    expect(json.body[1]!.text).toBe('Second paragraph');
  });

  it('should include heading hierarchy', () => {
    doc = Document.create();

    const h1 = new Paragraph().addText('Title');
    h1.setStyle('Heading1');
    doc.addParagraph(h1);

    doc.addParagraph(new Paragraph().addText('Body'));

    const h2 = new Paragraph().addText('Section');
    h2.setStyle('Heading2');
    doc.addParagraph(h2);

    const json = doc.toJSON();
    expect(json.stats.headings).toBe(2);
    expect(json.headings).toEqual([
      { level: 1, text: 'Title' },
      { level: 2, text: 'Section' },
    ]);
  });

  it('should include table info', () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('Before'));
    doc.addTable(new Table(3, 4));
    doc.addParagraph(new Paragraph().addText('After'));

    const json = doc.toJSON();
    expect(json.stats.tables).toBeGreaterThanOrEqual(1);
    const tableEntry = json.body.find((b) => b.type === 'table');
    expect(tableEntry).toBeDefined();
    expect(tableEntry!.text).toContain('3 rows');
  });

  it('should include style information', () => {
    doc = Document.create();
    const para = new Paragraph().addText('Styled text');
    para.setStyle('Heading1');
    doc.addParagraph(para);

    const json = doc.toJSON();
    expect(json.body[0]!.style).toBe('Heading1');
  });

  it('should be JSON-serializable', () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('Test'));

    const json = doc.toJSON();
    const serialized = JSON.stringify(json);
    const parsed = JSON.parse(serialized);

    expect(parsed.stats.paragraphs).toBe(1);
    expect(parsed.body[0].text).toBe('Test');
  });

  it('should include document properties', () => {
    doc = Document.create();
    doc.setTitle('Test Document');

    const json = doc.toJSON();
    expect(json.properties.title).toBe('Test Document');
  });
});
