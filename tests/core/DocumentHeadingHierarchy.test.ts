import { Document, Paragraph } from '../../src';

describe('Document.getHeadingHierarchy', () => {
  let doc: Document;

  afterEach(() => {
    doc?.dispose();
  });

  it('should return empty array when no headings exist', () => {
    doc = Document.create();
    doc.addParagraph(new Paragraph().addText('Normal text'));
    expect(doc.getHeadingHierarchy()).toEqual([]);
  });

  it('should detect headings by style', () => {
    doc = Document.create();

    const h1 = new Paragraph().addText('Chapter 1');
    h1.setStyle('Heading1');
    doc.addParagraph(h1);

    doc.addParagraph(new Paragraph().addText('Body text'));

    const h2 = new Paragraph().addText('Section 1.1');
    h2.setStyle('Heading2');
    doc.addParagraph(h2);

    const headings = doc.getHeadingHierarchy();
    expect(headings.length).toBe(2);
    expect(headings[0]!.level).toBe(1);
    expect(headings[0]!.text).toBe('Chapter 1');
    expect(headings[1]!.level).toBe(2);
    expect(headings[1]!.text).toBe('Section 1.1');
  });

  it('should return headings in document order', () => {
    doc = Document.create();

    const h2 = new Paragraph().addText('Sub-section');
    h2.setStyle('Heading2');
    doc.addParagraph(h2);

    const h1 = new Paragraph().addText('Main section');
    h1.setStyle('Heading1');
    doc.addParagraph(h1);

    const headings = doc.getHeadingHierarchy();
    expect(headings.length).toBe(2);
    expect(headings[0]!.text).toBe('Sub-section');
    expect(headings[1]!.text).toBe('Main section');
  });

  it('should detect skipped heading levels', () => {
    doc = Document.create();

    const h1 = new Paragraph().addText('Title');
    h1.setStyle('Heading1');
    doc.addParagraph(h1);

    // Skip H2, go straight to H3
    const h3 = new Paragraph().addText('Deep section');
    h3.setStyle('Heading3');
    doc.addParagraph(h3);

    const headings = doc.getHeadingHierarchy();
    expect(headings.length).toBe(2);

    // Detect the skip
    const hasSkip = headings[1]!.level - headings[0]!.level > 1;
    expect(hasSkip).toBe(true);
  });

  it('should include paragraph reference for modification', () => {
    doc = Document.create();
    const h1 = new Paragraph().addText('Original Title');
    h1.setStyle('Heading1');
    doc.addParagraph(h1);

    const headings = doc.getHeadingHierarchy();
    expect(headings[0]!.paragraph).toBe(h1);
  });
});
