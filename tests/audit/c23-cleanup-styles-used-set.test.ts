/**
 * CleanupHelper.cleanupStyles() must build a complete used-style set before
 * removing anything. Styles are "in use" without any body paragraph carrying
 * the ID: tables reference them via w:tblStyle (ECMA-376 §17.4.62), numbering
 * levels via w:pStyle (§17.9.23), headers/footers/footnotes live outside the
 * body walk, and basedOn/link/next chains plus built-in defaults like Normal
 * and TableGrid resolve implicitly. Removing any of these leaves dangling
 * references and lost formatting in programmatically created documents.
 */
import { Document } from '../../src/core/Document';
import { CleanupHelper } from '../../src/helpers/CleanupHelper';
import { Header } from '../../src/elements/Header';
import { Paragraph } from '../../src/elements/Paragraph';
import { AbstractNumbering } from '../../src/formatting/AbstractNumbering';
import { NumberingInstance } from '../../src/formatting/NumberingInstance';

const JSZip = require('jszip');

describe('C23: cleanupStyles preserves all in-use and protected styles', () => {
  let doc: Document;

  beforeEach(() => {
    doc = Document.create();
  });

  afterEach(() => {
    doc.dispose();
  });

  function styleIds(): string[] {
    return doc
      .getStylesManager()
      .getAllStyles()
      .map((s) => s.getStyleId());
  }

  it('keeps table styles referenced via w:tblStyle', () => {
    const stylesManager = doc.getStylesManager();
    stylesManager.getStyle('TableGrid'); // lazy-load built-in into the map
    const table = doc.createTable(2, 2);
    table.setStyle('TableGrid');

    new CleanupHelper(doc).run({ cleanupStyles: true });

    expect(styleIds()).toContain('TableGrid');
  });

  it('keeps built-in defaults like Normal even when unreferenced', () => {
    doc.addParagraph(new Paragraph());

    new CleanupHelper(doc).run({ cleanupStyles: true });

    expect(styleIds()).toContain('Normal');
  });

  it('keeps basedOn ancestors of used custom styles', () => {
    const stylesManager = doc.getStylesManager();
    stylesManager.createParagraphStyle('BaseX', 'Base X');
    stylesManager.createParagraphStyle('MidX', 'Mid X', 'BaseX');
    stylesManager.createParagraphStyle('LeafX', 'Leaf X', 'MidX');
    const para = new Paragraph();
    para.setStyle('LeafX');
    doc.addParagraph(para);

    new CleanupHelper(doc).run({ cleanupStyles: true });

    const ids = styleIds();
    expect(ids).toContain('LeafX');
    expect(ids).toContain('MidX');
    expect(ids).toContain('BaseX');
  });

  it('keeps linked character styles of used paragraph styles', () => {
    const stylesManager = doc.getStylesManager();
    const paraStyle = stylesManager.createParagraphStyle('LinkedParaX', 'Linked Para X');
    stylesManager.createCharacterStyle('LinkedCharX', 'Linked Char X');
    paraStyle.setLink('LinkedCharX');
    const para = new Paragraph();
    para.setStyle('LinkedParaX');
    doc.addParagraph(para);

    new CleanupHelper(doc).run({ cleanupStyles: true });

    expect(styleIds()).toContain('LinkedCharX');
  });

  it('keeps styles used only in headers', () => {
    doc.getStylesManager().createParagraphStyle('HdrStyleX', 'Header Style X');
    const header = new Header();
    const para = new Paragraph();
    para.setStyle('HdrStyleX');
    para.addText('Header text');
    header.addParagraph(para);
    doc.setHeader(header);

    new CleanupHelper(doc).run({ cleanupStyles: true });

    expect(styleIds()).toContain('HdrStyleX');
  });

  it('keeps styles used only in footnotes', () => {
    doc.getStylesManager().createParagraphStyle('FootStyleX', 'Footnote Style X');
    const footnote = doc.createFootnote('Footnote text');
    footnote.getParagraphs()[0]!.setStyle('FootStyleX');

    new CleanupHelper(doc).run({ cleanupStyles: true });

    expect(styleIds()).toContain('FootStyleX');
  });

  it('keeps styles referenced by numbering levels via w:pStyle', () => {
    doc.getStylesManager().createParagraphStyle('ListStyleX', 'List Style X');
    const abstractNum = AbstractNumbering.createBulletList(50);
    abstractNum.getAllLevels()[0]!.setParagraphStyle('ListStyleX');
    doc.getNumberingManager().addAbstractNumbering(abstractNum);
    doc.getNumberingManager().addInstance(new NumberingInstance(50, 50));

    new CleanupHelper(doc).run({ cleanupStyles: true });

    expect(styleIds()).toContain('ListStyleX');
  });

  it('still removes genuinely unused custom styles', () => {
    doc.getStylesManager().createParagraphStyle('OrphanX', 'Orphan X');
    doc.addParagraph(new Paragraph());

    const report = new CleanupHelper(doc).run({ cleanupStyles: true });

    expect(report.stylesRemoved).toBeGreaterThanOrEqual(1);
    expect(styleIds()).not.toContain('OrphanX');
  });

  it('round-trips a created document through all() without dropping Normal or TableGrid', async () => {
    doc.getStylesManager().getStyle('TableGrid');
    const heading = new Paragraph();
    heading.setStyle('Heading1');
    heading.addText('Title');
    doc.addParagraph(heading);
    const table = doc.createTable(2, 2);
    table.setStyle('TableGrid');

    new CleanupHelper(doc).all();

    const buffer = await doc.toBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const stylesXml = await zip.file('word/styles.xml')!.async('string');
    const documentXml = await zip.file('word/document.xml')!.async('string');

    expect(documentXml).toContain('<w:tblStyle w:val="TableGrid"/>');
    expect(stylesXml).toContain('w:styleId="TableGrid"');
    expect(stylesXml).toContain('w:styleId="Normal"');
  });
});
