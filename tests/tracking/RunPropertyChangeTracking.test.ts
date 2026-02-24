import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('Run Property Change Tracking', () => {
  // Regression: Bug fix — rPrChange must not be emitted as self-closing
  // Per ECMA-376 §17.13.5.32, rPrChange requires a child w:rPr (minOccurs=1).
  // When all previous properties are undefined (inherited from styles),
  // the output must omit rPrChange entirely rather than emit an empty one.
  describe('rPrChange XML output', () => {
    it('should NOT emit rPrChange when previous properties are all undefined', () => {
      const doc = Document.create();
      const para = new Paragraph();
      doc.addParagraph(para);
      // Run starts with NO explicit formatting (all inherited from style)
      const run = new Run('Test text');
      para.addRun(run);

      doc.enableTrackChanges({ author: 'TestAuthor', trackFormatting: true });
      // Apply formatting to a run that had no explicit formatting
      run.setFont('Verdana');
      run.setSize(14);
      run.setBold(true);
      run.setColor('000000');
      doc.flushPendingChanges();

      // Run should have a property change revision with undefined previous values
      const propChange = run.getPropertyChangeRevision();
      expect(propChange).toBeDefined();

      // But the XML output should NOT contain rPrChange
      // because all previous properties were undefined
      const xml = run.toXML();
      const xmlStr = XMLBuilder.elementToString(xml);
      expect(xmlStr).not.toContain('w:rPrChange');
    });

    it('should emit rPrChange with child w:rPr when previous properties are defined', () => {
      const doc = Document.create();
      const para = new Paragraph();
      doc.addParagraph(para);
      const run = new Run('Test text');
      run.setFont('Arial');
      run.setSize(10);
      run.setColor('FF0000');
      para.addRun(run);

      doc.enableTrackChanges({ author: 'TestAuthor', trackFormatting: true });
      run.setFont('Verdana');
      run.setSize(14);
      run.setColor('000000');
      doc.flushPendingChanges();

      const xml = run.toXML();
      const xmlStr = XMLBuilder.elementToString(xml);
      // rPrChange should be present
      expect(xmlStr).toContain('w:rPrChange');
      // It must contain a child w:rPr (not self-closing)
      expect(xmlStr).toMatch(/w:rPrChange[^/]*>.*<w:rPr/s);
      // Previous font should be recorded
      expect(xmlStr).toContain('w:ascii="Arial"');
    });

    it('should survive save/load round-trip without creating empty rPrChange', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      doc.addParagraph(para);
      const run = new Run('Round-trip test');
      para.addRun(run);

      doc.enableTrackChanges({ author: 'TestAuthor', trackFormatting: true });
      run.setFont('Verdana');
      run.setBold(true);
      doc.flushPendingChanges();

      // Save and reload
      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Verify no empty rPrChange in the saved XML
      const loadedRun = loaded.getParagraphs()[0]?.getRuns()[0];
      expect(loadedRun).toBeDefined();
      const xml = loadedRun!.toXML();
      const xmlStr = XMLBuilder.elementToString(xml);
      // Should NOT contain self-closing rPrChange
      expect(xmlStr).not.toMatch(/<w:rPrChange[^>]*\/>/);

      loaded.dispose();
    });
  });

  it('should create rPrChange when run font is changed with tracking enabled', () => {
    const doc = Document.create();
    const para = new Paragraph();
    doc.addParagraph(para);
    const run = new Run('Hello world');
    run.setFont('Arial');
    para.addRun(run);

    doc.enableTrackChanges({ author: 'TestAuthor', trackFormatting: true });
    run.setFont('Verdana');
    doc.flushPendingChanges();

    const propChange = run.getPropertyChangeRevision();
    expect(propChange).toBeDefined();
    expect(propChange!.author).toBe('TestAuthor');
    expect(propChange!.previousProperties).toBeDefined();
    expect(propChange!.previousProperties.font).toBe('Arial');
  });

  it('should create rPrChange with multiple changed properties', () => {
    const doc = Document.create();
    const para = new Paragraph();
    doc.addParagraph(para);
    const run = new Run('Hello');
    run.setFont('Arial');
    run.setSize(10);
    run.setColor('FF0000');
    para.addRun(run);

    doc.enableTrackChanges({ author: 'TestAuthor', trackFormatting: true });
    run.setFont('Verdana');
    run.setSize(12);
    run.setColor('000000');
    doc.flushPendingChanges();

    const propChange = run.getPropertyChangeRevision();
    expect(propChange).toBeDefined();
    expect(propChange!.previousProperties.font).toBe('Arial');
    expect(propChange!.previousProperties.size).toBe(10);
    expect(propChange!.previousProperties.color).toBe('FF0000');
  });

  it('should NOT create rPrChange when value does not change', () => {
    const doc = Document.create();
    const para = new Paragraph();
    doc.addParagraph(para);
    const run = new Run('Hello');
    run.setFont('Verdana');
    para.addRun(run);

    doc.enableTrackChanges({ author: 'TestAuthor', trackFormatting: true });
    run.setFont('Verdana');
    doc.flushPendingChanges();

    const propChange = run.getPropertyChangeRevision();
    expect(propChange).toBeUndefined();
  });

  it('should merge with existing rPrChange on the run', () => {
    const doc = Document.create();
    const para = new Paragraph();
    doc.addParagraph(para);
    const run = new Run('Hello');
    run.setFont('Arial');
    run.setBold(true);
    para.addRun(run);

    run.setPropertyChangeRevision({
      id: 99,
      author: 'OriginalAuthor',
      date: new Date('2025-01-01'),
      previousProperties: { bold: false },
    });

    doc.enableTrackChanges({
      author: 'TestAuthor',
      trackFormatting: true,
      clearExistingPropertyChanges: false,
    });

    run.setFont('Verdana');
    doc.flushPendingChanges();

    const propChange = run.getPropertyChangeRevision();
    expect(propChange).toBeDefined();
    // Verify original author/date preserved during merge
    expect(propChange!.author).toBe('OriginalAuthor');
    expect(propChange!.date).toEqual(new Date('2025-01-01'));
    // Verify previous properties merged
    expect(propChange!.previousProperties.bold).toBe(false);
    expect(propChange!.previousProperties.font).toBe('Arial');
  });
});
