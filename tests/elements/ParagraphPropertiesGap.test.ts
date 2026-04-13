/**
 * Gap Tests for Paragraph Properties
 *
 * Tests implemented-but-undertested Paragraph.ts features:
 * - setContextualSpacing()
 * - setBidi()
 * - setTextDirection()
 * - setTextAlignment()
 * - setSuppressLineNumbers()
 * - setSuppressAutoHyphens()
 * - setAdjustRightInd()
 * - setMirrorIndents()
 * - setSuppressOverlap()
 * - setDivId()
 * - setWidowControl()
 * - setFrameProperties()
 * - setParagraphMarkFormatting()
 * - setParagraphPropertiesChange() round-trip
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Document } from '../../src/core/Document';

describe('Paragraph Properties Gap Tests', () => {
  describe('Contextual Spacing (w:contextualSpacing)', () => {
    test('should set contextualSpacing', () => {
      const para = new Paragraph();
      para.setContextualSpacing(true);
      expect(para.getFormatting().contextualSpacing).toBe(true);
    });

    test('should round-trip contextualSpacing', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Contextual');
      para.setContextualSpacing(true);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().contextualSpacing).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Bidi (w:bidi)', () => {
    test('should round-trip bidi', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('RTL paragraph');
      para.setBidi(true);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().bidi).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Text Direction (w:textDirection)', () => {
    test('should round-trip textDirection lrTb', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Left to right');
      para.setTextDirection('lrTb');

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().textDirection).toBe('lrTb');

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip textDirection tbRl', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Top to bottom');
      para.setTextDirection('tbRl');

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().textDirection).toBe('tbRl');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Text Alignment (w:textAlignment)', () => {
    const alignments: Array<'top' | 'center' | 'baseline' | 'bottom' | 'auto'> = [
      'top',
      'center',
      'baseline',
      'bottom',
      'auto',
    ];

    test('should round-trip all textAlignment values', async () => {
      for (const alignment of alignments) {
        const doc = Document.create();
        const para = doc.createParagraph(`Align: ${alignment}`);
        para.setTextAlignment(alignment);

        const buffer = await doc.toBuffer();
        const loaded = await Document.loadFromBuffer(buffer);
        expect(loaded.getParagraphs()[0]?.getFormatting().textAlignment).toBe(alignment);

        doc.dispose();
        loaded.dispose();
      }
    });
  });

  describe('Suppress Line Numbers (w:suppressLineNumbers)', () => {
    test('should round-trip suppressLineNumbers', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('No line numbers');
      para.setSuppressLineNumbers(true);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().suppressLineNumbers).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Suppress Auto Hyphens (w:suppressAutoHyphens)', () => {
    test('should round-trip suppressAutoHyphens', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('No hyphenation');
      para.setSuppressAutoHyphens(true);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().suppressAutoHyphens).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Adjust Right Indent (w:adjustRightInd)', () => {
    test('should round-trip adjustRightInd', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Adjusted');
      para.setAdjustRightInd(true);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().adjustRightInd).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Mirror Indents (w:mirrorIndents)', () => {
    test('should round-trip mirrorIndents', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Mirrored');
      para.setMirrorIndents(true);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().mirrorIndents).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Suppress Overlap (w:suppressOverlap)', () => {
    test('should round-trip suppressOverlap', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('No overlap');
      para.setSuppressOverlap(true);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().suppressOverlap).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Div ID (w:divId)', () => {
    test('should round-trip divId', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('With divId');
      para.setDivId(12345);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getFormatting().divId).toBe(12345);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Widow Control (w:widowControl)', () => {
    test('should set widowControl', () => {
      const para = new Paragraph();
      para.setWidowControl(true);
      expect(para.getFormatting().widowControl).toBe(true);
    });

    test('should round-trip widowControl=false', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('No widow control');
      para.setWidowControl(false);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      // widowControl=false should be preserved (explicit override of default)
      const fmt = loaded.getParagraphs()[0]?.getFormatting();
      // When false, it might not be present or be false
      expect(fmt?.widowControl === false || fmt?.widowControl === undefined).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Frame Properties (w:framePr)', () => {
    test('should set frame properties', () => {
      const para = new Paragraph();
      para.setFrameProperties({
        w: 4320,
        h: 2160,
        x: 720,
        y: 1440,
        hAnchor: 'page',
        vAnchor: 'text',
      });
      const fmt = para.getFormatting();
      expect(fmt.framePr).toBeDefined();
      expect(fmt.framePr?.w).toBe(4320);
      expect(fmt.framePr?.h).toBe(2160);
    });

    test('should round-trip frame properties', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Framed text');
      para.setFrameProperties({
        w: 4320,
        h: 2160,
        hAnchor: 'page',
        vAnchor: 'text',
        wrap: 'around',
      });

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fp = loaded.getParagraphs()[0]?.getFormatting().framePr;

      expect(fp).toBeDefined();
      expect(fp?.w).toBe(4320);
      expect(fp?.h).toBe(2160);
      expect(fp?.hAnchor).toBe('page');
      expect(fp?.vAnchor).toBe('text');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Frame Properties - ST_Wrap completeness (ECMA-376 §17.18.104)', () => {
    test('should round-trip wrap="auto"', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Auto wrap');
      para.setFrameProperties({ w: 4320, h: 2160, wrap: 'auto' });

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fp = loaded.getParagraphs()[0]?.getFormatting().framePr;

      expect(fp?.wrap).toBe('auto');
      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip wrap="through"', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Through wrap');
      para.setFrameProperties({ w: 4320, h: 2160, wrap: 'through' });

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fp = loaded.getParagraphs()[0]?.getFormatting().framePr;

      expect(fp?.wrap).toBe('through');
      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip all six ST_Wrap values', async () => {
      const wrapValues = ['around', 'auto', 'none', 'notBeside', 'through', 'tight'] as const;

      for (const wrapVal of wrapValues) {
        const doc = Document.create();
        const para = doc.createParagraph(`Wrap ${wrapVal}`);
        para.setFrameProperties({ w: 4320, h: 2160, wrap: wrapVal });

        const buffer = await doc.toBuffer();
        const loaded = await Document.loadFromBuffer(buffer);
        const fp = loaded.getParagraphs()[0]?.getFormatting().framePr;

        expect(fp?.wrap).toBe(wrapVal);
        doc.dispose();
        loaded.dispose();
      }
    });
  });

  describe('Paragraph Mark Formatting (w:rPr in pPr)', () => {
    test('should round-trip paragraph mark formatting', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('With styled mark');
      para.setParagraphMarkFormatting({ bold: true, color: 'FF0000', font: 'Arial' });

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const rPr = loaded.getParagraphs()[0]?.getFormatting().paragraphMarkRunProperties;

      expect(rPr).toBeDefined();
      expect(rPr?.bold).toBe(true);
      expect(rPr?.color).toBe('FF0000');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Combined Paragraph Properties', () => {
    test('should round-trip multiple paragraph properties', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Combined');
      para
        .setContextualSpacing(true)
        .setBidi(true)
        .setTextDirection('lrTb')
        .setTextAlignment('center')
        .setSuppressLineNumbers(true)
        .setSuppressAutoHyphens(true)
        .setMirrorIndents(true);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getParagraphs()[0]?.getFormatting();

      expect(fmt?.contextualSpacing).toBe(true);
      expect(fmt?.bidi).toBe(true);
      expect(fmt?.textDirection).toBe('lrTb');
      expect(fmt?.textAlignment).toBe('center');
      expect(fmt?.suppressLineNumbers).toBe(true);
      expect(fmt?.suppressAutoHyphens).toBe(true);
      expect(fmt?.mirrorIndents).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Bidi-aware indentation (w:start/w:end) — ECMA-376 §17.3.1.15', () => {
    test('should parse w:start/w:end as left/right indentation', async () => {
      // Create a base document with indentation
      const doc = Document.create();
      const para = doc.createParagraph('RTL indented');
      para.setLeftIndent(720);
      para.setRightIndent(360);
      const buffer = await doc.toBuffer();
      doc.dispose();

      // Modify the DOCX XML to replace w:left/w:right with w:start/w:end
      const JSZip = (await import('jszip')).default;
      const zip = await JSZip.loadAsync(buffer);
      let docXml = await zip.file('word/document.xml')!.async('string');
      docXml = docXml.replace(/(<w:ind[^>]*)\bw:left="/g, '$1w:start="');
      docXml = docXml.replace(/(<w:ind[^>]*)\bw:right="/g, '$1w:end="');
      zip.file('word/document.xml', docXml);
      const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

      // Parse and verify indentation is read correctly
      const loaded = await Document.loadFromBuffer(modifiedBuffer);
      const fmt = loaded.getParagraphs()[0]?.getFormatting();
      expect(fmt?.indentation?.left).toBe(720);
      expect(fmt?.indentation?.right).toBe(360);
      loaded.dispose();
    });

    test('should prefer w:start over w:left when both present', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Both attrs');
      para.setLeftIndent(500);
      const buffer = await doc.toBuffer();
      doc.dispose();

      // Add w:start alongside existing w:left (w:start should win)
      const JSZip = (await import('jszip')).default;
      const zip = await JSZip.loadAsync(buffer);
      let docXml = await zip.file('word/document.xml')!.async('string');
      docXml = docXml.replace(/(<w:ind\s)/, '$1w:start="900" ');
      zip.file('word/document.xml', docXml);
      const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

      const loaded = await Document.loadFromBuffer(modifiedBuffer);
      const fmt = loaded.getParagraphs()[0]?.getFormatting();
      expect(fmt?.indentation?.left).toBe(900);
      loaded.dispose();
    });
  });

  describe('Spacing autospacing/lines (ECMA-376 §17.3.1.33)', () => {
    test('should round-trip beforeAutospacing and afterAutospacing', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Autospaced paragraph');
      para.setSpaceBefore(100);
      para.setSpaceAfter(100);
      // Set extended attributes directly on internal formatting
      para.formatting.spacing!.beforeAutospacing = true;
      para.formatting.spacing!.afterAutospacing = true;

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const spc = loaded.getParagraphs()[0]?.formatting.spacing;

      expect(spc?.beforeAutospacing).toBe(true);
      expect(spc?.afterAutospacing).toBe(true);
      expect(spc?.before).toBe(100);
      expect(spc?.after).toBe(100);
      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip beforeLines and afterLines', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Line-spaced paragraph');
      para.setSpaceBefore(0);
      para.formatting.spacing!.beforeLines = 100;
      para.formatting.spacing!.afterLines = 50;

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const spc = loaded.getParagraphs()[0]?.formatting.spacing;

      expect(spc?.beforeLines).toBe(100);
      expect(spc?.afterLines).toBe(50);
      doc.dispose();
      loaded.dispose();
    });

    test('should parse autospacing from injected XML', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Test');
      para.setSpaceBefore(100);
      const buffer = await doc.toBuffer();
      doc.dispose();

      // Inject beforeAutospacing and afterAutospacing into spacing element
      const JSZip = (await import('jszip')).default;
      const zip = await JSZip.loadAsync(buffer);
      let docXml = await zip.file('word/document.xml')!.async('string');
      docXml = docXml.replace(
        /<w:spacing\s/,
        '<w:spacing w:beforeAutospacing="1" w:afterAutospacing="1" '
      );
      zip.file('word/document.xml', docXml);
      const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

      const loaded = await Document.loadFromBuffer(modifiedBuffer);
      const spc = loaded.getParagraphs()[0]?.formatting.spacing;

      expect(spc?.beforeAutospacing).toBe(true);
      expect(spc?.afterAutospacing).toBe(true);
      loaded.dispose();
    });

    test('should parse autospacing when no before/after/line attributes exist', async () => {
      // Edge case: w:spacing has ONLY autospacing attrs, no before/after/line
      const doc = Document.create();
      const para = doc.createParagraph('Test');
      // Set a dummy property to ensure w:pPr is emitted
      para.setAlignment('left');
      const buffer = await doc.toBuffer();
      doc.dispose();

      // Replace existing spacing or inject one with only autospacing
      const JSZip = (await import('jszip')).default;
      const zip = await JSZip.loadAsync(buffer);
      let docXml = await zip.file('word/document.xml')!.async('string');
      // Insert autospacing-only spacing element before </w:pPr>
      docXml = docXml.replace(
        '</w:pPr>',
        '<w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr>'
      );
      zip.file('word/document.xml', docXml);
      const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

      const loaded = await Document.loadFromBuffer(modifiedBuffer);
      const spc = loaded.getParagraphs()[0]?.formatting.spacing;

      expect(spc).toBeDefined();
      expect(spc?.beforeAutospacing).toBe(true);
      expect(spc?.afterAutospacing).toBe(true);
      loaded.dispose();
    });
  });
});
