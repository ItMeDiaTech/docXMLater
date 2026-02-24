/**
 * Gap Tests for Run Properties
 *
 * Tests implemented-but-undertested Run.ts formatting properties:
 * - setAllCaps() / setSmallCaps()
 * - setStrike() (and dstrike via formatting)
 * - setHighlight() with various colors
 * - setWebHidden()
 * - setComplexScript()
 * - setSubscript() / setSuperscript()
 *
 * Note: RTL, vanish, noProof, snapToGrid, specVanish, fitText, eastAsianLayout
 * are already covered in RunAdvancedProperties.test.ts
 */

import { Run } from '../../src/elements/Run';
import { Paragraph } from '../../src/elements/Paragraph';
import { Document } from '../../src/core/Document';

describe('Run Properties Gap Tests', () => {
  describe('AllCaps (w:caps)', () => {
    test('should set allCaps', () => {
      const run = new Run('uppercase');
      run.setAllCaps();
      expect(run.getFormatting().allCaps).toBe(true);
    });

    test('should round-trip allCaps', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText('uppercase', { allCaps: true });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getRuns()[0]?.getFormatting().allCaps).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('SmallCaps (w:smallCaps)', () => {
    test('should set smallCaps', () => {
      const run = new Run('small caps');
      run.setSmallCaps();
      expect(run.getFormatting().smallCaps).toBe(true);
    });

    test('should round-trip smallCaps', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText('small caps', { smallCaps: true });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getRuns()[0]?.getFormatting().smallCaps).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Strikethrough (w:strike)', () => {
    test('should set strike', () => {
      const run = new Run('deleted');
      run.setStrike();
      expect(run.getFormatting().strike).toBe(true);
    });

    test('should round-trip strike', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText('strikethrough', { strike: true });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getRuns()[0]?.getFormatting().strike).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Double Strikethrough (w:dstrike)', () => {
    test('should set dstrike via formatting object', () => {
      const run = Run.createFromContent([{ type: 'text', value: 'double strike' }], {
        dstrike: true,
      });
      expect(run.getFormatting().dstrike).toBe(true);
    });

    test('should generate w:dstrike in XML', () => {
      const run = Run.createFromContent([{ type: 'text', value: 'double strike' }], {
        dstrike: true,
      });
      const xml = run.toXML();
      const rPr = (xml.children as any[])?.find((c: any) => c?.name === 'w:rPr');
      const dstrike = rPr?.children?.find((c: any) => c?.name === 'w:dstrike');
      expect(dstrike).toBeDefined();
    });
  });

  describe('Highlight (w:highlight)', () => {
    const highlightColors = [
      'yellow',
      'green',
      'cyan',
      'magenta',
      'blue',
      'red',
      'darkBlue',
      'darkCyan',
      'darkGreen',
      'darkMagenta',
      'darkRed',
      'darkYellow',
      'darkGray',
      'lightGray',
      'black',
      'white',
    ] as const;

    test('should set highlight color', () => {
      const run = new Run('highlighted');
      run.setHighlight('yellow');
      expect(run.getFormatting().highlight).toBe('yellow');
    });

    test('should round-trip all 16 highlight colors', async () => {
      for (const color of highlightColors) {
        const doc = Document.create();
        const para = new Paragraph();
        para.addText(`${color} highlight`, { highlight: color });
        doc.addParagraph(para);

        const buffer = await doc.toBuffer();
        const loaded = await Document.loadFromBuffer(buffer);
        const fmt = loaded.getParagraphs()[0]?.getRuns()[0]?.getFormatting();
        expect(fmt?.highlight).toBe(color);

        doc.dispose();
        loaded.dispose();
      }
    });
  });

  describe('WebHidden (w:webHidden)', () => {
    test('should set webHidden', () => {
      const run = new Run('web hidden');
      run.setWebHidden();
      expect(run.getFormatting().webHidden).toBe(true);
    });

    test('should round-trip webHidden', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText('web hidden', { webHidden: true });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getRuns()[0]?.getFormatting().webHidden).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('ComplexScript (w:cs)', () => {
    test('should set complexScript', () => {
      const run = new Run('complex');
      run.setComplexScript();
      expect(run.getFormatting().complexScript).toBe(true);
    });

    test('should round-trip complexScript', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText('complex script', { complexScript: true });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getParagraphs()[0]?.getRuns()[0]?.getFormatting().complexScript).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Subscript / Superscript (w:vertAlign)', () => {
    test('should set subscript', () => {
      const run = new Run('2');
      run.setSubscript();
      expect(run.getFormatting().subscript).toBe(true);
    });

    test('should set superscript', () => {
      const run = new Run('2');
      run.setSuperscript();
      expect(run.getFormatting().superscript).toBe(true);
    });

    test('should clear superscript when setting subscript', () => {
      const run = new Run('2');
      run.setSuperscript();
      run.setSubscript();
      expect(run.getFormatting().subscript).toBe(true);
      expect(run.getFormatting().superscript).toBeFalsy();
    });

    test('should clear subscript when setting superscript', () => {
      const run = new Run('2');
      run.setSubscript();
      run.setSuperscript();
      expect(run.getFormatting().superscript).toBe(true);
      expect(run.getFormatting().subscript).toBeFalsy();
    });

    test('should round-trip subscript', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText('H');
      para.addText('2', { subscript: true });
      para.addText('O');
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const runs = loaded.getParagraphs()[0]?.getRuns();
      expect(runs?.[1]?.getFormatting().subscript).toBe(true);

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip superscript', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText('E=mc');
      para.addText('2', { superscript: true });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const runs = loaded.getParagraphs()[0]?.getRuns();
      expect(runs?.[1]?.getFormatting().superscript).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Combined Run Properties', () => {
    test('should round-trip multiple properties', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText('Formatted Text', {
        allCaps: true,
        strike: true,
        highlight: 'cyan',
        complexScript: true,
      });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getParagraphs()[0]?.getRuns()[0]?.getFormatting();

      expect(fmt?.allCaps).toBe(true);
      expect(fmt?.strike).toBe(true);
      expect(fmt?.highlight).toBe('cyan');
      expect(fmt?.complexScript).toBe(true);

      doc.dispose();
      loaded.dispose();
    });

    test('should support method chaining for all gap properties', () => {
      const run = new Run('Chained')
        .setAllCaps()
        .setStrike()
        .setHighlight('yellow')
        .setWebHidden()
        .setComplexScript()
        .setSubscript();

      const fmt = run.getFormatting();
      expect(fmt.allCaps).toBe(true);
      expect(fmt.strike).toBe(true);
      expect(fmt.highlight).toBe('yellow');
      expect(fmt.webHidden).toBe(true);
      expect(fmt.complexScript).toBe(true);
      expect(fmt.subscript).toBe(true);
    });
  });
});
