/**
 * Tests for Paragraph and Run classes
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('Run', () => {
  describe('Basic functionality', () => {
    test('should create a run with text', () => {
      const run = new Run('Hello World');
      expect(run.getText()).toBe('Hello World');
    });

    test('should set text', () => {
      const run = new Run('Original');
      run.setText('Updated');
      expect(run.getText()).toBe('Updated');
    });

    test('should create run with formatting', () => {
      const run = new Run('Bold text', { bold: true });
      expect(run.getFormatting().bold).toBe(true);
    });
  });

  describe('Formatting methods', () => {
    test('should set bold', () => {
      const run = new Run('Text');
      run.setBold();
      expect(run.getFormatting().bold).toBe(true);
    });

    test('should set italic', () => {
      const run = new Run('Text');
      run.setItalic();
      expect(run.getFormatting().italic).toBe(true);
    });

    test('should set underline', () => {
      const run = new Run('Text');
      run.setUnderline();
      expect(run.getFormatting().underline).toBe(true);
    });

    test('should set underline with style', () => {
      const run = new Run('Text');
      run.setUnderline('double');
      expect(run.getFormatting().underline).toBe('double');
    });

    test('should set strike', () => {
      const run = new Run('Text');
      run.setStrike();
      expect(run.getFormatting().strike).toBe(true);
    });

    test('should set subscript', () => {
      const run = new Run('Text');
      run.setSubscript();
      expect(run.getFormatting().subscript).toBe(true);
      expect(run.getFormatting().superscript).toBe(false);
    });

    test('should set superscript', () => {
      const run = new Run('Text');
      run.setSuperscript();
      expect(run.getFormatting().superscript).toBe(true);
      expect(run.getFormatting().subscript).toBe(false);
    });

    test('should toggle between subscript and superscript', () => {
      const run = new Run('Text');
      run.setSubscript();
      expect(run.getFormatting().subscript).toBe(true);
      run.setSuperscript();
      expect(run.getFormatting().superscript).toBe(true);
      expect(run.getFormatting().subscript).toBe(false);
    });

    test('should set font', () => {
      const run = new Run('Text');
      run.setFont('Arial', 12);
      expect(run.getFormatting().font).toBe('Arial');
      expect(run.getFormatting().size).toBe(12);
    });

    test('should set font without size', () => {
      const run = new Run('Text');
      run.setFont('Times New Roman');
      expect(run.getFormatting().font).toBe('Times New Roman');
      expect(run.getFormatting().size).toBeUndefined();
    });

    test('should set size', () => {
      const run = new Run('Text');
      run.setSize(14);
      expect(run.getFormatting().size).toBe(14);
    });

    test('should set color', () => {
      const run = new Run('Text');
      run.setColor('#FF0000');
      expect(run.getFormatting().color).toBe('FF0000');
    });

    test('should set color without hash', () => {
      const run = new Run('Text');
      run.setColor('00FF00');
      expect(run.getFormatting().color).toBe('00FF00');
    });

    test('should set highlight', () => {
      const run = new Run('Text');
      run.setHighlight('yellow');
      expect(run.getFormatting().highlight).toBe('yellow');
    });

    test('should set small caps', () => {
      const run = new Run('Text');
      run.setSmallCaps();
      expect(run.getFormatting().smallCaps).toBe(true);
    });

    test('should set all caps', () => {
      const run = new Run('Text');
      run.setAllCaps();
      expect(run.getFormatting().allCaps).toBe(true);
    });
  });

  describe('Method chaining', () => {
    test('should support method chaining', () => {
      const run = new Run('Text')
        .setBold()
        .setItalic()
        .setColor('FF0000')
        .setSize(14);

      const formatting = run.getFormatting();
      expect(formatting.bold).toBe(true);
      expect(formatting.italic).toBe(true);
      expect(formatting.color).toBe('FF0000');
      expect(formatting.size).toBe(14);
    });
  });

  describe('XML generation', () => {
    test('should generate basic XML', () => {
      const run = new Run('Hello');
      const xml = run.toXML();

      expect(xml.name).toBe('w:r');
      expect(xml.children).toBeDefined();
    });

    test('should generate XML with formatting', () => {
      const run = new Run('Bold', { bold: true });
      const xml = run.toXML();

      const builder = new XMLBuilder();
      builder.element(xml.name, xml.attributes, xml.children);
      const xmlStr = builder.build();

      expect(xmlStr).toContain('<w:b/>');
      expect(xmlStr).toContain('Bold');
    });

    test('should preserve spaces with xml:space attribute', () => {
      const run = new Run('  Text with spaces  ');
      const xml = run.toXML();

      const builder = new XMLBuilder();
      builder.element(xml.name, xml.attributes, xml.children);
      const xmlStr = builder.build();

      expect(xmlStr).toContain('xml:space="preserve"');
    });
  });

  describe('Static methods', () => {
    test('should create run with static method', () => {
      const run = Run.create('Text', { bold: true });
      expect(run.getText()).toBe('Text');
      expect(run.getFormatting().bold).toBe(true);
    });
  });
});

describe('Paragraph', () => {
  describe('Basic functionality', () => {
    test('should create empty paragraph', () => {
      const para = new Paragraph();
      expect(para.getRuns().length).toBe(0);
      expect(para.getText()).toBe('');
    });

    test('should add run', () => {
      const para = new Paragraph();
      const run = new Run('Hello');
      para.addRun(run);

      expect(para.getRuns().length).toBe(1);
      expect(para.getText()).toBe('Hello');
    });

    test('should add text', () => {
      const para = new Paragraph();
      para.addText('Hello');
      para.addText(' World');

      expect(para.getRuns().length).toBe(2);
      expect(para.getText()).toBe('Hello World');
    });

    test('should set text', () => {
      const para = new Paragraph();
      para.addText('First');
      para.addText('Second');
      para.setText('Replaced');

      expect(para.getRuns().length).toBe(1);
      expect(para.getText()).toBe('Replaced');
    });

    test('should get combined text from multiple runs', () => {
      const para = new Paragraph();
      para.addText('Hello ');
      para.addText('World', { bold: true });
      para.addText('!', { italic: true });

      expect(para.getText()).toBe('Hello World!');
    });

    test('should include hyperlink text in getText', () => {
      const para = new Paragraph();
      para.addText('Click ');
      const link = Hyperlink.createExternal('https://example.com', 'here');
      para.addHyperlink(link);
      para.addText(' for more');

      expect(para.getText()).toBe('Click here for more');
    });

    test('should handle hyperlink-only paragraph', () => {
      const para = new Paragraph();
      const link = Hyperlink.createExternal('https://example.com', 'Link Text');
      para.addHyperlink(link);

      expect(para.getText()).toBe('Link Text');
    });

    test('should handle multiple hyperlinks and runs', () => {
      const para = new Paragraph();
      para.addText('See ');
      para.addHyperlink(Hyperlink.createExternal('https://site1.com', 'site 1'));
      para.addText(' and ');
      para.addHyperlink(Hyperlink.createExternal('https://site2.com', 'site 2'));
      para.addText('.');

      expect(para.getText()).toBe('See site 1 and site 2.');
    });

    test('should handle hyperlink before runs', () => {
      const para = new Paragraph();
      para.addHyperlink(Hyperlink.createExternal('https://example.com', 'Link'));
      para.addText(' followed by text');

      expect(para.getText()).toBe('Link followed by text');
    });

    test('should handle internal hyperlinks in getText', () => {
      const para = new Paragraph();
      para.addText('Go to ');
      para.addHyperlink(Hyperlink.createInternal('Section1', 'Section 1'));
      para.addText('.');

      expect(para.getText()).toBe('Go to Section 1.');
    });
  });

  describe('Formatting methods', () => {
    test('should set alignment', () => {
      const para = new Paragraph();
      para.setAlignment('center');
      expect(para.getFormatting().alignment).toBe('center');
    });

    test('should set left indent', () => {
      const para = new Paragraph();
      para.setLeftIndent(720); // 0.5 inch
      expect(para.getFormatting().indentation?.left).toBe(720);
    });

    test('should set right indent', () => {
      const para = new Paragraph();
      para.setRightIndent(360);
      expect(para.getFormatting().indentation?.right).toBe(360);
    });

    test('should set first line indent', () => {
      const para = new Paragraph();
      para.setFirstLineIndent(720);
      expect(para.getFormatting().indentation?.firstLine).toBe(720);
    });

    test('should set space before', () => {
      const para = new Paragraph();
      para.setSpaceBefore(240);
      expect(para.getFormatting().spacing?.before).toBe(240);
    });

    test('should set space after', () => {
      const para = new Paragraph();
      para.setSpaceAfter(240);
      expect(para.getFormatting().spacing?.after).toBe(240);
    });

    test('should set line spacing', () => {
      const para = new Paragraph();
      para.setLineSpacing(360, 'exact');
      expect(para.getFormatting().spacing?.line).toBe(360);
      expect(para.getFormatting().spacing?.lineRule).toBe('exact');
    });

    test('should set style', () => {
      const para = new Paragraph();
      para.setStyle('Heading1');
      expect(para.getFormatting().style).toBe('Heading1');
    });

    test('should set keep next', () => {
      const para = new Paragraph();
      para.setKeepNext();
      expect(para.getFormatting().keepNext).toBe(true);
    });

    test('should set keep lines', () => {
      const para = new Paragraph();
      para.setKeepLines();
      expect(para.getFormatting().keepLines).toBe(true);
    });

    test('should set page break before', () => {
      const para = new Paragraph();
      para.setPageBreakBefore();
      expect(para.getFormatting().pageBreakBefore).toBe(true);
    });
  });

  describe('Method chaining', () => {
    test('should support method chaining', () => {
      const para = new Paragraph()
        .setAlignment('center')
        .setSpaceBefore(240)
        .setSpaceAfter(240)
        .addText('Centered text');

      const formatting = para.getFormatting();
      expect(formatting.alignment).toBe('center');
      expect(formatting.spacing?.before).toBe(240);
      expect(formatting.spacing?.after).toBe(240);
      expect(para.getText()).toBe('Centered text');
    });
  });

  describe('XML generation', () => {
    test('should generate basic XML', () => {
      const para = new Paragraph();
      para.addText('Hello');
      const xml = para.toXML();

      expect(xml.name).toBe('w:p');
      expect(xml.children).toBeDefined();
    });

    test('should generate XML with empty run if no text', () => {
      const para = new Paragraph();
      const xml = para.toXML();

      const builder = new XMLBuilder();
      builder.element(xml.name, xml.attributes, xml.children);
      const xmlStr = builder.build();

      expect(xmlStr).toContain('<w:r>');
    });

    test('should generate XML with alignment', () => {
      const para = new Paragraph();
      para.setAlignment('center');
      para.addText('Centered');
      const xml = para.toXML();

      const builder = new XMLBuilder();
      builder.element(xml.name, xml.attributes, xml.children);
      const xmlStr = builder.build();

      expect(xmlStr).toContain('<w:jc w:val="center"/>');
    });

    test('should generate XML with multiple formatted runs', () => {
      const para = new Paragraph();
      para.addText('Normal ');
      para.addText('Bold', { bold: true });
      para.addText(' Italic', { italic: true });

      const xml = para.toXML();
      const builder = new XMLBuilder();
      builder.element(xml.name, xml.attributes, xml.children);
      const xmlStr = builder.build();

      expect(xmlStr).toContain('Normal');
      expect(xmlStr).toContain('Bold');
      expect(xmlStr).toContain('Italic');
      expect(xmlStr).toContain('<w:b/>');
      expect(xmlStr).toContain('<w:i/>');
    });
  });

  describe('Static methods', () => {
    test('should create paragraph with static method', () => {
      const para = Paragraph.create({ alignment: 'center' });
      expect(para.getFormatting().alignment).toBe('center');
    });
  });
});
