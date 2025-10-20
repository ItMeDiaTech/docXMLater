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

  describe('Detached paragraphs', () => {
    describe('Paragraph.create()', () => {
      test('should create empty detached paragraph', () => {
        const para = Paragraph.create();
        expect(para).toBeInstanceOf(Paragraph);
        expect(para.getText()).toBe('');
        expect(para.getContent()).toHaveLength(0);
      });

      test('should create detached paragraph with text', () => {
        const para = Paragraph.create('Hello World');
        expect(para.getText()).toBe('Hello World');
        expect(para.getRuns()).toHaveLength(1);
      });

      test('should create detached paragraph with text and formatting', () => {
        const para = Paragraph.create('Centered Text', { alignment: 'center' });
        expect(para.getText()).toBe('Centered Text');
        expect(para.getFormatting().alignment).toBe('center');
      });

      test('should create detached paragraph with just formatting', () => {
        const para = Paragraph.create({ alignment: 'right', style: 'Heading1' });
        expect(para.getText()).toBe('');
        expect(para.getFormatting().alignment).toBe('right');
        expect(para.getStyle()).toBe('Heading1');
      });

      test('should support method chaining after creation', () => {
        const para = Paragraph.create('Initial')
          .addText(' text', { bold: true })
          .setAlignment('justify')
          .setSpaceBefore(240);

        expect(para.getText()).toBe('Initial text');
        expect(para.getRuns()).toHaveLength(2);
        expect(para.getFormatting().alignment).toBe('justify');
        expect(para.getFormatting().spacing?.before).toBe(240);
      });
    });

    describe('Paragraph.createWithStyle()', () => {
      test('should create detached paragraph with style', () => {
        const para = Paragraph.createWithStyle('Heading Text', 'Heading1');
        expect(para.getText()).toBe('Heading Text');
        expect(para.getStyle()).toBe('Heading1');
      });

      test('should support method chaining with styled paragraph', () => {
        const para = Paragraph.createWithStyle('Title', 'Title')
          .setAlignment('center')
          .addText(' - Subtitle', { italic: true });

        expect(para.getText()).toBe('Title - Subtitle');
        expect(para.getStyle()).toBe('Title');
        expect(para.getFormatting().alignment).toBe('center');
      });
    });

    describe('Paragraph.createEmpty()', () => {
      test('should create empty detached paragraph', () => {
        const para = Paragraph.createEmpty();
        expect(para.getText()).toBe('');
        expect(para.getContent()).toHaveLength(0);
      });

      test('should allow adding content after creation', () => {
        const para = Paragraph.createEmpty()
          .addText('Added later')
          .setAlignment('center');

        expect(para.getText()).toBe('Added later');
        expect(para.getFormatting().alignment).toBe('center');
      });
    });

    describe('Paragraph.createFormatted()', () => {
      test('should create detached paragraph with run and paragraph formatting', () => {
        const para = Paragraph.createFormatted(
          'Important Text',
          { bold: true, color: 'FF0000' },
          { alignment: 'center' }
        );

        expect(para.getText()).toBe('Important Text');
        expect(para.getFormatting().alignment).toBe('center');

        const run = para.getRuns()[0];
        expect(run).toBeDefined();
        expect(run!.getFormatting().bold).toBe(true);
        expect(run!.getFormatting().color).toBe('FF0000');
      });

      test('should work without paragraph formatting', () => {
        const para = Paragraph.createFormatted(
          'Bold Text',
          { bold: true }
        );

        expect(para.getText()).toBe('Bold Text');
        const run = para.getRuns()[0];
        expect(run).toBeDefined();
        expect(run!.getFormatting().bold).toBe(true);
      });

      test('should work without run formatting', () => {
        const para = Paragraph.createFormatted(
          'Plain Text',
          undefined,
          { alignment: 'right' }
        );

        expect(para.getText()).toBe('Plain Text');
        expect(para.getFormatting().alignment).toBe('right');
      });
    });

    describe('Complex detached paragraph scenarios', () => {
      test('should build complex paragraph with multiple runs', () => {
        const para = Paragraph.create()
          .addText('Normal ', {})
          .addText('bold ', { bold: true })
          .addText('italic ', { italic: true })
          .addText('both', { bold: true, italic: true })
          .setAlignment('justify');

        expect(para.getText()).toBe('Normal bold italic both');
        expect(para.getRuns()).toHaveLength(4);
        expect(para.getFormatting().alignment).toBe('justify');
      });

      test('should clone detached paragraph', () => {
        const original = Paragraph.create('Original Text', { alignment: 'center' });
        const clone = original.clone();

        // Verify clone has same content
        expect(clone.getText()).toBe('Original Text');
        expect(clone.getFormatting().alignment).toBe('center');

        // Verify they are independent
        clone.addText(' - Modified');
        expect(original.getText()).toBe('Original Text');
        expect(clone.getText()).toBe('Original Text - Modified');
      });

      test('should set complex formatting on detached paragraph', () => {
        const para = Paragraph.create()
          .addText('Complex Paragraph')
          .setAlignment('center')
          .setLeftIndent(720)
          .setRightIndent(720)
          .setSpaceBefore(240)
          .setSpaceAfter(240)
          .setLineSpacing(360, 'exact')
          .setKeepNext(true)
          .setKeepLines(true);

        const formatting = para.getFormatting();
        expect(formatting.alignment).toBe('center');
        expect(formatting.indentation?.left).toBe(720);
        expect(formatting.indentation?.right).toBe(720);
        expect(formatting.spacing?.before).toBe(240);
        expect(formatting.spacing?.after).toBe(240);
        expect(formatting.spacing?.line).toBe(360);
        expect(formatting.spacing?.lineRule).toBe('exact');
        expect(formatting.keepNext).toBe(true);
        expect(formatting.keepLines).toBe(true);
      });

      test('should generate proper XML from detached paragraph', () => {
        const para = Paragraph.create('Test Content', { alignment: 'center' });
        const xml = para.toXML();

        expect(xml.name).toBe('w:p');
        expect(xml.children).toBeDefined();

        // Should have paragraph properties and at least one run
        const children = xml.children || [];
        expect(children.length).toBeGreaterThan(0);
      });
    });
  });
});
