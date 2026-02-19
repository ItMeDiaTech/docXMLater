/**
 * ECMA-376 Gap Analysis: Phase B & C Tests
 *
 * Tests cover:
 * B1: Page borders
 * B2: Form field data (w:ffData)
 * B4: webSettings.xml generation
 * C1: w14 run effects passthrough
 * C2: Expanded settings coverage
 */

import { Document } from '../../src/core/Document';
import { ComplexField } from '../../src/elements/Field';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Section } from '../../src/elements/Section';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: Creates a DOCX buffer, then injects custom document.xml content.
 */
async function createDocxWithDocumentXml(documentXml: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText('placeholder'));
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.updateFile(DOCX_PATHS.DOCUMENT, documentXml);
  return await zipHandler.toBuffer();
}

/**
 * Helper: Creates a DOCX buffer with custom settings.xml.
 */
async function createDocxWithSettings(settingsXml: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText('Test content'));
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.updateFile(DOCX_PATHS.SETTINGS, settingsXml);
  return await zipHandler.toBuffer();
}

function wrapInDocument(bodyContent: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="w14">
  <w:body>
    ${bodyContent}
  </w:body>
</w:document>`;
}

describe('ECMA-376 Gap Analysis Phase B', () => {
  describe('B1: Page borders', () => {
    it('should set and generate page borders on Section', () => {
      const section = new Section({
        pageBorders: {
          top: { style: 'single', size: 4, color: '000000', space: 1 },
          bottom: { style: 'single', size: 4, color: '000000', space: 1 },
          left: { style: 'double', size: 6, color: 'FF0000', space: 4 },
          right: { style: 'double', size: 6, color: 'FF0000', space: 4 },
        },
      });
      const xml = XMLBuilder.elementToString(section.toXML());
      expect(xml).toContain('<w:pgBorders');
      expect(xml).toContain('<w:top');
      expect(xml).toContain('<w:bottom');
      expect(xml).toContain('<w:left');
      expect(xml).toContain('<w:right');
    });

    it('should round-trip page borders through parse', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>test</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:pgBorders w:offsetFrom="page">
        <w:top w:val="single" w:sz="4" w:color="FF0000" w:space="24"/>
        <w:bottom w:val="single" w:sz="4" w:color="FF0000" w:space="24"/>
      </w:pgBorders>
    </w:sectPr>
  </w:body>
</w:document>`;
      const buffer = await createDocxWithDocumentXml(documentXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const section = doc.getSection();
      expect(section).toBeDefined();
      const borders = section.getProperties().pageBorders;
      expect(borders).toBeDefined();
      expect(borders!.top?.style).toBe('single');
      expect(borders!.top?.color).toBe('FF0000');
      expect(borders!.offsetFrom).toBe('page');
      doc.dispose();
    });
  });

  describe('B2: Form field data (w:ffData)', () => {
    it('should round-trip text input form field via ComplexField', async () => {
      const bodyXml = `<w:p>
        <w:r>
          <w:fldChar w:fldCharType="begin">
            <w:ffData>
              <w:name w:val="TextInput1"/>
              <w:enabled/>
              <w:calcOnExit w:val="0"/>
              <w:helpText w:type="text" w:val="Enter name"/>
              <w:textInput>
                <w:default w:val="John"/>
                <w:maxLength w:val="50"/>
                <w:format w:val="UPPERCASE"/>
              </w:textInput>
            </w:ffData>
          </w:fldChar>
        </w:r>
        <w:r><w:instrText> FORMTEXT </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>John</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const para = doc.getParagraphs()[0]!;
      const content = para.getContent();
      const field = content.find(c => c instanceof ComplexField) as ComplexField | undefined;
      expect(field).toBeDefined();
      const ffd = field!.getFormFieldData();
      expect(ffd).toBeDefined();
      expect(ffd!.name).toBe('TextInput1');
      expect(ffd!.enabled).toBe(true);
      expect(ffd!.helpText).toBe('Enter name');
      const ti = ffd!.fieldType;
      expect(ti?.type).toBe('textInput');
      if (ti?.type === 'textInput') {
        expect(ti.defaultValue).toBe('John');
        expect(ti.maxLength).toBe(50);
        expect(ti.format).toBe('UPPERCASE');
      }
      doc.dispose();
    });

    it('should round-trip checkbox form field via ComplexField', async () => {
      const bodyXml = `<w:p>
        <w:r>
          <w:fldChar w:fldCharType="begin">
            <w:ffData>
              <w:name w:val="Check1"/>
              <w:enabled/>
              <w:calcOnExit w:val="0"/>
              <w:checkBox>
                <w:sizeAuto/>
                <w:default w:val="0"/>
                <w:checked w:val="1"/>
              </w:checkBox>
            </w:ffData>
          </w:fldChar>
        </w:r>
        <w:r><w:instrText> FORMCHECKBOX </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const para = doc.getParagraphs()[0]!;
      const content = para.getContent();
      const field = content.find(c => c instanceof ComplexField) as ComplexField | undefined;
      expect(field).toBeDefined();
      const ffd = field!.getFormFieldData();
      expect(ffd).toBeDefined();
      const cb = ffd!.fieldType;
      expect(cb?.type).toBe('checkBox');
      if (cb?.type === 'checkBox') {
        expect(cb.checked).toBe(true);
        expect(cb.defaultChecked).toBe(false);
        expect(cb.size).toBe('auto');
      }
      doc.dispose();
    });

    it('should round-trip dropdown form field via ComplexField', async () => {
      const bodyXml = `<w:p>
        <w:r>
          <w:fldChar w:fldCharType="begin">
            <w:ffData>
              <w:name w:val="Dropdown1"/>
              <w:enabled/>
              <w:ddList>
                <w:result w:val="1"/>
                <w:default w:val="0"/>
                <w:listEntry w:val="Option A"/>
                <w:listEntry w:val="Option B"/>
                <w:listEntry w:val="Option C"/>
              </w:ddList>
            </w:ffData>
          </w:fldChar>
        </w:r>
        <w:r><w:instrText> FORMDROPDOWN </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const para = doc.getParagraphs()[0]!;
      const content = para.getContent();
      const field = content.find(c => c instanceof ComplexField) as ComplexField | undefined;
      expect(field).toBeDefined();
      const ffd = field!.getFormFieldData();
      expect(ffd).toBeDefined();
      const dd = ffd!.fieldType;
      expect(dd?.type).toBe('dropDownList');
      if (dd?.type === 'dropDownList') {
        expect(dd.result).toBe(1);
        expect(dd.defaultResult).toBe(0);
        expect(dd.listEntries).toEqual(['Option A', 'Option B', 'Option C']);
      }
      doc.dispose();
    });

    it('should generate ffData in toXML for text input', () => {
      const run = Run.createFromContent([
        {
          type: 'fieldChar',
          fieldCharType: 'begin',
          formFieldData: {
            name: 'TestField',
            enabled: true,
            calcOnExit: false,
            helpText: 'Help',
            fieldType: {
              type: 'textInput',
              defaultValue: 'Default',
              maxLength: 100,
              format: 'UPPERCASE',
            },
          },
        },
      ]);
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('<w:ffData>');
      expect(xml).toContain('<w:name');
      expect(xml).toContain('w:val="TestField"');
      expect(xml).toContain('<w:textInput>');
      expect(xml).toContain('w:val="Default"');
      expect(xml).toContain('<w:maxLength');
    });
  });

  describe('B4: webSettings.xml generation', () => {
    it('should include webSettings.xml in new documents', async () => {
      const doc = Document.create();
      doc.addParagraph(new Paragraph().addText('Test'));
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zipHandler = new ZipHandler();
      await zipHandler.loadFromBuffer(buffer);
      const webSettings = zipHandler.getFileAsString('word/webSettings.xml');
      expect(webSettings).toBeDefined();
      expect(webSettings).toContain('<w:webSettings');
      expect(webSettings).toContain('<w:optimizeForBrowser');
    });

    it('should include webSettings.xml relationship', async () => {
      const doc = Document.create();
      doc.addParagraph(new Paragraph().addText('Test'));
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zipHandler = new ZipHandler();
      await zipHandler.loadFromBuffer(buffer);
      const rels = zipHandler.getFileAsString('word/_rels/document.xml.rels');
      expect(rels).toContain('webSettings');
    });

    it('should include webSettings.xml content type', async () => {
      const doc = Document.create();
      doc.addParagraph(new Paragraph().addText('Test'));
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zipHandler = new ZipHandler();
      await zipHandler.loadFromBuffer(buffer);
      const contentTypes = zipHandler.getFileAsString('[Content_Types].xml');
      expect(contentTypes).toContain('webSettings.xml');
    });
  });
});

describe('ECMA-376 Gap Analysis Phase C', () => {
  describe('C1: w14 run effects passthrough', () => {
    it('should round-trip w14:textOutline through parse and save', async () => {
      const bodyXml = `<w:p><w:r><w:rPr>
        <w14:textOutline w14:w="9525" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr">
          <w14:solidFill><w14:srgbClr w14:val="000000"/></w14:solidFill>
          <w14:prstDash w14:val="solid"/>
          <w14:round/>
        </w14:textOutline>
      </w:rPr><w:t>Outlined</w:t></w:r></w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Verify the w14 property was parsed
      const runs = doc.getParagraphs()[0]?.getRuns() ?? [];
      expect(runs.length).toBeGreaterThan(0);
      const fmt = runs[0]!.getFormatting();
      expect(fmt.rawW14Properties).toBeDefined();
      expect(fmt.rawW14Properties!.length).toBeGreaterThan(0);
      expect(fmt.rawW14Properties![0]).toContain('w14:textOutline');

      doc.dispose();
    });

    it('should generate w14 properties in toXML', () => {
      const run = new Run('Test');
      run.addRawW14Property('<w14:ligatures w14:val="standard"/>');
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('w14:ligatures');
      expect(xml).toContain('w14:val="standard"');
    });

    it('should preserve multiple w14 effects', async () => {
      const bodyXml = `<w:p><w:r><w:rPr>
        <w14:ligatures w14:val="standard"/>
        <w14:numForm w14:val="oldStyle"/>
        <w14:cntxtAlts/>
      </w:rPr><w:t>Fancy</w:t></w:r></w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const runs = doc.getParagraphs()[0]?.getRuns() ?? [];
      const fmt = runs[0]!.getFormatting();
      expect(fmt.rawW14Properties).toBeDefined();
      expect(fmt.rawW14Properties!.length).toBe(3);
      doc.dispose();
    });
  });

  describe('C2: Expanded settings coverage', () => {
    it('should parse evenAndOddHeaders from settings.xml', async () => {
      const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:evenAndOddHeaders/>
  <w:defaultTabStop w:val="720"/>
</w:settings>`;
      const buffer = await createDocxWithSettings(settingsXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(doc.getEvenAndOddHeaders()).toBe(true);
      doc.dispose();
    });

    it('should parse mirrorMargins from settings.xml', async () => {
      const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:mirrorMargins/>
  <w:defaultTabStop w:val="720"/>
</w:settings>`;
      const buffer = await createDocxWithSettings(settingsXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(doc.getMirrorMargins()).toBe(true);
      doc.dispose();
    });

    it('should parse autoHyphenation from settings.xml', async () => {
      const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:autoHyphenation/>
  <w:defaultTabStop w:val="720"/>
</w:settings>`;
      const buffer = await createDocxWithSettings(settingsXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(doc.getAutoHyphenation()).toBe(true);
      doc.dispose();
    });

    it('should parse decimalSymbol and listSeparator', async () => {
      const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:decimalSymbol w:val=","/>
  <w:listSeparator w:val=";"/>
</w:settings>`;
      const buffer = await createDocxWithSettings(settingsXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(doc.getDecimalSymbol()).toBe(',');
      expect(doc.getListSeparator()).toBe(';');
      doc.dispose();
    });

    it('should set and merge evenAndOddHeaders into settings', async () => {
      const settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;
      const buffer = await createDocxWithSettings(settingsXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(doc.getEvenAndOddHeaders()).toBe(false);
      doc.setEvenAndOddHeaders(true);
      const savedBuffer = await doc.toBuffer();
      doc.dispose();

      const zipHandler = new ZipHandler();
      await zipHandler.loadFromBuffer(savedBuffer);
      const savedSettings = zipHandler.getFileAsString(DOCX_PATHS.SETTINGS);
      expect(savedSettings).toContain('<w:evenAndOddHeaders');
    });

    it('should default to false when settings are absent', () => {
      const doc = Document.create();
      expect(doc.getEvenAndOddHeaders()).toBe(false);
      expect(doc.getMirrorMargins()).toBe(false);
      expect(doc.getAutoHyphenation()).toBe(false);
      expect(doc.getDecimalSymbol()).toBeUndefined();
      expect(doc.getListSeparator()).toBeUndefined();
      doc.dispose();
    });
  });
});
