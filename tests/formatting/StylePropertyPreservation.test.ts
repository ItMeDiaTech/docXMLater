/**
 * Style Property Preservation Tests
 *
 * Tests that structural properties (basedOn, next, link, uiPriority) are preserved
 * when replacing an existing style via addStyle() with a style that doesn't
 * explicitly set those properties.
 */

import { Style } from '../../src/formatting/Style';
import { StylesManager } from '../../src/formatting/StylesManager';
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

describe('Style Property Preservation', () => {
  describe('addStyle() preserves structural properties from existing style', () => {
    it('should preserve basedOn when new style does not set it', () => {
      const manager = StylesManager.create();

      // Add a style with basedOn
      const original = Style.create({
        styleId: 'Heading1',
        name: 'heading 1',
        type: 'paragraph',
        basedOn: 'Normal',
        next: 'Heading4',
        link: 'Heading1Char',
        uiPriority: 9,
        runFormatting: { bold: true, size: 16 },
      });
      manager.addStyle(original);
      manager.resetModified(); // Simulate post-parse state

      // Replace with a style that only sets formatting (no basedOn)
      const replacement = Style.create({
        styleId: 'Heading1',
        name: 'heading 1',
        type: 'paragraph',
        runFormatting: { bold: true, size: 14, color: '000000' },
      });
      manager.addStyle(replacement);

      // Verify structural properties preserved
      const result = manager.getStyle('Heading1')!;
      const props = result.getProperties();
      expect(props.basedOn).toBe('Normal');
      expect(props.next).toBe('Heading4');
      expect(props.link).toBe('Heading1Char');
      expect(props.uiPriority).toBe(9);
      // Verify new formatting applied
      expect(props.runFormatting?.color).toBe('000000');
      expect(props.runFormatting?.size).toBe(14);
    });

    it('should preserve next when new style does not set it', () => {
      const manager = StylesManager.create();

      const original = Style.create({
        styleId: 'TestStyle',
        name: 'Test',
        type: 'paragraph',
        next: 'Normal',
      });
      manager.addStyle(original);
      manager.resetModified();

      const replacement = Style.create({
        styleId: 'TestStyle',
        name: 'Test',
        type: 'paragraph',
        runFormatting: { italic: true },
      });
      manager.addStyle(replacement);

      expect(manager.getStyle('TestStyle')!.getProperties().next).toBe('Normal');
    });

    it('should preserve link when new style does not set it', () => {
      const manager = StylesManager.create();

      const original = Style.create({
        styleId: 'Heading2',
        name: 'heading 2',
        type: 'paragraph',
        link: 'Heading2Char',
      });
      manager.addStyle(original);
      manager.resetModified();

      const replacement = Style.create({
        styleId: 'Heading2',
        name: 'heading 2',
        type: 'paragraph',
        runFormatting: { color: 'FF0000' },
      });
      manager.addStyle(replacement);

      expect(manager.getStyle('Heading2')!.getProperties().link).toBe('Heading2Char');
    });

    it('should preserve uiPriority when new style does not set it', () => {
      const manager = StylesManager.create();

      const original = Style.create({
        styleId: 'ListParagraph',
        name: 'List Paragraph',
        type: 'paragraph',
        uiPriority: 34,
        basedOn: 'Normal',
        link: 'ListParagraphChar',
      });
      manager.addStyle(original);
      manager.resetModified();

      const replacement = Style.create({
        styleId: 'ListParagraph',
        name: 'List Paragraph',
        type: 'paragraph',
        paragraphFormatting: { alignment: 'left' },
      });
      manager.addStyle(replacement);

      const props = manager.getStyle('ListParagraph')!.getProperties();
      expect(props.uiPriority).toBe(34);
      expect(props.basedOn).toBe('Normal');
      expect(props.link).toBe('ListParagraphChar');
    });

    it('should preserve isDefault flag from existing style', () => {
      const manager = StylesManager.create();

      const original = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        isDefault: true,
      });
      manager.addStyle(original);
      manager.resetModified();

      // Replace without isDefault
      const replacement = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        runFormatting: { font: 'Arial' },
      });
      manager.addStyle(replacement);

      expect(manager.getStyle('Normal')!.getIsDefault()).toBe(true);
    });

    it('should NOT override new style properties that are explicitly set', () => {
      const manager = StylesManager.create();

      const original = Style.create({
        styleId: 'Heading1',
        name: 'heading 1',
        type: 'paragraph',
        basedOn: 'Normal',
        next: 'Normal',
        link: 'Heading1Char',
        uiPriority: 9,
      });
      manager.addStyle(original);
      manager.resetModified();

      // New style explicitly sets different values
      const replacement = Style.create({
        styleId: 'Heading1',
        name: 'heading 1',
        type: 'paragraph',
        basedOn: 'Title',     // Explicitly different
        next: 'Heading2',     // Explicitly different
        link: 'NewCharStyle', // Explicitly different
        uiPriority: 1,        // Explicitly different
      });
      manager.addStyle(replacement);

      const props = manager.getStyle('Heading1')!.getProperties();
      expect(props.basedOn).toBe('Title');
      expect(props.next).toBe('Heading2');
      expect(props.link).toBe('NewCharStyle');
      expect(props.uiPriority).toBe(1);
    });

    it('should work correctly when adding a brand new style (no existing)', () => {
      const manager = StylesManager.create();

      const newStyle = Style.create({
        styleId: 'BrandNew',
        name: 'Brand New',
        type: 'paragraph',
        runFormatting: { bold: true },
      });
      manager.addStyle(newStyle);

      const props = manager.getStyle('BrandNew')!.getProperties();
      expect(props.basedOn).toBeUndefined();
      expect(props.next).toBeUndefined();
      expect(props.link).toBeUndefined();
      expect(props.uiPriority).toBeUndefined();
    });
  });

  describe('Style preservation in XML merge', () => {
    it('should preserve basedOn/next/link in merged XML output', async () => {
      // Create a DOCX with styles that have structural properties
      const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Heading4"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:rsid w:val="00AB1234"/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="240"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Heading1Char">
    <w:name w:val="Heading 1 Char"/>
    <w:link w:val="Heading1"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
</w:styles>`;

      // Create DOCX and inject custom styles.xml
      const doc = Document.create();
      doc.addParagraph(new Paragraph().addText('Test'));
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zipHandler = new ZipHandler();
      await zipHandler.loadFromBuffer(buffer);
      zipHandler.updateFile(DOCX_PATHS.STYLES, stylesXml);
      const customBuffer = await zipHandler.toBuffer();

      // Load, modify Heading1, save
      const doc2 = await Document.loadFromBuffer(customBuffer);

      // Add a replacement Heading1 that only changes formatting
      const newHeading1 = Style.create({
        styleId: 'Heading1',
        name: 'heading 1',
        type: 'paragraph',
        runFormatting: { bold: true, size: 14, color: '000000' },
      });
      doc2.addStyle(newHeading1);

      const outputBuffer = await doc2.toBuffer();
      doc2.dispose();

      // Extract and inspect the merged styles.xml
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(outputBuffer);
      const mergedStylesXml = zip2.getFileAsString(DOCX_PATHS.STYLES) || '';

      // The merged Heading1 should have preserved structural properties
      // AND the formatting should be updated
      expect(mergedStylesXml).toContain('w:styleId="Heading1"');
      // basedOn, next, link should appear in the merged style's toXML() output
      // (they were preserved from existing style via addStyle())
      expect(mergedStylesXml).toContain('w:val="Normal"'); // basedOn
      expect(mergedStylesXml).toContain('w:val="Heading1Char"'); // link preserved

      // Heading1Char should still be untouched (not in modifiedStyleIds)
      expect(mergedStylesXml).toContain('w:styleId="Heading1Char"');
    });
  });
});
