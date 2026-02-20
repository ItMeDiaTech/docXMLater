/**
 * Tests for Normal and NormalWeb Style Linking
 *
 * Validates that changes to Normal style are automatically
 * applied to NormalWeb (Normal (Web)) style when it exists.
 */

import { Document, Style, Paragraph, Run } from '../../src';

describe('Normal and NormalWeb Style Linking', () => {
  let doc: Document;

  beforeEach(() => {
    doc = Document.create();
  });

  afterEach(() => {
    doc.dispose();
  });

  describe('linkNormalWebToNormal option (default: true)', () => {
    it('should apply Normal changes to NormalWeb by default', () => {
      // Add NormalWeb style with different formatting
      const normalWeb = Style.create({
        styleId: 'NormalWeb',
        name: 'Normal (Web)',
        type: 'paragraph',
        basedOn: 'Normal',
        runFormatting: { font: 'Times New Roman', size: 10 },
        paragraphFormatting: { alignment: 'left' },
      });
      doc.addStyle(normalWeb);

      // Apply styles with default options (linkNormalWebToNormal defaults to true)
      doc.applyStyles({
        normal: {
          run: { font: 'Arial', size: 12 },
          paragraph: { alignment: 'left' },
        },
      });

      // Verify NormalWeb was updated with Normal's formatting
      const updated = doc.getStyle('NormalWeb');
      expect(updated).toBeDefined();
      expect(updated?.getRunFormatting()?.font).toBe('Arial');
      expect(updated?.getRunFormatting()?.size).toBe(12);
    });

    it('should apply paragraph formatting to NormalWeb', () => {
      const normalWeb = Style.create({
        styleId: 'NormalWeb',
        name: 'Normal (Web)',
        type: 'paragraph',
        paragraphFormatting: { alignment: 'left', spacing: { before: 0 } },
      });
      doc.addStyle(normalWeb);

      doc.applyStyles({
        normal: {
          run: { font: 'Verdana' },
          paragraph: { alignment: 'justify', spacing: { before: 100, after: 100 } },
        },
      });

      const updated = doc.getStyle('NormalWeb');
      expect(updated?.getParagraphFormatting()?.alignment).toBe('justify');
      expect(updated?.getParagraphFormatting()?.spacing?.before).toBe(100);
    });
  });

  describe('linkNormalWebToNormal: false', () => {
    it('should NOT apply Normal changes to NormalWeb when disabled', () => {
      // Add NormalWeb style with specific formatting
      const normalWeb = Style.create({
        styleId: 'NormalWeb',
        name: 'Normal (Web)',
        type: 'paragraph',
        runFormatting: { font: 'Times New Roman', size: 10 },
        paragraphFormatting: { alignment: 'left' },
      });
      doc.addStyle(normalWeb);

      // Apply styles with linking DISABLED
      doc.applyStyles({
        normal: {
          run: { font: 'Arial', size: 12 },
          paragraph: { alignment: 'center' },
        },
        linkNormalWebToNormal: false,
      });

      // Verify NormalWeb was NOT updated
      const notUpdated = doc.getStyle('NormalWeb');
      expect(notUpdated?.getRunFormatting()?.font).toBe('Times New Roman');
      expect(notUpdated?.getRunFormatting()?.size).toBe(10);
      expect(notUpdated?.getParagraphFormatting()?.alignment).toBe('left');
    });
  });

  describe('NormalWeb does not exist', () => {
    it('should not throw when NormalWeb does not exist', () => {
      // Apply styles without adding NormalWeb - should not throw
      expect(() => {
        doc.applyStyles({
          normal: { run: { font: 'Arial', size: 12 } },
        });
      }).not.toThrow();
    });

    it('should still apply Normal changes when NormalWeb does not exist', () => {
      const results = doc.applyStyles({
        normal: {
          run: { font: 'Arial', size: 12 },
          paragraph: { alignment: 'center' },
        },
      });

      expect(results.normal).toBe(true);

      // Verify Normal was updated
      const normal = doc.getStyle('Normal');
      expect(normal?.getRunFormatting()?.font).toBe('Arial');
    });
  });

  describe('Modified style tracking', () => {
    it('should mark NormalWeb as modified for merging', () => {
      const normalWeb = Style.create({
        styleId: 'NormalWeb',
        name: 'Normal (Web)',
        type: 'paragraph',
      });
      doc.addStyle(normalWeb);

      // Reset modified tracking
      doc.styles().resetModified();

      // Apply styles
      doc.applyStyles({
        normal: { run: { font: 'Arial' } },
      });

      // Verify NormalWeb is in modified set
      const modified = doc.styles().getModifiedStyleIds();
      expect(modified.has('NormalWeb')).toBe(true);
    });

    it('should NOT mark NormalWeb as modified when linking disabled', () => {
      const normalWeb = Style.create({
        styleId: 'NormalWeb',
        name: 'Normal (Web)',
        type: 'paragraph',
      });
      doc.addStyle(normalWeb);

      // Reset modified tracking
      doc.styles().resetModified();

      // Apply styles with linking disabled
      doc.applyStyles({
        normal: { run: { font: 'Arial' } },
        linkNormalWebToNormal: false,
      });

      // Verify NormalWeb is NOT in modified set
      const modified = doc.styles().getModifiedStyleIds();
      expect(modified.has('NormalWeb')).toBe(false);
    });
  });

  describe('Edge cases', () => {
    it('should handle NormalWeb with basedOn Normal', () => {
      // Real-world scenario: NormalWeb often has basedOn="Normal"
      const normalWeb = Style.create({
        styleId: 'NormalWeb',
        name: 'Normal (Web)',
        type: 'paragraph',
        basedOn: 'Normal',
        // Only override specific properties
        runFormatting: { size: 10 },
      });
      doc.addStyle(normalWeb);

      doc.applyStyles({
        normal: {
          run: { font: 'Georgia', size: 11 },
          paragraph: { alignment: 'justify' },
        },
      });

      const updated = doc.getStyle('NormalWeb');
      // Should have the new font and size from Normal linking
      expect(updated?.getRunFormatting()?.font).toBe('Georgia');
      expect(updated?.getRunFormatting()?.size).toBe(11);
    });

    it('should work with explicit linkNormalWebToNormal: true', () => {
      const normalWeb = Style.create({
        styleId: 'NormalWeb',
        name: 'Normal (Web)',
        type: 'paragraph',
        runFormatting: { font: 'Courier' },
      });
      doc.addStyle(normalWeb);

      // Explicitly set to true (same as default)
      doc.applyStyles({
        normal: { run: { font: 'Tahoma' } },
        linkNormalWebToNormal: true,
      });

      const updated = doc.getStyle('NormalWeb');
      expect(updated?.getRunFormatting()?.font).toBe('Tahoma');
    });
  });

  describe('NormalWeb paragraph-level formatting', () => {
    /**
     * These tests verify that applyStyles() clears and replaces direct
     * run formatting on NormalWeb paragraphs — not just style definitions.
     */

    function createDocWithNormalWebParagraphs(): Document {
      const doc = Document.create();
      // Add NormalWeb style
      const normalWeb = Style.create({
        styleId: 'NormalWeb',
        name: 'Normal (Web)',
        type: 'paragraph',
        basedOn: 'Normal',
        runFormatting: { font: 'Times New Roman', size: 12, color: '000000' },
      });
      doc.addStyle(normalWeb);
      return doc;
    }

    it('should clear and replace direct formatting on NormalWeb paragraphs', () => {
      const doc = createDocWithNormalWebParagraphs();

      // Add a NormalWeb paragraph with direct bold+red formatting
      const para = new Paragraph();
      para.setStyle('NormalWeb');
      const run = new Run('NormalWeb text');
      run.setBold(true);
      run.setColor('FF0000');
      run.setFont('Times New Roman');
      run.setSize(12);
      para.addRun(run);
      doc.addParagraph(para);

      // Apply styles — should clear direct formatting and apply Normal config
      // preserveBold defaults to true for Normal, so we must explicitly disable it
      doc.applyStyles({
        normal: {
          run: { font: 'Verdana', size: 10, color: '000000', bold: false, preserveBold: false },
          paragraph: { alignment: 'left' },
        },
      });

      // Run should now have Normal config formatting (bold cleared, color changed)
      const runs = doc.getParagraphs().filter(p => p.getStyle() === 'NormalWeb');
      expect(runs.length).toBeGreaterThan(0);
      const updatedRun = runs[0]!.getRuns()[0]!;
      expect(updatedRun.getFont()).toBe('Verdana');
      expect(updatedRun.getColor()).toBe('000000');
      expect(updatedRun.getFormatting().bold).toBeFalsy();

      doc.dispose();
    });

    it('should skip NormalWeb paragraphs when linkNormalWebToNormal is false', () => {
      const doc = createDocWithNormalWebParagraphs();

      const para = new Paragraph();
      para.setStyle('NormalWeb');
      const run = new Run('Keep my formatting');
      run.setBold(true);
      run.setColor('FF0000');
      run.setFont('Times New Roman');
      run.setSize(12);
      para.addRun(run);
      doc.addParagraph(para);

      // Apply styles with linking DISABLED
      doc.applyStyles({
        normal: {
          run: { font: 'Verdana', size: 10, color: '000000', bold: false },
          paragraph: { alignment: 'left' },
        },
        linkNormalWebToNormal: false,
      });

      // Run should retain original direct formatting
      const nwParas = doc.getParagraphs().filter(p => p.getStyle() === 'NormalWeb');
      expect(nwParas.length).toBeGreaterThan(0);
      const keptRun = nwParas[0]!.getRuns()[0]!;
      expect(keptRun.getFont()).toBe('Times New Roman');
      expect(keptRun.getColor()).toBe('FF0000');
      expect(keptRun.getFormatting().bold).toBe(true);

      doc.dispose();
    });

    it('should preserve white font on NormalWeb paragraphs when preserveWhiteFont is true', () => {
      const doc = createDocWithNormalWebParagraphs();

      const para = new Paragraph();
      para.setStyle('NormalWeb');
      const run = new Run('Hidden text');
      run.setColor('FFFFFF'); // White font (hidden)
      run.setFont('Times New Roman');
      run.setSize(12);
      para.addRun(run);
      doc.addParagraph(para);

      doc.applyStyles({
        normal: {
          run: { font: 'Verdana', size: 10, color: '000000', bold: false },
          paragraph: { alignment: 'left' },
        },
        preserveWhiteFont: true,
      });

      // White font should be preserved
      const nwParas = doc.getParagraphs().filter(p => p.getStyle() === 'NormalWeb');
      const whiteRun = nwParas[0]!.getRuns()[0]!;
      expect(whiteRun.getColor()).toBe('FFFFFF');
      // Other formatting should still be updated
      expect(whiteRun.getFont()).toBe('Verdana');

      doc.dispose();
    });

    it('should respect preserveBold flag on NormalWeb paragraphs', () => {
      const doc = createDocWithNormalWebParagraphs();

      const para = new Paragraph();
      para.setStyle('NormalWeb');
      const run = new Run('Bold text');
      run.setBold(true);
      run.setFont('Times New Roman');
      run.setSize(12);
      para.addRun(run);
      doc.addParagraph(para);

      // Default: preserveBold is true for Normal config
      doc.applyStyles({
        normal: {
          run: { font: 'Verdana', size: 10, color: '000000', bold: false, preserveBold: true },
          paragraph: { alignment: 'left' },
        },
      });

      // Bold should be preserved because preserveBold is true
      const nwParas = doc.getParagraphs().filter(p => p.getStyle() === 'NormalWeb');
      const boldRun = nwParas[0]!.getRuns()[0]!;
      expect(boldRun.getFormatting().bold).toBe(true);
      // Font and size should still be updated
      expect(boldRun.getFont()).toBe('Verdana');
      expect(boldRun.getSize()).toBe(10);

      doc.dispose();
    });

    it('should process mixed Normal and NormalWeb paragraphs identically', () => {
      const doc = createDocWithNormalWebParagraphs();

      // Add a Normal paragraph
      const normalPara = new Paragraph();
      normalPara.setStyle('Normal');
      const normalRun = new Run('Normal text');
      normalRun.setBold(true);
      normalRun.setColor('FF0000');
      normalRun.setFont('Times New Roman');
      normalRun.setSize(12);
      normalPara.addRun(normalRun);
      doc.addParagraph(normalPara);

      // Add a NormalWeb paragraph with same formatting
      const nwPara = new Paragraph();
      nwPara.setStyle('NormalWeb');
      const nwRun = new Run('NormalWeb text');
      nwRun.setBold(true);
      nwRun.setColor('FF0000');
      nwRun.setFont('Times New Roman');
      nwRun.setSize(12);
      nwPara.addRun(nwRun);
      doc.addParagraph(nwPara);

      doc.applyStyles({
        normal: {
          run: { font: 'Verdana', size: 10, color: '000000', bold: false, preserveBold: false },
          paragraph: { alignment: 'left' },
        },
      });

      // Both Normal and NormalWeb should get the same treatment
      const allParas = doc.getParagraphs();
      const normalParas = allParas.filter(p => p.getStyle() === 'Normal');
      const nwParas = allParas.filter(p => p.getStyle() === 'NormalWeb');

      // At least one of each
      expect(normalParas.length).toBeGreaterThan(0);
      expect(nwParas.length).toBeGreaterThan(0);

      // Compare formatting outcomes
      const nResult = normalParas[normalParas.length - 1]!.getRuns()[0]!;
      const nwResult = nwParas[0]!.getRuns()[0]!;

      expect(nResult.getFont()).toBe(nwResult.getFont());
      expect(nResult.getColor()).toBe(nwResult.getColor());
      expect(nResult.getSize()).toBe(nwResult.getSize());
      expect(nResult.getFormatting().bold).toBe(nwResult.getFormatting().bold);

      doc.dispose();
    });

    it('should update paragraphMarkRunProperties on NormalWeb paragraphs', () => {
      const doc = createDocWithNormalWebParagraphs();

      const para = new Paragraph();
      para.setStyle('NormalWeb');
      para.formatting.paragraphMarkRunProperties = {
        bold: true,
        font: 'Times New Roman',
        size: 12,
        color: 'FF0000',
      };
      const run = new Run('Text');
      run.setFont('Times New Roman');
      run.setSize(12);
      para.addRun(run);
      doc.addParagraph(para);

      doc.applyStyles({
        normal: {
          run: { font: 'Verdana', size: 10, color: '000000', bold: false, preserveBold: false },
          paragraph: { alignment: 'left' },
        },
      });

      // Paragraph mark properties should be updated
      const nwParas = doc.getParagraphs().filter(p => p.getStyle() === 'NormalWeb');
      const markProps = nwParas[0]!.formatting.paragraphMarkRunProperties;
      expect(markProps).toBeDefined();
      expect(markProps!.font).toBe('Verdana');
      expect(markProps!.color).toBe('000000');
      expect(markProps!.bold).toBeFalsy();

      doc.dispose();
    });

    it('should preserve center alignment on NormalWeb paragraphs when preserveCenterAlignment is true', () => {
      const doc = createDocWithNormalWebParagraphs();

      const para = new Paragraph();
      para.setStyle('NormalWeb');
      para.setAlignment('center');
      const run = new Run('Centered text');
      run.setFont('Times New Roman');
      run.setSize(12);
      para.addRun(run);
      doc.addParagraph(para);

      doc.applyStyles({
        normal: {
          run: { font: 'Verdana', size: 10, color: '000000' },
          paragraph: { alignment: 'left' },
          preserveCenterAlignment: true,
        },
      });

      // Center alignment should be preserved
      const nwParas = doc.getParagraphs().filter(p => p.getStyle() === 'NormalWeb');
      expect(nwParas[0]!.getAlignment()).toBe('center');

      doc.dispose();
    });
  });
});
