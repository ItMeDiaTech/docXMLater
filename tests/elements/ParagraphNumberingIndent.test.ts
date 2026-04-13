/**
 * Tests for Paragraph numbering and indentation conflict resolution
 * Validates that numbering properly clears conflicting indentation properties
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { defaultLogger } from '../../src/utils/logger';

describe('Paragraph - Numbering and Indentation Conflict Resolution', () => {
  // Mock logger to capture warnings
  let warnSpy: jest.SpyInstance;

  beforeEach(() => {
    warnSpy = jest.spyOn(defaultLogger, 'warn').mockImplementation(() => {});
  });

  afterEach(() => {
    warnSpy.mockRestore();
  });

  describe('setNumbering() clears conflicting indentation', () => {
    it('should clear left indentation when numbering is applied', () => {
      const para = new Paragraph();
      para.setLeftIndent(1440); // Set 1 inch indent
      para.setNumbering(1, 0);

      const formatting = para.getFormatting();
      expect(formatting.numbering).toEqual({ numId: 1, level: 0 });
      expect(formatting.indentation).toBeUndefined();
    });

    it('should clear firstLine indentation when numbering is applied', () => {
      const para = new Paragraph();
      para.setFirstLineIndent(720); // Set 0.5 inch first line indent
      para.setNumbering(1, 0);

      const formatting = para.getFormatting();
      expect(formatting.numbering).toEqual({ numId: 1, level: 0 });
      expect(formatting.indentation).toBeUndefined();
    });

    it('should clear hanging indentation when numbering is applied', () => {
      const para = new Paragraph();
      // Manually set hanging indent (no public method for this)
      para.setLeftIndent(720);
      const formatting = para.getFormatting();
      if (formatting.indentation) {
        formatting.indentation.hanging = 360;
      }

      para.setNumbering(1, 0);

      // Get formatting after setNumbering
      const updatedFormatting = para.getFormatting();
      expect(updatedFormatting.numbering).toEqual({ numId: 1, level: 0 });
      expect(updatedFormatting.indentation).toBeUndefined();
    });

    it('should clear multiple indentation properties when numbering is applied', () => {
      const para = new Paragraph();
      para.setLeftIndent(1440);
      para.setFirstLineIndent(720);
      para.setNumbering(1, 0);

      const formatting = para.getFormatting();
      expect(formatting.numbering).toEqual({ numId: 1, level: 0 });
      expect(formatting.indentation).toBeUndefined();
    });

    it('should preserve right indentation when numbering is applied', () => {
      const para = new Paragraph();
      para.setLeftIndent(720);
      para.setRightIndent(1440);
      para.setNumbering(1, 0);

      const formatting = para.getFormatting();
      expect(formatting.numbering).toEqual({ numId: 1, level: 0 });
      expect(formatting.indentation).toEqual({ right: 1440 });
    });

    it('should handle numbering when no indentation exists', () => {
      const para = new Paragraph();
      para.setNumbering(1, 0);

      const formatting = para.getFormatting();
      expect(formatting.numbering).toEqual({ numId: 1, level: 0 });
      expect(formatting.indentation).toBeUndefined();
    });
  });

  describe('Warnings when setting indentation on numbered paragraphs', () => {
    it('should warn when setting left indent on numbered paragraph', () => {
      const para = new Paragraph();
      para.setNumbering(1, 0);
      para.setLeftIndent(1440);

      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining('Setting left indentation on a numbered paragraph')
      );
    });

    it('should warn when setting first line indent on numbered paragraph', () => {
      const para = new Paragraph();
      para.setNumbering(1, 0);
      para.setFirstLineIndent(720);

      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining('Setting first line indentation on a numbered paragraph')
      );
    });

    it('should not warn when setting right indent on numbered paragraph', () => {
      const para = new Paragraph();
      para.setNumbering(1, 0);
      para.setRightIndent(1440);

      expect(warnSpy).not.toHaveBeenCalled();
    });

    it('should not warn when setting indent on non-numbered paragraph', () => {
      const para = new Paragraph();
      para.setLeftIndent(720);
      para.setFirstLineIndent(360);

      expect(warnSpy).not.toHaveBeenCalled();
    });
  });

  describe('Edge cases and complex scenarios', () => {
    it('should handle changing numbering levels', () => {
      const para = new Paragraph();
      para.setLeftIndent(720);
      para.setNumbering(1, 0); // Level 0

      const formatting1 = para.getFormatting();
      expect(formatting1.numbering?.level).toBe(0);
      expect(formatting1.indentation).toBeUndefined();

      para.setNumbering(1, 2); // Change to level 2

      const formatting2 = para.getFormatting();
      expect(formatting2.numbering?.level).toBe(2);
      expect(formatting2.indentation).toBeUndefined();
    });

    it('should handle numbering applied multiple times', () => {
      const para = new Paragraph();
      para.setLeftIndent(1440);
      para.setNumbering(1, 0);
      para.setNumbering(2, 1); // Change to different list

      const formatting = para.getFormatting();
      expect(formatting.numbering).toEqual({ numId: 2, level: 1 });
      expect(formatting.indentation).toBeUndefined();
    });

    it('should handle indent set before and after numbering', () => {
      const para = new Paragraph();

      // Set indent before numbering
      para.setLeftIndent(720);
      expect(para.getFormatting().indentation?.left).toBe(720);

      // Apply numbering - should clear indent
      para.setNumbering(1, 0);
      expect(para.getFormatting().indentation).toBeUndefined();

      // Try to set indent after numbering - should warn but set anyway
      para.setLeftIndent(1440);
      expect(warnSpy).toHaveBeenCalled();
      expect(para.getFormatting().indentation?.left).toBe(1440);
    });

    it('should preserve right indent across numbering changes', () => {
      const para = new Paragraph();
      para.setLeftIndent(720);
      para.setRightIndent(720);
      para.setNumbering(1, 0);

      expect(para.getFormatting().indentation).toEqual({ right: 720 });

      para.setNumbering(1, 1); // Change level
      expect(para.getFormatting().indentation).toEqual({ right: 720 });
    });
  });

  describe('XML generation with conflict resolution', () => {
    it('should not include left indent in XML when numbering is present', () => {
      const para = new Paragraph();
      para.addText('Test item');
      para.setLeftIndent(1440);
      para.setNumbering(1, 0);

      const xml = para.toXML();
      const xmlStr = JSON.stringify(xml);

      // Should have numbering
      expect(xmlStr).toContain('numPr');
      expect(xmlStr).toContain('numId');

      // Should not have left indent (because it was cleared)
      expect(para.getFormatting().indentation).toBeUndefined();
    });

    it('should include right indent in XML with numbering', () => {
      const para = new Paragraph();
      para.addText('Test item');
      para.setRightIndent(720);
      para.setNumbering(1, 0);

      const xml = para.toXML();
      const xmlStr = JSON.stringify(xml);

      // Should have numbering
      expect(xmlStr).toContain('numPr');

      // Should have right indent
      expect(xmlStr).toContain('"w:right":720');
    });
  });

  describe('Numbering validation', () => {
    it('should throw error for negative numId', () => {
      const para = new Paragraph();
      expect(() => para.setNumbering(-1, 0)).toThrow('Numbering ID must be non-negative');
    });

    it('should throw error for level < 0', () => {
      const para = new Paragraph();
      expect(() => para.setNumbering(1, -1)).toThrow('Level must be between 0 and 8');
    });

    it('should throw error for level > 8', () => {
      const para = new Paragraph();
      expect(() => para.setNumbering(1, 9)).toThrow('Level must be between 0 and 8');
    });

    it('should accept valid levels 0-8', () => {
      const para = new Paragraph();
      for (let level = 0; level <= 8; level++) {
        expect(() => para.setNumbering(1, level)).not.toThrow();
        expect(para.getFormatting().numbering?.level).toBe(level);
      }
    });
  });
});

describe('Paragraph numId=0 suppression round-trip (ECMA-376 §17.3.1.19)', () => {
  it('should serialize numId=0 when numberingSuppressed is true', () => {
    const para = new Paragraph();
    para.addText('Suppressed numbering');
    para.formatting.numberingSuppressed = true;

    const { XMLBuilder } = require('../../src/xml/XMLBuilder');
    const xml = XMLBuilder.elementToString(para.toXML());

    expect(xml).toContain('<w:numId w:val="0"');
    expect(xml).toContain('<w:ilvl w:val="0"');
  });

  it('should round-trip numId=0 through buffer', async () => {
    const { Document } = require('../../src/core/Document');

    const doc = Document.create();
    const para = doc.createParagraph('No numbering');
    para.formatting.numberingSuppressed = true;

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);
    const loadedPara = loaded.getParagraphs()[0];

    expect(loadedPara.formatting.numberingSuppressed).toBe(true);
    expect(loadedPara.formatting.numbering).toBeUndefined();

    doc.dispose();
    loaded.dispose();
  });

  it('should round-trip numId=0 from injected XML', async () => {
    const { Document } = require('../../src/core/Document');

    const doc = Document.create();
    const para = doc.createParagraph('Test');
    para.setAlignment('left'); // Ensure w:pPr is emitted
    const buffer = await doc.toBuffer();
    doc.dispose();

    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buffer);
    let docXml = await zip.file('word/document.xml')!.async('string');
    // Inject numPr with numId=0 to suppress numbering
    docXml = docXml.replace(
      '</w:pPr>',
      '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr>'
    );
    zip.file('word/document.xml', docXml);
    const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    // Load and verify suppression is detected
    const loaded = await Document.loadFromBuffer(modifiedBuffer);
    expect(loaded.getParagraphs()[0]?.formatting.numberingSuppressed).toBe(true);

    // Save and reload to verify round-trip
    const buffer2 = await loaded.toBuffer();
    const loaded2 = await Document.loadFromBuffer(buffer2);
    expect(loaded2.getParagraphs()[0]?.formatting.numberingSuppressed).toBe(true);

    loaded.dispose();
    loaded2.dispose();
  });
});
