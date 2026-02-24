/**
 * Tests for w:default="1" preservation on Normal style
 */

import { Style } from '../../src/formatting/Style';
import { StylesManager } from '../../src/formatting/StylesManager';
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

describe('Style Default Preservation', () => {
  describe('Style.getIsDefault() / setIsDefault()', () => {
    it('should return false by default', () => {
      const style = Style.create({
        styleId: 'Test',
        name: 'Test',
        type: 'paragraph',
      });
      expect(style.getIsDefault()).toBe(false);
    });

    it('should return true when isDefault is set', () => {
      const style = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        isDefault: true,
      });
      expect(style.getIsDefault()).toBe(true);
    });

    it('should allow setting isDefault', () => {
      const style = Style.create({
        styleId: 'Test',
        name: 'Test',
        type: 'paragraph',
      });
      style.setIsDefault(true);
      expect(style.getIsDefault()).toBe(true);

      style.setIsDefault(false);
      expect(style.getIsDefault()).toBe(false);
    });

    it("should emit w:default='1' in toXML() when isDefault is true", () => {
      const style = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        isDefault: true,
      });

      const xml = style.toXML();
      // XMLElement has attributes — check them
      expect(xml.attributes?.['w:default']).toBe('1');
    });

    it('should not emit w:default in toXML() when isDefault is false', () => {
      const style = Style.create({
        styleId: 'Custom',
        name: 'Custom',
        type: 'paragraph',
        isDefault: false,
      });

      const xml = style.toXML();
      expect(xml.attributes?.['w:default']).toBeUndefined();
    });
  });

  describe('StylesManager.addStyle() isDefault preservation', () => {
    it('should preserve isDefault when replacing existing default style', () => {
      const manager = new StylesManager(false);

      // Add Normal with isDefault: true
      const original = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        isDefault: true,
      });
      manager.addStyle(original);

      // Replace with new Normal without isDefault
      const replacement = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        runFormatting: { bold: true },
      });
      manager.addStyle(replacement);

      // Should have inherited isDefault from existing
      const result = manager.getStyle('Normal');
      expect(result?.getIsDefault()).toBe(true);
    });

    it('should not set isDefault when replacing non-default style', () => {
      const manager = new StylesManager(false);

      const original = Style.create({
        styleId: 'Custom1',
        name: 'Custom1',
        type: 'paragraph',
      });
      manager.addStyle(original);

      const replacement = Style.create({
        styleId: 'Custom1',
        name: 'Custom1 Updated',
        type: 'paragraph',
      });
      manager.addStyle(replacement);

      expect(manager.getStyle('Custom1')?.getIsDefault()).toBe(false);
    });

    it('should not override isDefault when replacement already has isDefault: true', () => {
      const manager = new StylesManager(false);

      const original = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        isDefault: true,
      });
      manager.addStyle(original);

      // Replacement also has isDefault: true
      const replacement = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        isDefault: true,
        runFormatting: { color: 'FF0000' },
      });
      manager.addStyle(replacement);

      expect(manager.getStyle('Normal')?.getIsDefault()).toBe(true);
    });
  });

  describe('Round-trip w:default preservation', () => {
    it("should preserve w:default='1' through load-save cycle", async () => {
      // Create a doc, save it, load it back
      const doc = Document.create();
      doc.createParagraph('Test');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Load and verify Normal has isDefault
      const doc2 = await Document.loadFromBuffer(buffer1);

      // Now add a new Normal-style replacement without explicitly setting isDefault
      const newNormal = Style.create({
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        runFormatting: { font: 'Arial', size: 11 },
      });
      doc2.getStylesManager().addStyle(newNormal);

      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Check the output XML for w:default="1" on Normal
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer2);
      const stylesXml = zip.getFileAsString(DOCX_PATHS.STYLES) || '';
      zip.clear();

      // Find the Normal style element and check for w:default
      const normalMatch = stylesXml.match(/<w:style[^>]*w:styleId="Normal"[^>]*>/);
      expect(normalMatch).not.toBeNull();
      expect(normalMatch![0]).toContain('w:default="1"');
    });
  });
});
