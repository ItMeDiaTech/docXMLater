/**
 * Tests for browser extension URL prefix sanitization in hyperlinks.
 *
 * Verifies that corrupted URLs (wrapped by browser extensions like Adobe Acrobat)
 * are automatically sanitized during document load and via validateAndFix().
 */

import { Document } from '../../src/core/Document';
import { Relationship } from '../../src/core/Relationship';
import { RelationshipManager } from '../../src/core/RelationshipManager';
import { Hyperlink } from '../../src/elements/Hyperlink';

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

describe('Hyperlink URL Sanitization', () => {
  describe('RelationshipManager.fromXml() auto-sanitization', () => {
    it('should strip Chrome extension prefix from hyperlink targets', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="${HYPERLINK_REL_TYPE}"
            Target="chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/https://example.com/doc.pdf"
            TargetMode="External"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
            Target="styles.xml"/>
        </Relationships>`;

      const manager = RelationshipManager.fromXml(xml);

      // Hyperlink should be sanitized
      const hyperlink = manager.getRelationship('rId1');
      expect(hyperlink).toBeDefined();
      expect(hyperlink!.getTarget()).toBe('https://example.com/doc.pdf');

      // Non-hyperlink relationship should be untouched
      const styles = manager.getRelationship('rId2');
      expect(styles).toBeDefined();
      expect(styles!.getTarget()).toBe('styles.xml');
    });

    it('should fix broken protocol after stripping extension prefix', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId20" Type="${HYPERLINK_REL_TYPE}"
            Target="chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/https:/aetnao365.sharepoint.com/sites/test"
            TargetMode="External"/>
        </Relationships>`;

      const manager = RelationshipManager.fromXml(xml);
      const rel = manager.getRelationship('rId20');
      expect(rel).toBeDefined();
      expect(rel!.getTarget()).toBe('https://aetnao365.sharepoint.com/sites/test');
    });

    it('should not modify clean hyperlink URLs', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="${HYPERLINK_REL_TYPE}"
            Target="https://example.com/clean-url"
            TargetMode="External"/>
        </Relationships>`;

      const manager = RelationshipManager.fromXml(xml);
      const rel = manager.getRelationship('rId1');
      expect(rel!.getTarget()).toBe('https://example.com/clean-url');
    });
  });

  describe('Hyperlink.validateAndFix() with extension-prefixed URLs', () => {
    it('should strip extension prefix and report fix', async () => {
      const hyperlink = new Hyperlink({
        url: 'chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/https://example.com/doc.pdf',
      });

      const result = await hyperlink.validateAndFix({ fixCommonIssues: true });
      expect(result.fixed.length).toBeGreaterThan(0);
      expect(result.fixed).toContain('Stripped Chrome extension URL prefix');
    });

    it('should strip extension prefix and fix broken protocol together', async () => {
      const hyperlink = new Hyperlink({
        url: 'chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/https:/sharepoint.com/sites/test',
      });

      const result = await hyperlink.validateAndFix({ fixCommonIssues: true });
      expect(result.fixed).toContain('Stripped Chrome extension URL prefix');
      expect(result.fixed).toContain('Fixed broken protocol (added missing slash)');
    });

    it('should not report fixes for clean URLs', async () => {
      const hyperlink = new Hyperlink({ url: 'https://example.com' });

      const result = await hyperlink.validateAndFix({ fixCommonIssues: true });
      const extensionFixes = result.fixed.filter((f) => f.includes('extension'));
      expect(extensionFixes).toHaveLength(0);
    });
  });

  describe('Round-trip sanitization', () => {
    it('should sanitize corrupted hyperlink URL through document load', async () => {
      // Create a document with a hyperlink so the relationship is referenced
      const doc = Document.create();
      const para = doc.createParagraph('');
      const link = para.addHyperlink(
        'chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/https:/aetnao365.sharepoint.com/sites/test'
      );
      link.setText('Click here');

      // Save to buffer
      const buffer = await doc.toBuffer();
      doc.dispose();

      // Load from buffer — RelationshipManager.fromXml() sanitizes at parse time
      const reloaded = await Document.loadFromBuffer(buffer);

      const hyperlinkRels = reloaded
        .getRelationshipManager()
        .getRelationshipsByType(HYPERLINK_REL_TYPE);
      expect(hyperlinkRels.length).toBeGreaterThan(0);
      const rel = hyperlinkRels[0]!;
      expect(rel.getTarget()).toBe('https://aetnao365.sharepoint.com/sites/test');

      reloaded.dispose();
    });
  });
});
