/**
 * Numbering XML Merge Tests
 *
 * Tests that numbering.xml is correctly preserved through load -> modify -> save cycles,
 * particularly extended namespace declarations and attributes from Word 2013+ features
 * (w15:restartNumberingAfterBreak, w16cid:durableId, mc:Ignorable, etc.).
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { AbstractNumbering } from '../../src/formatting/AbstractNumbering';
import { NumberingInstance } from '../../src/formatting/NumberingInstance';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: Creates a minimal DOCX buffer with custom numbering.xml content.
 * First creates a valid DOCX, then post-processes the ZIP to inject custom numbering.
 */
async function createDocxWithNumbering(numberingXml: string): Promise<Buffer> {
  const doc = Document.create();
  const para = new Paragraph().addText('Test numbered item');
  doc.addParagraph(para);
  const buffer = await doc.toBuffer();
  doc.dispose();

  // Post-process: replace numbering.xml in the saved ZIP
  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.updateFile(DOCX_PATHS.NUMBERING, numberingXml);
  return await zipHandler.toBuffer();
}

/** Numbering XML with extended namespaces and Word 2013+ attributes */
const NUMBERING_WITH_EXTENDED_NS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
             xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
             xmlns:v="urn:schemas-microsoft-com:vml"
             xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             xmlns:w10="urn:schemas-microsoft-com:office:word"
             xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
             xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
             xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
             xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
             xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
             xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
             xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
             xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
             mc:Ignorable="w14 w15 w16se w16cid wp14">
  <w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0">
    <w:nsid w:val="12345678"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0" w:tplc="AABBCCDD">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1" w15:restartNumberingAfterBreak="0">
    <w:nsid w:val="87654321"/>
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0" w:tplc="EEFF0011">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#xF0B7;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1" w16cid:durableId="111222333">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2" w16cid:durableId="444555666">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>`;

/** Simple numbering XML with only standard namespaces */
const NUMBERING_SIMPLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>`;

describe('Numbering XML Merge Tests', () => {
  describe('Namespace Preservation', () => {
    it('should preserve extended namespaces when numbering is not modified', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_WITH_EXTENDED_NS);
      const doc = await Document.loadFromBuffer(buffer);

      // Save without modifications
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      // Extract numbering.xml from output
      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // Should be identical to original — no modifications made
      expect(numberingXml).toContain('xmlns:w15=');
      expect(numberingXml).toContain('xmlns:w16cid=');
      expect(numberingXml).toContain('mc:Ignorable=');
      expect(numberingXml).toContain('w15:restartNumberingAfterBreak=');
      expect(numberingXml).toContain('w16cid:durableId=');
    });

    it('should preserve extended namespaces when adding new numbering definitions', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_WITH_EXTENDED_NS);
      const doc = await Document.loadFromBuffer(buffer);

      // Add a new numbered list (modifies numbering)
      const numId = doc.getNumberingManager().createBulletList();

      // Save
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      // Extract numbering.xml from output
      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // Root element should still have extended namespaces
      expect(numberingXml).toContain('xmlns:w15=');
      expect(numberingXml).toContain('xmlns:w16cid=');
      expect(numberingXml).toContain('mc:Ignorable=');

      // Original definitions should still have extended attributes
      expect(numberingXml).toContain('w15:restartNumberingAfterBreak="0"');
      expect(numberingXml).toContain('w16cid:durableId="111222333"');
      expect(numberingXml).toContain('w16cid:durableId="444555666"');

      // New definitions should also be present
      expect(numberingXml).toContain(`w:numId="${numId}"`);
    });

    it('should preserve unmodified abstractNum definitions when modifying others', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_WITH_EXTENDED_NS);
      const doc = await Document.loadFromBuffer(buffer);

      // Add a new abstractNum + instance (modifies numbering)
      doc.getNumberingManager().createNumberedList();

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // Original abstractNum 0 should be preserved with its extended attributes
      expect(numberingXml).toContain('w:abstractNumId="0"');
      expect(numberingXml).toContain('w:nsid w:val="12345678"');

      // Original abstractNum 1 should be preserved
      expect(numberingXml).toContain('w:abstractNumId="1"');
      expect(numberingXml).toContain('w:nsid w:val="87654321"');

      // Original num instances should be preserved with durableId
      expect(numberingXml).toContain('w:numId="1"');
      expect(numberingXml).toContain('w:numId="2"');
    });
  });

  describe('Merge with New Definitions', () => {
    it('should append new abstractNum before first w:num element', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_SIMPLE);
      const doc = await Document.loadFromBuffer(buffer);

      // Add a new list
      const numId = doc.getNumberingManager().createBulletList();

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // Original definition preserved
      expect(numberingXml).toContain('w:abstractNumId="0"');
      expect(numberingXml).toContain('w:numId="1"');

      // New definitions present
      expect(numberingXml).toContain(`w:numId="${numId}"`);

      // Structure should be valid: all abstractNum elements before num elements
      // Use ' w:abstractNumId' to avoid matching <w:abstractNumId> inside <w:num> children
      const firstAbstractNum = numberingXml.indexOf('<w:abstractNum ');
      const lastAbstractNum = numberingXml.lastIndexOf('<w:abstractNum ');
      const firstNum = numberingXml.indexOf('<w:num ');
      expect(firstAbstractNum).toBeLessThan(firstNum);
      expect(lastAbstractNum).toBeLessThan(firstNum);
    });

    it('should correctly merge when multiple new lists are added', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_SIMPLE);
      const doc = await Document.loadFromBuffer(buffer);

      // Add multiple new lists
      const numId1 = doc.getNumberingManager().createBulletList();
      const numId2 = doc.getNumberingManager().createNumberedList();

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // All should be present
      expect(numberingXml).toContain('w:abstractNumId="0"'); // original
      expect(numberingXml).toContain('w:numId="1"'); // original
      expect(numberingXml).toContain(`w:numId="${numId1}"`);
      expect(numberingXml).toContain(`w:numId="${numId2}"`);
    });
  });

  describe('Replacement of Existing Definitions', () => {
    it('should replace an existing abstractNum/num with same ID (no duplicates)', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_SIMPLE);
      const doc = await Document.loadFromBuffer(buffer);
      const mgr = doc.getNumberingManager();

      // Create a new abstractNum with same ID 0 to replace the existing one
      const replacement = AbstractNumbering.createNumberedList(0, 3);
      mgr.addAbstractNumbering(replacement);

      // Create a new num instance with same numId 1 to replace the existing one
      const instance = NumberingInstance.create({ numId: 1, abstractNumId: 0 });
      mgr.addInstance(instance);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // Should have exactly one abstractNumId="0" and one numId="1"
      const abstractNumMatches = numberingXml.match(/<w:abstractNum[\s][^>]*w:abstractNumId="0"/g);
      expect(abstractNumMatches).toHaveLength(1);

      const numMatches = numberingXml.match(/<w:num[\s][^>]*w:numId="1"/g);
      expect(numMatches).toHaveLength(1);

      // New content should be present (3 levels, not the original single level)
      // The replacement has 3 lvl elements
      const lvlMatches = numberingXml.match(/<w:lvl\s/g);
      expect(lvlMatches!.length).toBeGreaterThanOrEqual(3);
    });
  });

  describe('Removal of Definitions', () => {
    it('should remove deleted abstractNum and num from output XML', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_WITH_EXTENDED_NS);
      const doc = await Document.loadFromBuffer(buffer);
      const mgr = doc.getNumberingManager();

      // Remove abstractNum 1 (which cascades to remove num 2)
      mgr.removeAbstractNumbering(1);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // abstractNum 0 and num 1 should still be present
      expect(numberingXml).toContain('w:abstractNumId="0"');
      expect(numberingXml).toContain('w:numId="1"');

      // abstractNum 1 and num 2 should be removed
      expect(numberingXml).not.toMatch(/<w:abstractNum[^>]*w:abstractNumId="1"/);
      expect(numberingXml).not.toMatch(/<w:num[^>]*w:numId="2"/);
    });

    it('should remove a single num instance without removing its abstractNum', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_WITH_EXTENDED_NS);
      const doc = await Document.loadFromBuffer(buffer);
      const mgr = doc.getNumberingManager();

      // Remove only num instance 2 (keep abstractNum 1)
      mgr.removeInstance(2);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // abstractNum 1 should still be present
      expect(numberingXml).toContain('w:abstractNumId="1"');

      // num 2 should be removed
      expect(numberingXml).not.toMatch(/<w:num[^>]*w:numId="2"/);

      // num 1 should still be present
      expect(numberingXml).toContain('w:numId="1"');
    });
  });

  describe('Full Regeneration for New Documents', () => {
    it('should generate from scratch when no original numbering.xml exists', async () => {
      const doc = Document.create();

      // Add a list to a brand new document
      const numId = doc.getNumberingManager().createBulletList();
      const para = new Paragraph().addText('Bullet item');
      para.setNumbering(numId, 0);
      doc.addParagraph(para);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // Should have standard namespaces (generated from scratch)
      expect(numberingXml).toContain('xmlns:w=');
      expect(numberingXml).toContain(`w:numId="${numId}"`);
    });
  });

  describe('numIdMacAtCleanup ordering (CT_Numbering)', () => {
    const NUMBERING_WITH_MAC_CLEANUP = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:numIdMacAtCleanup w:val="8"/>
</w:numbering>`;

    it('should insert new w:num before w:numIdMacAtCleanup', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_WITH_MAC_CLEANUP);
      const doc = await Document.loadFromBuffer(buffer);

      // Add a new numbering definition — will insert a new w:num
      const numId = doc.getNumberingManager().createNumberedList();
      const para = new Paragraph().addText('New item');
      para.setNumbering(numId, 0);
      doc.addParagraph(para);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      // numIdMacAtCleanup should remain the last child
      const macPos = numberingXml.indexOf('<w:numIdMacAtCleanup');
      const closingPos = numberingXml.indexOf('</w:numbering>');
      const newNumPos = numberingXml.indexOf(`w:numId="${numId}"`);

      expect(macPos).toBeGreaterThan(0);
      expect(closingPos).toBeGreaterThan(macPos);
      expect(newNumPos).toBeGreaterThan(0);
      expect(newNumPos).toBeLessThan(macPos);
    });

    it('should keep numIdMacAtCleanup as last element when multiple nums are added', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_WITH_MAC_CLEANUP);
      const doc = await Document.loadFromBuffer(buffer);

      // Add two new lists
      const numId1 = doc.getNumberingManager().createNumberedList();
      const numId2 = doc.getNumberingManager().createBulletList();
      const para1 = new Paragraph().addText('Item 1');
      para1.setNumbering(numId1, 0);
      doc.addParagraph(para1);
      const para2 = new Paragraph().addText('Item 2');
      para2.setNumbering(numId2, 0);
      doc.addParagraph(para2);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      const macPos = numberingXml.indexOf('<w:numIdMacAtCleanup');
      const closingPos = numberingXml.indexOf('</w:numbering>');

      // Nothing between numIdMacAtCleanup and </w:numbering> except whitespace
      const between = numberingXml.slice(
        numberingXml.indexOf('>', macPos) + 1,
        closingPos
      ).trim();
      expect(between).toBe('');
    });

    it('should work normally when numIdMacAtCleanup is absent', async () => {
      const buffer = await createDocxWithNumbering(NUMBERING_SIMPLE);
      const doc = await Document.loadFromBuffer(buffer);

      const numId = doc.getNumberingManager().createNumberedList();
      const para = new Paragraph().addText('New');
      para.setNumbering(numId, 0);
      doc.addParagraph(para);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const numberingXml = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      expect(numberingXml).toContain(`w:numId="${numId}"`);
      expect(numberingXml).not.toContain('numIdMacAtCleanup');
    });

    it('should not place w:num elements after numIdMacAtCleanup even with abstractNum fallback', async () => {
      // Numbering with macAtCleanup but no existing w:num elements
      const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:numIdMacAtCleanup w:val="5"/>
</w:numbering>`;

      const buffer = await createDocxWithNumbering(numberingXml);
      const doc = await Document.loadFromBuffer(buffer);

      const numId = doc.getNumberingManager().createNumberedList();
      const para = new Paragraph().addText('Test');
      para.setNumbering(numId, 0);
      doc.addParagraph(para);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const output = zip.getFileAsString(DOCX_PATHS.NUMBERING) || '';

      const macPos = output.indexOf('<w:numIdMacAtCleanup');
      const numPos = output.lastIndexOf('<w:num ');

      expect(macPos).toBeGreaterThan(0);
      expect(numPos).toBeLessThan(macPos);
    });
  });
});
