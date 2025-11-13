/**
 * Tests for Run character style reference (Phase 4.1.1)
 * Tests character style linking and round-trip functionality
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import path from 'path';
import fs from 'fs';

describe('Run Character Style Reference - Round Trip Tests', () => {
  const testOutputDir = path.join(__dirname, '../output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  describe('Character Style Reference', () => {
    it('should set character style reference', () => {
      const run = new Run('Styled text');
      run.setCharacterStyle('Emphasis');

      const formatting = run.getFormatting();
      expect(formatting.characterStyle).toBe('Emphasis');
    });

    it('should round-trip character style reference', async () => {
      const doc = Document.create();
      const para = new Paragraph();

      const run = new Run('Text with character style', { characterStyle: 'Emphasis' });
      para.addRun(run);

      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);
      const paragraphs = loadedDoc.getParagraphs();

      expect(paragraphs).toHaveLength(1);
      const loadedRun = paragraphs[0]?.getRuns()[0];
      expect(loadedRun?.getFormatting().characterStyle).toBe('Emphasis');
    });

    it('should save character style to file and load correctly', async () => {
      const doc = Document.create();
      const para = new Paragraph();

      const run = new Run('Formatted text');
      run.setCharacterStyle('Strong');
      para.addRun(run);

      doc.addParagraph(para);

      const filePath = path.join(testOutputDir, 'run-character-style.docx');
      await doc.save(filePath);

      const loadedDoc = await Document.load(filePath);
      const paragraphs = loadedDoc.getParagraphs();
      const loadedRun = paragraphs[0]?.getRuns()[0];

      expect(loadedRun?.getFormatting().characterStyle).toBe('Strong');

      // Cleanup
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    });

    it('should support method chaining', () => {
      const run = new Run('Chained')
        .setCharacterStyle('Emphasis')
        .setBold()
        .setItalic();

      const formatting = run.getFormatting();
      expect(formatting.characterStyle).toBe('Emphasis');
      expect(formatting.bold).toBe(true);
      expect(formatting.italic).toBe(true);
    });

    it('should work with multiple runs with different styles', async () => {
      const doc = Document.create();
      const para = new Paragraph();

      para.addRun(new Run('Normal text '));
      para.addRun(new Run('Emphasized', { characterStyle: 'Emphasis' }));
      para.addRun(new Run(' and '));
      para.addRun(new Run('Strong', { characterStyle: 'Strong' }));

      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);
      const runs = loadedDoc.getParagraphs()[0]?.getRuns();

      expect(runs).toHaveLength(4);
      expect(runs?.[0]?.getFormatting().characterStyle).toBeUndefined();
      expect(runs?.[1]?.getFormatting().characterStyle).toBe('Emphasis');
      expect(runs?.[2]?.getFormatting().characterStyle).toBeUndefined();
      expect(runs?.[3]?.getFormatting().characterStyle).toBe('Strong');
    });
  });
});
