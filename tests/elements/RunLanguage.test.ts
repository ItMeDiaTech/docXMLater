/**
 * Tests for Run language (Phase 4.1.10)
 * Tests language code and round-trip functionality
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run, LanguageConfig } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import path from 'path';
import fs from 'fs';

describe('Run Language - Round Trip Tests', () => {
  const testOutputDir = path.join(__dirname, '../output');

  beforeAll(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  describe('Language', () => {
    it('should set language code', () => {
      const run = new Run('English text');
      run.setLanguage('en-US');

      const formatting = run.getFormatting();
      expect(formatting.language).toBe('en-US');
    });

    it('should round-trip French language through buffer', async () => {
      const doc = Document.create();
      const para = new Paragraph();

      const run = new Run('Texte français', {
        language: 'fr-FR',
      });
      para.addRun(run);

      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);
      const paragraphs = loadedDoc.getParagraphs();

      const loadedRun = paragraphs[0]?.getRuns()[0];
      expect(loadedRun?.getFormatting().language).toBe('fr-FR');
    });

    it('should round-trip Spanish language through file', async () => {
      const testFile = path.join(testOutputDir, 'test-language.docx');
      const doc = Document.create();
      const para = new Paragraph();

      const run = new Run('Texto español', {
        language: 'es-ES',
        bold: true,
      });
      para.addRun(run);

      doc.addParagraph(para);

      await doc.save(testFile);
      const loadedDoc = await Document.load(testFile);
      const paragraphs = loadedDoc.getParagraphs();

      const loadedRun = paragraphs[0]?.getRuns()[0];
      expect(loadedRun?.getFormatting().language).toBe('es-ES');
      expect(loadedRun?.getFormatting().bold).toBe(true);
    });
  });

  describe('CT_Language full support (ECMA-376 §17.3.2.20)', () => {
    it('should set and serialize LanguageConfig with all three attributes', () => {
      const run = new Run('Mixed text');
      run.setLanguage({ val: 'en-US', eastAsia: 'ja-JP', bidi: 'ar-SA' });

      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('w:val="en-US"');
      expect(xml).toContain('w:eastAsia="ja-JP"');
      expect(xml).toContain('w:bidi="ar-SA"');
    });

    it('should round-trip eastAsia and bidi language attributes through buffer', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run('Mixed language text');
      run.setLanguage({ val: 'en-US', eastAsia: 'zh-CN', bidi: 'he-IL' });
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const loadedRun = loaded.getParagraphs()[0]?.getRuns()[0];
      const lang = loadedRun?.getFormatting().language as LanguageConfig;

      expect(lang).toBeDefined();
      expect(lang.val).toBe('en-US');
      expect(lang.eastAsia).toBe('zh-CN');
      expect(lang.bidi).toBe('he-IL');

      doc.dispose();
      loaded.dispose();
    });

    it('should handle LanguageConfig with only eastAsia', () => {
      const run = new Run('Japanese text');
      run.setLanguage({ eastAsia: 'ja-JP' });

      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('w:eastAsia="ja-JP"');
      expect(xml).not.toContain('w:val=');
      expect(xml).not.toContain('w:bidi=');
    });

    it('should handle LanguageConfig with only bidi', () => {
      const run = new Run('Arabic text');
      run.setLanguage({ bidi: 'ar-SA' });

      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('w:bidi="ar-SA"');
      expect(xml).not.toContain('w:val=');
    });

    it('should preserve backward compatibility with string language', () => {
      const run = new Run('English text');
      run.setLanguage('en-US');

      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('w:val="en-US"');
      expect(xml).not.toContain('w:eastAsia');
      expect(xml).not.toContain('w:bidi');
    });
  });
});
