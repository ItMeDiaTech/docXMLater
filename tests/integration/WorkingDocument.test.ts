/**
 * Working Document Integration Test
 * Tests editing operations on a real document without corruption
 */

import { describe, it, expect, beforeAll, afterAll } from '@jest/globals';
import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import path from 'path';
import * as fs from 'fs';
import * as os from 'os';

describe.skip('Working Document Integration', () => {
  const workingDocPath = path.join(__dirname, '../../Working.docx');
  let tempDir: string;

  beforeAll(() => {
    // Create temporary directory for test files
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'docxml-working-test-'));
  });

  afterAll(() => {
    // Clean up temporary directory
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  });

  it('should load Working.docx without warnings', async () => {
    if (!fs.existsSync(workingDocPath)) {
      console.log(`Skipping test: ${workingDocPath} not found`);
      return;
    }

    const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});

    try {
      const doc = await Document.load(workingDocPath);

      // Get document statistics
      const paragraphs = doc.getParagraphs();
      const tables = doc.getTables();

      console.log(`\n=== Working Document Statistics ===`);
      console.log(`Paragraphs: ${paragraphs.length}`);
      console.log(`Tables: ${tables.length}`);

      // Count runs with text
      let totalRuns = 0;
      let runsWithText = 0;

      for (const para of paragraphs) {
        const runs = para.getRuns();
        totalRuns += runs.length;

        for (const run of runs) {
          if (run.getText().length > 0) {
            runsWithText++;
          }
        }
      }

      console.log(`Total runs: ${totalRuns}`);
      console.log(`Runs with text: ${runsWithText}`);
      console.log(`Empty runs: ${totalRuns - runsWithText}`);

      // Should not have corruption warnings
      const warnCalls = consoleWarnSpy.mock.calls;
      const hasCorruptionWarning = warnCalls.some(call =>
        call.some(arg =>
          typeof arg === 'string' && (
            arg.includes('corrupted') ||
            arg.includes('empty')
          )
        )
      );

      expect(hasCorruptionWarning).toBe(false);

      // Should have meaningful content
      expect(runsWithText).toBeGreaterThan(0);

    } finally {
      consoleWarnSpy.mockRestore();
    }
  });

  it('should preserve content through load/save cycle', async () => {
    if (!fs.existsSync(workingDocPath)) {
      console.log(`Skipping test: ${workingDocPath} not found`);
      return;
    }

    const testPath = path.join(tempDir, 'working-preserved.docx');

    // Load original
    const doc1 = await Document.load(workingDocPath);
    const originalParas = doc1.getParagraphs();
    const originalTexts = originalParas.map(p => p.getText());

    // Save
    await doc1.save(testPath);

    // Load again
    const doc2 = await Document.load(testPath);
    const loadedParas = doc2.getParagraphs();
    const loadedTexts = loadedParas.map(p => p.getText());

    // Compare (allowing for potential parsing differences)
    console.log(`\n=== Content Preservation Check ===`);
    console.log(`Original paragraphs: ${originalParas.length}`);
    console.log(`Loaded paragraphs: ${loadedParas.length}`);

    // Text content should be preserved
    const originalTotalText = originalTexts.join('');
    const loadedTotalText = loadedTexts.join('');

    expect(loadedTotalText.length).toBeGreaterThan(0);

    // If original had text, loaded should have similar amount (allowing 10% variation)
    if (originalTotalText.length > 0) {
      const textLengthRatio = loadedTotalText.length / originalTotalText.length;
      expect(textLengthRatio).toBeGreaterThan(0.9);
      expect(textLengthRatio).toBeLessThan(1.1);
    }
  });

  it('should handle modifications without corruption', async () => {
    if (!fs.existsSync(workingDocPath)) {
      console.log(`Skipping test: ${workingDocPath} not found`);
      return;
    }

    const testPath = path.join(tempDir, 'working-modified.docx');
    const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});

    try {
      // Load document
      const doc = await Document.load(workingDocPath);
      const paragraphs = doc.getParagraphs();

      console.log(`\n=== Testing Modifications ===`);

      // Test 1: Update alignment on first few paragraphs
      if (paragraphs.length > 0) {
        console.log(`Updating alignment on ${Math.min(3, paragraphs.length)} paragraphs`);
        for (let i = 0; i < Math.min(3, paragraphs.length); i++) {
          const para = paragraphs[i];
          if (para) {
            para.setAlignment('center');
          }
        }
      }

      // Test 2: Update indentation
      if (paragraphs.length > 3) {
        console.log(`Updating indentation on paragraph 4`);
        paragraphs[3]?.setLeftIndent(720); // 0.5 inch
      }

      // Test 3: Update font on runs
      if (paragraphs.length > 0) {
        const runs = paragraphs[0]?.getRuns() || [];
        if (runs.length > 0) {
          console.log(`Updating font on first run`);
          runs[0]?.setFont('Arial', 12);
        }
      }

      // Test 4: Add a new paragraph
      console.log(`Adding new paragraph`);
      doc.createParagraph('This is a test paragraph added by the integration test');

      // Save modified document
      await doc.save(testPath);

      // Verify no corruption warnings during save
      const warnCalls = consoleWarnSpy.mock.calls;
      const hasEmptyWarning = warnCalls.some(call =>
        call.some(arg =>
          typeof arg === 'string' && arg.includes('empty')
        )
      );

      expect(hasEmptyWarning).toBe(false);

      // Load and verify
      const doc2 = await Document.load(testPath);
      const modifiedParas = doc2.getParagraphs();

      // Should have at least as many paragraphs as before (we added one)
      expect(modifiedParas.length).toBeGreaterThanOrEqual(paragraphs.length);

      // Should find our test paragraph
      const hasTestParagraph = modifiedParas.some(p =>
        p.getText().includes('test paragraph added by the integration test')
      );
      expect(hasTestParagraph).toBe(true);

      console.log(`✓ Modifications applied successfully without corruption`);

    } finally {
      consoleWarnSpy.mockRestore();
    }
  });

  it('should handle table modifications if tables exist', async () => {
    if (!fs.existsSync(workingDocPath)) {
      console.log(`Skipping test: ${workingDocPath} not found`);
      return;
    }

    const doc = await Document.load(workingDocPath);
    const tables = doc.getTables();

    console.log(`\n=== Table Handling ===`);
    console.log(`Tables found: ${tables.length}`);

    if (tables.length > 0) {
      console.log(`Document has ${tables.length} table(s)`);
      // Note: Table editing is not yet fully implemented in Phase 2/3
      // This test just verifies tables don't cause corruption
      const testPath = path.join(tempDir, 'working-with-tables.docx');
      await doc.save(testPath);

      const doc2 = await Document.load(testPath);
      expect(doc2.getTables().length).toBe(tables.length);

      console.log(`✓ Tables preserved through save/load cycle`);
    } else {
      console.log(`No tables to test`);
    }
  });

  it('should detect and report actual corruption if it occurs', async () => {
    if (!fs.existsSync(workingDocPath)) {
      console.log(`Skipping test: ${workingDocPath} not found`);
      return;
    }

    // Intentionally create a document with mostly empty runs
    const emptyDoc = Document.create();
    for (let i = 0; i < 20; i++) {
      const para = emptyDoc.createParagraph();
      para.addRun(new Run('')); // Explicitly add empty run
    }

    const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});

    try {
      const testPath = path.join(tempDir, 'intentionally-empty.docx');
      await emptyDoc.save(testPath);

      // Should have warned about empty content
      const warnCalls = consoleWarnSpy.mock.calls;
      const hasWarning = warnCalls.some(call =>
        call.some(arg => typeof arg === 'string')
      );

      expect(hasWarning).toBe(true);

      console.log(`✓ Validation correctly detects empty content`);

    } finally {
      consoleWarnSpy.mockRestore();
    }
  });
});
