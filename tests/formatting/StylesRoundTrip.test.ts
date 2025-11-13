/**
 * Style Round-Trip Tests
 * Tests that styles are correctly preserved through load -> save -> load cycles
 * Specifically tests the color parsing bug fix (colors should not get mixed with sizes)
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import * as path from 'path';
import * as fs from 'fs/promises';

describe('Style Round-Trip Tests', () => {
  const TEMP_DIR = path.join(__dirname, '..', '..', 'temp-test-output');
  const TEST_FILE = path.join(__dirname, '..', '..', 'Test6_BaseFile.docx');

  // Check if test file exists
  let testFileExists = false;

  beforeAll(async () => {
    // Create temp directory for test outputs
    try {
      await fs.mkdir(TEMP_DIR, { recursive: true });
    } catch (error) {
      // Directory might already exist
    }

    // Check if test file exists
    try {
      await fs.access(TEST_FILE);
      testFileExists = true;
    } catch {
      testFileExists = false;
    }
  });

  afterAll(async () => {
    // Clean up temp directory
    try {
      const files = await fs.readdir(TEMP_DIR);
      for (const file of files) {
        await fs.unlink(path.join(TEMP_DIR, file));
      }
      await fs.rmdir(TEMP_DIR);
    } catch (error) {
      // Ignore cleanup errors
    }
  });

  describe('Color Preservation', () => {
    it('should preserve hex color "000000" (black) through round-trip', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      // Load document with styles
      const doc1 = await Document.load(TEST_FILE);
      const styles1 = doc1.getStyles();
      const heading1_1 = styles1.find(
        (s) => s.getStyleId() === 'Heading1' || s.getName() === 'Heading1'
      );

      expect(heading1_1).toBeDefined();
      const props1 = heading1_1!.getProperties();
      expect(props1.runFormatting?.color).toBe('000000');

      // Save and reload
      const tempFile = path.join(TEMP_DIR, 'round-trip-test-1.docx');
      await doc1.save(tempFile);

      const doc2 = await Document.load(tempFile);
      const styles2 = doc2.getStyles();
      const heading1_2 = styles2.find(
        (s) => s.getStyleId() === 'Heading1' || s.getName() === 'Heading1'
      );

      expect(heading1_2).toBeDefined();
      const props2 = heading1_2!.getProperties();
      expect(props2.runFormatting?.color).toBe('000000');
    });

    it('should preserve theme color "0f4761" through round-trip', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      const doc1 = await Document.load(TEST_FILE);
      const styles1 = doc1.getStyles();
      const heading3_1 = styles1.find(
        (s) => s.getStyleId() === 'Heading3' || s.getName() === 'Heading3'
      );

      expect(heading3_1).toBeDefined();
      const props1 = heading3_1!.getProperties();
      expect(props1.runFormatting?.color).toBe('0f4761');

      // Save and reload
      const tempFile = path.join(TEMP_DIR, 'round-trip-test-2.docx');
      await doc1.save(tempFile);

      const doc2 = await Document.load(tempFile);
      const styles2 = doc2.getStyles();
      const heading3_2 = styles2.find(
        (s) => s.getStyleId() === 'Heading3' || s.getName() === 'Heading3'
      );

      expect(heading3_2).toBeDefined();
      const props2 = heading3_2!.getProperties();
      expect(props2.runFormatting?.color).toBe('0f4761');
    });

    it('should not confuse color with size values', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      const doc = await Document.load(TEST_FILE);
      const styles = doc.getStyles();
      const heading1 = styles.find(
        (s) => s.getStyleId() === 'Heading1' || s.getName() === 'Heading1'
      );

      expect(heading1).toBeDefined();
      const props = heading1!.getProperties();

      // Heading1 has size 18pt (36 half-points) and color 000000
      // The bug was: color became "36" (the size value)
      expect(props.runFormatting?.size).toBe(18);
      expect(props.runFormatting?.color).toBe('000000');
      expect(props.runFormatting?.color).not.toBe('36'); // Should NOT be the size value
    });

    it('should preserve colors in all heading styles', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      const doc = await Document.load(TEST_FILE);
      const styles = doc.getStyles();

      // Test Heading1 (black)
      const h1 = styles.find((s) => s.getStyleId() === 'Heading1');
      if (h1) {
        const props = h1.getProperties();
        expect(props.runFormatting?.color).toBe('000000');
      }

      // Test Heading3 (theme color)
      const h3 = styles.find((s) => s.getStyleId() === 'Heading3');
      if (h3) {
        const props = h3.getProperties();
        expect(props.runFormatting?.color).toBe('0f4761');
      }

      // Test Heading4 (theme color)
      const h4 = styles.find((s) => s.getStyleId() === 'Heading4');
      if (h4) {
        const props = h4.getProperties();
        expect(props.runFormatting?.color).toBe('0f4761');
      }
    });
  });

  describe('Full Style Preservation', () => {
    it('should preserve all run formatting properties through round-trip', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      const doc1 = await Document.load(TEST_FILE);
      const styles1 = doc1.getStyles();
      const heading1_1 = styles1.find((s) => s.getStyleId() === 'Heading1');

      expect(heading1_1).toBeDefined();
      const props1 = heading1_1!.getProperties();

      // Save and reload
      const tempFile = path.join(TEMP_DIR, 'round-trip-full-props.docx');
      await doc1.save(tempFile);

      const doc2 = await Document.load(tempFile);
      const styles2 = doc2.getStyles();
      const heading1_2 = styles2.find((s) => s.getStyleId() === 'Heading1');

      expect(heading1_2).toBeDefined();
      const props2 = heading1_2!.getProperties();

      // Compare all properties
      expect(props2.runFormatting?.bold).toBe(props1.runFormatting?.bold);
      expect(props2.runFormatting?.size).toBe(props1.runFormatting?.size);
      expect(props2.runFormatting?.color).toBe(props1.runFormatting?.color);
      expect(props2.runFormatting?.font).toBe(props1.runFormatting?.font);
    });

    it('should preserve multiple styles with different colors', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      const doc1 = await Document.load(TEST_FILE);
      const tempFile = path.join(TEMP_DIR, 'round-trip-multiple-styles.docx');
      await doc1.save(tempFile);

      const doc2 = await Document.load(tempFile);
      const styles2 = doc2.getStyles();

      // Check multiple styles
      const stylesToCheck = ['Heading1', 'Heading3', 'Heading4'];

      for (const styleId of stylesToCheck) {
        const style1 = doc1
          .getStyles()
          .find((s) => s.getStyleId() === styleId);
        const style2 = styles2.find((s) => s.getStyleId() === styleId);

        if (style1 && style2) {
          const props1 = style1.getProperties();
          const props2 = style2.getProperties();

          // Colors should match
          expect(props2.runFormatting?.color).toBe(
            props1.runFormatting?.color
          );

          // Sizes should match
          expect(props2.runFormatting?.size).toBe(props1.runFormatting?.size);
        }
      }
    });
  });

  describe('Edge Cases', () => {
    it('should handle styles with no color', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      const doc = await Document.load(TEST_FILE);
      const styles = doc.getStyles();

      // Heading2 might not have explicit color
      const h2 = styles.find((s) => s.getStyleId() === 'Heading2');
      if (h2) {
        const props = h2.getProperties();
        // Should be undefined or valid hex, NOT a size value
        if (props.runFormatting?.color) {
          expect(props.runFormatting.color).toMatch(/^[0-9A-Fa-f]{6}$/);
        }
      }
    });

    it('should handle styles with size but no color', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      const doc = await Document.load(TEST_FILE);
      const styles = doc.getStyles();

      for (const style of styles) {
        const props = style.getProperties();

        // If style has size but no color, color should be undefined
        if (props.runFormatting?.size && !props.runFormatting?.color) {
          expect(props.runFormatting.color).toBeUndefined();
        }

        // If style has color, it should be valid hex
        if (props.runFormatting?.color) {
          expect(props.runFormatting.color).toMatch(/^[0-9A-Fa-f]{3,6}$/);
          // Should not be a 1-2 digit number (like 15, 28, 36) which are font sizes
          // But "000000" is valid (black color), so check length
          const colorNum = parseInt(props.runFormatting.color, 10);
          if (!isNaN(colorNum) && colorNum < 100 && props.runFormatting.color.length <= 2) {
            // Values like "15", "28", "36" (1-2 digits) are likely font sizes, not colors
            throw new Error(
              `Color value "${props.runFormatting.color}" looks like a font size`
            );
          }
        }
      }
    });

    it('should handle multiple round-trips without degradation', async () => {
      if (!testFileExists) {
        console.warn('Test skipped: Test6_BaseFile.docx not found');
        return;
      }
      let doc = await Document.load(TEST_FILE);
      const originalStyles = doc.getStyles();
      const heading1Original = originalStyles.find(
        (s) => s.getStyleId() === 'Heading1'
      );
      const originalColor = heading1Original?.getProperties().runFormatting
        ?.color;
      const originalSize = heading1Original?.getProperties().runFormatting
        ?.size;

      // Perform 3 round-trips
      for (let i = 0; i < 3; i++) {
        const tempFile = path.join(
          TEMP_DIR,
          `round-trip-multiple-${i}.docx`
        );
        await doc.save(tempFile);
        doc = await Document.load(tempFile);

        const styles = doc.getStyles();
        const heading1 = styles.find((s) => s.getStyleId() === 'Heading1');

        expect(heading1).toBeDefined();
        const props = heading1!.getProperties();
        expect(props.runFormatting?.color).toBe(originalColor);
        expect(props.runFormatting?.size).toBe(originalSize);
      }
    });
  });

  describe('Created Styles', () => {
    it('should preserve colors in programmatically created styles', async () => {
      const doc = Document.create();

      // Create style with specific color
      const customStyle = new Style({
        type: 'paragraph',
        styleId: 'CustomRed',
        name: 'Custom Red Style',
        runFormatting: {
          color: 'FF0000',
          size: 14,
        },
      });

      doc.addStyle(customStyle);

      // Save and reload
      const tempFile = path.join(TEMP_DIR, 'custom-style-test.docx');
      await doc.save(tempFile);

      const doc2 = await Document.load(tempFile);
      const styles = doc2.getStyles();
      const reloadedStyle = styles.find((s) => s.getStyleId() === 'CustomRed');

      expect(reloadedStyle).toBeDefined();
      const props = reloadedStyle!.getProperties();
      expect(props.runFormatting?.color).toBe('FF0000');
      expect(props.runFormatting?.size).toBe(14);
      expect(props.runFormatting?.color).not.toBe('14'); // NOT the size!
    });

    it('should handle black color (000000) without converting to 0', async () => {
      const doc = Document.create();

      const blackStyle = new Style({
        type: 'paragraph',
        styleId: 'BlackText',
        name: 'Black Text',
        runFormatting: {
          color: '000000',
          size: 12,
        },
      });

      doc.addStyle(blackStyle);

      const tempFile = path.join(TEMP_DIR, 'black-color-test.docx');
      await doc.save(tempFile);

      const doc2 = await Document.load(tempFile);
      const reloadedStyle = doc2
        .getStyles()
        .find((s) => s.getStyleId() === 'BlackText');

      expect(reloadedStyle).toBeDefined();
      const props = reloadedStyle!.getProperties();
      expect(props.runFormatting?.color).toBe('000000');
      expect(typeof props.runFormatting?.color).toBe('string');
    });
  });
});
