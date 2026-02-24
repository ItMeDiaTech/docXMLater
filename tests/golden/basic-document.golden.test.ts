/**
 * Golden file tests for document generation.
 *
 * Run with UPDATE_GOLDEN=true to regenerate expected files:
 *   UPDATE_GOLDEN=true npm test -- --testPathPattern=golden
 */

import * as fs from 'fs/promises';
import * as path from 'path';
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { expectDocxEqual } from '../helpers/compareDocx';

const GOLDEN_DIR = path.join(__dirname, 'fixtures');
const UPDATE_GOLDEN = process.env.UPDATE_GOLDEN === 'true';

async function ensureGoldenDir(): Promise<void> {
  await fs.mkdir(GOLDEN_DIR, { recursive: true });
}

/**
 * Helper: generate a doc, compare to golden file, or update if UPDATE_GOLDEN=true
 */
async function goldenTest(name: string, generate: () => Promise<Buffer>): Promise<void> {
  await ensureGoldenDir();
  const goldenPath = path.join(GOLDEN_DIR, `${name}.docx`);
  const actual = await generate();

  if (UPDATE_GOLDEN) {
    await fs.writeFile(goldenPath, actual);
    console.log(`Updated golden file: ${goldenPath}`);
    return;
  }

  let expected: Buffer;
  try {
    expected = await fs.readFile(goldenPath);
  } catch {
    throw new Error(
      `Golden file not found: ${goldenPath}\n` + `Run with UPDATE_GOLDEN=true to generate it.`
    );
  }

  await expectDocxEqual(expected, actual);
}

describe('Golden file tests', () => {
  it('basic empty document', async () => {
    await goldenTest('basic-empty', async () => {
      const doc = Document.create();
      const buffer = await doc.toBuffer();
      doc.dispose();
      return buffer;
    });
  });

  it('document with formatted paragraphs', async () => {
    await goldenTest('formatted-paragraphs', async () => {
      const doc = Document.create();
      const para1 = new Paragraph();
      para1.addText('Hello World', { bold: true });
      para1.setAlignment('center');
      doc.addParagraph(para1);

      const para2 = new Paragraph();
      para2.addText('Second paragraph', { italic: true, color: 'FF0000' });
      doc.addParagraph(para2);

      const buffer = await doc.toBuffer();
      doc.dispose();
      return buffer;
    });
  });

  it('document with styles', async () => {
    await goldenTest('styled-document', async () => {
      const doc = Document.create();
      const heading = new Paragraph();
      heading.addText('Heading text');
      heading.setStyle('Heading1');
      doc.addParagraph(heading);

      const body = new Paragraph();
      body.addText('Body content in Normal style');
      doc.addParagraph(body);

      const buffer = await doc.toBuffer();
      doc.dispose();
      return buffer;
    });
  });
});
