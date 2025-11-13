/**
 * Diagnostic test for ErrorDoc.docx corruption issue
 * Investigates why all text content was stripped from the document
 */

import { describe, it, expect } from '@jest/globals';
import { Document } from '../../src/core/Document';
import { XMLParser } from '../../src/xml/XMLParser';
import path from 'path';
import * as fs from 'fs';

describe.skip('ErrorDoc.docx Diagnostic', () => {
  const errorDocPath = path.join(__dirname, '../../ErrorDoc.docx');

  it('should exist', () => {
    expect(fs.existsSync(errorDocPath)).toBe(true);
  });

  it('should load and inspect document structure', async () => {
    const doc = await Document.load(errorDocPath);

    // Get all paragraphs
    const paragraphs = doc.getParagraphs();
    console.log(`\n=== Document Structure ===`);
    console.log(`Total paragraphs: ${paragraphs.length}`);

    // Count runs and empty runs
    let totalRuns = 0;
    let emptyRuns = 0;
    let runsWithText = 0;

    for (const para of paragraphs) {
      const runs = para.getRuns();
      totalRuns += runs.length;

      for (const run of runs) {
        const text = run.getText();
        if (text.length === 0) {
          emptyRuns++;
        } else {
          runsWithText++;
          console.log(`Found text: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"`);
        }
      }
    }

    console.log(`\nTotal runs: ${totalRuns}`);
    console.log(`Empty runs: ${emptyRuns}`);
    console.log(`Runs with text: ${runsWithText}`);

    // Check parse warnings
    const warnings = doc.getParseWarnings();
    console.log(`\nParse warnings: ${warnings.length}`);
    for (const warning of warnings) {
      console.log(`  ${warning.element}: ${warning.error.message}`);
    }

    expect(paragraphs.length).toBeGreaterThan(0);
  });

  it('should inspect raw XML from the document', async () => {
    const { ZipHandler } = await import('../../src/zip/ZipHandler');
    const zip = new ZipHandler();
    await zip.load(errorDocPath);

    const docXml = zip.getFileAsString('word/document.xml');
    if (!docXml) {
      throw new Error('word/document.xml not found in ErrorDoc.docx');
    }

    // Extract body content
    const bodyContent = XMLParser.extractBody(docXml);
    console.log(`\n=== Raw XML Inspection ===`);
    console.log(`Document XML size: ${docXml.length} bytes`);
    console.log(`Body content size: ${bodyContent.length} bytes`);

    // Count all <w:t> tags
    const textTagMatches = bodyContent.match(/<w:t[^>]*>/g) || [];
    console.log(`Number of <w:t> tags: ${textTagMatches.length}`);

    // Extract all text using XMLParser
    const paragraphXmls = XMLParser.extractElements(bodyContent, 'w:p');
    console.log(`Number of paragraphs: ${paragraphXmls.length}`);

    // Inspect first 3 paragraphs in detail
    console.log(`\n=== First 3 Paragraphs (Raw XML) ===`);
    for (let i = 0; i < Math.min(3, paragraphXmls.length); i++) {
      const paraXml = paragraphXmls[i];
      if (!paraXml) continue;

      console.log(`\nParagraph ${i + 1}:`);
      console.log(paraXml.substring(0, 500));

      // Extract runs
      const runXmls = XMLParser.extractElements(paraXml, 'w:r');
      console.log(`  Runs in paragraph: ${runXmls.length}`);

      for (let j = 0; j < Math.min(2, runXmls.length); j++) {
        const runXml = runXmls[j];
        if (!runXml) continue;

        const extractedText = XMLParser.extractText(runXml);
        console.log(`  Run ${j + 1} extracted text: "${extractedText}" (length: ${extractedText.length})`);

        // Check if <w:t> tag exists
        if (runXml.includes('<w:t')) {
          // Find the actual content between tags
          const textTagStart = runXml.indexOf('<w:t');
          const textTagEnd = runXml.indexOf('</w:t>', textTagStart);
          const textTagContent = runXml.substring(textTagStart, textTagEnd + 6);
          console.log(`  Raw <w:t> tag: ${textTagContent}`);
        }
      }
    }

    expect(paragraphXmls.length).toBeGreaterThan(0);
  });

  it('should test XMLParser.extractText() with sample XML', () => {
    console.log(`\n=== XMLParser.extractText() Tests ===`);

    // Test cases
    const testCases = [
      {
        name: 'Simple text',
        xml: '<w:r><w:t>Hello World</w:t></w:r>',
        expected: 'Hello World'
      },
      {
        name: 'Empty text',
        xml: '<w:r><w:t></w:t></w:r>',
        expected: ''
      },
      {
        name: 'Text with xml:space attribute',
        xml: '<w:r><w:t xml:space="preserve">Hello World</w:t></w:r>',
        expected: 'Hello World'
      },
      {
        name: 'Multiple text elements',
        xml: '<w:r><w:t>Hello</w:t><w:t> World</w:t></w:r>',
        expected: 'Hello World'
      },
      {
        name: 'Text with XML entities',
        xml: '<w:r><w:t>&lt;Hello&gt; &amp; &quot;World&quot;</w:t></w:r>',
        expected: '<Hello> & "World"' // Expected after unescaping
      },
      {
        name: 'Text with special chars',
        xml: '<w:r><w:t>Hello\nWorld\tTest</w:t></w:r>',
        expected: 'Hello\nWorld\tTest'
      }
    ];

    for (const testCase of testCases) {
      const result = XMLParser.extractText(testCase.xml);
      console.log(`\nTest: ${testCase.name}`);
      console.log(`  Input: ${testCase.xml}`);
      console.log(`  Expected: "${testCase.expected}"`);
      console.log(`  Got: "${result}"`);
      console.log(`  Match: ${result === testCase.expected ? '✓' : '✗'}`);
    }
  });
});
