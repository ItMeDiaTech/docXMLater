/**
 * Gap Tests for Run Content Elements
 *
 * Tests run content types that are parsed from XML but lack dedicated tests:
 * - addTab() / addBreak() / addCarriageReturn()
 * - softHyphen / noBreakHyphen (via createFromContent)
 * - symbol (via createFromContent)
 * - pageNumber (via createFromContent)
 * - positionTab (via createFromContent)
 * - lastRenderedPageBreak (via createFromContent)
 * - separator / continuationSeparator (via createFromContent)
 * - annotationRef (via createFromContent)
 * - dayShort/monthShort/yearShort/dayLong/monthLong/yearLong (via createFromContent)
 * - fieldChar (via createFromContent)
 */

import { Run, RunContent } from '../../src/elements/Run';
import { Paragraph } from '../../src/elements/Paragraph';
import { Document } from '../../src/core/Document';
import { XMLElement } from '../../src/xml/XMLBuilder';

function filterXMLElements(children?: (XMLElement | string)[]): XMLElement[] {
  return (children || []).filter((c): c is XMLElement => typeof c !== 'string');
}

describe('Run Content Gap Tests', () => {
  describe('addTab()', () => {
    test('should add tab to run', () => {
      const run = new Run('Before');
      run.addTab().appendText('After');
      const content = run.getContent();
      expect(content.some((c) => c.type === 'tab')).toBe(true);
    });

    test('should generate w:tab in XML', () => {
      const run = new Run('Before');
      run.addTab();
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:tab')).toBe(true);
    });

    test('should round-trip tab', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run('Before');
      run.addTab().appendText('After');
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const loadedContent = loaded.getParagraphs()[0]?.getRuns()[0]?.getContent();
      expect(loadedContent?.some((c) => c.type === 'tab')).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('addBreak()', () => {
    test('should add line break', () => {
      const run = new Run('Text');
      run.addBreak();
      const content = run.getContent();
      expect(content.some((c) => c.type === 'break')).toBe(true);
    });

    test('should round-trip page break', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run('Before');
      run.addBreak('page');
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const content = loaded.getParagraphs()[0]?.getRuns()[0]?.getContent();
      const breakContent = content?.find((c) => c.type === 'break');
      expect(breakContent).toBeDefined();
      expect(breakContent?.breakType).toBe('page');

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip column break', async () => {
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run('Before');
      run.addBreak('column');
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const content = loaded.getParagraphs()[0]?.getRuns()[0]?.getContent();
      const breakContent = content?.find((c) => c.type === 'break');
      expect(breakContent).toBeDefined();
      expect(breakContent?.breakType).toBe('column');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('addCarriageReturn()', () => {
    test('should add carriage return', () => {
      const run = new Run('Text');
      run.addCarriageReturn();
      expect(run.getContent().some((c) => c.type === 'carriageReturn')).toBe(true);
    });

    test('should generate w:cr in XML', () => {
      const run = new Run('Text');
      run.addCarriageReturn();
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:cr')).toBe(true);
    });
  });

  describe('softHyphen (via createFromContent)', () => {
    test('should generate w:softHyphen in XML', () => {
      const content: RunContent[] = [
        { type: 'text', value: 'long' },
        { type: 'softHyphen' },
        { type: 'text', value: 'word' },
      ];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:softHyphen')).toBe(true);
    });
  });

  describe('noBreakHyphen (via createFromContent)', () => {
    test('should generate w:noBreakHyphen in XML', () => {
      const content: RunContent[] = [
        { type: 'text', value: 'non' },
        { type: 'noBreakHyphen' },
        { type: 'text', value: 'breaking' },
      ];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:noBreakHyphen')).toBe(true);
    });
  });

  describe('symbol (via createFromContent)', () => {
    test('should generate w:sym in XML', () => {
      const content: RunContent[] = [
        { type: 'symbol', symbolFont: 'Wingdings', symbolChar: 'F0FC' },
      ];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      const sym = children.find((c) => c.name === 'w:sym');
      expect(sym).toBeDefined();
      expect(sym?.attributes?.['w:font']).toBe('Wingdings');
      expect(sym?.attributes?.['w:char']).toBe('F0FC');
    });
  });

  describe('pageNumber (via createFromContent)', () => {
    test('should generate w:pgNum in XML', () => {
      const content: RunContent[] = [{ type: 'pageNumber' }];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:pgNum')).toBe(true);
    });
  });

  describe('positionTab (via createFromContent)', () => {
    test('should generate w:ptab in XML with attributes', () => {
      const content: RunContent[] = [
        {
          type: 'positionTab',
          ptabAlignment: 'center',
          ptabRelativeTo: 'margin',
          ptabLeader: 'dot',
        },
      ];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      const ptab = children.find((c) => c.name === 'w:ptab');
      expect(ptab).toBeDefined();
      expect(ptab?.attributes?.['w:alignment']).toBe('center');
      expect(ptab?.attributes?.['w:relativeTo']).toBe('margin');
      expect(ptab?.attributes?.['w:leader']).toBe('dot');
    });
  });

  describe('lastRenderedPageBreak (via createFromContent)', () => {
    test('should generate w:lastRenderedPageBreak in XML', () => {
      const content: RunContent[] = [
        { type: 'lastRenderedPageBreak' },
        { type: 'text', value: 'After break' },
      ];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:lastRenderedPageBreak')).toBe(true);
    });
  });

  describe('separator / continuationSeparator (via createFromContent)', () => {
    test('should generate w:separator in XML', () => {
      const content: RunContent[] = [{ type: 'separator' }];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:separator')).toBe(true);
    });

    test('should generate w:continuationSeparator in XML', () => {
      const content: RunContent[] = [{ type: 'continuationSeparator' }];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:continuationSeparator')).toBe(true);
    });
  });

  describe('annotationRef (via createFromContent)', () => {
    test('should generate w:annotationRef in XML', () => {
      const content: RunContent[] = [{ type: 'annotationRef' }];
      const run = Run.createFromContent(content);
      const xml = run.toXML();
      const children = filterXMLElements(xml.children);
      expect(children.some((c) => c.name === 'w:annotationRef')).toBe(true);
    });
  });

  describe('Date fields (via createFromContent)', () => {
    test('should generate all short/long date fields', () => {
      const dateTypes: Array<{ type: RunContent['type']; xmlName: string }> = [
        { type: 'dayShort', xmlName: 'w:dayShort' },
        { type: 'dayLong', xmlName: 'w:dayLong' },
        { type: 'monthShort', xmlName: 'w:monthShort' },
        { type: 'monthLong', xmlName: 'w:monthLong' },
        { type: 'yearShort', xmlName: 'w:yearShort' },
        { type: 'yearLong', xmlName: 'w:yearLong' },
      ];

      for (const { type, xmlName } of dateTypes) {
        const content: RunContent[] = [{ type }];
        const run = Run.createFromContent(content);
        const xml = run.toXML();
        const children = filterXMLElements(xml.children);
        expect(children.some((c) => c.name === xmlName)).toBe(true);
      }
    });
  });

  describe('fieldChar (via createFromContent)', () => {
    test('should generate w:fldChar begin/separate/end', () => {
      const types: Array<'begin' | 'separate' | 'end'> = ['begin', 'separate', 'end'];

      for (const fldCharType of types) {
        const content: RunContent[] = [{ type: 'fieldChar', fieldCharType: fldCharType }];
        const run = Run.createFromContent(content);
        const xml = run.toXML();
        const children = filterXMLElements(xml.children);
        const fldChar = children.find((c) => c.name === 'w:fldChar');
        expect(fldChar).toBeDefined();
        expect(fldChar?.attributes?.['w:fldCharType']).toBe(fldCharType);
      }
    });
  });

  describe('Multiple content types combined', () => {
    test('should round-trip run with mixed content types', async () => {
      const doc = Document.create();
      const para = new Paragraph();

      const run = new Run('Before');
      run.addTab().appendText('Middle').addBreak().appendText('After');
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const content = loaded.getParagraphs()[0]?.getRuns()[0]?.getContent();

      expect(content).toBeDefined();
      expect(content!.some((c) => c.type === 'text' && c.value === 'Before')).toBe(true);
      expect(content!.some((c) => c.type === 'tab')).toBe(true);
      expect(content!.some((c) => c.type === 'text' && c.value === 'Middle')).toBe(true);
      expect(content!.some((c) => c.type === 'break')).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });
});
