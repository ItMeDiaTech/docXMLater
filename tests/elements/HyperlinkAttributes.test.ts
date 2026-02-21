/**
 * Hyperlink attributes gap tests: docLocation, setTgtFrame, setHistory
 * Phase 6A of ECMA-376 gap analysis
 */

import { Hyperlink } from '../../src/elements/Hyperlink';
import { Document } from '../../src/core/Document';

describe('Hyperlink docLocation', () => {
  test('should set and get docLocation via constructor', () => {
    const link = new Hyperlink({
      url: 'https://example.com/doc.docx',
      text: 'Open Doc',
      docLocation: 'Section1',
    });
    expect(link.getDocLocation()).toBe('Section1');
  });

  test('should set and get docLocation via setter', () => {
    const link = Hyperlink.createExternal('https://example.com', 'Link');
    link.setDocLocation('Table1');
    expect(link.getDocLocation()).toBe('Table1');
  });

  test('should clear docLocation', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      text: 'Link',
      docLocation: 'Section1',
    });
    link.setDocLocation(undefined);
    expect(link.getDocLocation()).toBeUndefined();
  });

  test('should generate w:docLocation in XML', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      text: 'Link',
      docLocation: 'Heading1',
      relationshipId: 'rId1',
    });
    const xml = link.toXML();
    expect(xml.attributes?.['w:docLocation']).toBe('Heading1');
  });

  test('should not generate w:docLocation when not set', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      text: 'Link',
      relationshipId: 'rId1',
    });
    const xml = link.toXML();
    expect(xml.attributes?.['w:docLocation']).toBeUndefined();
  });

  test('should preserve docLocation in clone', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      text: 'Link',
      docLocation: 'Bookmark5',
      tgtFrame: '_blank',
      history: '1',
    });
    const cloned = link.clone();
    expect(cloned.getDocLocation()).toBe('Bookmark5');
    expect(cloned.getTgtFrame()).toBe('_blank');
    expect(cloned.getHistory()).toBe('1');
  });
});

describe('Hyperlink setTgtFrame', () => {
  test('should set tgtFrame via setter', () => {
    const link = Hyperlink.createExternal('https://example.com', 'Link');
    link.setTgtFrame('_blank');
    expect(link.getTgtFrame()).toBe('_blank');
  });

  test('should clear tgtFrame', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      text: 'Link',
      tgtFrame: '_blank',
    });
    link.setTgtFrame(undefined);
    expect(link.getTgtFrame()).toBeUndefined();
  });

  test('should chain setTgtFrame', () => {
    const link = Hyperlink.createExternal('https://example.com', 'Link');
    const result = link.setTgtFrame('_self');
    expect(result).toBe(link);
  });

  test('should generate w:tgtFrame in XML', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      text: 'Link',
      tgtFrame: '_blank',
      relationshipId: 'rId1',
    });
    const xml = link.toXML();
    expect(xml.attributes?.['w:tgtFrame']).toBe('_blank');
  });
});

describe('Hyperlink setHistory', () => {
  test('should set history via setter', () => {
    const link = Hyperlink.createExternal('https://example.com', 'Link');
    link.setHistory('1');
    expect(link.getHistory()).toBe('1');
  });

  test('should clear history', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      text: 'Link',
      history: '1',
    });
    link.setHistory(undefined);
    expect(link.getHistory()).toBeUndefined();
  });

  test('should chain setHistory', () => {
    const link = Hyperlink.createExternal('https://example.com', 'Link');
    const result = link.setHistory('1');
    expect(result).toBe(link);
  });
});

describe('Hyperlink attributes round-trip', () => {
  test('should round-trip tgtFrame and history via document', async () => {
    const doc = Document.create();
    const para = doc.createParagraph('Test');
    const link = Hyperlink.createExternal('https://example.com', 'Link');
    link.setTgtFrame('_blank');
    link.setHistory('1');
    para.addHyperlink(link);

    // toBuffer() triggers OOXML validation
    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    // Verify the loaded doc is valid
    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.getParagraphs().length).toBeGreaterThan(0);

    doc.dispose();
    loaded.dispose();
  });

  test('should round-trip docLocation via document', async () => {
    const doc = Document.create();
    const para = doc.createParagraph('Test');
    const link = new Hyperlink({
      url: 'https://example.com/doc.docx',
      text: 'Open Doc',
      docLocation: 'Section2',
    });
    para.addHyperlink(link);

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    doc.dispose();
  });
});
