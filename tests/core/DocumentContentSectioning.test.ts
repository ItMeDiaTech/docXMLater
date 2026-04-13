/**
 * Tests for Document.extractByHeading() and Document.getElementsBetween()
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';

describe('Document.extractByHeading()', () => {
  describe('basic splitting by H1', () => {
    it('splits document into chapters at H1 headings', () => {
      const doc = Document.create();
      doc.addHeading('Chapter 1', 1);
      doc.createParagraph('Chapter 1 content.');
      doc.addHeading('Chapter 2', 1);
      doc.createParagraph('Chapter 2 content.');
      doc.createParagraph('More chapter 2.');

      const sections = doc.extractByHeading(1);

      expect(sections).toHaveLength(2);
      expect(sections[0]!.heading!.getText()).toBe('Chapter 1');
      expect(sections[0]!.level).toBe(1);
      expect(sections[0]!.content).toHaveLength(1);
      expect(sections[1]!.heading!.getText()).toBe('Chapter 2');
      expect(sections[1]!.content).toHaveLength(2);
      doc.dispose();
    });

    it('puts preamble content in a section with heading: undefined', () => {
      const doc = Document.create();
      doc.createParagraph('Preamble text.');
      doc.createParagraph('More preamble.');
      doc.addHeading('Chapter 1', 1);
      doc.createParagraph('Chapter content.');

      const sections = doc.extractByHeading(1);

      expect(sections).toHaveLength(2);
      expect(sections[0]!.heading).toBeUndefined();
      expect(sections[0]!.level).toBe(0);
      expect(sections[0]!.content).toHaveLength(2);
      expect(sections[1]!.heading!.getText()).toBe('Chapter 1');
      doc.dispose();
    });

    it('handles document with no headings', () => {
      const doc = Document.create();
      doc.createParagraph('Just text.');
      doc.createParagraph('More text.');

      const sections = doc.extractByHeading(1);

      expect(sections).toHaveLength(1);
      expect(sections[0]!.heading).toBeUndefined();
      expect(sections[0]!.content).toHaveLength(2);
      doc.dispose();
    });

    it('handles empty document', () => {
      const doc = Document.create();
      const sections = doc.extractByHeading(1);

      expect(sections).toHaveLength(0);
      doc.dispose();
    });

    it('handles consecutive headings with no content between', () => {
      const doc = Document.create();
      doc.addHeading('H1 First', 1);
      doc.addHeading('H1 Second', 1);
      doc.addHeading('H1 Third', 1);
      doc.createParagraph('Only content.');

      const sections = doc.extractByHeading(1);

      expect(sections).toHaveLength(3);
      expect(sections[0]!.content).toHaveLength(0);
      expect(sections[1]!.content).toHaveLength(0);
      expect(sections[2]!.content).toHaveLength(1);
      doc.dispose();
    });
  });

  describe('splitting by multiple heading levels', () => {
    it('splits at H1 and H2 when maxLevel=2', () => {
      const doc = Document.create();
      doc.addHeading('Chapter 1', 1);
      doc.createParagraph('Intro to chapter 1.');
      doc.addHeading('Section 1.1', 2);
      doc.createParagraph('Section content.');
      doc.addHeading('Section 1.2', 2);
      doc.createParagraph('More section content.');
      doc.addHeading('Chapter 2', 1);
      doc.createParagraph('Chapter 2 content.');

      const sections = doc.extractByHeading(2);

      expect(sections).toHaveLength(4);
      expect(sections[0]!.heading!.getText()).toBe('Chapter 1');
      expect(sections[0]!.level).toBe(1);
      expect(sections[1]!.heading!.getText()).toBe('Section 1.1');
      expect(sections[1]!.level).toBe(2);
      expect(sections[2]!.heading!.getText()).toBe('Section 1.2');
      expect(sections[2]!.level).toBe(2);
      expect(sections[3]!.heading!.getText()).toBe('Chapter 2');
      expect(sections[3]!.level).toBe(1);
      doc.dispose();
    });

    it('ignores lower-level headings when maxLevel is restrictive', () => {
      const doc = Document.create();
      doc.addHeading('Chapter', 1);
      doc.addHeading('Section', 2);
      doc.addHeading('Subsection', 3);
      doc.createParagraph('Content.');

      // Only split at H1
      const sections = doc.extractByHeading(1);

      expect(sections).toHaveLength(1);
      expect(sections[0]!.heading!.getText()).toBe('Chapter');
      // H2, H3, and paragraph are all content under this H1
      expect(sections[0]!.content).toHaveLength(3);
      doc.dispose();
    });

    it('splits at H1 through H3 when maxLevel=3', () => {
      const doc = Document.create();
      doc.addHeading('H1', 1);
      doc.addHeading('H2', 2);
      doc.addHeading('H3', 3);
      doc.createParagraph('Text under H3.');
      doc.addHeading('H4', 4); // not a split point
      doc.createParagraph('Text under H4.');

      const sections = doc.extractByHeading(3);

      expect(sections).toHaveLength(3);
      expect(sections[2]!.heading!.getText()).toBe('H3');
      // H4 heading + its text are content under H3
      expect(sections[2]!.content).toHaveLength(3);
      doc.dispose();
    });
  });

  describe('tables and mixed content', () => {
    it('includes tables in section content', () => {
      const doc = Document.create();
      doc.addHeading('Data Section', 1);
      doc.createParagraph('Intro.');
      doc.addTable(
        Table.fromArray([
          ['A', 'B'],
          ['1', '2'],
        ])
      );
      doc.createParagraph('Conclusion.');

      const sections = doc.extractByHeading(1);

      expect(sections).toHaveLength(1);
      expect(sections[0]!.content).toHaveLength(3); // para + table + para
      const hasTable = sections[0]!.content.some((e) => e instanceof Table);
      expect(hasTable).toBe(true);
      doc.dispose();
    });
  });

  describe('default maxLevel', () => {
    it('defaults to maxLevel=1 (only H1 splits)', () => {
      const doc = Document.create();
      doc.addHeading('Part 1', 1);
      doc.addHeading('Sub', 2);
      doc.addHeading('Part 2', 1);

      const sections = doc.extractByHeading();

      expect(sections).toHaveLength(2);
      expect(sections[0]!.heading!.getText()).toBe('Part 1');
      expect(sections[0]!.content).toHaveLength(1); // The H2
      expect(sections[1]!.heading!.getText()).toBe('Part 2');
      doc.dispose();
    });
  });
});

describe('Document.getElementsBetween()', () => {
  it('returns elements between two paragraphs', () => {
    const doc = Document.create();
    const start = doc.createParagraph('Start');
    doc.createParagraph('Middle 1');
    doc.createParagraph('Middle 2');
    const end = doc.createParagraph('End');

    const between = doc.getElementsBetween(start, end);

    expect(between).toHaveLength(2);
    expect((between[0] as Paragraph).getText()).toBe('Middle 1');
    expect((between[1] as Paragraph).getText()).toBe('Middle 2');
    doc.dispose();
  });

  it('returns elements between a paragraph and a table', () => {
    const doc = Document.create();
    const heading = doc.addHeading('Title', 1);
    doc.createParagraph('Content 1');
    doc.createParagraph('Content 2');
    const table = doc.createTable(1, 1);

    const between = doc.getElementsBetween(heading, table);

    expect(between).toHaveLength(2);
    doc.dispose();
  });

  it('returns empty array for adjacent elements', () => {
    const doc = Document.create();
    const first = doc.createParagraph('First');
    const second = doc.createParagraph('Second');

    const between = doc.getElementsBetween(first, second);

    expect(between).toHaveLength(0);
    doc.dispose();
  });

  it('returns empty array when start is after end', () => {
    const doc = Document.create();
    const first = doc.createParagraph('First');
    const second = doc.createParagraph('Second');

    const between = doc.getElementsBetween(second, first);

    expect(between).toEqual([]);
    doc.dispose();
  });

  it('returns empty array when element not found', () => {
    const doc = Document.create();
    doc.createParagraph('Existing');

    const orphan = new Paragraph().addText('Not in doc');
    const existing = doc.getParagraphs()[0]!;

    expect(doc.getElementsBetween(orphan, existing)).toEqual([]);
    expect(doc.getElementsBetween(existing, orphan)).toEqual([]);
    doc.dispose();
  });

  it('returns empty array when same element passed twice', () => {
    const doc = Document.create();
    const para = doc.createParagraph('Solo');

    expect(doc.getElementsBetween(para, para)).toEqual([]);
    doc.dispose();
  });

  it('works with heading-based chapter extraction', () => {
    const doc = Document.create();
    const h1 = doc.addHeading('Chapter 1', 1);
    doc.createParagraph('Para A');
    doc.addTable(Table.fromArray([['T1']]));
    doc.createParagraph('Para B');
    const h2 = doc.addHeading('Chapter 2', 1);
    doc.createParagraph('Para C');

    const chapter1Content = doc.getElementsBetween(h1, h2);

    expect(chapter1Content).toHaveLength(3);
    expect((chapter1Content[0] as Paragraph).getText()).toBe('Para A');
    expect(chapter1Content[1] instanceof Table).toBe(true);
    expect((chapter1Content[2] as Paragraph).getText()).toBe('Para B');
    doc.dispose();
  });
});

describe('extractByHeading + getElementsBetween integration', () => {
  it('can extract and process individual chapters', () => {
    const doc = Document.create();
    doc.addHeading('Introduction', 1);
    doc.createParagraph('Intro content.');
    doc.addHeading('Methods', 1);
    doc.createParagraph('Methods content.');
    doc.addTable(Table.fromArray([['Data', 'Value']]));
    doc.addHeading('Results', 1);
    doc.createParagraph('Results content.');

    const chapters = doc.extractByHeading(1);

    // Find the Methods chapter
    const methods = chapters.find((s) => s.heading?.getText() === 'Methods');
    expect(methods).toBeDefined();
    expect(methods!.content).toHaveLength(2); // paragraph + table

    // Count total content across all chapters
    const totalContent = chapters.reduce((sum, s) => sum + s.content.length, 0);
    expect(totalContent).toBe(4); // 1 + 2 + 1
    doc.dispose();
  });
});
