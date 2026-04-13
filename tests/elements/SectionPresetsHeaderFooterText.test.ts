/**
 * Tests for Section page size presets, Header.addText(), Footer.addText()
 */

import { Section } from '../../src/elements/Section';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { PAGE_SIZES } from '../../src/utils/units';

// ============================================================================
// Section Page Size Presets
// ============================================================================

describe('Section page size presets', () => {
  describe('Section.createLegal()', () => {
    it('creates a legal-sized section', () => {
      const section = Section.createLegal();
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(PAGE_SIZES.LEGAL.width);
      expect(props.pageSize?.height).toBe(PAGE_SIZES.LEGAL.height);
      expect(props.pageSize?.orientation).toBe('portrait');
    });

    it('has correct dimensions (8.5" x 14")', () => {
      const section = Section.createLegal();
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(12240); // 8.5 inches
      expect(props.pageSize?.height).toBe(20160); // 14 inches
    });
  });

  describe('Section.createTabloid()', () => {
    it('creates a tabloid-sized section', () => {
      const section = Section.createTabloid();
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(PAGE_SIZES.TABLOID.width);
      expect(props.pageSize?.height).toBe(PAGE_SIZES.TABLOID.height);
      expect(props.pageSize?.orientation).toBe('portrait');
    });

    it('has correct dimensions (11" x 17")', () => {
      const section = Section.createTabloid();
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(15840); // 11 inches
      expect(props.pageSize?.height).toBe(24480); // 17 inches
    });
  });

  describe('Section.createA3()', () => {
    it('creates an A3-sized section', () => {
      const section = Section.createA3();
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(PAGE_SIZES.A3.width);
      expect(props.pageSize?.height).toBe(PAGE_SIZES.A3.height);
      expect(props.pageSize?.orientation).toBe('portrait');
    });

    it('has correct dimensions (29.7cm x 42cm)', () => {
      const section = Section.createA3();
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(16838);
      expect(props.pageSize?.height).toBe(23811);
    });
  });

  describe('Section.createLandscape() extended', () => {
    it('creates landscape letter (default)', () => {
      const section = Section.createLandscape();
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(PAGE_SIZES.LETTER.height);
      expect(props.pageSize?.height).toBe(PAGE_SIZES.LETTER.width);
      expect(props.pageSize?.orientation).toBe('landscape');
    });

    it('creates landscape A4', () => {
      const section = Section.createLandscape('a4');
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(PAGE_SIZES.A4.height);
      expect(props.pageSize?.height).toBe(PAGE_SIZES.A4.width);
    });

    it('creates landscape legal', () => {
      const section = Section.createLandscape('legal');
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(PAGE_SIZES.LEGAL.height);
      expect(props.pageSize?.height).toBe(PAGE_SIZES.LEGAL.width);
      expect(props.pageSize?.orientation).toBe('landscape');
    });

    it('creates landscape tabloid', () => {
      const section = Section.createLandscape('tabloid');
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(PAGE_SIZES.TABLOID.height);
      expect(props.pageSize?.height).toBe(PAGE_SIZES.TABLOID.width);
    });

    it('creates landscape A3', () => {
      const section = Section.createLandscape('a3');
      const props = section.getProperties();

      expect(props.pageSize?.width).toBe(PAGE_SIZES.A3.height);
      expect(props.pageSize?.height).toBe(PAGE_SIZES.A3.width);
    });
  });

  describe('all presets produce valid sections', () => {
    it.each([
      ['Letter', Section.createLetter()],
      ['A4', Section.createA4()],
      ['Legal', Section.createLegal()],
      ['Tabloid', Section.createTabloid()],
      ['A3', Section.createA3()],
    ])('%s section has valid properties', (_name, section) => {
      const props = section.getProperties();
      expect(props.pageSize?.width).toBeGreaterThan(0);
      expect(props.pageSize?.height).toBeGreaterThan(0);
      expect(props.pageSize?.height).toBeGreaterThan(props.pageSize!.width!);
    });
  });
});

// ============================================================================
// Header.addText()
// ============================================================================

describe('Header.addText()', () => {
  it('adds plain text to header', () => {
    const header = new Header();
    header.addText('Company Name');

    const elements = header.getElements();
    expect(elements).toHaveLength(1);
    expect((elements[0] as Paragraph).getText()).toBe('Company Name');
  });

  it('adds formatted text', () => {
    const header = new Header();
    header.addText('Bold Header', { bold: true, font: 'Arial', size: 10 });

    const para = header.getElements()[0] as Paragraph;
    const runs = para.getRuns();
    expect(runs[0]!.getText()).toBe('Bold Header');
    expect(runs[0]!.getFormatting().bold).toBe(true);
    expect(runs[0]!.getFormatting().font).toBe('Arial');
  });

  it('returns paragraph for further customization', () => {
    const header = new Header();
    const para = header.addText('Right-aligned');
    para.setAlignment('right');

    expect(para.getAlignment()).toBe('right');
  });

  it('adds multiple text paragraphs', () => {
    const header = new Header();
    header.addText('Line 1');
    header.addText('Line 2');

    expect(header.getElementCount()).toBe(2);
  });

  it('works in a Document context', () => {
    const doc = Document.create();
    const header = new Header();
    header.addText('Document Title', { bold: true });
    doc.setHeader(header);

    // Header was set successfully
    expect(header.getElements()).toHaveLength(1);
    doc.dispose();
  });
});

// ============================================================================
// Footer.addText()
// ============================================================================

describe('Footer.addText()', () => {
  it('adds plain text to footer', () => {
    const footer = new Footer();
    footer.addText('Page 1');

    const elements = footer.getElements();
    expect(elements).toHaveLength(1);
    expect((elements[0] as Paragraph).getText()).toBe('Page 1');
  });

  it('adds formatted text', () => {
    const footer = new Footer();
    footer.addText('Confidential', { italic: true, color: '888888', size: 8 });

    const para = footer.getElements()[0] as Paragraph;
    const fmt = para.getRuns()[0]!.getFormatting();
    expect(fmt.italic).toBe(true);
    expect(fmt.color).toBe('888888');
    expect(fmt.size).toBe(8);
  });

  it('returns paragraph for alignment', () => {
    const footer = new Footer();
    const para = footer.addText('Centered footer');
    para.setAlignment('center');

    expect(para.getAlignment()).toBe('center');
  });

  it('works in a Document context', () => {
    const doc = Document.create();
    const footer = new Footer();
    footer.addText('Copyright 2024', { size: 8 });
    doc.setFooter(footer);

    expect(footer.getElements()).toHaveLength(1);
    doc.dispose();
  });
});

// ============================================================================
// Integration
// ============================================================================

describe('integration: sections + headers + footers', () => {
  it('creates a legal document with header and footer', async () => {
    const doc = Document.create();

    // Set legal page size
    doc.setSection(Section.createLegal());

    // Add header
    const header = new Header();
    header.addText('LEGAL DOCUMENT', { bold: true, font: 'Times New Roman', size: 10 });
    doc.setHeader(header);

    // Add footer
    const footer = new Footer();
    footer.addText('Confidential', { italic: true, size: 8 });
    doc.setFooter(footer);

    // Content
    doc.addHeading('Contract', 1);
    doc.createParagraph('Terms and conditions apply.');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    doc.dispose();
  });

  it('creates a landscape A3 poster document', () => {
    const doc = Document.create();
    doc.setSection(Section.createLandscape('a3'));

    const section = doc.getSection();
    const props = section.getProperties();

    // Landscape A3: width > height
    expect(props.pageSize?.width).toBeGreaterThan(props.pageSize!.height!);
    doc.dispose();
  });
});
