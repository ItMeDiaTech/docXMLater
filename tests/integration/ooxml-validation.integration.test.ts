/**
 * OOXML Schema Compliance Integration Tests
 *
 * Systematically exercises XML element ordering and schema compliance paths
 * most prone to regression. Each test creates a minimal Document, configures
 * an element to exercise all properties that map to XML children, calls
 * toBuffer() (which triggers OOXML validation via setup.ts monkey-patch),
 * and disposes.
 *
 * No assertions on XML structure — the validator handles correctness.
 * Tests verify the document generates without validation errors.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';
import { TableCell } from '../../src/elements/TableCell';
import { Image } from '../../src/elements/Image';
import { ImageRun } from '../../src/elements/ImageRun';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Bookmark } from '../../src/elements/Bookmark';
import { Comment } from '../../src/elements/Comment';
import { Revision } from '../../src/elements/Revision';
import { StructuredDocumentTag, ListItem } from '../../src/elements/StructuredDocumentTag';
import { Style } from '../../src/formatting/Style';
import { AbstractNumbering } from '../../src/formatting/AbstractNumbering';
import { NumberingInstance } from '../../src/formatting/NumberingInstance';
import { NumberingLevel } from '../../src/formatting/NumberingLevel';

/** 1x1 transparent PNG for image tests */
function createTestPng(): Buffer {
  return Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41,
    0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00,
    0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);
}

describe('OOXML Validation Integration', () => {

  // =========================================================================
  // 1. Table with full tblPr properties
  // =========================================================================
  describe('Table - full tblPr properties', () => {
    it('should produce valid OOXML with all table properties set', async () => {
      const doc = Document.create();

      const table = new Table(3, 3);
      table.setStyle('TableGrid');
      table.setWidth(5000);
      table.setWidthType('dxa');
      table.setAlignment('center');
      table.setCellSpacing(20);
      table.setIndent(720);
      table.setBorders({
        top: { style: 'single', size: 4, color: '000000' },
        bottom: { style: 'single', size: 4, color: '000000' },
        left: { style: 'single', size: 4, color: '000000' },
        right: { style: 'single', size: 4, color: '000000' },
        insideH: { style: 'single', size: 4, color: '000000' },
        insideV: { style: 'single', size: 4, color: '000000' },
      });
      table.setShading({ fill: 'F0F0F0', pattern: 'clear' });
      table.setLayout('fixed');
      table.setCellMargins({ top: 50, bottom: 50, left: 108, right: 108 });
      table.setTblLook('04A0');
      // tblStyleRowBandSize/tblStyleColBandSize are stored but not serialized
      // in direct tblPr (only valid in table style definitions per ECMA-376)
      table.setStyleRowBandSize(1);
      table.setStyleColBandSize(1);
      table.setCaption('Test Table');
      table.setDescription('A table exercising all tblPr children');

      // Populate cells
      for (let r = 0; r < 3; r++) {
        for (let c = 0; c < 3; c++) {
          table.getCell(r, c)!.createParagraph(`Cell ${r},${c}`);
        }
      }

      doc.addTable(table);
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 2. Table with tblPrEx (row-level property exceptions)
  // =========================================================================
  describe('Table - tblPrEx row-level exceptions', () => {
    it('should produce valid OOXML with table property exceptions on rows', async () => {
      const doc = Document.create();

      const table = new Table(3, 3);
      table.setWidth(5000);
      table.setBorders({
        top: { style: 'single', size: 4, color: '000000' },
        bottom: { style: 'single', size: 4, color: '000000' },
        left: { style: 'single', size: 4, color: '000000' },
        right: { style: 'single', size: 4, color: '000000' },
        insideH: { style: 'single', size: 4, color: '000000' },
        insideV: { style: 'single', size: 4, color: '000000' },
      });

      // Row 0: override borders + shading + spacing + width + indent + justification
      table.getRow(0)!.setTablePropertyExceptions({
        borders: {
          top: { style: 'double', size: 8, color: 'FF0000' },
          bottom: { style: 'double', size: 8, color: 'FF0000' },
          left: { style: 'double', size: 8, color: 'FF0000' },
          right: { style: 'double', size: 8, color: 'FF0000' },
        },
        shading: { fill: 'FFFF00', pattern: 'clear' },
        cellSpacing: 40,
        width: 6000,
        indentation: 360,
        justification: 'center',
      });

      // Row 1: different exceptions
      table.getRow(1)!.setTablePropertyExceptions({
        shading: { fill: 'E0E0E0', pattern: 'clear' },
        justification: 'right',
      });

      for (let r = 0; r < 3; r++) {
        for (let c = 0; c < 3; c++) {
          table.getCell(r, c)!.createParagraph(`Cell ${r},${c}`);
        }
      }

      doc.addTable(table);
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 3. Table with cell merging
  // =========================================================================
  describe('Table - cell merging (gridSpan + vMerge)', () => {
    it('should produce valid OOXML with horizontal and vertical merges', async () => {
      const doc = Document.create();

      const table = new Table(4, 4);

      // Horizontal merge: row 0, cols 0-1
      table.getCell(0, 0)!.setColumnSpan(2);
      table.getCell(0, 0)!.createParagraph('Merged horizontally');

      // Vertical merge: col 2, rows 0-2
      table.getCell(0, 2)!.setVerticalMerge('restart');
      table.getCell(0, 2)!.createParagraph('Merged vertically');
      table.getCell(1, 2)!.setVerticalMerge('continue');
      table.getCell(2, 2)!.setVerticalMerge('continue');

      // Cell properties: borders, shading, vertical alignment, text direction, margins
      table.getCell(0, 3)!.setBorders({
        top: { style: 'thick', size: 12, color: '0000FF' },
        bottom: { style: 'thick', size: 12, color: '0000FF' },
      });
      table.getCell(0, 3)!.setShading({ fill: 'CCFFCC', pattern: 'clear' });
      table.getCell(0, 3)!.setVerticalAlignment('center');
      table.getCell(0, 3)!.setTextDirection('btLr');
      table.getCell(0, 3)!.setMargins({ top: 100, bottom: 100, left: 200, right: 200 });
      table.getCell(0, 3)!.setNoWrap(true);
      table.getCell(0, 3)!.createParagraph('Formatted cell');

      // Fill remaining
      for (let r = 1; r < 4; r++) {
        for (let c = 0; c < 4; c++) {
          if (table.getCell(r, c)!.getParagraphs().length === 0) {
            table.getCell(r, c)!.createParagraph(`R${r}C${c}`);
          }
        }
      }

      doc.addTable(table);
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 4. Paragraph with full pPr properties
  // =========================================================================
  describe('Paragraph - full pPr properties', () => {
    it('should produce valid OOXML with all paragraph properties set', async () => {
      const doc = Document.create();

      // Set up numbering for the paragraph
      const abstract = AbstractNumbering.createNumberedList(1);
      const instance = NumberingInstance.create(1, 1);
      doc.getNumberingManager().addAbstractNumbering(abstract);
      doc.getNumberingManager().addInstance(instance);

      const para = new Paragraph();
      para.setKeepNext(true);
      para.setKeepLines(true);
      para.setPageBreakBefore(false);
      para.setSpaceBefore(240);
      para.setSpaceAfter(120);
      para.setLineSpacing(276, 'auto');
      para.setLeftIndent(720);
      para.setRightIndent(360);
      para.setFirstLineIndent(360);
      para.setContextualSpacing(true);
      para.setAlignment('both');
      para.setOutlineLevel(1);
      para.setNumbering(1, 0);
      para.setBorder({
        top: { style: 'single', size: 4, color: '000000' },
        bottom: { style: 'single', size: 4, color: '000000' },
        left: { style: 'single', size: 4, color: '000000' },
        right: { style: 'single', size: 4, color: '000000' },
      });
      para.setShading({ fill: 'FFFFCC', pattern: 'clear' });
      para.setTabs([
        { position: 1440, val: 'left', leader: 'none' },
        { position: 4320, val: 'center', leader: 'dot' },
        { position: 7200, val: 'right', leader: 'hyphen' },
      ]);
      para.setWidowControl(true);
      para.setBidi(false);

      para.addText('Paragraph with all properties', { bold: true });

      doc.addParagraph(para);
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 5. Paragraph with CJK properties
  // =========================================================================
  describe('Paragraph - CJK typography properties', () => {
    it('should produce valid OOXML with CJK properties set', async () => {
      const doc = Document.create();

      const para = new Paragraph();
      para.setKinsoku(true);
      para.setWordWrap(true);
      para.setOverflowPunct(true);
      para.setTopLinePunct(true);
      para.setAutoSpaceDE(true);
      para.setAutoSpaceDN(true);
      para.setAlignment('both');
      para.addText('CJK typography test paragraph');

      const para2 = new Paragraph();
      para2.setKinsoku(false);
      para2.setWordWrap(false);
      para2.setOverflowPunct(false);
      para2.setTopLinePunct(false);
      para2.setAutoSpaceDE(false);
      para2.setAutoSpaceDN(false);
      para2.addText('CJK properties disabled');

      doc.addParagraph(para);
      doc.addParagraph(para2);
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 6. Run with full rPr properties (via character style)
  // =========================================================================
  describe('Run - full rPr properties via character style', () => {
    it('should produce valid OOXML with all run properties set', async () => {
      const doc = Document.create();

      // Create a character style with run formatting
      const charStyle = new Style({
        styleId: 'CustomChar',
        name: 'Custom Character',
        type: 'character',
        runFormatting: {
          bold: true,
          italic: true,
          color: 'FF0000',
          size: 28,
        },
      });
      doc.getStylesManager().addStyle(charStyle);

      const para = new Paragraph();

      // Run with all direct formatting
      const run = new Run('Fully formatted run');
      run.setFont('Arial');
      run.setFontEastAsia('MS Mincho');
      run.setFontCs('Arial');
      run.setBold(true);
      run.setItalic(true);
      run.setAllCaps(false);
      run.setSmallCaps(true);
      run.setStrike(false);
      run.setColor('0000FF');
      run.setSize(24);
      run.setHighlight('yellow');
      run.setUnderline('single');
      run.setSubscript(false);
      run.setSuperscript(false);
      run.setCharacterSpacing(20);
      run.setKerning(24);
      run.setLanguage('en-US');
      run.setNoProof(true);
      run.setCharacterStyle('CustomChar');
      run.setShading({ fill: 'FFFF00' });

      para.addRun(run);

      // Second run with different properties
      const run2 = new Run('Another run');
      run2.setFont('Times New Roman');
      run2.setBold(false);
      run2.setItalic(false);
      run2.setAllCaps(true);
      run2.setStrike(true);
      run2.setColor('00FF00');
      run2.setSize(18);
      run2.setHighlight('green');
      run2.setUnderline('double');
      run2.setSuperscript(true);

      para.addRun(run2);

      doc.addParagraph(para);
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 7. Style with full metadata
  // =========================================================================
  describe('Style - full metadata ordering', () => {
    it('should produce valid OOXML with all style metadata set', async () => {
      const doc = Document.create();

      // Paragraph style with full metadata
      const paraStyle = new Style({
        styleId: 'CustomParagraph',
        name: 'Custom Paragraph',
        type: 'paragraph',
        paragraphFormatting: {
          alignment: 'center',
          spacing: { before: 240, after: 120 },
        },
        runFormatting: {
          bold: true,
          color: '333333',
          size: 24,
        },
      });
      paraStyle.setBasedOn('Normal');
      paraStyle.setNext('Normal');
      paraStyle.setLink('CustomParagraphChar');
      paraStyle.setUiPriority(10);
      paraStyle.setSemiHidden(false);
      paraStyle.setUnhideWhenUsed(false);
      paraStyle.setQFormat(true);
      paraStyle.setLocked(false);
      paraStyle.setPersonal(false);
      paraStyle.setAliases('MyPara,TestPara');

      // Linked character style
      const charStyle = new Style({
        styleId: 'CustomParagraphChar',
        name: 'Custom Paragraph Char',
        type: 'character',
        runFormatting: {
          bold: true,
          color: '333333',
          size: 24,
        },
      });
      charStyle.setLink('CustomParagraph');
      charStyle.setUiPriority(10);
      charStyle.setQFormat(false);
      charStyle.setLocked(true);

      doc.getStylesManager().addStyle(paraStyle);
      doc.getStylesManager().addStyle(charStyle);

      // Use the style
      const para = new Paragraph();
      para.setStyle('CustomParagraph');
      para.addText('Text with custom style');
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 8. Revision types (insert, delete, move, property change)
  // =========================================================================
  describe('Revisions - insert, delete, and property change', () => {
    it('should produce valid OOXML with multiple revision types', async () => {
      const doc = Document.create();

      const para1 = new Paragraph();

      // Insertion revision
      const insertRev = new Revision({
        type: 'insert',
        author: 'TestAuthor',
        date: new Date('2025-01-15T10:00:00Z'),
        content: [new Run('Inserted text')],
      });
      para1.addRevision(insertRev);

      // Normal text between revisions
      para1.addText(' normal text ');

      // Deletion revision
      const deleteRev = new Revision({
        type: 'delete',
        author: 'TestAuthor',
        date: new Date('2025-01-15T11:00:00Z'),
        content: [new Run('Deleted text')],
      });
      para1.addRevision(deleteRev);

      // Property change revision — silently skipped in paragraph serialization
      // (rPrChange is only valid inside w:rPr, not as direct child of w:p)
      const rPrChangeRev = new Revision({
        type: 'runPropertiesChange',
        author: 'TestAuthor',
        date: new Date('2025-01-15T12:00:00Z'),
        content: [new Run('formatted text')],
      });
      para1.addRevision(rPrChangeRev);

      doc.addParagraph(para1);

      // Second paragraph with interleaved insert/delete revisions
      const para2 = new Paragraph();
      para2.addText('Before ');
      para2.addRevision(new Revision({
        type: 'insert',
        author: 'Author2',
        date: new Date('2025-01-15T14:00:00Z'),
        content: [new Run('newly added')],
      }));
      para2.addText(' middle ');
      para2.addRevision(new Revision({
        type: 'delete',
        author: 'Author2',
        date: new Date('2025-01-15T15:00:00Z'),
        content: [new Run('removed')],
      }));
      para2.addText(' after');
      doc.addParagraph(para2);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 9. Hyperlinks in revisions
  // =========================================================================
  describe('Hyperlinks in revisions', () => {
    it('should produce valid OOXML with hyperlinks inside revision wrappers', async () => {
      const doc = Document.create();

      const para = new Paragraph();

      // Hyperlink inside an insertion
      const hyperlink = Hyperlink.createExternal('https://example.com', 'Example Link');
      const insertRev = new Revision({
        type: 'insert',
        author: 'TestAuthor',
        date: new Date('2025-01-15T10:00:00Z'),
        content: [],
      });
      insertRev.addHyperlink(hyperlink);
      para.addRevision(insertRev);

      para.addText(' text between ');

      // Hyperlink inside a deletion
      const hyperlink2 = Hyperlink.createExternal('https://old-link.com', 'Old Link');
      const deleteRev = new Revision({
        type: 'delete',
        author: 'TestAuthor',
        date: new Date('2025-01-15T11:00:00Z'),
        content: [],
      });
      deleteRev.addHyperlink(hyperlink2);
      para.addRevision(deleteRev);

      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 10. Footnotes with hyperlinks
  // =========================================================================
  describe('Footnotes with hyperlinks', () => {
    it('should produce valid OOXML with footnotes containing hyperlinks', async () => {
      const doc = Document.create();

      doc.createParagraph('Document body text');

      // Create footnote with hyperlink in its content
      const fn = doc.createFootnote('See ');
      const fnPara = fn.getParagraphs()[0]!;
      fnPara.addHyperlink(Hyperlink.createExternal('https://example.com', 'this link'));

      // Second footnote with just text
      doc.createFootnote('A plain footnote');

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 11. Endnotes with hyperlinks
  // =========================================================================
  describe('Endnotes with hyperlinks', () => {
    it('should produce valid OOXML with endnotes containing hyperlinks', async () => {
      const doc = Document.create();

      doc.createParagraph('Document body text');

      // Create endnote with hyperlink
      const en = doc.createEndnote('Reference: ');
      const enPara = en.getParagraphs()[0]!;
      enPara.addHyperlink(Hyperlink.createExternal('https://example.com', 'source'));

      // Second endnote
      doc.createEndnote('A plain endnote');

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 12. Comments with reply threading
  // =========================================================================
  describe('Comments - reply threading', () => {
    it('should produce valid OOXML with comment threads', async () => {
      const doc = Document.create();

      const para = doc.createParagraph('Text that has comments');

      // Create parent comment
      const comment = doc.createComment('Alice', 'Please review this section', 'A');
      para.addComment(comment);

      // Create reply to the comment
      const reply = doc.getCommentManager().createReply(
        comment.getId(),
        'Bob',
        'Looks good to me',
        'B'
      );
      // Mark as modified for save
      (doc as any)._commentsModified = true;

      // Create second independent comment
      const comment2 = doc.createComment('Charlie', 'Another comment', 'C');
      const para2 = doc.createParagraph('More text');
      para2.addComment(comment2);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 13. Headers and footers with content
  // =========================================================================
  describe('Headers and footers - multiple types with content', () => {
    it('should produce valid OOXML with default/first/even headers and footers', async () => {
      const doc = Document.create();

      // Default header with paragraphs, formatting, and external hyperlink
      const defaultHeader = Header.createDefault();
      const headerPara = defaultHeader.createParagraph('Default Header');
      headerPara.setAlignment('right');
      headerPara.addHyperlink(Hyperlink.createExternal('https://example.com', 'Example'));
      const headerPara2 = defaultHeader.createParagraph('Second line');
      headerPara2.setAlignment('center');
      headerPara2.addText(' with bold', { bold: true });

      // First page header
      const firstHeader = Header.createFirst();
      firstHeader.createParagraph('First Page Header').setAlignment('center');

      // Even page header
      const evenHeader = Header.createEven();
      evenHeader.createParagraph('Even Page Header').setAlignment('left');

      // Default footer
      const defaultFooter = Footer.createDefault();
      const footerPara = defaultFooter.createParagraph('Page ');
      footerPara.addPageNumber();
      footerPara.setAlignment('center');

      // First page footer
      const firstFooter = Footer.createFirst();
      firstFooter.createParagraph('Confidential').setAlignment('center');

      // Even page footer
      const evenFooter = Footer.createEven();
      evenFooter.createParagraph('Even Page Footer').setAlignment('right');

      doc.setHeader(defaultHeader);
      doc.setFirstPageHeader(firstHeader);
      doc.setEvenPageHeader(evenHeader);
      doc.setFooter(defaultFooter);
      doc.setFirstPageFooter(firstFooter);
      doc.setEvenPageFooter(evenFooter);

      doc.createParagraph('Document body text');

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 14. Settings with track changes
  // =========================================================================
  describe('Settings - track changes configuration', () => {
    it('should produce valid OOXML with track changes settings', async () => {
      const doc = Document.create();

      doc.enableTrackChanges({
        author: 'TestAuthor',
        trackFormatting: true,
        showInsertionsAndDeletions: true,
        showFormatting: true,
        showInkAnnotations: true,
      });

      doc.createParagraph('Content with track changes enabled');

      // Add tracked revision
      const para2 = new Paragraph();
      const insertRev = new Revision({
        type: 'insert',
        author: 'TestAuthor',
        date: new Date('2025-01-15T10:00:00Z'),
        content: [new Run('Tracked insertion')],
      });
      para2.addRevision(insertRev);
      doc.addParagraph(para2);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 15. Bookmarks with cross-references
  // =========================================================================
  describe('Bookmarks - start/end pairs and internal hyperlinks', () => {
    it('should produce valid OOXML with bookmarks and cross-references', async () => {
      const doc = Document.create();

      // Create bookmark
      const bookmark = doc.createBookmark('ImportantSection');
      const targetPara = doc.createParagraph('Bookmarked section content');
      targetPara.addBookmark(bookmark);

      // Second bookmark
      const bookmark2 = doc.createBookmark('AnotherSection');
      const targetPara2 = doc.createParagraph('Another bookmarked section');
      targetPara2.addBookmark(bookmark2);

      // Cross-reference via internal hyperlink
      const refPara = new Paragraph();
      refPara.addText('See ');
      refPara.addHyperlink(Hyperlink.createInternal(bookmark.getName(), 'Important Section'));
      refPara.addText(' and ');
      refPara.addHyperlink(Hyperlink.createInternal(bookmark2.getName(), 'Another Section'));
      doc.addParagraph(refPara);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 16. Images (inline and floating)
  // =========================================================================
  describe('Images - inline and floating', () => {
    it('should produce valid OOXML with inline image', async () => {
      const doc = Document.create();
      const image = await Image.fromBuffer(createTestPng(), 'png', 914400, 914400);
      doc.addImage(image);
      doc.createParagraph('Text after inline image');

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });

    it('should produce valid OOXML with image having borders and effects', async () => {
      const doc = Document.create();
      const image = await Image.fromBuffer(createTestPng(), 'png', 914400, 914400);

      image.setAltText('Test Image');
      image.setBorder({
        width: 1,
        fill: { type: 'srgbClr', value: '000000' },
      });

      doc.addImage(image);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 17. StructuredDocumentTag (content controls)
  // =========================================================================
  describe('StructuredDocumentTag - multiple control types', () => {
    it('should produce valid OOXML with rich text SDT', async () => {
      const doc = Document.create();

      const para = new Paragraph();
      para.addText('Rich text content in SDT');
      const sdt = StructuredDocumentTag.createRichText([para], {
        alias: 'RichTextControl',
        tag: 'rich-text-1',
      });
      doc.addStructuredDocumentTag(sdt);

      doc.createParagraph('Text after SDT');

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });

    it('should produce valid OOXML with date picker SDT', async () => {
      const doc = Document.create();

      const para = new Paragraph();
      para.addText('2025-01-15');
      const sdt = StructuredDocumentTag.createDatePicker('yyyy-MM-dd', [para], {
        alias: 'DatePicker',
        tag: 'date-1',
      });
      doc.addStructuredDocumentTag(sdt);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });

    it('should produce valid OOXML with dropdown SDT', async () => {
      const doc = Document.create();

      const items: ListItem[] = [
        { displayText: 'Option A', value: 'a' },
        { displayText: 'Option B', value: 'b' },
        { displayText: 'Option C', value: 'c' },
      ];
      const para = new Paragraph();
      para.addText('Option A');
      const sdt = StructuredDocumentTag.createDropDownList(items, [para], {
        alias: 'Dropdown',
        tag: 'dropdown-1',
      });
      doc.addStructuredDocumentTag(sdt);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });

    it('should produce valid OOXML with checkbox SDT', async () => {
      const doc = Document.create();

      const para = new Paragraph();
      para.addText('Checked');
      const sdt = StructuredDocumentTag.createCheckbox(true, [para], {
        alias: 'Checkbox',
        tag: 'checkbox-1',
      });
      doc.addStructuredDocumentTag(sdt);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // 18. Numbering with multiple levels
  // =========================================================================
  describe('Numbering - multi-level lists', () => {
    it('should produce valid OOXML with numbered and bullet lists', async () => {
      const doc = Document.create();

      // Create bullet list
      const bulletAbstract = AbstractNumbering.createBulletList(1, 3);
      const bulletInstance = NumberingInstance.create(1, 1);
      doc.getNumberingManager().addAbstractNumbering(bulletAbstract);
      doc.getNumberingManager().addInstance(bulletInstance);

      // Create numbered list
      const numAbstract = AbstractNumbering.createNumberedList(2, 3);
      const numInstance = NumberingInstance.create(2, 2);
      doc.getNumberingManager().addAbstractNumbering(numAbstract);
      doc.getNumberingManager().addInstance(numInstance);

      // Bullet list paragraphs
      for (let level = 0; level < 3; level++) {
        const para = new Paragraph();
        para.setNumbering(1, level);
        para.addText(`Bullet item level ${level}`);
        doc.addParagraph(para);
      }

      doc.createParagraph(''); // spacer

      // Numbered list paragraphs
      for (let level = 0; level < 3; level++) {
        const para = new Paragraph();
        para.setNumbering(2, level);
        para.addText(`Numbered item level ${level}`);
        doc.addParagraph(para);
      }

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });

  // =========================================================================
  // Combined: kitchen-sink document
  // =========================================================================
  describe('Combined - kitchen-sink document with all element types', () => {
    it('should produce valid OOXML with multiple element types in one document', async () => {
      const doc = Document.create();

      // Style
      const headingStyle = new Style({
        styleId: 'KitchenSinkHeading',
        name: 'Kitchen Sink Heading',
        type: 'paragraph',
        paragraphFormatting: { alignment: 'center', spacing: { before: 240, after: 120 } },
        runFormatting: { bold: true, size: 32, color: '1F4E79' },
      });
      headingStyle.setBasedOn('Normal');
      headingStyle.setQFormat(true);
      headingStyle.setUiPriority(5);
      doc.getStylesManager().addStyle(headingStyle);

      // Numbering
      const abstract = AbstractNumbering.createBulletList(1, 3);
      const instance = NumberingInstance.create(1, 1);
      doc.getNumberingManager().addAbstractNumbering(abstract);
      doc.getNumberingManager().addInstance(instance);

      // Header + Footer
      const header = Header.createDefault();
      header.createParagraph('Document Title').setAlignment('center');
      doc.setHeader(header);

      const footer = Footer.createDefault();
      const footerPara = footer.createParagraph('Page ');
      footerPara.addPageNumber();
      footerPara.setAlignment('center');
      doc.setFooter(footer);

      // Heading paragraph
      const heading = new Paragraph();
      heading.setStyle('KitchenSinkHeading');
      heading.addText('Kitchen Sink Integration Test');
      doc.addParagraph(heading);

      // Normal paragraph with full formatting
      const normalPara = new Paragraph();
      normalPara.setSpaceBefore(120);
      normalPara.setSpaceAfter(120);
      normalPara.setAlignment('both');
      normalPara.addText('This document exercises all element types in a single buffer.', {
        italic: true,
      });
      doc.addParagraph(normalPara);

      // Table
      const table = new Table(2, 2);
      table.setWidth(5000);
      table.setBorders({
        top: { style: 'single', size: 4, color: '000000' },
        bottom: { style: 'single', size: 4, color: '000000' },
        left: { style: 'single', size: 4, color: '000000' },
        right: { style: 'single', size: 4, color: '000000' },
        insideH: { style: 'single', size: 4, color: '000000' },
        insideV: { style: 'single', size: 4, color: '000000' },
      });
      table.getCell(0, 0)!.createParagraph('Header 1');
      table.getCell(0, 1)!.createParagraph('Header 2');
      table.getCell(1, 0)!.createParagraph('Data 1');
      table.getCell(1, 1)!.createParagraph('Data 2');
      doc.addTable(table);

      // List items
      for (let i = 0; i < 2; i++) {
        const listPara = new Paragraph();
        listPara.setNumbering(1, 0);
        listPara.addText(`List item ${i + 1}`);
        doc.addParagraph(listPara);
      }

      // Paragraph with hyperlink
      const linkPara = new Paragraph();
      linkPara.addText('Visit ');
      linkPara.addHyperlink(Hyperlink.createExternal('https://example.com', 'Example'));
      doc.addParagraph(linkPara);

      // Bookmark + cross-reference
      const bookmark = doc.createBookmark('TestBookmark');
      const bookmarkPara = doc.createParagraph('Bookmarked text');
      bookmarkPara.addBookmark(bookmark);

      // Comment
      const comment = doc.createComment('Author', 'Review needed');
      const commentPara = doc.createParagraph('Text with comment');
      commentPara.addComment(comment);

      // Footnote and endnote
      doc.createFootnote('A footnote reference');
      doc.createEndnote('An endnote reference');

      // Image
      const image = await Image.fromBuffer(createTestPng(), 'png', 914400, 914400);
      doc.addImage(image);

      // SDT
      const sdtPara = new Paragraph();
      sdtPara.addText('SDT content');
      const sdt = StructuredDocumentTag.createRichText([sdtPara], {
        alias: 'TestSDT',
        tag: 'test-sdt',
      });
      doc.addStructuredDocumentTag(sdt);

      // Revision
      const revPara = new Paragraph();
      revPara.addRevision(new Revision({
        type: 'insert',
        author: 'TestAuthor',
        date: new Date('2025-01-15T10:00:00Z'),
        content: [new Run('Tracked change')],
      }));
      doc.addParagraph(revPara);

      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);
      doc.dispose();
    });
  });
});
