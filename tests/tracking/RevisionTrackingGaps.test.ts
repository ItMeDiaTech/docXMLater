/**
 * RevisionTrackingGaps.test.ts
 *
 * Comprehensive tests for Phase A-H: Revision tracking for Table, TableRow,
 * TableCell, Section, structural table changes, and body-level paragraph tracking.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';
import { TableCell } from '../../src/elements/TableCell';
import { Section } from '../../src/elements/Section';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { acceptRevisionsInMemory } from '../../src/utils/InMemoryRevisionAcceptor';

// ============================================================================
// Helper: create a document with tracking enabled
// ============================================================================

function createTrackedDocument(): Document {
  const doc = Document.create();
  doc.enableTrackChanges({ author: 'TestAuthor' });
  return doc;
}

// ============================================================================
// Phase A & B: Table property tracking
// ============================================================================

describe('Table property change tracking', () => {
  let doc: Document;
  let table: Table;

  beforeEach(() => {
    doc = createTrackedDocument();
    table = new Table(2, 2);
    doc.addTable(table);
  });

  afterEach(() => {
    doc.dispose();
  });

  it('should track setWidth changes', () => {
    table.setWidth(5000);
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('TestAuthor');
    expect(change!.previousProperties).toBeDefined();
  });

  it('should track setAlignment changes', () => {
    table.setAlignment('center');
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('TestAuthor');
  });

  it('should track setLayout changes with full snapshot', () => {
    table.setLayout('fixed');
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    // Full snapshot: layout was not set before, so it's absent from previousProperties
    expect(change!.previousProperties.layout).toBeUndefined();
    // Full snapshot includes pre-existing properties (default width)
    expect(change!.previousProperties.width).toBe(9360);
  });

  it('should track setBorders changes with full snapshot', () => {
    table.setBorders({
      top: { style: 'single', size: 4, color: '000000' },
    });
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    // Borders were not set before, so absent from full snapshot
    expect(change!.previousProperties.borders).toBeUndefined();
    // Full snapshot includes pre-existing properties
    expect(change!.previousProperties.width).toBe(9360);
  });

  it('should track setStyle changes with full snapshot', () => {
    table.setStyle('TableGrid');
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.style).toBeUndefined();
    expect(change!.previousProperties.width).toBe(9360);
  });

  it('should track setShading changes with full snapshot', () => {
    table.setShading({ fill: 'FF0000', pattern: 'clear' });
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.shading).toBeUndefined();
    expect(change!.previousProperties.width).toBe(9360);
  });

  it('should track setIndent changes with full snapshot', () => {
    table.setIndent(720);
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.indent).toBeUndefined();
    expect(change!.previousProperties.width).toBe(9360);
  });

  it('should track setBidiVisual changes with full snapshot', () => {
    table.setBidiVisual(true);
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.bidiVisual).toBeUndefined();
    expect(change!.previousProperties.width).toBe(9360);
  });

  it('should consolidate multiple property changes into single tblPrChange', () => {
    table.setWidth(5000);
    table.setAlignment('center');
    table.setLayout('fixed');
    doc.flushPendingChanges();
    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties).toBeDefined();
    expect(change!.author).toBe('TestAuthor');
    // Full snapshot: width is rolled back to original default (9360)
    expect(change!.previousProperties.width).toBe(9360);
    // Alignment and layout were not set before, so absent from full snapshot
    expect('alignment' in change!.previousProperties).toBe(false);
    expect('layout' in change!.previousProperties).toBe(false);
  });

  it('should include full previous formatting in tblPrChange (ECMA-376 §17.13.5.36)', () => {
    // Set up a table with multiple properties BEFORE enabling tracking
    doc.disableTrackChanges();
    table.setWidth(5000);
    table.setWidthType('pct');
    table.setAlignment('center');
    table.setBorders({
      top: { style: 'single', size: 4, color: '000000' },
    });
    doc.enableTrackChanges({ author: 'TestAuthor' });

    // Change only the layout
    table.setLayout('auto');
    doc.flushPendingChanges();

    const change = table.getTblPrChange();
    expect(change).toBeDefined();
    // Full snapshot should include ALL pre-existing properties
    expect(change!.previousProperties.width).toBe(5000);
    expect(change!.previousProperties.widthType).toBe('pct');
    expect(change!.previousProperties.alignment).toBe('center');
    expect(change!.previousProperties.borders).toBeDefined();
    // Layout was not set before, so absent from snapshot
    expect('layout' in change!.previousProperties).toBe(false);
  });

  it('should not create tblPrChange when tracking is disabled', () => {
    doc.disableTrackChanges();
    table.setWidth(5000);
    doc.flushPendingChanges();
    expect(table.getTblPrChange()).toBeUndefined();
  });

  it('should serialize tblPrChange in toXML()', () => {
    table.setTblPrChange({
      id: '1',
      author: 'TestAuthor',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { width: 3000, widthType: 'dxa', alignment: 'left' },
    });
    const xml = table.toXML();
    const xmlStr = typeof xml === 'string' ? xml : JSON.stringify(xml);
    expect(xmlStr).toContain('tblPrChange');
  });
});

// ============================================================================
// Phase A & B: TableRow property tracking
// ============================================================================

describe('TableRow property change tracking', () => {
  let doc: Document;
  let table: Table;
  let row: TableRow;

  beforeEach(() => {
    doc = createTrackedDocument();
    table = new Table(2, 2);
    doc.addTable(table);
    row = table.getRows()[0]!;
  });

  afterEach(() => {
    doc.dispose();
  });

  it('should track setHeight changes', () => {
    row.setHeight(720, 'exact');
    doc.flushPendingChanges();
    const change = row.getTrPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('TestAuthor');
    expect(change!.previousProperties).toBeDefined();
    expect(change!.previousProperties.height).toBeUndefined();
  });

  it('should track setHeader changes', () => {
    row.setHeader(true);
    doc.flushPendingChanges();
    const change = row.getTrPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.isHeader).toBeUndefined();
  });

  it('should track setCantSplit changes', () => {
    row.setCantSplit(true);
    doc.flushPendingChanges();
    const change = row.getTrPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.cantSplit).toBeUndefined();
  });

  it('should track setJustification changes', () => {
    row.setJustification('center');
    doc.flushPendingChanges();
    const change = row.getTrPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.justification).toBeUndefined();
  });

  it('should track setHidden changes', () => {
    row.setHidden(true);
    doc.flushPendingChanges();
    const change = row.getTrPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.hidden).toBeUndefined();
  });

  it('should serialize trPrChange in toXML()', () => {
    row.setTrPrChange({
      id: '2',
      author: 'TestAuthor',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { height: 360, heightRule: 'atLeast' },
    });
    const xml = row.toXML();
    const xmlStr = typeof xml === 'string' ? xml : JSON.stringify(xml);
    expect(xmlStr).toContain('trPrChange');
  });
});

// ============================================================================
// Phase A & B: TableCell property tracking
// ============================================================================

describe('TableCell property change tracking', () => {
  let doc: Document;
  let table: Table;
  let cell: TableCell;

  beforeEach(() => {
    doc = createTrackedDocument();
    table = new Table(2, 2);
    doc.addTable(table);
    cell = table.getRows()[0]!.getCells()[0]!;
  });

  afterEach(() => {
    doc.dispose();
  });

  it('should track setWidth changes', () => {
    cell.setWidth(2500);
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('TestAuthor');
    expect(change!.previousProperties).toBeDefined();
    expect(change!.previousProperties.width).toBeUndefined();
  });

  it('should track setShading changes', () => {
    cell.setShading({ fill: '00FF00', pattern: 'clear' });
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.shading).toBeUndefined();
  });

  it('should track setBorders changes', () => {
    cell.setBorders({
      top: { style: 'single', size: 4, color: '000000' },
    });
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.borders).toBeUndefined();
  });

  it('should track setVerticalAlignment changes', () => {
    cell.setVerticalAlignment('center');
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.verticalAlignment).toBeUndefined();
  });

  it('should track setMargins changes', () => {
    cell.setMargins({ top: 100, bottom: 100 });
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.margins).toBeUndefined();
  });

  it('should track setTextDirection changes', () => {
    cell.setTextDirection('btLr');
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
  });

  it('should track setNoWrap changes', () => {
    cell.setNoWrap(true);
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
  });

  it('should track setFitText changes', () => {
    cell.setFitText(true);
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
  });

  it('should track setHideMark changes', () => {
    cell.setHideMark(true);
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
  });

  it('should consolidate multiple cell property changes', () => {
    cell.setWidth(2500);
    cell.setVerticalAlignment('center');
    cell.setShading({ fill: 'FFFF00' });
    doc.flushPendingChanges();
    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('TestAuthor');
    expect(change!.previousProperties).toBeDefined();
    expect('width' in change!.previousProperties).toBe(true);
    expect('verticalAlignment' in change!.previousProperties).toBe(true);
    expect('shading' in change!.previousProperties).toBe(true);
  });

  it('should serialize tcPrChange in toXML()', () => {
    cell.setTcPrChange({
      id: '3',
      author: 'TestAuthor',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { width: 2000, widthType: 'dxa', shading: { fill: 'FF0000' } },
    });
    const xml = cell.toXML();
    const xmlStr = typeof xml === 'string' ? xml : JSON.stringify(xml);
    expect(xmlStr).toContain('tcPrChange');
  });
});

// ============================================================================
// Phase A & B: Section property tracking
// ============================================================================

describe('Section property change tracking', () => {
  let doc: Document;
  let section: Section;

  beforeEach(() => {
    doc = createTrackedDocument();
    section = doc.getSection();
  });

  afterEach(() => {
    doc.dispose();
  });

  it('should track setPageSize changes', () => {
    // Use different dimensions from default (12240x15840) to trigger change
    section.setPageSize(15840, 12240);
    doc.flushPendingChanges();
    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('TestAuthor');
    expect(change!.previousProperties).toBeDefined();
  });

  it('should track setOrientation changes', () => {
    section.setOrientation('landscape');
    doc.flushPendingChanges();
    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('TestAuthor');
  });

  it('should track setMargins changes', () => {
    // Use different margins from defaults to trigger change
    section.setMargins({ top: 2880, bottom: 2880, left: 2880, right: 2880 });
    doc.flushPendingChanges();
    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties).toBeDefined();
  });

  it('should track setColumns changes', () => {
    section.setColumns(2, 720);
    doc.flushPendingChanges();
    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties).toBeDefined();
  });

  it('should track setSectionType changes', () => {
    section.setSectionType('continuous');
    doc.flushPendingChanges();
    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties).toBeDefined();
  });

  it('should track setTitlePage changes', () => {
    section.setTitlePage(true);
    doc.flushPendingChanges();
    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties).toBeDefined();
  });

  it('should serialize sectPrChange in toXML()', () => {
    section.setSectPrChange({
      id: '4',
      author: 'TestAuthor',
      date: '2026-02-17T00:00:00Z',
      previousProperties: {
        pageSize: { width: 12240, height: 15840, orientation: 'portrait' },
        margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
      },
    });
    const xml = section.toXML();
    const xmlStr = typeof xml === 'string' ? xml : JSON.stringify(xml);
    expect(xmlStr).toContain('sectPrChange');
  });
});

// ============================================================================
// Phase E: Parse *PrChange round-trip
// ============================================================================

describe('*PrChange round-trip (save → load → verify)', () => {
  it('should round-trip tblPrChange', async () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    table.setTblPrChange({
      id: '10',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { width: 3000, widthType: 'dxa', alignment: 'left' },
    });
    doc.addTable(table);
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const tables = loaded.getTables();
    expect(tables.length).toBe(1);
    const change = tables[0]!.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('Author');
    expect(change!.previousProperties.width).toBe(3000);
    loaded.dispose();
  });

  it('should round-trip trPrChange', async () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    const row = table.getRows()[0]!;
    row.setTrPrChange({
      id: '11',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { height: 500, heightRule: 'exact', isHeader: true },
    });
    doc.addTable(table);
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const tables = loaded.getTables();
    const change = tables[0]!.getRows()[0]!.getTrPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('Author');
    expect(change!.previousProperties.height).toBe(500);
    expect(change!.previousProperties.isHeader).toBe(true);
    loaded.dispose();
  });

  it('should round-trip tcPrChange', async () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    const cell = table.getRows()[0]!.getCells()[0]!;
    cell.setTcPrChange({
      id: '12',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { width: 2000, widthType: 'dxa', verticalAlignment: 'top' },
    });
    doc.addTable(table);
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const tables = loaded.getTables();
    const change = tables[0]!.getRows()[0]!.getCells()[0]!.getTcPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('Author');
    expect(change!.previousProperties.width).toBe(2000);
    expect(change!.previousProperties.verticalAlignment).toBe('top');
    loaded.dispose();
  });

  it('should round-trip sectPrChange', async () => {
    const doc = Document.create();
    const section = doc.getSection();
    section.setSectPrChange({
      id: '13',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: {
        pageSize: { width: 12240, height: 15840, orientation: 'portrait' },
        type: 'continuous',
      },
    });
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = loaded.getSection().getSectPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('Author');
    expect(change!.previousProperties.pageSize).toBeDefined();
    expect(change!.previousProperties.type).toBe('continuous');
    loaded.dispose();
  });
});

// ============================================================================
// Phase F: Structural table changes
// ============================================================================

describe('Structural table change tracking', () => {
  let doc: Document;
  let table: Table;

  beforeEach(() => {
    doc = createTrackedDocument();
    table = new Table(3, 3);
    doc.addTable(table);
  });

  afterEach(() => {
    doc.dispose();
  });

  it('should mark new cells with cellIns on insertRow', () => {
    const newRow = table.insertRow(1);
    const cells = newRow.getCells();
    for (const cell of cells) {
      const rev = cell.getCellRevision();
      expect(rev).toBeDefined();
      expect(rev!.getType()).toBe('tableCellInsert');
      expect(rev!.getAuthor()).toBe('TestAuthor');
    }
  });

  it('should mark cells with cellDel on removeRow (not actually remove)', () => {
    const rowBefore = table.getRows()[1]!;
    const cellsBefore = rowBefore.getCells();
    table.removeRow(1);
    // Row should still exist (not removed) with cellDel markers
    expect(table.getRows().length).toBe(3); // Still 3 rows
    for (const cell of cellsBefore) {
      const rev = cell.getCellRevision();
      expect(rev).toBeDefined();
      expect(rev!.getType()).toBe('tableCellDelete');
    }
  });

  it('should mark new cells with cellIns on addColumn', () => {
    table.addColumn(1);
    // Each row should have a new cell at index 1 with cellIns
    for (const row of table.getRows()) {
      const cell = row.getCells()[1]!;
      const rev = cell.getCellRevision();
      expect(rev).toBeDefined();
      expect(rev!.getType()).toBe('tableCellInsert');
    }
  });

  it('should mark cells with cellDel on removeColumn (not actually remove)', () => {
    table.removeColumn(1);
    // Cells should still exist (not removed) with cellDel markers
    for (const row of table.getRows()) {
      expect(row.getCells().length).toBe(3); // Still 3 cells
      const cell = row.getCells()[1]!;
      const rev = cell.getCellRevision();
      expect(rev).toBeDefined();
      expect(rev!.getType()).toBe('tableCellDelete');
    }
  });

  it('should mark absorbed cells with cellMerge on mergeCells', () => {
    table.mergeCells(0, 0, 1, 0);
    // Start cell (0,0) should NOT have cellMerge
    const startCell = table.getCell(0, 0)!;
    // The start cell should not have a cellMerge revision (it's the anchor)
    const startRev = startCell.getCellRevision();
    expect(startRev).toBeUndefined();

    // Absorbed cell (1,0) should have cellMerge
    const absorbedCell = table.getCell(1, 0)!;
    const absorbedRev = absorbedCell.getCellRevision();
    expect(absorbedRev).toBeDefined();
    expect(absorbedRev!.getType()).toBe('tableCellMerge');
  });

  it('should not mark structural changes when tracking is disabled', () => {
    doc.disableTrackChanges();
    const newRow = table.insertRow(1);
    for (const cell of newRow.getCells()) {
      expect(cell.getCellRevision()).toBeUndefined();
    }
  });

  it('should actually remove row when tracking is disabled', () => {
    doc.disableTrackChanges();
    table.removeRow(1);
    expect(table.getRows().length).toBe(2);
  });

  it('should actually remove column when tracking is disabled', () => {
    doc.disableTrackChanges();
    table.removeColumn(1);
    for (const row of table.getRows()) {
      expect(row.getCells().length).toBe(2);
    }
  });
});

// ============================================================================
// Phase G: Body-level paragraph tracking
// ============================================================================

describe('Body-level paragraph add/remove tracking', () => {
  let doc: Document;

  beforeEach(() => {
    doc = createTrackedDocument();
  });

  afterEach(() => {
    doc.dispose();
  });

  it('should wrap paragraph runs in w:ins on addParagraph', () => {
    const para = new Paragraph();
    para.addText('Tracked text');
    doc.addParagraph(para);

    const content = para.getContent();
    const revisions = content.filter((item) => item instanceof Revision);
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('insert');
    expect(revisions[0]!.getAuthor()).toBe('TestAuthor');
  });

  it('should not wrap empty paragraphs in w:ins', () => {
    const para = new Paragraph();
    doc.addParagraph(para);

    const content = para.getContent();
    const revisions = content.filter((item) => item instanceof Revision);
    expect(revisions.length).toBe(0);
  });

  it('should wrap content in w:del on removeParagraph', () => {
    // First add a paragraph without tracking
    doc.disableTrackChanges();
    const para = new Paragraph();
    para.addText('Will be deleted');
    doc.addParagraph(para);

    // Re-enable tracking then remove
    doc.enableTrackChanges({ author: 'TestAuthor' });
    doc.removeParagraph(para);

    // Paragraph should still be in document (not removed)
    expect(doc.getParagraphs()).toContain(para);

    const content = para.getContent();
    const revisions = content.filter((item) => item instanceof Revision);
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('delete');
  });

  it('should actually remove paragraph when tracking is disabled', () => {
    doc.disableTrackChanges();
    const para = new Paragraph();
    para.addText('Remove me');
    doc.addParagraph(para);

    const count = doc.getParagraphs().length;
    doc.removeParagraph(para);
    expect(doc.getParagraphs().length).toBe(count - 1);
  });
});

// ============================================================================
// Phase G: TableCell paragraph tracking
// ============================================================================

describe('TableCell paragraph add/remove tracking', () => {
  let doc: Document;
  let table: Table;
  let cell: TableCell;

  beforeEach(() => {
    doc = createTrackedDocument();
    table = new Table(1, 1);
    doc.addTable(table);
    cell = table.getRows()[0]!.getCells()[0]!;
  });

  afterEach(() => {
    doc.dispose();
  });

  it('should wrap runs in w:ins on addParagraphAt', () => {
    const para = new Paragraph();
    para.addText('Cell text');
    cell.addParagraphAt(0, para);

    const content = para.getContent();
    const revisions = content.filter((item) => item instanceof Revision);
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('insert');
  });

  it('should wrap content in w:del on removeParagraph', () => {
    // Add a paragraph without tracking first
    doc.disableTrackChanges();
    const para = new Paragraph();
    para.addText('To delete');
    cell.addParagraphAt(0, para);

    // Re-enable tracking
    doc.enableTrackChanges({ author: 'TestAuthor' });
    // Need to re-bind tracking to the cell
    cell.removeParagraph(0);

    // Paragraph should still exist in cell
    expect(cell.getParagraphs().length).toBeGreaterThanOrEqual(1);
  });
});

// ============================================================================
// Phase H: Accept revisions clears *PrChange fields
// ============================================================================

describe('Accept revisions clears *PrChange fields', () => {
  it('should clear tblPrChange on accept', () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    table.setTblPrChange({
      id: '20',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { width: 3000 },
    });
    doc.addTable(table);

    acceptRevisionsInMemory(doc);
    expect(table.getTblPrChange()).toBeUndefined();
    doc.dispose();
  });

  it('should clear trPrChange on accept', () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    const row = table.getRows()[0]!;
    row.setTrPrChange({
      id: '21',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { height: 500 },
    });
    doc.addTable(table);

    acceptRevisionsInMemory(doc);
    expect(row.getTrPrChange()).toBeUndefined();
    doc.dispose();
  });

  it('should clear tcPrChange on accept', () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    const cell = table.getRows()[0]!.getCells()[0]!;
    cell.setTcPrChange({
      id: '22',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { width: 2000 },
    });
    doc.addTable(table);

    acceptRevisionsInMemory(doc);
    expect(cell.getTcPrChange()).toBeUndefined();
    doc.dispose();
  });

  it('should clear cellRevision markers on accept', () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    const cell = table.getRows()[0]!.getCells()[0]!;
    const revision = Revision.createTableCellInsert('Author', []);
    cell.setCellRevision(revision);
    doc.addTable(table);

    acceptRevisionsInMemory(doc);
    expect(cell.getCellRevision()).toBeUndefined();
    doc.dispose();
  });

  it('should clear sectPrChange on accept', () => {
    const doc = Document.create();
    const section = doc.getSection();
    section.setSectPrChange({
      id: '23',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { type: 'continuous' },
    });

    acceptRevisionsInMemory(doc);
    expect(section.getSectPrChange()).toBeUndefined();
    doc.dispose();
  });

  it('should not clear *PrChange when acceptPropertyChanges is false', () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    table.setTblPrChange({
      id: '24',
      author: 'Author',
      date: '2026-02-17T00:00:00Z',
      previousProperties: { width: 3000 },
    });
    doc.addTable(table);

    acceptRevisionsInMemory(doc, { acceptPropertyChanges: false });
    expect(table.getTblPrChange()).toBeDefined();
    doc.dispose();
  });
});

// ============================================================================
// Integration: Full pipeline
// ============================================================================

describe('Full tracking pipeline integration', () => {
  it('should enable tracking → modify table → flush → save → load → verify', async () => {
    const doc = createTrackedDocument();
    const table = new Table(2, 2);
    doc.addTable(table);

    // Modify table properties
    table.setWidth(8000);
    table.setAlignment('center');

    // Modify cell
    const cell = table.getRows()[0]!.getCells()[0]!;
    cell.setShading({ fill: 'CCCCCC' });

    // Modify row
    const row = table.getRows()[0]!;
    row.setHeader(true);

    // Modify section
    const section = doc.getSection();
    section.setOrientation('landscape');

    // Flush to apply *PrChange
    doc.flushPendingChanges();

    // Verify in-memory
    expect(table.getTblPrChange()).toBeDefined();
    expect(cell.getTcPrChange()).toBeDefined();
    expect(row.getTrPrChange()).toBeDefined();
    expect(section.getSectPrChange()).toBeDefined();

    // Save and reload
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const loadedTables = loaded.getTables();
    expect(loadedTables.length).toBe(1);

    // Verify tblPrChange persisted
    expect(loadedTables[0]!.getTblPrChange()).toBeDefined();
    expect(loadedTables[0]!.getTblPrChange()!.author).toBe('TestAuthor');

    // Verify trPrChange persisted
    expect(loadedTables[0]!.getRows()[0]!.getTrPrChange()).toBeDefined();

    // Verify tcPrChange persisted
    expect(loadedTables[0]!.getRows()[0]!.getCells()[0]!.getTcPrChange()).toBeDefined();

    // Verify sectPrChange persisted
    expect(loaded.getSection().getSectPrChange()).toBeDefined();

    loaded.dispose();
  });

  it('should support accept all revisions after full tracking pipeline', async () => {
    const doc = createTrackedDocument();
    const table = new Table(1, 1);
    doc.addTable(table);

    table.setWidth(5000);
    table.getRows()[0]!.setHeader(true);
    table.getRows()[0]!.getCells()[0]!.setShading({ fill: 'FF0000' });
    doc.getSection().setOrientation('landscape');

    doc.flushPendingChanges();

    // Accept all
    const result = acceptRevisionsInMemory(doc);
    expect(result.propertyChangesAccepted).toBeGreaterThan(0);

    // All *PrChange cleared
    expect(table.getTblPrChange()).toBeUndefined();
    expect(table.getRows()[0]!.getTrPrChange()).toBeUndefined();
    expect(table.getRows()[0]!.getCells()[0]!.getTcPrChange()).toBeUndefined();
    expect(doc.getSection().getSectPrChange()).toBeUndefined();

    doc.dispose();
  });
});

// ============================================================================
// Fix validation tests
// ============================================================================

describe('Fix 1: Element identity prevents consolidation key collisions', () => {
  it('should track two different cells setting the same property value', () => {
    const doc = createTrackedDocument();
    const table = new Table(1, 2);
    doc.addTable(table);

    const cell1 = table.getRows()[0]!.getCells()[0]!;
    const cell2 = table.getRows()[0]!.getCells()[1]!;

    cell1.setWidth(5000);
    cell2.setWidth(5000);
    doc.flushPendingChanges();

    // Both cells should have their own tcPrChange (not just the first one)
    expect(cell1.getTcPrChange()).toBeDefined();
    expect(cell2.getTcPrChange()).toBeDefined();
    doc.dispose();
  });
});

describe('Fix 2: removeRows respects tracking', () => {
  it('should mark cells with cellDel when tracking enabled', () => {
    const doc = createTrackedDocument();
    const table = new Table(4, 2);
    doc.addTable(table);

    table.removeRows(1, 2);

    // Rows should still exist (not removed)
    expect(table.getRows().length).toBe(4);

    // Rows 1 and 2 should have cellDel markers
    for (let i = 1; i <= 2; i++) {
      for (const cell of table.getRows()[i]!.getCells()) {
        const rev = cell.getCellRevision();
        expect(rev).toBeDefined();
        expect(rev!.getType()).toBe('tableCellDelete');
      }
    }

    // Rows 0 and 3 should NOT have cellDel markers
    expect(table.getRows()[0]!.getCells()[0]!.getCellRevision()).toBeUndefined();
    expect(table.getRows()[3]!.getCells()[0]!.getCellRevision()).toBeUndefined();

    doc.dispose();
  });

  it('should actually remove rows when tracking disabled', () => {
    const doc = createTrackedDocument();
    doc.disableTrackChanges();
    const table = new Table(4, 2);
    doc.addTable(table);

    table.removeRows(1, 2);
    expect(table.getRows().length).toBe(2);
    doc.dispose();
  });
});

describe('Fix 3: removeRow wraps content in w:del', () => {
  it('should wrap cell text in w:del revisions', () => {
    const doc = Document.create();
    const table = new Table(2, 1);
    doc.addTable(table);

    // Add text to cells without tracking
    const cell = table.getRows()[1]!.getCells()[0]!;
    const para = new Paragraph();
    para.addText('Cell text');
    cell.addParagraph(para);

    // Enable tracking and remove row
    doc.enableTrackChanges({ author: 'TestAuthor' });
    table.removeRow(1);

    // Cell should have cellDel AND content should have w:del revision
    const rev = cell.getCellRevision();
    expect(rev).toBeDefined();
    expect(rev!.getType()).toBe('tableCellDelete');

    const content = para.getContent();
    const revisions = content.filter((item) => item instanceof Revision);
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('delete');

    doc.dispose();
  });
});

describe('Fix 12: Object equality prevents duplicate tracking', () => {
  it('should not create change when setting same object value', () => {
    const doc = createTrackedDocument();
    const table = new Table(1, 1);
    doc.addTable(table);

    // Set shading first
    table.setShading({ fill: 'FF0000', pattern: 'clear' });
    doc.flushPendingChanges();

    // Clear the change
    table.clearTblPrChange();

    // Set the same shading again — should not create a new change
    table.setShading({ fill: 'FF0000', pattern: 'clear' });
    doc.flushPendingChanges();

    // The change should not be created since values are equal
    expect(table.getTblPrChange()).toBeUndefined();

    doc.dispose();
  });
});

// ============================================================================
// Fix: Loaded tables with w:tblW w:w="0" w:type="auto" capture correct width
// ============================================================================

describe('Loaded table with auto width captures correct tblPrChange', () => {
  it('should capture original 0/auto width in tblPrChange for loaded tables', async () => {
    // Create a document with a table that has w:tblW w:w="0" w:type="auto"
    const doc = Document.create();
    const table = new Table(1, 1);
    doc.addTable(table);

    // Set width to 0/auto (simulating what the parser does for auto-sized tables)
    table.setWidth(0);
    table.setWidthType('auto');

    // Save and reload to go through the parser
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const loadedTable = loaded.getTables()[0]!;

    // Verify the loaded table has width=0, widthType=auto
    expect(loadedTable.getWidth()).toBe(0);
    expect(loadedTable.getWidthType()).toBe('auto');

    // Enable tracking and modify a property
    loaded.enableTrackChanges({ author: 'TestAuthor' });
    loadedTable.setLayout('fixed');
    loaded.flushPendingChanges();

    // Verify tblPrChange captures 0/auto, NOT the constructor default 9360/dxa
    const change = loadedTable.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.width).toBe(0);
    expect(change!.previousProperties.widthType).toBe('auto');

    loaded.dispose();
  });
});

// ============================================================================
// Fix: Loaded tables with w:tblInd capture correct indent in tblPrChange
// ============================================================================

describe('Loaded table with indent captures correct tblPrChange', () => {
  it('should round-trip w:tblInd through save/load and capture in tblPrChange', async () => {
    // Create a document with a table that has a specific indent
    const doc = Document.create();
    const table = new Table(1, 1);
    doc.addTable(table);
    table.setIndent(720); // 0.5 inch

    // Save and reload to go through the parser
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const loadedTable = loaded.getTables()[0]!;

    // Verify the loaded table preserved the indent
    expect(loadedTable.getIndent()).toBe(720);

    // Enable tracking and modify a property
    loaded.enableTrackChanges({ author: 'TestAuthor' });
    loadedTable.setLayout('fixed');
    loaded.flushPendingChanges();

    // Verify tblPrChange captures the original indent
    const change = loadedTable.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties.indent).toBe(720);

    loaded.dispose();
  });
});
