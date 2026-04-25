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
import { acceptRevisionsInMemory } from '../../src/processors/InMemoryRevisionAcceptor';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

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
    // Full snapshot: properties not set before the change are absent from snapshot
    // (they were undefined, meaning the cell had no explicit width/vAlign/shading)
    expect('width' in change!.previousProperties).toBe(false);
    expect('verticalAlignment' in change!.previousProperties).toBe(false);
    expect('shading' in change!.previousProperties).toBe(false);
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

  it('should round-trip extended sectPrChange properties (bidi, rtlGutter, docGrid, titlePage, vAlign)', async () => {
    const doc = Document.create();
    const section = doc.getSection();
    section.setSectPrChange({
      id: '14',
      author: 'Author',
      date: '2026-03-01T00:00:00Z',
      previousProperties: {
        pageSize: { width: 12240, height: 15840 },
        titlePage: true,
        verticalAlignment: 'center',
        textDirection: 'tbRl',
        bidi: true,
        rtlGutter: true,
        docGrid: { type: 'lines', linePitch: 360 },
        lineNumbering: { countBy: 5, start: 1, restart: 'newPage' },
        pageNumbering: { start: 1, format: 'decimal' },
      },
    });
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = loaded.getSection().getSectPrChange();
    expect(change).toBeDefined();
    const prev = change!.previousProperties;

    expect(prev.titlePage).toBe(true);
    expect(prev.verticalAlignment).toBe('center');
    expect(prev.textDirection).toBe('tbRl');
    expect(prev.bidi).toBe(true);
    expect(prev.rtlGutter).toBe(true);
    expect(prev.docGrid?.type).toBe('lines');
    expect(prev.docGrid?.linePitch).toBe(360);
    expect(prev.lineNumbering?.countBy).toBe(5);
    expect(prev.pageNumbering?.format).toBe('decimal');

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

// ============================================================================
// Table Scrunching Bug Fixes — sectPrChange full snapshot
// ============================================================================

describe('Fix A: sectPrChange preserves full original section properties', () => {
  let doc: Document;
  let section: Section;

  beforeEach(() => {
    doc = createTrackedDocument();
    section = doc.getSection();
  });

  afterEach(() => {
    doc.dispose();
  });

  it('should include full previous properties in sectPrChange (not just delta)', () => {
    // Change only orientation — sectPrChange should still include pageSize, margins, columns, type
    section.setOrientation('landscape');
    doc.flushPendingChanges();

    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    expect(change!.previousProperties).toBeDefined();

    // Full snapshot should include ALL section properties, not just orientation
    expect(change!.previousProperties.pageSize).toBeDefined();
    expect(change!.previousProperties.pageSize.width).toBe(12240);
    expect(change!.previousProperties.pageSize.height).toBe(15840);
    expect(change!.previousProperties.margins).toBeDefined();
    expect(change!.previousProperties.margins.top).toBe(1440);
    expect(change!.previousProperties.margins.left).toBe(1440);
    expect(change!.previousProperties.columns).toBeDefined();
    expect(change!.previousProperties.type).toBe('nextPage');
  });

  it('should preserve original portrait pageSize when switching to landscape', () => {
    section.setOrientation('landscape');
    doc.flushPendingChanges();

    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    // The original was portrait, so the previous state should be portrait
    expect(change!.previousProperties.pageSize.orientation).toBe('portrait');
    expect(change!.previousProperties.pageSize.width).toBe(12240);
    expect(change!.previousProperties.pageSize.height).toBe(15840);
  });

  it('should serialize full sectPrChange with pgSz, pgMar, cols in XML', () => {
    section.setSectPrChange({
      id: '100',
      author: 'TestAuthor',
      date: '2026-02-22T00:00:00Z',
      previousProperties: {
        pageSize: { width: 12240, height: 15840, orientation: 'portrait' },
        margins: { top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720 },
        columns: { count: 1, space: 720 },
        docGrid: { linePitch: 360 },
      },
    });
    const xml = section.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('sectPrChange');
    expect(xmlStr).toContain('pgSz');
    expect(xmlStr).toContain('pgMar');
    expect(xmlStr).toContain('cols');
    expect(xmlStr).toContain('docGrid');
    expect(xmlStr).toContain('linePitch');
  });

  it('should include gutter in sectPrChange pgMar when present', () => {
    section.setSectPrChange({
      id: '102',
      author: 'TestAuthor',
      date: '2026-02-22T00:00:00Z',
      previousProperties: {
        margins: {
          top: 1440,
          bottom: 1440,
          left: 1440,
          right: 1440,
          header: 720,
          footer: 720,
          gutter: 360,
        },
      },
    });
    const xml = section.toXML();
    const xmlStr = XMLBuilder.elementToString(xml);
    expect(xmlStr).toContain('w:gutter="360"');
  });

  it('should include docGrid in sectPrChange when present', () => {
    section.setSectPrChange({
      id: '101',
      author: 'TestAuthor',
      date: '2026-02-22T00:00:00Z',
      previousProperties: {
        docGrid: { linePitch: 360, type: 'default' },
      },
    });
    const xml = section.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('docGrid');
    expect(xmlStr).toContain('360');
  });

  it('should NOT produce empty sectPr in sectPrChange', () => {
    // Changing just the orientation should produce a sectPrChange with populated content
    section.setOrientation('landscape');
    doc.flushPendingChanges();

    const change = section.getSectPrChange();
    expect(change).toBeDefined();
    // Must NOT be empty — should have at least pageSize and margins
    expect(Object.keys(change!.previousProperties).length).toBeGreaterThan(0);
    expect(change!.previousProperties.pageSize).toBeDefined();
    expect(change!.previousProperties.margins).toBeDefined();
  });
});

// ============================================================================
// Table Scrunching Bug Fixes — tblGridChange
// ============================================================================

describe('Fix B: tblGridChange created when modifying table grid', () => {
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

  it('should create tblGridChange when setTableGrid is called with previous grid', () => {
    // First set a grid without tracking
    doc.disableTrackChanges();
    table.setTableGrid([4680, 4680]);
    doc.enableTrackChanges({ author: 'TestAuthor' });

    // Now modify the grid — should create tblGridChange
    table.setTableGrid([6480, 6480]);

    const gridChange = table.getTblGridChange();
    expect(gridChange).toBeDefined();
    expect(gridChange!.getAuthor()).toBe('TestAuthor');
    expect(gridChange!.getPreviousGrid()).toEqual([{ width: 4680 }, { width: 4680 }]);
  });

  it('should NOT create tblGridChange when no previous grid exists', () => {
    // Table has no explicit grid set — setTableGrid should not create tblGridChange
    table.setTableGrid([6480, 6480]);
    expect(table.getTblGridChange()).toBeUndefined();
  });

  it('should preserve first tblGridChange (original baseline)', () => {
    doc.disableTrackChanges();
    table.setTableGrid([4680, 4680]);
    doc.enableTrackChanges({ author: 'TestAuthor' });

    // First modification
    table.setTableGrid([6480, 6480]);
    // Second modification — should keep original grid, not intermediate
    table.setTableGrid([3240, 3240, 3240, 3240]);

    const gridChange = table.getTblGridChange();
    expect(gridChange).toBeDefined();
    expect(gridChange!.getPreviousGrid()).toEqual([{ width: 4680 }, { width: 4680 }]);
  });

  it('should serialize tblGridChange inside tblGrid element', () => {
    doc.disableTrackChanges();
    table.setTableGrid([4680, 4680]);
    doc.enableTrackChanges({ author: 'TestAuthor' });
    table.setTableGrid([6480, 6480]);

    const xml = table.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('tblGridChange');
    expect(xmlStr).toContain('4680');
  });

  it('should NOT create tblPrChange for tableGrid (uses tblGridChange instead)', () => {
    doc.disableTrackChanges();
    table.setTableGrid([4680, 4680]);
    doc.enableTrackChanges({ author: 'TestAuthor' });

    table.setTableGrid([6480, 6480]);
    doc.flushPendingChanges();

    // tblGridChange should exist, NOT tblPrChange for tableGrid
    expect(table.getTblGridChange()).toBeDefined();
    // tblPrChange should NOT contain tableGrid
    const tblPrChange = table.getTblPrChange();
    if (tblPrChange) {
      expect(tblPrChange.previousProperties.tableGrid).toBeUndefined();
    }
  });
});

describe('tblGridChange round-trip parsing', () => {
  it('should round-trip tblGridChange through document save/load', async () => {
    const { TableGridChange } = require('../../src/elements/TableGridChange');
    const table2 = new Table(1, 2);
    table2.getRows()[0]!.getCells()[0]!.createParagraph('A');
    table2.setTableGrid([3000, 4000]);
    table2.setTblGridChange(TableGridChange.create(1, [{ width: 2500 }, { width: 3500 }]));
    const doc2 = Document.create();
    doc2.addTable(table2);

    const buffer = await doc2.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const loadedTable = loaded.getTables()[0];
    const gridChange = loadedTable?.getTblGridChange();

    expect(gridChange).toBeDefined();
    const prevGrid = gridChange!.getPreviousGrid();
    expect(prevGrid).toHaveLength(2);
    expect(prevGrid[0]!.width).toBe(2500);
    expect(prevGrid[1]!.width).toBe(3500);

    doc2.dispose();
    loaded.dispose();
  });
});

// ============================================================================
// Table Scrunching Bug Fixes — tblLook extended attributes
// ============================================================================

describe('Fix C: tblLook includes extended attributes', () => {
  it('should serialize tblLook with extended attributes in tblPr', () => {
    const table = new Table(1, 1, { tblLook: '04A0' });
    const xml = table.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('w:firstRow');
    expect(xmlStr).toContain('w:lastRow');
    expect(xmlStr).toContain('w:firstColumn');
    expect(xmlStr).toContain('w:lastColumn');
    expect(xmlStr).toContain('w:noHBand');
    expect(xmlStr).toContain('w:noVBand');
  });

  it('should decode 04A0 correctly', () => {
    // 0x04A0 = 0000 0100 1010 0000
    // firstRow=1 (0x0020), lastRow=0, firstColumn=1 (0x0080), lastColumn=0
    // noHBand=0, noVBand=1 (0x0400)
    const table = new Table(1, 1, { tblLook: '04A0' });
    const flags = table.getTblLookFlags();
    expect(flags.firstRow).toBe(true);
    expect(flags.lastRow).toBe(false);
    expect(flags.firstColumn).toBe(true);
    expect(flags.lastColumn).toBe(false);
    expect(flags.noHBand).toBe(false);
    expect(flags.noVBand).toBe(true);
  });

  it('should serialize extended tblLook in tblPrChange', () => {
    const table = new Table(1, 1);
    table.setTblPrChange({
      id: '50',
      author: 'TestAuthor',
      date: '2026-02-22T00:00:00Z',
      previousProperties: { tblLook: '04A0' },
    });
    const xml = table.toXML();
    const xmlStr = JSON.stringify(xml);
    // tblPrChange should contain extended tblLook attributes
    expect(xmlStr).toContain('tblPrChange');
    expect(xmlStr).toContain('w:firstRow');
  });

  it('should serialize tblPrChange previousProperties in CT_TblPr schema order (ECMA-376 §17.4.58)', () => {
    const table = new Table(1, 1);
    table.setTblPrChange({
      id: '60',
      author: 'TestAuthor',
      date: '2026-03-01T00:00:00Z',
      previousProperties: {
        style: 'TableGrid',
        overlap: 'overlap',
        bidiVisual: true,
        tblStyleRowBandSize: 2,
        tblStyleColBandSize: 3,
        width: 5000,
        widthType: 'dxa',
        alignment: 'center',
        cellSpacing: 20,
        indent: 360,
        borders: {
          top: { style: 'single', size: 4, color: '000000' },
          bottom: { style: 'single', size: 4, color: '000000' },
        },
        shading: { pattern: 'clear', fill: 'FFFFFF' },
        layout: 'fixed',
        cellMargins: { top: 50, bottom: 50 },
        tblLook: '04A0',
        caption: 'Test Caption',
        description: 'Test Description',
      },
    });

    const xmlElement = table.toXML();
    const fullXml = XMLBuilder.elementToString(xmlElement);

    // Extract only the tblPrChange portion to avoid matching main tblPr elements
    const changeStart = fullXml.indexOf('<w:tblPrChange');
    expect(changeStart).toBeGreaterThan(-1);
    const xml = fullXml.substring(changeStart);

    // Extract positions of all tblPrChange child elements
    // These must appear in CT_TblPr order per ECMA-376
    const tblStylePos = xml.indexOf('<w:tblStyle');
    const tblOverlapPos = xml.indexOf('<w:tblOverlap');
    const bidiVisualPos = xml.indexOf('<w:bidiVisual');
    const rowBandPos = xml.indexOf('<w:tblStyleRowBandSize');
    const colBandPos = xml.indexOf('<w:tblStyleColBandSize');
    const tblWPos = xml.indexOf('<w:tblW');
    const jcPos = xml.indexOf('<w:jc');
    const cellSpacingPos = xml.indexOf('<w:tblCellSpacing');
    const tblIndPos = xml.indexOf('<w:tblInd');
    const tblBordersPos = xml.indexOf('<w:tblBorders');
    const shdPos = xml.indexOf('<w:shd');
    const tblLayoutPos = xml.indexOf('<w:tblLayout');
    const tblCellMarPos = xml.indexOf('<w:tblCellMar');
    const tblLookPos = xml.indexOf('<w:tblLook');
    const captionPos = xml.indexOf('<w:tblCaption');
    const descriptionPos = xml.indexOf('<w:tblDescription');

    // All elements should be present
    expect(tblStylePos).toBeGreaterThan(-1);
    expect(tblOverlapPos).toBeGreaterThan(-1);
    expect(bidiVisualPos).toBeGreaterThan(-1);
    expect(tblWPos).toBeGreaterThan(-1);
    expect(captionPos).toBeGreaterThan(-1);

    // Verify CT_TblPr order: tblStyle < tblOverlap < bidiVisual < rowBandSize < colBandSize
    //   < tblW < jc < tblCellSpacing < tblInd < tblBorders < shd < tblLayout
    //   < tblCellMar < tblLook < tblCaption < tblDescription
    expect(tblStylePos).toBeLessThan(tblOverlapPos);
    expect(tblOverlapPos).toBeLessThan(bidiVisualPos);
    expect(bidiVisualPos).toBeLessThan(rowBandPos);
    expect(rowBandPos).toBeLessThan(colBandPos);
    expect(colBandPos).toBeLessThan(tblWPos);
    expect(tblWPos).toBeLessThan(jcPos);
    expect(jcPos).toBeLessThan(cellSpacingPos);
    expect(cellSpacingPos).toBeLessThan(tblIndPos);
    expect(tblIndPos).toBeLessThan(tblBordersPos);
    expect(tblBordersPos).toBeLessThan(shdPos);
    expect(shdPos).toBeLessThan(tblLayoutPos);
    expect(tblLayoutPos).toBeLessThan(tblCellMarPos);
    expect(tblCellMarPos).toBeLessThan(tblLookPos);
    expect(tblLookPos).toBeLessThan(captionPos);
    expect(captionPos).toBeLessThan(descriptionPos);
  });
});

// ============================================================================
// Table Scrunching Bug Fixes — tcPrChange full snapshot
// ============================================================================

describe('Fix D: tcPrChange preserves full original cell properties', () => {
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

  it('should include full previous properties in tcPrChange (not just delta)', () => {
    const cell = table.getRows()[0]!.getCells()[0]!;

    // Set properties before tracking
    doc.disableTrackChanges();
    cell.setWidthType(5000, 'pct');
    cell.setShading({
      fill: 'BFBFBF',
      pattern: 'clear',
      themeFill: 'background1',
      themeFillShade: 'BF',
    });
    doc.enableTrackChanges({ author: 'TestAuthor' });

    // Change only the width
    cell.setWidthType(12960, 'dxa');
    doc.flushPendingChanges();

    const change = cell.getTcPrChange();
    expect(change).toBeDefined();
    // Full snapshot should include ALL pre-existing properties, not just width
    expect(change!.previousProperties.width).toBe(5000);
    expect(change!.previousProperties.widthType).toBe('pct');
    expect(change!.previousProperties.shading).toBeDefined();
    expect(change!.previousProperties.shading.fill).toBe('BFBFBF');
    expect(change!.previousProperties.shading.themeFill).toBe('background1');
    expect(change!.previousProperties.shading.themeFillShade).toBe('BF');
  });

  it('should preserve theme attributes in shading', () => {
    const cell = table.getRows()[0]!.getCells()[0]!;
    cell.setTcPrChange({
      id: '60',
      author: 'TestAuthor',
      date: '2026-02-22T00:00:00Z',
      previousProperties: {
        width: 5000,
        widthType: 'pct',
        shading: {
          fill: 'BFBFBF',
          pattern: 'clear',
          themeFill: 'background1',
          themeFillShade: 'BF',
        },
      },
    });
    const xml = cell.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('tcPrChange');
    expect(xmlStr).toContain('themeFill');
    expect(xmlStr).toContain('themeFillShade');
  });

  it('should serialize gridSpan and vMerge in tcPrChange', () => {
    const cell = table.getRows()[0]!.getCells()[0]!;
    cell.setTcPrChange({
      id: '61',
      author: 'TestAuthor',
      date: '2026-02-22T00:00:00Z',
      previousProperties: {
        width: 5000,
        widthType: 'dxa',
        columnSpan: 2,
        vMerge: 'restart',
      },
    });
    const xml = cell.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('tcPrChange');
    expect(xmlStr).toContain('gridSpan');
    expect(xmlStr).toContain('vMerge');
  });
});

describe('tcPrChange previousProperties parsing completeness', () => {
  it('should round-trip gridSpan, vMerge, margins, textDirection, noWrap in tcPrChange', async () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    const cell = table.getRows()[0]!.getCells()[0]!;
    cell.createParagraph('Test');
    cell.setTcPrChange({
      id: '70',
      author: 'TestAuthor',
      date: '2026-03-01T00:00:00Z',
      previousProperties: {
        width: 3000,
        widthType: 'dxa',
        columnSpan: 3,
        vMerge: 'restart',
        noWrap: true,
        margins: { top: 50, bottom: 50, left: 100, right: 100 },
        textDirection: 'btLr',
        fitText: true,
        verticalAlignment: 'center',
        hideMark: true,
      },
    });
    doc.addTable(table);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const loadedCell = loaded.getTables()[0]?.getRows()[0]?.getCells()[0];
    const change = loadedCell?.getTcPrChange();

    expect(change).toBeDefined();
    const prev = change!.previousProperties;
    expect(prev.width).toBe(3000);
    expect(prev.columnSpan).toBe(3);
    expect(prev.vMerge).toBe('restart');
    expect(prev.noWrap).toBe(true);
    expect(prev.margins).toBeDefined();
    expect(prev.margins.top).toBe(50);
    expect(prev.margins.left).toBe(100);
    expect(prev.textDirection).toBe('btLr');
    expect(prev.fitText).toBe(true);
    expect(prev.verticalAlignment).toBe('center');
    expect(prev.hideMark).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('trPrChange previousProperties parsing completeness', () => {
  it('should round-trip gridBefore, gridAfter, wBefore, wAfter, justification in trPrChange', async () => {
    const doc = Document.create();
    const table = new Table(1, 2);
    const row = table.getRows()[0]!;
    row.getCells()[0]!.createParagraph('Test');
    row.setTrPrChange({
      id: '80',
      author: 'TestAuthor',
      date: '2026-03-01T00:00:00Z',
      previousProperties: {
        height: 400,
        heightRule: 'atLeast',
        isHeader: true,
        gridBefore: 1,
        gridAfter: 2,
        wBefore: 500,
        wBeforeType: 'dxa',
        wAfter: 300,
        wAfterType: 'dxa',
        justification: 'center',
      },
    });
    doc.addTable(table);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const loadedRow = loaded.getTables()[0]?.getRows()[0];
    const change = loadedRow?.getTrPrChange();

    expect(change).toBeDefined();
    const prev = change!.previousProperties;
    expect(prev.height).toBe(400);
    expect(prev.heightRule).toBe('atLeast');
    expect(prev.isHeader).toBe(true);
    expect(prev.gridBefore).toBe(1);
    expect(prev.gridAfter).toBe(2);
    expect(prev.wBefore).toBe(500);
    expect(prev.wAfter).toBe(300);
    expect(prev.justification).toBe('center');

    doc.dispose();
    loaded.dispose();
  });
});

describe('tblPrChange cellSpacingType and cellMargins parsing', () => {
  it('should round-trip cellSpacingType in tblPrChange', async () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    table.getRows()[0]!.getCells()[0]!.createParagraph('Test');
    table.setTblPrChange({
      id: '90',
      author: 'TestAuthor',
      date: '2026-03-01T00:00:00Z',
      previousProperties: {
        style: 'TableGrid',
        width: 5000,
        widthType: 'dxa',
        cellSpacing: 20,
        cellSpacingType: 'dxa',
      },
    });
    doc.addTable(table);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = loaded.getTables()[0]?.getTblPrChange();

    expect(change).toBeDefined();
    expect(change!.previousProperties.cellSpacing).toBe(20);
    expect(change!.previousProperties.cellSpacingType).toBe('dxa');

    doc.dispose();
    loaded.dispose();
  });

  it('should round-trip cellMargins in tblPrChange', async () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    table.getRows()[0]!.getCells()[0]!.createParagraph('Test');
    table.setTblPrChange({
      id: '91',
      author: 'TestAuthor',
      date: '2026-03-01T00:00:00Z',
      previousProperties: {
        width: 5000,
        widthType: 'dxa',
        cellMargins: { top: 50, left: 115, bottom: 50, right: 115 },
      },
    });
    doc.addTable(table);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = loaded.getTables()[0]?.getTblPrChange();

    expect(change).toBeDefined();
    const cm = change!.previousProperties.cellMargins;
    expect(cm).toBeDefined();
    expect(cm.top).toBe(50);
    expect(cm.bottom).toBe(50);
    expect(cm.left).toBe(115);
    expect(cm.right).toBe(115);

    doc.dispose();
    loaded.dispose();
  });
});

describe('tblPrChange extended table properties parsing', () => {
  it('should round-trip tblpPr, bidiVisual, overlap, bandSizes, caption, description, tblLook in tblPrChange', async () => {
    const doc = Document.create();
    const table = new Table(1, 1);
    table.getRows()[0]!.getCells()[0]!.createParagraph('Test');
    table.setTblPrChange({
      id: '95',
      author: 'TestAuthor',
      date: '2026-03-15T00:00:00Z',
      previousProperties: {
        style: 'TableGrid',
        width: 5000,
        widthType: 'dxa',
        position: { x: 100, y: 200, horizontalAnchor: 'margin', leftFromText: 50 },
        overlap: 'overlap',
        bidiVisual: true,
        tblLook: '04A0',
        caption: 'Test Table',
        description: 'A test table description',
      },
    });
    doc.addTable(table);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const change = loaded.getTables()[0]?.getTblPrChange();

    expect(change).toBeDefined();
    const prev = change!.previousProperties;
    expect(prev.position).toBeDefined();
    expect(prev.position.x).toBe(100);
    expect(prev.position.horizontalAnchor).toBe('margin');
    expect(prev.overlap).toBe('overlap');
    expect(prev.bidiVisual).toBe(true);
    expect(prev.tblLook).toBeDefined();
    expect(prev.caption).toBe('Test Table');
    expect(prev.description).toBe('A test table description');

    doc.dispose();
    loaded.dispose();
  });
});
