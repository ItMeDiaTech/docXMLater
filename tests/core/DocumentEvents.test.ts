/**
 * Document event system: lifecycle (load/save) + mutation
 * (paragraphAdded, tableAdded, paragraphRemoved) hooks.
 */
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';

describe('Document events', () => {
  it('fires paragraphAdded on createParagraph()', () => {
    const doc = Document.create();
    const seen: string[] = [];
    doc.on('paragraphAdded', ({ paragraph }) => {
      seen.push(paragraph.getText());
    });
    doc.createParagraph('A');
    doc.createParagraph('B');
    expect(seen).toEqual(['A', 'B']);
    doc.dispose();
  });

  it('fires paragraphAdded on addParagraph()', () => {
    const doc = Document.create();
    let count = 0;
    doc.on('paragraphAdded', () => {
      count++;
    });
    doc.addParagraph(new Paragraph().addText('manual'));
    expect(count).toBe(1);
    doc.dispose();
  });

  it('fires tableAdded on addTable()', () => {
    const doc = Document.create();
    let lastTable: Table | undefined;
    doc.on('tableAdded', ({ table }) => {
      lastTable = table;
    });
    const t = new Table(2, 2);
    doc.addTable(t);
    expect(lastTable).toBe(t);
    doc.dispose();
  });

  it('returns an unsubscribe function from on()', () => {
    const doc = Document.create();
    let count = 0;
    const unsub = doc.on('paragraphAdded', () => {
      count++;
    });
    doc.createParagraph('x');
    unsub();
    doc.createParagraph('y');
    expect(count).toBe(1);
    doc.dispose();
  });

  it('off() removes a listener', () => {
    const doc = Document.create();
    let count = 0;
    const handler = () => count++;
    doc.on('paragraphAdded', handler);
    doc.createParagraph('x');
    doc.off('paragraphAdded', handler);
    doc.createParagraph('y');
    expect(count).toBe(1);
    doc.dispose();
  });

  it('catches listener errors without aborting the operation', () => {
    const doc = Document.create();
    doc.on('paragraphAdded', () => {
      throw new Error('listener fail');
    });
    // Must not throw out of createParagraph().
    const para = doc.createParagraph('survive');
    expect(para.getText()).toBe('survive');
    doc.dispose();
  });

  it('fires beforeSave + afterSave on toBuffer()', async () => {
    const doc = Document.create();
    doc.createParagraph('hello');
    const events: string[] = [];
    doc.on('beforeSave', () => events.push('before'));
    doc.on('afterSave', () => events.push('after'));
    await doc.toBuffer();
    expect(events).toEqual(['before', 'after']);
    doc.dispose();
  });

  it('exposes a working afterLoad listener API on a loaded document', async () => {
    // afterLoad fires inside the static factory before any caller can attach
    // a listener — test instead that the listener-count API is consistent on
    // a loaded document and that further listeners attach without errors.
    const seed = Document.create();
    seed.createParagraph('seed');
    const buf = await seed.toBuffer();
    seed.dispose();

    const doc = await Document.loadFromBuffer(buf);
    expect(doc.listenerCount('afterLoad')).toBe(0);
    const unsub = doc.on('afterLoad', () => {});
    expect(doc.listenerCount('afterLoad')).toBe(1);
    unsub();
    doc.dispose();
  });

  it('fires paragraphRemoved when removeParagraph splices a paragraph', () => {
    const doc = Document.create();
    const para = doc.createParagraph('to be removed');
    const seen: Paragraph[] = [];
    doc.on('paragraphRemoved', ({ paragraph }) => seen.push(paragraph));
    expect(doc.removeParagraph(para)).toBe(true);
    expect(seen).toEqual([para]);
    doc.dispose();
  });

  it('fires tableAdded / tableRemoved on full table lifecycle', () => {
    const doc = Document.create();
    const added: Table[] = [];
    const removed: Table[] = [];
    doc.on('tableAdded', ({ table }) => added.push(table));
    doc.on('tableRemoved', ({ table }) => removed.push(table));
    const t = doc.createTable(2, 2);
    expect(added).toEqual([t]);
    expect(doc.removeTable(t)).toBe(true);
    expect(removed).toEqual([t]);
    doc.dispose();
  });

  it('fires paragraphAdded on insertParagraphAt and addBodyElement', () => {
    const doc = Document.create();
    const seen: string[] = [];
    doc.on('paragraphAdded', ({ paragraph }) => seen.push(paragraph.getText()));
    doc.insertParagraphAt(0, new Paragraph().addText('inserted'));
    doc.addBodyElement(new Paragraph().addText('via-addBody'));
    expect(seen).toEqual(['inserted', 'via-addBody']);
    doc.dispose();
  });

  it('fires removed + added on replaceElement', () => {
    const doc = Document.create();
    const old = doc.createParagraph('old');
    const events: string[] = [];
    doc.on('paragraphRemoved', () => events.push('removed'));
    doc.on('paragraphAdded', () => events.push('added'));
    expect(doc.replaceElement(old, new Paragraph().addText('new'))).toBe(true);
    // replace fires removed first, then added — order matters for audit logs.
    expect(events).toEqual(['removed', 'added']);
    doc.dispose();
  });

  it('disposes listeners on dispose()', () => {
    const doc = Document.create();
    doc.on('paragraphAdded', () => {});
    doc.on('paragraphAdded', () => {});
    expect(doc.listenerCount('paragraphAdded')).toBe(2);
    doc.dispose();
    expect(doc.listenerCount('paragraphAdded')).toBe(0);
  });
});

describe('Document.setDefaults', () => {
  afterEach(() => {
    Document.resetDefaults();
  });

  it('applies default font to new paragraphs', () => {
    Document.setDefaults({ font: 'Verdana', fontSize: 12 });
    const doc = Document.create();
    const para = doc.createParagraph('default font');
    const run = para.getRuns()[0]!;
    expect(run.getFormatting().font).toBe('Verdana');
    expect(run.getFormatting().size).toBe(12);
    doc.dispose();
  });

  it('does not override explicitly-set run formatting', () => {
    Document.setDefaults({ font: 'Verdana' });
    const doc = Document.create();
    const para = new Paragraph();
    para.addText('manual');
    para.getRuns()[0]!.setFont('Arial');
    doc.addParagraph(para);
    expect(para.getRuns()[0]!.getFormatting().font).toBe('Arial');
    doc.dispose();
  });

  it('returns a read-only copy via getDefaults()', () => {
    Document.setDefaults({ font: 'Verdana' });
    const d = Document.getDefaults();
    expect(d.font).toBe('Verdana');
    // Mutating the returned object should not affect future calls.
    (d as { font?: string }).font = 'Hacked';
    expect(Document.getDefaults().font).toBe('Verdana');
  });

  it('resetDefaults() clears all entries', () => {
    Document.setDefaults({ font: 'Verdana', fontSize: 14, fontColor: 'FF0000' });
    Document.resetDefaults();
    expect(Document.getDefaults()).toEqual({});
  });

  it('does not affect paragraphs created without text content', () => {
    Document.setDefaults({ font: 'Verdana' });
    const doc = Document.create();
    const para = doc.createParagraph(); // empty — no runs to format
    expect(para.getRuns().length).toBe(0);
    doc.dispose();
  });
});
