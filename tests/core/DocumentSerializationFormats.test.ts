/**
 * Tests for Document.toBase64(), Document.toDataUri(), Document.loadFromBase64()
 */

import { Document } from '../../src/core/Document';

describe('Document.toBase64()', () => {
  it('returns a non-empty base64 string', async () => {
    const doc = Document.create();
    doc.createParagraph('Hello');

    const base64 = await doc.toBase64();

    expect(typeof base64).toBe('string');
    expect(base64.length).toBeGreaterThan(0);
    doc.dispose();
  });

  it('produces valid base64 that decodes to a DOCX buffer', async () => {
    const doc = Document.create();
    doc.createParagraph('Test content');

    const base64 = await doc.toBase64();
    const buffer = Buffer.from(base64, 'base64');

    // DOCX files are ZIP archives starting with PK signature
    expect(buffer[0]).toBe(0x50); // 'P'
    expect(buffer[1]).toBe(0x4b); // 'K'
    doc.dispose();
  });

  it('decodes to a valid DOCX with same content as toBuffer', async () => {
    const doc = Document.create();
    doc.createParagraph('Comparison test');

    const base64 = await doc.toBase64();
    const decoded = Buffer.from(base64, 'base64');

    // Both should be valid ZIPs with PK signature
    expect(decoded[0]).toBe(0x50);
    expect(decoded[1]).toBe(0x4b);
    // Size should be in the same ballpark
    const buffer = await doc.toBuffer();
    expect(Math.abs(decoded.length - buffer.length)).toBeLessThan(100);
    doc.dispose();
  });

  it('handles empty document', async () => {
    const doc = Document.create();

    const base64 = await doc.toBase64();

    expect(base64.length).toBeGreaterThan(0); // Still has ZIP structure
    doc.dispose();
  });

  it('handles document with tables and formatting', async () => {
    const doc = Document.create();
    doc.addHeading('Report', 1);
    doc.createParagraph('Content with **bold**.');
    doc.addTable(
      (await import('../../src/elements/Table')).Table.fromArray([
        ['A', 'B'],
        ['1', '2'],
      ])
    );

    const base64 = await doc.toBase64();

    expect(base64.length).toBeGreaterThan(100);
    doc.dispose();
  });
});

describe('Document.toDataUri()', () => {
  it('returns a valid data URI with correct MIME type', async () => {
    const doc = Document.create();
    doc.createParagraph('Hello');

    const uri = await doc.toDataUri();

    expect(uri).toMatch(
      /^data:application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document;base64,/
    );
    doc.dispose();
  });

  it('contains base64 content after the header', async () => {
    const doc = Document.create();
    doc.createParagraph('Content');

    const uri = await doc.toDataUri();
    const base64Part = uri.split(',')[1]!;

    // Should be valid base64
    expect(base64Part.length).toBeGreaterThan(0);
    const buffer = Buffer.from(base64Part, 'base64');
    expect(buffer[0]).toBe(0x50); // 'P' - ZIP signature
    doc.dispose();
  });

  it('has correct structure: MIME type prefix + base64 content', async () => {
    const doc = Document.create();
    doc.createParagraph('Structure test');

    const uri = await doc.toDataUri();
    const prefix =
      'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,';

    expect(uri.startsWith(prefix)).toBe(true);

    // The base64 part should decode to a valid ZIP
    const base64Part = uri.slice(prefix.length);
    const buffer = Buffer.from(base64Part, 'base64');
    expect(buffer[0]).toBe(0x50); // 'P'
    expect(buffer[1]).toBe(0x4b); // 'K'
    doc.dispose();
  });
});

describe('Document.loadFromBase64()', () => {
  it('loads a document from base64 string', async () => {
    const original = Document.create();
    original.createParagraph('Round trip content');

    const base64 = await original.toBase64();
    original.dispose();

    const loaded = await Document.loadFromBase64(base64);

    expect(loaded.toPlainText()).toContain('Round trip content');
    loaded.dispose();
  });

  it('preserves document structure through base64 round-trip', async () => {
    const original = Document.create();
    original.addHeading('Title', 1);
    original.createParagraph('Introduction text.');
    original.addHeading('Section', 2);
    original.createParagraph('Section content.');

    const base64 = await original.toBase64();
    original.dispose();

    const loaded = await Document.loadFromBase64(base64);

    expect(loaded.getStatistics().headings).toBe(2);
    expect(loaded.toPlainText()).toContain('Introduction text.');
    expect(loaded.toPlainText()).toContain('Section content.');
    loaded.dispose();
  });

  it('preserves tables through base64 round-trip', async () => {
    const original = Document.create();
    const { Table } = await import('../../src/elements/Table');
    original.addTable(
      Table.fromArray([
        ['Name', 'Value'],
        ['Alpha', '100'],
      ])
    );

    const base64 = await original.toBase64();
    original.dispose();

    const loaded = await Document.loadFromBase64(base64);

    expect(loaded.getTables()).toHaveLength(1);
    loaded.dispose();
  });

  it('is the inverse of toBase64', async () => {
    const doc = Document.create();
    doc.createParagraph('Invertibility test');
    doc.setDefaultFont('Georgia', 12);

    const base64 = await doc.toBase64();
    doc.dispose();

    const restored = await Document.loadFromBase64(base64);
    const reEncoded = await restored.toBase64();
    restored.dispose();

    // Re-encoding should produce a valid document
    const final = await Document.loadFromBase64(reEncoded);
    expect(final.toPlainText()).toContain('Invertibility test');
    final.dispose();
  });
});

describe('serialization format interop', () => {
  it('toBuffer → base64 → loadFromBase64 round-trip', async () => {
    const doc = Document.create();
    doc.createParagraph('Interop test');

    const buffer = await doc.toBuffer();
    doc.dispose();

    const base64 = buffer.toString('base64');
    const loaded = await Document.loadFromBase64(base64);

    expect(loaded.toPlainText()).toContain('Interop test');
    loaded.dispose();
  });

  it('loadFromBase64 → toDataUri produces valid URI', async () => {
    const doc = Document.create();
    doc.createParagraph('URI test');

    const base64 = await doc.toBase64();
    doc.dispose();

    const loaded = await Document.loadFromBase64(base64);
    const uri = await loaded.toDataUri();

    expect(uri).toMatch(/^data:application/);
    loaded.dispose();
  });
});
