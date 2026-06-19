/**
 * Empty-table cleanup after revision acceptance must not delete tables whose
 * content contributes no paragraph text: images (figure/logo layout tables),
 * fields, and nested tables stored as raw XML passthrough. Paragraph.getText()
 * only surfaces Run/Hyperlink text, so a text-only emptiness check silently
 * removes such tables in the default acceptAllRevisions() flow.
 */
import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { Image } from '../../src/elements/Image';
import { ImageRun } from '../../src/elements/ImageRun';
import { Field } from '../../src/elements/Field';
import { acceptRevisionsInMemory } from '../../src/processors/InMemoryRevisionAcceptor';

// 1x1 transparent PNG
const PNG_1X1 = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const NESTED_TABLE_XML =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Nested content</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

describe('C6: empty-table cleanup preserves non-text content', () => {
  it('does not remove a table whose only content is an image', async () => {
    const doc = Document.create();
    try {
      const image = await Image.fromBuffer(PNG_1X1, 'png', 914400, 914400);
      const table = doc.createTable(1, 1);
      table.getRows()[0]?.getCells()[0]?.createParagraph().addRun(new ImageRun(image));

      const result = acceptRevisionsInMemory(doc, { cleanupEmptyTables: true });

      expect(result.emptyTablesRemoved).toBe(0);
      expect(doc.getTables().length).toBe(1);
    } finally {
      doc.dispose();
    }
  });

  it('does not remove a table whose cell content lives in raw nested content', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(1, 1);
      table.getRows()[0]?.getCells()[0]?.addRawNestedContent(0, NESTED_TABLE_XML, 'table');

      const result = acceptRevisionsInMemory(doc, { cleanupEmptyTables: true });

      expect(result.emptyTablesRemoved).toBe(0);
      expect(doc.getTables().length).toBe(1);
    } finally {
      doc.dispose();
    }
  });

  it('does not remove a table whose only content is a field', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(1, 1);
      table
        .getRows()[0]
        ?.getCells()[0]
        ?.createParagraph()
        .addField(new Field({ type: 'PAGE' }));

      const result = acceptRevisionsInMemory(doc, { cleanupEmptyTables: true });

      expect(result.emptyTablesRemoved).toBe(0);
      expect(doc.getTables().length).toBe(1);
    } finally {
      doc.dispose();
    }
  });

  it('preserves an image-only table through the default acceptAllRevisions() flow', async () => {
    const doc = Document.create();
    try {
      const image = await Image.fromBuffer(PNG_1X1, 'png', 914400, 914400);
      const table = doc.createTable(1, 1);
      table.getRows()[0]?.getCells()[0]?.createParagraph().addRun(new ImageRun(image));

      await doc.acceptAllRevisions();

      expect(doc.getTables().length).toBe(1);
    } finally {
      doc.dispose();
    }
  });

  it('still removes a table emptied by accepted tracked deletions', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(1, 1);
      const para = table.getRows()[0]?.getCells()[0]?.createParagraph();
      para?.addRevision(Revision.createDeletion('Author', new Run('Deleted content')));

      const result = acceptRevisionsInMemory(doc, { cleanupEmptyTables: true });

      expect(result.emptyTablesRemoved).toBe(1);
      expect(doc.getTables().length).toBe(0);
    } finally {
      doc.dispose();
    }
  });
});
