/**
 * Tests for validateBookmarkPairs() — ensures balanced bookmark pairs on save
 */

import { Document } from "../../src/core/Document";
import { Paragraph } from "../../src/elements/Paragraph";
import { Bookmark } from "../../src/elements/Bookmark";
import { Table } from "../../src/elements/Table";

describe("Bookmark Pair Validation", () => {
  describe("validateBookmarkPairs()", () => {
    it("should add missing bookmarkEnd for orphaned bookmarkStart", () => {
      const doc = Document.create();

      const para1 = new Paragraph();
      para1.addText("Start paragraph");
      para1.addBookmarkStart(new Bookmark({ id: 10, name: "orphaned_start", skipNormalization: true }));
      doc.addParagraph(para1);

      const para2 = new Paragraph();
      para2.addText("End paragraph");
      doc.addParagraph(para2);

      // No bookmarkEnd for ID 10 — orphaned start
      const repairs = doc.validateBookmarkPairs();
      expect(repairs).toBe(1);

      // Verify end was added to last paragraph
      const paras = doc.getParagraphs();
      const lastPara = paras[paras.length - 1]!;
      const ends = lastPara.getBookmarksEnd();
      expect(ends.some(bm => bm.getId() === 10)).toBe(true);
    });

    it("should remove orphaned bookmarkEnd (no matching start)", () => {
      const doc = Document.create();

      const para1 = new Paragraph();
      para1.addText("Paragraph with orphaned end");
      para1.addBookmarkEnd(new Bookmark({ id: 99, name: "orphaned_end", skipNormalization: true }));
      doc.addParagraph(para1);

      const repairs = doc.validateBookmarkPairs();
      expect(repairs).toBe(1);

      // Verify end was removed
      const ends = doc.getParagraphs()[0]!.getBookmarksEnd();
      expect(ends.some(bm => bm.getId() === 99)).toBe(false);
    });

    it("should not modify balanced bookmark pairs", () => {
      const doc = Document.create();

      const para1 = new Paragraph();
      para1.addText("Bookmarked content");
      para1.addBookmarkStart(new Bookmark({ id: 5, name: "balanced", skipNormalization: true }));
      para1.addBookmarkEnd(new Bookmark({ id: 5, name: "balanced", skipNormalization: true }));
      doc.addParagraph(para1);

      const repairs = doc.validateBookmarkPairs();
      expect(repairs).toBe(0);
    });

    it("should handle multiple orphaned bookmarks", () => {
      const doc = Document.create();

      const para1 = new Paragraph();
      para1.addBookmarkStart(new Bookmark({ id: 1, name: "bm1", skipNormalization: true }));
      para1.addBookmarkStart(new Bookmark({ id: 2, name: "bm2", skipNormalization: true }));
      para1.addBookmarkEnd(new Bookmark({ id: 1, name: "bm1", skipNormalization: true }));
      // ID 2 has no end
      doc.addParagraph(para1);

      const para2 = new Paragraph();
      para2.addBookmarkEnd(new Bookmark({ id: 3, name: "bm3", skipNormalization: true }));
      // ID 3 has no start
      doc.addParagraph(para2);

      const repairs = doc.validateBookmarkPairs();
      expect(repairs).toBe(2); // 1 orphaned start + 1 orphaned end
    });

    it("should scan bookmarks inside table cells", () => {
      const doc = Document.create();

      const table = new Table(1, 1);
      const cell = table.getCell(0, 0);
      if (cell) {
        const cellPara = new Paragraph();
        cellPara.addText("Cell content");
        cellPara.addBookmarkStart(new Bookmark({ id: 20, name: "in_table", skipNormalization: true }));
        cell.addParagraph(cellPara);
        // No bookmarkEnd
      }
      doc.addTable(table);

      // Add trailing paragraph for the end to be placed
      const lastPara = new Paragraph();
      lastPara.addText("Last");
      doc.addParagraph(lastPara);

      const repairs = doc.validateBookmarkPairs();
      expect(repairs).toBe(1);

      // bookmarkEnd should be added to last paragraph
      const ends = lastPara.getBookmarksEnd();
      expect(ends.some(bm => bm.getId() === 20)).toBe(true);
    });

    it("should ensure bookmark count equality after validation", async () => {
      const doc = Document.create();

      // Create several bookmark starts without ends
      for (let i = 0; i < 5; i++) {
        const para = new Paragraph();
        para.addText(`Para ${i}`);
        para.addBookmarkStart(new Bookmark({ id: i + 100, name: `bm_${i}`, skipNormalization: true }));
        doc.addParagraph(para);
      }

      // Add one balanced pair
      const balPara = new Paragraph();
      balPara.addBookmarkStart(new Bookmark({ id: 200, name: "balanced", skipNormalization: true }));
      balPara.addBookmarkEnd(new Bookmark({ id: 200, name: "balanced", skipNormalization: true }));
      doc.addParagraph(balPara);

      doc.validateBookmarkPairs();

      // Collect all start/end IDs
      const allStarts = new Set<number>();
      const allEnds = new Set<number>();
      for (const para of doc.getParagraphs()) {
        for (const bm of para.getBookmarksStart()) allStarts.add(bm.getId());
        for (const bm of para.getBookmarksEnd()) allEnds.add(bm.getId());
      }

      expect(allStarts.size).toBe(allEnds.size);
      for (const id of allStarts) {
        expect(allEnds.has(id)).toBe(true);
      }
    });

    it("should produce valid DOCX after validation", async () => {
      const doc = Document.create();

      const para1 = new Paragraph();
      para1.addText("Content");
      para1.addBookmarkStart(new Bookmark({ id: 50, name: "test_bm", skipNormalization: true }));
      doc.addParagraph(para1);

      // No matching bookmarkEnd — validation will fix it on save
      const buffer = await doc.toBuffer();
      expect(buffer.length).toBeGreaterThan(0);

      // Re-load to verify it's valid
      const doc2 = await Document.loadFromBuffer(buffer);
      expect(doc2.getParagraphs().length).toBeGreaterThanOrEqual(1);
      doc2.dispose();
      doc.dispose();
    });
  });
});
