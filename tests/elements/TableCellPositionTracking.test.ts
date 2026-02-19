/**
 * Tests for TableCell position tracking and trailing blank removal
 * Critical bug fix: rawNestedContent positions must update when paragraphs are added/removed
 */

import { TableCell } from "../../src/elements/TableCell";
import { Paragraph } from "../../src/elements/Paragraph";
import { Revision } from "../../src/elements/Revision";
import { ImageRun } from "../../src/elements/ImageRun";
import { Image } from "../../src/elements/Image";

describe("TableCell Position Tracking", () => {
  describe("removeParagraph position updates", () => {
    it("should decrement rawNestedContent positions when paragraph is removed before nested content", () => {
      const cell = new TableCell();

      // Add paragraphs
      cell.createParagraph("Para 0");
      cell.createParagraph("Para 1 - blank");
      cell.createParagraph("Para 2");

      // Simulate nested content at position 2 (after Para 1)
      cell.addRawNestedContent(2, "<w:tbl>nested table</w:tbl>", "table");

      // Verify initial position
      const initialContent = cell.getRawNestedContent();
      expect(initialContent[0]?.position).toBe(2);

      // Remove paragraph at index 1 (before nested content)
      cell.removeParagraph(1);

      // Position should decrement from 2 to 1
      const updatedContent = cell.getRawNestedContent();
      expect(updatedContent[0]?.position).toBe(1);
    });

    it("should NOT change rawNestedContent position when paragraph is removed after nested content", () => {
      const cell = new TableCell();

      // Add paragraphs
      cell.createParagraph("Para 0");
      cell.createParagraph("Para 1");
      cell.createParagraph("Para 2 - to remove");

      // Simulate nested content at position 1 (after Para 0)
      cell.addRawNestedContent(1, "<w:tbl>nested table</w:tbl>", "table");

      // Verify initial position
      expect(cell.getRawNestedContent()[0]?.position).toBe(1);

      // Remove paragraph at index 2 (after nested content position)
      cell.removeParagraph(2);

      // Position should remain 1
      expect(cell.getRawNestedContent()[0]?.position).toBe(1);
    });

    it("should handle multiple rawNestedContent items correctly", () => {
      const cell = new TableCell();

      // Add paragraphs
      cell.createParagraph("Para 0");
      cell.createParagraph("Para 1 - blank");
      cell.createParagraph("Para 2");
      cell.createParagraph("Para 3");

      // Add nested content at different positions
      cell.addRawNestedContent(1, "<w:tbl>table1</w:tbl>", "table");
      cell.addRawNestedContent(3, "<w:tbl>table2</w:tbl>", "table");

      // Verify initial positions
      const content = cell.getRawNestedContent();
      expect(content[0]?.position).toBe(1);
      expect(content[1]?.position).toBe(3);

      // Remove paragraph at index 1
      cell.removeParagraph(1);

      // First item position should stay 1 (removed at same position)
      // Second item position should decrement from 3 to 2
      const updated = cell.getRawNestedContent();
      expect(updated[0]?.position).toBe(1);
      expect(updated[1]?.position).toBe(2);
    });
  });

  describe("addParagraphAt position updates", () => {
    it("should increment rawNestedContent positions when paragraph is inserted before nested content", () => {
      const cell = new TableCell();

      // Add paragraphs
      cell.createParagraph("Para 0");
      cell.createParagraph("Para 1");

      // Simulate nested content at position 1 (after Para 0)
      cell.addRawNestedContent(1, "<w:tbl>nested table</w:tbl>", "table");

      // Verify initial position
      expect(cell.getRawNestedContent()[0]?.position).toBe(1);

      // Insert new paragraph at index 0
      const newPara = new Paragraph();
      newPara.addText("Inserted Para");
      cell.addParagraphAt(0, newPara);

      // Position should increment from 1 to 2
      expect(cell.getRawNestedContent()[0]?.position).toBe(2);
    });

    it("should NOT change rawNestedContent position when paragraph is added after nested content", () => {
      const cell = new TableCell();

      // Add paragraphs
      cell.createParagraph("Para 0");
      cell.createParagraph("Para 1");

      // Simulate nested content at position 1 (after Para 0)
      cell.addRawNestedContent(1, "<w:tbl>nested table</w:tbl>", "table");

      // Verify initial position
      expect(cell.getRawNestedContent()[0]?.position).toBe(1);

      // Insert new paragraph at end
      const newPara = new Paragraph();
      newPara.addText("Appended Para");
      cell.addParagraphAt(10, newPara); // Index beyond array pushes to end

      // Position should remain 1
      expect(cell.getRawNestedContent()[0]?.position).toBe(1);
    });
  });

  describe("removeTrailingBlankParagraphs", () => {
    it("should remove trailing blank paragraphs from cell", () => {
      const cell = new TableCell();

      cell.createParagraph("Content");
      cell.createParagraph(""); // blank
      cell.createParagraph(""); // blank

      expect(cell.getParagraphs().length).toBe(3);

      const removed = cell.removeTrailingBlankParagraphs();

      expect(removed).toBe(2);
      expect(cell.getParagraphs().length).toBe(1);
      expect(cell.getParagraphs()[0]?.getText()).toBe("Content");
    });

    it("should keep at least one paragraph in cell", () => {
      const cell = new TableCell();

      cell.createParagraph(""); // blank
      cell.createParagraph(""); // blank

      expect(cell.getParagraphs().length).toBe(2);

      const removed = cell.removeTrailingBlankParagraphs();

      // Should only remove 1, keeping at least one paragraph
      expect(removed).toBe(1);
      expect(cell.getParagraphs().length).toBe(1);
    });

    it("should stop at non-blank paragraph", () => {
      const cell = new TableCell();

      cell.createParagraph("Content 1");
      cell.createParagraph("Content 2");
      cell.createParagraph(""); // blank at end

      const removed = cell.removeTrailingBlankParagraphs();

      expect(removed).toBe(1);
      expect(cell.getParagraphs().length).toBe(2);
    });

    it("should respect preserve flag when ignorePreserveFlag is false", () => {
      const cell = new TableCell();

      cell.createParagraph("Content");
      const blankPara = cell.createParagraph("");
      blankPara.setPreserved(true);

      const removed = cell.removeTrailingBlankParagraphs({
        ignorePreserveFlag: false,
      });

      expect(removed).toBe(0);
      expect(cell.getParagraphs().length).toBe(2);
    });

    it("should ignore preserve flag when ignorePreserveFlag is true", () => {
      const cell = new TableCell();

      cell.createParagraph("Content");
      const blankPara = cell.createParagraph("");
      blankPara.setPreserved(true);

      const removed = cell.removeTrailingBlankParagraphs({
        ignorePreserveFlag: true,
      });

      expect(removed).toBe(1);
      expect(cell.getParagraphs().length).toBe(1);
    });

    it("should NOT remove paragraph with revision containing ImageRun", async () => {
      const cell = new TableCell();

      cell.createParagraph("Content");

      // Create a paragraph that only contains a revision with an ImageRun
      const imagePara = new Paragraph();
      // 1x1 transparent PNG
      const imageBuffer = Buffer.from([
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82,
      ]);
      const image = await Image.fromBuffer(imageBuffer, 'png', 914400, 914400);
      const imageRun = new ImageRun(image);
      const revision = new Revision({
        id: 1,
        author: "Test Author",
        date: new Date(),
        type: "insert",
        content: [imageRun],
      });
      imagePara.addRevision(revision);
      cell.addParagraph(imagePara);

      expect(cell.getParagraphs().length).toBe(2);

      const removed = cell.removeTrailingBlankParagraphs({ ignorePreserveFlag: true });

      // Should NOT remove the paragraph because it contains a revision with an image
      expect(removed).toBe(0);
      expect(cell.getParagraphs().length).toBe(2);
    });

    it("should NOT remove paragraph with both insert and delete revisions containing images", async () => {
      const cell = new TableCell();

      cell.createParagraph("Content");

      // Simulate the real-world pattern: deleted old image + inserted new image
      const imagePara = new Paragraph();
      const imageBuffer = Buffer.from([
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82,
      ]);
      const newImage = await Image.fromBuffer(imageBuffer, 'png', 914400, 914400);
      const oldImage = await Image.fromBuffer(imageBuffer, 'png', 914400, 914400);

      // w:ins with new image
      imagePara.addRevision(new Revision({
        id: 1,
        author: "Test Author",
        date: new Date(),
        type: "insert",
        content: [new ImageRun(newImage)],
      }));
      // w:del with old image
      imagePara.addRevision(new Revision({
        id: 2,
        author: "Test Author",
        date: new Date(),
        type: "delete",
        content: [new ImageRun(oldImage)],
      }));

      cell.addParagraph(imagePara);

      expect(cell.getParagraphs().length).toBe(2);

      const removed = cell.removeTrailingBlankParagraphs({ ignorePreserveFlag: true });

      expect(removed).toBe(0);
      expect(cell.getParagraphs().length).toBe(2);
    });

    it("should NOT remove blank if raw nested content is positioned after it", () => {
      const cell = new TableCell();

      cell.createParagraph("Content");
      cell.createParagraph(""); // blank

      // Nested content positioned at the end (after the blank)
      cell.addRawNestedContent(2, "<w:tbl>nested</w:tbl>", "table");

      const removed = cell.removeTrailingBlankParagraphs();

      // Should NOT remove the blank because nested content depends on it
      expect(removed).toBe(0);
      expect(cell.getParagraphs().length).toBe(2);
    });
  });
});
