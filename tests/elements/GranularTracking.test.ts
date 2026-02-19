/**
 * Integration tests for character-level granular tracked changes.
 *
 * Verifies that Run.setText() with tracking enabled produces fine-grained
 * diff-based revisions instead of whole-run delete+insert.
 */

import { Document } from "../../src/core/Document";
import { Run } from "../../src/elements/Run";
import { Revision } from "../../src/elements/Revision";

describe("Granular Tracking", () => {
  describe("Run.setText() fine-grained tracking", () => {
    it("should produce granular revisions for space removal", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("word  word");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("word word");

      // Should have: equalRun("word ") + deleteRev(" ") + equalRun("word")
      const content = para.getContent();
      expect(content.length).toBe(3);

      // First: unchanged run "word "
      expect(content[0]).toBeInstanceOf(Run);
      expect((content[0] as Run).getText()).toBe("word ");

      // Second: delete revision for the extra space
      expect(content[1]).toBeInstanceOf(Revision);
      const delRev = content[1] as Revision;
      expect(delRev.getType()).toBe("delete");
      expect(delRev.getRuns()[0]!.getText()).toBe(" ");

      // Third: unchanged run "word"
      expect(content[2]).toBeInstanceOf(Run);
      expect((content[2] as Run).getText()).toBe("word");
    });

    it("should produce granular revisions for word replacement", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("The quick fox");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("The slow fox");

      const content = para.getContent();
      // "The " (equal) + "quick" (delete) + "slow" (insert) + " fox" (equal) = 4 items
      expect(content.length).toBe(4);

      expect(content[0]).toBeInstanceOf(Run);
      expect((content[0] as Run).getText()).toBe("The ");

      expect(content[1]).toBeInstanceOf(Revision);
      expect((content[1] as Revision).getType()).toBe("delete");
      expect((content[1] as Revision).getRuns()[0]!.getText()).toBe("quick");

      expect(content[2]).toBeInstanceOf(Revision);
      expect((content[2] as Revision).getType()).toBe("insert");
      expect((content[2] as Revision).getRuns()[0]!.getText()).toBe("slow");

      expect(content[3]).toBeInstanceOf(Run);
      expect((content[3] as Run).getText()).toBe(" fox");
    });

    it("should fall back to whole-run replacement when no common text", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("abc");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("xyz");

      // Should fall back to whole-run: delete("abc") + insert("xyz")
      const content = para.getContent();
      expect(content.length).toBe(2);
      expect(content[0]).toBeInstanceOf(Revision);
      expect(content[1]).toBeInstanceOf(Revision);

      expect((content[0] as Revision).getType()).toBe("delete");
      expect((content[0] as Revision).getRuns()[0]!.getText()).toBe("abc");

      expect((content[1] as Revision).getType()).toBe("insert");
      expect((content[1] as Revision).getRuns()[0]!.getText()).toBe("xyz");
    });

    it("should preserve formatting on all split segments", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("Hello World");
      const run = para.getRuns()[0]!;
      run.setBold(true);
      run.setFont("Arial", 12);
      run.setColor("FF0000");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("Hello Earth");

      const content = para.getContent();
      // "Hello " (equal) + "World" (delete) + "Earth" (insert) = 3 items
      expect(content.length).toBe(3);

      // Check unchanged run has same formatting
      const equalRun = content[0] as Run;
      expect(equalRun.getBold()).toBe(true);
      expect(equalRun.getFormatting().font).toBe("Arial");
      expect(equalRun.getFormatting().size).toBe(12);
      expect(equalRun.getFormatting().color).toBe("FF0000");

      // Check delete revision's run has same formatting
      const delRev = content[1] as Revision;
      const delRun = delRev.getRuns()[0]!;
      expect(delRun.getBold()).toBe(true);
      expect(delRun.getFormatting().font).toBe("Arial");

      // Check insert revision's run has same formatting
      const insRev = content[2] as Revision;
      const insRun = insRev.getRuns()[0]!;
      expect(insRun.getBold()).toBe(true);
      expect(insRun.getFormatting().font).toBe("Arial");
    });

    it("should propagate tracking context to unchanged runs", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("Hello World");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("Hello Earth");

      const content = para.getContent();
      // First item is the unchanged "Hello " run
      const equalRun = content[0] as Run;
      expect(equalRun).toBeInstanceOf(Run);
      expect(equalRun._getParentParagraph()).toBe(para);

      // The unchanged run should also support tracking for future edits
      // Let's edit it and see if it creates revisions
      equalRun.setText("Bye ");

      // After editing the equal run, it should create its own tracked changes
      // "Hello " → "Bye " diff: delete "Hello" + insert "Bye" + equal " "
      // So paragraph content becomes:
      //   [delete "Hello", insert "Bye", equal " ", delete "World", insert "Earth"]
      const newContent = para.getContent();
      expect(newContent.length).toBe(5);

      // First three items replace the original "Hello " equal run
      expect(newContent[0]).toBeInstanceOf(Revision);
      expect((newContent[0] as Revision).getType()).toBe("delete");
      expect((newContent[0] as Revision).getRuns()[0]!.getText()).toBe("Hello");

      expect(newContent[1]).toBeInstanceOf(Revision);
      expect((newContent[1] as Revision).getType()).toBe("insert");
      expect((newContent[1] as Revision).getRuns()[0]!.getText()).toBe("Bye");

      expect(newContent[2]).toBeInstanceOf(Run);
      expect((newContent[2] as Run).getText()).toBe(" ");
    });

    it("should handle suffix-only change", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("Hello World");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("Goodbye World");

      const content = para.getContent();
      // "Hello" (delete) + "Goodbye" (insert) + " World" (equal) = 3 items
      expect(content.length).toBe(3);

      expect(content[0]).toBeInstanceOf(Revision);
      expect((content[0] as Revision).getType()).toBe("delete");
      expect((content[0] as Revision).getRuns()[0]!.getText()).toBe("Hello");

      expect(content[1]).toBeInstanceOf(Revision);
      expect((content[1] as Revision).getType()).toBe("insert");
      expect((content[1] as Revision).getRuns()[0]!.getText()).toBe("Goodbye");

      expect(content[2]).toBeInstanceOf(Run);
      expect((content[2] as Run).getText()).toBe(" World");
    });

    it("should handle insertion with no deletion (text appended)", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("hello");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("hello world");

      const content = para.getContent();
      // "hello" (equal) + " world" (insert) = 2 items
      expect(content.length).toBe(2);

      expect(content[0]).toBeInstanceOf(Run);
      expect((content[0] as Run).getText()).toBe("hello");

      expect(content[1]).toBeInstanceOf(Revision);
      expect((content[1] as Revision).getType()).toBe("insert");
      expect((content[1] as Revision).getRuns()[0]!.getText()).toBe(" world");
    });

    it("should handle deletion with no insertion (text removed from end)", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("hello world");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("hello");

      const content = para.getContent();
      // "hello" (equal) + " world" (delete) = 2 items
      expect(content.length).toBe(2);

      expect(content[0]).toBeInstanceOf(Run);
      expect((content[0] as Run).getText()).toBe("hello");

      expect(content[1]).toBeInstanceOf(Revision);
      expect((content[1] as Revision).getType()).toBe("delete");
      expect((content[1] as Revision).getRuns()[0]!.getText()).toBe(" world");
    });

    it("should register revisions with RevisionManager", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("word  word");

      doc.enableTrackChanges({ author: "Editor" });

      const revManager = doc.getRevisionManager();
      const beforeCount = revManager.getAllRevisions().length;

      const runs = para.getRuns();
      runs[0]!.setText("word word");

      // Should have registered exactly 1 revision (delete for the extra space)
      const afterCount = revManager.getAllRevisions().length;
      expect(afterCount - beforeCount).toBe(1);
      expect(revManager.getDeletionCount()).toBe(1);
    });
  });

  describe("findAndReplaceAll with tracking", () => {
    it("should not produce duplicate revisions when global tracking is enabled", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("The quick brown fox");

      // Enable global tracking
      doc.enableTrackChanges({ author: "Editor" });

      const revManager = doc.getRevisionManager();
      const beforeCount = revManager.getAllRevisions().length;

      const result = doc.findAndReplaceAll("quick", "slow", {
        trackChanges: true,
        author: "Editor",
      });

      expect(result.count).toBe(1);

      // The key assertion: no duplicate revisions
      const afterCount = revManager.getAllRevisions().length;
      const newRevisions = afterCount - beforeCount;

      // Should have created revisions from setText (fine-grained),
      // NOT duplicate revisions from manual creation + setText
      // With "The quick brown fox" → "The slow brown fox":
      // delete "quick" + insert "slow" = 2 revisions
      expect(newRevisions).toBe(2);

      // The returned revisions should match what's in the manager
      expect(result.revisions).toBeDefined();
      expect(result.revisions!.length).toBe(2);
    });

    it("should create revisions via options when global tracking is disabled", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("old value");

      // Global tracking is NOT enabled

      const result = doc.findAndReplaceAll("old", "new", {
        trackChanges: true,
        author: "Editor",
      });

      expect(result.count).toBe(1);
      expect(result.revisions).toBeDefined();
      expect(result.revisions!.length).toBe(2); // delete + insert

      // With option-level tracking, revisions are now embedded in the paragraph.
      // "old value" → "new value" diff: delete "old" + insert "new" + equal " value"
      const content = para.getContent();
      expect(content.length).toBe(3);

      expect(content[0]).toBeInstanceOf(Revision);
      expect((content[0] as Revision).getType()).toBe("delete");
      expect((content[0] as Revision).getRuns()[0]!.getText()).toBe("old");

      expect(content[1]).toBeInstanceOf(Revision);
      expect((content[1] as Revision).getType()).toBe("insert");
      expect((content[1] as Revision).getRuns()[0]!.getText()).toBe("new");

      expect(content[2]).toBeInstanceOf(Run);
      expect((content[2] as Run).getText()).toBe(" value");
    });

    it("should produce fine-grained revisions through findAndReplaceAll", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("Hello World");

      doc.enableTrackChanges({ author: "Editor" });

      doc.findAndReplaceAll("World", "Earth", {
        trackChanges: true,
        author: "Editor",
      });

      // The paragraph content should show fine-grained tracking:
      // "Hello " (unchanged) + "World" (deleted) + "Earth" (inserted)
      const content = para.getContent();
      expect(content.length).toBe(3);

      expect(content[0]).toBeInstanceOf(Run);
      expect((content[0] as Run).getText()).toBe("Hello ");

      expect(content[1]).toBeInstanceOf(Revision);
      expect((content[1] as Revision).getType()).toBe("delete");

      expect(content[2]).toBeInstanceOf(Revision);
      expect((content[2] as Revision).getType()).toBe("insert");
    });
  });

  describe("replaceText with tracking", () => {
    it("should produce fine-grained revisions when global tracking is enabled", () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("Hello World");

      doc.enableTrackChanges({ author: "Editor" });

      doc.replaceText("World", "Earth");

      const content = para.getContent();
      // Should have fine-grained tracking via setText():
      // "Hello " (equal) + "World" (delete) + "Earth" (insert) = 3 items
      expect(content.length).toBe(3);

      expect(content[0]).toBeInstanceOf(Run);
      expect((content[0] as Run).getText()).toBe("Hello ");

      expect(content[1]).toBeInstanceOf(Revision);
      expect((content[1] as Revision).getType()).toBe("delete");

      expect(content[2]).toBeInstanceOf(Revision);
      expect((content[2] as Revision).getType()).toBe("insert");
    });
  });

  describe("DOCX output verification", () => {
    it("should produce valid DOCX with granular tracked changes", async () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("word  word");

      doc.enableTrackChanges({ author: "Editor" });

      const runs = para.getRuns();
      runs[0]!.setText("word word");

      // Should produce valid buffer without errors
      const buffer = await doc.toBuffer();
      expect(buffer).toBeDefined();
      expect(buffer.length).toBeGreaterThan(0);

      // Load the document back to verify structure
      const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      expect(loaded).toBeDefined();

      // Check that the document has paragraphs
      const paragraphs = loaded.getAllParagraphs();
      expect(paragraphs.length).toBeGreaterThan(0);

      loaded.dispose();
    });

    it("should produce valid DOCX with word replacement tracking", async () => {
      const doc = new Document();
      const para = doc.createParagraph();
      para.addText("The quick brown fox jumps over the lazy dog");

      doc.enableTrackChanges({ author: "Editor" });

      doc.replaceText("quick", "slow");
      doc.replaceText("lazy", "energetic");

      const buffer = await doc.toBuffer();
      expect(buffer).toBeDefined();
      expect(buffer.length).toBeGreaterThan(0);

      const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      expect(loaded).toBeDefined();
      loaded.dispose();
    });
  });
});
