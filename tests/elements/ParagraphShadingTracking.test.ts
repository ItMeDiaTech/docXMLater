/**
 * ParagraphShadingTracking - Tests for paragraph shading change tracking
 */

import { Document } from "../../src/core/Document";
import { Paragraph } from "../../src/elements/Paragraph";

describe("Paragraph Shading Tracking", () => {
  let doc: Document;

  beforeEach(() => {
    doc = Document.create();
  });

  afterEach(() => {
    doc.dispose();
  });

  it("should track shading change when tracking is enabled", () => {
    const para = new Paragraph();
    para.addText("Test paragraph");
    doc.addParagraph(para);
    doc.enableTrackChanges({ author: "TestAuthor" });

    para.setShading({ fill: "FFFF00", pattern: "solid" });
    doc.flushPendingChanges();

    const formatting = para.getFormatting();
    expect(formatting.pPrChange).toBeDefined();
    expect(formatting.pPrChange!.author).toBe("TestAuthor");
  });

  it("should capture previous shading value", () => {
    const para = new Paragraph();
    para.addText("Test paragraph");
    para.setShading({ fill: "FF0000", pattern: "clear" });
    doc.addParagraph(para);
    doc.enableTrackChanges({ author: "TestAuthor" });

    para.setShading({ fill: "00FF00", pattern: "solid" });
    doc.flushPendingChanges();

    const formatting = para.getFormatting();
    expect(formatting.pPrChange).toBeDefined();
    // Current should be the new value
    expect(formatting.shading!.fill).toBe("00FF00");
    expect(formatting.shading!.pattern).toBe("solid");
  });

  it("should not create pPrChange when tracking is disabled", () => {
    const para = new Paragraph();
    para.addText("Test paragraph");
    doc.addParagraph(para);

    para.setShading({ fill: "FFFF00", pattern: "solid" });

    const formatting = para.getFormatting();
    expect(formatting.pPrChange).toBeUndefined();
  });

  it("should track theme shading changes", () => {
    const para = new Paragraph();
    para.addText("Test paragraph");
    doc.addParagraph(para);
    doc.enableTrackChanges({ author: "TestAuthor" });

    para.setShading({
      fill: "D9E2F3",
      pattern: "clear",
      themeFill: "accent1",
      themeFillTint: "33",
    });
    doc.flushPendingChanges();

    const formatting = para.getFormatting();
    expect(formatting.pPrChange).toBeDefined();
    expect(formatting.shading!.themeFill).toBe("accent1");
    expect(formatting.shading!.themeFillTint).toBe("33");
  });
});
