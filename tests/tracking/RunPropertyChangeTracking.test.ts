import { Document } from "../../src/core/Document";
import { Paragraph } from "../../src/elements/Paragraph";
import { Run } from "../../src/elements/Run";

describe("Run Property Change Tracking", () => {
  it("should create rPrChange when run font is changed with tracking enabled", () => {
    const doc = Document.create();
    const para = new Paragraph();
    doc.addParagraph(para);
    const run = new Run("Hello world");
    run.setFont("Arial");
    para.addRun(run);

    doc.enableTrackChanges({ author: "TestAuthor", trackFormatting: true });
    run.setFont("Verdana");
    doc.flushPendingChanges();

    const propChange = run.getPropertyChangeRevision();
    expect(propChange).toBeDefined();
    expect(propChange!.author).toBe("TestAuthor");
    expect(propChange!.previousProperties).toBeDefined();
    expect(propChange!.previousProperties.font).toBe("Arial");
  });

  it("should create rPrChange with multiple changed properties", () => {
    const doc = Document.create();
    const para = new Paragraph();
    doc.addParagraph(para);
    const run = new Run("Hello");
    run.setFont("Arial");
    run.setSize(10);
    run.setColor("FF0000");
    para.addRun(run);

    doc.enableTrackChanges({ author: "TestAuthor", trackFormatting: true });
    run.setFont("Verdana");
    run.setSize(12);
    run.setColor("000000");
    doc.flushPendingChanges();

    const propChange = run.getPropertyChangeRevision();
    expect(propChange).toBeDefined();
    expect(propChange!.previousProperties.font).toBe("Arial");
    expect(propChange!.previousProperties.size).toBe(10);
    expect(propChange!.previousProperties.color).toBe("FF0000");
  });

  it("should NOT create rPrChange when value does not change", () => {
    const doc = Document.create();
    const para = new Paragraph();
    doc.addParagraph(para);
    const run = new Run("Hello");
    run.setFont("Verdana");
    para.addRun(run);

    doc.enableTrackChanges({ author: "TestAuthor", trackFormatting: true });
    run.setFont("Verdana");
    doc.flushPendingChanges();

    const propChange = run.getPropertyChangeRevision();
    expect(propChange).toBeUndefined();
  });

  it("should merge with existing rPrChange on the run", () => {
    const doc = Document.create();
    const para = new Paragraph();
    doc.addParagraph(para);
    const run = new Run("Hello");
    run.setFont("Arial");
    run.setBold(true);
    para.addRun(run);

    run.setPropertyChangeRevision({
      id: 99,
      author: "OriginalAuthor",
      date: new Date("2025-01-01"),
      previousProperties: { bold: false },
    });

    doc.enableTrackChanges({
      author: "TestAuthor",
      trackFormatting: true,
      clearExistingPropertyChanges: false,
    });

    run.setFont("Verdana");
    doc.flushPendingChanges();

    const propChange = run.getPropertyChangeRevision();
    expect(propChange).toBeDefined();
    // Verify original author/date preserved during merge
    expect(propChange!.author).toBe("OriginalAuthor");
    expect(propChange!.date).toEqual(new Date("2025-01-01"));
    // Verify previous properties merged
    expect(propChange!.previousProperties.bold).toBe(false);
    expect(propChange!.previousProperties.font).toBe("Arial");
  });
});
