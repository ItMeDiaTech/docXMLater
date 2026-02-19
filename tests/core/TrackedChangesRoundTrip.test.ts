/**
 * Tracked Changes ECMA-376 Round-Trip Tests
 *
 * Tests for verifying that all run and paragraph property changes survive
 * the load-parse-serialize cycle with correct OOXML element names and order.
 */

import { Run, RunFormatting } from "../../src/elements/Run";
import { Revision } from "../../src/elements/Revision";
import { Paragraph } from "../../src/elements/Paragraph";
import { Document } from "../../src/core/Document";
import { XMLBuilder } from "../../src/xml/XMLBuilder";

/**
 * Helper to serialize a Run to XML string for inspection
 */
function serializeRun(run: Run): string {
  const xmlElement = run.toXML();
  return XMLBuilder.elementToString(xmlElement);
}

/**
 * Helper to create a run with a property change revision
 */
function createRunWithPropChange(
  text: string,
  currentFormatting: RunFormatting,
  previousProperties: Partial<RunFormatting>
): Run {
  const run = new Run(text, currentFormatting);
  run.setPropertyChangeRevision({
    id: 1,
    author: "TestAuthor",
    date: new Date("2024-01-01T00:00:00Z"),
    previousProperties,
  });
  return run;
}

describe("Tracked Changes Round-Trip", () => {
  describe("rPrChange serialization — full property coverage", () => {
    it("should serialize bold in rPrChange as w:b", () => {
      const run = createRunWithPropChange("text", {}, { bold: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:rPrChange");
      expect(xml).toContain("<w:b ");
      expect(xml).not.toContain("<w:bold");
    });

    it("should serialize italic in rPrChange as w:i", () => {
      const run = createRunWithPropChange("text", {}, { italic: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:i ");
      expect(xml).not.toContain("<w:italic");
    });

    it("should serialize dstrike in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { dstrike: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:dstrike ");
    });

    it("should serialize outline in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { outline: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:outline ");
    });

    it("should serialize shadow in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { shadow: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:shadow ");
    });

    it("should serialize emboss in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { emboss: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:emboss ");
    });

    it("should serialize imprint in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { imprint: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:imprint ");
    });

    it("should serialize vanish in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { vanish: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:vanish ");
    });

    it("should serialize specVanish in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { specVanish: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:specVanish ");
    });

    it("should serialize webHidden in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { webHidden: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:webHidden ");
    });

    it("should serialize rtl in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { rtl: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:rtl ");
    });

    it("should serialize complexScript in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { complexScript: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:cs ");
    });

    it("should serialize complexScriptBold in rPrChange as w:bCs", () => {
      const run = createRunWithPropChange("text", {}, { complexScriptBold: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:bCs ");
    });

    it("should serialize complexScriptItalic in rPrChange as w:iCs", () => {
      const run = createRunWithPropChange("text", {}, { complexScriptItalic: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:iCs ");
    });

    it("should serialize characterSpacing in rPrChange as w:spacing", () => {
      const run = createRunWithPropChange("text", {}, { characterSpacing: 40 });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:spacing ");
      expect(xml).toContain('w:val="40"');
    });

    it("should serialize scaling in rPrChange as w:w", () => {
      const run = createRunWithPropChange("text", {}, { scaling: 150 });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:w ");
      expect(xml).toContain('w:val="150"');
    });

    it("should serialize position in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { position: 6 });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:position ");
      expect(xml).toContain('w:val="6"');
    });

    it("should serialize kerning in rPrChange as w:kern", () => {
      const run = createRunWithPropChange("text", {}, { kerning: 24 });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:kern ");
    });

    it("should serialize noProof in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { noProof: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:noProof ");
    });

    it("should serialize snapToGrid in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { snapToGrid: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:snapToGrid ");
    });

    it("should serialize emphasis in rPrChange as w:em", () => {
      const run = createRunWithPropChange("text", {}, { emphasis: "dot" as any });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:em ");
      expect(xml).toContain('w:val="dot"');
    });

    it("should serialize shading in rPrChange as w:shd", () => {
      const run = createRunWithPropChange("text", {}, {
        shading: { pattern: "clear", fill: "FFFF00", color: "auto" },
      });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:shd ");
      expect(xml).toContain('w:fill="FFFF00"');
    });

    it("should serialize border in rPrChange as w:bdr", () => {
      const run = createRunWithPropChange("text", {}, {
        border: { style: "single" as any, size: 4, color: "000000", space: 1 },
      });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:bdr ");
      expect(xml).toContain('w:val="single"');
    });

    it("should serialize characterStyle in rPrChange as w:rStyle", () => {
      const run = createRunWithPropChange("text", {}, { characterStyle: "Emphasis" });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:rStyle ");
      expect(xml).toContain('w:val="Emphasis"');
    });

    it("should serialize font with extended attrs in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, {
        font: "Arial",
        fontHAnsi: "Arial",
        fontEastAsia: "MS Mincho",
        fontCs: "Arial Unicode MS",
        fontHint: "eastAsia",
      });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:rFonts ");
      expect(xml).toContain('w:ascii="Arial"');
      expect(xml).toContain('w:hAnsi="Arial"');
      expect(xml).toContain('w:eastAsia="MS Mincho"');
      expect(xml).toContain('w:cs="Arial Unicode MS"');
      expect(xml).toContain('w:hint="eastAsia"');
    });

    it("should serialize allCaps in rPrChange as w:caps", () => {
      const run = createRunWithPropChange("text", {}, { allCaps: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:caps ");
      expect(xml).not.toContain("<w:allCaps");
    });

    it("should serialize smallCaps in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { smallCaps: true });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:smallCaps ");
    });

    it("should serialize underline in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { underline: "double" });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:u ");
      expect(xml).toContain('w:val="double"');
    });

    it("should serialize color with theme attrs in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, {
        color: "FF0000",
        themeColor: "accent1" as any,
        themeTint: 128,
      });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:color ");
      expect(xml).toContain('w:val="FF0000"');
      expect(xml).toContain('w:themeColor="accent1"');
      expect(xml).toContain('w:themeTint="80"'); // 128 = 0x80
    });

    it("should serialize highlight in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { highlight: "yellow" });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:highlight ");
      expect(xml).toContain('w:val="yellow"');
    });

    it("should serialize effect in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { effect: "shimmer" });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:effect ");
      expect(xml).toContain('w:val="shimmer"');
    });

    it("should serialize fitText in rPrChange", () => {
      const run = createRunWithPropChange("text", {}, { fitText: 2400 });
      const xml = serializeRun(run);
      expect(xml).toContain("<w:fitText ");
      expect(xml).toContain('w:val="2400"');
    });

    it("should serialize ALL 30+ properties in one rPrChange", () => {
      const prevProps: Partial<RunFormatting> = {
        characterStyle: "Emphasis",
        font: "Arial",
        fontHAnsi: "Arial",
        fontEastAsia: "MS Mincho",
        fontCs: "Arial Unicode MS",
        bold: true,
        complexScriptBold: true,
        italic: true,
        complexScriptItalic: true,
        allCaps: true,
        smallCaps: true,
        shading: { pattern: "clear", fill: "FFFF00" },
        emphasis: "dot" as any,
        outline: true,
        shadow: true,
        emboss: true,
        imprint: true,
        noProof: true,
        snapToGrid: true,
        vanish: true,
        webHidden: true,
        specVanish: true,
        rtl: true,
        complexScript: true,
        strike: true,
        dstrike: true,
        underline: "double",
        characterSpacing: 40,
        scaling: 150,
        position: 6,
        kerning: 24,
        size: 14,
        color: "FF0000",
        highlight: "yellow",
        subscript: true,
        effect: "shimmer",
        fitText: 2400,
        border: { style: "single" as any, size: 4, color: "000000" },
      };

      const run = createRunWithPropChange("text", {}, prevProps);
      const xml = serializeRun(run);

      // Verify all major elements are present
      expect(xml).toContain("<w:rPrChange");
      expect(xml).toContain("<w:rStyle ");
      expect(xml).toContain("<w:rFonts ");
      expect(xml).toContain("<w:bdr ");
      expect(xml).toContain("<w:b ");
      expect(xml).toContain("<w:bCs ");
      expect(xml).toContain("<w:i ");
      expect(xml).toContain("<w:iCs ");
      expect(xml).toContain("<w:caps ");
      expect(xml).toContain("<w:smallCaps ");
      expect(xml).toContain("<w:shd ");
      expect(xml).toContain("<w:em ");
      expect(xml).toContain("<w:outline ");
      expect(xml).toContain("<w:shadow ");
      expect(xml).toContain("<w:emboss ");
      expect(xml).toContain("<w:imprint ");
      expect(xml).toContain("<w:noProof ");
      expect(xml).toContain("<w:snapToGrid ");
      expect(xml).toContain("<w:vanish ");
      expect(xml).toContain("<w:webHidden ");
      expect(xml).toContain("<w:specVanish ");
      expect(xml).toContain("<w:rtl ");
      expect(xml).toContain("<w:cs ");
      expect(xml).toContain("<w:strike ");
      expect(xml).toContain("<w:dstrike ");
      expect(xml).toContain("<w:u ");
      expect(xml).toContain("<w:spacing ");
      expect(xml).toContain("<w:w ");
      expect(xml).toContain("<w:position ");
      expect(xml).toContain("<w:kern ");
      expect(xml).toContain("<w:effect ");
      expect(xml).toContain("<w:sz ");
      expect(xml).toContain("<w:szCs ");
      expect(xml).toContain("<w:color ");
      expect(xml).toContain("<w:highlight ");
      expect(xml).toContain("<w:vertAlign ");
      expect(xml).toContain("<w:fitText ");
    });
  });

  describe("Revision.toXML() for runPropertiesChange", () => {
    it("should use correct OOXML element names via generateRunPropertiesXML", () => {
      const rev = Revision.createRunPropertiesChange(
        "TestAuthor",
        new Run("text"),
        { bold: true, font: "Arial", allCaps: true, size: 12 }
      );
      const xmlElement = rev.toXML();
      if (!xmlElement) throw new Error("Revision.toXML() returned null");
      const xml = XMLBuilder.elementToString(xmlElement);

      // Should use w:b, not w:bold
      expect(xml).toContain("<w:b ");
      expect(xml).not.toContain("<w:bold");
      // Should use w:rFonts, not w:font
      expect(xml).toContain("<w:rFonts ");
      expect(xml).not.toContain("<w:font");
      // Should use w:caps, not w:allCaps
      expect(xml).toContain("<w:caps ");
      expect(xml).not.toContain("<w:allCaps");
      // Should use w:sz, not w:size
      expect(xml).toContain("<w:sz ");
      expect(xml).not.toContain("<w:size");
    });

    it("should include all extended properties in Revision.toXML()", () => {
      const rev = Revision.createRunPropertiesChange(
        "TestAuthor",
        new Run("text"),
        {
          dstrike: true,
          outline: true,
          shadow: true,
          emboss: true,
          imprint: true,
          vanish: true,
          rtl: true,
          complexScript: true,
          characterSpacing: 20,
          kerning: 16,
          noProof: true,
        }
      );
      const xmlElement = rev.toXML();
      if (!xmlElement) throw new Error("Revision.toXML() returned null");
      const xml = XMLBuilder.elementToString(xmlElement);

      expect(xml).toContain("<w:dstrike ");
      expect(xml).toContain("<w:outline ");
      expect(xml).toContain("<w:shadow ");
      expect(xml).toContain("<w:emboss ");
      expect(xml).toContain("<w:imprint ");
      expect(xml).toContain("<w:vanish ");
      expect(xml).toContain("<w:rtl ");
      expect(xml).toContain("<w:cs ");
      expect(xml).toContain("<w:spacing ");
      expect(xml).toContain("<w:kern ");
      expect(xml).toContain("<w:noProof ");
    });
  });

  describe("rPrChange full round-trip (load → parse → save)", () => {
    it("should preserve all rPrChange properties through document round-trip", async () => {
      // Create a document with a run that has property change revision
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run("test text", { bold: true });
      run.setPropertyChangeRevision({
        id: 42,
        author: "RoundTripAuthor",
        date: new Date("2024-06-15T10:00:00Z"),
        previousProperties: {
          italic: true,
          dstrike: true,
          outline: true,
          shadow: true,
          emboss: true,
          vanish: true,
          rtl: true,
          complexScriptBold: true,
          characterSpacing: 30,
          kerning: 20,
          noProof: true,
          snapToGrid: true,
          font: "Times New Roman",
          fontHAnsi: "Times New Roman",
          fontCs: "Arial",
          size: 14,
          color: "0000FF",
          highlight: "green",
          characterStyle: "Strong",
          shading: { pattern: "clear", fill: "E0E0E0" },
          border: { style: "single" as any, size: 4, color: "333333" },
        },
      });
      para.addRun(run);
      doc.addParagraph(para);

      // Save to buffer
      const buffer = await doc.toBuffer();
      doc.dispose();

      // Reload from buffer (preserve revisions)
      const doc2 = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const paras = doc2.getParagraphs();
      expect(paras.length).toBeGreaterThan(0);

      const runs = paras[0]!.getRuns();
      expect(runs.length).toBeGreaterThan(0);

      const propChange = runs[0]!.getPropertyChangeRevision();
      expect(propChange).toBeDefined();
      expect(propChange!.author).toBe("RoundTripAuthor");

      const prev = propChange!.previousProperties as Partial<RunFormatting>;
      expect(prev.italic).toBe(true);
      expect(prev.dstrike).toBe(true);
      expect(prev.outline).toBe(true);
      expect(prev.shadow).toBe(true);
      expect(prev.emboss).toBe(true);
      expect(prev.vanish).toBe(true);
      expect(prev.rtl).toBe(true);
      expect(prev.complexScriptBold).toBe(true);
      expect(prev.characterSpacing).toBe(30);
      expect(prev.kerning).toBe(20);
      expect(prev.noProof).toBe(true);
      expect(prev.snapToGrid).toBe(true);
      expect(prev.font).toBe("Times New Roman");
      expect(prev.fontHAnsi).toBe("Times New Roman");
      expect(prev.fontCs).toBe("Arial");
      expect(prev.size).toBe(14);
      expect(prev.color).toBe("0000FF");
      expect(prev.highlight).toBe("green");
      expect(prev.characterStyle).toBe("Strong");
      expect(prev.shading).toBeDefined();
      expect(prev.shading!.fill).toBe("E0E0E0");
      expect(prev.border).toBeDefined();
      expect(prev.border!.style).toBe("single");

      doc2.dispose();
    });
  });

  describe("pPrChange textAlignment round-trip", () => {
    it("should preserve textAlignment through pPrChange round-trip", async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText("test");
      para.setParagraphPropertiesChange({
        id: "10",
        author: "TestAuthor",
        date: "2024-01-01T00:00:00Z",
        previousProperties: {
          alignment: "center",
          textAlignment: "baseline",
          textDirection: "tbRl",
        },
      });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const paras = doc2.getParagraphs();
      const change = paras[0]!.formatting.pPrChange;

      expect(change).toBeDefined();
      expect(change!.previousProperties).toBeDefined();
      expect((change!.previousProperties as any).alignment).toBe("center");
      expect((change!.previousProperties as any).textAlignment).toBe("baseline");
      expect((change!.previousProperties as any).textDirection).toBe("tbRl");

      doc2.dispose();
    });
  });

  describe("pPrChange element ordering", () => {
    it("should serialize pPrChange properties in CT_PPrBase schema order", () => {
      const para = new Paragraph();
      para.addText("test");
      para.setParagraphPropertiesChange({
        id: "10",
        author: "TestAuthor",
        date: "2024-01-01T00:00:00Z",
        previousProperties: {
          style: "Heading1",
          keepNext: true,
          keepLines: true,
          widowControl: true,
          numbering: { level: 0, numId: 1 },
          suppressLineNumbers: true,
          bidi: true,
          spacing: { before: 240, after: 120 },
          indentation: { left: 720 },
          contextualSpacing: true,
          alignment: "center",
          textDirection: "tbRl",
          textAlignment: "baseline",
          outlineLevel: 1,
        },
      });

      const xmlElement = para.toXML();
      const xml = XMLBuilder.elementToString(xmlElement);

      // Verify order by checking relative positions
      const pStylePos = xml.indexOf("<w:pStyle");
      const keepNextPos = xml.indexOf("<w:keepNext");
      const keepLinesPos = xml.indexOf("<w:keepLines");
      const widowControlPos = xml.indexOf("<w:widowControl");
      const numPrPos = xml.indexOf("<w:numPr");
      const suppressLinePos = xml.indexOf("<w:suppressLineNumbers");
      const bidiPos = xml.indexOf("<w:bidi");
      const spacingPos = xml.indexOf("<w:spacing");
      const indPos = xml.indexOf("<w:ind");
      const contextualPos = xml.indexOf("<w:contextualSpacing");
      const jcPos = xml.indexOf("<w:jc");
      const textDirPos = xml.indexOf("<w:textDirection");
      const textAlignPos = xml.indexOf("<w:textAlignment");
      const outlineLvlPos = xml.indexOf("<w:outlineLvl");

      // All positions should be found
      expect(pStylePos).toBeGreaterThan(-1);
      expect(keepNextPos).toBeGreaterThan(-1);
      expect(jcPos).toBeGreaterThan(-1);
      expect(outlineLvlPos).toBeGreaterThan(-1);

      // Verify CT_PPrBase order
      expect(pStylePos).toBeLessThan(keepNextPos);
      expect(keepNextPos).toBeLessThan(keepLinesPos);
      expect(keepLinesPos).toBeLessThan(widowControlPos);
      expect(widowControlPos).toBeLessThan(numPrPos);
      expect(numPrPos).toBeLessThan(suppressLinePos);
      expect(suppressLinePos).toBeLessThan(bidiPos);
      expect(bidiPos).toBeLessThan(spacingPos);
      expect(spacingPos).toBeLessThan(indPos);
      expect(indPos).toBeLessThan(contextualPos);
      expect(contextualPos).toBeLessThan(jcPos);
      expect(jcPos).toBeLessThan(textDirPos);
      expect(textDirPos).toBeLessThan(textAlignPos);
      expect(textAlignPos).toBeLessThan(outlineLvlPos);
    });
  });

  describe("people.xml author coverage", () => {
    it("should include rPrChange authors in people.xml", async () => {
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run("text");
      run.setPropertyChangeRevision({
        id: 1,
        author: "RPrChangeOnlyAuthor",
        date: new Date("2024-01-01T00:00:00Z"),
        previousProperties: { bold: true },
      });
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();

      // Extract people.xml from the buffer
      const JSZip = (await import("jszip")).default;
      const zip = await JSZip.loadAsync(buffer);
      const peopleXml = await zip.file("word/people.xml")?.async("string");

      expect(peopleXml).toBeDefined();
      expect(peopleXml).toContain("RPrChangeOnlyAuthor");

      doc.dispose();
    });
  });

  describe("Parser coverage for new rPrChange properties", () => {
    it("should parse w:webHidden in rPrChange", async () => {
      // Create document with webHidden in rPrChange
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run("text");
      run.setPropertyChangeRevision({
        id: 1,
        author: "Test",
        date: new Date("2024-01-01T00:00:00Z"),
        previousProperties: { webHidden: true },
      });
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const runs2 = doc2.getParagraphs()[0]!.getRuns();
      const propChange2 = runs2[0]!.getPropertyChangeRevision();
      const prev2 = propChange2?.previousProperties as Partial<RunFormatting>;
      expect(prev2.webHidden).toBe(true);
      doc2.dispose();
    });

    it("should parse w:cs (complexScript) in rPrChange", async () => {
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run("text");
      run.setPropertyChangeRevision({
        id: 1,
        author: "Test",
        date: new Date("2024-01-01T00:00:00Z"),
        previousProperties: { complexScript: true },
      });
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const runs2 = doc2.getParagraphs()[0]!.getRuns();
      const propChange2 = runs2[0]!.getPropertyChangeRevision();
      const prev2 = propChange2?.previousProperties as Partial<RunFormatting>;
      expect(prev2.complexScript).toBe(true);
      doc2.dispose();
    });

    it("should parse extended rFonts attrs in rPrChange", async () => {
      const doc = Document.create();
      const para = new Paragraph();
      const run = new Run("text");
      run.setPropertyChangeRevision({
        id: 1,
        author: "Test",
        date: new Date("2024-01-01T00:00:00Z"),
        previousProperties: {
          font: "Arial",
          fontHAnsi: "Calibri",
          fontEastAsia: "MS Mincho",
          fontCs: "Arial Unicode MS",
          fontHint: "eastAsia",
        },
      });
      para.addRun(run);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const runs2 = doc2.getParagraphs()[0]!.getRuns();
      const propChange2 = runs2[0]!.getPropertyChangeRevision();
      const prev2 = propChange2?.previousProperties as Partial<RunFormatting>;
      expect(prev2.font).toBe("Arial");
      expect(prev2.fontHAnsi).toBe("Calibri");
      expect(prev2.fontEastAsia).toBe("MS Mincho");
      expect(prev2.fontCs).toBe("Arial Unicode MS");
      expect(prev2.fontHint).toBe("eastAsia");
      doc2.dispose();
    });

    it("should parse w:textAlignment in pPrChange", async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText("test");
      para.setParagraphPropertiesChange({
        id: "10",
        author: "Test",
        date: "2024-01-01T00:00:00Z",
        previousProperties: {
          textAlignment: "center",
        },
      });
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const change = doc2.getParagraphs()[0]!.formatting.pPrChange;
      expect(change).toBeDefined();
      expect((change!.previousProperties as any).textAlignment).toBe("center");
      doc2.dispose();
    });
  });
});
