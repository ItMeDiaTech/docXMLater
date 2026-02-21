/**
 * ShadingThemeRoundTrip - Tests for theme shading attributes, nil pattern, and auto value round-trip
 */

import { Document } from "../../src/core/Document";
import { Paragraph } from "../../src/elements/Paragraph";
import { Run } from "../../src/elements/Run";
import { Table } from "../../src/elements/Table";
import { ZipHandler } from "../../src/zip/ZipHandler";
import { buildShadingAttributes } from "../../src/elements/CommonTypes";

/**
 * Helper: creates a document buffer with custom XML injected into word/document.xml
 */
async function createDocWithShadingXml(shdXml: string, elementType: 'paragraph' | 'run' | 'table' | 'cell' | 'tblPrEx' = 'paragraph'): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText("placeholder"));
  const buffer = await doc.toBuffer();
  doc.dispose();

  // Post-process: inject shading into the document.xml
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  const docXml = zip.getFileAsString("word/document.xml")!;

  let modifiedXml: string;
  if (elementType === 'paragraph') {
    if (docXml.includes('<w:pPr>')) {
      modifiedXml = docXml.replace('<w:pPr>', `<w:pPr>${shdXml}`);
    } else {
      // No pPr exists, inject one after <w:p> (or <w:p ...>)
      modifiedXml = docXml.replace(/<w:p(\s[^>]*)?>/, `<w:p$1><w:pPr>${shdXml}</w:pPr>`);
    }
  } else if (elementType === 'run') {
    modifiedXml = docXml.replace(
      '<w:rPr>',
      `<w:rPr>${shdXml}`
    );
    // If no rPr exists, add one
    if (modifiedXml === docXml) {
      modifiedXml = docXml.replace(
        '<w:t',
        `<w:rPr>${shdXml}</w:rPr><w:t`
      );
    }
  } else if (elementType === 'table') {
    // Add table wrapper
    modifiedXml = docXml.replace(
      '</w:body>',
      `<w:tbl><w:tblPr>${shdXml}<w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body>`
    );
  } else if (elementType === 'cell') {
    modifiedXml = docXml.replace(
      '</w:body>',
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr>${shdXml}</w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body>`
    );
  } else if (elementType === 'tblPrEx') {
    modifiedXml = docXml.replace(
      '</w:body>',
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tblPrEx>${shdXml}</w:tblPrEx><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body>`
    );
  }

  zip.addFile("word/document.xml", modifiedXml!);
  return zip.toBuffer();
}

const THEME_SHD = '<w:shd w:val="clear" w:fill="F2F2F2" w:color="auto" w:themeFill="background1" w:themeFillShade="F2" w:themeColor="text1" w:themeTint="0D" w:themeFillTint="AA" w:themeShade="BB"/>';

describe("Shading Theme Round-Trip", () => {
  describe("Paragraph theme shading", () => {
    it("should parse all 9 shading attributes from paragraph", async () => {
      const buffer = await createDocWithShadingXml(THEME_SHD, 'paragraph');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const para = doc.getParagraphs()[0]!;
      const shading = para.getFormatting().shading;

      expect(shading).toBeDefined();
      expect(shading!.pattern).toBe("clear");
      expect(shading!.fill).toBe("F2F2F2");
      expect(shading!.color).toBe("auto");
      expect(shading!.themeFill).toBe("background1");
      expect(shading!.themeFillShade).toBe("F2");
      expect(shading!.themeColor).toBe("text1");
      expect(shading!.themeTint).toBe("0D");
      expect(shading!.themeFillTint).toBe("AA");
      expect(shading!.themeShade).toBe("BB");
      doc.dispose();
    });

    it("should round-trip theme shading on paragraph", async () => {
      const buffer = await createDocWithShadingXml(THEME_SHD, 'paragraph');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const buf2 = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buf2, { revisionHandling: 'preserve' });
      const shading = doc2.getParagraphs()[0]!.getFormatting().shading;
      expect(shading!.themeFill).toBe("background1");
      expect(shading!.themeFillShade).toBe("F2");
      expect(shading!.themeColor).toBe("text1");
      doc2.dispose();
    });
  });

  describe("Run theme shading", () => {
    it("should parse theme shading from run", async () => {
      const buffer = await createDocWithShadingXml(THEME_SHD, 'run');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const runs = doc.getParagraphs()[0]!.getRuns();
      const shading = runs[0]?.getFormatting().shading;

      expect(shading).toBeDefined();
      expect(shading!.themeFill).toBe("background1");
      expect(shading!.themeColor).toBe("text1");
      doc.dispose();
    });
  });

  describe("Table theme shading", () => {
    it("should parse theme shading from table", async () => {
      const buffer = await createDocWithShadingXml(THEME_SHD, 'table');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const body = doc.getBodyElements();
      const tables = body.filter((e: any): e is Table => e instanceof Table);
      expect(tables.length).toBeGreaterThan(0);
      const shading = tables[0]!.getShading();

      expect(shading).toBeDefined();
      expect(shading!.themeFill).toBe("background1");
      expect(shading!.themeColor).toBe("text1");
      doc.dispose();
    });
  });

  describe("Cell theme shading", () => {
    it("should parse theme shading from table cell", async () => {
      const buffer = await createDocWithShadingXml(THEME_SHD, 'cell');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const body = doc.getBodyElements();
      const tables = body.filter((e: any): e is Table => e instanceof Table);
      const cell = tables[0]!.getRow(0)!.getCell(0)!;
      const shading = cell.getShading();

      expect(shading).toBeDefined();
      expect(shading!.themeFill).toBe("background1");
      expect(shading!.themeFillShade).toBe("F2");
      doc.dispose();
    });

    it("should round-trip cell theme shading", async () => {
      const buffer = await createDocWithShadingXml(THEME_SHD, 'cell');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const buf2 = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buf2, { revisionHandling: 'preserve' });
      const body = doc2.getBodyElements();
      const tables = body.filter((e: any): e is Table => e instanceof Table);
      const shading = tables[0]!.getRow(0)!.getCell(0)!.getShading();
      expect(shading!.themeFill).toBe("background1");
      doc2.dispose();
    });
  });

  describe("TableRow tblPrEx theme shading", () => {
    it("should parse theme shading from table property exceptions", async () => {
      const buffer = await createDocWithShadingXml(THEME_SHD, 'tblPrEx');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const body = doc.getBodyElements();
      const tables = body.filter((e: any): e is Table => e instanceof Table);
      const row = tables[0]!.getRow(0)!;
      const exceptions = row.getTablePropertyExceptions();

      expect(exceptions).toBeDefined();
      expect(exceptions!.shading).toBeDefined();
      expect(exceptions!.shading!.themeFill).toBe("background1");
      expect(exceptions!.shading!.themeColor).toBe("text1");
      doc.dispose();
    });
  });

  describe("nil pattern round-trip", () => {
    it("should round-trip nil pattern on paragraph", async () => {
      const doc = Document.create();
      const para = new Paragraph();
      para.addText("Test");
      para.setShading({ pattern: "nil" });
      doc.addParagraph(para);

      const buf = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buf, { revisionHandling: 'preserve' });
      const shading = doc2.getParagraphs()[0]!.getFormatting().shading;
      expect(shading).toBeDefined();
      expect(shading!.pattern).toBe("nil");
      doc2.dispose();
    });

    it("should round-trip nil pattern on run", () => {
      const run = new Run("Test");
      run.setShading({ pattern: "nil" });
      const xml = run.toXML();
      const xmlStr = JSON.stringify(xml);
      // toXML returns an XMLElement object; check the serialized structure contains val=nil
      expect(xmlStr).toContain('"w:val":"nil"');
    });
  });

  describe("auto fill/color round-trip", () => {
    it("should round-trip auto fill value", async () => {
      const shdXml = '<w:shd w:val="clear" w:fill="auto" w:color="auto"/>';
      const buffer = await createDocWithShadingXml(shdXml, 'paragraph');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const shading = doc.getParagraphs()[0]!.getFormatting().shading;

      expect(shading).toBeDefined();
      expect(shading!.fill).toBe("auto");
      expect(shading!.color).toBe("auto");
      doc.dispose();
    });

    it("should preserve auto through save/reload cycle", async () => {
      const shdXml = '<w:shd w:val="clear" w:fill="auto" w:color="auto"/>';
      const buffer = await createDocWithShadingXml(shdXml, 'paragraph');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const buf2 = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buf2, { revisionHandling: 'preserve' });
      const shading = doc2.getParagraphs()[0]!.getFormatting().shading;
      expect(shading!.fill).toBe("auto");
      expect(shading!.color).toBe("auto");
      doc2.dispose();
    });
  });

  describe("pPrChange theme shading", () => {
    it("should parse all theme attrs from pPrChange shading", async () => {
      const pPrChangeXml = `<w:pPrChange w:id="1" w:author="Test" w:date="2024-01-01T00:00:00Z"><w:pPr>${THEME_SHD}</w:pPr></w:pPrChange>`;
      const doc = Document.create();
      doc.addParagraph(new Paragraph().addText("placeholder"));
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      let docXml = zip.getFileAsString("word/document.xml")!;
      if (docXml.includes('</w:pPr>')) {
        docXml = docXml.replace('</w:pPr>', `${pPrChangeXml}</w:pPr>`);
      } else {
        // No pPr exists, inject one with pPrChange after <w:p> (or <w:p ...>)
        docXml = docXml.replace(/<w:p(\s[^>]*)?>/, `<w:p$1><w:pPr>${pPrChangeXml}</w:pPr>`);
      }
      zip.addFile("word/document.xml", docXml);
      const modifiedBuf = await zip.toBuffer();

      const doc2 = await Document.loadFromBuffer(modifiedBuf, { revisionHandling: 'preserve' });
      const para = doc2.getParagraphs()[0]!;
      const pPrChange = para.getFormatting().pPrChange;
      expect(pPrChange).toBeDefined();
      expect(pPrChange!.previousProperties?.shading).toBeDefined();
      const shading = pPrChange!.previousProperties!.shading!;
      expect(shading.pattern).toBe("clear");
      expect(shading.themeFill).toBe("background1");
      expect(shading.themeFillShade).toBe("F2");
      expect(shading.themeColor).toBe("text1");
      expect(shading.themeTint).toBe("0D");
      doc2.dispose();
    });
  });

  describe("buildShadingAttributes helper", () => {
    it("should produce correct XML attributes", () => {
      const attrs = buildShadingAttributes({
        pattern: "clear",
        fill: "F2F2F2",
        color: "auto",
        themeFill: "accent1",
        themeFillTint: "40",
      });
      expect(attrs["w:val"]).toBe("clear");
      expect(attrs["w:fill"]).toBe("F2F2F2");
      expect(attrs["w:color"]).toBe("auto");
      expect(attrs["w:themeFill"]).toBe("accent1");
      expect(attrs["w:themeFillTint"]).toBe("40");
    });

    it("should always include w:val and omit other undefined fields", () => {
      const attrs = buildShadingAttributes({
        fill: "FF0000",
      });
      // w:val is always required per ECMA-376, defaults to "clear"
      expect(Object.keys(attrs)).toEqual(["w:val", "w:fill"]);
      expect(attrs["w:val"]).toBe("clear");
    });
  });
});
