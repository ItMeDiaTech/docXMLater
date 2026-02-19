/**
 * StyleShadingParsing - Tests for style-level shading parsing including
 * paragraph/run shading in tblStylePr and theme attributes in styles
 */

import { Document } from "../../src/core/Document";
import { Paragraph } from "../../src/elements/Paragraph";
import { ZipHandler } from "../../src/zip/ZipHandler";

/**
 * Helper: Inject a table style with shading into styles.xml
 */
async function createDocWithStyleShading(styleXml: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText("placeholder"));
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  const stylesXml = zip.getFileAsString("word/styles.xml")!;

  // Inject the custom style before closing </w:styles>
  const modifiedStyles = stylesXml.replace("</w:styles>", `${styleXml}</w:styles>`);
  zip.addFile("word/styles.xml", modifiedStyles);
  return zip.toBuffer();
}

describe("Style Shading Parsing", () => {
  describe("Paragraph shading in tblStylePr", () => {
    it("should parse paragraph shading from conditional table style", async () => {
      const styleXml = `
        <w:style w:type="table" w:styleId="TestTableStyle">
          <w:name w:val="Test Table Style"/>
          <w:tblStylePr w:type="firstRow">
            <w:pPr>
              <w:shd w:val="clear" w:fill="4472C4" w:themeFill="accent1"/>
            </w:pPr>
          </w:tblStylePr>
        </w:style>
      `;

      const buffer = await createDocWithStyleShading(styleXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const stylesManager = doc.getStylesManager();
      const style = stylesManager.getStyle("TestTableStyle");

      expect(style).toBeDefined();
      const props = style!.getProperties();
      const conditionals = props.tableStyle?.conditionalFormatting;
      expect(conditionals).toBeDefined();

      const firstRow = conditionals!.find((c) => c.type === "firstRow");
      expect(firstRow).toBeDefined();
      expect(firstRow!.paragraphFormatting?.shading).toBeDefined();
      expect(firstRow!.paragraphFormatting!.shading!.fill).toBe("4472C4");
      expect(firstRow!.paragraphFormatting!.shading!.themeFill).toBe("accent1");
      doc.dispose();
    });
  });

  describe("Run shading in tblStylePr", () => {
    it("should parse run shading from conditional table style", async () => {
      const styleXml = `
        <w:style w:type="table" w:styleId="TestTableStyle2">
          <w:name w:val="Test Table Style 2"/>
          <w:tblStylePr w:type="lastCol">
            <w:rPr>
              <w:shd w:val="solid" w:fill="FF0000" w:color="FFFFFF"/>
            </w:rPr>
          </w:tblStylePr>
        </w:style>
      `;

      const buffer = await createDocWithStyleShading(styleXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const stylesManager = doc.getStylesManager();
      const style = stylesManager.getStyle("TestTableStyle2");

      expect(style).toBeDefined();
      const props = style!.getProperties();
      const conditionals = props.tableStyle?.conditionalFormatting;
      const lastCol = conditionals!.find((c) => c.type === "lastCol");
      expect(lastCol).toBeDefined();
      expect(lastCol!.runFormatting?.shading).toBeDefined();
      expect(lastCol!.runFormatting!.shading!.pattern).toBe("solid");
      expect(lastCol!.runFormatting!.shading!.fill).toBe("FF0000");
      doc.dispose();
    });
  });

  describe("Theme attributes in style shading", () => {
    it("should parse theme shading from table cell style", async () => {
      const styleXml = `
        <w:style w:type="table" w:styleId="ThemeTableStyle">
          <w:name w:val="Theme Table Style"/>
          <w:tblStylePr w:type="firstRow">
            <w:tcPr>
              <w:shd w:val="clear" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/>
            </w:tcPr>
          </w:tblStylePr>
        </w:style>
      `;

      const buffer = await createDocWithStyleShading(styleXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const stylesManager = doc.getStylesManager();
      const style = stylesManager.getStyle("ThemeTableStyle");

      expect(style).toBeDefined();
      const props = style!.getProperties();
      const conditionals = props.tableStyle?.conditionalFormatting;
      const firstRow = conditionals!.find((c) => c.type === "firstRow");
      expect(firstRow).toBeDefined();
      expect(firstRow!.cellFormatting?.shading).toBeDefined();
      expect(firstRow!.cellFormatting!.shading!.themeFill).toBe("accent1");
      expect(firstRow!.cellFormatting!.shading!.themeFillTint).toBe("33");
      doc.dispose();
    });

    it("should parse shading from table style default formatting", async () => {
      const styleXml = `
        <w:style w:type="table" w:styleId="DefaultShadingStyle">
          <w:name w:val="Default Shading Style"/>
          <w:tblPr>
            <w:tblW w:w="0" w:type="auto"/>
            <w:shd w:val="clear" w:fill="E8E8E8" w:themeFill="background1" w:themeFillShade="E8"/>
          </w:tblPr>
        </w:style>
      `;

      const buffer = await createDocWithStyleShading(styleXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: "preserve" });
      const stylesManager = doc.getStylesManager();
      const style = stylesManager.getStyle("DefaultShadingStyle");

      expect(style).toBeDefined();
      const props = style!.getProperties();
      expect(props.tableStyle?.table?.shading).toBeDefined();
      expect(props.tableStyle!.table!.shading!.fill).toBe("E8E8E8");
      expect(props.tableStyle!.table!.shading!.themeFill).toBe("background1");
      expect(props.tableStyle!.table!.shading!.themeFillShade).toBe("E8");
      doc.dispose();
    });
  });
});
