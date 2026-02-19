/**
 * Tests for app.xml preservation during round-trip
 */

import { Document } from "../../src/core/Document";
import { ZipHandler } from "../../src/zip/ZipHandler";
import { DOCX_PATHS } from "../../src/zip/types";

describe("App.xml Preservation", () => {
  /**
   * Helper: create a minimal DOCX buffer with custom app.xml
   */
  async function createDocxWithAppXml(appXml: string): Promise<Buffer> {
    const doc = Document.create();
    doc.createParagraph("Test content");
    const buffer = await doc.toBuffer();

    // Post-process: inject custom app.xml
    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    zip.updateFile(DOCX_PATHS.APP_PROPS, appXml);
    const modifiedBuffer = await zip.toBuffer();
    zip.clear();

    return modifiedBuffer;
  }

  const richAppXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Template>Normal.dotm</Template>
  <TotalTime>42</TotalTime>
  <Pages>5</Pages>
  <Words>1234</Words>
  <Characters>7890</Characters>
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <Lines>99</Lines>
  <Paragraphs>15</Paragraphs>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Title</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>My Document Title</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company>Acme Corp</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <CharactersWithSpaces>9012</CharactersWithSpaces>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`;

  it("should preserve original app.xml when no properties are modified", async () => {
    const buffer = await createDocxWithAppXml(richAppXml);
    const doc = await Document.loadFromBuffer(buffer);

    // Save without modifying any app properties
    const outputBuffer = await doc.toBuffer();
    doc.dispose();

    // Read back and check app.xml is preserved
    const zip = new ZipHandler();
    await zip.loadFromBuffer(outputBuffer);
    const savedAppXml = zip.getFileAsString(DOCX_PATHS.APP_PROPS);
    zip.clear();

    expect(savedAppXml).toBe(richAppXml);
  });

  it("should preserve HeadingPairs and TotalTime when company is updated", async () => {
    const buffer = await createDocxWithAppXml(richAppXml);
    const doc = await Document.loadFromBuffer(buffer);

    doc.setCompany("New Company");
    const outputBuffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(outputBuffer);
    const savedAppXml = zip.getFileAsString(DOCX_PATHS.APP_PROPS) || "";
    zip.clear();

    // Company should be updated
    expect(savedAppXml).toContain("<Company>New Company</Company>");

    // Original metadata should be preserved
    expect(savedAppXml).toContain("<TotalTime>42</TotalTime>");
    expect(savedAppXml).toContain("<Pages>5</Pages>");
    expect(savedAppXml).toContain("<Words>1234</Words>");
    expect(savedAppXml).toContain("<HeadingPairs>");
    expect(savedAppXml).toContain("My Document Title");
    expect(savedAppXml).toContain("<Template>Normal.dotm</Template>");
  });

  it("should update Application and AppVersion when changed", async () => {
    const buffer = await createDocxWithAppXml(richAppXml);
    const doc = await Document.loadFromBuffer(buffer);

    doc.setApplication("Custom App");
    doc.setAppVersion("2.0.0");
    const outputBuffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(outputBuffer);
    const savedAppXml = zip.getFileAsString(DOCX_PATHS.APP_PROPS) || "";
    zip.clear();

    expect(savedAppXml).toContain("<Application>Custom App</Application>");
    expect(savedAppXml).toContain("<AppVersion>2.0.0</AppVersion>");
    // Company should remain unchanged
    expect(savedAppXml).toContain("<Company>Acme Corp</Company>");
  });

  it("should generate app.xml from scratch for new documents", async () => {
    const doc = Document.create();
    doc.createParagraph("New doc");
    doc.setCompany("Test Corp");

    const outputBuffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(outputBuffer);
    const savedAppXml = zip.getFileAsString(DOCX_PATHS.APP_PROPS) || "";
    zip.clear();

    expect(savedAppXml).toContain("<Company>Test Corp</Company>");
    expect(savedAppXml).toContain("<Application>docxmlater</Application>");
  });

  it("should reset _originalAppPropsXml on dispose", async () => {
    const buffer = await createDocxWithAppXml(richAppXml);
    const doc = await Document.loadFromBuffer(buffer);

    // Verify the original app.xml was loaded
    expect((doc as any)._originalAppPropsXml).toBeDefined();

    doc.dispose();

    // After dispose, the preserved app.xml should be cleared
    expect((doc as any)._originalAppPropsXml).toBeUndefined();
  });
});
