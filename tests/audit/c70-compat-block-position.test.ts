/**
 * Regression: upgradeCompatBlock() must insert a new w:compat block at its
 * CT_Settings sequence position, not as the last child of w:settings.
 * CT_Settings (ECMA-376 17.15.1.78) is an xsd:sequence where w:compat
 * precedes w:rsids, m:mathPr, w:themeFontLang, w:clrSchemeMapping,
 * w:shapeDefaults, w:decimalSymbol, and w:listSeparator; the no-compat path
 * previously appended the block immediately before </w:settings>, producing
 * schema-invalid part content whenever any of those siblings were present.
 */
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { CompatibilityUpgrader } from '../../src/processors/CompatibilityUpgrader';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

const SETTINGS_NO_COMPAT_THEMEFONTLANG = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:themeFontLang w:val="en-US"/>
</w:settings>`;

const SETTINGS_NO_COMPAT_TRAILING_SIBLINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:defaultTabStop w:val="720"/>
  <w:rsids>
    <w:rsidRoot w:val="00A1B2C3"/>
    <w:rsid w:val="00A1B2C3"/>
  </w:rsids>
  <w:themeFontLang w:val="en-US"/>
  <w:decimalSymbol w:val="."/>
  <w:listSeparator w:val=","/>
</w:settings>`;

const SETTINGS_NO_COMPAT_NO_SIBLINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

async function createDocxWithSettings(settingsXml: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText('Test content'));
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.updateFile(DOCX_PATHS.SETTINGS, settingsXml);
  return await zipHandler.toBuffer();
}

describe('CompatibilityUpgrader w:compat insertion position (CT_Settings sequence)', () => {
  it('inserts w:compat before w:themeFontLang when no compat block exists', () => {
    const result = CompatibilityUpgrader.upgradeCompatBlock(SETTINGS_NO_COMPAT_THEMEFONTLANG, 12);

    const compatIndex = result.xml.indexOf('<w:compat>');
    const themeFontLangIndex = result.xml.indexOf('<w:themeFontLang');
    expect(compatIndex).toBeGreaterThan(-1);
    expect(themeFontLangIndex).toBeGreaterThan(-1);
    expect(compatIndex).toBeLessThan(themeFontLangIndex);
    // The block still lands after the preceding sequence members
    expect(compatIndex).toBeGreaterThan(result.xml.indexOf('<w:characterSpacingControl'));
  });

  it('inserts w:compat before the earliest trailing sibling (w:rsids)', () => {
    const result = CompatibilityUpgrader.upgradeCompatBlock(
      SETTINGS_NO_COMPAT_TRAILING_SIBLINGS,
      12
    );

    const compatIndex = result.xml.indexOf('<w:compat>');
    const rsidsIndex = result.xml.indexOf('<w:rsids>');
    expect(compatIndex).toBeGreaterThan(-1);
    expect(rsidsIndex).toBeGreaterThan(-1);
    expect(compatIndex).toBeLessThan(rsidsIndex);
    // All original siblings survive in order
    expect(result.xml.indexOf('<w:themeFontLang')).toBeGreaterThan(
      result.xml.indexOf('</w:rsids>')
    );
    expect(result.xml.indexOf('<w:decimalSymbol')).toBeLessThan(
      result.xml.indexOf('<w:listSeparator')
    );
  });

  it('falls back to inserting before </w:settings> when no trailing siblings exist', () => {
    const result = CompatibilityUpgrader.upgradeCompatBlock(SETTINGS_NO_COMPAT_NO_SIBLINGS, 12);

    const compatIndex = result.xml.indexOf('<w:compat>');
    expect(compatIndex).toBeGreaterThan(-1);
    expect(compatIndex).toBeGreaterThan(result.xml.indexOf('<w:defaultTabStop'));
    expect(compatIndex).toBeLessThan(result.xml.indexOf('</w:settings>'));
    expect(result.xml).toContain('compatibilityMode');
    expect(result.xml).toContain('w:val="15"');
  });

  it('persists the ordered w:compat block through upgradeToModernFormat() save', async () => {
    const buffer = await createDocxWithSettings(SETTINGS_NO_COMPAT_TRAILING_SIBLINGS);
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    try {
      const report = doc.upgradeToModernFormat();
      expect(report.changed).toBe(true);

      const savedBuffer = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(savedBuffer);
      const settingsXml = zip.getFileAsString(DOCX_PATHS.SETTINGS)!;

      const compatIndex = settingsXml.indexOf('<w:compat>');
      const rsidsIndex = settingsXml.indexOf('<w:rsids>');
      const themeFontLangIndex = settingsXml.indexOf('<w:themeFontLang');
      expect(compatIndex).toBeGreaterThan(-1);
      expect(rsidsIndex).toBeGreaterThan(-1);
      expect(themeFontLangIndex).toBeGreaterThan(-1);
      expect(compatIndex).toBeLessThan(rsidsIndex);
      expect(compatIndex).toBeLessThan(themeFontLangIndex);
    } finally {
      doc.dispose();
    }
  });
});
