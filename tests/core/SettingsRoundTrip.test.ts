/**
 * Settings.xml Round-Trip and Compatibility Mode Tests
 *
 * Tests cover:
 * 1. Phase 1a: Parsing managed settings from settings.xml on load
 * 2. Phase 1b: Merge and save flow for settings.xml
 * 3. Phase 2: Compatibility mode detection API
 * 4. Phase 3: Modern format upgrade (upgradeToModernFormat)
 *
 * Per MS-DOCX spec, the default compatibilityMode when absent is 12 (Word 2007).
 * See: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/90138c4d-eb18-4edc-aa6c-dfb799cb1d0d
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { CompatibilityMode } from '../../src/types/compatibility-types';
import { CompatibilityUpgrader } from '../../src/utils/CompatibilityUpgrader';
import { LEGACY_COMPAT_ELEMENTS, LEGACY_COMPAT_ELEMENT_NAMES, MODERN_COMPAT_SETTINGS, MS_WORD_COMPAT_URI } from '../../src/constants/legacyCompatFlags';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: Creates a minimal DOCX buffer with custom settings.xml content.
 * First creates a valid DOCX, then post-processes the ZIP to inject custom settings.
 * This avoids the save flow overwriting the custom settings via updateSettingsXml().
 */
async function createDocxWithSettings(settingsXml: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText('Test content'));
  const buffer = await doc.toBuffer();
  doc.dispose();

  // Post-process: replace settings.xml in the saved ZIP
  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.updateFile(DOCX_PATHS.SETTINGS, settingsXml);
  return await zipHandler.toBuffer();
}

/** Minimal settings.xml with compatibilityMode=15 (modern) */
const SETTINGS_MODE_15 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
  </w:compat>
  <w:themeFontLang w:val="en-US"/>
</w:settings>`;

/** Settings.xml with compatibilityMode=12 (Word 2007) */
const SETTINGS_MODE_12 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:useFELayout/>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/>
  </w:compat>
  <w:themeFontLang w:val="en-US"/>
</w:settings>`;

/** Settings.xml with compatibilityMode=14 (Word 2010) and legacy flags */
const SETTINGS_MODE_14 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:adjustLineHeightInTable/>
    <w:useFELayout/>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
  </w:compat>
  <w:themeFontLang w:val="en-US"/>
</w:settings>`;

/** Settings.xml with no w:compat block at all */
const SETTINGS_NO_COMPAT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:themeFontLang w:val="en-US"/>
</w:settings>`;

/** Settings.xml with track changes enabled and protection */
const SETTINGS_WITH_TRACKING_AND_PROTECTION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="100"/>
  <w:revisionView w:insDel="1" w:formatting="0" w:inkAnnotations="1"/>
  <w:trackRevisions/>
  <w:doNotTrackFormatting/>
  <w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
  <w:rsids>
    <w:rsidRoot w:val="00A1B2C3"/>
    <w:rsid w:val="00A1B2C3"/>
    <w:rsid w:val="00D4E5F6"/>
    <w:rsid w:val="00789ABC"/>
  </w:rsids>
  <w:themeFontLang w:val="en-US"/>
</w:settings>`;

/** Settings with evenAndOddHeaders and other framework-unmanaged elements */
const SETTINGS_WITH_EXTRA_ELEMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="150"/>
  <w:proofState w:spelling="clean" w:grammar="clean"/>
  <w:defaultTabStop w:val="360"/>
  <w:evenAndOddHeaders/>
  <w:drawingGridHorizontalSpacing w:val="120"/>
  <w:noPunctuationKerning/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
  <w:themeFontLang w:val="en-US"/>
  <w:decimalSymbol w:val="."/>
  <w:listSeparator w:val=","/>
</w:settings>`;

describe('Settings.xml Round-Trip and Compatibility Mode', () => {

  // ======================================================================
  // Phase 1a: Parsing managed settings from settings.xml on load
  // ======================================================================
  describe('Phase 1a: Parse settings on load', () => {

    it('should parse trackRevisions as enabled', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // trackChangesEnabled should be true from <w:trackRevisions/>
      // We verify via the public isTrackChangesEnabled() or by enabling and checking it stays
      // Actually, let's verify by saving and checking the output
      const outputBuffer = await doc.toBuffer();
      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('<w:trackRevisions');
      reloaded.dispose();
      doc.dispose();
    });

    it('should parse doNotTrackFormatting flag', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Save without modifications - doNotTrackFormatting should be preserved
      const outputBuffer = await doc.toBuffer();
      const settingsXml = doc.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      // Since no modifications, original XML should be returned as-is
      expect(settingsXml).toContain('<w:doNotTrackFormatting');
      doc.dispose();
    });

    it('should parse document protection settings', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.isProtected()).toBe(true);
      const prot = doc.getProtection();
      expect(prot).toBeDefined();
      expect(prot!.edit).toBe('trackedChanges');
      expect(prot!.enforcement).toBe(true);
      doc.dispose();
    });

    it('should parse RSIDs from settings.xml', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getRsidRoot()).toBe('00A1B2C3');
      const rsids = doc.getRsids();
      expect(rsids).toContain('00A1B2C3');
      expect(rsids).toContain('00D4E5F6');
      expect(rsids).toContain('00789ABC');
      expect(rsids.length).toBe(3);
      doc.dispose();
    });

    it('should parse revisionView settings', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Save and check that revisionView with formatting="0" is preserved
      const outputBuffer = await doc.toBuffer();
      const settingsXml = doc.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      // No modifications made, so original XML should be returned
      expect(settingsXml).toContain('<w:revisionView');
      doc.dispose();
    });
  });

  // ======================================================================
  // Phase 1b: Settings merge and save flow
  // ======================================================================
  describe('Phase 1b: Settings merge and save flow', () => {

    it('should preserve original settings.xml when no modifications made', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_EXTRA_ELEMENTS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Save without any modifications
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      // Reload and verify all original elements are preserved
      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('<w:evenAndOddHeaders');
      expect(settingsXml).toContain('<w:drawingGridHorizontalSpacing');
      expect(settingsXml).toContain('<w:noPunctuationKerning');
      expect(settingsXml).toContain('<w:proofState');
      expect(settingsXml).toContain('w:percent="150"');
      expect(settingsXml).toContain('w:val="360"'); // defaultTabStop
      expect(settingsXml).toContain('<w:decimalSymbol');
      expect(settingsXml).toContain('<w:listSeparator');
      reloaded.dispose();
    });

    it('should preserve compatibility mode when enabling track changes', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_12);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Enable track changes - this should modify settings but preserve compat mode
      doc.enableTrackChanges({ author: 'TestAuthor' });
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      // Verify compat mode is still 12 and trackRevisions was added
      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('w:val="12"');
      expect(settingsXml).toContain('<w:trackRevisions');
      reloaded.dispose();
    });

    it('should not remove existing trackRevisions when modifying protection', async () => {
      // This is the critical C4 fix test - protectDocument() should not remove trackRevisions
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Modify only protection - trackRevisions should be preserved
      doc.protectDocument({ edit: 'readOnly' });
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      // trackRevisions should still be present (was parsed from settings on load)
      expect(settingsXml).toContain('<w:trackRevisions');
      // Protection should be updated
      expect(settingsXml).toContain('w:edit="readOnly"');
      reloaded.dispose();
    });

    it('should not remove RSIDs when enabling track changes', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Enable track changes - RSIDs should be preserved
      doc.enableTrackChanges({ author: 'NewAuthor' });
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('00A1B2C3');
      expect(settingsXml).toContain('00D4E5F6');
      expect(settingsXml).toContain('00789ABC');
      reloaded.dispose();
    });

    it('should write trackRevisions to settings.xml when enableTrackChanges is called', async () => {
      // This verifies the fix for the bug where enableTrackChanges() had no effect on settings.xml
      const buffer = await createDocxWithSettings(SETTINGS_MODE_15);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Original should NOT have trackRevisions
      const originalSettings = doc.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(originalSettings).not.toContain('<w:trackRevisions');

      // Enable track changes
      doc.enableTrackChanges({ author: 'TestAuthor' });
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      // Saved settings should now have trackRevisions
      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('<w:trackRevisions');
      reloaded.dispose();
    });

    it('should remove trackRevisions from settings.xml when disableTrackChanges is called', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      doc.disableTrackChanges();
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).not.toContain('<w:trackRevisions');
      reloaded.dispose();
    });

    it('should preserve framework-unmanaged elements during merge', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_EXTRA_ELEMENTS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Trigger a settings modification
      doc.enableTrackChanges({ author: 'Test' });
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      // All framework-unmanaged elements should still be present
      expect(settingsXml).toContain('<w:evenAndOddHeaders');
      expect(settingsXml).toContain('<w:drawingGridHorizontalSpacing');
      expect(settingsXml).toContain('<w:noPunctuationKerning');
      expect(settingsXml).toContain('<w:proofState');
      expect(settingsXml).toContain('<w:decimalSymbol');
      expect(settingsXml).toContain('<w:listSeparator');
      // And the new trackRevisions should also be present
      expect(settingsXml).toContain('<w:trackRevisions');
      reloaded.dispose();
    });

    it('should generate settings from scratch for new documents', async () => {
      const doc = Document.create();
      doc.addParagraph(new Paragraph().addText('Hello'));
      doc.enableTrackChanges({ author: 'Author' });
      const buffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('<w:trackRevisions');
      expect(settingsXml).toContain('compatibilityMode');
      expect(settingsXml).toContain('w:val="15"');
      reloaded.dispose();
    });

    it('should handle document protection with crypto attributes', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_15);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      doc.protectDocument({
        edit: 'readOnly',
        enforcement: true,
        cryptProviderType: 'rsaAES',
        cryptAlgorithmClass: 'hash',
        cryptAlgorithmType: 'typeAny',
        cryptAlgorithmSid: 14,
        cryptSpinCount: 100000,
      });
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('w:edit="readOnly"');
      expect(settingsXml).toContain('w:enforcement="1"');
      expect(settingsXml).toContain('w:cryptProviderType="rsaAES"');
      expect(settingsXml).toContain('w:cryptAlgorithmSid="14"');
      expect(settingsXml).toContain('w:cryptSpinCount="100000"');
      reloaded.dispose();
    });

    it('should handle unprotectDocument removing protection', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.isProtected()).toBe(true);
      doc.unprotectDocument();
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      const settingsXml = reloaded.getZipHandler().getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).not.toContain('<w:documentProtection');
      reloaded.dispose();
    });
  });

  // ======================================================================
  // Phase 2: Compatibility mode detection API
  // ======================================================================
  describe('Phase 2: Compatibility mode detection API', () => {

    it('should detect mode 15 (Word 2013+)', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_15);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);
      expect(doc.getCompatibilityMode()).toBe(15);
      expect(doc.isCompatibilityMode()).toBe(false);
      doc.dispose();
    });

    it('should detect mode 12 (Word 2007)', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_12);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2007);
      expect(doc.getCompatibilityMode()).toBe(12);
      expect(doc.isCompatibilityMode()).toBe(true);
      doc.dispose();
    });

    it('should detect mode 14 (Word 2010)', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_14);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2010);
      expect(doc.getCompatibilityMode()).toBe(14);
      expect(doc.isCompatibilityMode()).toBe(true);
      doc.dispose();
    });

    it('should default to mode 12 when w:compat block is absent (per MS-DOCX spec)', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_NO_COMPAT);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Per MS-DOCX spec: "The default value of this element is 12"
      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2007);
      expect(doc.isCompatibilityMode()).toBe(true);
      doc.dispose();
    });

    it('should return mode 15 for new documents', async () => {
      const doc = Document.create();
      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);
      expect(doc.isCompatibilityMode()).toBe(false);
      doc.dispose();
    });

    it('should parse legacy compat flags', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_14);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const info = doc.getCompatibilityInfo();
      expect(info.legacyFlags).toContain('useFELayout');
      expect(info.legacyFlags).toContain('adjustLineHeightInTable');
      expect(info.legacyFlags.length).toBe(2);
      doc.dispose();
    });

    it('should parse w:compatSetting entries', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_15);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const info = doc.getCompatibilityInfo();
      expect(info.compatSettings.length).toBe(5);
      expect(info.compatSettings[0]!.name).toBe('compatibilityMode');
      expect(info.compatSettings[0]!.val).toBe('15');
      expect(info.compatSettings[1]!.name).toBe('overrideTableStyleFontSizeAndJustification');
      doc.dispose();
    });

    it('should return full CompatibilityInfo for documents without w:compat', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_NO_COMPAT);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const info = doc.getCompatibilityInfo();
      expect(info.mode).toBe(CompatibilityMode.Word2007);
      expect(info.isLegacyMode).toBe(true);
      expect(info.compatSettings).toEqual([]);
      expect(info.legacyFlags).toEqual([]);
      doc.dispose();
    });

    it('should preserve compatibility mode through round-trip', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_12);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Verify initial mode
      expect(doc.getCompatibilityMode()).toBe(12);

      // Save and reload
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      // Mode should still be 12, not upgraded to 15
      expect(reloaded.getCompatibilityMode()).toBe(12);
      reloaded.dispose();
    });

    it('should preserve compatibility mode even when track changes is modified', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_14);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(14);

      // Enable track changes triggers settings merge
      doc.enableTrackChanges({ author: 'Test' });
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const reloaded = await Document.loadFromBuffer(outputBuffer, { revisionHandling: 'preserve' });
      // Mode should still be 14 even after settings merge
      expect(reloaded.getCompatibilityMode()).toBe(14);
      // Legacy flags should be preserved too
      const info = reloaded.getCompatibilityInfo();
      expect(info.legacyFlags).toContain('useFELayout');
      expect(info.legacyFlags).toContain('adjustLineHeightInTable');
      reloaded.dispose();
    });
  });

  // ======================================================================
  // dispose() cleanup tests
  // ======================================================================
  // Phase 3: Modern Format Upgrade
  // ======================================================================
  describe('Phase 3: upgradeToModernFormat()', () => {

    it('should upgrade mode 12 document to mode 15', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_12);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2007);
      expect(doc.isCompatibilityMode()).toBe(true);

      const report = doc.upgradeToModernFormat();

      expect(report.changed).toBe(true);
      expect(report.previousMode).toBe(CompatibilityMode.Word2007);
      expect(report.newMode).toBe(15);
      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);
      expect(doc.isCompatibilityMode()).toBe(false);

      doc.dispose();
    });

    it('should upgrade mode 14 document and remove legacy flags', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_14);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2010);
      const info = doc.getCompatibilityInfo();
      expect(info.legacyFlags).toContain('useFELayout');
      expect(info.legacyFlags).toContain('adjustLineHeightInTable');

      const report = doc.upgradeToModernFormat();

      expect(report.changed).toBe(true);
      expect(report.previousMode).toBe(CompatibilityMode.Word2010);
      expect(report.removedFlags).toContain('useFELayout');
      expect(report.removedFlags).toContain('adjustLineHeightInTable');
      expect(doc.getCompatibilityInfo().legacyFlags).toHaveLength(0);

      doc.dispose();
    });

    it('should be a no-op for mode 15 documents with no legacy flags', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_15);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);

      const report = doc.upgradeToModernFormat();

      expect(report.changed).toBe(false);
      expect(report.removedFlags).toHaveLength(0);
      expect(report.addedSettings).toHaveLength(0);

      doc.dispose();
    });

    it('should create compat block when absent', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_NO_COMPAT);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2007); // default

      const report = doc.upgradeToModernFormat();

      expect(report.changed).toBe(true);
      expect(report.previousMode).toBe(CompatibilityMode.Word2007);
      expect(report.addedSettings.length).toBeGreaterThan(0);
      expect(report.addedSettings).toContain('compatibilityMode');
      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);

      doc.dispose();
    });

    it('should persist upgrade through save/reload cycle', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_12);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      doc.upgradeToModernFormat();
      const savedBuffer = await doc.toBuffer();
      doc.dispose();

      // Reload and verify
      const doc2 = await Document.loadFromBuffer(savedBuffer, { revisionHandling: 'preserve' });
      expect(doc2.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);
      expect(doc2.isCompatibilityMode()).toBe(false);

      // Verify the settings.xml has modern compat settings
      const zip = new ZipHandler();
      await zip.loadFromBuffer(savedBuffer);
      const settingsXml = zip.getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('compatibilityMode');
      expect(settingsXml).toContain('w:val="15"');
      expect(settingsXml).not.toMatch(/<w:useFELayout/);

      doc2.dispose();
    });

    it('should preserve extra compat settings during upgrade', async () => {
      // Use a valid compatSetting name with a non-Microsoft URI to test that
      // non-Microsoft settings survive the upgrade. The name must be valid per
      // OOXML schema to pass validation.
      const settingsWithCustomUri = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:useFELayout/>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/>
    <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://custom.vendor.com/word" w:val="1"/>
  </w:compat>
</w:settings>`;

      const buffer = await createDocxWithSettings(settingsWithCustomUri);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const report = doc.upgradeToModernFormat();
      expect(report.changed).toBe(true);
      expect(report.removedFlags).toContain('useFELayout');

      // Save and check that non-Microsoft URI setting was preserved
      const savedBuffer = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(savedBuffer);
      const settingsXml = zip.getFileAsString(DOCX_PATHS.SETTINGS);

      expect(settingsXml).toContain('http://custom.vendor.com/word');
      expect(settingsXml).toContain('w:val="15"');

      doc.dispose();
    });

    it('should expand namespaces during upgrade', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_12);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const report = doc.upgradeToModernFormat();

      // Namespaces should be expanded if they weren't already present
      expect(report.namespacesExpanded).toBeDefined();
      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);

      doc.dispose();
    });

    it('should preserve other settings.xml elements during upgrade', async () => {
      const settingsWithExtras = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="150"/>
  <w:proofState w:spelling="clean"/>
  <w:trackRevisions/>
  <w:defaultTabStop w:val="360"/>
  <w:evenAndOddHeaders/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:adjustLineHeightInTable/>
    <w:useFELayout/>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>
  </w:compat>
  <w:decimalSymbol w:val="."/>
  <w:listSeparator w:val=","/>
</w:settings>`;

      const buffer = await createDocxWithSettings(settingsWithExtras);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      doc.upgradeToModernFormat();
      const savedBuffer = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(savedBuffer);
      const settingsXml = zip.getFileAsString(DOCX_PATHS.SETTINGS);

      // Verify non-compat elements are preserved
      expect(settingsXml).toContain('w:zoom');
      expect(settingsXml).toContain('w:evenAndOddHeaders');
      expect(settingsXml).toContain('w:decimalSymbol');
      expect(settingsXml).toContain('w:listSeparator');
      expect(settingsXml).toContain('w:trackRevisions');
      // Legacy flags should be gone
      expect(settingsXml).not.toMatch(/<w:useFELayout/);
      expect(settingsXml).not.toMatch(/<w:adjustLineHeightInTable/);

      doc.dispose();
    });

    it('should report accurate removedFlags matching actual document', async () => {
      // Settings with many legacy flags
      const settingsWithManyFlags = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:noTabHangInd/>
    <w:noLeading/>
    <w:spaceForUL/>
    <w:doNotExpandShiftReturn/>
    <w:adjustLineHeightInTable/>
    <w:adjustLineHeightInTable/>
    <w:useFELayout/>
    <w:cachedColBalance/>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/>
  </w:compat>
</w:settings>`;

      const buffer = await createDocxWithSettings(settingsWithManyFlags);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      expect(doc.getCompatibilityMode()).toBe(11);

      const report = doc.upgradeToModernFormat();

      expect(report.changed).toBe(true);
      expect(report.previousMode).toBe(11);
      expect(report.removedFlags).toHaveLength(8);
      expect(report.removedFlags).toContain('noTabHangInd');
      expect(report.removedFlags).toContain('noLeading');
      expect(report.removedFlags).toContain('spaceForUL');
      expect(report.removedFlags).toContain('doNotExpandShiftReturn');
      expect(report.removedFlags).toContain('adjustLineHeightInTable');
      expect(report.removedFlags).toContain('useFELayout');
      expect(report.removedFlags).toContain('adjustLineHeightInTable');
      expect(report.removedFlags).toContain('cachedColBalance');

      doc.dispose();
    });

    it('should upgrade then allow further modifications before save', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_MODE_12);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Upgrade first
      doc.upgradeToModernFormat();

      // Then make modifications
      doc.enableTrackChanges({ author: 'Test Author' });
      doc.addParagraph(new Paragraph().addText('New content after upgrade'));

      const savedBuffer = await doc.toBuffer();

      // Reload and verify both upgrade and modifications persisted
      const doc2 = await Document.loadFromBuffer(savedBuffer, { revisionHandling: 'preserve' });
      expect(doc2.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);
      expect(doc2.getParagraphs().length).toBeGreaterThanOrEqual(2);

      // Verify settings.xml has both upgrade and track changes
      const zip = new ZipHandler();
      await zip.loadFromBuffer(savedBuffer);
      const settingsXml = zip.getFileAsString(DOCX_PATHS.SETTINGS);
      expect(settingsXml).toContain('w:val="15"');
      expect(settingsXml).toContain('w:trackRevisions');

      doc.dispose();
      doc2.dispose();
    });

    it('should handle new document upgrade gracefully', () => {
      const doc = Document.create();

      // New documents default to mode 15
      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2013Plus);

      // Upgrade should be a no-op
      const report = doc.upgradeToModernFormat();
      expect(report.changed).toBe(false);

      doc.dispose();
    });
  });

  // ======================================================================
  // Phase 3: CompatibilityUpgrader unit tests
  // ======================================================================
  describe('Phase 3: CompatibilityUpgrader', () => {

    it('should contain 65 legacy compat elements in catalog', () => {
      expect(LEGACY_COMPAT_ELEMENTS).toHaveLength(65);
      expect(LEGACY_COMPAT_ELEMENT_NAMES.size).toBe(65);
    });

    it('should have all elements prefixed with w:', () => {
      for (const elem of LEGACY_COMPAT_ELEMENTS) {
        expect(elem.startsWith('w:')).toBe(true);
      }
    });

    it('should have 5 modern compat settings (excluding useWord2013TrackBottomHyphenation)', () => {
      expect(MODERN_COMPAT_SETTINGS).toHaveLength(5);
      expect(MODERN_COMPAT_SETTINGS[0]!.name).toBe('compatibilityMode');
      expect(MODERN_COMPAT_SETTINGS[0]!.val).toBe('15');
    });

    it('should upgrade XML with legacy flags', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat>
    <w:adjustLineHeightInTable/>
    <w:useFELayout/>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/>
  </w:compat>
</w:settings>`;

      const result = CompatibilityUpgrader.upgradeCompatBlock(xml, 12);

      expect(result.report.changed).toBe(true);
      expect(result.report.removedFlags).toContain('useFELayout');
      expect(result.report.removedFlags).toContain('adjustLineHeightInTable');
      expect(result.xml).toContain('w:val="15"');
      expect(result.xml).not.toContain('useFELayout');
      expect(result.xml).not.toContain('adjustLineHeightInTable');
    });

    it('should insert compat block when missing', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`;

      const result = CompatibilityUpgrader.upgradeCompatBlock(xml, 12);

      expect(result.report.changed).toBe(true);
      expect(result.xml).toContain('<w:compat>');
      expect(result.xml).toContain('compatibilityMode');
      expect(result.xml).toContain('w:val="15"');
    });

    it('should preserve non-Microsoft URI settings', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/>
    <w:compatSetting w:name="vendorSetting" w:uri="http://other.vendor.com" w:val="custom"/>
  </w:compat>
</w:settings>`;

      const result = CompatibilityUpgrader.upgradeCompatBlock(xml, 12);

      expect(result.xml).toContain('vendorSetting');
      expect(result.xml).toContain('http://other.vendor.com');
      expect(result.xml).toContain('w:val="15"');
    });

    it('should expand modern namespaces', () => {
      const namespaces: Record<string, string> = {
        'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      };

      const result = CompatibilityUpgrader.ensureModernNamespaces(namespaces);

      expect(result.expanded).toBe(true);
      expect(result.namespaces['xmlns:w14']).toBeDefined();
      expect(result.namespaces['xmlns:w15']).toBeDefined();
      expect(result.namespaces['xmlns:w16se']).toBeDefined();
      expect(result.namespaces['xmlns:w16cid']).toBeDefined();
    });

    it('should not expand namespaces if already present', () => {
      const namespaces: Record<string, string> = {
        'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'xmlns:w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
        'xmlns:w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
      };

      const result = CompatibilityUpgrader.ensureModernNamespaces(namespaces);

      expect(result.expanded).toBe(false);
    });

    it('should handle self-closing w:compat tag', () => {
      const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:compat/>
</w:settings>`;

      const result = CompatibilityUpgrader.upgradeCompatBlock(xml, 12);

      expect(result.report.changed).toBe(true);
      expect(result.xml).toContain('<w:compat>');
      expect(result.xml).toContain('compatibilityMode');
    });
  });

  // ======================================================================
  describe('dispose() cleanup', () => {

    it('should reset all settings-related fields on dispose', async () => {
      const buffer = await createDocxWithSettings(SETTINGS_WITH_TRACKING_AND_PROTECTION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Verify fields are populated
      expect(doc.isProtected()).toBe(true);
      expect(doc.getRsids().length).toBeGreaterThan(0);
      expect(doc.getCompatibilityMode()).toBe(15);

      doc.dispose();

      // After dispose, compatibility mode should return default (12 since no _compatInfo)
      expect(doc.getCompatibilityMode()).toBe(CompatibilityMode.Word2007);
      expect(doc.isCompatibilityMode()).toBe(true);
      expect(doc.isProtected()).toBe(false);
      expect(doc.getRsids().length).toBe(0);
      expect(doc.getRsidRoot()).toBeUndefined();
    });
  });
});
