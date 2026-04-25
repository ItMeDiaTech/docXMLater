/**
 * Example: Compatibility mode.
 *
 * Word stores a `compatibilityMode` setting (per `w:compat` in
 * settings.xml) that controls layout behaviour bug-compatibility with
 * older versions: 11=Word2003, 12=Word2007 (default), 14=Word2010,
 * 15=Word2013+. The legacy 1st-edition `<w:compat>` block carries 65
 * boolean flags; the 2nd-edition replaces most with explicit
 * `<w:compatSetting>` entries.
 *
 * docxmlater can:
 *   - read the current mode (`getCompatibilityMode`,
 *     `getCompatibilityInfo`)
 *   - upgrade legacy documents to the 2nd-edition format
 *     (`upgradeToModernFormat`) — replaces the 65 legacy boolean flags
 *     with the modern `compatSetting` form Word 2013+ writes natively,
 *     preserving non-Microsoft custom settings.
 *
 * Run: `npx ts-node examples/20-compatibility-mode/compatibility-mode.ts <input.docx>`
 */

import { Document } from '../../src';

async function main() {
  const inputPath = process.argv[2];
  if (!inputPath) {
    // eslint-disable-next-line no-console
    console.log(
      'Usage: ts-node examples/20-compatibility-mode/compatibility-mode.ts <input.docx>'
    );
    // eslint-disable-next-line no-console
    console.log('Demonstrating with a freshly-created document instead:');
    const fresh = Document.create();
    fresh.createParagraph('New document').setStyle('Heading1');
    const info = fresh.getCompatibilityInfo();
    // eslint-disable-next-line no-console
    console.log('  mode:', info.mode);
    // eslint-disable-next-line no-console
    console.log('  isLegacyMode:', info.isLegacyMode);
    // eslint-disable-next-line no-console
    console.log('  legacyFlags:', info.legacyFlags);
    // eslint-disable-next-line no-console
    console.log('  compatSettings count:', info.compatSettings.length);
    fresh.dispose();
    return;
  }

  const doc = await Document.load(inputPath);

  const before = doc.getCompatibilityInfo();
  // eslint-disable-next-line no-console
  console.log('Before upgrade:');
  // eslint-disable-next-line no-console
  console.log('  mode:', before.mode);
  // eslint-disable-next-line no-console
  console.log('  isLegacyMode:', before.isLegacyMode);
  // eslint-disable-next-line no-console
  console.log('  legacyFlags (count):', before.legacyFlags.length);
  // eslint-disable-next-line no-console
  console.log('  compatSettings (count):', before.compatSettings.length);

  const report = doc.upgradeToModernFormat();
  // eslint-disable-next-line no-console
  console.log('Upgrade report:');
  // eslint-disable-next-line no-console
  console.log('  changed:', report.changed);
  // eslint-disable-next-line no-console
  console.log('  previousMode → newMode:', report.previousMode, '→', report.newMode);
  // eslint-disable-next-line no-console
  console.log('  removedFlags:', report.removedFlags);
  // eslint-disable-next-line no-console
  console.log('  addedSettings:', report.addedSettings);
  // eslint-disable-next-line no-console
  console.log('  namespacesExpanded:', report.namespacesExpanded);

  await doc.save('examples/20-compatibility-mode/upgraded.docx');
  doc.dispose();
  // eslint-disable-next-line no-console
  console.log('Wrote examples/20-compatibility-mode/upgraded.docx');
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exit(1);
});
