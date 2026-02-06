/**
 * CompatibilityUpgrader — Upgrades w:compat block in settings.xml to modern format.
 *
 * Equivalent to Word's File > Info > Convert operation for the compatibility
 * settings portion. Removes legacy boolean compat flags, updates compatSetting
 * entries, and manages modern namespaces.
 *
 * @module
 */

import {
  LEGACY_COMPAT_ELEMENT_NAMES,
  MODERN_COMPAT_SETTINGS,
  MS_WORD_COMPAT_URI,
} from '../constants/legacyCompatFlags';
import type { CompatSetting } from '../types/compatibility-types';

/**
 * Report of what changed during an upgrade operation.
 */
export interface UpgradeReport {
  /** Previous compatibility mode value (11, 12, 14, or 15) */
  previousMode: number;

  /** New compatibility mode (always 15 after upgrade) */
  newMode: number;

  /** Legacy compat flags that were removed (without w: prefix) */
  removedFlags: string[];

  /** Modern settings that were added (setting names) */
  addedSettings: string[];

  /** Whether namespaces were expanded */
  namespacesExpanded: boolean;

  /** Whether any changes were actually made (false if already mode 15 with no legacy flags) */
  changed: boolean;
}

/**
 * Upgrades w:compat blocks in settings.xml to modern Word 2013+ format.
 */
export class CompatibilityUpgrader {
  /**
   * Upgrades the w:compat block in settings.xml to mode 15.
   *
   * Operations performed:
   * 1. Remove all legacy boolean compat elements
   * 2. Remove existing Microsoft-URI w:compatSetting entries
   * 3. Insert modern w:compatSetting entries
   * 4. Preserve non-Microsoft w:compatSetting entries (custom URIs)
   *
   * @param settingsXml - The raw settings.xml content
   * @param previousMode - The current compatibility mode value (for the report)
   * @returns Object with upgraded XML and an UpgradeReport
   */
  static upgradeCompatBlock(
    settingsXml: string,
    previousMode: number = 12
  ): { xml: string; report: UpgradeReport } {
    const removedFlags: string[] = [];
    const addedSettings: string[] = [];

    // Check if already mode 15 with no legacy flags
    const compatBlockMatch = settingsXml.match(/<w:compat>([\s\S]*?)<\/w:compat>/);
    const selfClosingCompat = settingsXml.match(/<w:compat\s*\/>/);

    if (!compatBlockMatch && !selfClosingCompat) {
      // No w:compat block — insert a full modern block before </w:settings>
      const modernBlock = CompatibilityUpgrader.buildModernCompatBlock(MODERN_COMPAT_SETTINGS);
      const xml = settingsXml.replace(
        /<\/w:settings>/,
        modernBlock + '\n</w:settings>'
      );

      for (const s of MODERN_COMPAT_SETTINGS) {
        addedSettings.push(s.name);
      }

      return {
        xml,
        report: {
          previousMode,
          newMode: 15,
          removedFlags,
          addedSettings,
          namespacesExpanded: false,
          changed: true,
        },
      };
    }

    // Handle self-closing <w:compat/>
    if (selfClosingCompat) {
      const modernBlock = CompatibilityUpgrader.buildModernCompatBlock(MODERN_COMPAT_SETTINGS);
      const xml = settingsXml.replace(/<w:compat\s*\/>/, modernBlock);

      for (const s of MODERN_COMPAT_SETTINGS) {
        addedSettings.push(s.name);
      }

      return {
        xml,
        report: {
          previousMode,
          newMode: 15,
          removedFlags,
          addedSettings,
          namespacesExpanded: false,
          changed: true,
        },
      };
    }

    // Process existing w:compat block
    const compatBlock = compatBlockMatch![1] ?? '';

    // 1. Find and record legacy flags that will be removed
    const legacyFlagRegex = /<w:(\w+)(?:\s[^>]*)?\/?>/g;
    let flagMatch;
    while ((flagMatch = legacyFlagRegex.exec(compatBlock)) !== null) {
      const tagName = flagMatch[1];
      if (tagName && tagName !== 'compatSetting' && LEGACY_COMPAT_ELEMENT_NAMES.has(tagName)) {
        removedFlags.push(tagName);
      }
    }

    // 2. Collect non-Microsoft compatSetting entries to preserve
    const preservedSettings: CompatSetting[] = [];
    const settingRegex = /<w:compatSetting\s+([^>]*)\/?\s*>/g;
    let settingMatch;
    while ((settingMatch = settingRegex.exec(compatBlock)) !== null) {
      const attrs = settingMatch[1] ?? '';
      const nameMatch = attrs.match(/w:name\s*=\s*"([^"]*)"/);
      const uriMatch = attrs.match(/w:uri\s*=\s*"([^"]*)"/);
      const valMatch = attrs.match(/w:val\s*=\s*"([^"]*)"/);

      if (nameMatch?.[1] && uriMatch?.[1] && valMatch?.[1]) {
        // Preserve settings with non-Microsoft URIs
        if (uriMatch[1] !== MS_WORD_COMPAT_URI) {
          preservedSettings.push({
            name: nameMatch[1],
            uri: uriMatch[1],
            val: valMatch[1],
          });
        }
      }
    }

    // 3. Determine which modern settings are actually new
    const existingMsSettings = new Set<string>();
    const existingSettingRegex = /<w:compatSetting\s+([^>]*)\/?\s*>/g;
    let existingMatch;
    while ((existingMatch = existingSettingRegex.exec(compatBlock)) !== null) {
      const attrs = existingMatch[1] ?? '';
      const nameMatch = attrs.match(/w:name\s*=\s*"([^"]*)"/);
      const uriMatch = attrs.match(/w:uri\s*=\s*"([^"]*)"/);
      if (nameMatch?.[1] && uriMatch?.[1] && uriMatch[1] === MS_WORD_COMPAT_URI) {
        existingMsSettings.add(nameMatch[1]);
      }
    }

    for (const s of MODERN_COMPAT_SETTINGS) {
      if (!existingMsSettings.has(s.name)) {
        addedSettings.push(s.name);
      }
    }
    // compatibilityMode value change counts as "added" if it existed but with different val
    if (existingMsSettings.has('compatibilityMode') && previousMode !== 15) {
      if (!addedSettings.includes('compatibilityMode')) {
        addedSettings.push('compatibilityMode');
      }
    }

    // 4. Build the new compat block
    const allSettings = [...MODERN_COMPAT_SETTINGS, ...preservedSettings];
    const modernBlock = CompatibilityUpgrader.buildModernCompatBlock(allSettings);

    // 5. Replace the old block
    const changed = removedFlags.length > 0 || addedSettings.length > 0 || previousMode !== 15;
    const xml = settingsXml.replace(/<w:compat>[\s\S]*?<\/w:compat>/, modernBlock);

    return {
      xml,
      report: {
        previousMode,
        newMode: 15,
        removedFlags,
        addedSettings,
        namespacesExpanded: false,
        changed,
      },
    };
  }

  /**
   * Ensures namespace declarations include all modern extended namespaces
   * (w14, w15, w16se, w16cid, etc.) and returns the expanded set.
   *
   * @param namespaces - Current namespace record from document
   * @returns Updated namespace record with modern namespaces added
   */
  static ensureModernNamespaces(
    namespaces: Record<string, string>
  ): { namespaces: Record<string, string>; expanded: boolean } {
    const result = { ...namespaces };
    let expanded = false;

    const modernNamespaces: Record<string, string> = {
      'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
      'xmlns:w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
      'xmlns:w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    };

    for (const [key, value] of Object.entries(modernNamespaces)) {
      if (!result[key]) {
        result[key] = value;
        expanded = true;
      }
    }

    return { namespaces: result, expanded };
  }

  /**
   * Builds a <w:compat> block XML string from an array of compat settings.
   * Does not include legacy boolean flags (they are stripped during upgrade).
   */
  private static buildModernCompatBlock(settings: CompatSetting[]): string {
    const entries = settings.map(
      s => `    <w:compatSetting w:name="${s.name}" w:uri="${s.uri}" w:val="${s.val}"/>`
    );
    return `<w:compat>\n${entries.join('\n')}\n  </w:compat>`;
  }
}
