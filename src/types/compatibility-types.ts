/**
 * Word compatibility mode versions.
 *
 * These map to Microsoft's internal version numbers used in the
 * w:compatSetting name="compatibilityMode" value within settings.xml.
 *
 * Per MS-DOCX specification, the default when absent is 12 (Word 2007).
 * See: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/90138c4d-eb18-4edc-aa6c-dfb799cb1d0d
 */
export enum CompatibilityMode {
  /** Word 2003 and earlier — uses MS-DOC feature set */
  Word2003 = 11,
  /** Word 2007 (ECMA-376 1st edition) — DEFAULT when absent */
  Word2007 = 12,
  /** Word 2010 (ISO/IEC 29500 with exclusions) */
  Word2010 = 14,
  /** Word 2013/2016/2019/365 (full modern features) */
  Word2013Plus = 15,
}

/**
 * A single w:compatSetting entry (name/uri/val triple).
 *
 * These are the modern extensibility mechanism for compatibility settings,
 * added in ECMA-376 2nd edition. They coexist with legacy boolean flags
 * in the Transitional schema.
 */
export interface CompatSetting {
  name: string;
  uri: string;
  val: string;
}

/**
 * Parsed compatibility settings from the w:compat block in settings.xml.
 */
export interface CompatibilityInfo {
  /** The numeric compatibility mode value (11, 12, 14, or 15) */
  mode: CompatibilityMode;

  /** Whether the document is in a legacy compatibility mode (below 15) */
  isLegacyMode: boolean;

  /** Modern w:compatSetting entries (name/uri/val triples) */
  compatSettings: CompatSetting[];

  /** Legacy boolean compat flags that are present and enabled (without w: prefix) */
  legacyFlags: string[];
}
