/**
 * Styles gap tests: latent styles, personalCompose/personalReply
 * Phase 5 of ECMA-376 gap analysis
 */

import { Style } from '../../src/formatting/Style';
import { StylesManager } from '../../src/formatting/StylesManager';
import { Document } from '../../src/core/Document';

describe('Style personalCompose', () => {
  test('should set and get personalCompose', () => {
    const style = Style.create({
      styleId: 'PersonalComposeStyle',
      name: 'Personal Compose Style',
      type: 'paragraph',
    });
    style.setPersonalCompose(true);
    expect(style.getPersonalCompose()).toBe(true);
  });

  test('should default personalCompose to false', () => {
    const style = Style.create({
      styleId: 'TestStyle',
      name: 'Test',
      type: 'paragraph',
    });
    expect(style.getPersonalCompose()).toBe(false);
  });

  test('should generate w:personalCompose in XML', () => {
    const style = Style.create({
      styleId: 'ComposeStyle',
      name: 'Compose',
      type: 'paragraph',
      personalCompose: true,
    });

    const xml = style.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('w:personalCompose');
  });

  test('should not generate w:personalCompose when false', () => {
    const style = Style.create({
      styleId: 'RegularStyle',
      name: 'Regular',
      type: 'paragraph',
    });

    const xml = style.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).not.toContain('w:personalCompose');
  });
});

describe('Style personalReply', () => {
  test('should set and get personalReply', () => {
    const style = Style.create({
      styleId: 'PersonalReplyStyle',
      name: 'Personal Reply Style',
      type: 'paragraph',
    });
    style.setPersonalReply(true);
    expect(style.getPersonalReply()).toBe(true);
  });

  test('should generate w:personalReply in XML', () => {
    const style = Style.create({
      styleId: 'ReplyStyle',
      name: 'Reply',
      type: 'paragraph',
      personalReply: true,
    });

    const xml = style.toXML();
    const xmlStr = JSON.stringify(xml);
    expect(xmlStr).toContain('w:personalReply');
  });
});

describe('Latent Styles', () => {
  test('should set and get latent styles config', () => {
    const sm = StylesManager.create();
    sm.setLatentStyles({
      defaultLockedState: false,
      defaultUiPriority: 99,
      defaultSemiHidden: true,
      defaultUnhideWhenUsed: true,
      defaultQFormat: false,
      count: 376,
    });

    const config = sm.getLatentStyles();
    expect(config?.defaultUiPriority).toBe(99);
    expect(config?.defaultSemiHidden).toBe(true);
    expect(config?.count).toBe(376);
  });

  test('should add and retrieve latent style exceptions', () => {
    const sm = StylesManager.create();
    sm.addLatentStyleException({ name: 'Normal', qFormat: true, uiPriority: 0 });
    sm.addLatentStyleException({
      name: 'heading 1',
      qFormat: true,
      semiHidden: false,
      uiPriority: 9,
    });

    const exceptions = sm.getLatentStyleExceptions();
    expect(exceptions).toHaveLength(2);
    expect(exceptions[0]!.name).toBe('Normal');
    expect(exceptions[1]!.uiPriority).toBe(9);
  });

  test('should replace existing exception for same name', () => {
    const sm = StylesManager.create();
    sm.addLatentStyleException({ name: 'Normal', qFormat: true, uiPriority: 0 });
    sm.addLatentStyleException({ name: 'Normal', qFormat: false, uiPriority: 5 });

    const exceptions = sm.getLatentStyleExceptions();
    expect(exceptions).toHaveLength(1);
    expect(exceptions[0]!.uiPriority).toBe(5);
    expect(exceptions[0]!.qFormat).toBe(false);
  });

  test('should generate latent styles in XML', () => {
    const sm = StylesManager.create();
    sm.setLatentStyles({
      defaultLockedState: false,
      defaultUiPriority: 99,
      defaultSemiHidden: true,
      defaultUnhideWhenUsed: true,
      count: 376,
    });
    sm.addLatentStyleException({ name: 'Normal', qFormat: true, uiPriority: 0 });

    const xml = sm.generateStylesXml();
    expect(xml).toContain('w:latentStyles');
    expect(xml).toContain('w:defUIPriority="99"');
    expect(xml).toContain('w:lsdException');
    expect(xml).toContain('w:name="Normal"');
  });

  test('should generate valid OOXML with latent styles in document', async () => {
    const doc = Document.create();
    const sm = doc.getStylesManager();
    sm.setLatentStyles({
      defaultLockedState: false,
      defaultUiPriority: 99,
      defaultSemiHidden: true,
      defaultUnhideWhenUsed: true,
      count: 376,
    });
    sm.addLatentStyleException({ name: 'Normal', qFormat: true, uiPriority: 0 });
    sm.addLatentStyleException({
      name: 'heading 1',
      qFormat: true,
      semiHidden: false,
      unhideWhenUsed: false,
      uiPriority: 9,
    });

    doc.createParagraph('Document with latent styles');

    // toBuffer() triggers OOXML validation
    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);

    doc.dispose();
  });
});

describe('Style personalCompose/Reply round-trip', () => {
  test('should round-trip personalCompose through document', async () => {
    const doc = Document.create();
    const sm = doc.getStylesManager();

    const style = Style.create({
      styleId: 'ComposeStyle',
      name: 'Compose Style',
      type: 'paragraph',
      personalCompose: true,
      personal: true,
    });
    sm.addStyle(style);

    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);
    const loadedSm = loaded.getStylesManager();
    const loadedStyle = loadedSm.getStyle('ComposeStyle');

    expect(loadedStyle).toBeDefined();
    expect(loadedStyle?.getPersonalCompose()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });

  test('should round-trip personalReply through document', async () => {
    const doc = Document.create();
    const sm = doc.getStylesManager();

    const style = Style.create({
      styleId: 'ReplyStyle',
      name: 'Reply Style',
      type: 'paragraph',
      personalReply: true,
      personal: true,
    });
    sm.addStyle(style);

    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);
    const loadedSm = loaded.getStylesManager();
    const loadedStyle = loadedSm.getStyle('ReplyStyle');

    expect(loadedStyle).toBeDefined();
    expect(loadedStyle?.getPersonalReply()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('Style rFonts full attribute parsing (ECMA-376 §17.3.2.26)', () => {
  test('should round-trip all w:rFonts attributes through style definitions', async () => {
    // Create a document, then inject a style with full rFonts into the XML
    const doc = Document.create();
    doc.createParagraph('Test');
    const buffer = await doc.toBuffer();
    doc.dispose();

    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buffer);
    let stylesXml = await zip.file('word/styles.xml')!.async('string');

    // Inject a character style with comprehensive rFonts
    const customStyle = `<w:style w:type="character" w:styleId="CJKMixed">
      <w:name w:val="CJK Mixed"/>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei" w:cs="Arial" w:hint="eastAsia" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>
      </w:rPr>
    </w:style>`;
    stylesXml = stylesXml.replace('</w:styles>', `${customStyle}</w:styles>`);
    zip.file('word/styles.xml', stylesXml);
    const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    const loaded = await Document.loadFromBuffer(modifiedBuffer);
    const stylesManager = (loaded as any).stylesManager;
    const style = stylesManager?.getStyle('CJKMixed');
    const runFmt = style?.getRunFormatting();

    expect(runFmt).toBeDefined();
    expect(runFmt?.font).toBe('Calibri');
    expect(runFmt?.fontHAnsi).toBe('Calibri');
    expect(runFmt?.fontEastAsia).toBe('Microsoft YaHei');
    expect(runFmt?.fontCs).toBe('Arial');
    expect(runFmt?.fontHint).toBe('eastAsia');
    expect(runFmt?.fontAsciiTheme).toBe('minorHAnsi');
    expect(runFmt?.fontHAnsiTheme).toBe('minorHAnsi');
    expect(runFmt?.fontEastAsiaTheme).toBe('minorEastAsia');
    expect(runFmt?.fontCsTheme).toBe('minorBidi');

    loaded.dispose();
  });
});

describe('Style w:color theme attributes parsing (ECMA-376 §17.3.2.6)', () => {
  test('should round-trip themeColor, themeTint, themeShade through style definitions', async () => {
    const doc = Document.create();
    doc.createParagraph('Test');
    const buffer = await doc.toBuffer();
    doc.dispose();

    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buffer);
    let stylesXml = await zip.file('word/styles.xml')!.async('string');

    // Inject a character style with theme color attributes
    const customStyle = `<w:style w:type="character" w:styleId="ThemeAccent">
      <w:name w:val="Theme Accent"/>
      <w:rPr>
        <w:color w:val="4472C4" w:themeColor="accent1" w:themeTint="BF" w:themeShade="A0"/>
      </w:rPr>
    </w:style>`;
    stylesXml = stylesXml.replace('</w:styles>', `${customStyle}</w:styles>`);
    zip.file('word/styles.xml', stylesXml);
    const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    const loaded = await Document.loadFromBuffer(modifiedBuffer);
    const stylesManager = (loaded as any).stylesManager;
    const style = stylesManager?.getStyle('ThemeAccent');
    const runFmt = style?.getRunFormatting();

    expect(runFmt).toBeDefined();
    expect(runFmt?.color).toBe('4472C4');
    expect(runFmt?.themeColor).toBe('accent1');
    // themeTint "BF" = 191 decimal
    expect(runFmt?.themeTint).toBe(0xbf);
    // themeShade "A0" = 160 decimal
    expect(runFmt?.themeShade).toBe(0xa0);

    loaded.dispose();
  });
});

describe('Style w:u underline theme attributes parsing (ECMA-376 §17.3.2.40)', () => {
  test('should round-trip underline color and theme attributes through style definitions', async () => {
    const doc = Document.create();
    doc.createParagraph('Test');
    const buffer = await doc.toBuffer();
    doc.dispose();

    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buffer);
    let stylesXml = await zip.file('word/styles.xml')!.async('string');

    const customStyle = `<w:style w:type="character" w:styleId="ThemedUnderline">
      <w:name w:val="Themed Underline"/>
      <w:rPr>
        <w:u w:val="single" w:color="4472C4" w:themeColor="accent1" w:themeTint="BF" w:themeShade="A0"/>
      </w:rPr>
    </w:style>`;
    stylesXml = stylesXml.replace('</w:styles>', `${customStyle}</w:styles>`);
    zip.file('word/styles.xml', stylesXml);
    const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    const loaded = await Document.loadFromBuffer(modifiedBuffer);
    const stylesManager = (loaded as any).stylesManager;
    const style = stylesManager?.getStyle('ThemedUnderline');
    const runFmt = style?.getRunFormatting();

    expect(runFmt).toBeDefined();
    expect(runFmt?.underline).toBe('single');
    expect(runFmt?.underlineColor).toBe('4472C4');
    expect(runFmt?.underlineThemeColor).toBe('accent1');
    expect(runFmt?.underlineThemeTint).toBe(0xbf);
    expect(runFmt?.underlineThemeShade).toBe(0xa0);

    loaded.dispose();
  });
});

describe('Style w:szCs parsing (ECMA-376 §17.3.2.40)', () => {
  test('should parse complex script font size from style definitions', async () => {
    const doc = Document.create();
    doc.createParagraph('Test');
    const buffer = await doc.toBuffer();
    doc.dispose();

    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buffer);
    let stylesXml = await zip.file('word/styles.xml')!.async('string');

    // Inject style with szCs different from sz
    const customStyle = `<w:style w:type="character" w:styleId="ArabicText">
      <w:name w:val="Arabic Text"/>
      <w:rPr>
        <w:sz w:val="24"/>
        <w:szCs w:val="28"/>
      </w:rPr>
    </w:style>`;
    stylesXml = stylesXml.replace('</w:styles>', `${customStyle}</w:styles>`);
    zip.file('word/styles.xml', stylesXml);
    const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    const loaded = await Document.loadFromBuffer(modifiedBuffer);
    const stylesManager = (loaded as any).stylesManager;
    const style = stylesManager?.getStyle('ArabicText');
    const runFmt = style?.getRunFormatting();

    expect(runFmt).toBeDefined();
    expect(runFmt?.size).toBe(12); // 24 half-points = 12pt
    expect(runFmt?.sizeCs).toBe(14); // 28 half-points = 14pt

    loaded.dispose();
  });
});

describe('Style paragraph borders and tabs parsing', () => {
  test('should parse paragraph borders from style definitions', async () => {
    const doc = Document.create();
    doc.createParagraph('Test');
    const buffer = await doc.toBuffer();
    doc.dispose();

    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buffer);
    let stylesXml = await zip.file('word/styles.xml')!.async('string');

    const customStyle = `<w:style w:type="paragraph" w:styleId="CodeBlock">
      <w:name w:val="Code Block"/>
      <w:pPr>
        <w:pBdr>
          <w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/>
          <w:left w:val="single" w:sz="4" w:space="4" w:color="000000"/>
          <w:right w:val="single" w:sz="4" w:space="4" w:color="000000"/>
          <w:between w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/>
        </w:pBdr>
      </w:pPr>
    </w:style>`;
    stylesXml = stylesXml.replace('</w:styles>', `${customStyle}</w:styles>`);
    zip.file('word/styles.xml', stylesXml);
    const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    const loaded = await Document.loadFromBuffer(modifiedBuffer);
    const stylesManager = (loaded as any).stylesManager;
    const style = stylesManager?.getStyle('CodeBlock');
    const paraFmt = style?.getParagraphFormatting();

    expect(paraFmt).toBeDefined();
    expect(paraFmt?.borders).toBeDefined();
    expect(paraFmt?.borders?.top?.style).toBe('single');
    expect(paraFmt?.borders?.top?.size).toBe(4);
    expect(paraFmt?.borders?.bottom?.style).toBe('single');
    expect(paraFmt?.borders?.left?.space).toBe(4);
    expect(paraFmt?.borders?.between?.color).toBe('CCCCCC');

    loaded.dispose();
  });

  test('should parse tab stops from style definitions', async () => {
    const doc = Document.create();
    doc.createParagraph('Test');
    const buffer = await doc.toBuffer();
    doc.dispose();

    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buffer);
    let stylesXml = await zip.file('word/styles.xml')!.async('string');

    const customStyle = `<w:style w:type="paragraph" w:styleId="TabbedPara">
      <w:name w:val="Tabbed Paragraph"/>
      <w:pPr>
        <w:tabs>
          <w:tab w:val="left" w:pos="720"/>
          <w:tab w:val="center" w:pos="4320"/>
          <w:tab w:val="right" w:pos="8640" w:leader="dot"/>
        </w:tabs>
      </w:pPr>
    </w:style>`;
    stylesXml = stylesXml.replace('</w:styles>', `${customStyle}</w:styles>`);
    zip.file('word/styles.xml', stylesXml);
    const modifiedBuffer = await zip.generateAsync({ type: 'nodebuffer' });

    const loaded = await Document.loadFromBuffer(modifiedBuffer);
    const stylesManager = (loaded as any).stylesManager;
    const style = stylesManager?.getStyle('TabbedPara');
    const paraFmt = style?.getParagraphFormatting();

    expect(paraFmt).toBeDefined();
    expect(paraFmt?.tabs).toBeDefined();
    expect(paraFmt?.tabs).toHaveLength(3);
    expect(paraFmt?.tabs?.[0]).toMatchObject({ val: 'left', position: 720 });
    expect(paraFmt?.tabs?.[1]).toMatchObject({ val: 'center', position: 4320 });
    expect(paraFmt?.tabs?.[2]).toMatchObject({ val: 'right', position: 8640, leader: 'dot' });

    loaded.dispose();
  });
});
