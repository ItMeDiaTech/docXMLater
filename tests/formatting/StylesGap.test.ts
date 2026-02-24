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
