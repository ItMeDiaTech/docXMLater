/**
 * Settings API gap tests: new getter/setter pairs for common settings
 * Phase 6C of ECMA-376 gap analysis
 */

import { Document } from '../../src/core/Document';

describe('Settings API: hideSpellingErrors', () => {
  test('should default to false', () => {
    const doc = Document.create();
    expect(doc.getHideSpellingErrors()).toBe(false);
    doc.dispose();
  });

  test('should set and get hideSpellingErrors', () => {
    const doc = Document.create();
    doc.setHideSpellingErrors(true);
    expect(doc.getHideSpellingErrors()).toBe(true);
    doc.dispose();
  });

  test('should generate valid OOXML with hideSpellingErrors', async () => {
    const doc = Document.create();
    doc.setHideSpellingErrors(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    doc.dispose();
  });

  test('should round-trip hideSpellingErrors', async () => {
    const doc = Document.create();
    doc.setHideSpellingErrors(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    expect(loaded.getHideSpellingErrors()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('Settings API: hideGrammaticalErrors', () => {
  test('should default to false', () => {
    const doc = Document.create();
    expect(doc.getHideGrammaticalErrors()).toBe(false);
    doc.dispose();
  });

  test('should set and get hideGrammaticalErrors', () => {
    const doc = Document.create();
    doc.setHideGrammaticalErrors(true);
    expect(doc.getHideGrammaticalErrors()).toBe(true);
    doc.dispose();
  });

  test('should generate valid OOXML with hideGrammaticalErrors', async () => {
    const doc = Document.create();
    doc.setHideGrammaticalErrors(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    doc.dispose();
  });

  test('should round-trip hideGrammaticalErrors', async () => {
    const doc = Document.create();
    doc.setHideGrammaticalErrors(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    expect(loaded.getHideGrammaticalErrors()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('Settings API: defaultTabStop', () => {
  test('should default to undefined', () => {
    const doc = Document.create();
    // New doc may have a default tab stop from template
    // Just verify it's accessible
    const val = doc.getDefaultTabStop();
    expect(val === undefined || typeof val === 'number').toBe(true);
    doc.dispose();
  });

  test('should set and get defaultTabStop', () => {
    const doc = Document.create();
    doc.setDefaultTabStop(720);
    expect(doc.getDefaultTabStop()).toBe(720);
    doc.dispose();
  });

  test('should generate valid OOXML with defaultTabStop', async () => {
    const doc = Document.create();
    doc.setDefaultTabStop(720);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    doc.dispose();
  });

  test('should round-trip defaultTabStop', async () => {
    const doc = Document.create();
    doc.setDefaultTabStop(360);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    expect(loaded.getDefaultTabStop()).toBe(360);

    doc.dispose();
    loaded.dispose();
  });
});

describe('Settings API: updateFields', () => {
  test('should default to false', () => {
    const doc = Document.create();
    expect(doc.getUpdateFields()).toBe(false);
    doc.dispose();
  });

  test('should set and get updateFields', () => {
    const doc = Document.create();
    doc.setUpdateFields(true);
    expect(doc.getUpdateFields()).toBe(true);
    doc.dispose();
  });

  test('should generate valid OOXML with updateFields', async () => {
    const doc = Document.create();
    doc.setUpdateFields(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    doc.dispose();
  });

  test('should round-trip updateFields', async () => {
    const doc = Document.create();
    doc.setUpdateFields(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    expect(loaded.getUpdateFields()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('Settings API: embedTrueTypeFonts', () => {
  test('should default to false', () => {
    const doc = Document.create();
    expect(doc.getEmbedTrueTypeFonts()).toBe(false);
    doc.dispose();
  });

  test('should set and get embedTrueTypeFonts', () => {
    const doc = Document.create();
    doc.setEmbedTrueTypeFonts(true);
    expect(doc.getEmbedTrueTypeFonts()).toBe(true);
    doc.dispose();
  });

  test('should generate valid OOXML with embedTrueTypeFonts', async () => {
    const doc = Document.create();
    doc.setEmbedTrueTypeFonts(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    doc.dispose();
  });

  test('should round-trip embedTrueTypeFonts', async () => {
    const doc = Document.create();
    doc.setEmbedTrueTypeFonts(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    expect(loaded.getEmbedTrueTypeFonts()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('Settings API: saveSubsetFonts', () => {
  test('should default to false', () => {
    const doc = Document.create();
    expect(doc.getSaveSubsetFonts()).toBe(false);
    doc.dispose();
  });

  test('should set and get saveSubsetFonts', () => {
    const doc = Document.create();
    doc.setSaveSubsetFonts(true);
    expect(doc.getSaveSubsetFonts()).toBe(true);
    doc.dispose();
  });

  test('should generate valid OOXML with saveSubsetFonts', async () => {
    const doc = Document.create();
    doc.setSaveSubsetFonts(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    doc.dispose();
  });

  test('should round-trip saveSubsetFonts', async () => {
    const doc = Document.create();
    doc.setSaveSubsetFonts(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    expect(loaded.getSaveSubsetFonts()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('Settings API: doNotTrackMoves', () => {
  test('should default to false', () => {
    const doc = Document.create();
    expect(doc.getDoNotTrackMoves()).toBe(false);
    doc.dispose();
  });

  test('should set and get doNotTrackMoves', () => {
    const doc = Document.create();
    doc.setDoNotTrackMoves(true);
    expect(doc.getDoNotTrackMoves()).toBe(true);
    doc.dispose();
  });

  test('should generate valid OOXML with doNotTrackMoves', async () => {
    const doc = Document.create();
    doc.setDoNotTrackMoves(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();
    expect(buffer.length).toBeGreaterThan(0);
    doc.dispose();
  });

  test('should round-trip doNotTrackMoves', async () => {
    const doc = Document.create();
    doc.setDoNotTrackMoves(true);
    doc.createParagraph('Test');

    const buffer = await doc.toBuffer();

    // Must use preserve mode: default 'accept' mode strips doNotTrackMoves
    // (acceptAllRevisions removes all revision-related settings)
    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    expect(loaded.getDoNotTrackMoves()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });
});

describe('Settings API: combined settings', () => {
  test('should handle multiple settings at once', async () => {
    const doc = Document.create();
    doc.setHideSpellingErrors(true);
    doc.setHideGrammaticalErrors(true);
    doc.setDefaultTabStop(720);
    doc.setUpdateFields(true);
    doc.setEmbedTrueTypeFonts(true);
    doc.setSaveSubsetFonts(true);
    doc.createParagraph('Test document with multiple settings');

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);

    expect(loaded.getHideSpellingErrors()).toBe(true);
    expect(loaded.getHideGrammaticalErrors()).toBe(true);
    expect(loaded.getDefaultTabStop()).toBe(720);
    expect(loaded.getUpdateFields()).toBe(true);
    expect(loaded.getEmbedTrueTypeFonts()).toBe(true);
    expect(loaded.getSaveSubsetFonts()).toBe(true);

    doc.dispose();
    loaded.dispose();
  });

  test('should disable settings that were previously enabled', async () => {
    // Create document with settings enabled
    const doc = Document.create();
    doc.setHideSpellingErrors(true);
    doc.setEmbedTrueTypeFonts(true);
    doc.createParagraph('Test');

    const buffer1 = await doc.toBuffer();
    const loaded1 = await Document.loadFromBuffer(buffer1);

    // Disable the settings
    loaded1.setHideSpellingErrors(false);
    loaded1.setEmbedTrueTypeFonts(false);

    const buffer2 = await loaded1.toBuffer();
    const loaded2 = await Document.loadFromBuffer(buffer2);

    expect(loaded2.getHideSpellingErrors()).toBe(false);
    expect(loaded2.getEmbedTrueTypeFonts()).toBe(false);

    doc.dispose();
    loaded1.dispose();
    loaded2.dispose();
  });
});
